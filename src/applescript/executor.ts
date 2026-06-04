import { execFileSync } from 'node:child_process';
import {
    OutlookMcpError,
    ErrorCode,
    OutlookBridgeUnhealthyError,
} from '../utils/errors.js';
import { recordLatency, waitForSlot } from './throttle.js';

export interface AppleScriptResult {
    readonly success: boolean;
    readonly output: string;
    readonly error?: string;
}

export interface ExecuteOptions {
    readonly timeoutMs?: number;
    /**
     * When true, skip the adaptive inter-call throttle. Use only for the
     * throttle's own probe path or for synchronous health checks that must
     * not contribute to bridge stress measurements (e.g. isOutlookRunning).
     */
    readonly skipThrottle?: boolean;
}

const DEFAULT_TIMEOUT_MS = 30000;

const ERROR_PATTERNS = {
    notRunning: /not running|application isn't running/i,
    permissionDenied: /not authorized|permission denied|assistive access/i,
    timeout: /timed out|timeout/i,
    handlerFailed: /AppleEvent handler failed/i,
} as const;

function categorizeError(errorMessage: string): 'not_running' | 'permission_denied' | 'timeout' | 'unknown' {
    if (ERROR_PATTERNS.notRunning.test(errorMessage)) {
        return 'not_running';
    }
    if (ERROR_PATTERNS.permissionDenied.test(errorMessage)) {
        return 'permission_denied';
    }
    if (ERROR_PATTERNS.timeout.test(errorMessage)) {
        return 'timeout';
    }
    return 'unknown';
}

export function escapeForAppleScript(value: string): string {
    return value
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\r\n/g, '" & return & "')
        .replace(/\n/g, '" & linefeed & "')
        .replace(/\r/g, '" & return & "');
}

export function executeAppleScript(script: string, options: ExecuteOptions = {}): AppleScriptResult {
    const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    const execOptions = {
        encoding: 'utf8' as const,
        timeout: timeoutMs,
        maxBuffer: 50 * 1024 * 1024, // 50MB for large results
        input: script,
    };

    // Throttle: pace AppleScript calls to give Outlook's bridge breathing room.
    // Default 100–1000ms gap (adaptive per bridge state). Skipped for the
    // throttle's own internal probes to avoid feedback loops.
    if (!options.skipThrottle) {
        waitForSlot();
    }

    const startNs = process.hrtime.bigint();
    try {
        // Execute via osascript with script passed on stdin (no shell interpretation)
        const output = execFileSync('osascript', [], execOptions);
        const elapsedMs = Number((process.hrtime.bigint() - startNs) / 1_000_000n);
        if (!options.skipThrottle) {
            recordLatency(elapsedMs);
        }
        return {
            success: true,
            output: (output as string).trim(),
        };
    }
    catch (error: unknown) {
        const elapsedMs = Number((process.hrtime.bigint() - startNs) / 1_000_000n);
        if (!options.skipThrottle) {
            recordLatency(elapsedMs);
        }
        const errorMessage = error instanceof Error ? error.message : String(error);
        const stderr = (error as { stderr?: Buffer | string | undefined })?.stderr;
        let stderrText = '';
        if (stderr instanceof Buffer) {
            stderrText = stderr.toString('utf8');
        }
        else if (typeof stderr === 'string') {
            stderrText = stderr;
        }
        const fullError = `${errorMessage}\n${stderrText}`.trim();
        return {
            success: false,
            output: '',
            error: fullError,
        };
    }
}

export function executeAppleScriptOrThrow(script: string, options: ExecuteOptions = {}): string {
    const result = executeAppleScript(script, options);
    if (!result.success) {
        const errorType = categorizeError(result.error ?? '');
        throw new AppleScriptExecutionError(result.error ?? 'Unknown error', errorType);
    }
    return result.output;
}

/**
 * Like executeAppleScriptOrThrow, but retries once on a clean AppleEvent timeout
 * (-1712 from Outlook's `with timeout`) or generic timeout error.
 *
 * Outlook's first heavy AppleScript call after idle takes 30–45s while it pages
 * in indexes; subsequent calls are 1–4s. This retry handles the cold start by
 * issuing a warm-up call (`count of messages of outbox`) to ensure the bridge
 * is responsive, then retrying the original script.
 *
 * The warm-up query was chosen specifically because it's the same one the
 * watchdog (`outlook-safe-restart.sh`) uses to detect bridge degradation: it
 * touches the message store, not just app metadata. The earlier `return name`
 * was a metadata-only call and could falsely pass while real queries still
 * hung (verified 2026-05-12 — window enumeration worked while `count of
 * messages of outbox` returned -1712).
 *
 * Use only for read-only operations — never for mutations, where retry-after-
 * timeout could double-apply.
 */
export function executeAppleScriptWithRetry(script: string, options: ExecuteOptions = {}): string {
    const result = executeAppleScript(script, options);
    if (result.success) {
        return result.output;
    }
    const errorType = categorizeError(result.error ?? '');
    const isTimeout = errorType === 'timeout' || /-1712|AppleEvent timed out|spawnSync osascript ETIMEDOUT/i.test(result.error ?? '');
    if (!isTimeout) {
        throw new AppleScriptExecutionError(result.error ?? 'Unknown error', errorType);
    }
    // Cold-start fallback: warm up the bridge, then retry once with the same timeout.
    // The warmup itself skips the throttle so we don't double-count its latency.
    executeAppleScript(WARMUP_SCRIPT, { timeoutMs: 5000, skipThrottle: true });
    const retry = executeAppleScript(script, options);
    if (!retry.success) {
        const retryType = categorizeError(retry.error ?? '');
        throw new AppleScriptExecutionError(retry.error ?? 'Unknown error', retryType);
    }
    return retry.output;
}

const WARMUP_SCRIPT = `tell application "Microsoft Outlook" to count of messages of outbox`;

export function isOutlookRunning(): boolean {
    const script = `
tell application "System Events"
  set isRunning to (name of processes) contains "Microsoft Outlook"
  return isRunning
end tell
`;
    // This probe targets System Events, not Outlook itself — don't let its
    // timing pollute the bridge-stress window or take a throttle slot.
    const result = executeAppleScript(script, { skipThrottle: true });
    return result.success && result.output.toLowerCase() === 'true';
}

// =============================================================================
// Write-path health gate
// =============================================================================

const GROUND_TRUTH_PROBE_QUERY = 'tell application "Microsoft Outlook" to count of messages of outbox';
const WRITE_GATE_PROBE_TIMEOUT_MS = 2500;
/** A probe slower than this (or errored) means the bridge is too degraded to safely mutate. */
const WRITE_GATE_HEALTHY_MS = 800;

export interface BridgeHealth {
    readonly healthy: boolean;
    readonly probeMs: number;
    readonly probeSucceeded: boolean;
    readonly error?: string;
}

/**
 * Fast ground-truth probe of bridge health, mirroring the watchdog's
 * `count of messages of outbox` (touches the message store, not just app
 * metadata). Healthy bridges answer in <200ms; a stuck bridge errors at -1712
 * within the 2.5s window. Skips the throttle so it never waits behind a queued
 * slot, and skips latency recording so the probe itself doesn't move the window.
 */
export function probeBridgeHealth(): BridgeHealth {
    const start = Date.now();
    const r = executeAppleScript(GROUND_TRUTH_PROBE_QUERY, {
        timeoutMs: WRITE_GATE_PROBE_TIMEOUT_MS,
        skipThrottle: true,
    });
    const probeMs = Date.now() - start;
    return {
        healthy: r.success && probeMs < WRITE_GATE_HEALTHY_MS,
        probeMs,
        probeSucceeded: r.success,
        ...(r.error != null && { error: r.error }),
    };
}

/**
 * Gate a mutating operation (send, move, delete) on a fresh health probe.
 * Throws OutlookBridgeUnhealthyError WITHOUT attempting the mutation when the
 * bridge is degraded — so a `send` never hangs on a stuck bridge and gets
 * SIGKILLed mid-AppleEvent (the sequence that corrupts the bridge). This is the
 * cheap, unconditional pre-write check; read paths use the repository's
 * `gateExpensive` (which only probes when the rolling median already looks bad).
 */
export function assertBridgeHealthyForWrite(operation: string): void {
    const h = probeBridgeHealth();
    if (h.healthy) return;
    throw new OutlookBridgeUnhealthyError(operation, h.probeMs, h.probeSucceeded);
}


/** Maps errorType strings to structured ErrorCode values. */
const ERROR_TYPE_TO_CODE: Record<AppleScriptErrorType, ErrorCode> = {
    not_running: ErrorCode.OUTLOOK_NOT_RUNNING,
    permission_denied: ErrorCode.APPLESCRIPT_PERMISSION_DENIED,
    timeout: ErrorCode.APPLESCRIPT_TIMEOUT,
    unknown: ErrorCode.APPLESCRIPT_ERROR,
};

type AppleScriptErrorType = 'not_running' | 'permission_denied' | 'timeout' | 'unknown';

/**
 * Thrown when an osascript invocation fails.
 * Integrates with the OutlookMcpError hierarchy so callers
 * get a structured `code` field alongside the raw `errorType`.
 */
export class AppleScriptExecutionError extends OutlookMcpError {
    readonly code: ErrorCode;
    readonly errorType: AppleScriptErrorType;
    constructor(message: string, errorType: AppleScriptErrorType) {
        super(message);
        this.errorType = errorType;
        this.code = ERROR_TYPE_TO_CODE[errorType];
    }
}
