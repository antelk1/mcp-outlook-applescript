import { execSync } from 'node:child_process';
import {
    OutlookMcpError,
    ErrorCode,
} from '../utils/errors.js';

export interface AppleScriptResult {
    readonly success: boolean;
    readonly output: string;
    readonly error?: string;
}

export interface ExecuteOptions {
    readonly timeoutMs?: number;
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
        stdio: ['pipe', 'pipe', 'pipe'] as ['pipe', 'pipe', 'pipe'],
    };
    try {
        // Execute via osascript
        const output = execSync(`osascript -e '${script.replace(/'/g, "'\"'\"'")}'`, execOptions);
        return {
            success: true,
            output: (output as string).trim(),
        };
    }
    catch (error: unknown) {
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

export function isOutlookRunning(): boolean {
    const script = `
tell application "System Events"
  set isRunning to (name of processes) contains "Microsoft Outlook"
  return isRunning
end tell
`;
    const result = executeAppleScript(script);
    return result.success && result.output.toLowerCase() === 'true';
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
