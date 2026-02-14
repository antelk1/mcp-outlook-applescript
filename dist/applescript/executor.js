import { execFileSync } from 'node:child_process';
import { OutlookMcpError, ErrorCode, } from '../utils/errors.js';
const DEFAULT_TIMEOUT_MS = 30000;
const ERROR_PATTERNS = {
    notRunning: /not running|application isn't running/i,
    permissionDenied: /not authorized|permission denied|assistive access/i,
    timeout: /timed out|timeout/i,
    handlerFailed: /AppleEvent handler failed/i,
};
function categorizeError(errorMessage) {
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
export function escapeForAppleScript(value) {
    return value
        .replace(/\\/g, '\\\\')
        .replace(/"/g, '\\"')
        .replace(/\r\n/g, '" & return & "')
        .replace(/\n/g, '" & linefeed & "')
        .replace(/\r/g, '" & return & "');
}
export function executeAppleScript(script, options = {}) {
    const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    const execOptions = {
        encoding: 'utf8',
        timeout: timeoutMs,
        maxBuffer: 50 * 1024 * 1024, // 50MB for large results
        input: script,
    };
    try {
        // Execute via osascript with script passed on stdin (no shell interpretation)
        const output = execFileSync('osascript', [], execOptions);
        return {
            success: true,
            output: output.trim(),
        };
    }
    catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        const stderr = error?.stderr;
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
export function executeAppleScriptOrThrow(script, options = {}) {
    const result = executeAppleScript(script, options);
    if (!result.success) {
        const errorType = categorizeError(result.error ?? '');
        throw new AppleScriptExecutionError(result.error ?? 'Unknown error', errorType);
    }
    return result.output;
}
export function isOutlookRunning() {
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
const ERROR_TYPE_TO_CODE = {
    not_running: ErrorCode.OUTLOOK_NOT_RUNNING,
    permission_denied: ErrorCode.APPLESCRIPT_PERMISSION_DENIED,
    timeout: ErrorCode.APPLESCRIPT_TIMEOUT,
    unknown: ErrorCode.APPLESCRIPT_ERROR,
};
/**
 * Thrown when an osascript invocation fails.
 * Integrates with the OutlookMcpError hierarchy so callers
 * get a structured `code` field alongside the raw `errorType`.
 */
export class AppleScriptExecutionError extends OutlookMcpError {
    code;
    errorType;
    constructor(message, errorType) {
        super(message);
        this.errorType = errorType;
        this.code = ERROR_TYPE_TO_CODE[errorType];
    }
}
