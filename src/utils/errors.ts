/**
 * Custom error classes for the Outlook MCP server.
 */

/**
 * Error codes for categorizing errors.
 */
export const ErrorCode = {
    UNKNOWN: 'UNKNOWN',
    VALIDATION_ERROR: 'VALIDATION_ERROR',
    NOT_FOUND: 'NOT_FOUND',
    OUTLOOK_NOT_RUNNING: 'OUTLOOK_NOT_RUNNING',
    APPLESCRIPT_PERMISSION_DENIED: 'APPLESCRIPT_PERMISSION_DENIED',
    APPLESCRIPT_TIMEOUT: 'APPLESCRIPT_TIMEOUT',
    APPLESCRIPT_ERROR: 'APPLESCRIPT_ERROR',
    OUTLOOK_BRIDGE_STRESSED: 'OUTLOOK_BRIDGE_STRESSED',
    OUTLOOK_QUERY_REFUSED: 'OUTLOOK_QUERY_REFUSED',
    ATTACHMENT_NOT_FOUND: 'ATTACHMENT_NOT_FOUND',
    ATTACHMENT_TOO_LARGE: 'ATTACHMENT_TOO_LARGE',
    ATTACHMENT_SAVE_ERROR: 'ATTACHMENT_SAVE_ERROR',
    MAIL_SEND_ERROR: 'MAIL_SEND_ERROR',
    APPROVAL_EXPIRED: 'APPROVAL_EXPIRED',
    APPROVAL_INVALID: 'APPROVAL_INVALID',
    TARGET_CHANGED: 'TARGET_CHANGED',
} as const;
export type ErrorCode = (typeof ErrorCode)[keyof typeof ErrorCode];

/**
 * Base class for all Outlook MCP errors.
 */
export abstract class OutlookMcpError extends Error {
    abstract readonly code: ErrorCode;
    constructor(message: string) {
        super(message);
        this.name = this.constructor.name;
        Error.captureStackTrace(this, this.constructor);
    }
}

/**
 * Generic wrapper for unexpected errors.
 */
export class UnknownError extends OutlookMcpError {
    readonly code = ErrorCode.UNKNOWN;
    constructor(message: string, readonly cause?: Error | undefined) {
        super(message);
    }
}

/**
 * Thrown for input validation errors.
 */
export class ValidationError extends OutlookMcpError {
    readonly code = ErrorCode.VALIDATION_ERROR;
    constructor(message: string) {
        super(message);
    }
}

/**
 * Thrown when a requested resource is not found.
 */
export class NotFoundError extends OutlookMcpError {
    readonly code = ErrorCode.NOT_FOUND;
    constructor(resourceType: string, id: number | string) {
        super(`${resourceType} with ID ${id} not found`);
    }
}

/**
 * Type guard to check if an error is an OutlookMcpError.
 */
export function isOutlookMcpError(error: unknown): error is OutlookMcpError {
    return error instanceof OutlookMcpError;
}

/**
 * Wraps an unknown error in an OutlookMcpError if needed.
 */
export function wrapError(error: unknown, defaultMessage: string): OutlookMcpError {
    if (isOutlookMcpError(error)) {
        return error;
    }
    if (error instanceof Error) {
        return new UnknownError(error.message, error);
    }
    return new UnknownError(defaultMessage);
}

// =============================================================================
// AppleScript Errors
// =============================================================================

/**
 * Thrown when Outlook is not running and needs to be.
 */
export class OutlookNotRunningError extends OutlookMcpError {
    readonly code = ErrorCode.OUTLOOK_NOT_RUNNING;
    constructor() {
        super('Microsoft Outlook is not running. ' +
            'Please start Outlook and try again.');
    }
}

/**
 * Thrown when AppleScript automation permission is denied.
 */
export class AppleScriptPermissionError extends OutlookMcpError {
    readonly code = ErrorCode.APPLESCRIPT_PERMISSION_DENIED;
    constructor() {
        super('Automation permission denied for Microsoft Outlook. ' +
            'Please grant access in System Settings > Privacy & Security > Automation.');
    }
}

/**
 * Thrown when AppleScript execution times out.
 */
export class AppleScriptTimeoutError extends OutlookMcpError {
    readonly code = ErrorCode.APPLESCRIPT_TIMEOUT;
    constructor(operation: string) {
        super(`AppleScript operation timed out: ${operation}. ` +
            'This may happen with large data sets. Try reducing the limit.');
    }
}

/**
 * Thrown for general AppleScript errors.
 */
export class AppleScriptError extends OutlookMcpError {
    readonly code = ErrorCode.APPLESCRIPT_ERROR;
    constructor(message: string, readonly cause?: Error | undefined) {
        super(message);
    }
}

/**
 * Thrown when adaptive backoff classifies the Outlook AppleScript bridge as
 * degraded and refuses to issue further expensive operations. This is a
 * structurally different signal from APPLESCRIPT_TIMEOUT: the timeout means
 * "this specific call took too long"; BRIDGE_STRESSED means "the bridge
 * itself is stuck and the safety brake has engaged before we make it worse."
 *
 * Callers (notably the LLM-driven assistant) should treat this as actionable:
 * the user should run `~/.local/bin/outlook-safe-restart.sh` rather than retry.
 */
export class OutlookBridgeStressedError extends OutlookMcpError {
    readonly code = ErrorCode.OUTLOOK_BRIDGE_STRESSED;
    constructor(medianMs: number, operation: string) {
        super(`Outlook AppleScript bridge looks degraded ` +
            `(rolling median latency ${medianMs}ms over recent calls). ` +
            `Refused: ${operation}. ` +
            `Diagnose first: \`~/.local/bin/outlook-safe-restart.sh --check\` ` +
            `runs an independent probe. If --check says healthy, the MCP's ` +
            `in-memory state is stale — restart Claude Code to clear it ` +
            `(or wait up to 5 minutes for the rolling window to age out). ` +
            `If --check says degraded, run \`outlook-safe-restart.sh\` (no args) ` +
            `to safely restart Outlook — the safety guards will refuse if a ` +
            `draft is open.`);
    }
}

/**
 * Thrown by predictive guards that refuse an operation based on its INPUTS
 * rather than current bridge state. The canonical case is an unscoped
 * `searchEmails` call (no `folder_id`) against a mailbox with tens of
 * thousands of messages — Outlook would need to scan every message in every
 * folder, which has been empirically shown to destabilize the AppleScript
 * bridge in a single call (verified 2026-05-12).
 *
 * Structurally different from BRIDGE_STRESSED: that means "the bridge is
 * currently degraded"; QUERY_REFUSED means "this operation is known to be
 * dangerous regardless of current state — narrow it before retrying."
 */
export class OutlookQueryRefusedError extends OutlookMcpError {
    readonly code = ErrorCode.OUTLOOK_QUERY_REFUSED;
    constructor(operation: string, reason: string, suggestion: string) {
        super(`Refused ${operation}: ${reason}. ${suggestion}`);
    }
}

// =============================================================================
// Attachment and Email Errors
// =============================================================================

/**
 * Thrown when an attachment file cannot be found.
 */
export class AttachmentNotFoundError extends OutlookMcpError {
    readonly code = ErrorCode.ATTACHMENT_NOT_FOUND;
    constructor(path: string) {
        super(`Attachment file not found: ${path}. Please check the file path exists.`);
    }
}

/**
 * Thrown when an attachment exceeds the size limit.
 */
export class AttachmentTooLargeError extends OutlookMcpError {
    readonly code = ErrorCode.ATTACHMENT_TOO_LARGE;
    constructor(name: string, sizeBytes: number, maxBytes: number) {
        super(`Attachment "${name}" is ${Math.round(sizeBytes / 1024 / 1024)}MB ` +
            `which exceeds the maximum size of ${Math.round(maxBytes / 1024 / 1024)}MB.`);
    }
}

/**
 * Thrown when saving an attachment to disk fails.
 */
export class AttachmentSaveError extends OutlookMcpError {
    readonly code = ErrorCode.ATTACHMENT_SAVE_ERROR;
    constructor(name: string, reason: string) {
        super(`Failed to save attachment "${name}": ${reason}`);
    }
}

/**
 * Thrown when sending an email fails.
 */
export class MailSendError extends OutlookMcpError {
    readonly code = ErrorCode.MAIL_SEND_ERROR;
    constructor(reason: string) {
        super(`Failed to send email: ${reason}`);
    }
}

// =============================================================================
// Approval Errors
// =============================================================================

/**
 * Thrown when an approval token has expired.
 */
export class ApprovalExpiredError extends OutlookMcpError {
    readonly code = ErrorCode.APPROVAL_EXPIRED;
    constructor() {
        super('Approval token has expired. Please prepare the operation again.');
    }
}

/**
 * Thrown when an approval token is invalid.
 */
export class ApprovalInvalidError extends OutlookMcpError {
    readonly code = ErrorCode.APPROVAL_INVALID;
    constructor(reason: string) {
        super(`Invalid approval token: ${reason}`);
    }
}

/**
 * Thrown when the target has been modified since the approval was generated.
 */
export class TargetChangedError extends OutlookMcpError {
    readonly code = ErrorCode.TARGET_CHANGED;
    constructor() {
        super('The target has been modified since the approval was generated. ' +
            'Please prepare the operation again.');
    }
}
