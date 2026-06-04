import { existsSync, writeFileSync, unlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomBytes } from 'node:crypto';
import { executeAppleScriptOrThrow, assertBridgeHealthyForWrite, AppleScriptExecutionError } from './executor.js';
import * as scripts from './scripts.js';
import { parseSendEmailResult } from './parser.js';
import { AppleScriptError, AttachmentNotFoundError, MailSendError, MailSendIndeterminateError } from '../utils/errors.js';
/**
 * Node-level timeout for the send osascript. MUST be larger than the inner
 * AppleScript `with timeout` (45s in scripts.sendEmail) so the AppleEvent times
 * out cleanly (-1712) instead of Node SIGKILLing osascript mid-send — the
 * latter corrupts the bridge for all subsequent calls.
 */
const SEND_NODE_TIMEOUT_MS = 60_000;
/** Sends email through Outlook via AppleScript. */
export class AppleScriptMailSender {
    /**
     * Validates attachment paths, builds the send-email AppleScript, executes
     * it, and parses the result.
     * @param params - Email composition parameters (recipients, body, attachments, etc.).
     * @returns The sent message's ID and timestamp.
     * @throws AttachmentNotFoundError if any attachment or inline image path does not exist.
     * @throws MailSendError if Outlook reports a send failure.
     */
    sendEmail(params) {
        if (params.attachments != null) {
            for (const attachment of params.attachments) {
                if (!existsSync(attachment.path)) {
                    throw new AttachmentNotFoundError(attachment.path);
                }
            }
        }
        if (params.inlineImages != null) {
            for (const image of params.inlineImages) {
                if (!existsSync(image.path)) {
                    throw new AttachmentNotFoundError(image.path);
                }
            }
        }
        // Pre-send health gate: refuse with guidance if the bridge is degraded,
        // BEFORE composing anything. Never fire a send into a stuck bridge — it
        // hangs and gets SIGKILLed mid-AppleEvent, corrupting the bridge.
        assertBridgeHealthyForWrite('send_email');
        // For HTML bodies, write to a temp file so AppleScript can read it
        // instead of embedding HTML in the AppleScript string literal (which
        // breaks the parser on quotes, colons, and other special chars).
        let bodyFilePath;
        if (params.bodyType === 'html') {
            const id = randomBytes(8).toString('hex');
            bodyFilePath = join(tmpdir(), `mcp-outlook-html-${id}.html`);
            writeFileSync(bodyFilePath, params.body, 'utf8');
        }
        try {
            let scriptParams = {
                to: params.to,
                subject: params.subject,
                body: params.body,
                bodyType: params.bodyType,
                ...(bodyFilePath != null && { bodyFilePath }),
            };
            if (params.cc != null)
                scriptParams = { ...scriptParams, cc: params.cc };
            if (params.bcc != null)
                scriptParams = { ...scriptParams, bcc: params.bcc };
            if (params.replyTo != null)
                scriptParams = { ...scriptParams, replyTo: params.replyTo };
            if (params.attachments != null)
                scriptParams = { ...scriptParams, attachments: params.attachments };
            if (params.inlineImages != null)
                scriptParams = { ...scriptParams, inlineImages: params.inlineImages };
            if (params.accountId != null)
                scriptParams = { ...scriptParams, accountId: params.accountId };
            const script = scripts.sendEmail(scriptParams);
            let output;
            try {
                output = executeAppleScriptOrThrow(script, { timeoutMs: SEND_NODE_TIMEOUT_MS });
            }
            catch (err) {
                // Node killed osascript at 60s — even the 45s inner timeout
                // didn't fire. Outcome indeterminate; do not let callers retry.
                if (err instanceof AppleScriptExecutionError && err.errorType === 'timeout') {
                    throw new MailSendIndeterminateError('osascript was killed at the Node 60s timeout');
                }
                throw err;
            }
            const result = parseSendEmailResult(output);
            if (result == null) {
                throw new AppleScriptError('Failed to parse send email response');
            }
            if (!result.success) {
                const reason = result.error ?? 'Unknown error';
                // Inner `with timeout` fired (-1712): the send was issued but its
                // reply timed out, so it may or may not have queued. Signal
                // INDETERMINATE so the caller verifies rather than blind-retrying.
                if (/-1712|AppleEvent timed out|timed out/i.test(reason)) {
                    throw new MailSendIndeterminateError(`AppleScript send timed out: ${reason}`);
                }
                throw new MailSendError(reason);
            }
            return {
                messageId: result.messageId ?? '',
                sentAt: result.sentAt ?? '',
            };
        }
        finally {
            // Clean up temp file
            if (bodyFilePath != null) {
                try {
                    unlinkSync(bodyFilePath);
                }
                catch { /* ignore cleanup errors */ }
            }
        }
    }
}
/** Creates a new AppleScriptMailSender instance. */
export function createMailSender() {
    return new AppleScriptMailSender();
}
