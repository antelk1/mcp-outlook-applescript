import { executeAppleScript, executeAppleScriptOrThrow } from './executor.js';
import { AppleScriptError } from '../utils/errors.js';
import * as scripts from './scripts.js';
import * as parser from './parser.js';
import type { IContentReader } from '../tools/mail.js';
import type { IEventContentReader, EventDetails } from '../tools/calendar.js';
import type { IContactContentReader, ContactDetails } from '../tools/contacts.js';
import type { ITaskContentReader, TaskDetails } from '../tools/tasks.js';
import type { INoteContentReader, NoteDetails } from '../tools/notes.js';
import type { AttachmentInfo } from '../types/mail.js';
import type { SaveAttachmentResult } from './parser.js';

// =============================================================================
// Path Prefixes & Helpers
// =============================================================================

export const EMAIL_PATH_PREFIX = 'applescript-email:' as const;
export const EVENT_PATH_PREFIX = 'applescript-event:' as const;
export const CONTACT_PATH_PREFIX = 'applescript-contact:' as const;
export const TASK_PATH_PREFIX = 'applescript-task:' as const;
export const NOTE_PATH_PREFIX = 'applescript-note:' as const;

/** Extracts a numeric entity ID from a prefixed synthetic path string. */
function extractId(path: string | null, prefix: string): number | null {
    if (path == null || !path.startsWith(prefix)) {
        return null;
    }
    const idStr = path.substring(prefix.length);
    const id = parseInt(idStr, 10);
    return isNaN(id) ? null : id;
}

export function createEmailPath(id: number): string {
    return `${EMAIL_PATH_PREFIX}${id}`;
}

export function createEventPath(id: number): string {
    return `${EVENT_PATH_PREFIX}${id}`;
}

export function createContactPath(id: number): string {
    return `${CONTACT_PATH_PREFIX}${id}`;
}

export function createTaskPath(id: number): string {
    return `${TASK_PATH_PREFIX}${id}`;
}

export function createNotePath(id: number): string {
    return `${NOTE_PATH_PREFIX}${id}`;
}

// =============================================================================
// Generic Content Reader Pipeline
// =============================================================================

/**
 * Shared pipeline for all entity content readers.
 *
 * Steps: extract ID from path prefix → generate AppleScript →
 * execute → parse raw output → map to domain result.
 * Returns null at any failure point rather than throwing.
 */
function readEntityDetails<TParsed, TResult>(
    dataFilePath: string | null,
    prefix: string,
    buildScript: (id: number) => string,
    parse: (output: string) => TParsed | null,
    mapResult: (parsed: TParsed) => TResult,
): TResult | null {
    const id = extractId(dataFilePath, prefix);
    if (id == null) {
        return null;
    }
    try {
        const script = buildScript(id);
        const result = executeAppleScript(script);
        if (!result.success) {
            return null;
        }
        const parsed = parse(result.output);
        if (parsed == null) {
            return null;
        }
        return mapResult(parsed);
    }
    catch {
        return null;
    }
}

// =============================================================================
// Entity Content Readers
// =============================================================================

/** Reads the full body of an email by its synthetic path. */
export class AppleScriptEmailContentReader implements IContentReader {
    readEmailBody(dataFilePath: string | null): string | null {
        return readEntityDetails(
            dataFilePath,
            EMAIL_PATH_PREFIX,
            scripts.getMessage,
            parser.parseEmail,
            (email) => email.htmlContent ?? email.plainContent,
        );
    }
}

/** Reads full event details (title, location, attendees, etc.) by synthetic path. */
export class AppleScriptEventContentReader implements IEventContentReader {
    readEventDetails(dataFilePath: string | null): EventDetails | null {
        return readEntityDetails(
            dataFilePath,
            EVENT_PATH_PREFIX,
            scripts.getEvent,
            parser.parseEvent,
            (event) => ({
                title: event.subject,
                location: event.location,
                description: event.htmlContent ?? event.plainContent,
                organizer: event.organizer,
                attendees: event.attendees.map((a) => ({
                    email: a.email,
                    name: a.name,
                    status: 'unknown' as const,
                })),
            }),
        );
    }
}

/** Reads full contact details (emails, phones, addresses) by synthetic path. */
export class AppleScriptContactContentReader implements IContactContentReader {
    readContactDetails(dataFilePath: string | null): ContactDetails | null {
        return readEntityDetails(
            dataFilePath,
            CONTACT_PATH_PREFIX,
            scripts.getContact,
            parser.parseContact,
            (contact) => {
                const emails = contact.emails.map((e) => ({
                    type: 'work',
                    address: e,
                }));
                const phones: Array<{ type: string; number: string }> = [];
                if (contact.homePhone != null) {
                    phones.push({ type: 'home', number: contact.homePhone });
                }
                if (contact.workPhone != null) {
                    phones.push({ type: 'work', number: contact.workPhone });
                }
                if (contact.mobilePhone != null) {
                    phones.push({ type: 'mobile', number: contact.mobilePhone });
                }
                const addresses: Array<{
                    type: string;
                    street: string | null;
                    city: string | null;
                    state: string | null;
                    postalCode: string | null;
                    country: string | null;
                }> = [];
                if (contact.homeStreet != null ||
                    contact.homeCity != null ||
                    contact.homeState != null ||
                    contact.homeZip != null ||
                    contact.homeCountry != null) {
                    addresses.push({
                        type: 'home',
                        street: contact.homeStreet,
                        city: contact.homeCity,
                        state: contact.homeState,
                        postalCode: contact.homeZip,
                        country: contact.homeCountry,
                    });
                }
                return {
                    firstName: contact.firstName,
                    lastName: contact.lastName,
                    middleName: contact.middleName,
                    nickname: contact.nickname,
                    company: contact.company,
                    jobTitle: contact.jobTitle,
                    department: contact.department,
                    emails,
                    phones,
                    addresses,
                    notes: contact.notes,
                };
            },
        );
    }
}

/** Reads task body and completion info by synthetic path. */
export class AppleScriptTaskContentReader implements ITaskContentReader {
    readTaskDetails(dataFilePath: string | null): TaskDetails | null {
        return readEntityDetails(
            dataFilePath,
            TASK_PATH_PREFIX,
            scripts.getTask,
            parser.parseTask,
            (task) => ({
                body: task.htmlContent ?? task.plainContent,
                completedDate: task.completedDate,
                reminderDate: null,
                categories: [],
            }),
        );
    }
}

/** Reads note content with a short preview by synthetic path. */
export class AppleScriptNoteContentReader implements INoteContentReader {
    readNoteDetails(dataFilePath: string | null): NoteDetails | null {
        return readEntityDetails(
            dataFilePath,
            NOTE_PATH_PREFIX,
            scripts.getNote,
            parser.parseNote,
            (note) => {
                const body = note.htmlContent ?? note.plainContent ?? '';
                const preview = body.substring(0, 200);
                return {
                    title: note.name,
                    body,
                    preview,
                    createdDate: note.createdDate,
                    categories: [],
                };
            },
        );
    }
}

// =============================================================================
// Attachment Reader (separate pattern — not entity-based)
// =============================================================================

export interface IAttachmentReader {
    listAttachments(emailId: number): AttachmentInfo[];
    saveAttachment(emailId: number, attachmentIndex: number, savePath: string): SaveAttachmentResult;
}

/** Reads attachment metadata and saves attachment files via AppleScript. */
export class AppleScriptAttachmentReader implements IAttachmentReader {
    listAttachments(emailId: number): AttachmentInfo[] {
        try {
            const script = scripts.listAttachments(emailId);
            const result = executeAppleScript(script);
            if (!result.success) {
                return [];
            }
            const rows = parser.parseAttachments(result.output);
            return rows.map((r) => ({
                index: r.index,
                name: r.name,
                size: r.fileSize,
                contentType: r.contentType,
            }));
        }
        catch {
            return [];
        }
    }

    saveAttachment(emailId: number, attachmentIndex: number, savePath: string): SaveAttachmentResult {
        const script = scripts.saveAttachment(emailId, attachmentIndex, savePath);
        const output = executeAppleScriptOrThrow(script);
        const result = parser.parseSaveAttachmentResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse save attachment response');
        }
        return result;
    }
}

// =============================================================================
// Factory
// =============================================================================

export interface AppleScriptContentReaders {
    readonly email: IContentReader;
    readonly event: IEventContentReader;
    readonly contact: IContactContentReader;
    readonly task: ITaskContentReader;
    readonly note: INoteContentReader;
    readonly attachment: IAttachmentReader;
}

/** Creates all content reader instances for the AppleScript backend. */
export function createAppleScriptContentReaders(): AppleScriptContentReaders {
    return {
        email: new AppleScriptEmailContentReader(),
        event: new AppleScriptEventContentReader(),
        contact: new AppleScriptContactContentReader(),
        task: new AppleScriptTaskContentReader(),
        note: new AppleScriptNoteContentReader(),
        attachment: new AppleScriptAttachmentReader(),
    };
}
