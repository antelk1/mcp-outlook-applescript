import { DELIMITERS } from './scripts.js';

// ---------------------------------------------------------------------------
// Row interfaces — typed representations of AppleScript output records
// ---------------------------------------------------------------------------

/** Represents a mail folder returned from AppleScript. */
export interface AppleScriptFolderRow {
    readonly id: number;
    readonly name: string | null;
    readonly unreadCount: number;
}

/** Represents an email attachment's metadata from AppleScript. */
export interface AppleScriptAttachmentRow {
    readonly index: number;
    readonly name: string;
    readonly fileSize: number;
    readonly contentType: string;
}

/** Represents a single email message returned from AppleScript. */
export interface AppleScriptEmailRow {
    readonly id: number;
    readonly folderId: number | null;
    readonly subject: string | null;
    readonly senderName: string | null;
    readonly senderEmail: string | null;
    readonly toRecipients: string | null;
    readonly ccRecipients: string | null;
    readonly preview: string | null;
    readonly isRead: boolean;
    readonly dateReceived: string | null;
    readonly dateSent: string | null;
    readonly priority: string;
    readonly htmlContent: string | null;
    readonly plainContent: string | null;
    readonly hasHtml: boolean;
    readonly attachments: string[];
    readonly attachmentDetails: AppleScriptAttachmentRow[];
    readonly flagStatus: number | null;
}

/** Represents a calendar folder returned from AppleScript. */
export interface AppleScriptCalendarRow {
    readonly id: number;
    readonly name: string | null;
}

/** Represents a calendar event returned from AppleScript. */
export interface AppleScriptEventRow {
    readonly id: number;
    readonly calendarId: number | null;
    readonly subject: string | null;
    readonly startTime: string | null;
    readonly endTime: string | null;
    readonly location: string | null;
    readonly isAllDay: boolean;
    readonly isRecurring: boolean;
    readonly organizer: string | null;
    readonly htmlContent: string | null;
    readonly plainContent: string | null;
    readonly attendees: Array<{
        email: string;
        name: string;
    }>;
}

/** Represents a contact record returned from AppleScript. */
export interface AppleScriptContactRow {
    readonly id: number;
    readonly displayName: string | null;
    readonly firstName: string | null;
    readonly lastName: string | null;
    readonly middleName: string | null;
    readonly nickname: string | null;
    readonly company: string | null;
    readonly jobTitle: string | null;
    readonly department: string | null;
    readonly notes: string | null;
    readonly emails: string[];
    readonly homePhone: string | null;
    readonly workPhone: string | null;
    readonly mobilePhone: string | null;
    readonly homeStreet: string | null;
    readonly homeCity: string | null;
    readonly homeState: string | null;
    readonly homeZip: string | null;
    readonly homeCountry: string | null;
}

/** Represents a task record returned from AppleScript. */
export interface AppleScriptTaskRow {
    readonly id: number;
    readonly folderId: number | null;
    readonly name: string | null;
    readonly isCompleted: boolean;
    readonly dueDate: string | null;
    readonly startDate: string | null;
    readonly completedDate: string | null;
    readonly priority: string;
    readonly htmlContent: string | null;
    readonly plainContent: string | null;
}

/** Represents a note record returned from AppleScript. */
export interface AppleScriptNoteRow {
    readonly id: number;
    readonly folderId: number | null;
    readonly name: string | null;
    readonly createdDate: string | null;
    readonly modifiedDate: string | null;
    readonly preview: string | null;
    readonly htmlContent: string | null;
    readonly plainContent: string | null;
}

// ---------------------------------------------------------------------------
// Result interfaces — outcome types for mutating AppleScript operations
// ---------------------------------------------------------------------------

/** Outcome of responding to a calendar event invitation. */
export interface RespondToEventResult {
    readonly success: boolean;
    readonly eventId?: number;
    readonly error?: string;
}

/** Outcome of deleting a calendar event. */
export interface DeleteEventResult {
    readonly success: boolean;
    readonly eventId?: number;
    readonly error?: string;
}

/** Outcome of saving an email attachment to disk. */
export interface SaveAttachmentResult {
    readonly success: boolean;
    readonly name?: string;
    readonly savedTo?: string;
    readonly fileSize?: number;
    readonly error?: string;
}

/** Represents an Outlook account (Exchange, IMAP, etc.). */
export interface AppleScriptAccountRow {
    readonly id: number;
    readonly name: string | null;
    readonly email: string | null;
    readonly type: string;
}

/** Extends a folder row with its parent account ID and total message count. */
export interface AppleScriptFolderWithAccountRow extends AppleScriptFolderRow {
    readonly accountId: number;
    readonly messageCount: number;
}

/** Outcome of updating a calendar event's properties. */
export interface UpdateEventResult {
    readonly success: boolean;
    readonly id?: number;
    readonly updatedFields?: readonly string[];
    readonly error?: string;
}

/** Outcome of sending an email message. */
export interface SendEmailResult {
    readonly success: boolean;
    readonly messageId?: string;
    readonly sentAt?: string;
    readonly error?: string;
}

// ---------------------------------------------------------------------------
// Primitive parsers — convert raw string tokens into typed values
// ---------------------------------------------------------------------------

/**
 * Splits raw delimiter-separated AppleScript output into an array of key-value records.
 * @param output - Raw string from osascript stdout.
 * @returns Array of string-keyed dictionaries, one per record.
 */
function parseRawOutput(output: string): Record<string, string>[] {
    if (output.trim().length === 0) {
        return [];
    }
    const records: Record<string, string>[] = [];
    const recordStrings = output.split(DELIMITERS.RECORD).filter((s) => s.length > 0);
    for (const recordStr of recordStrings) {
        const record: Record<string, string> = {};
        const fieldStrings = recordStr.split(DELIMITERS.FIELD);
        for (const fieldStr of fieldStrings) {
            const [key, value] = fieldStr.split(DELIMITERS.EQUALS);
            if (key !== undefined && value !== undefined) {
                record[key] = value;
            }
        }
        if (Object.keys(record).length > 0) {
            records.push(record);
        }
    }
    return records;
}

/**
 * Converts a raw string token to a number, defaulting to 0 for missing or invalid values.
 * @param value - Raw token from the parsed record.
 * @returns Parsed integer or 0.
 */
function parseNumber(value: string | undefined): number {
    if (value === undefined || value === '' || value === DELIMITERS.NULL) {
        return 0;
    }
    const num = parseInt(value, 10);
    return isNaN(num) ? 0 : num;
}

/**
 * Converts a raw string token to a number or null for missing/invalid values.
 * @param value - Raw token from the parsed record.
 * @returns Parsed integer or null.
 */
function parseNumberOrNull(value: string | undefined): number | null {
    if (value === undefined || value === '' || value === DELIMITERS.NULL) {
        return null;
    }
    const num = parseInt(value, 10);
    return isNaN(num) ? null : num;
}

/**
 * Converts a raw string token to a boolean (case-insensitive "true" check).
 * @param value - Raw token from the parsed record.
 * @returns True only when the token is literally "true".
 */
function parseBoolean(value: string | undefined): boolean {
    if (value === undefined) {
        return false;
    }
    return value.toLowerCase() === 'true';
}

/**
 * Converts a raw string token to a trimmed string or null for empty/sentinel values.
 * @param value - Raw token from the parsed record.
 * @returns The string value, or null if empty, null-sentinel, or "missing value".
 */
function parseString(value: string | undefined): string | null {
    if (value === undefined || value === '' || value === DELIMITERS.NULL || value === 'missing value') {
        return null;
    }
    return value;
}

/**
 * Splits a comma-separated token into a trimmed string array.
 * @param value - Comma-delimited raw token.
 * @returns Array of non-empty trimmed strings.
 */
function parseList(value: string | undefined): string[] {
    if (value === undefined || value === '' || value === DELIMITERS.NULL) {
        return [];
    }
    return value.split(',').filter((s) => s.length > 0).map((s) => s.trim());
}

/**
 * Parses a comma-separated list of pipe-delimited attachment detail tokens.
 * Each token has the format: index|name|fileSize|contentType.
 * @param value - Raw attachment details string.
 * @returns Array of structured attachment rows.
 */
function parseAttachmentDetails(value: string | undefined): AppleScriptAttachmentRow[] {
    if (value === undefined || value === '' || value === DELIMITERS.NULL) {
        return [];
    }
    const items = value.split(',').filter((s) => s.length > 0);
    return items.map((item) => {
        const parts = item.split('|');
        return {
            index: parseInt(parts[0] ?? '0', 10) || 0,
            name: parts[1] ?? '',
            fileSize: parseInt(parts[2] ?? '0', 10) || 0,
            contentType: parts[3] ?? 'application/octet-stream',
        };
    });
}

/**
 * Parses a comma-separated list of pipe-delimited attendee tokens.
 * Each token has the format: email|name.
 * @param value - Raw attendees string.
 * @returns Array of attendee objects with email and name.
 */
function parseAttendees(value: string | undefined): Array<{ email: string; name: string }> {
    if (value === undefined || value === '' || value === DELIMITERS.NULL) {
        return [];
    }
    const attendees: Array<{ email: string; name: string }> = [];
    const items = value.split(',').filter((s) => s.length > 0);
    for (const item of items) {
        const parts = item.split('|');
        if (parts[0] !== undefined) {
            attendees.push({
                email: parts[0].trim(),
                name: parts[1]?.trim() ?? '',
            });
        }
    }
    return attendees;
}

// ---------------------------------------------------------------------------
// Entity parsers — transform raw output into typed row arrays
// ---------------------------------------------------------------------------

/**
 * Parses AppleScript output into an array of mail folder rows.
 * @param output - Raw osascript stdout.
 * @returns Typed folder rows.
 */
export function parseFolders(output: string): AppleScriptFolderRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        name: parseString(r['name']),
        unreadCount: parseNumber(r['unreadCount']),
    }));
}

/**
 * Parses AppleScript output into an array of email rows.
 * @param output - Raw osascript stdout.
 * @returns Typed email rows with attachments, flags, and content.
 */
export function parseEmails(output: string): AppleScriptEmailRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        folderId: parseNumberOrNull(r['folderId']),
        subject: parseString(r['subject']),
        senderName: parseString(r['senderName']),
        senderEmail: parseString(r['senderEmail']),
        toRecipients: parseString(r['toRecipients']),
        ccRecipients: parseString(r['ccRecipients']),
        preview: parseString(r['preview']),
        isRead: parseBoolean(r['isRead']),
        dateReceived: parseString(r['dateReceived']),
        dateSent: parseString(r['dateSent']),
        priority: r['priority'] ?? 'normal',
        htmlContent: parseString(r['htmlContent']),
        plainContent: parseString(r['plainContent']),
        hasHtml: parseBoolean(r['hasHtml']),
        attachments: parseList(r['attachments']),
        attachmentDetails: parseAttachmentDetails(r['attachmentDetails']),
        flagStatus: parseNumberOrNull(r['flagStatus']),
    }));
}

/**
 * Parses AppleScript output into a single email row, or null if none found.
 * @param output - Raw osascript stdout.
 * @returns The first email row, or null.
 */
export function parseEmail(output: string): AppleScriptEmailRow | null {
    const emails = parseEmails(output);
    return emails[0] ?? null;
}

/**
 * Parses AppleScript output into an array of calendar rows.
 * @param output - Raw osascript stdout.
 * @returns Typed calendar rows.
 */
export function parseCalendars(output: string): AppleScriptCalendarRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        name: parseString(r['name']),
    }));
}

/**
 * Parses AppleScript output into an array of calendar event rows.
 * @param output - Raw osascript stdout.
 * @returns Typed event rows with attendees and content.
 */
export function parseEvents(output: string): AppleScriptEventRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        calendarId: parseNumberOrNull(r['calendarId']),
        subject: parseString(r['subject']),
        startTime: parseString(r['startTime']),
        endTime: parseString(r['endTime']),
        location: parseString(r['location']),
        isAllDay: parseBoolean(r['isAllDay']),
        isRecurring: parseBoolean(r['isRecurring']),
        organizer: parseString(r['organizer']),
        htmlContent: parseString(r['htmlContent']),
        plainContent: parseString(r['plainContent']),
        attendees: parseAttendees(r['attendees']),
    }));
}

/**
 * Parses AppleScript output into a single event row, or null if none found.
 * @param output - Raw osascript stdout.
 * @returns The first event row, or null.
 */
export function parseEvent(output: string): AppleScriptEventRow | null {
    const events = parseEvents(output);
    return events[0] ?? null;
}

/**
 * Parses AppleScript output into an array of contact rows.
 * @param output - Raw osascript stdout.
 * @returns Typed contact rows with phone numbers, emails, and addresses.
 */
export function parseContacts(output: string): AppleScriptContactRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        displayName: parseString(r['displayName']),
        firstName: parseString(r['firstName']),
        lastName: parseString(r['lastName']),
        middleName: parseString(r['middleName']),
        nickname: parseString(r['nickname']),
        company: parseString(r['company']),
        jobTitle: parseString(r['jobTitle']),
        department: parseString(r['department']),
        notes: parseString(r['notes']),
        emails: parseList(r['emails']),
        homePhone: parseString(r['homePhone']),
        workPhone: parseString(r['workPhone']),
        mobilePhone: parseString(r['mobilePhone']),
        homeStreet: parseString(r['homeStreet']),
        homeCity: parseString(r['homeCity']),
        homeState: parseString(r['homeState']),
        homeZip: parseString(r['homeZip']),
        homeCountry: parseString(r['homeCountry']),
    }));
}

/**
 * Parses AppleScript output into a single contact row, or null if none found.
 * @param output - Raw osascript stdout.
 * @returns The first contact row, or null.
 */
export function parseContact(output: string): AppleScriptContactRow | null {
    const contacts = parseContacts(output);
    return contacts[0] ?? null;
}

/**
 * Parses AppleScript output into an array of task rows.
 * @param output - Raw osascript stdout.
 * @returns Typed task rows with dates, priority, and content.
 */
export function parseTasks(output: string): AppleScriptTaskRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        folderId: parseNumberOrNull(r['folderId']),
        name: parseString(r['name']),
        isCompleted: parseBoolean(r['isCompleted']),
        dueDate: parseString(r['dueDate']),
        startDate: parseString(r['startDate']),
        completedDate: parseString(r['completedDate']),
        priority: r['priority'] ?? 'normal',
        htmlContent: parseString(r['htmlContent']),
        plainContent: parseString(r['plainContent']),
    }));
}

/**
 * Parses AppleScript output into a single task row, or null if none found.
 * @param output - Raw osascript stdout.
 * @returns The first task row, or null.
 */
export function parseTask(output: string): AppleScriptTaskRow | null {
    const tasks = parseTasks(output);
    return tasks[0] ?? null;
}

/**
 * Parses AppleScript output into an array of note rows.
 * @param output - Raw osascript stdout.
 * @returns Typed note rows with dates, preview, and content.
 */
export function parseNotes(output: string): AppleScriptNoteRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        folderId: parseNumberOrNull(r['folderId']),
        name: parseString(r['name']),
        createdDate: parseString(r['createdDate']),
        modifiedDate: parseString(r['modifiedDate']),
        preview: parseString(r['preview']),
        htmlContent: parseString(r['htmlContent']),
        plainContent: parseString(r['plainContent']),
    }));
}

/**
 * Parses AppleScript output into a single note row, or null if none found.
 * @param output - Raw osascript stdout.
 * @returns The first note row, or null.
 */
export function parseNote(output: string): AppleScriptNoteRow | null {
    const notes = parseNotes(output);
    return notes[0] ?? null;
}

/**
 * Extracts an integer count from raw AppleScript output.
 * @param output - Raw osascript stdout containing a single number.
 * @returns The parsed count, or 0 if unparseable.
 */
export function parseCount(output: string): number {
    const trimmed = output.trim();
    const num = parseInt(trimmed, 10);
    return isNaN(num) ? 0 : num;
}

// ---------------------------------------------------------------------------
// Account parsers
// ---------------------------------------------------------------------------

/**
 * Parses AppleScript output into an array of Outlook account rows.
 * @param output - Raw osascript stdout.
 * @returns Typed account rows with id, name, email, and type.
 */
export function parseAccounts(output: string): AppleScriptAccountRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        name: parseString(r['name']),
        email: parseString(r['email']),
        type: r['type'] ?? 'exchange',
    }));
}

/**
 * Extracts the default account ID from AppleScript output.
 * @param output - Raw osascript stdout in "id=<number>" format.
 * @returns The account ID, or null if the output indicates an error.
 */
export function parseDefaultAccountId(output: string): number | null {
    const trimmed = output.trim();
    if (trimmed.startsWith('error')) {
        return null;
    }
    const parts = trimmed.split(DELIMITERS.EQUALS);
    if (parts[0] === 'id' && parts[1] !== undefined) {
        return parseNumberOrNull(parts[1]);
    }
    return null;
}

// ---------------------------------------------------------------------------
// Mutation result parsers — interpret outcomes of write operations
// ---------------------------------------------------------------------------

/**
 * Parses the result of a create-event AppleScript call.
 * @param output - Raw osascript stdout.
 * @returns Object with the new event's id and calendarId, or null on failure.
 */
export function parseCreateEventResult(output: string): {
    id: number;
    calendarId: number | null;
} | null {
    const records = parseRawOutput(output);
    if (records.length === 0)
        return null;
    const r = records[0];
    const id = parseNumber(r['id']);
    if (id === 0)
        return null;
    return {
        id,
        calendarId: parseNumberOrNull(r['calendarId']),
    };
}

/**
 * Parses AppleScript output into folder rows that include account ownership and message counts.
 * @param output - Raw osascript stdout.
 * @returns Typed folder rows extended with accountId and messageCount.
 */
export function parseFoldersWithAccount(output: string): AppleScriptFolderWithAccountRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        id: parseNumber(r['id']),
        name: parseString(r['name']),
        unreadCount: parseNumber(r['unreadCount']),
        messageCount: parseNumber(r['messageCount']),
        accountId: parseNumber(r['accountId']),
    }));
}

/**
 * Parses the result of an event RSVP response AppleScript call.
 * @param output - Raw osascript stdout.
 * @returns Structured result with success status and eventId or error, or null on empty output.
 */
export function parseRespondToEventResult(output: string): RespondToEventResult | null {
    const records = parseRawOutput(output);
    if (records.length === 0)
        return null;
    const record = records[0];
    if (!record)
        return null;
    const success = record['success'] === 'true';
    if (success) {
        return {
            success: true,
            eventId: parseNumber(record['eventId']),
        };
    }
    else {
        return {
            success: false,
            error: record['error'] ?? 'Unknown error',
        };
    }
}

/**
 * Parses the result of a delete-event AppleScript call.
 * @param output - Raw osascript stdout.
 * @returns Structured result with success status and eventId or error, or null on empty output.
 */
export function parseDeleteEventResult(output: string): DeleteEventResult | null {
    const records = parseRawOutput(output);
    if (records.length === 0)
        return null;
    const record = records[0];
    if (!record)
        return null;
    const success = record['success'] === 'true';
    if (success) {
        return {
            success: true,
            eventId: parseNumber(record['eventId']),
        };
    }
    else {
        return {
            success: false,
            error: record['error'] ?? 'Unknown error',
        };
    }
}

/**
 * Parses the result of an update-event AppleScript call.
 * @param output - Raw osascript stdout.
 * @returns Structured result with success status, updated field names, or error. Null on empty output.
 */
export function parseUpdateEventResult(output: string): UpdateEventResult | null {
    const records = parseRawOutput(output);
    if (records.length === 0)
        return null;
    const record = records[0];
    if (!record)
        return null;
    const success = record['success'] === 'true';
    if (success) {
        const fieldsStr = record['updatedFields'] ?? '';
        const fields = fieldsStr.length > 0 ? fieldsStr.split(',') : [];
        return {
            success: true,
            id: parseNumber(record['eventId']),
            updatedFields: fields,
        };
    }
    else {
        return {
            success: false,
            error: record['error'] ?? 'Unknown error',
        };
    }
}

/**
 * Parses the result of a send-email AppleScript call.
 * @param output - Raw osascript stdout.
 * @returns Structured result with success status and messageId/sentAt or error. Null on empty output.
 */
export function parseSendEmailResult(output: string): SendEmailResult | null {
    const records = parseRawOutput(output);
    if (records.length === 0)
        return null;
    const record = records[0];
    if (!record)
        return null;
    const success = record['success'] === 'true';
    if (success) {
        return {
            success: true,
            messageId: record['messageId'] ?? '',
            sentAt: record['sentAt'] ?? '',
        };
    }
    else {
        return {
            success: false,
            error: record['error'] ?? 'Unknown error',
        };
    }
}

/**
 * Parses AppleScript output into an array of attachment metadata rows.
 * @param output - Raw osascript stdout.
 * @returns Typed attachment rows with index, name, size, and content type.
 */
export function parseAttachments(output: string): AppleScriptAttachmentRow[] {
    const records = parseRawOutput(output);
    return records.map((r) => ({
        index: parseNumber(r['index']),
        name: r['name'] ?? '',
        fileSize: parseNumber(r['fileSize']),
        contentType: r['contentType'] ?? 'application/octet-stream',
    }));
}

/**
 * Parses the result of a save-attachment AppleScript call.
 * @param output - Raw osascript stdout.
 * @returns Structured result with file name, path, and size on success, or error. Null on empty output.
 */
export function parseSaveAttachmentResult(output: string): SaveAttachmentResult | null {
    const records = parseRawOutput(output);
    if (records.length === 0)
        return null;
    const record = records[0];
    if (!record)
        return null;
    const success = record['success'] === 'true';
    if (success) {
        return {
            success: true,
            name: record['name'] ?? '',
            savedTo: record['savedTo'] ?? '',
            fileSize: parseNumber(record['fileSize']),
        };
    }
    else {
        return {
            success: false,
            error: record['error'] ?? 'Unknown error',
        };
    }
}
