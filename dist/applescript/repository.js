import { executeAppleScript, executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import * as parser from './parser.js';
import { appleTimestampToIso, isoToAppleTimestamp } from '../utils/dates.js';
import { createEmailPath, createEventPath, createContactPath, createTaskPath, createNotePath, } from './content-readers.js';
import { TtlCache } from './cache.js';
import { isExpensiveOperationAllowed, currentMedianLatency, clearWindow, recordLatency } from './throttle.js';
import { OutlookBridgeStressedError, OutlookQueryRefusedError } from '../utils/errors.js';
/**
 * Cheap "ground-truth" query for confirming bridge health. Same shape as the
 * watchdog script's probe — touches Outlook's message store (not just app
 * metadata), so it actually exercises the path that degrades. On a healthy
 * bridge it returns in <200ms; on a stuck bridge it times out at -1712.
 */
const GROUND_TRUTH_PROBE_QUERY = 'tell application "Microsoft Outlook" to count of messages of outbox';
const GROUND_TRUTH_PROBE_TIMEOUT_MS = 2000;
/** Probe latencies under this threshold count as "bridge is actually healthy". */
const GROUND_TRUTH_PROBE_HEALTHY_MS = 500;
// Default search folders cache (60 min TTL). The user's primary inbox and
// sent-items folder IDs are effectively static within a session; resolving
// them once via AppleScript and reusing avoids a per-call lookup.
const DEFAULT_SEARCH_FOLDERS_TTL_MS = 60 * 60 * 1000;
const DEFAULT_SEARCH_FOLDERS_CACHE_KEY = 'ids';
/**
 * Per-folder message-count safeguard. A `searchEmails`/`searchEmailsInFolder`
 * call against a folder larger than this is empirically dangerous: Outlook's
 * per-message `whose subject contains` scan can time out at -1712 and leave
 * the bridge degraded (verified 2026-05-12). Date filters bound the scan
 * enough to be safe, so we allow >50k searches when `after` or `before` is
 * provided.
 */
const LARGE_FOLDER_THRESHOLD = 50_000;
/**
 * Refuse expensive operations when the bridge is degraded — but verify against
 * a live ground-truth probe before refusing, so we don't false-positive on
 * stale evidence.
 *
 * The rolling median can stay "degraded" for up to 5 minutes after a single
 * heavy `whose` query (which is naturally slow on large folders) inflates the
 * window. The independent watchdog probe meanwhile shows the bridge is fine
 * because it queries cheap metadata. Without the live re-check here, expensive
 * calls would be refused for minutes based on already-disproved evidence
 * (verified 2026-05-13 in the rewards_advisor session — MCP refused with
 * 29.8s rolling median, watchdog probe at the same instant returned 170ms).
 *
 * Flow:
 * 1. Median says healthy/normal/stressed → allow immediately (fast path)
 * 2. Median says degraded → run a cheap `count of messages of outbox` probe
 *    a. Probe fast (<500ms) → bridge is actually healthy; clear stale window
 *       and allow the operation. Seed with the fresh probe sample so future
 *       classifications use post-recovery data.
 *    b. Probe slow or errored → degradation is real; refuse with both numbers
 *       in the error so the caller knows we cross-checked.
 */
function gateExpensive(operation) {
    if (isExpensiveOperationAllowed())
        return;
    const probeStart = Date.now();
    const probe = executeAppleScript(GROUND_TRUTH_PROBE_QUERY, {
        timeoutMs: GROUND_TRUTH_PROBE_TIMEOUT_MS,
        skipThrottle: true,
    });
    const probeMs = Date.now() - probeStart;
    if (probe.success && probeMs < GROUND_TRUTH_PROBE_HEALTHY_MS) {
        // Bridge is genuinely healthy. The window held stale evidence
        // (probably one or two slow `whose` queries). Reset it and seed
        // with the fresh probe sample so future classifications use
        // post-recovery data.
        clearWindow();
        recordLatency(probeMs);
        return;
    }
    // Both rolling median AND live probe confirm degradation. Refuse.
    throw new OutlookBridgeStressedError(currentMedianLatency(), probeMs, probe.success, operation);
}
// Folder list cache (5 min TTL). Folders rarely change — extending from the
// previous 30s removes hundreds of redundant AppleEvents per typical session.
const FOLDER_LIST_TTL_MS = 5 * 60 * 1000;
const FOLDER_LIST_CACHE_KEY = 'all';
function priorityToNumber(priority) {
    switch (priority.toLowerCase()) {
        case 'high':
            return 1;
        case 'low':
            return -1;
        default:
            return 0;
    }
}
function toFolderRow(asFolder) {
    return {
        id: asFolder.id,
        name: asFolder.name,
        parentId: null,
        specialType: 0,
        folderType: 1,
        accountId: 1,
        messageCount: asFolder.messageCount,
        unreadCount: asFolder.unreadCount,
    };
}
function calendarToFolderRow(asCal) {
    return {
        id: asCal.id,
        name: asCal.name,
        parentId: null,
        specialType: 0,
        folderType: 2,
        accountId: 1,
        messageCount: 0,
        unreadCount: 0,
    };
}
function toEmailRow(asEmail) {
    return {
        id: asEmail.id,
        folderId: asEmail.folderId ?? 0,
        subject: asEmail.subject,
        sender: asEmail.senderName,
        senderAddress: asEmail.senderEmail,
        recipients: asEmail.toRecipients,
        displayTo: asEmail.toRecipients,
        toAddresses: asEmail.toRecipients,
        ccAddresses: asEmail.ccRecipients,
        preview: asEmail.preview,
        isRead: asEmail.isRead ? 1 : 0,
        timeReceived: isoToAppleTimestamp(asEmail.dateReceived),
        timeSent: isoToAppleTimestamp(asEmail.dateSent),
        hasAttachment: asEmail.attachments.length > 0 ? 1 : 0,
        size: 0,
        priority: priorityToNumber(asEmail.priority),
        flagStatus: asEmail.flagStatus ?? 0,
        categories: null,
        messageId: null,
        conversationId: null,
        dataFilePath: createEmailPath(asEmail.id),
    };
}
function toEventRow(asEvent) {
    return {
        id: asEvent.id,
        folderId: asEvent.calendarId ?? 0,
        startDate: isoToAppleTimestamp(asEvent.startTime),
        endDate: isoToAppleTimestamp(asEvent.endTime),
        isRecurring: asEvent.isRecurring ? 1 : 0,
        hasReminder: 0,
        attendeeCount: asEvent.attendees.length,
        uid: null,
        masterRecordId: null,
        recurrenceId: null,
        dataFilePath: createEventPath(asEvent.id),
    };
}
function toContactRow(asContact) {
    return {
        id: asContact.id,
        folderId: 0,
        displayName: asContact.displayName,
        sortName: asContact.lastName ?? asContact.displayName,
        contactType: null,
        dataFilePath: createContactPath(asContact.id),
    };
}
function toTaskRow(asTask) {
    return {
        id: asTask.id,
        folderId: asTask.folderId ?? 0,
        name: asTask.name,
        isCompleted: asTask.isCompleted ? 1 : 0,
        dueDate: isoToAppleTimestamp(asTask.dueDate),
        startDate: isoToAppleTimestamp(asTask.startDate),
        priority: priorityToNumber(asTask.priority),
        hasReminder: null,
        dataFilePath: createTaskPath(asTask.id),
    };
}
function toNoteRow(asNote) {
    return {
        id: asNote.id,
        folderId: asNote.folderId ?? 0,
        modifiedDate: isoToAppleTimestamp(asNote.modifiedDate),
        dataFilePath: createNotePath(asNote.id),
    };
}
/**
 * Deduplicate email rows by ID, preserving first occurrence.
 * searchMessages phases 1 and 2 may return overlapping results — this is
 * intentional: doing dedup in TypeScript (O(n) via Set) instead of AppleScript
 * (O(n × offset) via list scans) eliminates the main performance bottleneck.
 */
export function deduplicateEmailRows(rows) {
    const seen = new Set();
    return rows.filter(r => {
        if (seen.has(r.id))
            return false;
        seen.add(r.id);
        return true;
    });
}
/**
 * Case-insensitive containment check against subject, sender display name, and
 * sender email. Used by the date-first search path to filter results in TS
 * after AppleScript has done the cheap date predicate. Matches the original
 * Phase 1 + Phase 2 semantics (subject match OR sender match) in a single pass.
 */
export function applySubjectSenderFilter(rows, query) {
    const needle = query.toLowerCase();
    return rows.filter(r => {
        const subject = (r.subject ?? '').toLowerCase();
        const sender = (r.sender ?? '').toLowerCase();
        const senderAddress = (r.senderAddress ?? '').toLowerCase();
        return subject.includes(needle) || sender.includes(needle) || senderAddress.includes(needle);
    });
}
/**
 * Calculate timeout for search operations, scaling with offset.
 *
 * Layering (outermost to innermost):
 *   - MCP client (Claude Code): typically 90–120s for tool calls
 *   - Node side (this function):  65–80s
 *   - AppleScript inner block:    55s (`with timeout of 55 seconds` in scripts.ts)
 *
 * AppleScript's inner timeout MUST fire before Node's, otherwise Node SIGKILLs
 * osascript mid-Apple-Event and corrupts Outlook's scripting bridge, causing
 * cascading -609 errors on subsequent calls. The 10s gap between Node (65s)
 * and AS-internal (55s) is the safety margin.
 *
 * Why 55s for AS-internal: Outlook's first heavy AppleScript call after idle
 * (cold start) takes 30–45s while it pages in the subject index. 55s comfortably
 * covers this so cold-start searches succeed instead of timing out.
 */
export function searchTimeoutMs(offset) {
    return Math.min(80000, 65000 + Math.floor(offset / 25) * 5000);
}
export class AppleScriptRepository {
    folderListCache = new TtlCache(FOLDER_LIST_TTL_MS);
    defaultSearchFoldersCache = new TtlCache(DEFAULT_SEARCH_FOLDERS_TTL_MS);
    /**
     * Resolves and caches the IDs of the user's primary INBOX and Sent Items
     * folders via `inbox of default account` / `folder "Sent Items" of default
     * account`. These are the folders that an unscoped `searchEmails` call
     * defaults to — Archive is deliberately excluded per user policy.
     */
    getDefaultSearchFolders() {
        const cached = this.defaultSearchFoldersCache.get(DEFAULT_SEARCH_FOLDERS_CACHE_KEY);
        if (cached != null)
            return cached;
        const output = executeAppleScriptOrThrow(scripts.GET_DEFAULT_SEARCH_FOLDER_IDS, { timeoutMs: 8000 });
        const inboxMatch = output.match(/inbox=(\d+)/);
        const sentMatch = output.match(/sentItems=(\d+)?/);
        if (inboxMatch == null) {
            throw new Error(`Failed to resolve default inbox from AppleScript: ${output}`);
        }
        const resolved = {
            inboxId: parseInt(inboxMatch[1], 10),
            sentItemsId: sentMatch && sentMatch[1] ? parseInt(sentMatch[1], 10) : null,
        };
        this.defaultSearchFoldersCache.set(DEFAULT_SEARCH_FOLDERS_CACHE_KEY, resolved);
        return resolved;
    }
    /**
     * Refuses a search against a folder larger than LARGE_FOLDER_THRESHOLD
     * UNLESS a date filter is provided (which bounds the scan enough to be
     * safe). Surfaces a structured OUTLOOK_QUERY_REFUSED with a concrete
     * remedy — pass `after` and/or `before`.
     */
    gateLargeFolderSearch(folderId, after, before, operation) {
        const hasDateFilter = after != null || before != null;
        if (hasDateFilter)
            return;
        const folder = this.listFolders().find(f => f.id === folderId);
        const count = folder?.messageCount ?? 0;
        if (count > LARGE_FOLDER_THRESHOLD) {
            throw new OutlookQueryRefusedError(operation, `folder "${folder?.name ?? folderId}" contains ${count.toLocaleString()} messages (>${LARGE_FOLDER_THRESHOLD.toLocaleString()} threshold) and unbounded search would risk destabilizing the Outlook AppleScript bridge`, `Pass \`after\` and/or \`before\` (ISO 8601) to bound the search by date — e.g., \`after: "2025-01-01T00:00:00Z"\` to search the last year. Date-bounded searches at this size are safe.`);
        }
    }
    listFolders() {
        const cached = this.folderListCache.get(FOLDER_LIST_CACHE_KEY);
        if (cached != null)
            return cached;
        const output = executeAppleScriptOrThrow(scripts.LIST_MAIL_FOLDERS);
        const folders = parser.parseFolders(output).map(toFolderRow);
        this.folderListCache.set(FOLDER_LIST_CACHE_KEY, folders);
        return folders;
    }
    getFolder(id) {
        return this.listFolders().find((f) => f.id === id);
    }
    listEmails(folderId, limit, offset, after, before) {
        gateExpensive(`listEmails(folder=${folderId}, limit=${limit})`);
        const script = scripts.listMessages(folderId, limit, offset, false, after, before);
        const timeoutMs = (after != null || before != null) ? 65000 : 50000;
        const output = executeAppleScriptOrThrow(script, { timeoutMs });
        return parser.parseEmails(output).map(toEmailRow);
    }
    listUnreadEmails(folderId, limit, offset, after, before) {
        gateExpensive(`listUnreadEmails(folder=${folderId}, limit=${limit})`);
        const script = scripts.listMessages(folderId, limit, offset, true, after, before);
        const timeoutMs = (after != null || before != null) ? 65000 : 50000;
        const output = executeAppleScriptOrThrow(script, { timeoutMs });
        return parser.parseEmails(output).map(toEmailRow);
    }
    /**
     * Unscoped `searchEmails` auto-scopes to the user's primary INBOX AND
     * Sent Items folders. Archive is deliberately excluded — the user must
     * pass `folder_id` explicitly to search Archive (user policy 2026-05-12).
     *
     * Each folder is searched in a separate AppleScript call (so the throttle
     * paces them naturally); results are merged, dedup'd by id, and sorted
     * by `dateReceived` descending before truncating to `limit`.
     *
     * The 50k safeguard applies to each folder individually — if either is
     * >50k AND no date filter is provided, OUTLOOK_QUERY_REFUSED is thrown.
     */
    searchEmails(query, limit, offset, after, before, includeBodySearch) {
        const { inboxId, sentItemsId } = this.getDefaultSearchFolders();
        const folders = sentItemsId != null ? [inboxId, sentItemsId] : [inboxId];
        const dateFirstMode = after != null || before != null;
        const merged = [];
        for (const folderId of folders) {
            this.gateLargeFolderSearch(folderId, after, before, `searchEmails(query="${query}") on default folder ${folderId}`);
            gateExpensive(`searchEmails(query="${query}", folder=${folderId}, limit=${limit})`);
            // In date-first mode the AppleScript template does NOT apply offset/limit
            // (it walks all date-matched messages up to DATE_FIRST_MAX_SCAN), so pass
            // 0/0 here — TS applies pagination after subject/sender filtering.
            const asLimit = dateFirstMode ? 0 : limit;
            const asOffset = dateFirstMode ? 0 : offset;
            const script = scripts.searchMessages(query, folderId, asLimit, asOffset, after, before, includeBodySearch);
            const output = executeAppleScriptOrThrow(script, { timeoutMs: searchTimeoutMs(offset) });
            let rows = parser.parseEmails(output).map(toEmailRow);
            if (dateFirstMode) {
                rows = applySubjectSenderFilter(rows, query);
            }
            merged.push(...rows);
        }
        const deduped = deduplicateEmailRows(merged);
        deduped.sort((a, b) => (Number(b.timeReceived) || 0) - (Number(a.timeReceived) || 0));
        // Apply offset+limit here: in date-first mode we deferred pagination from
        // AppleScript; in subject-first mode AppleScript already applied them but
        // slicing to `limit` is harmless if we're already at or under limit.
        return deduped.slice(offset, offset + limit);
    }
    searchEmailsInFolder(folderId, query, limit, offset, after, before, includeBodySearch) {
        this.gateLargeFolderSearch(folderId, after, before, `searchEmailsInFolder(folder=${folderId}, query="${query}")`);
        gateExpensive(`searchEmailsInFolder(folder=${folderId}, query="${query}", limit=${limit})`);
        const dateFirstMode = after != null || before != null;
        const asLimit = dateFirstMode ? 0 : limit;
        const asOffset = dateFirstMode ? 0 : offset;
        const script = scripts.searchMessages(query, folderId, asLimit, asOffset, after, before, includeBodySearch);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: searchTimeoutMs(offset) });
        let rows = parser.parseEmails(output).map(toEmailRow);
        if (dateFirstMode) {
            rows = applySubjectSenderFilter(rows, query);
            rows.sort((a, b) => (Number(b.timeReceived) || 0) - (Number(a.timeReceived) || 0));
            rows = rows.slice(offset, offset + limit);
        }
        return deduplicateEmailRows(rows);
    }
    getEmail(id) {
        try {
            const script = scripts.getMessage(id);
            const output = executeAppleScriptOrThrow(script);
            const email = parser.parseEmail(output);
            return email != null ? toEmailRow(email) : undefined;
        }
        catch {
            return undefined;
        }
    }
    getUnreadCount() {
        const folders = this.listFolders();
        return folders.reduce((sum, f) => sum + f.unreadCount, 0);
    }
    getUnreadCountByFolder(folderId) {
        try {
            const script = scripts.getUnreadCount(folderId);
            const output = executeAppleScriptOrThrow(script);
            return parser.parseCount(output);
        }
        catch {
            return 0;
        }
    }
    listCalendars() {
        const output = executeAppleScriptOrThrow(scripts.LIST_CALENDARS);
        return parser.parseCalendars(output).map(calendarToFolderRow);
    }
    listEvents(limit, offset = 0) {
        const script = scripts.listEvents(null, null, null, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseEvents(output).map(toEventRow);
    }
    listEventsByFolder(folderId, limit, offset = 0) {
        const script = scripts.listEvents(folderId, null, null, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseEvents(output).map(toEventRow);
    }
    listEventsByDateRange(startDate, endDate, limit, offset = 0) {
        // Server-side date filtering via AppleScript whose clause
        const startIso = appleTimestampToIso(startDate);
        const endIso = appleTimestampToIso(endDate);
        if (startIso == null || endIso == null) {
            return this.listEvents(limit, offset);
        }
        const script = scripts.listEvents(null, startIso, endIso, limit, offset);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 60000 });
        return parser.parseEvents(output).map(toEventRow);
    }
    getEvent(id) {
        try {
            const script = scripts.getEvent(id);
            const output = executeAppleScriptOrThrow(script);
            const event = parser.parseEvent(output);
            return event != null ? toEventRow(event) : undefined;
        }
        catch {
            return undefined;
        }
    }
    listContacts(limit, offset) {
        const script = scripts.listContacts(limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseContacts(output).map(toContactRow);
    }
    searchContacts(query, limit, offset = 0) {
        const script = scripts.searchContacts(query, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseContacts(output).map(toContactRow);
    }
    getContact(id) {
        try {
            const script = scripts.getContact(id);
            const output = executeAppleScriptOrThrow(script);
            const contact = parser.parseContact(output);
            return contact != null ? toContactRow(contact) : undefined;
        }
        catch {
            return undefined;
        }
    }
    listTasks(limit, offset) {
        const script = scripts.listTasks(limit, offset, true);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }
    listIncompleteTasks(limit, offset) {
        const script = scripts.listTasks(limit, offset, false);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }
    searchTasks(query, limit, offset = 0) {
        const script = scripts.searchTasks(query, limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseTasks(output).map(toTaskRow);
    }
    getTask(id) {
        try {
            const script = scripts.getTask(id);
            const output = executeAppleScriptOrThrow(script);
            const task = parser.parseTask(output);
            return task != null ? toTaskRow(task) : undefined;
        }
        catch {
            return undefined;
        }
    }
    listNotes(limit, offset) {
        const script = scripts.listNotes(limit, offset);
        const output = executeAppleScriptOrThrow(script);
        return parser.parseNotes(output).map(toNoteRow);
    }
    searchNotes(query, limit, offset = 0) {
        const script = scripts.searchNotes(query, limit, offset);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 30000 });
        return parser.parseNotes(output).map(toNoteRow);
    }
    searchEvents(query, limit, offset = 0, after, before) {
        const script = scripts.searchEvents(query, limit, offset, after, before);
        const output = executeAppleScriptOrThrow(script, { timeoutMs: 30000 });
        return parser.parseEvents(output).map(toEventRow);
    }
    getNote(id) {
        try {
            const script = scripts.getNote(id);
            const output = executeAppleScriptOrThrow(script);
            const note = parser.parseNote(output);
            return note != null ? toNoteRow(note) : undefined;
        }
        catch {
            return undefined;
        }
    }
    moveEmail(emailId, destinationFolderId) {
        const script = scripts.moveMessage(emailId, destinationFolderId);
        executeAppleScriptOrThrow(script);
    }
    deleteEmail(emailId) {
        const script = scripts.deleteMessage(emailId);
        executeAppleScriptOrThrow(script);
    }
    archiveEmail(emailId) {
        const script = scripts.archiveMessage(emailId);
        executeAppleScriptOrThrow(script);
    }
    junkEmail(emailId) {
        const script = scripts.junkMessage(emailId);
        executeAppleScriptOrThrow(script);
    }
    markEmailRead(emailId, isRead) {
        const script = scripts.setMessageReadStatus(emailId, isRead);
        executeAppleScriptOrThrow(script);
    }
    setEmailFlag(emailId, flagStatus) {
        const script = scripts.setMessageFlag(emailId, flagStatus);
        executeAppleScriptOrThrow(script);
    }
    setEmailCategories(emailId, categories) {
        const script = scripts.setMessageCategories(emailId, categories);
        executeAppleScriptOrThrow(script);
    }
    createFolder(name, parentFolderId) {
        const script = scripts.createMailFolder(name, parentFolderId);
        const output = executeAppleScriptOrThrow(script);
        const newFolderId = parseInt(output.trim(), 10);
        this.folderListCache.invalidateAll();
        return {
            id: newFolderId,
            name,
            parentId: parentFolderId ?? null,
            specialType: 0,
            folderType: 1,
            accountId: 1,
            messageCount: 0,
            unreadCount: 0,
        };
    }
    deleteFolder(folderId) {
        const script = scripts.deleteMailFolder(folderId);
        executeAppleScriptOrThrow(script);
        this.folderListCache.invalidateAll();
    }
    renameFolder(folderId, newName) {
        const script = scripts.renameMailFolder(folderId, newName);
        executeAppleScriptOrThrow(script);
        this.folderListCache.invalidateAll();
    }
    moveFolder(folderId, destinationParentId) {
        const script = scripts.moveMailFolder(folderId, destinationParentId);
        executeAppleScriptOrThrow(script);
        this.folderListCache.invalidateAll();
    }
    emptyFolder(folderId) {
        const script = scripts.emptyMailFolder(folderId);
        executeAppleScriptOrThrow(script);
    }
}
export function createAppleScriptRepository() {
    return new AppleScriptRepository();
}
