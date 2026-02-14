import { z } from 'zod';
import { paginate } from '../types/index.js';
import { appleTimestampToIso } from '../utils/dates.js';
// ---------------------------------------------------------------------------
// Zod input schemas for note MCP tools
// ---------------------------------------------------------------------------
export const ListNotesInput = z.strictObject({
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of notes to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of notes to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
});
export const GetNoteInput = z.strictObject({
    note_id: z.number().int().positive().describe('The note ID to retrieve (e.g., from list_notes or search_notes)'),
});
export const SearchNotesInput = z.strictObject({
    query: z.string().min(1).describe('Search query text matched against note titles/names only — does not search body content (e.g., "meeting notes")'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of notes to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of notes to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
});
/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullNoteContentReader = {
    readNoteDetails: () => null,
};
// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------
/** Converts a raw note repository row and its rich details into a NoteSummary domain object. */
function transformNoteSummary(row, details) {
    return {
        id: row.id,
        folderId: row.folderId,
        title: details?.title ?? null,
        preview: details?.preview ?? null,
        modifiedDate: appleTimestampToIso(row.modifiedDate),
    };
}
/** Converts a raw note repository row and its rich details into a full Note domain object. */
function transformNote(row, details) {
    const summary = transformNoteSummary(row, details);
    return {
        ...summary,
        body: details?.body ?? null,
        createdDate: details?.createdDate ?? null,
        categories: details?.categories ?? [],
    };
}
// ---------------------------------------------------------------------------
// NotesTools -- provides read operations for Outlook notes
// ---------------------------------------------------------------------------
/** Exposes note read operations backed by a repository and an optional content reader. */
export class NotesTools {
    repository;
    contentReader;
    constructor(repository, contentReader = nullNoteContentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
    }
    /** Returns a paginated list of note summaries with titles and previews. */
    listNotes(params) {
        const { limit, offset } = params;
        const rows = this.repository.listNotes(limit + 1, offset);
        const items = rows.map((row) => {
            const details = this.contentReader.readNoteDetails(row.dataFilePath);
            return transformNoteSummary(row, details);
        });
        return paginate(items, limit);
    }
    /** Retrieves a single note by ID with full details, or null if not found. */
    getNote(params) {
        const { note_id } = params;
        const row = this.repository.getNote(note_id);
        if (row == null) {
            return null;
        }
        const details = this.contentReader.readNoteDetails(row.dataFilePath);
        return transformNote(row, details);
    }
    /** Searches notes by title/name and returns matching summaries up to the given limit. */
    searchNotes(params) {
        const { query, limit, offset } = params;
        const rows = this.repository.searchNotes(query, limit + 1, offset);
        const items = rows.map(row => {
            const details = this.contentReader.readNoteDetails(row.dataFilePath);
            return transformNoteSummary(row, details);
        });
        return paginate(items, limit);
    }
}
/** Factory that creates a NotesTools instance with the given repository and optional content reader. */
export function createNotesTools(repository, contentReader = nullNoteContentReader) {
    return new NotesTools(repository, contentReader);
}
