import { z } from 'zod';
import type { IRepository } from '../database/repository.js';
import type { NoteSummary, Note, PaginatedResult } from '../types/index.js';
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

/** Validated parameters for listing notes with pagination. */
export type ListNotesParams = z.infer<typeof ListNotesInput>;
/** Validated parameters for retrieving a single note. */
export type GetNoteParams = z.infer<typeof GetNoteInput>;
/** Validated parameters for searching notes by title/name. */
export type SearchNotesParams = z.infer<typeof SearchNotesInput>;

/** Reads rich note details (title, body, preview, created date, categories) from a data file. */
export interface INoteContentReader {
    readNoteDetails(dataFilePath: string | null): NoteDetails | null;
}

/** Rich note details extracted from a note's data file. */
export interface NoteDetails {
    readonly title: string | null;
    readonly body: string | null;
    readonly preview: string | null;
    readonly createdDate: string | null;
    readonly categories: readonly string[];
}

/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullNoteContentReader: INoteContentReader = {
    readNoteDetails: () => null,
};

// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------

/** Converts a raw note repository row and its rich details into a NoteSummary domain object. */
function transformNoteSummary(row: ReturnType<IRepository['getNote']> & {}, details: NoteDetails | null): NoteSummary {
    return {
        id: row.id,
        folderId: row.folderId,
        title: details?.title ?? null,
        preview: details?.preview ?? null,
        modifiedDate: appleTimestampToIso(row.modifiedDate),
    };
}

/** Converts a raw note repository row and its rich details into a full Note domain object. */
function transformNote(row: ReturnType<IRepository['getNote']> & {}, details: NoteDetails | null): Note {
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
    private readonly repository: IRepository;
    private readonly contentReader: INoteContentReader;

    constructor(repository: IRepository, contentReader: INoteContentReader = nullNoteContentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
    }

    /** Returns a paginated list of note summaries with titles and previews. */
    listNotes(params: ListNotesParams): PaginatedResult<NoteSummary> {
        const { limit, offset } = params;
        const rows = this.repository.listNotes(limit + 1, offset);
        const items = rows.map((row) => {
            const details = this.contentReader.readNoteDetails(row.dataFilePath);
            return transformNoteSummary(row, details);
        });
        return paginate(items, limit);
    }

    /** Retrieves a single note by ID with full details, or null if not found. */
    getNote(params: GetNoteParams): Note | null {
        const { note_id } = params;
        const row = this.repository.getNote(note_id);
        if (row == null) {
            return null;
        }
        const details = this.contentReader.readNoteDetails(row.dataFilePath);
        return transformNote(row, details);
    }

    /** Searches notes by title/name and returns matching summaries up to the given limit. */
    searchNotes(params: SearchNotesParams): PaginatedResult<NoteSummary> {
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
export function createNotesTools(repository: IRepository, contentReader: INoteContentReader = nullNoteContentReader): NotesTools {
    return new NotesTools(repository, contentReader);
}
