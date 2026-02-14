import { z } from 'zod';
import { paginate } from '../types/index.js';
// ---------------------------------------------------------------------------
// Zod input schemas for contact MCP tools
// ---------------------------------------------------------------------------
export const ListContactsInput = z.strictObject({
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of contacts to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of contacts to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
});
export const SearchContactsInput = z.strictObject({
    query: z.string().min(1).describe('Search query text matched against contact names (e.g., "John")'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of contacts to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of contacts to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
});
export const GetContactInput = z.strictObject({
    contact_id: z.number().int().positive().describe('The contact ID to retrieve (e.g., from list_contacts or search_contacts)'),
});
/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullContactContentReader = {
    readContactDetails: () => null,
};
// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------
/** Converts a raw contact repository row into a ContactSummary domain object. */
function transformContactSummary(row) {
    return {
        id: row.id,
        folderId: row.folderId,
        displayName: row.displayName,
        sortName: row.sortName,
        contactType: (row.contactType ?? 0),
    };
}
/** Converts a raw contact repository row and its rich details into a full Contact domain object. */
function transformContact(row, details) {
    const summary = transformContactSummary(row);
    return {
        ...summary,
        firstName: details?.firstName ?? null,
        lastName: details?.lastName ?? null,
        middleName: details?.middleName ?? null,
        nickname: details?.nickname ?? null,
        company: details?.company ?? null,
        jobTitle: details?.jobTitle ?? null,
        department: details?.department ?? null,
        emails: details?.emails?.map((e) => ({ type: e.type, address: e.address })) ?? [],
        phones: details?.phones?.map((p) => ({ type: p.type, number: p.number })) ?? [],
        addresses: details?.addresses?.map((a) => ({
            type: a.type,
            street: a.street,
            city: a.city,
            state: a.state,
            postalCode: a.postalCode,
            country: a.country,
        })) ?? [],
        notes: details?.notes ?? null,
    };
}
// ---------------------------------------------------------------------------
// ContactsTools -- provides read operations for Outlook contacts
// ---------------------------------------------------------------------------
/** Exposes contact read operations backed by a repository and an optional content reader. */
export class ContactsTools {
    repository;
    contentReader;
    constructor(repository, contentReader = nullContactContentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
    }
    /** Returns a paginated list of contact summaries. */
    listContacts(params) {
        const { limit, offset } = params;
        const rows = this.repository.listContacts(limit + 1, offset);
        return paginate(rows.map(transformContactSummary), limit);
    }
    /** Searches contacts by name and returns matching summaries up to the given limit. */
    searchContacts(params) {
        const { query, limit, offset } = params;
        const rows = this.repository.searchContacts(query, limit + 1, offset);
        return paginate(rows.map(transformContactSummary), limit);
    }
    /** Retrieves a single contact by ID with full details, or null if not found. */
    getContact(params) {
        const { contact_id } = params;
        const row = this.repository.getContact(contact_id);
        if (row == null) {
            return null;
        }
        const details = this.contentReader.readContactDetails(row.dataFilePath);
        return transformContact(row, details);
    }
}
/** Factory that creates a ContactsTools instance with the given repository and optional content reader. */
export function createContactsTools(repository, contentReader = nullContactContentReader) {
    return new ContactsTools(repository, contentReader);
}
