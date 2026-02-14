import { z } from 'zod';
import type { IRepository } from '../database/repository.js';
import type { ContactSummary, Contact, EmailType, PhoneType, AddressType, PaginatedResult } from '../types/index.js';
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

/** Validated parameters for listing contacts with pagination. */
export type ListContactsParams = z.infer<typeof ListContactsInput>;
/** Validated parameters for searching contacts by name. */
export type SearchContactsParams = z.infer<typeof SearchContactsInput>;
/** Validated parameters for retrieving a single contact. */
export type GetContactParams = z.infer<typeof GetContactInput>;

/** Reads rich contact details (name fields, emails, phones, addresses, notes) from a data file. */
export interface IContactContentReader {
    readContactDetails(dataFilePath: string | null): ContactDetails | null;
}

/** Rich contact details extracted from a contact's data file. */
export interface ContactDetails {
    readonly firstName: string | null;
    readonly lastName: string | null;
    readonly middleName: string | null;
    readonly nickname: string | null;
    readonly company: string | null;
    readonly jobTitle: string | null;
    readonly department: string | null;
    readonly emails: readonly { type: string; address: string }[];
    readonly phones: readonly { type: string; number: string }[];
    readonly addresses: readonly {
        type: string;
        street: string | null;
        city: string | null;
        state: string | null;
        postalCode: string | null;
        country: string | null;
    }[];
    readonly notes: string | null;
}

/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullContactContentReader: IContactContentReader = {
    readContactDetails: () => null,
};

// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------

/** Converts a raw contact repository row into a ContactSummary domain object. */
function transformContactSummary(row: ReturnType<IRepository['getContact']> & {}): ContactSummary {
    return {
        id: row.id,
        folderId: row.folderId,
        displayName: row.displayName,
        sortName: row.sortName,
        contactType: (row.contactType ?? 0) as ContactSummary['contactType'],
    };
}

/** Converts a raw contact repository row and its rich details into a full Contact domain object. */
function transformContact(row: ReturnType<IRepository['getContact']> & {}, details: ContactDetails | null): Contact {
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
        emails: details?.emails?.map((e) => ({ type: e.type as EmailType, address: e.address })) ?? [],
        phones: details?.phones?.map((p) => ({ type: p.type as PhoneType, number: p.number })) ?? [],
        addresses: details?.addresses?.map((a) => ({
            type: a.type as AddressType,
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
    private readonly repository: IRepository;
    private readonly contentReader: IContactContentReader;

    constructor(repository: IRepository, contentReader: IContactContentReader = nullContactContentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
    }

    /** Returns a paginated list of contact summaries. */
    listContacts(params: ListContactsParams): PaginatedResult<ContactSummary> {
        const { limit, offset } = params;
        const rows = this.repository.listContacts(limit + 1, offset);
        return paginate(rows.map(transformContactSummary), limit);
    }

    /** Searches contacts by name and returns matching summaries up to the given limit. */
    searchContacts(params: SearchContactsParams): PaginatedResult<ContactSummary> {
        const { query, limit, offset } = params;
        const rows = this.repository.searchContacts(query, limit + 1, offset);
        return paginate(rows.map(transformContactSummary), limit);
    }

    /** Retrieves a single contact by ID with full details, or null if not found. */
    getContact(params: GetContactParams): Contact | null {
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
export function createContactsTools(repository: IRepository, contentReader: IContactContentReader = nullContactContentReader): ContactsTools {
    return new ContactsTools(repository, contentReader);
}
