/**
 * Domain types for Outlook notes: summaries and full note records.
 */

/** Lightweight note representation used in list and search results. */
export interface NoteSummary {
    readonly id: number;
    readonly folderId: number;
    readonly title: string | null;
    readonly preview: string | null;
    readonly modifiedDate: string | null;
}

/** Complete note record including body content, creation date, and categories. */
export interface Note extends NoteSummary {
    readonly body: string | null;
    readonly createdDate: string | null;
    readonly categories: readonly string[];
}
