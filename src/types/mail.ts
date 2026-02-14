/**
 * Domain types for Outlook mail: folders, messages, attachments, and flags.
 */

/** Numeric identifiers for Outlook's built-in special folders. */
export const SpecialFolderType = {
    Inbox: 1,
    Outbox: 2,
    Calendar: 4,
    Sent: 8,
    Deleted: 9,
    Drafts: 10,
    Junk: 12,
} as const;
export type SpecialFolderTypeValue = (typeof SpecialFolderType)[keyof typeof SpecialFolderType];

/** Numeric priority levels matching Outlook's internal representation. */
export const Priority = {
    High: 1,
    Normal: 3,
    Low: 5,
} as const;
export type PriorityValue = (typeof Priority)[keyof typeof Priority];

/** Numeric flag states for email follow-up tracking. */
export const FlagStatus = {
    None: 0,
    Flagged: 1,
    Completed: 2,
} as const;
export type FlagStatusValue = (typeof FlagStatus)[keyof typeof FlagStatus];

/** A mail folder with its message and unread counts. */
export interface Folder {
    readonly id: number;
    readonly name: string;
    readonly parentId: number | null;
    readonly specialType: number;
    readonly folderType: number;
    readonly accountId: number;
    readonly messageCount: number;
    readonly unreadCount: number;
}

/** Lightweight email representation used in list and search results. */
export interface EmailSummary {
    readonly id: number;
    readonly folderId: number;
    readonly subject: string | null;
    readonly sender: string | null;
    readonly senderAddress: string | null;
    readonly preview: string | null;
    readonly isRead: boolean;
    readonly timeReceived: string | null;
    readonly timeSent: string | null;
    readonly hasAttachment: boolean;
    readonly priority: PriorityValue;
    readonly flagStatus: FlagStatusValue;
    readonly categories: readonly string[];
}

/** Complete email record including body content and recipient details. */
export interface Email extends EmailSummary {
    readonly recipients: string | null;
    readonly displayTo: string | null;
    readonly toAddresses: string | null;
    readonly ccAddresses: string | null;
    readonly size: number;
    readonly messageId: string | null;
    readonly conversationId: number | null;
    readonly body: string | null;
    readonly htmlBody: string | null;
}

/** Metadata for a single email attachment. */
export interface AttachmentInfo {
    /** Position in the message's attachment list (1-based, matches AppleScript ordering). */
    readonly index: number;
    /** Original filename of the attachment. */
    readonly name: string;
    /** Size in bytes. */
    readonly size: number;
    /** MIME type (e.g., "application/pdf"). */
    readonly contentType: string;
}

/** Upper bound for downloading a single attachment (25 MB). */
export const MAX_ATTACHMENT_DOWNLOAD_SIZE = 25 * 1024 * 1024;

/** Upper bound for total attachment size when sending an email (25 MB). */
export const MAX_TOTAL_ATTACHMENT_SIZE = 25 * 1024 * 1024;

/** Aggregated unread message count, optionally broken down by folder. */
export interface UnreadCount {
    readonly total: number;
    readonly byFolder?: Record<number, number>;
}
