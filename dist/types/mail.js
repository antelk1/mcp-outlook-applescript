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
};
/** Numeric priority levels matching Outlook's internal representation. */
export const Priority = {
    High: 1,
    Normal: 3,
    Low: 5,
};
/** Numeric flag states for email follow-up tracking. */
export const FlagStatus = {
    None: 0,
    Flagged: 1,
    Completed: 2,
};
/** Upper bound for downloading a single attachment (25 MB). */
export const MAX_ATTACHMENT_DOWNLOAD_SIZE = 25 * 1024 * 1024;
/** Upper bound for total attachment size when sending an email (25 MB). */
export const MAX_TOTAL_ATTACHMENT_SIZE = 25 * 1024 * 1024;
