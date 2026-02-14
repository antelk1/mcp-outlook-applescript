/** AppleScript backend for Outlook on Mac — executor, scripts, parsers, and repositories. */
// Executor
export { executeAppleScript, executeAppleScriptOrThrow, escapeForAppleScript, isOutlookRunning, AppleScriptExecutionError, } from './executor.js';
// Scripts
export { DELIMITERS } from './scripts.js';
// Parser
export { parseFolders, parseEmails, parseEmail, parseCalendars, parseEvents, parseEvent, parseContacts, parseContact, parseTasks, parseTask, parseNotes, parseNote, parseCount, parseAccounts, parseDefaultAccountId, parseFoldersWithAccount, parseRespondToEventResult, parseDeleteEventResult, parseUpdateEventResult, parseSendEmailResult, parseAttachments, parseSaveAttachmentResult, } from './parser.js';
// Repository
export { AppleScriptRepository, createAppleScriptRepository, } from './repository.js';
// Account Repository
export { AccountRepository, createAccountRepository, } from './account-repository.js';
// Calendar Writer
export { AppleScriptCalendarWriter, createCalendarWriter, } from './calendar-writer.js';
// Calendar Manager
export { AppleScriptCalendarManager, createCalendarManager, } from './calendar-manager.js';
// Content Readers
export { AppleScriptEmailContentReader, AppleScriptEventContentReader, AppleScriptContactContentReader, AppleScriptTaskContentReader, AppleScriptNoteContentReader, AppleScriptAttachmentReader, createAppleScriptContentReaders, createEmailPath, createEventPath, createContactPath, createTaskPath, createNotePath, EMAIL_PATH_PREFIX, EVENT_PATH_PREFIX, CONTACT_PATH_PREFIX, TASK_PATH_PREFIX, NOTE_PATH_PREFIX, } from './content-readers.js';
// Mail Sender
export { AppleScriptMailSender, createMailSender, } from './mail-sender.js';
