/**
 * Repository interfaces and row types.
 *
 * These define the contract that the AppleScript repository implements.
 */

export interface FolderRow {
  readonly id: number;
  readonly name: string | null;
  readonly parentId: number | null;
  readonly specialType: number;
  readonly folderType: number;
  readonly accountId: number;
  readonly messageCount: number;
  readonly unreadCount: number;
}

export interface EmailRow {
  readonly id: number;
  readonly folderId: number;
  readonly subject: string | null;
  readonly sender: string | null;
  readonly senderAddress: string | null;
  readonly recipients: string | null;
  readonly displayTo: string | null;
  readonly toAddresses: string | null;
  readonly ccAddresses: string | null;
  readonly preview: string | null;
  readonly isRead: number;
  readonly timeReceived: number | null;
  readonly timeSent: number | null;
  readonly hasAttachment: number;
  readonly size: number;
  readonly priority: number;
  readonly flagStatus: number;
  readonly categories: Buffer | null;
  readonly messageId: string | null;
  readonly conversationId: number | null;
  readonly dataFilePath: string | null;
}

export interface EventRow {
  readonly id: number;
  readonly folderId: number;
  readonly startDate: number | null;
  readonly endDate: number | null;
  readonly isRecurring: number;
  readonly hasReminder: number;
  readonly attendeeCount: number;
  readonly uid: string | null;
  readonly masterRecordId: number | null;
  readonly recurrenceId: number | null;
  readonly dataFilePath: string | null;
}

export interface ContactRow {
  readonly id: number;
  readonly folderId: number;
  readonly displayName: string | null;
  readonly sortName: string | null;
  readonly contactType: number | null;
  readonly dataFilePath: string | null;
}

export interface TaskRow {
  readonly id: number;
  readonly folderId: number;
  readonly name: string | null;
  readonly isCompleted: number;
  readonly dueDate: number | null;
  readonly startDate: number | null;
  readonly priority: number;
  readonly hasReminder: number | null;
  readonly dataFilePath: string | null;
}

export interface NoteRow {
  readonly id: number;
  readonly folderId: number;
  readonly modifiedDate: number | null;
  readonly dataFilePath: string | null;
}

export interface IRepository {
  listFolders(): FolderRow[];
  getFolder(id: number): FolderRow | undefined;
  listEmails(folderId: number, limit: number, offset: number, after?: string, before?: string): EmailRow[];
  listUnreadEmails(folderId: number, limit: number, offset: number, after?: string, before?: string): EmailRow[];
  searchEmails(query: string, limit: number, offset: number, after?: string, before?: string): EmailRow[];
  searchEmailsInFolder(folderId: number, query: string, limit: number, offset: number, after?: string, before?: string): EmailRow[];
  getEmail(id: number): EmailRow | undefined;
  getUnreadCount(): number;
  getUnreadCountByFolder(folderId: number): number;
  listCalendars(): FolderRow[];
  listEvents(limit: number, offset: number): EventRow[];
  listEventsByFolder(folderId: number, limit: number, offset: number): EventRow[];
  listEventsByDateRange(startDate: number, endDate: number, limit: number, offset: number): EventRow[];
  getEvent(id: number): EventRow | undefined;
  listContacts(limit: number, offset: number): ContactRow[];
  searchContacts(query: string, limit: number, offset: number): ContactRow[];
  getContact(id: number): ContactRow | undefined;
  listTasks(limit: number, offset: number): TaskRow[];
  listIncompleteTasks(limit: number, offset: number): TaskRow[];
  searchTasks(query: string, limit: number, offset: number): TaskRow[];
  getTask(id: number): TaskRow | undefined;
  listNotes(limit: number, offset: number): NoteRow[];
  searchNotes(query: string, limit: number, offset: number): NoteRow[];
  searchEvents(query: string, limit: number, offset: number, after?: string, before?: string): EventRow[];
  getNote(id: number): NoteRow | undefined;
}

export interface IWriteableRepository extends IRepository {
  moveEmail(emailId: number, destinationFolderId: number): void;
  deleteEmail(emailId: number): void;
  archiveEmail(emailId: number): void;
  junkEmail(emailId: number): void;
  markEmailRead(emailId: number, isRead: boolean): void;
  setEmailFlag(emailId: number, flagStatus: number): void;
  setEmailCategories(emailId: number, categories: string[]): void;
  createFolder(name: string, parentFolderId?: number): FolderRow;
  deleteFolder(folderId: number): void;
  renameFolder(folderId: number, newName: string): void;
  moveFolder(folderId: number, destinationParentId: number): void;
  emptyFolder(folderId: number): void;
}
