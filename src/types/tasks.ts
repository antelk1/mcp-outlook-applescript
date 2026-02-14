/**
 * Domain types for Outlook tasks: summaries and full task records.
 */
import type { PriorityValue } from './mail.js';

/** Lightweight task representation used in list and search results. */
export interface TaskSummary {
    readonly id: number;
    readonly folderId: number;
    readonly name: string | null;
    readonly isCompleted: boolean;
    readonly dueDate: string | null;
    readonly priority: PriorityValue;
}

/** Complete task record including dates, reminder settings, body, and categories. */
export interface Task extends TaskSummary {
    readonly startDate: string | null;
    readonly completedDate: string | null;
    readonly hasReminder: boolean;
    readonly reminderDate: string | null;
    readonly body: string | null;
    readonly categories: readonly string[];
}
