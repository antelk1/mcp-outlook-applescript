import { z } from 'zod';
import { paginate } from '../types/index.js';
import { appleTimestampToIso } from '../utils/dates.js';
// ---------------------------------------------------------------------------
// Zod input schemas for task MCP tools
// ---------------------------------------------------------------------------
export const ListTasksInput = z.strictObject({
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of tasks to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of tasks to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
    include_completed: z.boolean().default(true).describe('Whether to include completed tasks. Set to false to see only incomplete tasks. Defaults to true if omitted.'),
});
export const SearchTasksInput = z.strictObject({
    query: z.string().min(1).describe('Search query text matched against task names (e.g., "review")'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of tasks to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of tasks to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
});
export const GetTaskInput = z.strictObject({
    task_id: z.number().int().positive().describe('The task ID to retrieve (e.g., from list_tasks or search_tasks)'),
});
/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullTaskContentReader = {
    readTaskDetails: () => null,
};
// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------
/** Converts a raw task repository row into a TaskSummary domain object. */
function transformTaskSummary(row) {
    return {
        id: row.id,
        folderId: row.folderId,
        name: row.name,
        isCompleted: row.isCompleted === 1,
        dueDate: appleTimestampToIso(row.dueDate),
        priority: row.priority,
    };
}
/** Converts a raw task repository row and its rich details into a full Task domain object. */
function transformTask(row, details) {
    const summary = transformTaskSummary(row);
    return {
        ...summary,
        startDate: appleTimestampToIso(row.startDate),
        completedDate: details?.completedDate ?? null,
        hasReminder: row.hasReminder === 1,
        reminderDate: details?.reminderDate ?? null,
        body: details?.body ?? null,
        categories: details?.categories ?? [],
    };
}
// ---------------------------------------------------------------------------
// TasksTools -- provides read operations for Outlook tasks
// ---------------------------------------------------------------------------
/** Exposes task read operations backed by a repository and an optional content reader. */
export class TasksTools {
    repository;
    contentReader;
    constructor(repository, contentReader = nullTaskContentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
    }
    /** Returns a paginated list of task summaries, optionally excluding completed tasks. */
    listTasks(params) {
        const { limit, offset, include_completed } = params;
        const rows = include_completed
            ? this.repository.listTasks(limit + 1, offset)
            : this.repository.listIncompleteTasks(limit + 1, offset);
        return paginate(rows.map(transformTaskSummary), limit);
    }
    /** Searches tasks by name and returns matching summaries up to the given limit. */
    searchTasks(params) {
        const { query, limit, offset } = params;
        const rows = this.repository.searchTasks(query, limit + 1, offset);
        return paginate(rows.map(transformTaskSummary), limit);
    }
    /** Retrieves a single task by ID with full details, or null if not found. */
    getTask(params) {
        const { task_id } = params;
        const row = this.repository.getTask(task_id);
        if (row == null) {
            return null;
        }
        const details = this.contentReader.readTaskDetails(row.dataFilePath);
        return transformTask(row, details);
    }
}
/** Factory that creates a TasksTools instance with the given repository and optional content reader. */
export function createTasksTools(repository, contentReader = nullTaskContentReader) {
    return new TasksTools(repository, contentReader);
}
