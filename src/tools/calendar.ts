import { z } from 'zod';
import type { IRepository } from '../database/repository.js';
import type { CalendarFolder, EventSummary, Event, Attendee, PaginatedResult } from '../types/index.js';
import { paginate } from '../types/index.js';
import { appleTimestampToIso, isoToAppleTimestamp } from '../utils/dates.js';

// ---------------------------------------------------------------------------
// Zod input schemas for calendar MCP tools
// ---------------------------------------------------------------------------

export const ListCalendarsInput = z.strictObject({});

export const ListEventsInput = z.strictObject({
    calendar_id: z.number().int().positive().optional().describe('Calendar folder ID to filter by (e.g., from list_calendars). If omitted, returns events from all calendars.'),
    start_date: z.string().optional().describe('Start date filter in ISO 8601 format (e.g., "2025-01-01T00:00:00Z"). If omitted, no start date filter is applied.'),
    end_date: z.string().optional().describe('End date filter in ISO 8601 format (e.g., "2025-12-31T23:59:59Z"). If omitted, no end date filter is applied.'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of events to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of events to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
});

export const GetEventInput = z.strictObject({
    event_id: z.number().int().positive().describe('The event ID to retrieve (e.g., from list_events or search_events)'),
});

export const SearchEventsInput = z.strictObject({
    query: z.string().min(1).describe('Search query text matched against event titles (e.g., "standup")'),
    limit: z
        .number()
        .int()
        .min(1)
        .max(100)
        .default(25)
        .describe('Maximum number of events to return, 1-100 (e.g., 10). Defaults to 25 if omitted.'),
    offset: z.number().int().min(0).default(0).describe('Number of events to skip for pagination (e.g., 25 for page 2). Defaults to 0 if omitted.'),
    after: z.string().optional().describe('Only include events starting on or after this ISO 8601 date (e.g., "2025-01-01T00:00:00Z"). If omitted, no start date filter.'),
    before: z.string().optional().describe('Only include events starting on or before this ISO 8601 date (e.g., "2025-12-31T23:59:59Z"). If omitted, no end date filter.'),
});

export const RespondToEventInput = z.strictObject({
    event_id: z.number().int().positive().describe('The event ID to respond to (e.g., from list_events or search_events)'),
    response: z.enum(['accept', 'decline', 'tentative']).describe('Your RSVP response: "accept", "decline", or "tentative"'),
    send_response: z.boolean().default(true).describe('Whether to send your response notification to the organizer. Defaults to true if omitted.'),
    comment: z.string().optional().describe('Optional comment to include with your response. If omitted, no comment is sent.'),
});

const DayOfWeek = z.enum([
    'sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday',
]);

const RecurrenceEndInput = z.discriminatedUnion('type', [
    z.strictObject({ type: z.literal('no_end') }),
    z.strictObject({
        type: z.literal('end_date'),
        date: z.string().describe('End date in ISO 8601 format'),
    }),
    z.strictObject({
        type: z.literal('end_after_count'),
        count: z.number().int().min(1).max(999).describe('Number of occurrences'),
    }),
]);

export const RecurrenceInput = z.strictObject({
    frequency: z.enum(['daily', 'weekly', 'monthly', 'yearly']).describe('How often the event repeats'),
    interval: z.number().int().min(1).max(999).default(1).describe('Number of frequency units between occurrences (e.g., 2 for every 2 weeks)'),
    days_of_week: z.array(DayOfWeek).min(1).optional().describe('Days of the week for weekly recurrence (e.g., ["monday", "wednesday"])'),
    day_of_month: z.number().int().min(1).max(31).optional().describe('Day of the month for monthly recurrence (e.g., 15 for the 15th)'),
    week_of_month: z.enum(['first', 'second', 'third', 'fourth', 'last']).optional().describe('Week of the month for ordinal monthly recurrence (e.g., "third" for 3rd Thursday)'),
    day_of_week_monthly: DayOfWeek.optional().describe('Day of week for ordinal monthly recurrence (used with week_of_month)'),
    end: RecurrenceEndInput.default({ type: 'no_end' }).describe('When the recurrence ends'),
}).superRefine((data, ctx) => {
    if (data.frequency === 'weekly' && (data.days_of_week == null || data.days_of_week.length === 0)) {
        ctx.addIssue({
            code: z.ZodIssueCode.custom,
            message: 'days_of_week is required for weekly recurrence',
            path: ['days_of_week'],
        });
    }
    if (data.frequency === 'monthly') {
        const hasOrdinal = data.week_of_month != null;
        const hasDayOfWeekMonthly = data.day_of_week_monthly != null;
        if (hasOrdinal !== hasDayOfWeekMonthly) {
            ctx.addIssue({
                code: z.ZodIssueCode.custom,
                message: 'week_of_month and day_of_week_monthly must both be provided for ordinal monthly recurrence',
                path: ['week_of_month'],
            });
        }
    }
    if (data.frequency !== 'monthly') {
        if (data.day_of_month != null || data.week_of_month != null || data.day_of_week_monthly != null) {
            ctx.addIssue({
                code: z.ZodIssueCode.custom,
                message: 'day_of_month, week_of_month, and day_of_week_monthly are only valid for monthly recurrence',
                path: ['frequency'],
            });
        }
    }
    if (data.frequency !== 'weekly' && data.days_of_week != null) {
        ctx.addIssue({
            code: z.ZodIssueCode.custom,
            message: 'days_of_week is only valid for weekly recurrence',
            path: ['days_of_week'],
        });
    }
});

const isoDateString = z
    .string()
    .refine((s) => !isNaN(Date.parse(s)), { message: 'Must be a valid ISO 8601 date string' });

export const CreateEventInput = z.strictObject({
    title: z.string().min(1).describe('Event title/subject (e.g., "Team Standup")'),
    start_date: isoDateString.describe('Start date in ISO 8601 UTC format (e.g., "2025-06-15T14:00:00Z")'),
    end_date: isoDateString.describe('End date in ISO 8601 UTC format (e.g., "2025-06-15T15:00:00Z"). Must be after start_date.'),
    calendar_id: z.number().int().positive().optional().describe('Calendar ID to create the event in (e.g., from list_calendars). If omitted, uses the default calendar.'),
    location: z.string().optional().describe('Event location (e.g., "Conference Room A"). If omitted, no location is set.'),
    description: z.string().optional().describe('Event description/body text. If omitted, no description is set.'),
    is_all_day: z.boolean().optional().default(false).describe('Whether this is an all-day event. Defaults to false if omitted.'),
    recurrence: RecurrenceInput.optional().describe('Recurrence pattern to make this a repeating event. If omitted, creates a single non-recurring event.'),
}).refine((data) => new Date(data.start_date).getTime() < new Date(data.end_date).getTime(), { message: 'start_date must be before end_date', path: ['start_date'] });

const ApplyTo = z.enum(['this_instance', 'all_in_series']).default('this_instance');

export const DeleteEventInput = z.strictObject({
    event_id: z.number().int().positive().describe('The event ID to delete (e.g., from list_events or search_events)'),
    apply_to: ApplyTo.describe('For recurring events: "this_instance" deletes one occurrence, "all_in_series" deletes the entire series. Defaults to "this_instance" if omitted.'),
});

export const UpdateEventInput = z.strictObject({
    event_id: z.number().int().positive().describe('The event ID to update (e.g., from list_events or search_events)'),
    apply_to: ApplyTo.describe('For recurring events: "this_instance" updates one occurrence, "all_in_series" updates the entire series. Defaults to "this_instance" if omitted.'),
    title: z.string().optional().describe('New event title. If omitted, title is unchanged.'),
    start_date: isoDateString.optional().describe('New start date in ISO 8601 UTC format (e.g., "2025-06-15T14:00:00Z"). If omitted, start date is unchanged.'),
    end_date: isoDateString.optional().describe('New end date in ISO 8601 UTC format (e.g., "2025-06-15T15:00:00Z"). If omitted, end date is unchanged.'),
    location: z.string().optional().describe('New location. If omitted, location is unchanged.'),
    description: z.string().optional().describe('New description. If omitted, description is unchanged.'),
    is_all_day: z.boolean().optional().describe('Whether the event is all day. If omitted, all-day status is unchanged.'),
}).refine((data) => {
    if (data.start_date != null && data.end_date != null) {
        return new Date(data.start_date).getTime() < new Date(data.end_date).getTime();
    }
    return true;
}, { message: 'start_date must be before end_date', path: ['start_date'] });

/** Validated parameters for listing calendars. */
export type ListCalendarsParams = z.infer<typeof ListCalendarsInput>;
/** Validated parameters for listing events. */
export type ListEventsParams = z.infer<typeof ListEventsInput>;
/** Validated parameters for retrieving a single event. */
export type GetEventParams = z.infer<typeof GetEventInput>;
/** Validated parameters for searching events. */
export type SearchEventsParams = z.infer<typeof SearchEventsInput>;
/** Validated parameters for creating an event. */
export type CreateEventParams = z.infer<typeof CreateEventInput>;
/** Validated recurrence configuration for an event. */
export type RecurrenceParams = z.infer<typeof RecurrenceInput>;
/** Validated parameters for responding to an event invitation. */
export type RespondToEventParams = z.infer<typeof RespondToEventInput>;
/** Validated parameters for deleting an event. */
export type DeleteEventParams = z.infer<typeof DeleteEventInput>;
/** Validated parameters for updating an event. */
export type UpdateEventParams = z.infer<typeof UpdateEventInput>;

/** Describes the result returned after successfully creating a calendar event. */
export interface CreateEventResult {
    readonly id: number;
    readonly title: string;
    readonly start_date: string;
    readonly end_date: string;
    readonly calendar_id: number | null;
    readonly location: string | null;
    readonly description: string | null;
    readonly is_all_day: boolean;
    readonly is_recurring: boolean;
}

/** Reads rich event details (title, location, description, attendees) from a data file. */
export interface IEventContentReader {
    readEventDetails(dataFilePath: string | null): EventDetails | null;
}

/** Rich event details extracted from an event's data file. */
export interface EventDetails {
    readonly title: string | null;
    readonly location: string | null;
    readonly description: string | null;
    readonly organizer: string | null;
    readonly attendees: readonly Attendee[];
}

/** No-op content reader that always returns null. Used when no data-file reader is available. */
export const nullEventContentReader: IEventContentReader = {
    readEventDetails: () => null,
};

// ---------------------------------------------------------------------------
// Row-to-domain transformers
// ---------------------------------------------------------------------------

/** Converts a raw calendar repository row into a CalendarFolder domain object. */
function transformCalendar(row: ReturnType<IRepository['listCalendars']>[number]): CalendarFolder {
    return {
        id: row.id,
        name: row.name ?? 'Unnamed',
        accountId: row.accountId,
    };
}

/** Converts a raw event repository row into an EventSummary, optionally attaching a content-reader title. */
function transformEventSummary(row: ReturnType<IRepository['getEvent']> & {}, title: string | null = null): EventSummary {
    return {
        id: row.id,
        folderId: row.folderId,
        title: title,
        startDate: appleTimestampToIso(row.startDate),
        endDate: appleTimestampToIso(row.endDate),
        isRecurring: row.isRecurring === 1,
        hasReminder: row.hasReminder === 1,
        attendeeCount: row.attendeeCount,
        uid: row.uid,
    };
}

/** Converts a raw event repository row and its rich details into a full Event domain object. */
function transformEvent(row: ReturnType<IRepository['getEvent']> & {}, details: EventDetails | null): Event {
    const summary = transformEventSummary(row, details?.title ?? null);
    return {
        ...summary,
        location: details?.location ?? null,
        description: details?.description ?? null,
        organizer: details?.organizer ?? null,
        attendees: details?.attendees ?? [],
        masterRecordId: row.masterRecordId ?? null,
        recurrenceId: row.recurrenceId ?? null,
    };
}

// ---------------------------------------------------------------------------
// CalendarTools -- provides read operations for calendars and events
// ---------------------------------------------------------------------------

/** Exposes calendar and event read operations backed by a repository and an optional content reader. */
export class CalendarTools {
    private readonly repository: IRepository;
    private readonly contentReader: IEventContentReader;

    constructor(repository: IRepository, contentReader: IEventContentReader = nullEventContentReader) {
        this.repository = repository;
        this.contentReader = contentReader;
    }

    /** Returns all calendar folders visible across configured accounts. */
    listCalendars(_params: ListCalendarsParams): CalendarFolder[] {
        const rows = this.repository.listCalendars();
        return rows.map(transformCalendar);
    }

    /** Returns event summaries, optionally filtered by calendar, date range, or a result limit. */
    listEvents(params: ListEventsParams): PaginatedResult<EventSummary> {
        const { calendar_id, start_date, end_date, limit, offset } = params;
        let rows;
        if (start_date != null && end_date != null) {
            const startTimestamp = isoToAppleTimestamp(start_date);
            const endTimestamp = isoToAppleTimestamp(end_date);
            if (startTimestamp != null && endTimestamp != null) {
                rows = this.repository.listEventsByDateRange(startTimestamp, endTimestamp, limit + 1, offset);
            }
            else {
                rows = this.repository.listEvents(limit + 1, offset);
            }
        }
        else if (calendar_id != null) {
            rows = this.repository.listEventsByFolder(calendar_id, limit + 1, offset);
        }
        else {
            rows = this.repository.listEvents(limit + 1, offset);
        }
        const items = rows.map((row) => {
            const details = this.contentReader.readEventDetails(row.dataFilePath);
            return transformEventSummary(row, details?.title ?? null);
        });
        return paginate(items, limit);
    }

    /** Retrieves a single event by ID with full details, or null if not found. */
    getEvent(params: GetEventParams): Event | null {
        const { event_id } = params;
        const row = this.repository.getEvent(event_id);
        if (row == null) {
            return null;
        }
        const details = this.contentReader.readEventDetails(row.dataFilePath);
        return transformEvent(row, details);
    }

    /** Searches events by title text and returns matching summaries up to the given limit. */
    searchEvents(params: SearchEventsParams): PaginatedResult<EventSummary> {
        const { query, limit, offset, after, before } = params;
        const rows = this.repository.searchEvents(query, limit + 1, offset, after, before);
        const items = rows.map(row => {
            const details = this.contentReader.readEventDetails(row.dataFilePath);
            const title = details?.title ?? '';
            return transformEventSummary(row, title);
        });
        return paginate(items, limit);
    }
}

/** Factory that creates a CalendarTools instance with the given repository and optional content reader. */
export function createCalendarTools(repository: IRepository, contentReader: IEventContentReader = nullEventContentReader): CalendarTools {
    return new CalendarTools(repository, contentReader);
}
