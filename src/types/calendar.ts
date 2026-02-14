/**
 * Domain types for Outlook calendars: folders, events, and attendees.
 */

/** A calendar folder belonging to an Outlook account. */
export interface CalendarFolder {
    readonly id: number;
    readonly name: string;
    readonly accountId: number;
}

/** Lightweight event representation used in list and search results. */
export interface EventSummary {
    readonly id: number;
    readonly folderId: number;
    readonly title: string | null;
    readonly startDate: string | null;
    readonly endDate: string | null;
    readonly isRecurring: boolean;
    readonly hasReminder: boolean;
    readonly attendeeCount: number;
    readonly uid: string | null;
}

/** Complete event record including description, location, and attendee list. */
export interface Event extends EventSummary {
    readonly location: string | null;
    readonly description: string | null;
    readonly organizer: string | null;
    readonly attendees: readonly Attendee[];
    readonly masterRecordId: number | null;
    readonly recurrenceId: number | null;
}

/** A single attendee on a calendar event with their RSVP status. */
export interface Attendee {
    readonly name: string | null;
    readonly email: string | null;
    readonly status: AttendeeStatus;
}

/** Possible RSVP response states for an event attendee. */
export const AttendeeStatus = {
    Unknown: 'unknown',
    Accepted: 'accepted',
    Declined: 'declined',
    Tentative: 'tentative',
    None: 'none',
} as const;
export type AttendeeStatus = (typeof AttendeeStatus)[keyof typeof AttendeeStatus];
