import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseRespondToEventResult, parseDeleteEventResult, parseUpdateEventResult } from './parser.js';
import { AppleScriptError } from '../utils/errors.js';
import type { RespondToEventResult } from './parser.js';

/** Accepted RSVP response values for calendar invitations. */
export type ResponseType = 'accept' | 'decline' | 'tentative';

/** Scope for recurring event operations: single instance or entire series. */
export type ApplyToScope = 'this_instance' | 'all_in_series';

/** Contract for managing existing calendar events (RSVP, delete, update). */
export interface ICalendarManager {
    /** Sends an RSVP response to a calendar invitation. */
    respondToEvent(eventId: number, response: ResponseType, sendResponse: boolean, comment?: string): RespondToEventResult;
    /** Deletes a calendar event, applying to one instance or the full series. */
    deleteEvent(eventId: number, applyTo: ApplyToScope): void;
    /** Updates one or more fields on an existing calendar event. */
    updateEvent(eventId: number, updates: EventUpdates, applyTo: ApplyToScope): UpdatedEvent;
}

/** Mutable fields that can be changed on an existing calendar event. */
export interface EventUpdates {
    readonly title?: string;
    readonly startDate?: string;
    readonly endDate?: string;
    readonly location?: string;
    readonly description?: string;
    readonly isAllDay?: boolean;
}

/** Result of a successful event update, listing which fields were changed. */
export interface UpdatedEvent {
    readonly id: number;
    readonly updatedFields: readonly string[];
}

/** Manages calendar events in Outlook via AppleScript. */
export class AppleScriptCalendarManager implements ICalendarManager {
    /**
     * Sends an RSVP for a calendar invitation.
     * @param eventId - Outlook event ID to respond to.
     * @param response - Accept, decline, or tentatively accept.
     * @param sendResponse - Whether to notify the organizer.
     * @param comment - Optional message to include with the RSVP.
     * @returns Parsed RSVP result including success status and event details.
     */
    respondToEvent(eventId: number, response: ResponseType, sendResponse: boolean, comment?: string): RespondToEventResult {
        const params = comment != null
            ? { eventId, response, sendResponse, comment }
            : { eventId, response, sendResponse };
        const script = scripts.respondToEvent(params);
        const output = executeAppleScriptOrThrow(script);
        const result = parseRespondToEventResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse RSVP response');
        }
        if (!result.success) {
            throw new AppleScriptError(result.error ?? 'RSVP operation failed');
        }
        return result;
    }

    /**
     * Removes a calendar event from Outlook.
     * @param eventId - Outlook event ID to delete.
     * @param applyTo - Whether to delete just this instance or the whole series.
     */
    deleteEvent(eventId: number, applyTo: ApplyToScope): void {
        const script = scripts.deleteEvent({ eventId, applyTo });
        const output = executeAppleScriptOrThrow(script);
        const result = parseDeleteEventResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse delete response');
        }
        if (!result.success) {
            throw new AppleScriptError(result.error ?? 'Delete operation failed');
        }
    }

    /**
     * Applies field-level changes to an existing calendar event.
     * @param eventId - Outlook event ID to update.
     * @param updates - Fields to modify (only non-null values are applied).
     * @param applyTo - Whether changes affect this instance or the whole series.
     * @returns The event ID and list of fields that were successfully changed.
     */
    updateEvent(eventId: number, updates: EventUpdates, applyTo: ApplyToScope): UpdatedEvent {
        const scriptUpdates = {
            ...(updates.title != null && { title: updates.title }),
            ...(updates.location != null && { location: updates.location }),
            ...(updates.description != null && { description: updates.description }),
            ...(updates.startDate != null && { startDate: updates.startDate }),
            ...(updates.endDate != null && { endDate: updates.endDate }),
            ...(updates.isAllDay != null && { isAllDay: updates.isAllDay }),
        };
        const scriptParams = {
            eventId,
            applyTo,
            updates: scriptUpdates,
        };
        const script = scripts.updateEvent(scriptParams);
        const output = executeAppleScriptOrThrow(script);
        const result = parseUpdateEventResult(output);
        if (result == null) {
            throw new AppleScriptError('Failed to parse update response');
        }
        if (!result.success) {
            throw new AppleScriptError(result.error ?? 'Update operation failed');
        }
        return {
            id: result.id!,
            updatedFields: result.updatedFields ?? [],
        };
    }
}

/** Creates a new AppleScriptCalendarManager instance. */
export function createCalendarManager(): ICalendarManager {
    return new AppleScriptCalendarManager();
}
