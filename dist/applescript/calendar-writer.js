import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseCreateEventResult } from './parser.js';
import { isoToDateComponents } from '../utils/dates.js';
/** Creates calendar events in Outlook via AppleScript. */
export class AppleScriptCalendarWriter {
    /**
     * Creates a calendar event with the specified properties.
     * Decomposes ISO dates into year/month/day/hour/minute components for
     * AppleScript, attaches optional recurrence rules, and parses the result.
     * @param params - Event properties including title, dates, and optional recurrence.
     * @returns The new event's ID and the calendar it was created in.
     */
    createEvent(params) {
        const start = isoToDateComponents(params.startDate);
        const end = isoToDateComponents(params.endDate);
        const scriptParams = {
            title: params.title,
            startYear: start.year,
            startMonth: start.month,
            startDay: start.day,
            startHours: start.hours,
            startMinutes: start.minutes,
            endYear: end.year,
            endMonth: end.month,
            endDay: end.day,
            endHours: end.hours,
            endMinutes: end.minutes,
        };
        if (params.calendarId != null)
            scriptParams.calendarId = params.calendarId;
        if (params.location != null)
            scriptParams.location = params.location;
        if (params.description != null)
            scriptParams.description = params.description;
        if (params.isAllDay != null)
            scriptParams.isAllDay = params.isAllDay;
        if (params.recurrence != null) {
            const rec = params.recurrence;
            const recurrenceScript = {
                frequency: rec.frequency,
                interval: rec.interval,
            };
            const mut = recurrenceScript;
            if (rec.daysOfWeek != null)
                mut.daysOfWeek = rec.daysOfWeek;
            if (rec.dayOfMonth != null)
                mut.dayOfMonth = rec.dayOfMonth;
            if (rec.weekOfMonth != null)
                mut.weekOfMonth = rec.weekOfMonth;
            if (rec.dayOfWeekMonthly != null)
                mut.dayOfWeekMonthly = rec.dayOfWeekMonthly;
            if (rec.endAfterCount != null)
                mut.endAfterCount = rec.endAfterCount;
            if (rec.endDate != null)
                mut.endDate = isoToDateComponents(rec.endDate);
            scriptParams.recurrence = recurrenceScript;
        }
        const script = scripts.createEvent(scriptParams);
        const output = executeAppleScriptOrThrow(script);
        const result = parseCreateEventResult(output);
        if (result == null) {
            throw new Error('Failed to parse create event result');
        }
        return result;
    }
}
/** Creates a new AppleScriptCalendarWriter instance. */
export function createCalendarWriter() {
    return new AppleScriptCalendarWriter();
}
