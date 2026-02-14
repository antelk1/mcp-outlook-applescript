/**
 * Date conversion utilities for Outlook's Apple epoch timestamps.
 *
 * Outlook for Mac stores timestamps as seconds since the Apple epoch
 * (January 1, 2001, 00:00:00 UTC), while JavaScript uses milliseconds
 * since the Unix epoch (January 1, 1970, 00:00:00 UTC).
 */
/**
 * Seconds between Unix epoch (1970-01-01) and Apple epoch (2001-01-01).
 * Calculated as: Date.UTC(2001, 0, 1) / 1000 = 978307200
 */
export const APPLE_EPOCH_OFFSET = 978307200;
/**
 * Converts an Apple epoch timestamp to an ISO 8601 string.
 *
 * @param timestamp - Seconds since Apple epoch (2001-01-01), or null/undefined
 * @returns ISO 8601 formatted date string, or null if input is null/undefined
 *
 * @example
 * ```ts
 * appleTimestampToIso(0);
 * // Returns: '2001-01-01T00:00:00.000Z'
 *
 * appleTimestampToIso(null);
 * // Returns: null
 * ```
 */
export function appleTimestampToIso(timestamp: number | null | undefined): string | null {
    if (timestamp === null || timestamp === undefined) {
        return null;
    }
    const unixTimestampMs = (timestamp + APPLE_EPOCH_OFFSET) * 1000;
    return new Date(unixTimestampMs).toISOString();
}
/**
 * Converts an ISO 8601 string to an Apple epoch timestamp.
 *
 * @param isoString - ISO 8601 formatted date string, or null/undefined
 * @returns Seconds since Apple epoch (2001-01-01), or null if input is null/undefined
 */
export function isoToAppleTimestamp(isoString: string | null | undefined): number | null {
    if (isoString === null || isoString === undefined) {
        return null;
    }
    const date = new Date(isoString);
    if (isNaN(date.getTime())) {
        return null;
    }
    const unixTimestampSec = Math.floor(date.getTime() / 1000);
    return unixTimestampSec - APPLE_EPOCH_OFFSET;
}
/**
 * Converts a JavaScript Date to an Apple epoch timestamp.
 *
 * @param date - Date object, or null/undefined
 * @returns Seconds since Apple epoch (2001-01-01), or null if input is null/undefined
 */
/** Decomposes an ISO 8601 string into individual UTC date/time components. */
export function isoToDateComponents(isoString: string): {
    year: number;
    month: number;
    day: number;
    hours: number;
    minutes: number;
} {
    const date = new Date(isoString);
    return {
        year: date.getUTCFullYear(),
        month: date.getUTCMonth() + 1,
        day: date.getUTCDate(),
        hours: date.getUTCHours(),
        minutes: date.getUTCMinutes(),
    };
}
