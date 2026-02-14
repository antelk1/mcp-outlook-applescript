/**
 * Unit tests for the 4 bug fixes in mcp-outlook-applescript.
 *
 * These tests validate generated AppleScript strings and parser behaviour
 * WITHOUT requiring Microsoft Outlook to be running.
 *
 * Bug 1: setMessageFlag -- invalid AppleScript enum syntax
 * Bug 2: listTasks     -- double "is" in whose clause / isCompleted read-back
 * Bug 3: listEventsByDateRange -- timeout due to fetching 1000 events
 * Bug 4: flagStatus read-back  -- listMessages/getMessage/searchMessages/parser
 */

import { describe, it, expect, vi } from 'vitest';

// ---------------------------------------------------------------------------
// Imports under test
// ---------------------------------------------------------------------------

// scripts.ts exports pure template-generating functions (no Outlook dependency)
import {
  setMessageFlag,
  listMessages,
  getMessage,
  searchMessages,
  searchContacts,
  listTasks,
  searchTasks,
  getTask,
  searchEvents,
  searchNotes,
  listEvents,
  buildAppleScriptDateVar,
} from '../src/applescript/scripts.js';

// parser.ts exports pure parsing functions
import {
  parseEmails,
  parseEmail,
} from '../src/applescript/parser.js';

// repository.ts -- we need AppleScriptRepository for Bug 3 tests
import { AppleScriptRepository } from '../src/applescript/repository.js';

// pagination
import { paginate } from '../src/types/pagination.js';

// =============================================================================
// Bug 1: setMessageFlag -- valid AppleScript enum values
// =============================================================================

describe('Bug 1: setMessageFlag generates valid AppleScript enum values', () => {
  it('flagStatus 1 (flagged) produces "set todo flag of m to not completed"', () => {
    const script = setMessageFlag(123, 1);
    expect(script).toContain('set todo flag of m to not completed');
  });

  it('flagStatus 2 (completed) produces "set todo flag of m to completed"', () => {
    const script = setMessageFlag(123, 2);
    expect(script).toContain('set todo flag of m to completed');
  });

  it('flagStatus 0 (not flagged) produces "set todo flag of m to not flagged"', () => {
    const script = setMessageFlag(123, 0);
    expect(script).toContain('set todo flag of m to not flagged');
  });

  it('default/unknown flagStatus (e.g. 99) falls back to "not flagged"', () => {
    const script = setMessageFlag(123, 99);
    expect(script).toContain('set todo flag of m to not flagged');
  });

  it('uses the correct message id in the script', () => {
    const script = setMessageFlag(456, 1);
    expect(script).toContain('set m to message id 456');
  });

  it('does NOT contain old broken enum values like "flag marked"', () => {
    for (const status of [0, 1, 2]) {
      const script = setMessageFlag(100, status);
      expect(script).not.toContain('flag marked');
      expect(script).not.toContain('flag complete');
      expect(script).not.toContain('flag not flagged');
    }
  });

  it('does NOT contain bare numeric todo flag values', () => {
    // Verify we are using symbolic names, not bare numbers like "set todo flag of m to 1"
    for (const status of [0, 1, 2]) {
      const script = setMessageFlag(100, status);
      // The script should NOT have patterns like "to 0", "to 1", "to 2" as the flag value
      // (they should be symbolic: "not flagged", "not completed", "completed")
      const flagLine = script.split('\n').find(l => l.includes('set todo flag of m to'));
      expect(flagLine).toBeDefined();
      // After "to " should come a symbolic name, not a bare digit
      const afterTo = flagLine.trim().replace(/.*set todo flag of m to /, '');
      expect(afterTo).not.toMatch(/^\d+$/);
    }
  });

  it('returns a well-formed AppleScript tell block', () => {
    const script = setMessageFlag(123, 1);
    expect(script).toContain('tell application "Microsoft Outlook"');
    expect(script).toContain('end tell');
    expect(script).toContain('return "ok"');
  });
});

// =============================================================================
// Bug 2: listTasks -- correct whose clause and isCompleted read-back
// =============================================================================

describe('Bug 2: listTasks generates correct whose clause', () => {
  it('includeCompleted=false uses "whose todo flag is not completed"', () => {
    const script = listTasks(10, 0, false);
    expect(script).toContain('whose todo flag is not completed');
  });

  it('includeCompleted=false does NOT produce double "is" pattern', () => {
    const script = listTasks(10, 0, false);
    // The old bug: "whose is completed is false"
    expect(script).not.toContain('whose is completed is false');
    expect(script).not.toContain('whose is completed = false');
  });

  it('includeCompleted=true produces no whose filter', () => {
    const script = listTasks(10, 0, true);
    expect(script).not.toContain('whose');
  });

  it('uses "(todo flag of t is completed)" for isCompleted property read', () => {
    const scriptFiltered = listTasks(10, 0, false);
    const scriptAll = listTasks(10, 0, true);
    // Both should read completion status the same way
    expect(scriptFiltered).toContain('(todo flag of t is completed)');
    expect(scriptAll).toContain('(todo flag of t is completed)');
  });

  it('does NOT use old "is completed of t" syntax for property read', () => {
    const script = listTasks(10, 0, false);
    expect(script).not.toContain('is completed of t');
  });

  it('respects limit and offset parameters', () => {
    const script = listTasks(5, 3, false);
    // startIdx should be offset+1 = 4, endIdx should be limit+offset = 8
    expect(script).toContain('set startIdx to 4');
    expect(script).toContain('set endIdx to 8');
  });
});

describe('Bug 2: searchTasks uses correct isCompleted read-back', () => {
  it('contains "(todo flag of t is completed)" for property read', () => {
    const script = searchTasks('test query', 10, 0);
    expect(script).toContain('(todo flag of t is completed)');
  });

  it('does NOT contain "is completed of t"', () => {
    const script = searchTasks('test query', 10, 0);
    expect(script).not.toContain('is completed of t');
  });

  it('does NOT use a whose clause that filters by completion', () => {
    // searchTasks filters by name, not completion status
    const script = searchTasks('test query', 10, 0);
    expect(script).toContain('whose name contains');
    expect(script).not.toContain('whose todo flag');
    expect(script).not.toContain('whose is completed');
  });
});

describe('Bug 2: getTask uses correct isCompleted read-back', () => {
  it('contains "(todo flag of t is completed)" for property read', () => {
    const script = getTask(42);
    expect(script).toContain('(todo flag of t is completed)');
  });

  it('does NOT contain "is completed of t"', () => {
    const script = getTask(42);
    expect(script).not.toContain('is completed of t');
  });

  it('uses the correct task id', () => {
    const script = getTask(99);
    expect(script).toContain('set t to task id 99');
  });
});

// =============================================================================
// Bug 3: listEventsByDateRange -- now uses server-side AppleScript whose clause
// =============================================================================

describe('Bug 3: listEventsByDateRange uses server-side date filtering', () => {
  it('generates AppleScript with whose start time clause for date range', () => {
    // listEvents with date params generates a whose clause
    const script = listEvents(null, '2025-06-01T00:00:00Z', '2025-12-31T23:59:59Z', 10, 0);
    expect(script).toContain('whose start time');
    expect(script).toContain('afterDate');
    expect(script).toContain('beforeDate');
  });

  it('does NOT generate whose clause when no date filter', () => {
    const script = listEvents(null, null, null, 10, 0);
    expect(script).not.toContain('whose');
    expect(script).not.toContain('afterDate');
  });

  it('generates correct date variable construction', () => {
    const script = listEvents(null, '2025-06-15T14:30:00Z', null, 10, 0);
    expect(script).toContain('set year of afterDate to 2025');
    expect(script).toContain('set month of afterDate to 6');
    expect(script).toContain('set day of afterDate to 15');
    expect(script).toContain('set hours of afterDate to 14');
    expect(script).toContain('set minutes of afterDate to 30');
  });

  it('supports offset in date-filtered queries', () => {
    const script = listEvents(null, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z', 10, 5);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
  });
});

// =============================================================================
// Bug 4: flagStatus read-back in email scripts and parser
// =============================================================================

describe('Bug 4: listMessages includes flagStatus in generated AppleScript', () => {
  it('generates the flag reading block with set mFlag to "0"', () => {
    const script = listMessages(125, 5, 0, false);
    expect(script).toContain('set mFlag to "0"');
  });

  it('generates the todo flag conversion logic', () => {
    const script = listMessages(125, 5, 0, false);
    expect(script).toContain('set f to (todo flag of m) as string');
    expect(script).toContain('if f is "completed"');
    expect(script).toContain('set mFlag to "2"');
    expect(script).toContain('if f is "not flagged"');
    // Default else branch sets to "1" (flagged/marked)
    expect(script).toContain('set mFlag to "1"');
  });

  it('includes flagStatus in the output record', () => {
    const script = listMessages(125, 5, 0, false);
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });

  it('unreadOnly flag still works correctly', () => {
    const script = listMessages(125, 5, 0, true);
    expect(script).toContain('whose is read is false');
    // flagStatus should still be present
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });
});

describe('Bug 4: getMessage includes flagStatus', () => {
  it('generates the flag reading block', () => {
    const script = getMessage(123);
    expect(script).toContain('set mFlag to "0"');
    expect(script).toContain('set f to (todo flag of m) as string');
    expect(script).toContain('if f is "completed"');
    expect(script).toContain('set mFlag to "2"');
  });

  it('includes flagStatus in the output record', () => {
    const script = getMessage(123);
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });

  it('uses the correct message id', () => {
    const script = getMessage(789);
    expect(script).toContain('set m to message id 789');
  });
});

describe('Bug 4: searchMessages includes flagStatus in both phases', () => {
  it('includes flagStatus in Phase 1 (subject matches)', () => {
    const script = searchMessages('test', null, 5);
    // Phase 1 is the first block; it should have the flag logic
    const phase1Marker = '-- Phase 1: Subject matches';
    const phase2Marker = '-- Phase 2: Sender matches';
    const phase1Section = script.slice(
      script.indexOf(phase1Marker),
      script.indexOf(phase2Marker)
    );
    expect(phase1Section).toContain('set mFlag to "0"');
    expect(phase1Section).toContain('flagStatus{{=}}" & mFlag');
  });

  it('includes flagStatus in Phase 2 (sender matches)', () => {
    const script = searchMessages('test', null, 5);
    const phase2Marker = '-- Phase 2: Sender matches';
    const phase2Section = script.slice(script.indexOf(phase2Marker));
    expect(phase2Section).toContain('set mFlag to "0"');
    expect(phase2Section).toContain('flagStatus{{=}}" & mFlag');
  });

  it('includes todo flag conversion logic in both phases', () => {
    const script = searchMessages('test', null, 5);
    // Should have the todo flag logic appearing twice (once per phase)
    const matches = script.match(/set f to \(todo flag of m\) as string/g);
    expect(matches).not.toBeNull();
    expect(matches.length).toBe(2);
  });

  it('works with a folder ID specified', () => {
    const script = searchMessages('query', 42, 10);
    expect(script).toContain('of mail folder id 42');
    expect(script).toContain('flagStatus{{=}}" & mFlag');
  });
});

describe('Bug 4: parseEmails extracts flagStatus from AppleScript output', () => {
  it('parses flagStatus as a number from delimited output', () => {
    const output = '{{RECORD}}id{{=}}1{{FIELD}}subject{{=}}Test Email{{FIELD}}flagStatus{{=}}2';
    const results = parseEmails(output);
    expect(results).toHaveLength(1);
    expect(results[0].flagStatus).toBe(2);
  });

  it('parses flagStatus 0 (not flagged)', () => {
    const output = '{{RECORD}}id{{=}}10{{FIELD}}subject{{=}}Hello{{FIELD}}flagStatus{{=}}0';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBe(0);
  });

  it('parses flagStatus 1 (flagged)', () => {
    const output = '{{RECORD}}id{{=}}20{{FIELD}}subject{{=}}Important{{FIELD}}flagStatus{{=}}1';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBe(1);
  });

  it('returns null when flagStatus is missing', () => {
    const output = '{{RECORD}}id{{=}}30{{FIELD}}subject{{=}}No Flag';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBeNull();
  });

  it('returns null when flagStatus is empty string', () => {
    const output = '{{RECORD}}id{{=}}40{{FIELD}}subject{{=}}Empty Flag{{FIELD}}flagStatus{{=}}';
    const results = parseEmails(output);
    expect(results[0].flagStatus).toBeNull();
  });

  it('parses multiple records each with their own flagStatus', () => {
    const output =
      '{{RECORD}}id{{=}}1{{FIELD}}subject{{=}}A{{FIELD}}flagStatus{{=}}0' +
      '{{RECORD}}id{{=}}2{{FIELD}}subject{{=}}B{{FIELD}}flagStatus{{=}}1' +
      '{{RECORD}}id{{=}}3{{FIELD}}subject{{=}}C{{FIELD}}flagStatus{{=}}2';
    const results = parseEmails(output);
    expect(results).toHaveLength(3);
    expect(results[0].flagStatus).toBe(0);
    expect(results[1].flagStatus).toBe(1);
    expect(results[2].flagStatus).toBe(2);
  });
});

describe('Bug 4: parseEmail (single) extracts flagStatus', () => {
  it('parses a single email with flagStatus', () => {
    const output = '{{RECORD}}id{{=}}99{{FIELD}}subject{{=}}Single{{FIELD}}flagStatus{{=}}1';
    const result = parseEmail(output);
    expect(result).not.toBeNull();
    expect(result.flagStatus).toBe(1);
  });

  it('returns null for empty input', () => {
    const result = parseEmail('');
    expect(result).toBeNull();
  });
});

// =============================================================================
// Response Envelope: paginate() helper
// =============================================================================

describe('paginate() helper', () => {
  it('returns hasMore=true when items exceed limit', () => {
    const items = [1, 2, 3, 4, 5, 6]; // 6 items, limit is 5
    const result = paginate(items, 5);
    expect(result.hasMore).toBe(true);
    expect(result.count).toBe(5);
    expect(result.items).toHaveLength(5);
    expect(result.items).toEqual([1, 2, 3, 4, 5]);
  });

  it('returns hasMore=false when items are at or below limit', () => {
    const items = [1, 2, 3];
    const result = paginate(items, 5);
    expect(result.hasMore).toBe(false);
    expect(result.count).toBe(3);
    expect(result.items).toHaveLength(3);
  });

  it('returns hasMore=false when items equal limit exactly', () => {
    const items = [1, 2, 3, 4, 5];
    const result = paginate(items, 5);
    expect(result.hasMore).toBe(false);
    expect(result.count).toBe(5);
  });

  it('handles empty array', () => {
    const result = paginate([], 10);
    expect(result.hasMore).toBe(false);
    expect(result.count).toBe(0);
    expect(result.items).toEqual([]);
  });
});

// =============================================================================
// Offset in search AppleScript templates
// =============================================================================

describe('Offset support in search AppleScript templates', () => {
  it('searchContacts generates startIdx/endIdx from offset', () => {
    const script = searchContacts('john', 10, 5);
    expect(script).toContain('set startIdx to 6');
    expect(script).toContain('set endIdx to 15');
  });

  it('searchTasks generates startIdx/endIdx from offset', () => {
    const script = searchTasks('review', 10, 3);
    expect(script).toContain('set startIdx to 4');
    expect(script).toContain('set endIdx to 13');
  });

  it('searchNotes generates startIdx/endIdx from offset', () => {
    const script = searchNotes('meeting', 5, 10);
    expect(script).toContain('set startIdx to 11');
    expect(script).toContain('set endIdx to 15');
  });

  it('searchEvents generates startIdx/endIdx from offset', () => {
    const script = searchEvents('standup', 25, 0);
    expect(script).toContain('set startIdx to 1');
    expect(script).toContain('set endIdx to 25');
  });

  it('searchEvents with offset=25 generates startIdx 26', () => {
    const script = searchEvents('standup', 25, 25);
    expect(script).toContain('set startIdx to 26');
    expect(script).toContain('set endIdx to 50');
  });
});

// =============================================================================
// buildAppleScriptDateVar helper
// =============================================================================

describe('buildAppleScriptDateVar', () => {
  it('generates correct date components from ISO string', () => {
    const result = buildAppleScriptDateVar('afterDate', '2025-06-15T14:30:00Z');
    expect(result).toContain('set afterDate to current date');
    expect(result).toContain('set year of afterDate to 2025');
    expect(result).toContain('set month of afterDate to 6');
    expect(result).toContain('set day of afterDate to 15');
    expect(result).toContain('set hours of afterDate to 14');
    expect(result).toContain('set minutes of afterDate to 30');
    expect(result).toContain('set seconds of afterDate to 0');
  });

  it('handles midnight correctly', () => {
    const result = buildAppleScriptDateVar('beforeDate', '2025-01-01T00:00:00Z');
    expect(result).toContain('set year of beforeDate to 2025');
    expect(result).toContain('set month of beforeDate to 1');
    expect(result).toContain('set day of beforeDate to 1');
    expect(result).toContain('set hours of beforeDate to 0');
    expect(result).toContain('set minutes of beforeDate to 0');
  });

  it('sets day to 1 first to avoid month overflow', () => {
    const result = buildAppleScriptDateVar('d', '2025-12-31T23:59:00Z');
    const lines = result.split('\n');
    // "set day of d to 1" should come before "set year of d to ..."
    const dayTo1Idx = lines.findIndex(l => l.includes('set day of d to 1'));
    const yearIdx = lines.findIndex(l => l.includes('set year of d to'));
    expect(dayTo1Idx).toBeLessThan(yearIdx);
  });
});

// =============================================================================
// Date filtering in email AppleScript
// =============================================================================

describe('Date filtering in listMessages', () => {
  it('generates whose clause with time received >= afterDate', () => {
    const script = listMessages(100, 10, 0, false, '2025-06-01T00:00:00Z');
    expect(script).toContain('time received ≥ afterDate');
    expect(script).toContain('set year of afterDate to 2025');
    expect(script).toContain('set month of afterDate to 6');
  });

  it('generates whose clause with time received <= beforeDate', () => {
    const script = listMessages(100, 10, 0, false, undefined, '2025-12-31T23:59:59Z');
    expect(script).toContain('time received ≤ beforeDate');
    expect(script).toContain('set year of beforeDate to 2025');
    expect(script).toContain('set month of beforeDate to 12');
  });

  it('combines after + before + unreadOnly in whose clause', () => {
    const script = listMessages(100, 10, 0, true, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z');
    expect(script).toContain('is read is false and time received ≥ afterDate and time received ≤ beforeDate');
  });

  it('does NOT use whose clause when no date filter and not unread', () => {
    const script = listMessages(100, 10, 0, false);
    expect(script).not.toContain('whose');
    expect(script).toContain('messages 1 thru');
  });
});

describe('Date filtering in searchMessages', () => {
  it('extends whose clause with date filter in phase 1', () => {
    const script = searchMessages('test', null, 10, 0, '2025-06-01T00:00:00Z');
    // Phase 1 whose clause should include both subject and date
    expect(script).toContain('subject contains "test" and time received ≥ afterDate');
  });

  it('includes date check in phase 2 sender loop', () => {
    const script = searchMessages('test', null, 10, 0, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z');
    expect(script).toContain('mDateObj < afterDate');
    expect(script).toContain('mDateObj > beforeDate');
  });

  it('works without date filter (backward compatible)', () => {
    const script = searchMessages('test', null, 10);
    expect(script).toContain('whose subject contains "test"');
    expect(script).not.toContain('afterDate');
    expect(script).not.toContain('beforeDate');
  });

  it('uses native offset in AppleScript (startIdx/endIdx)', () => {
    const script = searchMessages('test', null, 10, 25);
    expect(script).toContain('set skipCount to 25');
    expect(script).toContain('set phase1Start to skipCount + 1');
    expect(script).toContain('set phase2Skip to skipCount - phase1Total');
  });
});

describe('Date filtering in searchEvents', () => {
  it('extends whose clause with after filter', () => {
    const script = searchEvents('standup', 10, 0, '2025-06-01T00:00:00Z');
    expect(script).toContain('subject contains "standup" and start time ≥ afterDate');
  });

  it('extends whose clause with both after and before', () => {
    const script = searchEvents('standup', 10, 0, '2025-01-01T00:00:00Z', '2025-12-31T23:59:59Z');
    expect(script).toContain('subject contains "standup" and start time ≥ afterDate and start time ≤ beforeDate');
  });
});

// =============================================================================
// search_notes description fix
// =============================================================================

describe('search_notes description fix', () => {
  it('searchNotes AppleScript searches by name (not content)', () => {
    const script = searchNotes('meeting', 10, 0);
    expect(script).toContain('whose name contains');
    // Should NOT contain "content contains" or "plain text content contains"
    expect(script).not.toContain('content contains');
  });
});
