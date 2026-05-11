# mcp-outlook-applescript

MCP server for Microsoft Outlook on Mac. 49 tools for mail, calendar, contacts, tasks, and notes via AppleScript. TypeScript with strict mode, Zod validation, and security-hardened input escaping.

## Commands

| Command | Purpose |
|---------|---------|
| `npm run build` | Compile TypeScript to `dist/` and set shebang |
| `npm run typecheck` | Type-check without emitting |
| `npm test` | Run 121 unit tests (vitest, no Outlook needed) |
| `bash scripts/audit.sh` | 28 static quality checks (build, security, package, functional) |
| `node scripts/doc-quality.mjs` | Score all 49 tool descriptions against 6-criterion rubric |
| `node scripts/smoke-test.mjs` | Live integration tests (requires Outlook running) |

## Testing

| When | Run | What it checks |
|------|-----|----------------|
| Any code change | `npm run build && npm test && bash scripts/audit.sh` | Types, 121 unit tests, 28 static checks (security, package, bug-fix preservation) |
| Before release, or after changing AppleScript/repository/parser | `node scripts/smoke-test.mjs` (Outlook must be running) | Live integration: 45/49 tools via JSON-RPC against real Outlook |
| After changing tool descriptions or Zod schemas | Give an MCP client a task without naming tools — it must discover and call them from descriptions | Tool discoverability, schema usability, response format |

## Architecture

```
src/                                 # 34 TypeScript files
├── index.ts                         # MCP server entry point, tool dispatch
├── applescript/                     # AppleScript backend
│   ├── executor.ts                  #   osascript runner, escapeForAppleScript()
│   ├── throttle.ts                  #   adaptive inter-call pacing + bridge-state classifier
│   ├── cache.ts                     #   generic TTL cache (folder list, etc.)
│   ├── scripts.ts                   #   AppleScript templates (pure functions)
│   ├── parser.ts                    #   delimiter-based output parser
│   ├── repository.ts                #   IRepository implementation
│   ├── account-repository.ts        #   account enumeration
│   ├── account-scripts.ts           #   account-related AppleScript
│   ├── calendar-manager.ts          #   event update/delete/respond
│   ├── calendar-writer.ts           #   event creation
│   ├── content-readers.ts           #   single-item content readers
│   ├── mail-sender.ts               #   send email with attachments
│   └── index.ts                     #   barrel exports
├── approval/                        # Two-step prepare/confirm for destructive ops
│   ├── token-manager.ts             #   single-use token store (5-min TTL)
│   ├── hash.ts                      #   content hashing for change detection
│   └── types.ts                     #   ApprovalToken, OperationType, TargetType
├── database/
│   └── repository.ts                # IRepository + IWriteableRepository interfaces
├── parsers/
│   └── html-stripper.ts             # HTML to plain text
├── tools/                           # Zod input schemas + tool logic
│   ├── mail.ts                      #   mail read tools + send_email
│   ├── calendar.ts                  #   calendar tools
│   ├── contacts.ts                  #   contact tools
│   ├── tasks.ts                     #   task tools
│   ├── notes.ts                     #   note tools
│   ├── mailbox-organization.ts      #   destructive ops + low-risk mutations
│   └── index.ts                     #   barrel exports
├── types/                           # Domain types + pagination envelope
│   └── pagination.ts                #   PaginatedResult<T>, paginate()
└── utils/
    ├── dates.ts                     # Apple epoch (2001) <-> ISO conversion
    └── errors.ts                    # ErrorCode enum, OutlookMcpError hierarchy
```

### Data flow

1. `src/index.ts` registers 49 tools on the MCP server with Zod schemas
2. Each handler calls a tool class (`MailTools`, `CalendarTools`, etc.) which calls `IRepository`
3. `IRepository` (in `applescript/repository.ts`) calls `scripts.ts` to build an AppleScript string
4. `executor.ts` runs it via `osascript` (stdin, not shell) and returns raw output
5. `parser.ts` splits output on `{{RECORD}}` / `{{FIELD}}` / `{{=}}` delimiters into typed rows
6. Tool class maps rows to domain types and wraps in `paginate()` for list/search results

### Lazy initialization

The server does not connect to Outlook on startup. The first tool call triggers `initializeAppleScriptBackend()`, which checks `isOutlookRunning()` and creates all repositories and tool instances.

## Key conventions

### Zod strict schemas

Every tool uses `z.strictObject()` with `.describe()` on every field. No `as any`, no `args as {...}` casts. Schemas live in `src/tools/<category>.ts` and are re-exported through `src/tools/index.ts`.

### Input escaping

All user-provided strings interpolated into AppleScript templates must pass through `escapeForAppleScript()` in `executor.ts`. This escapes backslashes, double quotes, and converts newlines to AppleScript `linefeed`/`return` concatenation. The audit script verifies this for `sendEmail()` specifically (addresses, replyTo, attachment paths, contentId).

### Pagination envelope

All list/search tools return `{ items, count, hasMore }`. The `paginate()` helper in `src/types/pagination.ts` implements this by over-fetching `limit + 1` rows and trimming. Callers page by incrementing `offset` by `limit`.

### Approval system for destructive operations

Destructive operations (delete, move, archive, junk, empty folder) use a two-step flow:

1. `prepare_*` tool: reads the target, hashes it, generates a single-use token (5-minute TTL), returns a preview
2. `confirm_*` tool: validates the token, checks the hash matches current state, executes the operation, consumes the token

Token state is in-memory only (`ApprovalTokenManager`). Expired tokens are garbage-collected when the store exceeds 100 entries.

### Bridge protection (added 2026-05-12)

Outlook for Mac's AppleScript scripting bridge degrades under dense bursts of AppleEvents. A typical `search_emails` call fires 500–1500 AppleEvents in ~30s; sustained MCP usage was empirically observed to kill the bridge in 7h (vs the documented 4–12d for normal usage). Three cooperating mechanisms in this codebase reduce the load and surface degradation cleanly:

1. **Adaptive throttle** (`applescript/throttle.ts`)
   - Module-level rolling window of last 10 AppleScript call latencies
   - Classifies bridge state from p95 latency: `healthy` (<500ms) / `normal` (<1000ms) / `stressed` (<2000ms) / `degraded` (≥2000ms)
   - Enforces a minimum inter-call gap that widens with stress: 100ms healthy → 1000ms degraded
   - Implemented with `Atomics.wait` on a SharedArrayBuffer — sync sleep, no CPU burn, no subprocess fork; lets the executor stay synchronous all the way up
   - The `skipThrottle: true` option exists for internal probes (warmup, `isOutlookRunning`) that shouldn't pollute the bridge-stress window or take a throttle slot

2. **TTL folder cache** (`applescript/cache.ts` + `repository.ts`)
   - Folder list cached for 5 minutes (extended from prior 30s); folders rarely change so repeat reads serve from memory
   - Invalidated on `createFolder` / `deleteFolder` / `renameFolder` / `moveFolder`
   - Generic `TtlCache<T>` is available for future caches if needed

3. **Safety brake** (`gateExpensive()` in `repository.ts`, `OutlookBridgeStressedError`)
   - At the `degraded` tier the throttle widens AND the heavy read methods (`listEmails` / `listUnreadEmails` / `searchEmails` / `searchEmailsInFolder`) refuse to execute
   - Refusal throws `OutlookBridgeStressedError` with a recovery hint: run `outlook-safe-restart.sh` (the operator's watchdog script) rather than retry
   - This is the safety brake — it surfaces a structured signal BEFORE a heavy call pushes Outlook over the edge into -1712 territory, while a graceful Outlook quit is still possible

The warmup query in `executeAppleScriptWithRetry` was deliberately chosen as `count of messages of outbox` (the same probe used by the watchdog script). The earlier `return name` was app-metadata-only and could falsely pass while the message store was hung — verified empirically when window enumeration succeeded but `count of messages of outbox` returned -1712 simultaneously.

### Error handling

All errors extend `OutlookMcpError` (in `utils/errors.ts`) which carries a structured `ErrorCode`. The `handle()` wrapper in `index.ts` catches all errors and returns `{ text, isError: true }` with `CODE: message` format. Error categories: `OUTLOOK_NOT_RUNNING`, `APPLESCRIPT_PERMISSION_DENIED`, `APPLESCRIPT_TIMEOUT`, `APPLESCRIPT_ERROR`, `OUTLOOK_BRIDGE_STRESSED`, `OUTLOOK_QUERY_REFUSED`, `NOT_FOUND`, `VALIDATION_ERROR`, `APPROVAL_*`, `ATTACHMENT_*`, `MAIL_SEND_ERROR`.

Three structurally-different timeout/refusal codes — distinguish carefully:
- `APPLESCRIPT_TIMEOUT` — "this specific call took too long". Retry might work (cold-start case) or might mean the bridge is unhappy.
- `OUTLOOK_BRIDGE_STRESSED` — "the adaptive throttle classifies the bridge as currently degraded based on rolling p95 latency; we're refusing to push more heavy ops at it". Recovery is via the watchdog/manual Outlook restart, NOT retry.
- `OUTLOOK_QUERY_REFUSED` — "this query's INPUTS are known dangerous regardless of current bridge state — narrow them before retrying". Currently fires for unscoped `searchEmails` on mailboxes > 50k messages.

### Delimiter-based output protocol

AppleScript outputs are not JSON. Scripts emit records separated by `{{RECORD}}`, fields by `{{FIELD}}`, key-value pairs by `{{=}}`, and null sentinels as `{{NULL}}`. The parser in `parser.ts` splits on these and converts to typed row interfaces.

## How to add a new tool

1. **`src/applescript/scripts.ts`** -- Add a template function. Use `escapeForAppleScript()` on all user inputs. Use shared output constants (`DELIMITERS`, `FLAG_STATUS_BLOCK`, etc.).
2. **`src/database/repository.ts`** -- Add the method to `IRepository` (reads) or `IWriteableRepository` (mutations).
3. **`src/applescript/repository.ts`** -- Implement: call the script, execute via `executeAppleScriptOrThrow`, parse, map to row types.
4. **`src/tools/<category>.ts`** -- Add a `z.strictObject()` schema with `.describe()` on every field. Add the method to the tool class. Use `paginate()` for list/search results.
5. **`src/tools/index.ts`** -- Export the new schema.
6. **`src/index.ts`** -- Register with `server.tool()`. Write a description covering: what it does, when to use it, return format, pagination, related tools.
7. **`test/unit.test.ts`** -- Test the AppleScript template, input escaping, and offset/limit handling.
8. **`scripts/smoke-test.mjs`** -- Update expected tool count and list. Add a happy-path test.
9. **Update tool count** in `README.md`, `CLAUDE.md`, and `scripts/smoke-test.mjs`.

## AppleScript gotchas

- **`whose` clause limitations:** Works on scalar properties (subject, display name) but NOT on collection properties (email addresses, attachments). For collections, iterate with a loop and try/catch. See the `searchMessages` two-phase pattern: phase 1 filters by subject via `whose`, phase 2 scans for sender matches.
- **`plain text content` is slow:** ~100ms per message. Only use in single-item `get_*` tools and `listMessages` (folder listing with preview), never in search results.
- **Date construction:** Use component-based approach (`set year of d to X`, `set month of d to Y`, ...) for locale safety. See `buildAppleScriptDateVar()`. Setting day to 1 first avoids month-overflow bugs (e.g., setting month to February when day is 31).
- **Timeout management:** Node-side default is 30s (45–60s for list/search). `searchTimeoutMs()` caps at 50s so the server returns before the MCP client's ~60s request timeout fires. The search AppleScript body is wrapped in `with timeout of 45 seconds` so Outlook returns a clean timeout error (−1712) instead of being SIGKILLed by Node — SIGKILL mid-conversation leaves Outlook's AppleScript bridge in a broken state and causes `Connection is invalid (−609)` on subsequent calls until Outlook self-heals (~10–30s).
- **Flag enums:** Outlook uses `not completed` / `completed` / `not flagged` for todo flag (not `flag marked` / `flag complete`). The `is completed` property does not exist on task objects; use `todo flag is completed` instead.
- **Apple epoch:** Outlook timestamps are seconds since 2001-01-01, not Unix epoch. Use `appleTimestampToIso()` / `isoToAppleTimestamp()`.

## Dependencies

| Dependency | Purpose |
|------------|---------|
| `@modelcontextprotocol/sdk` | MCP server framework (stdio transport) |
| `zod` | Input validation for all 49 tool handlers |
| `typescript` (dev) | Build toolchain (`strict: true`, `ES2022`, `NodeNext`) |
| `vitest` (dev) | Test runner |
| `@types/node` (dev) | Node.js type definitions |

No native dependencies. No network calls. No database.

## Known issues

- **Large mailboxes (50k+ messages) + unscoped search:** Phase 1 of `searchMessages` uses `messages whose subject contains …` with no `folder_id` — Outlook must scan every message. On an 84k-msg Inbox this can take 30–50s on a cold start. Always pass `folder_id` and/or `after`/`before` when possible. `SENDER_SCAN_LIMIT` is 50 (Phase 2) — sender-only matches are limited to the first 50 messages of the folder. **Verified 2026-05-12**: a single unscoped `search_emails(query="supabase")` across an 84k-msg mailbox didn't just take long — it timed out at -1712 AND left the bridge degraded for subsequent calls until manual Outlook restart. One bad query design is enough to break the bridge.
- **Safety brake is reactive, not predictive:** the `OutlookBridgeStressedError` brake only fires once the rolling p95 latency window crosses 2000ms — meaning the FIRST heavy call always gets through, regardless of how dangerous it is. Two predictive measures close the gap:
  - `searchEmails(query, ...)` with no `folder_id` no longer performs a mailbox-wide scan. Instead it auto-scopes to **INBOX + Sent Items** of the default account (resolved via `inbox of default account` and `folder "Sent Items" of default account`, cached for 60 minutes). Archive is deliberately excluded per user policy 2026-05-12 — searches there must be explicit.
  - `gateLargeFolderSearch` refuses any search (auto-scoped OR explicit `searchEmailsInFolder`) against a folder with >50,000 messages UNLESS `after` and/or `before` is provided. The structured `OUTLOOK_QUERY_REFUSED` error tells the caller to pass a date filter. This makes large folders (e.g., Archive at 56k+) searchable, but only when date-bounded.
- **Bridge can be "healthy by probe" and "degraded by search" simultaneously:** the watchdog/probe script's `count of messages of outbox` is a metadata-class query that often succeeds even when `whose` filters and per-row property reads time out. Don't trust a healthy probe as proof that searches will work — only an actual search call reveals search-path degradation.
- **Cold-start latency:** Outlook's first heavy AppleScript query after idle takes 30–45s while it pages in the message subject index. Subsequent calls are 1–4s. The `with timeout of 55 seconds` AS-internal cap absorbs cold starts; Node's 65–80s outer timeout has a 10s safety margin so AS-internal always fires first.
- **Zombie server processes:** Prior to the 2026-04-24 fix, the server had no `transport.onclose`/`SIGTERM` handlers, so Node processes leaked after every Claude Code session that disconnected. If you see many instances in `pgrep -af mcp-outlook-applescript`, they are from old versions — rebuild and restart. Current code exits cleanly on transport close or signal.
- **`osascript ETIMEDOUT` followed by `-609`:** If Node's sync `execFileSync` kills osascript mid-Apple-Event, Outlook's scripting bridge breaks briefly. The fix is the AppleScript-internal `with timeout of 55 seconds` (set inside `searchMessages`), which returns a clean -1712 error before Node's 65–80s kill-timer fires. If you still see −609, wait 10–30s for Outlook to self-heal, then retry.

## Outlook 16+ AppleScript gotchas (1.2.0)

The following AppleScript patterns DO NOT work in Outlook for Mac 16+ and must be replaced with two-step variants. See `SENDER_BLOCK` and `DATE_BLOCK` helpers in `scripts.ts`:

- **`address of sender of m`** — fails with "Can't make ... into type specifier". The sender record has class `«class radd»` for the email address, but property paths through `sender of m` don't resolve. **Fix:** assign sender to a local variable first, then access fields:
  ```applescript
  set _s to sender of m
  set mSender to «class radd» of _s   -- email address
  set mSenderName to name of _s       -- friendly name
  ```
  This bug silently produced empty `senderEmail`/`senderName` in all search and list results before 1.2.0, and broke Phase 2 sender-search entirely (the `if mSender contains query` check always failed because `mSender` was always "").
- **`time received` of a sent item** is `missing value`. Sent Items have `time sent` instead. `DATE_BLOCK` tries `time received` first, then falls back to `time sent`.
- **`messages whose sender contains "X"`** raises "Can't make X into type email address". `whose` cannot index into the sender record, so Phase 2 (sender-only matches) must use a manual loop with `SENDER_SCAN_LIMIT` cap.
- **`address of r` for recipients** has the same bug pattern as sender. Recipients in `get_email` may need a similar two-step fix in a future change (current code silently returns empty recipient lists in some cases).
