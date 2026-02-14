#!/usr/bin/env bash
# =============================================================================
# audit.sh — Formalized audit for mcp-outlook-applescript
#
# Runs checks from 4 reviewer perspectives:
#   1. External Reviewer (build, types, code quality)
#   2. Security Reviewer (injection, secrets, eval)
#   3. Bar Raiser (package quality, licensing)
#   4. End User (functional, bug fixes preserved)
#
# Exit codes: 0 = all passed, 1 = one or more failures
# Usage: bash scripts/audit.sh
# =============================================================================
set -uo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_DIR"

# Counters
PASS=0
FAIL=0
WARN=0

pass() { echo "  PASS  $1"; ((PASS++)); }
fail() { echo "  FAIL  $1"; ((FAIL++)); }
warn() { echo "  WARN  $1"; ((WARN++)); }

# =============================================================================
echo ""
echo "=== BUILD & TYPES (External Reviewer) ==="
echo ""

# 1. TypeScript compiles with zero errors
if npx tsc --noEmit 2>/dev/null; then
  pass "tsc --noEmit: zero type errors"
else
  fail "tsc --noEmit: type errors found"
fi

# 2. All tests pass
VITEST_OUTPUT=$(npx vitest run 2>&1) || true
if echo "$VITEST_OUTPUT" | grep -q "Tests.*passed"; then
  VITEST_COUNT=$(echo "$VITEST_OUTPUT" | grep "Tests" | grep -oE '[0-9]+ passed' | head -1)
  pass "vitest run: $VITEST_COUNT"
else
  fail "vitest run: test failures"
  echo "$VITEST_OUTPUT" | tail -5 | sed 's/^/         /'
fi

# 3. No 'as any' in src/
AS_ANY_COUNT=$(grep -rn 'as any' src/ --include='*.ts' 2>/dev/null | wc -l | tr -d ' ')
if [ "$AS_ANY_COUNT" -eq 0 ]; then
  pass "No 'as any' casts in src/"
else
  fail "Found $AS_ANY_COUNT 'as any' cast(s) in src/"
  grep -rn 'as any' src/ --include='*.ts' 2>/dev/null | head -5 | sed 's/^/         /'
fi

# 4. No @ts-ignore or @ts-expect-error
TS_SUPPRESS=$(grep -rn '@ts-ignore\|@ts-expect-error' src/ --include='*.ts' 2>/dev/null | wc -l | tr -d ' ')
if [ "$TS_SUPPRESS" -eq 0 ]; then
  pass "No @ts-ignore / @ts-expect-error in src/"
else
  fail "Found $TS_SUPPRESS type suppression(s) in src/"
  grep -rn '@ts-ignore\|@ts-expect-error' src/ --include='*.ts' 2>/dev/null | head -5 | sed 's/^/         /'
fi

# 5. No console.log (except top-level error handler in index.ts)
CONSOLE_LOG=$(grep -rn 'console\.log' src/ --include='*.ts' 2>/dev/null | wc -l | tr -d ' ')
if [ "$CONSOLE_LOG" -eq 0 ]; then
  pass "No console.log in src/"
else
  fail "Found $CONSOLE_LOG console.log(s) in src/"
  grep -rn 'console\.log' src/ --include='*.ts' 2>/dev/null | head -5 | sed 's/^/         /'
fi

# 6. Consistent node: prefix on Node.js imports
BARE_NODE_IMPORTS=$(grep -rn "from 'fs'\|from 'path'\|from 'child_process'\|from 'os'\|from 'crypto'\|from 'util'" src/ --include='*.ts' 2>/dev/null | wc -l | tr -d ' ')
if [ "$BARE_NODE_IMPORTS" -eq 0 ]; then
  pass "All Node.js imports use node: prefix"
else
  fail "Found $BARE_NODE_IMPORTS bare Node.js import(s) (missing node: prefix)"
  grep -rn "from 'fs'\|from 'path'\|from 'child_process'\|from 'os'\|from 'crypto'\|from 'util'" src/ --include='*.ts' 2>/dev/null | head -5 | sed 's/^/         /'
fi

# 7. Server version matches package.json
PKG_VERSION=$(node -e "process.stdout.write(JSON.parse(require('fs').readFileSync('package.json','utf8')).version)")
if grep -q "version: '$PKG_VERSION'" src/index.ts 2>/dev/null; then
  pass "Server version matches package.json ($PKG_VERSION)"
else
  SRV_VERSION=$(grep -o "version: '[^']*'" src/index.ts 2>/dev/null | head -1 || echo "not found")
  fail "Server version mismatch: package.json=$PKG_VERSION, src/index.ts has $SRV_VERSION"
fi

# =============================================================================
echo ""
echo "=== SECURITY ==="
echo ""

# 8. npm audit
AUDIT_OUTPUT=$(npm audit --omit=dev 2>&1 || true)
if echo "$AUDIT_OUTPUT" | grep -q "found 0 vulnerabilities"; then
  pass "npm audit: zero vulnerabilities"
else
  VULN_LINE=$(echo "$AUDIT_OUTPUT" | grep -E "^\d+ vulnerabilit" | head -1 || echo "unknown")
  warn "npm audit: $VULN_LINE"
fi

# 9. No eval() or new Function()
EVAL_COUNT=$(grep -rn 'eval(' src/ --include='*.ts' 2>/dev/null | wc -l | tr -d ' ')
NEW_FUNC_COUNT=$(grep -rn 'new Function(' src/ --include='*.ts' 2>/dev/null | wc -l | tr -d ' ')
DYNAMIC_CODE=$((EVAL_COUNT + NEW_FUNC_COUNT))
if [ "$DYNAMIC_CODE" -eq 0 ]; then
  pass "No eval() or new Function() in src/"
else
  fail "Found $DYNAMIC_CODE dynamic code execution(s) in src/"
fi

# 10. All string interpolation into AppleScript uses escapeForAppleScript()
# Check sendEmail specifically: email addresses, replyTo, attachments, contentId
# Look for raw ${...} inside double-quoted AppleScript strings that aren't escaped
UNESCAPED=$(grep -n 'address:"${' src/applescript/scripts.ts 2>/dev/null | grep -v 'escapeForAppleScript' | wc -l | tr -d ' ')
if [ "$UNESCAPED" -eq 0 ]; then
  pass "sendEmail() email addresses are escaped"
else
  fail "Found $UNESCAPED unescaped email address interpolation(s) in sendEmail()"
fi

# Check replyTo escaping
REPLY_UNESCAPED=$(grep -n 'reply to.*"${' src/applescript/scripts.ts 2>/dev/null | grep -v 'escapeForAppleScript' | wc -l | tr -d ' ')
if [ "$REPLY_UNESCAPED" -eq 0 ]; then
  pass "sendEmail() replyTo is escaped"
else
  fail "sendEmail() replyTo has unescaped interpolation"
fi

# Check attachment path escaping in sendEmail()
# Lines in sendEmail use escapeForAppleScript(att.path) or escapeForAppleScript(img.path)
# Lines in saveAttachment use ${escapedPath} which is pre-escaped via escapeForAppleScript
ATT_UNESCAPED=$(grep -n 'POSIX file "${' src/applescript/scripts.ts 2>/dev/null | grep -v 'escapeForAppleScript\|escapedPath' | wc -l | tr -d ' ')
if [ "$ATT_UNESCAPED" -eq 0 ]; then
  pass "sendEmail() attachment paths are escaped"
else
  fail "Found $ATT_UNESCAPED unescaped attachment path(s) in sendEmail()"
fi

# Check contentId escaping
CID_UNESCAPED=$(grep -n 'content id.*"${' src/applescript/scripts.ts 2>/dev/null | grep -v 'escapeForAppleScript' | wc -l | tr -d ' ')
if [ "$CID_UNESCAPED" -eq 0 ]; then
  pass "sendEmail() contentId is escaped"
else
  fail "sendEmail() contentId has unescaped interpolation"
fi

# 11. No hardcoded secrets
SECRETS=$(grep -rniE '(password|secret|api_key|apikey|token)\s*[:=]\s*["\x27][^"\x27]{8,}' src/ --include='*.ts' 2>/dev/null | grep -v 'type\|interface\|description\|@param\|token_id\|tokenId\|token_manager\|ApprovalToken' | wc -l | tr -d ' ')
if [ "$SECRETS" -eq 0 ]; then
  pass "No hardcoded secrets detected"
else
  warn "Found $SECRETS possible hardcoded secret(s)"
fi

# 12. .env in .gitignore
if grep -q '\.env' .gitignore 2>/dev/null; then
  pass ".env listed in .gitignore"
else
  fail ".env not listed in .gitignore"
fi

# Newline escaping in escapeForAppleScript
if grep -q 'linefeed' src/applescript/executor.ts 2>/dev/null; then
  pass "escapeForAppleScript() handles newlines (via AppleScript linefeed/return)"
else
  fail "escapeForAppleScript() does not escape newlines"
fi

# =============================================================================
echo ""
echo "=== PACKAGE QUALITY (Bar Raiser) ==="
echo ""

# 13. LICENSE file exists
if [ -f LICENSE ]; then
  pass "LICENSE file exists"
else
  fail "LICENSE file missing"
fi

# 14. README.md has no Graph API references
GRAPH_REFS=$(grep -ciE 'graph api|microsoft graph|device.code|device_code' README.md 2>/dev/null || true)
GRAPH_REFS=$(echo "$GRAPH_REFS" | tr -d '[:space:]')
if [ -z "$GRAPH_REFS" ] || [ "$GRAPH_REFS" -eq 0 ] 2>/dev/null; then
  pass "README.md has no Graph API references"
else
  warn "README.md mentions Graph API $GRAPH_REFS time(s)"
fi

# 15. No Graph-related dependencies in package.json
GRAPH_DEPS=$(grep -cE '@microsoft/microsoft-graph|@azure/msal|@azure/identity' package.json 2>/dev/null || true)
GRAPH_DEPS=$(echo "$GRAPH_DEPS" | tr -d '[:space:]')
if [ -z "$GRAPH_DEPS" ] || [ "$GRAPH_DEPS" -eq 0 ] 2>/dev/null; then
  pass "No Graph API dependencies in package.json"
else
  fail "Found $GRAPH_DEPS Graph-related dependency(ies)"
fi

# 16. npm pack contains no test files or dev artifacts
PACK_FILES=$(npm pack --dry-run 2>&1 | grep -E '\.test\.|\.spec\.|run-.*\.mjs|CLAUDE\.md|TEST-PLAN' | wc -l | tr -d ' ')
if [ "$PACK_FILES" -eq 0 ]; then
  pass "npm pack: no test/dev files in package"
else
  fail "npm pack: found $PACK_FILES test/dev file(s) in package"
  npm pack --dry-run 2>&1 | grep -E '\.test\.|\.spec\.|run-.*\.mjs|CLAUDE\.md|TEST-PLAN' | head -5 | sed 's/^/         /'
fi

# 17. Source maps: either work or are disabled
if grep -q '"sourceMap": true' tsconfig.json 2>/dev/null; then
  # Source maps enabled — check if src/ is in package
  if grep -q '"src"' package.json 2>/dev/null; then
    pass "Source maps enabled with src/ in package files"
  else
    fail "Source maps enabled but src/ not in package (maps will be broken)"
  fi
else
  pass "Source maps disabled (saves package size)"
fi

# =============================================================================
echo ""
echo "=== FUNCTIONAL (End User) ==="
echo ""

# 18. dist/ is gitignored (build artifacts don't belong in version control)
if grep -q '^dist/' .gitignore 2>/dev/null; then
  pass "dist/ is listed in .gitignore"
else
  fail "dist/ is not gitignored — build artifacts should not be committed"
fi

# 19. Build produces valid output (shebang + syntax)
# Since dist/ is gitignored, we build first, then verify
if npm run build --silent 2>/dev/null; then
  if head -1 dist/index.js 2>/dev/null | grep -q '#!/usr/bin/env node'; then
    pass "dist/index.js has shebang after build"
  else
    fail "dist/index.js missing shebang after build"
  fi
  if node --check dist/index.js 2>/dev/null; then
    pass "dist/index.js parses without syntax errors"
  else
    fail "dist/index.js has syntax errors"
  fi
else
  fail "npm run build failed"
fi

# 20. Bug fix 1: not completed in scripts.ts
if grep -q "'not completed'" src/applescript/scripts.ts 2>/dev/null; then
  pass "Bug fix 1 preserved: 'not completed' in setMessageFlag"
else
  fail "Bug fix 1 MISSING: 'not completed' not found in scripts.ts"
fi

# 21. Bug fix 2: whose todo flag is not completed
if grep -q 'whose todo flag is not completed' src/applescript/scripts.ts 2>/dev/null; then
  pass "Bug fix 2 preserved: 'whose todo flag is not completed' in listTasks"
else
  fail "Bug fix 2 MISSING: 'whose todo flag is not completed' not found"
fi

# 22. Bug fix 3: listEventsByDateRange uses server-side date filtering (not client-side)
# The original bug was client-side JS filtering capped at 100 events.
# The fix now delegates to AppleScript whose clause via listEvents with date params.
if grep -q 'listEventsByDateRange' src/applescript/repository.ts 2>/dev/null && \
   grep -q "scripts.listEvents(null, startIso, endIso" src/applescript/repository.ts 2>/dev/null; then
  pass "Bug fix 3 preserved: listEventsByDateRange uses server-side AppleScript date filter"
else
  fail "Bug fix 3 MISSING: listEventsByDateRange server-side date filter not found in repository.ts"
fi

# 23. Bug fix 4: todo flag of m (flagStatus read-back)
if grep -q 'todo flag of m' src/applescript/scripts.ts 2>/dev/null; then
  pass "Bug fix 4 preserved: 'todo flag of m' in message scripts"
else
  fail "Bug fix 4 MISSING: 'todo flag of m' not found"
fi

# =============================================================================
echo ""
echo "=== SUMMARY ==="
echo ""
echo "  Passed: $PASS"
echo "  Failed: $FAIL"
echo "  Warnings: $WARN"
echo ""

if [ "$FAIL" -gt 0 ]; then
  echo "  RESULT: FAIL ($FAIL issue(s) must be fixed)"
  exit 1
else
  echo "  RESULT: PASS (all checks passed)"
  exit 0
fi
