#!/usr/bin/env node
// =============================================================================
// doc-quality.mjs — Automated MCP tool documentation quality scorer
//
// Spawns dist/index.js, fetches tools/list, scores every tool description
// against a 6-criterion rubric (D1-D6) and every parameter against a
// 3-criterion rubric (P1-P3).
//
// Usage: node scripts/doc-quality.mjs
// Exit: 0 if all tools score >= 5/6, 1 otherwise
// =============================================================================

import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const SERVER_PATH = join(__dirname, '..', 'dist', 'index.js');
const TIMEOUT_MS = 15_000;
const TOOL_PASS_THRESHOLD = 5; // minimum D-score for grade A

// =============================================================================
// JSON-RPC transport (reused from smoke-test.mjs)
// =============================================================================

let nextId = 1;
let serverDead = false;
let serverProcess;
const pending = new Map();

function startServer() {
    serverProcess = spawn('node', [SERVER_PATH], {
        stdio: ['pipe', 'pipe', 'pipe'],
        env: { ...process.env },
    });
    serverProcess.on('error', (err) => { console.error(`Server error: ${err.message}`); process.exit(1); });
    serverProcess.on('exit', () => { serverDead = true; rejectAll('Server exited'); });
    serverProcess.stdin.on('error', () => {});
    const rl = createInterface({ input: serverProcess.stdout });
    rl.on('line', (line) => {
        if (!line.trim()) return;
        try {
            const msg = JSON.parse(line);
            if (msg.id != null && pending.has(msg.id)) {
                const { resolve, timer } = pending.get(msg.id);
                clearTimeout(timer);
                pending.delete(msg.id);
                resolve(msg);
            }
        } catch { /* ignore non-JSON lines */ }
    });
    rl.on('close', () => { serverDead = true; rejectAll('stdout closed'); });
    serverProcess.stderr.on('data', () => {}); // suppress
}

function rejectAll(reason) {
    for (const [, { reject, timer }] of pending) { clearTimeout(timer); reject(new Error(reason)); }
    pending.clear();
}

function request(method, params = {}) {
    if (serverDead) return Promise.reject(new Error('Server not running'));
    const id = nextId++;
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timeout: ${method}`)); }, TIMEOUT_MS);
        pending.set(id, { resolve, reject, timer });
        serverProcess.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    });
}

function notify(method, params = {}) {
    if (serverDead) return;
    serverProcess.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, params }) + '\n');
}

function stopServer() {
    rejectAll('Shutting down');
    if (serverProcess && !serverDead) { serverProcess.stdin.end(); serverProcess.kill('SIGTERM'); }
}

// =============================================================================
// Known tool names (for D6 cross-reference checking)
// =============================================================================

const ALL_TOOL_NAMES = [
    'list_accounts', 'list_folders', 'list_emails', 'search_emails', 'get_email',
    'get_unread_count', 'list_attachments', 'download_attachment',
    'list_calendars', 'list_events', 'get_event', 'search_events',
    'create_event', 'respond_to_event', 'delete_event', 'update_event',
    'list_contacts', 'search_contacts', 'get_contact',
    'list_tasks', 'search_tasks', 'get_task',
    'list_notes', 'get_note', 'search_notes',
    'send_email',
    'prepare_delete_email', 'confirm_delete_email',
    'prepare_move_email', 'confirm_move_email',
    'prepare_archive_email', 'confirm_archive_email',
    'prepare_junk_email', 'confirm_junk_email',
    'prepare_delete_folder', 'confirm_delete_folder',
    'prepare_empty_folder', 'confirm_empty_folder',
    'prepare_batch_delete_emails', 'prepare_batch_move_emails', 'confirm_batch_operation',
    'mark_email_read', 'mark_email_unread',
    'set_email_flag', 'clear_email_flag', 'set_email_categories',
    'create_folder', 'rename_folder', 'move_folder',
];

// =============================================================================
// Scoring Engine — Tool Descriptions (D1-D6)
// =============================================================================

const VERB_REGEX = /^(List|Get|Search|Create|Delete|Update|Set|Clear|Mark|Move|Rename|Prepare|Confirm|Send|Download|Respond|Retrieve|Return|Fetch|Find|Check|Count|Remove|Archive|Empty|Sort|Filter|Display|Show|Read)/i;

function scoreToolDescription(toolName, description) {
    const d = description || '';
    const lower = d.toLowerCase();
    const sentences = d.split(/\.\s+/).filter(s => s.trim().length > 0);

    const results = {
        D1: false, D2: false, D3: false, D4: false, D5: false, D6: false,
        total: 0, details: [],
    };

    // D1: What it does — clear verb phrase, >= 20 chars
    if (d.length >= 20 && VERB_REGEX.test(d.trim())) {
        results.D1 = true;
    } else {
        results.details.push('D1: Missing clear verb phrase opening (>= 20 chars)');
    }

    // D2: When to use / not use — multi-sentence OR contains usage guidance
    const hasUsageHint = /use this|when you|use .+ instead|to (find|get|check|retrieve|view|see|fetch|search|look up|discover|identify)/i.test(d);
    if (hasUsageHint || sentences.length >= 2) {
        results.D2 = true;
    } else {
        results.details.push('D2: Missing when-to-use guidance or multi-sentence description');
    }

    // D3: What it returns — describes output
    const hasReturnInfo = /returns?|response|result|output|provides|includes|contains|fields?|array of|object with|JSON|each .+ has|per-email/i.test(d);
    if (hasReturnInfo) {
        results.D3 = true;
    } else {
        results.details.push('D3: Missing description of what the tool returns');
    }

    // D4: Error behavior — mentions errors or failure conditions
    const hasErrorInfo = /error|not found|fail|invalid|if .+ does not exist|does not|cannot|missing|expired|mismatch|reject|unavailable|must be running|require/i.test(d);
    if (hasErrorInfo) {
        results.D4 = true;
    } else {
        results.details.push('D4: Missing error/failure conditions');
    }

    // D5: Side effects / caveats — notes state changes, limitations, prerequisites
    const hasSideEffects = /move|delete|create|update|mark|change|will|must|require|send|save|write|modif|permanent|irreversible|Outlook|prerequisite|caveat|limit|note:|caution|warning|overwrite|replace|clear|set|flag|archive|junk|empty|rename|token|expire|approve/i.test(d);
    if (hasSideEffects) {
        results.D5 = true;
    } else {
        results.details.push('D5: Missing side effects, caveats, or prerequisites');
    }

    // D6: Related tools — cross-references another tool or workflow step
    const mentionsOtherTool = ALL_TOOL_NAMES.some(name => {
        if (name === toolName) return false;
        return lower.includes(name);
    });
    const hasWorkflowRef = /after|before|then use|see also|related|companion|pair|workflow|first call|followed by|use .+ instead|combine with|together with/i.test(d);
    if (mentionsOtherTool || hasWorkflowRef) {
        results.D6 = true;
    } else {
        results.details.push('D6: Missing cross-reference to related tools');
    }

    results.total = [results.D1, results.D2, results.D3, results.D4, results.D5, results.D6]
        .filter(Boolean).length;

    return results;
}

// =============================================================================
// Scoring Engine — Parameter Descriptions (P1-P3)
// =============================================================================

function scoreParamDescription(toolName, paramName, paramSchema, isRequired) {
    const d = paramSchema?.description || '';
    const lower = d.toLowerCase();

    const results = { P1: false, P2: false, P3: false, total: 0, details: [], description: d };

    // P1: Semantic meaning — describes what the parameter controls (>= 10 chars)
    if (d.length >= 10) {
        results.P1 = true;
    } else {
        results.details.push('P1: Description too short or missing (need >= 10 chars)');
    }

    // P2: Format / example — includes an example value
    const hasExample = /e\.g\.|for example|such as|like |`[^`]+`|"[^"]+"|format:|ISO 8601/i.test(d);
    if (hasExample) {
        results.P2 = true;
    } else {
        results.details.push('P2: Missing example value (use "e.g." notation)');
    }

    // P3: Default behavior — for optional params, states what happens when omitted
    if (isRequired) {
        // Required params get P3 for free
        results.P3 = true;
    } else {
        const hasDefault = /default|if omitted|if not (specified|provided|set)|when omitted|optional|if not given|omit/i.test(d);
        if (hasDefault) {
            results.P3 = true;
        } else {
            results.details.push('P3: Missing default/omission behavior for optional param');
        }
    }

    results.total = [results.P1, results.P2, results.P3].filter(Boolean).length;
    return results;
}

// =============================================================================
// Grade assignment
// =============================================================================

function grade(score) {
    if (score >= 5) return 'A';
    if (score >= 3) return 'B';
    if (score >= 1) return 'C';
    return 'F';
}

// =============================================================================
// Main
// =============================================================================

async function run() {
    console.log('');
    console.log('=== MCP Tool Documentation Quality ===');
    console.log('');

    startServer();

    // Initialize
    await request('initialize', {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: { name: 'doc-quality', version: '1.0.0' },
    });
    notify('notifications/initialized');

    // Get tools list
    const toolsResp = await request('tools/list');
    const tools = toolsResp.result?.tools;
    if (!Array.isArray(tools)) {
        console.error('Failed to get tools list');
        stopServer();
        process.exit(1);
    }

    // Score each tool
    const toolScores = [];
    const paramScores = [];

    for (const tool of tools) {
        const score = scoreToolDescription(tool.name, tool.description);
        toolScores.push({ name: tool.name, ...score });

        // Score parameters
        const schema = tool.inputSchema;
        if (schema?.properties) {
            const required = new Set(schema.required || []);
            for (const [paramName, paramSchema] of Object.entries(schema.properties)) {
                const pScore = scoreParamDescription(tool.name, paramName, paramSchema, required.has(paramName));
                paramScores.push({ tool: tool.name, param: paramName, required: required.has(paramName), ...pScore });
            }
        }
    }

    // =========================================================================
    // Report: Tool Scores
    // =========================================================================
    console.log('TOOL SCORES');

    const sortedTools = [...toolScores].sort((a, b) => a.name.localeCompare(b.name));
    for (const t of sortedTools) {
        const g = grade(t.total);
        const pass = t.total >= TOOL_PASS_THRESHOLD;
        const icon = pass ? '\u2713' : '\u2717';
        const d1 = t.D1 ? '\u2713' : '\u00b7';
        const d2 = t.D2 ? '\u2713' : '\u00b7';
        const d3 = t.D3 ? '\u2713' : '\u00b7';
        const d4 = t.D4 ? '\u2713' : '\u00b7';
        const d5 = t.D5 ? '\u2713' : '\u00b7';
        const d6 = t.D6 ? '\u2713' : '\u00b7';
        console.log(`  ${icon} ${t.name.padEnd(30)} [${g}] ${t.total}/6  D1${d1} D2${d2} D3${d3} D4${d4} D5${d5} D6${d6}`);
    }

    // =========================================================================
    // Report: Parameter Scores (only failing)
    // =========================================================================
    console.log('');
    console.log('PARAMETER SCORES (failing only)');

    const failingParams = paramScores.filter(p => {
        if (p.required) return p.total < 2; // Required: need P1+P3 (P3 is free)
        return p.total < 2; // Optional: need P1+P3
    });

    if (failingParams.length === 0) {
        console.log('  All parameters pass minimum criteria.');
    } else {
        for (const p of failingParams) {
            const p1 = p.P1 ? '\u2713' : '\u00b7';
            const p2 = p.P2 ? '\u2713' : '\u00b7';
            const p3 = p.P3 ? '\u2713' : '\u00b7';
            const desc = p.description ? `"${p.description.slice(0, 60)}"` : '(none)';
            console.log(`  \u2717 ${p.tool}.${p.param.padEnd(25)} P1${p1} P2${p2} P3${p3}  ${desc}`);
        }
    }

    // =========================================================================
    // Report: Failing Tools with Suggestions
    // =========================================================================
    const failingTools = sortedTools.filter(t => t.total < TOOL_PASS_THRESHOLD);

    if (failingTools.length > 0) {
        console.log('');
        console.log(`FAILING TOOLS (score < ${TOOL_PASS_THRESHOLD}):`);
        for (const t of failingTools) {
            const missing = t.details.map(d => d.split(':')[0]).join(', ');
            console.log(`  ${t.name} (${t.total}/6): Missing ${missing}`);
            for (const detail of t.details) {
                console.log(`    ${detail}`);
            }
        }
    }

    // =========================================================================
    // Summary
    // =========================================================================
    console.log('');
    console.log('=== SUMMARY ===');

    const gradeA = sortedTools.filter(t => t.total >= 5).length;
    const gradeB = sortedTools.filter(t => t.total >= 3 && t.total < 5).length;
    const gradeC = sortedTools.filter(t => t.total >= 1 && t.total < 3).length;
    const gradeF = sortedTools.filter(t => t.total === 0).length;
    const mean = (sortedTools.reduce((sum, t) => sum + t.total, 0) / sortedTools.length).toFixed(1);

    console.log(`  Tools >= 5 (A): ${gradeA}/${sortedTools.length}`);
    console.log(`  Tools 3-4 (B): ${gradeB}/${sortedTools.length}`);
    console.log(`  Tools 1-2 (C): ${gradeC}/${sortedTools.length}`);
    console.log(`  Tools 0   (F): ${gradeF}/${sortedTools.length}`);
    console.log(`  Mean score: ${mean}/6`);

    const paramTotal = paramScores.length;
    const paramPassing = paramScores.filter(p => p.P1).length;
    const paramWithExamples = paramScores.filter(p => p.P2).length;
    const optionalParams = paramScores.filter(p => !p.required);
    const optionalWithDefaults = optionalParams.filter(p => p.P3).length;

    console.log('');
    console.log(`  Parameters with P1 (semantic): ${paramPassing}/${paramTotal}`);
    console.log(`  Parameters with P2 (example):  ${paramWithExamples}/${paramTotal}`);
    console.log(`  Optional params with P3 (default): ${optionalWithDefaults}/${optionalParams.length}`);

    const allPass = failingTools.length === 0;
    console.log('');
    console.log(`  RESULT: ${allPass ? 'PASS' : 'FAIL'} (target: all tools >= ${TOOL_PASS_THRESHOLD})`);
    console.log('');

    stopServer();
    process.exit(allPass ? 0 : 1);
}

run().catch((err) => {
    console.error('Unhandled error:', err);
    stopServer();
    process.exit(1);
});
