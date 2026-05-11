/**
 * Adaptive throttle and bridge-state classification for the Outlook AppleScript
 * bridge.
 *
 * ## Why this exists
 *
 * Outlook for Mac's AppleScript scripting bridge degrades under dense bursts of
 * AppleEvents. Empirically verified 2026-05-12: a typical `search_emails` call
 * fires 500–1500 AppleEvents inside ~30s. After a few thousand AppleEvents in
 * dense succession, Outlook's bridge stops responding (-1712 AppleEvent timeout
 * on every subsequent call) and only a graceful Outlook quit/relaunch recovers
 * it. Anecdotally this happens after 7h of MCP usage vs. the documented 4–12d
 * for normal usage — meaning the MCP's call pattern is what's killing it.
 *
 * ## Strategy
 *
 * Two cooperating mechanisms:
 *
 * 1. **Inter-call pacing** — enforce a minimum gap between AppleScript
 *    invocations. Gives Outlook time to clean up bridge state between calls.
 *    Pacing is adaptive: tightens when the bridge is fast, widens when it's
 *    slow. On a healthy bridge it's almost imperceptible (100ms).
 *
 * 2. **Adaptive backoff** — record per-call latency in a rolling window,
 *    classify bridge health (healthy/normal/stressed/degraded), and at the
 *    degraded tier refuse expensive operations entirely. This is the safety
 *    brake: it surfaces a clear `OUTLOOK_BRIDGE_STRESSED` error to callers
 *    BEFORE the bridge dies completely, while a graceful restart is still
 *    possible.
 *
 * ## State machine
 *
 *   healthy   (p95 < 500ms)   →  100ms throttle, no refusals
 *   normal    (p95 < 1000ms)  →  200ms throttle, no refusals
 *   stressed  (p95 < 2000ms)  →  500ms throttle, no refusals
 *   degraded  (p95 ≥ 2000ms)  →  1000ms throttle, refuse expensive ops
 *
 * Thresholds chosen empirically: a fresh Outlook bridge probes at 90–160ms;
 * a 4d-uptime bridge probes at 1–3s; a degenerate bridge probes at 2+ s and
 * eventually times out altogether.
 *
 * ## Sync sleep implementation
 *
 * `executeAppleScript` uses `execFileSync` and is synchronous all the way up.
 * To throttle without going async (which would require touching every caller)
 * we use `Atomics.wait` on a SharedArrayBuffer — this blocks the thread for
 * exactly the requested duration without busy-waiting and without forking a
 * subprocess. Node 16+.
 */
const WINDOW_CAPACITY = 10;
const latencyWindow = [];
let lastCallTime = 0;
const sleepBuffer = new Int32Array(new SharedArrayBuffer(4));
const STATE_CONFIG = {
    healthy: { throttleMs: 100, refuseExpensive: false },
    normal: { throttleMs: 200, refuseExpensive: false },
    stressed: { throttleMs: 500, refuseExpensive: false },
    degraded: { throttleMs: 1000, refuseExpensive: true },
};
function p95(samples) {
    if (samples.length === 0)
        return 0;
    const sorted = [...samples].sort((a, b) => a - b);
    const idx = Math.floor(sorted.length * 0.95);
    return sorted[Math.min(idx, sorted.length - 1)] ?? 0;
}
function classifyBridgeState(p95Latency, sampleCount) {
    // With fewer than 3 samples we can't classify reliably — default to normal.
    // This means the first few calls of a session don't get unnecessary throttle.
    if (sampleCount < 3)
        return 'normal';
    if (p95Latency < 500)
        return 'healthy';
    if (p95Latency < 1000)
        return 'normal';
    if (p95Latency < 2000)
        return 'stressed';
    return 'degraded';
}
export function currentBridgeState() {
    return classifyBridgeState(p95(latencyWindow), latencyWindow.length);
}
export function currentP95() {
    return p95(latencyWindow);
}
export function currentThrottleMs() {
    return STATE_CONFIG[currentBridgeState()].throttleMs;
}
export function isExpensiveOperationAllowed() {
    return !STATE_CONFIG[currentBridgeState()].refuseExpensive;
}
/**
 * Blocks the calling thread until the throttle slot opens up.
 * Uses Atomics.wait — no CPU burn, no subprocess overhead.
 */
export function waitForSlot() {
    const now = Date.now();
    const gap = currentThrottleMs();
    const elapsed = now - lastCallTime;
    if (elapsed < gap) {
        const sleepMs = gap - elapsed;
        // Atomics.wait on an unchanged value blocks for exactly `sleepMs`.
        // The buffer value is 0 and we wait for value 0 → immediate timeout-wait.
        Atomics.wait(sleepBuffer, 0, 0, sleepMs);
    }
    lastCallTime = Date.now();
}
export function recordLatency(ms) {
    latencyWindow.push(ms);
    if (latencyWindow.length > WINDOW_CAPACITY) {
        latencyWindow.shift();
    }
}
export function bridgeStateSnapshot() {
    const state = currentBridgeState();
    return {
        state,
        p95LatencyMs: p95(latencyWindow),
        sampleCount: latencyWindow.length,
        windowCapacity: WINDOW_CAPACITY,
        throttleMs: STATE_CONFIG[state].throttleMs,
        refusingExpensiveOps: STATE_CONFIG[state].refuseExpensive,
    };
}
/**
 * Test-only: reset the throttle state. Not exported via index.ts.
 */
export function _resetForTests() {
    latencyWindow.length = 0;
    lastCallTime = 0;
}
