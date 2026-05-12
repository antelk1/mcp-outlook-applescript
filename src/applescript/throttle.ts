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
 * 2. **Adaptive backoff** — record per-call latency in a rolling window with
 *    time-decay (samples older than 5 minutes are excluded from classification
 *    so a single old slow call doesn't dominate the metric forever), classify
 *    bridge health (healthy/normal/stressed/degraded) from the **median** of
 *    recent samples, and at the degraded tier refuse expensive operations
 *    entirely. This is the safety brake: it surfaces a clear
 *    `OUTLOOK_BRIDGE_STRESSED` error to callers BEFORE the bridge dies
 *    completely, while a graceful restart is still possible.
 *
 * ## Why median, not p95
 *
 * An earlier version used `p95()` with `idx = floor(N * 0.95)`. For N ≤ 20
 * that math returns `samples[N-1]` — i.e., the maximum, not the 95th
 * percentile. One slow call (say a 13s cold-start search) set the metric to
 * its own value for the full 5-minute decay window and refused everything,
 * even when 9 subsequent calls were fast. Verified 2026-05-12.
 *
 * Median is the correct metric for "is the bridge typically slow right now?":
 * one outlier moves it by zero, sustained degradation moves it sharply.
 * Cold-start latency doesn't trigger refusals; genuine bridge stress does.
 *
 * ## State machine
 *
 *   healthy   (median < 500ms)   →  100ms throttle, no refusals
 *   normal    (median < 1000ms)  →  200ms throttle, no refusals
 *   stressed  (median < 2000ms)  →  500ms throttle, no refusals
 *   degraded  (median ≥ 2000ms)  →  1000ms throttle, refuse expensive ops
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
/**
 * Samples older than this are excluded from bridge-state classification.
 * Without time-decay, a single slow call (e.g. a 13s cold-start search) would
 * dominate p95 for as long as it stayed in the rolling window — even if every
 * subsequent call was fast. The state should reflect *recent* evidence, not
 * cumulative history.
 */
const SAMPLE_DECAY_MS = 5 * 60 * 1000;

interface LatencySample {
    readonly ms: number;
    readonly timestamp: number;
}

const latencyWindow: LatencySample[] = [];
let lastCallTime = 0;

const sleepBuffer = new Int32Array(new SharedArrayBuffer(4));

export type BridgeState = 'healthy' | 'normal' | 'stressed' | 'degraded';

interface ThrottleConfig {
    readonly throttleMs: number;
    readonly refuseExpensive: boolean;
}

const STATE_CONFIG: Record<BridgeState, ThrottleConfig> = {
    healthy:  { throttleMs: 100,  refuseExpensive: false },
    normal:   { throttleMs: 200,  refuseExpensive: false },
    stressed: { throttleMs: 500,  refuseExpensive: false },
    degraded: { throttleMs: 1000, refuseExpensive: true  },
};

function median(samples: number[]): number {
    if (samples.length === 0) return 0;
    const sorted = [...samples].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    if (sorted.length % 2 === 1) {
        return sorted[mid] ?? 0;
    }
    const lo = sorted[mid - 1] ?? 0;
    const hi = sorted[mid] ?? 0;
    return (lo + hi) / 2;
}

function recentLatencies(): number[] {
    const cutoff = Date.now() - SAMPLE_DECAY_MS;
    return latencyWindow.filter(s => s.timestamp >= cutoff).map(s => s.ms);
}

function classifyBridgeState(medianLatency: number, sampleCount: number): BridgeState {
    // With fewer than 3 *recent* samples we can't classify reliably — default
    // to normal. This means a fresh session doesn't get unnecessary throttle,
    // and stale samples don't cause persistent false-positive degradation.
    if (sampleCount < 3) return 'normal';
    if (medianLatency < 500) return 'healthy';
    if (medianLatency < 1000) return 'normal';
    if (medianLatency < 2000) return 'stressed';
    return 'degraded';
}

export function currentBridgeState(): BridgeState {
    const recent = recentLatencies();
    return classifyBridgeState(median(recent), recent.length);
}

export function currentMedianLatency(): number {
    return median(recentLatencies());
}

export function currentThrottleMs(): number {
    return STATE_CONFIG[currentBridgeState()].throttleMs;
}

export function isExpensiveOperationAllowed(): boolean {
    return !STATE_CONFIG[currentBridgeState()].refuseExpensive;
}

/**
 * Blocks the calling thread until the throttle slot opens up.
 * Uses Atomics.wait — no CPU burn, no subprocess overhead.
 */
export function waitForSlot(): void {
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

export function recordLatency(ms: number): void {
    latencyWindow.push({ ms, timestamp: Date.now() });
    if (latencyWindow.length > WINDOW_CAPACITY) {
        latencyWindow.shift();
    }
}

export function bridgeStateSnapshot() {
    const state = currentBridgeState();
    const recent = recentLatencies();
    return {
        state,
        medianLatencyMs: median(recent),
        recentSampleCount: recent.length,
        totalSampleCount: latencyWindow.length,
        windowCapacity: WINDOW_CAPACITY,
        sampleDecayMs: SAMPLE_DECAY_MS,
        throttleMs: STATE_CONFIG[state].throttleMs,
        refusingExpensiveOps: STATE_CONFIG[state].refuseExpensive,
    };
}

/**
 * Test-only: reset the throttle state. Not exported via index.ts.
 */
export function _resetForTests(): void {
    latencyWindow.length = 0;
    lastCallTime = 0;
}
