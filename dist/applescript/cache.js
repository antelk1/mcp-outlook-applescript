/**
 * Generic TTL cache for AppleScript query results.
 *
 * Reduces AppleEvent volume by serving repeat reads from memory when the
 * underlying state is unlikely to have changed. Each entry has its own TTL —
 * folder structure rarely changes (5min), account list almost never (1hr),
 * recent message metadata can change quickly (5min).
 *
 * Keys are caller-supplied strings; values are arbitrary. The cache is not
 * concerned with what's stored, only when it's stale.
 */
export class TtlCache {
    defaultTtlMs;
    entries = new Map();
    constructor(defaultTtlMs) {
        this.defaultTtlMs = defaultTtlMs;
    }
    get(key) {
        const entry = this.entries.get(key);
        if (entry == null)
            return undefined;
        if (Date.now() >= entry.expiresAt) {
            this.entries.delete(key);
            return undefined;
        }
        return entry.value;
    }
    set(key, value, ttlMs) {
        const ttl = ttlMs ?? this.defaultTtlMs;
        this.entries.set(key, {
            value,
            expiresAt: Date.now() + ttl,
        });
    }
    invalidate(key) {
        this.entries.delete(key);
    }
    invalidateAll() {
        this.entries.clear();
    }
    size() {
        return this.entries.size;
    }
}
