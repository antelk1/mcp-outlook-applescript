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

interface CacheEntry<T> {
    readonly value: T;
    readonly expiresAt: number;
}

export class TtlCache<T> {
    private readonly entries = new Map<string, CacheEntry<T>>();

    constructor(private readonly defaultTtlMs: number) {}

    get(key: string): T | undefined {
        const entry = this.entries.get(key);
        if (entry == null) return undefined;
        if (Date.now() >= entry.expiresAt) {
            this.entries.delete(key);
            return undefined;
        }
        return entry.value;
    }

    set(key: string, value: T, ttlMs?: number): void {
        const ttl = ttlMs ?? this.defaultTtlMs;
        this.entries.set(key, {
            value,
            expiresAt: Date.now() + ttl,
        });
    }

    invalidate(key: string): void {
        this.entries.delete(key);
    }

    invalidateAll(): void {
        this.entries.clear();
    }

    size(): number {
        return this.entries.size;
    }
}
