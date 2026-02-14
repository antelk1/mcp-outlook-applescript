import { randomUUID } from 'node:crypto';
// =============================================================================
// Constants
// =============================================================================
/** Tokens expire after 5 minutes by default. */
const DEFAULT_TTL_MS = 5 * 60 * 1000;
/** Triggers garbage collection of expired tokens when the store exceeds this size. */
const CLEANUP_THRESHOLD = 100;
// =============================================================================
// Token Manager
// =============================================================================
/**
 * In-memory store for single-use approval tokens.
 *
 * Each token authorizes exactly one destructive operation on one target.
 * Tokens are automatically purged once the store exceeds CLEANUP_THRESHOLD.
 */
export class ApprovalTokenManager {
    tokens = new Map();
    ttlMs;
    constructor(ttlMs = DEFAULT_TTL_MS) {
        this.ttlMs = ttlMs;
    }
    /** Creates and stores a new approval token for the given operation and target. */
    generateToken(params) {
        if (this.tokens.size > CLEANUP_THRESHOLD) {
            this.cleanupExpiredTokens();
        }
        const now = Date.now();
        const token = {
            tokenId: randomUUID(),
            operation: params.operation,
            targetType: params.targetType,
            targetId: params.targetId,
            targetHash: params.targetHash,
            createdAt: now,
            expiresAt: now + this.ttlMs,
            metadata: Object.freeze({ ...params.metadata }),
        };
        this.tokens.set(token.tokenId, token);
        return token;
    }
    /** Checks a token's validity without consuming it. Verifies existence, expiry, operation, and target. */
    validateToken(tokenId, operation, targetId) {
        const token = this.tokens.get(tokenId);
        if (token == null) {
            return { valid: false, error: 'NOT_FOUND' };
        }
        if (Date.now() > token.expiresAt) {
            return { valid: false, error: 'EXPIRED' };
        }
        if (token.operation !== operation) {
            return { valid: false, error: 'OPERATION_MISMATCH' };
        }
        if (token.targetId !== targetId) {
            return { valid: false, error: 'TARGET_MISMATCH' };
        }
        return { valid: true, token };
    }
    /** Validates a token and removes it from the store on success (one-time use). */
    consumeToken(tokenId, operation, targetId) {
        const result = this.validateToken(tokenId, operation, targetId);
        if (result.valid) {
            this.tokens.delete(tokenId);
        }
        return result;
    }
    /** Purges all expired tokens from the in-memory store. */
    cleanupExpiredTokens() {
        const now = Date.now();
        for (const [tokenId, token] of this.tokens) {
            if (now > token.expiresAt) {
                this.tokens.delete(tokenId);
            }
        }
    }
    /** Number of tokens currently held (includes expired tokens not yet purged). */
    get size() {
        return this.tokens.size;
    }
}
