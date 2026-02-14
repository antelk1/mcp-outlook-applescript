import { randomUUID } from 'node:crypto';
import type { OperationType, TargetType, ApprovalToken, ValidationResult } from './types.js';

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
    private readonly tokens = new Map<string, ApprovalToken>();
    private readonly ttlMs: number;

    constructor(ttlMs: number = DEFAULT_TTL_MS) {
        this.ttlMs = ttlMs;
    }

    /** Creates and stores a new approval token for the given operation and target. */
    generateToken(params: {
        operation: OperationType;
        targetType: TargetType;
        targetId: number;
        targetHash: string;
        metadata?: Record<string, unknown>;
    }): ApprovalToken {
        if (this.tokens.size > CLEANUP_THRESHOLD) {
            this.cleanupExpiredTokens();
        }
        const now = Date.now();
        const token: ApprovalToken = {
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
    validateToken(tokenId: string, operation: OperationType, targetId: number): ValidationResult {
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
    consumeToken(tokenId: string, operation: OperationType, targetId: number): ValidationResult {
        const result = this.validateToken(tokenId, operation, targetId);
        if (result.valid) {
            this.tokens.delete(tokenId);
        }
        return result;
    }

    /** Purges all expired tokens from the in-memory store. */
    cleanupExpiredTokens(): void {
        const now = Date.now();
        for (const [tokenId, token] of this.tokens) {
            if (now > token.expiresAt) {
                this.tokens.delete(tokenId);
            }
        }
    }

    /** Number of tokens currently held (includes expired tokens not yet purged). */
    get size(): number {
        return this.tokens.size;
    }
}
