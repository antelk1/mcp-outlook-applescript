/** Destructive mailbox operations that require two-phase approval. */
export type OperationType = 'delete_email' | 'move_email' | 'archive_email' | 'junk_email' | 'delete_folder' | 'empty_folder' | 'batch_delete_emails' | 'batch_move_emails';

/** Resource categories that approval tokens can target. */
export type TargetType = 'email' | 'folder';

/** Time-limited, single-use authorization for one destructive operation on one target. */
export interface ApprovalToken {
    readonly tokenId: string;
    readonly operation: OperationType;
    readonly targetType: TargetType;
    readonly targetId: number;
    readonly targetHash: string;
    readonly createdAt: number;
    readonly expiresAt: number;
    readonly metadata: Readonly<Record<string, unknown>>;
}

/** Specific failure reasons when a token does not pass validation. */
export type ValidationErrorReason = 'EXPIRED' | 'NOT_FOUND' | 'OPERATION_MISMATCH' | 'TARGET_MISMATCH' | 'TARGET_CHANGED' | 'ALREADY_CONSUMED';

/** Outcome of validating or consuming an approval token. */
export interface ValidationResult {
    readonly valid: boolean;
    readonly error?: ValidationErrorReason;
    readonly token?: ApprovalToken;
}
