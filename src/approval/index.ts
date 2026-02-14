/** Two-phase approval system for destructive mailbox operations. */
export {
    type OperationType,
    type TargetType,
    type ApprovalToken,
    type ValidationErrorReason,
    type ValidationResult,
} from './types.js';
export { hashEmailForApproval, hashFolderForApproval } from './hash.js';
export { ApprovalTokenManager } from './token-manager.js';
