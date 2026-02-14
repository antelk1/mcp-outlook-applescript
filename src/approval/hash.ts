import { createHash } from 'node:crypto';

/**
 * Produces a truncated SHA-256 fingerprint of an email's key properties.
 * Used to detect modifications between the prepare and confirm steps.
 */
export function hashEmailForApproval(email: {
    id: number;
    subject: string | null;
    folderId: number;
    timeReceived: number | null;
}): string {
    return createHash('sha256')
        .update(`${email.id}:${email.subject ?? ''}:${email.folderId}:${email.timeReceived ?? 0}`)
        .digest('hex')
        .slice(0, 16);
}

/**
 * Produces a truncated SHA-256 fingerprint of a folder's key properties.
 * Used to detect modifications between the prepare and confirm steps.
 */
export function hashFolderForApproval(folder: {
    id: number;
    name: string | null;
    messageCount: number;
}): string {
    return createHash('sha256')
        .update(`${folder.id}:${folder.name ?? ''}:${folder.messageCount}`)
        .digest('hex')
        .slice(0, 16);
}
