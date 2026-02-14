import { executeAppleScriptOrThrow } from './executor.js';
import { LIST_ACCOUNTS, GET_DEFAULT_ACCOUNT, listMailFoldersByAccounts } from './account-scripts.js';
import { parseAccounts, parseDefaultAccountId, parseFoldersWithAccount, } from './parser.js';
import type { AppleScriptAccountRow, AppleScriptFolderWithAccountRow } from './parser.js';

/** Contract for querying Outlook account and folder metadata. */
export interface IAccountRepository {
    /** Retrieves all configured Outlook accounts (Exchange, IMAP, POP). */
    listAccounts(): AppleScriptAccountRow[];
    /** Returns the numeric ID of the default account, or null if none exists. */
    getDefaultAccountId(): number | null;
    /** Fetches mail folders belonging to the specified accounts. */
    listMailFoldersByAccounts(accountIds: number[]): AppleScriptFolderWithAccountRow[];
}

/** Queries Outlook account and folder data via AppleScript. */
export class AccountRepository implements IAccountRepository {
    /** Runs the list-accounts script and parses the result into account rows. */
    listAccounts(): AppleScriptAccountRow[] {
        const output = executeAppleScriptOrThrow(LIST_ACCOUNTS);
        return parseAccounts(output);
    }

    /** Runs the default-account script and extracts its numeric ID. */
    getDefaultAccountId(): number | null {
        const output = executeAppleScriptOrThrow(GET_DEFAULT_ACCOUNT);
        return parseDefaultAccountId(output);
    }

    /**
     * Fetches mail folders for a set of accounts.
     * @param accountIds - Outlook account IDs to query folders for.
     * @returns Folder rows tagged with their parent account ID.
     */
    listMailFoldersByAccounts(accountIds: number[]): AppleScriptFolderWithAccountRow[] {
        if (accountIds.length === 0) {
            return [];
        }
        const script = listMailFoldersByAccounts(accountIds);
        const output = executeAppleScriptOrThrow(script);
        return parseFoldersWithAccount(output);
    }
}

/** Creates a new AccountRepository instance. */
export function createAccountRepository(): IAccountRepository {
    return new AccountRepository();
}
