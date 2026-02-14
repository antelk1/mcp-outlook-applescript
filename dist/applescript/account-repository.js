import { executeAppleScriptOrThrow } from './executor.js';
import { LIST_ACCOUNTS, GET_DEFAULT_ACCOUNT, listMailFoldersByAccounts } from './account-scripts.js';
import { parseAccounts, parseDefaultAccountId, parseFoldersWithAccount, } from './parser.js';
/** Queries Outlook account and folder data via AppleScript. */
export class AccountRepository {
    /** Runs the list-accounts script and parses the result into account rows. */
    listAccounts() {
        const output = executeAppleScriptOrThrow(LIST_ACCOUNTS);
        return parseAccounts(output);
    }
    /** Runs the default-account script and extracts its numeric ID. */
    getDefaultAccountId() {
        const output = executeAppleScriptOrThrow(GET_DEFAULT_ACCOUNT);
        return parseDefaultAccountId(output);
    }
    /**
     * Fetches mail folders for a set of accounts.
     * @param accountIds - Outlook account IDs to query folders for.
     * @returns Folder rows tagged with their parent account ID.
     */
    listMailFoldersByAccounts(accountIds) {
        if (accountIds.length === 0) {
            return [];
        }
        const script = listMailFoldersByAccounts(accountIds);
        const output = executeAppleScriptOrThrow(script);
        return parseFoldersWithAccount(output);
    }
}
/** Creates a new AccountRepository instance. */
export function createAccountRepository() {
    return new AccountRepository();
}
