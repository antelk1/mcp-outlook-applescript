/**
 * Domain types for Outlook contacts: records, emails, phones, and addresses.
 */

/** Distinguishes individual contacts from distribution lists. */
export const ContactType = {
    Person: 0,
    DistributionList: 1,
} as const;
export type ContactTypeValue = (typeof ContactType)[keyof typeof ContactType];

/** Lightweight contact representation used in list and search results. */
export interface ContactSummary {
    readonly id: number;
    readonly folderId: number;
    readonly displayName: string | null;
    readonly sortName: string | null;
    readonly contactType: ContactTypeValue;
}

/** Complete contact record with name fields, communication details, and notes. */
export interface Contact extends ContactSummary {
    readonly firstName: string | null;
    readonly lastName: string | null;
    readonly middleName: string | null;
    readonly nickname: string | null;
    readonly company: string | null;
    readonly jobTitle: string | null;
    readonly department: string | null;
    readonly emails: readonly ContactEmail[];
    readonly phones: readonly ContactPhone[];
    readonly addresses: readonly ContactAddress[];
    readonly notes: string | null;
}

/** An email address associated with a contact, tagged by usage type. */
export interface ContactEmail {
    readonly type: EmailType;
    readonly address: string;
}

/** Classification for a contact's email address. */
export const EmailType = {
    Work: 'work',
    Home: 'home',
    Other: 'other',
} as const;
export type EmailType = (typeof EmailType)[keyof typeof EmailType];

/** A phone number associated with a contact, tagged by usage type. */
export interface ContactPhone {
    readonly type: PhoneType;
    readonly number: string;
}

/** Classification for a contact's phone number. */
export const PhoneType = {
    Work: 'work',
    Home: 'home',
    Mobile: 'mobile',
    Fax: 'fax',
    Other: 'other',
} as const;
export type PhoneType = (typeof PhoneType)[keyof typeof PhoneType];

/** A postal address associated with a contact, tagged by usage type. */
export interface ContactAddress {
    readonly type: AddressType;
    readonly street: string | null;
    readonly city: string | null;
    readonly state: string | null;
    readonly postalCode: string | null;
    readonly country: string | null;
}

/** Classification for a contact's postal address. */
export const AddressType = {
    Work: 'work',
    Home: 'home',
    Other: 'other',
} as const;
export type AddressType = (typeof AddressType)[keyof typeof AddressType];
