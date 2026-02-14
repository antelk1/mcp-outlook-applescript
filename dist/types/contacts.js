/**
 * Domain types for Outlook contacts: records, emails, phones, and addresses.
 */
/** Distinguishes individual contacts from distribution lists. */
export const ContactType = {
    Person: 0,
    DistributionList: 1,
};
/** Classification for a contact's email address. */
export const EmailType = {
    Work: 'work',
    Home: 'home',
    Other: 'other',
};
/** Classification for a contact's phone number. */
export const PhoneType = {
    Work: 'work',
    Home: 'home',
    Mobile: 'mobile',
    Fax: 'fax',
    Other: 'other',
};
/** Classification for a contact's postal address. */
export const AddressType = {
    Work: 'work',
    Home: 'home',
    Other: 'other',
};
