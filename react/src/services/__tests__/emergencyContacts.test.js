import { describe, it, expect } from 'vitest';
import { contactListsEqual } from '../emergencyContacts';

const contact = (over = {}) => ({
    id: 'ec_1',
    name: 'Jane Doe',
    number: '(786)-234-8749',
    relationship: 'Parent',
    ...over
});

describe('contactListsEqual', () => {
    it('treats identical lists as equal', () => {
        expect(contactListsEqual([contact()], [contact()])).toBe(true);
        expect(contactListsEqual([], [])).toBe(true);
    });

    it('ignores surrounding whitespace so trim-only edits are not saved', () => {
        expect(contactListsEqual(
            [contact({ name: '  Jane Doe ' })],
            [contact()]
        )).toBe(true);
    });

    it('detects field edits', () => {
        expect(contactListsEqual([contact()], [contact({ name: 'Janet Doe' })])).toBe(false);
        expect(contactListsEqual([contact()], [contact({ number: '(305)-111-2222' })])).toBe(false);
        expect(contactListsEqual([contact()], [contact({ relationship: 'Spouse' })])).toBe(false);
    });

    it('detects added and removed contacts', () => {
        expect(contactListsEqual([contact()], [])).toBe(false);
        expect(contactListsEqual([], [contact()])).toBe(false);
        expect(contactListsEqual(
            [contact()],
            [contact(), contact({ id: 'ec_2' })]
        )).toBe(false);
    });

    it('ignores non-visible fields like dateAdded', () => {
        expect(contactListsEqual(
            [contact({ dateAdded: '2026-01-01T00:00:00Z' })],
            [contact({ dateAdded: '2026-06-30T00:00:00Z' })]
        )).toBe(true);
    });

    it('tolerates non-array input', () => {
        expect(contactListsEqual(null, [])).toBe(true);
        expect(contactListsEqual(undefined, [contact()])).toBe(false);
    });
});
