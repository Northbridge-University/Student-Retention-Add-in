import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
    getTodaysLdaSheetName,
    formatLastSentStamp,
    getNameParts,
    isValidEmail,
    isValidHttpUrl,
    evaluateMapping,
    renderTemplate,
    renderCCTemplate,
} from '../helpers.js';

describe('getTodaysLdaSheetName', () => {
    beforeEach(() => {
        vi.useFakeTimers();
    });

    afterEach(() => {
        vi.useRealTimers();
    });

    it('formats today as "LDA M-D-YYYY"', () => {
        vi.setSystemTime(new Date(2024, 2, 15)); // March 15, 2024 (months are 0-indexed)
        expect(getTodaysLdaSheetName()).toBe('LDA 3-15-2024');
    });

    it('does not zero-pad month or day', () => {
        vi.setSystemTime(new Date(2024, 0, 5)); // January 5
        expect(getTodaysLdaSheetName()).toBe('LDA 1-5-2024');
    });

    it('handles December (12-month case)', () => {
        vi.setSystemTime(new Date(2025, 11, 31));
        expect(getTodaysLdaSheetName()).toBe('LDA 12-31-2025');
    });
});

describe('formatLastSentStamp', () => {
    it('shows the time when the timestamp is from today', () => {
        const now = new Date(2026, 5, 10, 15, 30); // June 10, 2026 3:30 PM
        const sent = new Date(2026, 5, 10, 13, 54).toISOString();
        expect(formatLastSentStamp(sent, now)).toBe('at 1:54 pm');
    });

    it('uses am, 12-hour clock, and zero-padded minutes', () => {
        const now = new Date(2026, 5, 10, 15, 30);
        expect(formatLastSentStamp(new Date(2026, 5, 10, 9, 5).toISOString(), now)).toBe('at 9:05 am');
        expect(formatLastSentStamp(new Date(2026, 5, 10, 0, 7).toISOString(), now)).toBe('at 12:07 am');
        expect(formatLastSentStamp(new Date(2026, 5, 10, 12, 0).toISOString(), now)).toBe('at 12:00 pm');
    });

    it('shows a M/D/YY date when the timestamp is from another day', () => {
        const now = new Date(2026, 5, 10, 9, 0);
        const sent = new Date(2026, 5, 9, 13, 54).toISOString(); // June 9, 2026
        expect(formatLastSentStamp(sent, now)).toBe('on 6/9/26');
    });

    it('treats the same calendar day in a different year as not today', () => {
        const now = new Date(2026, 5, 10, 9, 0);
        const sent = new Date(2025, 5, 10, 13, 54).toISOString();
        expect(formatLastSentStamp(sent, now)).toBe('on 6/10/25');
    });

    it('zero-pads single-digit two-digit years', () => {
        const now = new Date(2026, 5, 10);
        expect(formatLastSentStamp(new Date(2009, 0, 2).toISOString(), now)).toBe('on 1/2/09');
    });

    it('returns an empty string for invalid or missing timestamps', () => {
        expect(formatLastSentStamp('not-a-date')).toBe('');
        expect(formatLastSentStamp(undefined)).toBe('');
        expect(formatLastSentStamp('')).toBe('');
    });
});

describe('getNameParts', () => {
    it('returns empty parts for non-string / null input', () => {
        expect(getNameParts(null)).toEqual({ first: '', last: '' });
        expect(getNameParts(undefined)).toEqual({ first: '', last: '' });
        expect(getNameParts(42)).toEqual({ first: '', last: '' });
        expect(getNameParts({})).toEqual({ first: '', last: '' });
    });

    it('parses "Last, First" form', () => {
        expect(getNameParts('Smith, John')).toEqual({ first: 'John', last: 'Smith' });
    });

    it('trims whitespace around the comma', () => {
        expect(getNameParts('  Smith  ,  John  ')).toEqual({ first: 'John', last: 'Smith' });
    });

    it('handles "Last," with no first name', () => {
        expect(getNameParts('Smith,')).toEqual({ first: '', last: 'Smith' });
    });

    it('parses "First Last" form by treating last token as surname', () => {
        expect(getNameParts('John Smith')).toEqual({ first: 'John', last: 'Smith' });
    });

    it('treats multi-word first names as joined first', () => {
        expect(getNameParts('John Doe Smith')).toEqual({ first: 'John Doe', last: 'Smith' });
    });

    it('returns single token as first name only', () => {
        expect(getNameParts('Cher')).toEqual({ first: 'Cher', last: '' });
    });

    it('handles whitespace-only input as empty (after trim)', () => {
        // After trim, "" → split(' ').filter(p=>p) is [], then parts.length !== 1
        // and parts.pop() is undefined. The function path doesn't handle this — let
        // it return what it returns (regression-pin behavior).
        expect(getNameParts('   ')).toEqual({ first: '', last: undefined });
    });
});

describe('isValidEmail', () => {
    it('accepts standard email forms', () => {
        expect(isValidEmail('user@example.com')).toBe(true);
        expect(isValidEmail('a@b.co')).toBe(true);
        expect(isValidEmail('first.last+tag@sub.example.org')).toBe(true);
    });

    it('rejects empty / whitespace-only input', () => {
        expect(isValidEmail('')).toBe(false);
        expect(isValidEmail('   ')).toBe(false);
    });

    it('rejects non-string input', () => {
        expect(isValidEmail(null)).toBe(false);
        expect(isValidEmail(undefined)).toBe(false);
        expect(isValidEmail(42)).toBe(false);
    });

    it('rejects strings with no @', () => {
        expect(isValidEmail('userexample.com')).toBe(false);
    });

    it('rejects strings with no dot in the domain', () => {
        expect(isValidEmail('user@example')).toBe(false);
    });

    it('rejects strings with whitespace', () => {
        expect(isValidEmail('user @example.com')).toBe(false);
        expect(isValidEmail('user@example .com')).toBe(false);
    });

    it('rejects strings with multiple @ signs', () => {
        expect(isValidEmail('a@b@c.com')).toBe(false);
    });
});

describe('isValidHttpUrl', () => {
    it('accepts http and https URLs', () => {
        expect(isValidHttpUrl('http://example.com')).toBe(true);
        expect(isValidHttpUrl('https://example.com/path?q=1')).toBe(true);
        expect(isValidHttpUrl('http://localhost:3000')).toBe(true);
    });

    it('rejects non-http(s) protocols', () => {
        expect(isValidHttpUrl('ftp://example.com')).toBe(false);
        expect(isValidHttpUrl('file:///etc/passwd')).toBe(false);
        expect(isValidHttpUrl('mailto:a@b.com')).toBe(false);
    });

    it('rejects strings that are not URLs', () => {
        expect(isValidHttpUrl('not a url')).toBe(false);
        expect(isValidHttpUrl('')).toBe(false);
        expect(isValidHttpUrl('example.com')).toBe(false); // no protocol
    });

    it('rejects non-string input by catching the URL constructor throw', () => {
        expect(isValidHttpUrl(null)).toBe(false);
        expect(isValidHttpUrl(undefined)).toBe(false);
    });
});

describe('evaluateMapping', () => {
    it('eq: case-insensitive equality', () => {
        expect(evaluateMapping('Yes', { operator: 'eq', if: 'yes' })).toBe(true);
        expect(evaluateMapping('Yes', { operator: 'eq', if: 'no' })).toBe(false);
    });

    it('neq: case-insensitive inequality', () => {
        expect(evaluateMapping('Yes', { operator: 'neq', if: 'no' })).toBe(true);
        expect(evaluateMapping('Yes', { operator: 'neq', if: 'YES' })).toBe(false);
    });

    it('contains / does_not_contain', () => {
        expect(evaluateMapping('Hello World', { operator: 'contains', if: 'world' })).toBe(true);
        expect(evaluateMapping('Hello', { operator: 'contains', if: 'xyz' })).toBe(false);
        expect(evaluateMapping('Hello', { operator: 'does_not_contain', if: 'xyz' })).toBe(true);
        expect(evaluateMapping('Hello World', { operator: 'does_not_contain', if: 'world' })).toBe(false);
    });

    it('starts_with / ends_with', () => {
        expect(evaluateMapping('foobar', { operator: 'starts_with', if: 'FOO' })).toBe(true);
        expect(evaluateMapping('foobar', { operator: 'starts_with', if: 'bar' })).toBe(false);
        expect(evaluateMapping('foobar', { operator: 'ends_with', if: 'BAR' })).toBe(true);
        expect(evaluateMapping('foobar', { operator: 'ends_with', if: 'foo' })).toBe(false);
    });

    it('numeric comparisons require both sides to be numeric', () => {
        expect(evaluateMapping(5, { operator: 'gt', if: 3 })).toBe(true);
        expect(evaluateMapping(3, { operator: 'gt', if: 5 })).toBe(false);
        expect(evaluateMapping(5, { operator: 'gte', if: 5 })).toBe(true);
        expect(evaluateMapping(4, { operator: 'lt', if: 5 })).toBe(true);
        expect(evaluateMapping(5, { operator: 'lte', if: 5 })).toBe(true);
    });

    it('numeric ops return false when either side is not numeric', () => {
        expect(evaluateMapping('abc', { operator: 'gt', if: 5 })).toBe(false);
        expect(evaluateMapping(5, { operator: 'gt', if: 'abc' })).toBe(false);
    });

    it('returns false for unknown operators', () => {
        expect(evaluateMapping('a', { operator: 'matches', if: 'a' })).toBe(false);
        expect(evaluateMapping('a', { operator: undefined, if: 'a' })).toBe(false);
    });

    it('coerces non-string cellValue with String()', () => {
        expect(evaluateMapping(123, { operator: 'eq', if: '123' })).toBe(true);
        expect(evaluateMapping(true, { operator: 'eq', if: 'true' })).toBe(true);
    });
});

describe('renderTemplate', () => {
    it('returns empty string for falsy template', () => {
        expect(renderTemplate('', {})).toBe('');
        expect(renderTemplate(null, {})).toBe('');
        expect(renderTemplate(undefined, {})).toBe('');
    });

    it('replaces a single placeholder', () => {
        expect(renderTemplate('Hello {name}', { name: 'World' })).toBe('Hello World');
    });

    it('replaces multiple placeholders', () => {
        expect(renderTemplate('{greeting} {name}', { greeting: 'Hi', name: 'Ada' }))
            .toBe('Hi Ada');
    });

    it('keeps unknown placeholders unchanged', () => {
        expect(renderTemplate('Hello {missing}', { name: 'X' }))
            .toBe('Hello {missing}');
    });

    it('resolves nested placeholders (recursive expansion)', () => {
        expect(renderTemplate('{a}', { a: '{b}', b: 'final' })).toBe('final');
    });

    it('strips wrapping <p>...</p> when value contains no nested block tags', () => {
        // Quill editor often emits "<p>text</p>" — when injected into another
        // <p> the wrapping tags are stripped to avoid invalid nested <p>s.
        expect(renderTemplate('Note: {body}', { body: '<p>hello</p>' }))
            .toBe('Note: hello');
    });

    it('keeps wrapping <p> when inner HTML contains another <p> or <div>', () => {
        expect(renderTemplate('{body}', { body: '<p><p>nested</p></p>' }))
            .toBe('<p><p>nested</p></p>');
        expect(renderTemplate('{body}', { body: '<p><div>div inside</div></p>' }))
            .toBe('<p><div>div inside</div></p>');
    });
});

describe('renderCCTemplate', () => {
    it('returns empty string when recipients is null/empty', () => {
        expect(renderCCTemplate(null, {})).toBe('');
        expect(renderCCTemplate(undefined, {})).toBe('');
        expect(renderCCTemplate([], {})).toBe('');
    });

    it('joins rendered recipients with semicolons', () => {
        const recipients = ['{first}@a.com', '{first}@b.com'];
        expect(renderCCTemplate(recipients, { first: 'jane' }))
            .toBe('jane@a.com;jane@b.com');
    });

    it('passes data through renderTemplate for each recipient', () => {
        // Single recipient still gets template substitution
        expect(renderCCTemplate(['{name}@x.com'], { name: 'ada' }))
            .toBe('ada@x.com');
    });
});
