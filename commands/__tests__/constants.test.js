import { describe, it, expect } from 'vitest';
import {
    CONSTANTS,
    parseDate,
    normalizeName,
    formatToLastFirst,
} from '../src/constants.js';

describe('parseDate', () => {
    it('returns null for null/undefined/empty', () => {
        expect(parseDate(null)).toBeNull();
        expect(parseDate(undefined)).toBeNull();
        expect(parseDate('')).toBeNull();
        expect(parseDate(0)).toBeNull();
    });

    it('passes through a Date instance unchanged', () => {
        const d = new Date('2024-06-15T00:00:00Z');
        expect(parseDate(d)).toBe(d);
    });

    it('parses Excel serial numbers above the 1970 threshold', () => {
        // 25569 == 1970-01-01; 44927 == 2023-01-01
        const result = parseDate(44927);
        expect(result).toBeInstanceOf(Date);
        expect(result.getUTCFullYear()).toBe(2023);
        expect(result.getUTCMonth()).toBe(0);
        expect(result.getUTCDate()).toBe(1);
    });

    it('rejects Excel serial numbers below the 1970 threshold', () => {
        expect(parseDate(1000)).toBeNull();
        expect(parseDate(25569)).toBeNull(); // boundary: not > 25569
    });

    it('parses ISO date strings', () => {
        const result = parseDate('2024-03-15');
        expect(result).toBeInstanceOf(Date);
        expect(result.getUTCFullYear()).toBe(2024);
    });

    it('returns null for unparseable strings', () => {
        expect(parseDate('not a date')).toBeNull();
        expect(parseDate('xyz')).toBeNull();
    });

    it('returns null for unsupported types', () => {
        expect(parseDate(true)).toBeNull();
        expect(parseDate({})).toBeNull();
        expect(parseDate([])).toBeNull();
    });
});

describe('normalizeName', () => {
    it('returns empty string for non-string / null inputs', () => {
        expect(normalizeName(null)).toBe('');
        expect(normalizeName(undefined)).toBe('');
        expect(normalizeName(123)).toBe('');
        expect(normalizeName({})).toBe('');
    });

    it('returns empty string for empty/whitespace input', () => {
        expect(normalizeName('')).toBe('');
        expect(normalizeName('   ')).toBe('');
    });

    it('flips "Last, First" to "first last"', () => {
        expect(normalizeName('Smith, John')).toBe('john smith');
        expect(normalizeName('DOE, JANE')).toBe('jane doe');
    });

    it('lowercases and trims "First Last" without flipping', () => {
        expect(normalizeName('John Smith')).toBe('john smith');
        expect(normalizeName('  John Smith  ')).toBe('john smith');
    });

    it('lowercases a single name', () => {
        expect(normalizeName('Cher')).toBe('cher');
    });
});

describe('formatToLastFirst', () => {
    it('returns empty string for non-string / null inputs', () => {
        expect(formatToLastFirst(null)).toBe('');
        expect(formatToLastFirst(undefined)).toBe('');
        expect(formatToLastFirst(42)).toBe('');
    });

    it('keeps "Last, First" form intact (with normalized spacing)', () => {
        expect(formatToLastFirst('Smith, John')).toBe('Smith, John');
        expect(formatToLastFirst('Smith,John')).toBe('Smith, John');
        expect(formatToLastFirst('  Smith  ,  John  ')).toBe('Smith, John');
    });

    it('converts "First Last" to "Last, First"', () => {
        expect(formatToLastFirst('John Smith')).toBe('Smith, John');
    });

    it('treats the last token as the surname for multi-word first names', () => {
        expect(formatToLastFirst('John Doe Smith')).toBe('Smith, John Doe');
    });

    it('returns a single-token name unchanged (after trim)', () => {
        expect(formatToLastFirst('Cher')).toBe('Cher');
        expect(formatToLastFirst('  Cher  ')).toBe('Cher');
    });
});

describe('CONSTANTS', () => {
    it('exposes the expected top-level keys', () => {
        expect(CONSTANTS.MASTER_LIST_SHEET).toBe('Master List');
        expect(CONSTANTS.HISTORY_SHEET).toBe('Student History');
        expect(CONSTANTS.SETTINGS_KEY).toBe('studentRetentionSettings');
    });

    it('exposes alias arrays for student name / id / number', () => {
        expect(Array.isArray(CONSTANTS.STUDENT_NAME_COLS)).toBe(true);
        expect(CONSTANTS.STUDENT_NAME_COLS.length).toBeGreaterThan(0);
        expect(Array.isArray(CONSTANTS.STUDENT_ID_COLS)).toBe(true);
        expect(Array.isArray(CONSTANTS.STUDENT_NUMBER_COLS)).toBe(true);
    });

    it('COLUMN_MAPPINGS values are all non-empty alias arrays', () => {
        const entries = Object.entries(CONSTANTS.COLUMN_MAPPINGS);
        expect(entries.length).toBeGreaterThan(0);
        for (const [key, aliases] of entries) {
            expect(Array.isArray(aliases), `${key} should be an array`).toBe(true);
            expect(aliases.length, `${key} should have at least one alias`).toBeGreaterThan(0);
            for (const alias of aliases) {
                expect(typeof alias, `aliases in ${key} should be strings`).toBe('string');
            }
        }
    });

    it('includes the column mappings consumed by the rest of the codebase', () => {
        // These keys are referenced from master-list-import / conditional-formatting / ribbon-actions —
        // a missing key would break those modules silently (findColumnIndex would just return -1).
        const required = [
            'grade', 'gradeBook', 'lastLda', 'daysOut', 'assigned',
            'courseMissingAssignments', 'hold', 'primaryPhone', 'otherPhone',
            'nextAssignmentDue', 'lastCourseGrade', 'expectedStartDate',
            'attendance', 'letterGrade', 'lastCourseLetterGrade', 'enrollGpa',
            'courseStart', 'courseEnd',
        ];
        for (const key of required) {
            expect(CONSTANTS.COLUMN_MAPPINGS[key], `COLUMN_MAPPINGS.${key} missing`).toBeDefined();
        }
    });
});
