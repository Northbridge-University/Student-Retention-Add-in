/*
 * Tests for the scan-filter derivation helpers (commands/src/scan-filter-derivation.js):
 * resolving today's LDA sheet name and deriving the extension's scan-filter
 * settings (minimum Days Out, failing list presence) from sheet values.
 */
import { describe, it, expect } from 'vitest';
import {
    todaysLdaSheetBaseName,
    pickTodaysLdaSheet,
    deriveScanFilterFromGrid
} from '../src/scan-filter-derivation.js';

const JUNE_12 = new Date(2026, 5, 12);

describe('todaysLdaSheetBaseName', () => {
    it('formats without zero padding, matching ldaProcessor date mode', () => {
        expect(todaysLdaSheetBaseName(JUNE_12)).toBe('LDA 6-12-2026');
        expect(todaysLdaSheetBaseName(new Date(2026, 10, 3))).toBe('LDA 11-3-2026');
    });
});

describe('pickTodaysLdaSheet', () => {
    it('finds the exact base name', () => {
        expect(pickTodaysLdaSheet(['Master List', 'LDA 6-12-2026'], JUNE_12)).toBe('LDA 6-12-2026');
    });

    it('finds zero-padded and slash date variants', () => {
        expect(pickTodaysLdaSheet(['LDA 06-12-2026'], JUNE_12)).toBe('LDA 06-12-2026');
        expect(pickTodaysLdaSheet(['LDA 6/12/2026'], JUNE_12)).toBe('LDA 6/12/2026');
    });

    it('prefers the highest "(n)" copy over the base sheet', () => {
        const names = ['LDA 6-12-2026', 'LDA 6-12-2026 (2)', 'LDA 6-12-2026 (3)'];
        expect(pickTodaysLdaSheet(names, JUNE_12)).toBe('LDA 6-12-2026 (3)');
    });

    it('returns null when only other days exist', () => {
        expect(pickTodaysLdaSheet(['LDA 6-11-2026', 'Master List'], JUNE_12)).toBeNull();
    });

    it('does not match unrelated sheets with similar prefixes', () => {
        expect(pickTodaysLdaSheet(['LDA 6-12-2026 backup', 'LDA 6-12-2026 (two)'], JUNE_12)).toBeNull();
    });
});

describe('deriveScanFilterFromGrid', () => {
    const HEADERS = ['Student Name', 'Grade', 'Days Out', 'Outreach'];
    const FAILING_TITLE_ROW = ['Failing Students (Active)', '', '', ''];
    const ATTENDANCE_TITLE_ROW = ['Low Attendance Students', '', '', ''];
    const BLANK = ['', '', '', ''];

    it('uses the minimum Days Out of the main table', () => {
        const grid = [
            HEADERS,
            ['A', 80, 12, ''],
            ['B', 70, 6, ''],
            ['C', 90, 4, '']
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: '>=4', includeFailing: false });
    });

    it('detects the failing section and excludes it from the minimum', () => {
        const grid = [
            HEADERS,
            ['A', 80, 9, ''],
            ['B', 70, 5, ''],
            BLANK, BLANK, BLANK,
            FAILING_TITLE_ROW,
            HEADERS,
            ['F1', 40, 1, ''],
            ['F2', 55, 0, '']
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: '>=5', includeFailing: true });
    });

    it('ignores the attendance section without flagging failing', () => {
        const grid = [
            HEADERS,
            ['A', 80, 7, ''],
            BLANK, BLANK, BLANK,
            ATTENDANCE_TITLE_ROW,
            HEADERS,
            ['L1', 80, 2, '']
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: '>=7', includeFailing: false });
    });

    it('matches the Days Out header case- and space-insensitively', () => {
        const grid = [
            ['student name', 'GRADE', 'DAYSOUT', 'outreach'],
            ['A', 80, 3, '']
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: '>=3', includeFailing: false });
    });

    it('floors fractional minimums for the extension filter parser', () => {
        const grid = [
            HEADERS,
            ['A', 80, 4.5, '']
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: '>=4', includeFailing: false });
    });

    it('returns null daysOutFilter for a headers-only sheet', () => {
        expect(deriveScanFilterFromGrid([HEADERS])).toEqual({ daysOutFilter: null, includeFailing: false });
    });

    it('returns null daysOutFilter when no Days Out column exists', () => {
        const grid = [
            ['Student Name', 'Grade'],
            ['A', 80]
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: null, includeFailing: false });
    });

    it('skips non-numeric Days Out cells in the main table', () => {
        const grid = [
            HEADERS,
            ['A', 80, 'n/a', ''],
            ['B', 70, 6, '']
        ];
        expect(deriveScanFilterFromGrid(grid)).toEqual({ daysOutFilter: '>=6', includeFailing: false });
    });

    it('handles an empty grid', () => {
        expect(deriveScanFilterFromGrid([])).toEqual({ daysOutFilter: null, includeFailing: false });
    });
});
