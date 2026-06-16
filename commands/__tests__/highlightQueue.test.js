import { describe, it, expect } from 'vitest';
import { coalesceHighlightRequests } from '../../shared/highlightQueue.js';
import { sheetNameDateVariants } from '../../shared/excel-helpers.js';

const base = {
    startCol: 'Student Name',
    endCol: 'Outreach',
    targetSheet: 'LDA 6-10-2026',
    color: '#92d050',
    editColumn: 'Outreach',
};

function single(syStudentId, overrides = {}) {
    return {
        ...base,
        studentName: `Student ${syStudentId}`,
        syStudentId,
        editText: `Submitted HW ${syStudentId}`,
        ...overrides,
    };
}

describe('coalesceHighlightRequests', () => {
    it('passes a lone single payload through in single shape', () => {
        const out = coalesceHighlightRequests([single('100')]);
        expect(out).toHaveLength(1);
        expect(out[0].syStudentId).toBe('100');
        expect(out[0].students).toBeUndefined();
        expect(out[0].editText).toBe('Submitted HW 100');
    });

    it('merges singles with identical sheet/format into one bulk payload', () => {
        const out = coalesceHighlightRequests([single('1'), single('2'), single('3')]);
        expect(out).toHaveLength(1);
        expect(out[0].students.map(s => s.syStudentId)).toEqual(['1', '2', '3']);
        expect(out[0].targetSheet).toBe(base.targetSheet);
        expect(out[0].color).toBe(base.color);
        // per-student editText preserved
        expect(out[0].students[1].editText).toBe('Submitted HW 2');
    });

    it('merges a bulk payload with trailing singles for the same group', () => {
        const bulk = {
            ...base,
            students: [
                { studentName: 'A', syStudentId: '1', editText: 'Submitted X' },
                { studentName: 'B', syStudentId: '2', editText: 'Submitted Y' },
            ],
        };
        const out = coalesceHighlightRequests([bulk, single('3')]);
        expect(out).toHaveLength(1);
        expect(out[0].students.map(s => s.syStudentId)).toEqual(['1', '2', '3']);
    });

    it('keeps different target sheets in separate payloads', () => {
        const out = coalesceHighlightRequests([
            single('1'),
            single('2', { targetSheet: 'Campus A' }),
        ]);
        expect(out).toHaveLength(2);
        expect(out.map(p => p.targetSheet).sort()).toEqual(['Campus A', 'LDA 6-10-2026']);
    });

    it('keeps different colors / columns in separate payloads', () => {
        const out = coalesceHighlightRequests([
            single('1'),
            single('2', { color: '#FFFF00' }),
            single('3', { endCol: 'Phone' }),
        ]);
        expect(out).toHaveLength(3);
    });

    it('dedupes the same student within a group, latest request wins', () => {
        const out = coalesceHighlightRequests([
            single('1', { editText: 'old text' }),
            single('1', { editText: 'new text' }),
        ]);
        expect(out).toHaveLength(1);
        expect(out[0].syStudentId).toBe('1');
        expect(out[0].editText).toBe('new text');
    });

    it('does not confuse numeric and string column indexes when grouping', () => {
        const out = coalesceHighlightRequests([
            single('1', { startCol: 0, endCol: 5 }),
            single('2', { startCol: '0', endCol: '5' }),
        ]);
        expect(out).toHaveLength(2);
    });

    it('passes malformed payloads through unmerged so the handler reports errors', () => {
        const noSheet = single('1', { targetSheet: undefined });
        const noStudent = { ...base, studentName: 'Ghost' }; // no syStudentId
        const out = coalesceHighlightRequests([noSheet, noStudent, single('2')]);
        expect(out).toHaveLength(3);
        expect(out.filter(p => p.targetSheet === undefined)).toHaveLength(1);
    });

    it('ignores null/non-object entries', () => {
        const out = coalesceHighlightRequests([null, undefined, 'junk', single('1')]);
        expect(out).toHaveLength(1);
        expect(out[0].syStudentId).toBe('1');
    });

    it('omits color/editColumn keys when the source payloads omitted them', () => {
        const out = coalesceHighlightRequests([
            single('1', { color: undefined, editColumn: undefined, editText: undefined }),
            single('2', { color: undefined, editColumn: undefined, editText: undefined }),
        ]);
        expect(out).toHaveLength(1);
        expect('color' in out[0]).toBe(false);
        expect('editColumn' in out[0]).toBe(false);
    });
});

describe('sheetNameDateVariants', () => {
    it('returns the original name first', () => {
        expect(sheetNameDateVariants('LDA 6-10-2026')[0]).toBe('LDA 6-10-2026');
    });

    it('generates zero-padded and non-padded variants with both separators', () => {
        const variants = sheetNameDateVariants('LDA 6-10-2026');
        expect(variants).toContain('LDA 06-10-2026');
        expect(variants).toContain('LDA 6-10-2026');
        expect(variants).toContain('LDA 06/10/2026');
        expect(variants).toContain('LDA 6/10/2026');
    });

    it('round-trips: padded form generates the non-padded form', () => {
        const variants = sheetNameDateVariants('LDA 06-09-2026');
        expect(variants).toContain('LDA 6-9-2026');
    });

    it('handles names without a prefix', () => {
        const variants = sheetNameDateVariants('1/5/2026');
        expect(variants).toContain('01/05/2026');
        expect(variants).toContain('1-5-2026');
    });

    it('returns non-date names unchanged as a single variant', () => {
        expect(sheetNameDateVariants('Master List')).toEqual(['Master List']);
        expect(sheetNameDateVariants('Campus A')).toEqual(['Campus A']);
    });

    it('returns non-string input unchanged', () => {
        expect(sheetNameDateVariants(null)).toEqual([null]);
        expect(sheetNameDateVariants(undefined)).toEqual([undefined]);
    });
});
