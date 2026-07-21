import { describe, it, expect } from 'vitest';
import {
    sanitizeRiskIndexSettings,
    DEFAULT_RISK_INDEX_SETTINGS,
    todayKey,
    normalizeLdaValue,
    buildDailySnapshot,
    mergeSnapshot,
    countRiskEpisodes,
    countRiskEpisodesForStudent,
    lookupRiskCount,
    parseLdaSheetName,
    resolveSnapshotIndices,
    mergeImportedDays,
    historyToCsv,
} from '../riskIndex.js';

describe('sanitizeRiskIndexSettings', () => {
    it('returns defaults for missing/invalid input', () => {
        expect(sanitizeRiskIndexSettings(undefined)).toEqual(DEFAULT_RISK_INDEX_SETTINGS);
        expect(sanitizeRiskIndexSettings(null)).toEqual(DEFAULT_RISK_INDEX_SETTINGS);
        expect(sanitizeRiskIndexSettings('nope')).toEqual(DEFAULT_RISK_INDEX_SETTINGS);
    });

    it('keeps explicit values', () => {
        expect(sanitizeRiskIndexSettings({ enabled: false, threshold: 10, showColumn: false }))
            .toEqual({ enabled: false, threshold: 10, showColumn: false });
    });

    it('falls back to default threshold for invalid numbers', () => {
        expect(sanitizeRiskIndexSettings({ threshold: 'abc' }).threshold).toBe(14);
        expect(sanitizeRiskIndexSettings({ threshold: 0 }).threshold).toBe(14);
        expect(sanitizeRiskIndexSettings({ threshold: -3 }).threshold).toBe(14);
        expect(sanitizeRiskIndexSettings({ threshold: 14.9 }).threshold).toBe(14);
    });

    it('fills only missing fields with defaults', () => {
        expect(sanitizeRiskIndexSettings({ enabled: false }))
            .toEqual({ enabled: false, threshold: 14, showColumn: true });
    });
});

describe('todayKey', () => {
    it('formats a local calendar date as YYYY-MM-DD', () => {
        expect(todayKey(new Date(2026, 6, 21))).toBe('2026-07-21');
        expect(todayKey(new Date(2026, 0, 5))).toBe('2026-01-05');
    });
});

describe('normalizeLdaValue', () => {
    it('returns null for empty values', () => {
        expect(normalizeLdaValue(null)).toBeNull();
        expect(normalizeLdaValue(undefined)).toBeNull();
        expect(normalizeLdaValue('')).toBeNull();
        expect(normalizeLdaValue('   ')).toBeNull();
    });

    it('converts Excel date serials without timezone drift', () => {
        // 45000 = 2023-03-15 (44927 = 2023-01-01)
        expect(normalizeLdaValue(45000)).toBe('2023-03-15');
        expect(normalizeLdaValue(44927)).toBe('2023-01-01');
    });

    it('parses common date strings', () => {
        expect(normalizeLdaValue('7/10/2026')).toBe('2026-07-10');
        expect(normalizeLdaValue('12/1/2025')).toBe('2025-12-01');
    });

    it('keeps ISO strings as-is', () => {
        expect(normalizeLdaValue('2026-07-10')).toBe('2026-07-10');
    });

    it('keeps unparseable strings verbatim so they still key episodes', () => {
        expect(normalizeLdaValue('N/A')).toBe('N/A');
    });

    it('stringifies out-of-range numbers instead of treating them as serials', () => {
        expect(normalizeLdaValue(7)).toBe('7');
    });
});

describe('buildDailySnapshot', () => {
    const headers = ['SyStudentID', 'Student Name', 'LDA', 'Days Out'];
    const idx = { idIdx: 0, ldaIdx: 2, daysOutIdx: 3 };

    it('captures id, normalized LDA, and days out per student', () => {
        const values = [
            headers,
            [101, 'Ada', 45000, 14],
            ['102 ', 'Ben', '7/10/2026', 3],
        ];
        expect(buildDailySnapshot(values, idx)).toEqual({
            '101': { lda: '2023-03-15', daysOut: 14 },
            '102': { lda: '2026-07-10', daysOut: 3 },
        });
    });

    it('skips rows with no id or non-numeric days out', () => {
        const values = [
            headers,
            ['', 'NoId', 45000, 14],
            [103, 'NoDays', 45000, ''],
            [104, 'TextDays', 45000, 'n/a'],
        ];
        expect(buildDailySnapshot(values, idx)).toEqual({});
    });

    it('handles a missing LDA column', () => {
        const values = [headers, [105, 'Eve', null, 7]];
        expect(buildDailySnapshot(values, { ...idx, ldaIdx: -1 })).toEqual({
            '105': { lda: null, daysOut: 7 },
        });
    });

    it('returns empty for missing required indices', () => {
        expect(buildDailySnapshot([headers, [1, 'x', 2, 3]], { idIdx: -1, ldaIdx: 2, daysOutIdx: 3 })).toEqual({});
        expect(buildDailySnapshot(null, idx)).toEqual({});
    });
});

describe('mergeSnapshot', () => {
    it('adds a new date entry', () => {
        const history = mergeSnapshot({ version: 1, days: {} }, '2026-07-21', { '1': { lda: 'a', daysOut: 5 } });
        expect(history).toEqual({ version: 1, days: { '2026-07-21': { '1': { lda: 'a', daysOut: 5 } } } });
    });

    it('overwrites the same day on re-runs and keeps other days', () => {
        const day1 = { '1': { lda: 'a', daysOut: 5 } };
        let history = mergeSnapshot({ version: 1, days: { '2026-07-20': day1 } }, '2026-07-21', { '1': { lda: 'a', daysOut: 6 } });
        history = mergeSnapshot(history, '2026-07-21', { '1': { lda: 'a', daysOut: 7 } });
        expect(Object.keys(history.days)).toEqual(['2026-07-20', '2026-07-21']);
        expect(history.days['2026-07-21']['1'].daysOut).toBe(7);
        expect(history.days['2026-07-20']).toEqual(day1);
    });

    it('tolerates malformed prior history', () => {
        expect(mergeSnapshot(undefined, '2026-07-21', {}).days['2026-07-21']).toEqual({});
        expect(mergeSnapshot({ bogus: true }, '2026-07-21', {}).days['2026-07-21']).toEqual({});
    });
});

describe('countRiskEpisodes', () => {
    const historyOf = (days) => ({ version: 1, days });

    it('counts consecutive days at threshold with the same LDA as one episode', () => {
        const history = historyOf({
            '2026-07-20': { '1': { lda: '2026-07-06', daysOut: 14 } },
            '2026-07-21': { '1': { lda: '2026-07-06', daysOut: 15 } },
            '2026-07-22': { '1': { lda: '2026-07-06', daysOut: 16 } },
        });
        expect(countRiskEpisodes(history, 14).get('1')).toBe(1);
    });

    it('counts a re-drop (new LDA) as a second episode', () => {
        const history = historyOf({
            '2026-07-20': { '1': { lda: '2026-07-06', daysOut: 14 } },
            '2026-07-25': { '1': { lda: '2026-07-24', daysOut: 1 } },
            '2026-08-10': { '1': { lda: '2026-07-27', daysOut: 14 } },
        });
        expect(countRiskEpisodes(history, 14).get('1')).toBe(2);
    });

    it('ignores students who never reach the threshold', () => {
        const history = historyOf({
            '2026-07-20': { '1': { lda: 'x', daysOut: 13 }, '2': { lda: 'y', daysOut: 14 } },
        });
        const counts = countRiskEpisodes(history, 14);
        expect(counts.has('1')).toBe(false);
        expect(counts.get('2')).toBe(1);
    });

    it('collapses missing-LDA hits into a single episode', () => {
        const history = historyOf({
            '2026-07-20': { '1': { lda: null, daysOut: 14 } },
            '2026-07-21': { '1': { lda: null, daysOut: 15 } },
        });
        expect(countRiskEpisodes(history, 14).get('1')).toBe(1);
    });

    it('respects a custom threshold', () => {
        const history = historyOf({
            '2026-07-20': { '1': { lda: 'a', daysOut: 10 } },
        });
        expect(countRiskEpisodes(history, 10).get('1')).toBe(1);
        expect(countRiskEpisodes(history, 14).has('1')).toBe(false);
    });

    it('handles empty or malformed history', () => {
        expect(countRiskEpisodes(undefined, 14).size).toBe(0);
        expect(countRiskEpisodes({ days: { '2026-07-20': null } }, 14).size).toBe(0);
    });
});

describe('countRiskEpisodesForStudent', () => {
    const history = {
        version: 1,
        days: {
            '2026-07-20': { '1': { lda: 'a', daysOut: 14 }, '2': { lda: 'b', daysOut: 20 } },
            '2026-07-21': { '1': { lda: 'a', daysOut: 15 } },
            '2026-08-10': { '1': { lda: 'c', daysOut: 14 } },
        },
    };

    it('matches the full-map count for a single student', () => {
        expect(countRiskEpisodesForStudent(history, '1', 14)).toBe(countRiskEpisodes(history, 14).get('1'));
        expect(countRiskEpisodesForStudent(history, '1', 14)).toBe(2);
        expect(countRiskEpisodesForStudent(history, '2', 14)).toBe(1);
    });

    it('accepts numeric ids and trims strings', () => {
        expect(countRiskEpisodesForStudent(history, 1, 14)).toBe(2);
        expect(countRiskEpisodesForStudent(history, ' 1 ', 14)).toBe(2);
    });

    it('returns 0 for unknown or empty ids', () => {
        expect(countRiskEpisodesForStudent(history, '999', 14)).toBe(0);
        expect(countRiskEpisodesForStudent(history, '', 14)).toBe(0);
        expect(countRiskEpisodesForStudent(null, '1', 14)).toBe(0);
    });
});

describe('lookupRiskCount', () => {
    it('returns the first matching id variant', () => {
        const map = new Map([['12345', 3], ['S-99', 1]]);
        expect(lookupRiskCount(map, ['12345'])).toBe(3);
        expect(lookupRiskCount(map, ['nope', 'S-99'])).toBe(1);
        expect(lookupRiskCount(map, ['nope'])).toBe(0);
        expect(lookupRiskCount(map, [])).toBe(0);
    });
});

describe('parseLdaSheetName', () => {
    it('parses date-mode LDA sheet names into date keys', () => {
        expect(parseLdaSheetName('LDA 7-21-2026')).toBe('2026-07-21');
        expect(parseLdaSheetName('LDA 12-1-2025')).toBe('2025-12-01');
        expect(parseLdaSheetName('lda 1-05-2026')).toBe('2026-01-05');
    });

    it('accepts same-day re-run suffixes', () => {
        expect(parseLdaSheetName('LDA 7-21-2026 (2)')).toBe('2026-07-21');
    });

    it('rejects non-LDA and campus-mode sheet names', () => {
        expect(parseLdaSheetName('Master List')).toBeNull();
        expect(parseLdaSheetName('Downtown Campus')).toBeNull();
        expect(parseLdaSheetName('LDA')).toBeNull();
        expect(parseLdaSheetName('LDA 13-40-2026')).toBeNull();
        expect(parseLdaSheetName('LDA 7-21-26')).toBeNull();
        expect(parseLdaSheetName('')).toBeNull();
        expect(parseLdaSheetName(null)).toBeNull();
    });
});

describe('resolveSnapshotIndices', () => {
    it('finds SyStudentID, LDA, and Days Out columns by alias', () => {
        expect(resolveSnapshotIndices(['Student Name', 'SyStudentID', 'LDA', 'Days Out']))
            .toEqual({ idIdx: 1, ldaIdx: 2, daysOutIdx: 3 });
    });

    it('normalizes header case and whitespace', () => {
        expect(resolveSnapshotIndices(['STUDENT ID', 'Last Date of Attendance', ' days  out ']))
            .toEqual({ idIdx: 0, ldaIdx: 1, daysOutIdx: 2 });
    });

    it('falls back to Student Number when no SyStudentID column exists', () => {
        expect(resolveSnapshotIndices(['Student Number', 'LDA', 'Days Out']).idIdx).toBe(0);
    });

    it('returns -1 for missing columns', () => {
        expect(resolveSnapshotIndices(['Foo', 'Bar'])).toEqual({ idIdx: -1, ldaIdx: -1, daysOutIdx: -1 });
        expect(resolveSnapshotIndices(undefined)).toEqual({ idIdx: -1, ldaIdx: -1, daysOutIdx: -1 });
    });
});

describe('mergeImportedDays', () => {
    const existing = { version: 1, days: { '2026-07-21': { '1': { lda: 'real', daysOut: 14 } } } };

    it('adds new dates and never overwrites existing snapshots', () => {
        const { history, added, skipped } = mergeImportedDays(existing, {
            '2026-07-21': { '1': { lda: 'backfilled', daysOut: 99 } },
            '2026-07-14': { '1': { lda: 'old', daysOut: 7 } },
        });
        expect(added).toEqual(['2026-07-14']);
        expect(skipped).toEqual(['2026-07-21']);
        expect(history.days['2026-07-21']['1'].lda).toBe('real');
        expect(history.days['2026-07-14']['1'].lda).toBe('old');
    });

    it('handles empty inputs', () => {
        expect(mergeImportedDays(undefined, {})).toEqual({ history: { version: 1, days: {} }, added: [], skipped: [] });
        const noop = mergeImportedDays(existing, undefined);
        expect(noop.history).toEqual(existing);
        expect(noop.added).toEqual([]);
    });
});

describe('historyToCsv', () => {
    it('flattens the history sorted by date then id, with a header row', () => {
        const csv = historyToCsv({
            version: 1,
            days: {
                '2026-07-21': { '20': { lda: '2026-07-07', daysOut: 14 }, '10': { lda: null, daysOut: 3 } },
                '2026-07-14': { '10': { lda: '2026-07-01', daysOut: 13 } },
            },
        });
        expect(csv.split('\n')).toEqual([
            'Date,SyStudentID,LDA,Days Out',
            '2026-07-14,10,2026-07-01,13',
            '2026-07-21,10,,3',
            '2026-07-21,20,2026-07-07,14',
        ]);
    });

    it('escapes cells containing commas or quotes', () => {
        const csv = historyToCsv({
            version: 1,
            days: { '2026-07-21': { 'a,b': { lda: 'x"y', daysOut: 5 } } },
        });
        expect(csv.split('\n')[1]).toBe('2026-07-21,"a,b","x""y",5');
    });

    it('returns just the header for empty history', () => {
        expect(historyToCsv(undefined)).toBe('Date,SyStudentID,LDA,Days Out');
        expect(historyToCsv({ version: 1, days: {} })).toBe('Date,SyStudentID,LDA,Days Out');
    });
});
