/*
 * scan-filter-derivation.js
 *
 * Pure helpers for deriving Chrome extension scan-filter settings from a
 * day's LDA sheet. No Excel.run / chrome / window usage here so these can
 * be unit tested directly (see __tests__/scan-filter-derivation.test.js).
 *
 * The LDA sheet layout these helpers understand is the one ldaProcessor.js
 * (react taskpane) writes in 'date' mode:
 *   - row 0: headers (includes a "Days Out" column)
 *   - rows 1..n: the main LDA table, sorted by Days Out descending
 *   - blank spacer rows
 *   - optional bold title cell "Failing Students (Active)" in column 0,
 *     followed by the failing table
 *   - optional "Low Attendance Students" section in the same shape
 */
import { findColumnIndex, normalizeHeader, sheetNameDateVariants } from '../../shared/excel-helpers.js';
import { CONSTANTS } from './constants.js';

// Section titles ldaProcessor writes into column 0 (normalized forms)
const FAILING_TITLE = normalizeHeader('Failing Students (Active)');
const ATTENDANCE_TITLE = normalizeHeader('Low Attendance Students');

/**
 * Base name of today's LDA sheet, matching ldaProcessor's 'date' naming:
 * "LDA M-D-YYYY" (no zero padding).
 * @param {Date} [date] - Defaults to today
 * @returns {string}
 */
export function todaysLdaSheetBaseName(date = new Date()) {
    return `LDA ${date.getMonth() + 1}-${date.getDate()}-${date.getFullYear()}`;
}

/**
 * Finds today's LDA sheet among the workbook's sheet names. Tolerates the
 * zero-padded / slash date variants via sheetNameDateVariants, plus the
 * " (2)", " (3)" suffixes ldaProcessor appends when the base name is taken
 * — the highest counter (most recently created) wins.
 *
 * @param {string[]} sheetNames - All worksheet names in the workbook
 * @param {Date} [date] - Defaults to today
 * @returns {string|null} The sheet name to inspect, or null if none found
 */
export function pickTodaysLdaSheet(sheetNames, date = new Date()) {
    const variants = sheetNameDateVariants(todaysLdaSheetBaseName(date));

    let best = null;
    let bestCounter = -1;
    for (const name of sheetNames) {
        for (const variant of variants) {
            if (name === variant) {
                // Base name (counter 0) loses to any "(n)" copy
                if (bestCounter < 0) { best = name; bestCounter = 0; }
            } else if (name.startsWith(`${variant} (`)) {
                const match = name.slice(variant.length).match(/^ \((\d+)\)$/);
                if (match) {
                    const counter = parseInt(match[1], 10);
                    if (counter > bestCounter) { best = name; bestCounter = counter; }
                }
            }
        }
    }
    return best;
}

/**
 * Derives scan-filter settings from an LDA sheet's used-range values.
 *
 * - daysOutFilter: ">=N" where N is the minimum numeric Days Out value in
 *   the MAIN table only (the failing/attendance sections hold students
 *   below the threshold and must not drag the minimum down). Null when the
 *   main table has no numeric Days Out values (e.g. headers only).
 * - includeFailing: whether a "Failing Students (Active)" section exists
 *   anywhere on the sheet. ldaProcessor only writes that title when the
 *   failing list was enabled AND produced rows.
 *
 * @param {Array<Array<*>>} values - 2D used-range values, headers in row 0
 * @returns {{ daysOutFilter: string|null, includeFailing: boolean }}
 */
export function deriveScanFilterFromGrid(values) {
    if (!Array.isArray(values) || values.length === 0) {
        return { daysOutFilter: null, includeFailing: false };
    }

    const normalizedHeaders = values[0].map(normalizeHeader);
    const daysOutIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.daysOut);

    const isBlankRow = (row) => row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    const sectionTitleOf = (row) => {
        const first = normalizeHeader(String(row[0] ?? ''));
        return first === FAILING_TITLE || first === ATTENDANCE_TITLE ? first : null;
    };

    // Walk the main table: contiguous rows after the header, ending at the
    // first blank row or section title.
    let minDaysOut = null;
    if (daysOutIdx !== -1) {
        for (let r = 1; r < values.length; r++) {
            const row = values[r];
            if (isBlankRow(row) || sectionTitleOf(row)) break;
            const value = row[daysOutIdx];
            if (typeof value === 'number' && !isNaN(value)) {
                minDaysOut = minDaysOut === null ? value : Math.min(minDaysOut, value);
            }
        }
    }

    // The failing section can sit anywhere below the main table.
    const includeFailing = values.some(row => sectionTitleOf(row) === FAILING_TITLE);

    return {
        // Floor fractional values: the extension's filter parser only
        // accepts integers, and >=floor(min) still includes the min row.
        daysOutFilter: minDaysOut !== null ? `>=${Math.floor(minDaysOut)}` : null,
        includeFailing
    };
}
