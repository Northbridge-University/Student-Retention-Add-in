/*
 * Timestamp: 2026-01-26 00:00:00
 * Version: 2.22.0
 * Author: Gemini (for Victor)
 * Description: Core logic for creating LDA reports.
 * Update: Optimize formatting performance by using range-based operations instead of cell-by-cell.
 *         Consecutive cells with the same color are merged into single range operations, reducing
 *         API calls from thousands to hundreds per batch. This dramatically speeds up formatting.
 */

import { getWorkbookSettings } from '../utility/getSettings';
import { saveDocumentSetting } from '../utility/documentSettings';
import { defaultColumns } from '../settings/DefaultSettings';
import { MASTER_LIST_SHEET, HISTORY_SHEET, BATCH_SIZE } from '../../../../shared/constants.js';
import { STUDENT_ID_ALIASES, STUDENT_NUMBER_ALIASES, LAST_LDA_ALIASES, SAP_STATUS_ALIASES } from '../../../../shared/columnAliases.js';
import { sanitizeRiskIndexSettings, buildDailySnapshot, mergeSnapshot, countRiskEpisodes, lookupRiskCount, loadLdaHistory, saveLdaHistory, todayKey } from './riskIndex.js';

const SHEET_NAMES = {
    MASTER_LIST: MASTER_LIST_SHEET,
    HISTORY: HISTORY_SHEET
};

// Row-count ceiling for formatting batches. Set modestly so typical data
// stays under the ~5MB per-sync payload cap. For high-density rows (lots of
// highlights/formulas), MAX_OPS_PER_SYNC triggers an earlier flush.
const FORMAT_BATCH_SIZE = 100;

// Hard cap on queued API ops between syncs. Excel Web rejects syncs whose
// payload exceeds its internal limit with "The request payload size has
// exceeded the limit." Keeping ops bounded avoids that, independent of batch
// size.
const MAX_OPS_PER_SYNC = 400;

// Fixed output-column widths. Office.js column widths are in points; pixels
// convert at ~0.75pt/px (96 DPI). Keyed by normalized column name (lowercased,
// whitespace removed) so both settings names and raw master headers match.
const FIXED_COLUMN_WIDTHS = {
    outreach: 217.5,            // 290px (also wrapped, see WRAP_COLUMN_NAMES)
    assigned: 57,              // 76px
    programversion: 267,       // 356px
    missingassignments: 51.75, // 69px
    grade: 45.75,              // 61px (Last Course Grade — settings name is "Grade")
    lastcoursegrade: 45.75,    // 61px (when shown under its raw master header)
    nextassignmentdue: 92.25,  // 123px
    ri: 30,                    // 40px (Risk Index)
};
const WRAP_COLUMN_NAMES = new Set(['outreach']);
const normalizeColName = (name) => String(name || '').toLowerCase().replace(/\s+/g, '');

/**
 * Helper to convert Excel serial date to MM-DD-YY string
 */
function formatExcelDate(serial) {
    if (!serial || isNaN(serial)) return serial;
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    const utcDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60000);
    const mm = String(utcDate.getMonth() + 1).padStart(2, '0');
    const dd = String(utcDate.getDate()).padStart(2, '0');
    const yy = String(utcDate.getFullYear()).slice(-2);
    return `${mm}-${dd}-${yy}`;
}

/**
 * Formats an Excel serial date (or string) into "Month Dayth, Year" format.
 * e.g. "March 2nd, 2026"
 */
function formatLongDate(value) {
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    let d = null;
    if (typeof value === 'number') {
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        d = new Date(date.valueOf() + date.getTimezoneOffset() * 60000);
    } else if (typeof value === 'string') {
        const match = value.trim().match(/^(\d{2})-(\d{2})-(\d{2})$/);
        if (match) {
            d = new Date(2000 + parseInt(match[3], 10), parseInt(match[1], 10) - 1, parseInt(match[2], 10));
        }
    }
    if (!d || isNaN(d.getTime())) return String(value);
    const day = d.getDate();
    const suffix = (day === 1 || day === 21 || day === 31) ? 'st'
        : (day === 2 || day === 22) ? 'nd'
        : (day === 3 || day === 23) ? 'rd' : 'th';
    return `${monthNames[d.getMonth()]} ${day}${suffix}, ${d.getFullYear()}`;
}

/**
 * Helper to format a date value into a friendly relative day string.
 * - "today" / "tomorrow" for 0/1 days out
 * - "this [Day]" for 2-6 days out
 * - "next [Day]" for 7-13 days out
 * - Falls back to "Month Dayth" (e.g. March 14th) for anything else
 */
function formatFriendlyDate(value) {
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

    // Check if the value is already "Today" or "Tomorrow" (case-insensitive)
    if (typeof value === 'string') {
        const lower = value.trim().toLowerCase();
        if (lower === 'today') return 'today';
        if (lower === 'tomorrow') return 'tomorrow';
    }

    // Try to parse into a Date object
    let targetDate = null;

    if (typeof value === 'number') {
        // Excel serial number
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        targetDate = new Date(date.valueOf() + date.getTimezoneOffset() * 60000);
    } else if (typeof value === 'string') {
        // Try MM-DD-YY format
        const match = value.trim().match(/^(\d{2})-(\d{2})-(\d{2})$/);
        if (match) {
            const year = 2000 + parseInt(match[3], 10);
            const month = parseInt(match[1], 10) - 1;
            const day = parseInt(match[2], 10);
            targetDate = new Date(year, month, day);
        }
    }

    if (!targetDate || isNaN(targetDate.getTime())) {
        // Can't parse, return as-is
        return typeof value === 'number' ? formatExcelDate(value) : value;
    }

    // Compare with today (strip time)
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    targetDate.setHours(0, 0, 0, 0);

    const diffDays = Math.round((targetDate - today) / (1000 * 60 * 60 * 24));

    if (diffDays === 0) return 'today';
    if (diffDays === 1) return 'tomorrow';
    if (diffDays >= 2 && diffDays <= 6) return `this ${dayNames[targetDate.getDay()]}`;
    if (diffDays >= 7 && diffDays <= 13) return `next ${dayNames[targetDate.getDay()]}`;

    // Beyond 13 days or in the past, use "Month Dayth" format
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'];
    const day = targetDate.getDate();
    const suffix = (day === 1 || day === 21 || day === 31) ? 'st'
        : (day === 2 || day === 22) ? 'nd'
        : (day === 3 || day === 23) ? 'rd' : 'th';
    return `${monthNames[targetDate.getMonth()]} ${day}${suffix}`;
}

// Format a note author for an outreach message. 'initials' collapses a name to
// its lowercase initials (e.g. "Victor Blanco" -> "vb"); anything else returns
// the trimmed full name. Returns the " - <author>" suffix, or '' when blank.
function formatAuthorSuffix(name, authorFormat) {
    const trimmed = String(name || '').trim();
    if (!trimmed) return '';
    if (authorFormat === 'initials') {
        const initials = trimmed.split(/\s+/).map(w => w[0]).join('').toLowerCase();
        if (initials) return ` - ${initials}`;
    }
    return ` - ${trimmed}`;
}

// Look up a tag map under any of the candidate IDs (master rows may carry
// both a SyStudentId and a Student Number, and history rows may have been
// keyed under either). Returns the first matching entry, or null.
function lookupTagEntry(map, sIds) {
    if (!map || !Array.isArray(sIds)) return null;
    for (const id of sIds) {
        if (id && map.has(id)) return map.get(id);
    }
    return null;
}

/**
 * Helper to generate the Retention Outreach Message.
 * Centralized logic for LDA tags, Missing Assignments, and Explicit DNC.
 * @param {string[]} sIds - Candidate Student IDs (SyStudentId and Student Number)
 * @param {Map} ldaMap - The map of LDA data
 * @param {number} missingVal - The value from "Missing Assignments" column
 * @param {string} tableContext - 'LDA_Table' or 'Failing_Table'
 * @param {Map} dncMap - Map of ID -> Tag Text
 * @param {Map} dncMessageMap - Map of ID -> { message, isQuote } for the qualifying DNC comment
 * @returns {string|null} - The formatted message or null
 */
function getRetentionMessage(sIds, ldaMap, missingVal, tableContext, dncMap, dncMessageMap, nextAssignmentDueVal, nextAssignmentDueColumnAllBlank, includeNextAssignmentDue = true, expandedDNC = false, authorFormat = 'initials') {
    // Priority 1: Explicit DNC (Highest Priority - Stop everything)
    const dncTag = lookupTagEntry(dncMap, sIds);
    if (dncTag) {
        // Split comma-separated tags and check each individually
        const individualTags = dncTag.split(',').map(t => t.trim());
        if (expandedDNC) {
            // Expanded DNC: only a general "DNC" or "DNC - Phone" tag populates the
            // Outreach column, using the comment text (or "Do not contact" for a quote)
            // plus the note author. Other variants write nothing here (their phone
            // columns are still struck through downstream).
            const triggersOutreach = individualTags.some(tag => tag === 'dnc' || tag === 'dnc - phone');
            if (triggersOutreach) {
                const dncMsg = lookupTagEntry(dncMessageMap, sIds);
                const author = formatAuthorSuffix(dncMsg && dncMsg.createdBy, authorFormat);
                if (dncMsg && dncMsg.message && !dncMsg.isQuote) {
                    return `${dncMsg.message}${author}`;
                }
                return `Do not contact${author}`;
            }
        } else {
            // Legacy DNC: any DNC except the phone-specific variants → "Do not contact".
            const hasExcludableDnc = individualTags.some(tag =>
                tag.includes('dnc') && tag !== 'dnc - phone' && tag !== 'dnc - other phone'
            );
            if (hasExcludableDnc) {
                return "Do not contact";
            }
        }
    }

    // Priority 2: Waiting for Grade (Strictly for Failing Table)
    //if (tableContext === 'Failing_Table' && typeof missingVal === 'number' && missingVal === 0) {
    //    return "[Retention] Waiting for Prof. to grade";
    //}

    // Priority 3: LDA Tag (Applies to any table)
    const ldaObj = lookupTagEntry(ldaMap, sIds);
    if (ldaObj && ldaObj.date) {
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'];
        const day = ldaObj.date.getDate();
        const suffix = (day === 1 || day === 21 || day === 31) ? 'st'
            : (day === 2 || day === 22) ? 'nd'
            : (day === 3 || day === 23) ? 'rd' : 'th';
        const friendlyDate = `${monthNames[ldaObj.date.getMonth()]} ${day}${suffix}`;
        const author = formatAuthorSuffix(ldaObj.createdBy, authorFormat);
        // When the engagement date is tomorrow, phrase it as "by tomorrow".
        const tomorrow = new Date();
        tomorrow.setHours(0, 0, 0, 0);
        tomorrow.setDate(tomorrow.getDate() + 1);
        if (ldaObj.date.getTime() === tomorrow.getTime()) {
            return `Student will engage by tomorrow${author}`;
        }
        return `Student will engage on ${friendlyDate}${author}`;
    }

    // Priority 4: Zero missing assignments with a next assignment due date
    // Skipped entirely when the user disables "Include Next Assignment Due".
    if (includeNextAssignmentDue && typeof missingVal === 'number' && missingVal === 0 && nextAssignmentDueVal) {
        const formattedDate = formatFriendlyDate(nextAssignmentDueVal);
        return `Student's next assignment is due ${formattedDate}.`;
    }

    // Priority 5: Zero missing assignments but NO next assignment due date (abnormality flag)
    // Only triggers if the column has data for other students (not entirely blank).
    // Also skipped when the user disables "Include Next Assignment Due" since the
    // flag is specifically tied to the next-assignment-due information.
    if (includeNextAssignmentDue && typeof missingVal === 'number' && missingVal === 0 && !nextAssignmentDueVal && !nextAssignmentDueColumnAllBlank) {
        return "Student has 0 missing assignments but they have no next assignment due date. Please check their Grade Book.";
    }

    return null;
}

/**
 * True when a SAP status value flags a student for the SAP List — i.e. it is
 * present and not "SAP Met" (case-insensitive, trimmed). Blank/empty is treated
 * as passing (not flagged), matching the AdSAPStatus convention.
 * @param {*} value - raw AdSAPStatus cell value
 * @returns {boolean}
 */
export function isSapFlagged(value) {
    const s = String(value ?? '').trim();
    if (s === '') return false;
    return s.toLowerCase() !== 'sap met';
}

/**
 * Main function to create the LDA report.
 * @param {Object} userOverrides - The settings from the UI (daysOut, includeFailingList, etc.)
 * @param {Function} onProgress - Callback to update UI steps: (stepId, status) => void
 * @param {Function} onBatchProgress - Optional callback for batch progress: (current, total, phase, tableName) => void
 * @param {Function} onCampusProgress - Optional callback for multi-campus progress: (campusName, campusIndex, totalCampuses, status) => void
 */
export async function createLDA(userOverrides, onProgress, onBatchProgress = null, onCampusProgress = null) {
    // Suppress StudentView's Outreach change handler for the entire LDA run.
    // Without this, when a StudentView is open at the same time, activating
    // the newly-created LDA sheet causes StudentView to rebind its onChanged
    // listener to it, and the programmatic Outreach writes below get
    // interpreted as user edits — causing addComment to fire for every row
    // and crashing the add-in. Cleared in finally so it's always reset.
    const hadWindow = (typeof window !== 'undefined');
    const previousSuppress = hadWindow ? window.__srkSuppressOutreachHandler : undefined;
    if (hadWindow) {
        window.__srkSuppressOutreachHandler = true;
    }
    try {
        // --- STEP 1: Validate Settings & Environment ---
        if (onProgress) onProgress('validate', 'active');
        
        const workbookSettings = getWorkbookSettings(defaultColumns);

        // If columns were auto-initialized from defaults, persist them so Settings tab stays in sync
        if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            const existing = Office.context.document.settings.get('workbookSettings');
            if (!existing || !Array.isArray(existing.columns) || existing.columns.length === 0) {
                saveDocumentSetting('workbookSettings', workbookSettings, 'ldaProcessor.autoInitColumns');
            }
        }

        const settings = {
            daysOut: userOverrides.daysOut ?? 5,
            includeFailingList: userOverrides.includeFailingList ?? false,
            includeSapList: userOverrides.includeSapList ?? false,
            includeAttendanceList: userOverrides.includeAttendanceList ?? false,
            includeLDATag: userOverrides.includeLDATag ?? true,
            includeDNCTag: userOverrides.includeDNCTag ?? true,
            includeNextAssignmentDue: userOverrides.includeNextAssignmentDue ?? true,
            // Advanced toggles: expanded DNC outreach messages (off = legacy
            // "Do not contact") and compact/fixed column widths (off = autosize only).
            expandedDNC: userOverrides.expandedDNC ?? false,
            compactColumnWidth: userOverrides.compactColumnWidth ?? true,
            // 'initials' = lowercase initials (e.g. "vb"), 'full' = full name.
            authorFormat: userOverrides.authorFormat ?? 'initials',
            // Highlight colors for the LDA-tag and DNC outreach comments.
            ldaTagColor: userOverrides.ldaTagColor ?? '#FFFF00',
            dncColor: userOverrides.dncColor ?? '#f2bdbd',
            sheetNameMode: userOverrides.sheetNameMode ?? 'date',
            columns: workbookSettings.columns,
            advisorAssignment: userOverrides.advisorAssignment ?? { enabled: false, advisors: [] },
            riskIndex: sanitizeRiskIndexSettings(userOverrides.riskIndex)
        };

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) {
                throw new Error(`Workbook missing "${SHEET_NAMES.MASTER_LIST}" sheet.`);
            }
            if (onProgress) onProgress('validate', 'completed');

            // --- STEP 2: Read Master List & Scan Colors ---
            if (onProgress) onProgress('read', 'active');

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();

            // First, get dimensions only
            masterRange.load("rowCount, columnCount");
            await context.sync();

            const totalRows = masterRange.rowCount;
            const totalCols = masterRange.columnCount;

            // Read values and formulas in batches to avoid response payload limits
            let masterValues = [];
            let masterFormulas = [];

            for (let startRow = 0; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = masterSheet.getRangeByIndexes(startRow, 0, rowsToRead, totalCols);
                batchRange.load("values, formulas");
                await context.sync();

                masterValues = masterValues.concat(batchRange.values);
                masterFormulas = masterFormulas.concat(batchRange.formulas);

                // Report progress for large datasets
                if (onBatchProgress && totalRows > BATCH_SIZE) {
                    const currentBatch = Math.floor(startRow / BATCH_SIZE) + 1;
                    const totalBatches = Math.ceil(totalRows / BATCH_SIZE);
                    onBatchProgress(currentBatch, totalBatches, 'reading', 'Master_List');
                }
            }

            const headers = masterValues[0];

            // Read cell colors from a sample to build color map
            // getCellProperties returns ~100-200 bytes per cell; cap total cells to stay under ~5MB payload limit
            const MAX_COLOR_CELLS = 12000;
            const colorSampleRows = Math.min(
                Math.max(50, Math.floor(MAX_COLOR_CELLS / totalCols)),
                500,
                totalRows
            );
            const colorSampleRange = masterSheet.getRangeByIndexes(0, 0, colorSampleRows, totalCols);
            let colorSampleProps;
            try {
                colorSampleProps = colorSampleRange.getCellProperties({
                    format: { fill: { color: true } }
                });
            } catch (e) {
                console.warn('getCellProperties failed, skipping color mapping', e);
                colorSampleProps = null;
            }

            let historyData = null;
            const hasHistory = sheets.items.some(s => s.name === SHEET_NAMES.HISTORY);
            if (hasHistory) {
                const histSheet = sheets.getItem(SHEET_NAMES.HISTORY);
                const histRange = histSheet.getUsedRange();
                histRange.load("values");
                historyData = histRange;
            }

            await context.sync();

            // getCellProperties may have failed (null) or may not have resolved
            let masterColors = null;
            try {
                if (colorSampleProps && colorSampleProps.value) {
                    masterColors = colorSampleProps.value;
                }
            } catch (e) {
                console.warn('Color sample resolution failed, skipping color mapping', e);
                masterColors = null;
            }
            
            // --- UPDATED: Space-Insensitive Matching ---
            const cleanStr = (s) => String(s || '').trim().toLowerCase();
            const stripStr = (s) => cleanStr(s).replace(/\s+/g, '');

            const getColIndex = (settingName) => {
                const targetNameStripped = stripStr(settingName);

                // 1. Find Config
                const colConfig = settings.columns.find(c => 
                    stripStr(c.name) === targetNameStripped
                );
                // If checking for a raw header that isn't in settings config but passed as a name
                let aliases = [];
                if (colConfig) {
                    if (Array.isArray(colConfig.alias)) {
                        aliases = colConfig.alias;
                    } else if (colConfig.alias) {
                        aliases = [colConfig.alias];
                    }
                }
                
                const candidates = [settingName, ...aliases]; // settingName might be a raw header now

                // 2. Find match in Excel Headers
                for (const rawCand of candidates) {
                    const candStripped = stripStr(rawCand);
                    const idx = headers.findIndex(h => stripStr(h) === candStripped);
                    if (idx !== -1) return idx;
                }

                return -1;
            };

            // --- Implicit Column Detection ---
            // 1. Filter settings columns to only those found in the Master List (skip missing)
            const matchedColumns = settings.columns.filter(col => getColIndex(col.name) !== -1);

            // 2. Mark which Master List indices are covered by matched settings columns
            const usedMasterIndices = new Set();
            matchedColumns.forEach(col => {
                const idx = getColIndex(col.name);
                if (idx !== -1) usedMasterIndices.add(idx);
            });

            // 3. Find unused Master List headers and append them as hidden
            const outputColumns = [...matchedColumns]; // Start with matched user config

            headers.forEach((h, idx) => {
                if (!usedMasterIndices.has(idx) && h && String(h).trim() !== "") {
                    // This is a column in Master List not in settings.
                    // Add it, but hide it.
                    outputColumns.push({
                        name: h,
                        hidden: true
                    });
                }
            });

            // --- Retrieve Key Indices ---
            const daysOutIdx = getColIndex('Days Out');
            const gradeIdx = getColIndex('Grade');
            // Resolve SyStudentId and Student Number as separate column
            // indices on the Master List. We need both because Student
            // History rows written before the SyStudentId migration are
            // keyed under Student Number, while newer rows are keyed under
            // SyStudentId — so a master row with both columns must be
            // checked against the tag maps under either value.
            const findColIdxByAliases = (aliases) => {
                for (const candidate of aliases) {
                    const idx = getColIndex(candidate);
                    if (idx !== -1) return idx;
                }
                return -1;
            };
            const sySidIdx = findColIdxByAliases(STUDENT_ID_ALIASES);
            const studentNumberIdx = findColIdxByAliases(STUDENT_NUMBER_ALIASES);
            // Prefer SyStudentId; fall back to Student Number for legacy
            // workbooks that don't have a SyStudentId column.
            const studentIdIdx = sySidIdx !== -1 ? sySidIdx : studentNumberIdx;

            // --- ProgramVersion index for advisor assignment ---
            const pvAliases = ['programversion', 'program', 'progversdescrip'];
            const programVersionIdx = headers.findIndex(h => pvAliases.includes(stripStr(h)));
            
            // Look for "Missing Assignments" column
            let missingIdx = getColIndex('Missing Assignments');
            if (missingIdx === -1) {
                 // Fallback scan
                 missingIdx = headers.findIndex(h => String(h).trim().toLowerCase().includes('missing'));
            }

            // Look for "Next Assignment Due" column
            let nextAssignmentDueIdx = getColIndex('Next Assignment Due');
            if (nextAssignmentDueIdx === -1) {
                nextAssignmentDueIdx = headers.findIndex(h =>
                    String(h).trim().toLowerCase().replace(/\s+/g, '') === 'nextassignmentdue'
                );
            }

            // Check if the entire "Next Assignment Due" column is blank (skip abnormality check if so)
            let nextAssignmentDueColumnAllBlank = true;
            if (nextAssignmentDueIdx !== -1) {
                for (let i = 1; i < masterValues.length; i++) {
                    const val = masterValues[i][nextAssignmentDueIdx];
                    if (val !== null && val !== undefined && val !== '') {
                        nextAssignmentDueColumnAllBlank = false;
                        break;
                    }
                }
            }

            // Look for "Attendance %" column
            let attendanceIdx = getColIndex('Attendance %');
            if (attendanceIdx === -1) {
                attendanceIdx = headers.findIndex(h => {
                    const stripped = String(h || '').trim().toLowerCase().replace(/\s+/g, '');
                    return stripped === 'attendance%' || stripped === 'attendancepercent' || stripped === 'attendance';
                });
            }

            // Look for the SAP status column (AdSAPStatus) for the SAP List table.
            let sapIdx = getColIndex('AdSAPStatus');
            if (sapIdx === -1) {
                sapIdx = headers.findIndex(h => SAP_STATUS_ALIASES.some(a => stripStr(h) === stripStr(a)));
            }

            if (daysOutIdx === -1) throw new Error("Could not find 'Days Out' column in Master List. Check Settings.");

            // --- Risk Index: daily LDA/Days Out snapshot + episode counts ---
            // One snapshot per calendar day for the whole Master List (a same-day
            // re-run overwrites that day's entry), stored under the LDAHistory
            // document setting. The counts feed the RI output column and the
            // student view badge.
            let riskIndexMap = new Map();
            const riskEnabled = settings.riskIndex.enabled && studentIdIdx !== -1;
            if (riskEnabled) {
                const ldaColIdx = findColIdxByAliases(LAST_LDA_ALIASES);
                const snapshot = buildDailySnapshot(masterValues, { idIdx: studentIdIdx, ldaIdx: ldaColIdx, daysOutIdx });
                const ldaHistory = mergeSnapshot(loadLdaHistory(), todayKey(), snapshot);
                saveLdaHistory(ldaHistory);
                riskIndexMap = countRiskEpisodes(ldaHistory, settings.riskIndex.threshold);
            }
            if (riskEnabled && settings.riskIndex.showColumn) {
                // Place RI right after Days Out; the column is synthetic (no
                // Master List source), flagged so buildOutputRow fills it.
                const riColumn = { name: 'RI', riskIndex: true };
                const daysOutOutIdx = outputColumns.findIndex(c => getColIndex(c.name) === daysOutIdx);
                if (daysOutOutIdx !== -1) {
                    outputColumns.splice(daysOutOutIdx + 1, 0, riColumn);
                } else {
                    outputColumns.push(riColumn);
                }
            }

            // Look for "Course Start" column
            let courseStartIdx = getColIndex('Course Start');
            if (courseStartIdx === -1) {
                courseStartIdx = headers.findIndex(h => {
                    const stripped = String(h || '').trim().toLowerCase().replace(/\s+/g, '');
                    return stripped === 'coursestart' || stripped === 'coursestartdate' || stripped === 'coursebegindate';
                });
            }

            // Compute Course Start mode (most frequent value = baseline)
            let courseStartModeValue = null;
            if (courseStartIdx !== -1) {
                const freqMap = new Map();
                for (let i = 1; i < masterValues.length; i++) {
                    const val = masterValues[i][courseStartIdx];
                    if (val !== null && val !== undefined && val !== '') {
                        const key = String(val);
                        freqMap.set(key, (freqMap.get(key) || 0) + 1);
                    }
                }
                let maxCount = 0;
                let modeKey = null;
                for (const [key, count] of freqMap) {
                    if (count > maxCount) {
                        maxCount = count;
                        modeKey = key;
                    }
                }
                if (modeKey !== null) {
                    for (let i = 1; i < masterValues.length; i++) {
                        if (String(masterValues[i][courseStartIdx]) === modeKey) {
                            courseStartModeValue = masterValues[i][courseStartIdx];
                            break;
                        }
                    }
                    console.log(`LDA: Course Start baseline (mode): ${typeof courseStartModeValue === 'number' ? formatExcelDate(courseStartModeValue) : courseStartModeValue} (${maxCount} occurrences)`);
                }
            }

            if (onProgress) onProgress('read', 'completed');


            // --- STEP 2b: Build Color Map (Value -> Color) ---
            const columnColorMaps = new Map();

            // Re-sort outputColumns to ensure processing order (though less critical now)
            // We only really care about mapping indices for color scanning
            
            const outputColIndices = outputColumns
                .map(c => getColIndex(c.name))
                .filter(idx => idx !== -1);

            // --- EXCLUDED COLORS SET ---
            const EXCLUDED_COLORS = new Set([
                '#ffffff', // White
                '#add8e6', // LightBlue (Standard Excel/HTML name)
                '#fc0019', // Bright Red
                '#ff0000', // Red
                '#ff0d0d', // Red variant
                '#ff1a1a', // Red variant
                '#fe0000', // Red variant
                '#ff2400', // Scarlet
                '#cc0000', // Dark Red
                '#ee0000', // Red variant
                '#dd0000', // Red variant
                '#e60000'  // Red variant
            ]);

            // Only iterate over the sampled color rows (masterColors has limited rows)
            if (masterColors) {
                const colorRowLimit = Math.min(masterColors.length, masterValues.length);
                for (let r = 1; r < colorRowLimit; r++) {
                    outputColIndices.forEach(cIdx => {
                        const val = masterValues[r][cIdx];
                        const rawColor = masterColors[r]?.[cIdx]?.format?.fill?.color;
                        const normColor = String(rawColor || '').toLowerCase();

                        // Check if value exists, color exists, and color is NOT in the excluded list
                        if (val && rawColor && !EXCLUDED_COLORS.has(normColor)) {
                            if (!columnColorMaps.has(cIdx)) {
                                columnColorMaps.set(cIdx, new Map());
                            }
                            columnColorMaps.get(cIdx).set(String(val), rawColor);
                        }
                    });
                }
            }

            // --- STEP 3: Filtering (days out + failing + attendance) ---
            if (onProgress) onProgress('filter', 'active');

            // When daysOut < 5 and includeFailingList is on, students who are failing
            // AND have days out < 5 should go to the failing table instead of the
            // LDA table to avoid duplicates on both tables.
            const shouldDeferToFailing = settings.daysOut < 5
                && settings.includeFailingList
                && gradeIdx !== -1;

            const dataRows = [];
            for (let i = 1; i < masterValues.length; i++) {
                const daysOutVal = masterValues[i][daysOutIdx];
                if (typeof daysOutVal === 'number' && daysOutVal >= settings.daysOut) {
                    // If failing list has priority, check whether this student is failing
                    // and in the overlap zone (days out < 5) — if so, skip them here
                    // so they land on the failing table instead.
                    if (shouldDeferToFailing && daysOutVal < 5) {
                        const gradeVal = masterValues[i][gradeIdx];
                        const isFailing = (typeof gradeVal === 'number')
                            && (gradeVal < 0.60 || (gradeVal >= 1 && gradeVal < 60));
                        if (isFailing) continue;
                    }

                    dataRows.push({
                        values: masterValues[i],
                        formulas: masterFormulas[i],
                        originalIndex: i
                    });
                }
            }

            dataRows.sort((a, b) => (b.values[daysOutIdx] || 0) - (a.values[daysOutIdx] || 0));

            // --- STEP 4: Filtering by Grades (Failing) ---
            let failingRows = [];
            if (settings.includeFailingList && gradeIdx !== -1) {
                for (let i = 1; i < masterValues.length; i++) {
                    const gradeVal = masterValues[i][gradeIdx];
                    const daysOutVal = masterValues[i][daysOutIdx];

                    const isFailing = (typeof gradeVal === 'number') && (gradeVal < 0.60 || (gradeVal >= 1 && gradeVal < 60));
                    const isRecent = (typeof daysOutVal === 'number') && (daysOutVal <= 4);

                    if (isFailing && isRecent) {
                        failingRows.push({
                            values: masterValues[i],
                            formulas: masterFormulas[i],
                            originalIndex: i
                        });
                    }
                }
                failingRows.sort((a, b) => (a.values[gradeIdx] || 0) - (b.values[gradeIdx] || 0));
            }

            // --- STEP 4b: Filtering by Attendance ---
            let attendanceRows = [];
            if (settings.includeAttendanceList && attendanceIdx !== -1) {
                // Build set of LDA student IDs (students on the days-out table)
                const ldaStudentIds = new Set();
                if (studentIdIdx !== -1) {
                    for (const row of dataRows) {
                        const sId = row.values[studentIdIdx];
                        if (sId) ldaStudentIds.add(sId);
                    }
                }

                // Track which student IDs end up in the attendance list
                const attendanceStudentIds = new Set();

                // Scan ALL master list rows for low attendance (not just failing)
                for (let i = 1; i < masterValues.length; i++) {
                    const sId = studentIdIdx !== -1 ? masterValues[i][studentIdIdx] : null;

                    // Skip students already on the LDA (days out) table
                    if (sId && ldaStudentIds.has(sId)) continue;

                    const attVal = masterValues[i][attendanceIdx];
                    // Parse attendance value - handle both decimal (0.59) and whole number (59) formats
                    let attPercent = null;
                    if (typeof attVal === 'number') {
                        attPercent = attVal <= 1 ? attVal * 100 : attVal;
                    }

                    if (attPercent !== null && attPercent < 60) {
                        attendanceRows.push({
                            values: masterValues[i],
                            formulas: masterFormulas[i],
                            originalIndex: i,
                            attendancePercent: attPercent
                        });
                        if (sId) attendanceStudentIds.add(sId);
                    }
                }

                // Sort attendance rows from lowest to highest attendance
                attendanceRows.sort((a, b) => (a.attendancePercent || 0) - (b.attendancePercent || 0));

                // Remove students with low attendance from the failing list (no duplicates)
                if (attendanceStudentIds.size > 0) {
                    failingRows = failingRows.filter(row => {
                        const sId = studentIdIdx !== -1 ? row.values[studentIdIdx] : null;
                        return !sId || !attendanceStudentIds.has(sId);
                    });
                }
            }

            // --- STEP 4c: Filtering by SAP status ---
            // Include students whose AdSAPStatus is not "SAP Met" and not blank.
            // Priority is LDA > SAP > Failing: exclude only students already on the
            // LDA (days-out) table here, then drop any SAP-flagged student from the
            // Failing list below so SAP wins over Failing. Sorted by SAP status value
            // so like statuses group together.
            let sapRows = [];
            if (settings.includeSapList && sapIdx !== -1) {
                const ldaIds = new Set();
                const ldaIndices = new Set();
                for (const row of dataRows) {
                    const sId = studentIdIdx !== -1 ? row.values[studentIdIdx] : null;
                    if (sId) ldaIds.add(sId);
                    ldaIndices.add(row.originalIndex);
                }
                const sapIds = new Set();
                const sapIndices = new Set();
                for (let i = 1; i < masterValues.length; i++) {
                    if (!isSapFlagged(masterValues[i][sapIdx])) continue;
                    const sId = studentIdIdx !== -1 ? masterValues[i][studentIdIdx] : null;
                    if (sId && ldaIds.has(sId)) continue;
                    if (ldaIndices.has(i)) continue;
                    sapRows.push({
                        values: masterValues[i],
                        formulas: masterFormulas[i],
                        originalIndex: i
                    });
                    if (sId) sapIds.add(sId);
                    sapIndices.add(i);
                }
                sapRows.sort((a, b) =>
                    String(a.values[sapIdx] || '').toLowerCase().localeCompare(String(b.values[sapIdx] || '').toLowerCase())
                );

                // SAP wins over Failing: remove SAP-flagged students from the Failing
                // list so a student never appears on both tables.
                if (sapIndices.size > 0) {
                    failingRows = failingRows.filter(row => {
                        const sId = studentIdIdx !== -1 ? row.values[studentIdIdx] : null;
                        if (sId && sapIds.has(sId)) return false;
                        return !sapIndices.has(row.originalIndex);
                    });
                }
            }

            if (onProgress) onProgress('filter', 'completed');

            // --- Advisor Auto-Assignment Map ---
            // When enabled, pre-compute which advisor each student row maps to
            const advisorAssignmentMap = new Map(); // originalIndex -> { name, color }
            const advisorConfig = settings.advisorAssignment;
            if (advisorConfig && advisorConfig.enabled && advisorConfig.advisors && advisorConfig.advisors.length > 0) {
                // Filter out advisors excluded today (by day of week)
                const todayDayKey = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'][new Date().getDay()];
                const activeAdvisors = advisorConfig.advisors.filter(a => !(a.excludeDays || []).includes(todayDayKey));
                if (activeAdvisors.length === 0) {
                    console.log('LDA: All advisors are excluded today, skipping auto-assignment');
                }

                // Build student list with metadata for multi-filter assignment
                const taggedStudents = [];
                for (const r of dataRows) {
                    taggedStudents.push({
                        programVersion: programVersionIdx !== -1 ? String(r.values[programVersionIdx] || '').trim() : '',
                        daysOut: typeof r.values[daysOutIdx] === 'number' ? r.values[daysOutIdx] : 0,
                        listType: 'lda',
                        originalIndex: r.originalIndex
                    });
                }
                for (const r of failingRows) {
                    taggedStudents.push({
                        programVersion: programVersionIdx !== -1 ? String(r.values[programVersionIdx] || '').trim() : '',
                        daysOut: typeof r.values[daysOutIdx] === 'number' ? r.values[daysOutIdx] : 0,
                        listType: 'failing',
                        originalIndex: r.originalIndex
                    });
                }
                for (const r of attendanceRows) {
                    taggedStudents.push({
                        programVersion: programVersionIdx !== -1 ? String(r.values[programVersionIdx] || '').trim() : '',
                        daysOut: typeof r.values[daysOutIdx] === 'number' ? r.values[daysOutIdx] : 0,
                        listType: 'attendance',
                        originalIndex: r.originalIndex
                    });
                }
                for (const r of sapRows) {
                    taggedStudents.push({
                        programVersion: programVersionIdx !== -1 ? String(r.values[programVersionIdx] || '').trim() : '',
                        daysOut: typeof r.values[daysOutIdx] === 'number' ? r.values[daysOutIdx] : 0,
                        listType: 'sap',
                        originalIndex: r.originalIndex
                    });
                }

                if (activeAdvisors.length > 0) {
                    // Use the shared algorithm with only active (non-excluded) advisors
                    const results = assignStudentsToAdvisors(taggedStudents, activeAdvisors);

                    // Re-run inline to build the per-student map
                    const hasAnyFilter = (a) => (a.programVersions?.length > 0) || (a.listPreference?.length > 0) || (a.daysOutMin != null) || (a.daysOutMax != null);
                    const advisorMatches = (a, s) => {
                        if (a.programVersions?.length > 0) {
                            const pvSet = new Set(a.programVersions.map(pv => pv.toLowerCase()));
                            if (!pvSet.has((s.programVersion || '').toLowerCase())) return false;
                        }
                        if (a.listPreference?.length > 0) {
                            if (!a.listPreference.includes(s.listType)) return false;
                        }
                        if (a.daysOutMin != null && typeof s.daysOut === 'number' && s.daysOut < a.daysOutMin) return false;
                        if (a.daysOutMax != null && typeof s.daysOut === 'number' && s.daysOut > a.daysOutMax) return false;
                        return true;
                    };
                    const getFilterSig = (a) => {
                        const parts = [];
                        if (a.programVersions?.length) parts.push('pv:' + [...a.programVersions].sort().join(',').toLowerCase());
                        if (a.listPreference?.length) parts.push('list:' + [...a.listPreference].sort().join(','));
                        if (a.daysOutMin != null) parts.push('doMin:' + a.daysOutMin);
                        if (a.daysOutMax != null) parts.push('doMax:' + a.daysOutMax);
                        return parts.join('|');
                    };

                    const assignedArr = new Array(taggedStudents.length).fill(null);
                    const processedSigs = new Set();
                    const sigGroups = new Map();
                    for (const a of activeAdvisors) {
                        if (!hasAnyFilter(a)) continue;
                        const sig = getFilterSig(a);
                        if (!sigGroups.has(sig)) sigGroups.set(sig, []);
                        sigGroups.get(sig).push(a);
                    }
                    for (const a of activeAdvisors) {
                        if (!hasAnyFilter(a)) continue;
                        const sig = getFilterSig(a);
                        if (processedSigs.has(sig)) continue;
                        processedSigs.add(sig);
                        const group = sigGroups.get(sig);
                        let rr = 0;
                        for (let i = 0; i < taggedStudents.length; i++) {
                            if (assignedArr[i]) continue;
                            if (advisorMatches(group[0], taggedStudents[i])) {
                                assignedArr[i] = group[rr % group.length];
                                rr++;
                            }
                        }
                    }
                    let rrIdx = 0;
                    for (let i = 0; i < taggedStudents.length; i++) {
                        if (!assignedArr[i]) {
                            let placed = false;
                            for (let attempt = 0; attempt < activeAdvisors.length; attempt++) {
                                const adv = activeAdvisors[(rrIdx + attempt) % activeAdvisors.length];
                                if (!hasAnyFilter(adv) || advisorMatches(adv, taggedStudents[i])) {
                                    assignedArr[i] = adv;
                                    rrIdx = (rrIdx + attempt + 1);
                                    placed = true;
                                    break;
                                }
                            }
                            if (!placed) {
                                assignedArr[i] = activeAdvisors[rrIdx % activeAdvisors.length];
                                rrIdx++;
                            }
                        }
                    }

                    // Even Split redistribution for inline assignment
                    if (advisorConfig.evenSplit && activeAdvisors.length > 1) {
                        const total = taggedStudents.length;
                        const base = Math.floor(total / activeAdvisors.length);
                        const extra = total % activeAdvisors.length;
                        const targets = new Map();
                        activeAdvisors.forEach((a, idx) => targets.set(a.id, base + (idx < extra ? 1 : 0)));

                        const inlineCounts = new Map();
                        activeAdvisors.forEach(a => inlineCounts.set(a.id, 0));
                        for (let i = 0; i < taggedStudents.length; i++) {
                            inlineCounts.set(assignedArr[i].id, inlineCounts.get(assignedArr[i].id) + 1);
                        }

                        const perAdvisor = new Map();
                        activeAdvisors.forEach(a => perAdvisor.set(a.id, []));
                        for (let i = 0; i < taggedStudents.length; i++) {
                            perAdvisor.get(assignedArr[i].id).push(i);
                        }

                        // Helper: move student between advisors
                        const inlineMoveStudent = (stuIdx, fromId, toId) => {
                            const advObj = activeAdvisors.find(a => a.id === toId);
                            assignedArr[stuIdx] = advObj;
                            inlineCounts.set(fromId, inlineCounts.get(fromId) - 1);
                            inlineCounts.set(toId, inlineCounts.get(toId) + 1);
                            const fromList = perAdvisor.get(fromId);
                            fromList.splice(fromList.indexOf(stuIdx), 1);
                            perAdvisor.get(toId).push(stuIdx);
                        };

                        // Pass 1: Move to under-allocated advisors that MATCH filters
                        for (const overAdv of activeAdvisors.filter(a => inlineCounts.get(a.id) > targets.get(a.id))) {
                            const surplus = inlineCounts.get(overAdv.id) - targets.get(overAdv.id);
                            if (surplus <= 0) continue;
                            const indices = [...perAdvisor.get(overAdv.id)];
                            indices.sort((a, b) => (taggedStudents[a].daysOut || 0) - (taggedStudents[b].daysOut || 0));
                            let moved = 0;
                            for (let si = 0; si < indices.length && moved < surplus; si++) {
                                const stuIdx = indices[si];
                                const eligible = activeAdvisors.filter(ua =>
                                    ua.id !== overAdv.id &&
                                    inlineCounts.get(ua.id) < targets.get(ua.id) &&
                                    (!hasAnyFilter(ua) || advisorMatches(ua, taggedStudents[stuIdx]))
                                );
                                if (eligible.length === 0) continue;
                                eligible.sort((a, b) => {
                                    const aMax = a.daysOutMax != null ? a.daysOutMax : Infinity;
                                    const bMax = b.daysOutMax != null ? b.daysOutMax : Infinity;
                                    if (aMax !== Infinity && bMax !== Infinity) return aMax - bMax;
                                    if (aMax !== Infinity) return -1;
                                    if (bMax !== Infinity) return 1;
                                    return 0;
                                });
                                inlineMoveStudent(stuIdx, overAdv.id, eligible[0].id);
                                moved++;
                            }
                        }

                        // Pass 2: Force-balance remaining surplus ignoring filters
                        for (const overAdv of activeAdvisors.filter(a => inlineCounts.get(a.id) > targets.get(a.id))) {
                            let surplus = inlineCounts.get(overAdv.id) - targets.get(overAdv.id);
                            if (surplus <= 0) continue;
                            const indices = [...perAdvisor.get(overAdv.id)];
                            indices.sort((a, b) => (taggedStudents[b].daysOut || 0) - (taggedStudents[a].daysOut || 0));
                            for (let si = 0; si < indices.length && surplus > 0; si++) {
                                const stuIdx = indices[si];
                                let bestTarget = null;
                                let bestDeficit = 0;
                                for (const ua of activeAdvisors) {
                                    if (ua.id === overAdv.id) continue;
                                    const deficit = targets.get(ua.id) - inlineCounts.get(ua.id);
                                    if (deficit > bestDeficit) {
                                        bestDeficit = deficit;
                                        bestTarget = ua;
                                    }
                                }
                                if (!bestTarget) break;
                                inlineMoveStudent(stuIdx, overAdv.id, bestTarget.id);
                                surplus--;
                            }
                        }
                    }

                    for (let i = 0; i < taggedStudents.length; i++) {
                        advisorAssignmentMap.set(taggedStudents[i].originalIndex, {
                            name: assignedArr[i].name,
                            color: assignedArr[i].color
                        });
                    }
                    console.log(`LDA: Auto-assigned ${taggedStudents.length} students across ${activeAdvisors.length} active advisors (${advisorConfig.advisors.length - activeAdvisors.length} excluded today)`);
                }
            }

            // Pin specific output columns to fixed widths (and wrap where configured)
            // after autofit, so messages wrap and key columns keep a consistent size.
            const applyFixedColumnWidths = (sheet) => {
                outputColumns.forEach((colConfig, idx) => {
                    const key = normalizeColName(colConfig.name);
                    const width = FIXED_COLUMN_WIDTHS[key];
                    if (width == null) return;
                    const col = sheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn();
                    col.format.columnWidth = width;
                    if (WRAP_COLUMN_NAMES.has(key)) col.format.wrapText = true;
                });
            };

            // --- Detect Multi-Campus Mode ---
            let isMultiCampus = false;
            let campusList = [];
            const campusIdx = getColIndex('Campus');

            if (settings.sheetNameMode === 'campus' && campusIdx !== -1) {
                const campusSet = new Set();
                for (let i = 1; i < masterValues.length; i++) {
                    const val = String(masterValues[i][campusIdx] || '').trim();
                    if (val) campusSet.add(val);
                }
                campusList = Array.from(campusSet).sort();
                isMultiCampus = campusList.length > 1;
            }


            if (isMultiCampus) {
                // ============================================================
                // MULTI-CAMPUS FLOW: Create one LDA sheet per campus
                // ============================================================

                // --- STEP 5 (Tags - global): Applying LDA & DNC Tags ---
                if (onProgress) onProgress('tags', 'active');

                const dncMap = new Map();
                const dncMessageMap = new Map();
                const ldaFollowUpMap = new Map();

                if (historyData && studentIdIdx !== -1) {
                    const hValues = historyData.values;
                    if (hValues.length > 0) {
                        const hHeaders = hValues[0].map(h => String(h).toLowerCase().trim());
                        // Collect every ID-like column (SyStudentId, Student Number, Student ID, etc.).
                        // Student History rows written by newer code land in SyStudentId while older
                        // rows landed in Student Number; the Master List only exposes one of these,
                        // so we index the map under every non-empty variant from each history row.
                        const hIdIndices = [];
                        hHeaders.forEach((h, idx) => {
                            if ((h.includes('student') && h.includes('id')) || h.includes('number')) {
                                hIdIndices.push(idx);
                            }
                        });
                        const hTagIdx = hHeaders.indexOf('tag');
                        const hCommentIdx = hHeaders.indexOf('comment');
                        const hCreatedByIdx = hHeaders.indexOf('created by');

                        if (hIdIndices.length > 0 && hTagIdx !== -1) {
                            const todayTime = new Date().setHours(0,0,0,0);
                            const ldaRegex = /\blda\b.*?(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;

                            for (let i = hValues.length - 1; i > 0; i--) {
                                // Normalize to strings so Map lookups match regardless of whether
                                // Excel returned the ID as a Number or a String.
                                const hids = hIdIndices
                                    .map(idx => hValues[i][idx])
                                    .filter(v => v !== null && v !== undefined && v !== '')
                                    .map(v => String(v).trim())
                                    .filter(v => v.length > 0);
                                const htagRaw = String(hValues[i][hTagIdx] || '');
                                const htagLower = htagRaw.toLowerCase().trim();

                                if (settings.includeDNCTag && hids.length > 0 && htagLower.includes('dnc')) {
                                    // Accumulate DNC tags so multiple entries (e.g. "DNC - Phone" and "DNC - Other Phone") are preserved
                                    for (const hid of hids) {
                                        dncMap.set(hid, dncMap.has(hid) ? dncMap.get(hid) + ', ' + htagLower : htagLower);
                                    }

                                    // Capture the comment text for the most recent qualifying DNC entry
                                    // (a bare "DNC" or "DNC - Phone" tag). That text becomes the Outreach
                                    // message — unless the comment is a Quote, in which case we only flag it.
                                    // Rows are walked newest-first, so the first match per student wins.
                                    const rowTags = htagLower.split(',').map(t => t.trim());
                                    const isQualifyingDnc = rowTags.includes('dnc') || rowTags.includes('dnc - phone');
                                    if (isQualifyingDnc) {
                                        const isQuote = rowTags.includes('quote');
                                        const message = hCommentIdx !== -1 ? String(hValues[i][hCommentIdx] || '').trim() : '';
                                        const createdBy = hCreatedByIdx !== -1 ? String(hValues[i][hCreatedByIdx] || '').trim() : '';
                                        for (const hid of hids) {
                                            if (!dncMessageMap.has(hid)) {
                                                dncMessageMap.set(hid, { message, isQuote, createdBy });
                                            }
                                        }
                                    }
                                }

                                if (settings.includeLDATag && hids.length > 0) {
                                    const match = htagRaw.match(ldaRegex);
                                    if (match) {
                                        const ldaDate = new Date(match[1]);
                                        if (!isNaN(ldaDate.getTime())) {
                                            ldaDate.setHours(0,0,0,0);
                                            if (ldaDate >= todayTime) {
                                                for (const hid of hids) {
                                                    if (!ldaFollowUpMap.has(hid)) {
                                                        ldaFollowUpMap.set(hid, { date: ldaDate, text: htagRaw, createdBy: hCreatedByIdx !== -1 ? String(hValues[i][hCreatedByIdx] || '').trim() : '' });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (onProgress) onProgress('tags', 'completed');

                // --- Shared format helpers (used by all campus sheets) ---
                const dateColumnIndices = new Set();
                outputColumns.forEach((colConfig) => {
                    const masterIdx = getColIndex(colConfig.name);
                    if (masterIdx === -1) return;
                    const excelHeader = String(masterValues[0][masterIdx] || '').toLowerCase();
                    const isExcluded = /id|no\.|num|code|zip|postal|social|ssn|phone|grade|score|credit|fee|days|count/i.test(excelHeader);
                    if (isExcluded) return;
                    let dateCount = 0;
                    let numCount = 0;
                    const limit = Math.min(masterValues.length, 100);
                    for (let i = 1; i < limit; i++) {
                        const val = masterValues[i][masterIdx];
                        if (typeof val === 'number') {
                            numCount++;
                            if (val > 10958 && val < 73051) dateCount++;
                        }
                    }
                    if (numCount > 0 && (dateCount / numCount) > 0.5) {
                        dateColumnIndices.add(colConfig.name);
                    }
                });

                const outreachColIndex = outputColumns.findIndex(c => c.name === 'Outreach');

                const buildOutputRow = (rowObj, tableContext) => {
                    const cells = [];
                    const formulas = [];
                    let rowColor = null;
                    let cellHighlights = [];
                    let comments = [];
                    // Normalize to string to match the normalized Student History keys.
                    const rawSId = rowObj.values[studentIdIdx];
                    const sId = (rawSId === null || rawSId === undefined || rawSId === '') ? rawSId : String(rawSId).trim();
                    // Pull the master's Student Number too (when it's a distinct
                    // column) so legacy history rows keyed under it still match.
                    const rawSNumber = (studentNumberIdx !== -1 && studentNumberIdx !== studentIdIdx) ? rowObj.values[studentNumberIdx] : null;
                    const sNumber = (rawSNumber === null || rawSNumber === undefined || rawSNumber === '') ? null : String(rawSNumber).trim();
                    const sIds = [sId, sNumber].filter((v, i, a) => v && a.indexOf(v) === i);
                    const missingVal = (missingIdx !== -1) ? rowObj.values[missingIdx] : null;
                    const nextAssignmentDueVal = (nextAssignmentDueIdx !== -1) ? rowObj.values[nextAssignmentDueIdx] : null;
                    const retentionMsg = getRetentionMessage(sIds, ldaFollowUpMap, missingVal, tableContext, dncMap, dncMessageMap, nextAssignmentDueVal, nextAssignmentDueColumnAllBlank, settings.includeNextAssignmentDue, settings.expandedDNC, settings.authorFormat);
                    const isGradeBookFlag = retentionMsg && retentionMsg.includes("Please check their Grade Book");
                    const isRetentionActive = !!retentionMsg && !isGradeBookFlag;
                    const dncTagText = lookupTagEntry(dncMap, sIds);
                    const dncTagsArr = dncTagText ? dncTagText.split(',').map(t => t.trim()) : [];
                    const isGeneralDnc = dncTagsArr.includes('dnc');
                    // Expanded DNC: a "DNC" or "DNC - Phone" writes its outreach message but gets
                    // no row highlight — only the Outreach cell is filled (settings.dncColor),
                    // and its phone strikethrough uses the same color with default text.
                    const isDncOutreach = settings.expandedDNC && (isGeneralDnc || dncTagsArr.includes('dnc - phone'));
                    const suppressRetentionHighlight = isDncOutreach;
                    // LDA-tag follow-up rows highlight in settings.ldaTagColor (default yellow).
                    const isLdaFollowUp = !!retentionMsg && retentionMsg.startsWith('Student will engage');
                    const isNextAssignmentDue = retentionMsg && retentionMsg.startsWith("Student's next assignment is due");

                    // --- Course Start Baseline Check ---
                    let courseStartMsg = null;
                    if (courseStartIdx !== -1 && courseStartModeValue !== null) {
                        const csVal = rowObj.values[courseStartIdx];
                        const csIsEmpty = (csVal === null || csVal === undefined || csVal === '');
                        if (csIsEmpty) {
                            courseStartMsg = 'Course Start is not listed';
                        } else if (String(csVal) !== String(courseStartModeValue)) {
                            const csNum = typeof csVal === 'number' ? csVal : parseFloat(csVal);
                            const modeNum = typeof courseStartModeValue === 'number' ? courseStartModeValue : parseFloat(courseStartModeValue);
                            const formattedDate = formatLongDate(csVal);
                            if (!isNaN(csNum) && !isNaN(modeNum) && csNum < modeNum) {
                                courseStartMsg = `Course Start is listed as ${formattedDate}`;
                            } else {
                                courseStartMsg = `Course Start is listed as ${formattedDate}`;
                            }
                        }
                    }

                    let partialRowColor = "#FFEDD5";
                    if (retentionMsg === "Do not contact") {
                        partialRowColor = "#FFC7CE";
                    } else if (isLdaFollowUp) {
                        partialRowColor = settings.ldaTagColor;
                    } else if (isNextAssignmentDue) {
                        partialRowColor = "#e2efda";
                    }
                    if (isRetentionActive && !suppressRetentionHighlight && outreachColIndex === -1) {
                        rowColor = partialRowColor;
                    }
                    outputColumns.forEach((colConfig, colOutIdx) => {
                        const masterIdx = getColIndex(colConfig.name);
                        let val = (masterIdx !== -1) ? rowObj.values[masterIdx] : (colConfig.riskIndex ? lookupRiskCount(riskIndexMap, sIds) : "");
                        let form = (masterIdx !== -1) ? rowObj.formulas[masterIdx] : null;
                        if (colConfig.name === 'Gradebook' && form && String(form).startsWith('=HYPERLINK')) {
                            // Keep formula
                        } else if (colConfig.name === 'Gradebook' && val && String(val).startsWith('http')) {
                            form = `=HYPERLINK("${val}", "Link")`;
                            val = "Link";
                        }
                        if (isRetentionActive && !suppressRetentionHighlight && outreachColIndex !== -1 && colOutIdx <= outreachColIndex) {
                            cellHighlights.push({ colIndex: colOutIdx, color: partialRowColor });
                        }
                        if (masterIdx !== -1 && val) {
                            const colMap = columnColorMaps.get(masterIdx);
                            if (colMap && colMap.has(String(val))) {
                                cellHighlights.push({ colIndex: colOutIdx, color: colMap.get(String(val)) });
                            }
                        }
                        // --- Advisor Auto-Assignment Override ---
                        if (colConfig.name === 'Assigned' && advisorAssignmentMap.size > 0) {
                            const advisorInfo = advisorAssignmentMap.get(rowObj.originalIndex);
                            if (advisorInfo) {
                                val = advisorInfo.name;
                                form = null;
                                cellHighlights.push({ colIndex: colOutIdx, color: advisorInfo.color });
                            }
                        }
                        if (colConfig.name === 'Outreach') {
                            if (retentionMsg && !isGradeBookFlag) {
                                val = retentionMsg;
                            }
                            if (isGradeBookFlag) {
                                comments.push({ colIndex: colOutIdx, text: retentionMsg });
                            }
                            if (courseStartMsg) {
                                comments.push({ colIndex: colOutIdx, text: courseStartMsg });
                            }
                            // Expanded DNC fills only the Outreach cell (no row highlight).
                            if (isDncOutreach) {
                                cellHighlights.push({ colIndex: colOutIdx, color: settings.dncColor });
                            }
                        }
                        // --- DNC Highlight (Highest Priority - Phone Columns) ---
                        // Only strikethrough the specific phone column matching the DNC type.
                        // General "DNC" strikes both. "DNC - Phone" only strikes Phone, etc.
                        if (settings.includeDNCTag && dncTagText) {
                            const dncTags = dncTagText.split(',').map(t => t.trim());
                            const hasGeneralDnc = dncTags.some(t => t === 'dnc');
                            const hasPhoneDnc = dncTags.some(t => t === 'dnc - phone');
                            const hasOtherPhoneDnc = dncTags.some(t => t === 'dnc - other phone');
                            // Expanded DNC: strikethrough fill follows the DNC color with default
                            // text. Legacy keeps the original red fill + dark-red text.
                            const strikeColor = settings.expandedDNC ? settings.dncColor : "#FFC7CE";
                            const strikeFontColor = settings.expandedDNC ? null : "#9C0006";

                            if (colConfig.name === 'Phone' && (hasGeneralDnc || hasPhoneDnc)) {
                                cellHighlights.push({ colIndex: colOutIdx, color: strikeColor, strikethrough: true, fontColor: strikeFontColor });
                            }
                            if (colConfig.name === 'Other Phone' && (hasGeneralDnc || hasOtherPhoneDnc)) {
                                cellHighlights.push({ colIndex: colOutIdx, color: strikeColor, strikethrough: true, fontColor: strikeFontColor });
                            }
                        }
                        cells.push(val);
                        formulas.push(form);
                    });
                    return { cells, formulas, rowColor, cellHighlights, comments };
                };

                // --- Process each campus ---
                if (onProgress) onProgress('createSheet', 'active');

                for (let ci = 0; ci < campusList.length; ci++) {
                    const campusName = campusList[ci];
                    if (onCampusProgress) onCampusProgress(campusName, ci, campusList.length, 'active');

                    // Filter data for this campus
                    const campusDataRows = dataRows.filter(r =>
                        String(r.values[campusIdx] || '').trim() === campusName
                    );
                    const campusFailingRows = failingRows.filter(r =>
                        String(r.values[campusIdx] || '').trim() === campusName
                    );
                    const campusAttendanceRows = attendanceRows.filter(r =>
                        String(r.values[campusIdx] || '').trim() === campusName
                    );
                    const campusSapRows = sapRows.filter(r =>
                        String(r.values[campusIdx] || '').trim() === campusName
                    );

                    // Create sheet for this campus
                    let sheetName = campusName;
                    let counter = 2;
                    sheets.load("items/name");
                    await context.sync();
                    const existingNames = sheets.items.map(s => s.name);
                    while (existingNames.includes(sheetName)) {
                        sheetName = `${campusName} (${counter++})`;
                    }

                    const newSheet = sheets.add(sheetName);
                    if (ci === 0) newSheet.activate();
                    await context.sync();

                    // Write LDA table for this campus
                    if (campusDataRows.length > 0) {
                        await writeTable(
                            context, newSheet, 0, `LDA_${ci}`,
                            outputColumns,
                            campusDataRows.map(r => buildOutputRow(r, 'LDA_Table')),
                            masterSheet, getColIndex, dateColumnIndices, null
                        );
                    } else {
                        const headerRange = newSheet.getRangeByIndexes(0, 0, 1, outputColumns.length);
                        headerRange.values = [outputColumns.map(c => c.name)];
                    }

                    // Write Failing table for this campus (if applicable)
                    let nextRow = campusDataRows.length + 4;
                    if (settings.includeFailingList && campusFailingRows.length > 0) {
                        const title = newSheet.getRangeByIndexes(nextRow - 1, 0, 1, 1);
                        title.values = [["Failing Students (Active)"]];
                        title.format.font.bold = true;

                        await writeTable(
                            context, newSheet, nextRow, `Failing_${ci}`,
                            outputColumns,
                            campusFailingRows.map(r => buildOutputRow(r, 'Failing_Table')),
                            masterSheet, getColIndex, dateColumnIndices, null
                        );
                        nextRow += campusFailingRows.length + 4;
                    }

                    // Write Attendance table for this campus (if applicable)
                    if (settings.includeAttendanceList && campusAttendanceRows.length > 0) {
                        const attTitle = newSheet.getRangeByIndexes(nextRow - 1, 0, 1, 1);
                        attTitle.values = [["Low Attendance Students"]];
                        attTitle.format.font.bold = true;

                        await writeTable(
                            context, newSheet, nextRow, `Attendance_${ci}`,
                            outputColumns,
                            campusAttendanceRows.map(r => buildOutputRow(r, 'Attendance_Table')),
                            masterSheet, getColIndex, dateColumnIndices, null
                        );
                        nextRow += campusAttendanceRows.length + 4;
                    }

                    // Write SAP table for this campus (if applicable)
                    if (settings.includeSapList && campusSapRows.length > 0) {
                        const sapTitle = newSheet.getRangeByIndexes(nextRow - 1, 0, 1, 1);
                        sapTitle.values = [["SAP Students (Not Met)"]];
                        sapTitle.format.font.bold = true;

                        await writeTable(
                            context, newSheet, nextRow, `SAP_${ci}`,
                            outputColumns,
                            campusSapRows.map(r => buildOutputRow(r, 'SAP_Table')),
                            masterSheet, getColIndex, dateColumnIndices, null
                        );
                        nextRow += campusSapRows.length + 4;
                    }

                    // Autofit & hide columns
                    newSheet.getUsedRange().getEntireColumn().format.autofitColumns();

                    // Pin configured columns to fixed widths (Outreach also wraps) after autofit.
                    if (settings.compactColumnWidth) applyFixedColumnWidths(newSheet);
                    await context.sync();

                    // Batch hidden column operations to avoid queue overflow
                    const campusHiddenIndices = outputColumns
                        .map((colConfig, idx) => colConfig.hidden ? idx : -1)
                        .filter(idx => idx !== -1);
                    const CAMPUS_HIDE_BATCH = 50;
                    for (let hi = 0; hi < campusHiddenIndices.length; hi += CAMPUS_HIDE_BATCH) {
                        const batch = campusHiddenIndices.slice(hi, hi + CAMPUS_HIDE_BATCH);
                        batch.forEach(idx => {
                            newSheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn().columnHidden = true;
                        });
                        await context.sync();
                    }
                    if (onCampusProgress) onCampusProgress(campusName, ci, campusList.length, 'completed');
                }

                if (onProgress) onProgress('createSheet', 'completed');
                if (onProgress) onProgress('writing', 'completed');
                if (onProgress) onProgress('highlights', 'completed');
                if (onProgress) onProgress('hyperlinks', 'completed');
                if (onProgress) onProgress('finalize', 'completed');

            } else {
                // ============================================================
                // SINGLE / DATE FLOW (original behavior)
                // ============================================================

                // --- STEP 5: Creating Sheet ---
                if (onProgress) onProgress('createSheet', 'active');

                let baseName;
                if (settings.sheetNameMode === 'campus') {
                    let campusName = '';
                    if (campusIdx !== -1) {
                        for (let i = 1; i < masterValues.length; i++) {
                            const val = String(masterValues[i][campusIdx] || '').trim();
                            if (val) { campusName = val; break; }
                        }
                    }
                    baseName = campusName || 'Campus';
                } else {
                    const today = new Date();
                    const dateStr = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
                    baseName = `LDA ${dateStr}`;
                }
                let sheetName = baseName;

                let counter = 2;
                const existingNames = sheets.items.map(s => s.name);
                while (existingNames.includes(sheetName)) {
                    sheetName = `${baseName} (${counter++})`;
                }

                const newSheet = sheets.add(sheetName);
                newSheet.activate();
                await context.sync();
                if (onProgress) onProgress('createSheet', 'completed');


                // --- STEP 6: Applying Tags (Data Processing) ---
                if (onProgress) onProgress('tags', 'active');

                const dncMap = new Map();
                const dncMessageMap = new Map();
                const ldaFollowUpMap = new Map();

                if (historyData && studentIdIdx !== -1) {
                    const hValues = historyData.values;
                    if (hValues.length > 0) {
                        const hHeaders = hValues[0].map(h => String(h).toLowerCase().trim());
                        // Collect every ID-like column (SyStudentId, Student Number, Student ID, etc.).
                        // Student History rows written by newer code land in SyStudentId while older
                        // rows landed in Student Number; the Master List only exposes one of these,
                        // so we index the map under every non-empty variant from each history row.
                        const hIdIndices = [];
                        hHeaders.forEach((h, idx) => {
                            if ((h.includes('student') && h.includes('id')) || h.includes('number')) {
                                hIdIndices.push(idx);
                            }
                        });
                        const hTagIdx = hHeaders.indexOf('tag');
                        const hCommentIdx = hHeaders.indexOf('comment');
                        const hCreatedByIdx = hHeaders.indexOf('created by');

                        if (hIdIndices.length > 0 && hTagIdx !== -1) {
                            const todayTime = new Date().setHours(0,0,0,0);
                            const ldaRegex = /\blda\b.*?(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})/i;

                            for (let i = hValues.length - 1; i > 0; i--) {
                                // Normalize to strings so Map lookups match regardless of whether
                                // Excel returned the ID as a Number or a String.
                                const hids = hIdIndices
                                    .map(idx => hValues[i][idx])
                                    .filter(v => v !== null && v !== undefined && v !== '')
                                    .map(v => String(v).trim())
                                    .filter(v => v.length > 0);
                                const htagRaw = String(hValues[i][hTagIdx] || '');
                                const htagLower = htagRaw.toLowerCase().trim();

                                if (settings.includeDNCTag && hids.length > 0 && htagLower.includes('dnc')) {
                                    // Accumulate DNC tags so multiple entries (e.g. "DNC - Phone" and "DNC - Other Phone") are preserved
                                    for (const hid of hids) {
                                        dncMap.set(hid, dncMap.has(hid) ? dncMap.get(hid) + ', ' + htagLower : htagLower);
                                    }

                                    // Capture the comment text for the most recent qualifying DNC entry
                                    // (a bare "DNC" or "DNC - Phone" tag). That text becomes the Outreach
                                    // message — unless the comment is a Quote, in which case we only flag it.
                                    // Rows are walked newest-first, so the first match per student wins.
                                    const rowTags = htagLower.split(',').map(t => t.trim());
                                    const isQualifyingDnc = rowTags.includes('dnc') || rowTags.includes('dnc - phone');
                                    if (isQualifyingDnc) {
                                        const isQuote = rowTags.includes('quote');
                                        const message = hCommentIdx !== -1 ? String(hValues[i][hCommentIdx] || '').trim() : '';
                                        const createdBy = hCreatedByIdx !== -1 ? String(hValues[i][hCreatedByIdx] || '').trim() : '';
                                        for (const hid of hids) {
                                            if (!dncMessageMap.has(hid)) {
                                                dncMessageMap.set(hid, { message, isQuote, createdBy });
                                            }
                                        }
                                    }
                                }

                                if (settings.includeLDATag && hids.length > 0) {
                                    const match = htagRaw.match(ldaRegex);
                                    if (match) {
                                        const dateString = match[1];
                                        const ldaDate = new Date(dateString);
                                        if (!isNaN(ldaDate.getTime())) {
                                            ldaDate.setHours(0,0,0,0);
                                            if (ldaDate >= todayTime) {
                                                for (const hid of hids) {
                                                    if (!ldaFollowUpMap.has(hid)) {
                                                        ldaFollowUpMap.set(hid, { date: ldaDate, text: htagRaw, createdBy: hCreatedByIdx !== -1 ? String(hValues[i][hCreatedByIdx] || '').trim() : '' });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (onProgress) onProgress('tags', 'completed');


                // --- STEP 7: Writing Tables (data only; highlights & hyperlinks run after) ---

                // --- 7a. Data-Driven Date Column Detection ---
                const dateColumnIndices = new Set();
                outputColumns.forEach((colConfig) => {
                    const masterIdx = getColIndex(colConfig.name);
                    if (masterIdx === -1) return;

                    const excelHeader = String(masterValues[0][masterIdx] || '').toLowerCase();
                    const isExcluded = /id|no\.|num|code|zip|postal|social|ssn|phone|grade|score|credit|fee|days|count/i.test(excelHeader);
                    if (isExcluded) return;

                    let dateCount = 0;
                    let numCount = 0;
                    const limit = Math.min(masterValues.length, 100);
                    for (let i = 1; i < limit; i++) {
                        const val = masterValues[i][masterIdx];
                        if (typeof val === 'number') {
                            numCount++;
                            if (val > 10958 && val < 73051) {
                                dateCount++;
                            }
                        }
                    }
                    if (numCount > 0 && (dateCount / numCount) > 0.5) {
                        dateColumnIndices.add(colConfig.name);
                    }
                });

                // Determine Outreach Column Index (It might be missing if implicit columns pushed it)
                const outreachColIndex = outputColumns.findIndex(c => c.name === 'Outreach');

                // Row Builder
                const buildOutputRow = (rowObj, tableContext) => {
                    const cells = [];
                    const formulas = [];
                    let rowColor = null;
                    let cellHighlights = [];
                    let comments = [];

                    // Normalize to string to match the normalized Student History keys.
                    const rawSId = rowObj.values[studentIdIdx];
                    const sId = (rawSId === null || rawSId === undefined || rawSId === '') ? rawSId : String(rawSId).trim();
                    // Pull the master's Student Number too (when it's a distinct
                    // column) so legacy history rows keyed under it still match.
                    const rawSNumber = (studentNumberIdx !== -1 && studentNumberIdx !== studentIdIdx) ? rowObj.values[studentNumberIdx] : null;
                    const sNumber = (rawSNumber === null || rawSNumber === undefined || rawSNumber === '') ? null : String(rawSNumber).trim();
                    const sIds = [sId, sNumber].filter((v, i, a) => v && a.indexOf(v) === i);

                    // 1. Get critical values
                    const missingVal = (missingIdx !== -1) ? rowObj.values[missingIdx] : null;
                    const nextAssignmentDueVal = (nextAssignmentDueIdx !== -1) ? rowObj.values[nextAssignmentDueIdx] : null;

                    // 2. Generate Retention Message using helper
                    const retentionMsg = getRetentionMessage(sIds, ldaFollowUpMap, missingVal, tableContext, dncMap, dncMessageMap, nextAssignmentDueVal, nextAssignmentDueColumnAllBlank, settings.includeNextAssignmentDue, settings.expandedDNC, settings.authorFormat);
                    const dncTagText = lookupTagEntry(dncMap, sIds);
                    const dncTagsArr = dncTagText ? dncTagText.split(',').map(t => t.trim()) : [];
                    const isGeneralDnc = dncTagsArr.includes('dnc');
                    // Expanded DNC: a "DNC" or "DNC - Phone" writes its outreach message but gets
                    // no row highlight — only the Outreach cell is filled (settings.dncColor),
                    // and its phone strikethrough uses the same color with default text.
                    const isDncOutreach = settings.expandedDNC && (isGeneralDnc || dncTagsArr.includes('dnc - phone'));
                    const suppressRetentionHighlight = isDncOutreach;
                    // LDA-tag follow-up rows highlight in settings.ldaTagColor (default yellow).
                    const isLdaFollowUp = !!retentionMsg && retentionMsg.startsWith('Student will engage');

                    // 2b. Course Start Baseline Check
                    let courseStartMsg = null;
                    if (courseStartIdx !== -1 && courseStartModeValue !== null) {
                        const csVal = rowObj.values[courseStartIdx];
                        const csIsEmpty = (csVal === null || csVal === undefined || csVal === '');
                        if (csIsEmpty) {
                            courseStartMsg = 'Course Start is not listed';
                        } else if (String(csVal) !== String(courseStartModeValue)) {
                            const csNum = typeof csVal === 'number' ? csVal : parseFloat(csVal);
                            const modeNum = typeof courseStartModeValue === 'number' ? courseStartModeValue : parseFloat(courseStartModeValue);
                            const formattedDate = formatLongDate(csVal);
                            if (!isNaN(csNum) && !isNaN(modeNum) && csNum < modeNum) {
                                courseStartMsg = `Course Start is listed as ${formattedDate}`;
                            } else {
                                courseStartMsg = `Course Start is listed as ${formattedDate}`;
                            }
                        }
                    }

                    // 3. Determine Highlighting Logic
                    const isLda = !!lookupTagEntry(ldaFollowUpMap, sIds);
                    const isGradeBookFlag = retentionMsg && retentionMsg.includes("Please check their Grade Book");
                    const isRetentionActive = !!retentionMsg && !isGradeBookFlag;
                    const isNextAssignmentDue = retentionMsg && retentionMsg.startsWith("Student's next assignment is due");

                    // Determine Row/Partial Color:
                    let partialRowColor = "#FFEDD5"; // Orange Default
                    if (retentionMsg === "Do not contact") {
                        partialRowColor = "#FFC7CE"; // Red for legacy DNC (expanded off)
                    } else if (isLdaFollowUp) {
                        partialRowColor = settings.ldaTagColor; // Yellow (same as Contacted) for LDA tag
                    } else if (isNextAssignmentDue) {
                        partialRowColor = "#e2efda"; // Light green for zero missing + next assignment due
                    }

                    // Fallback row color if Outreach column is missing
                    if (isRetentionActive && !suppressRetentionHighlight && outreachColIndex === -1) {
                         rowColor = partialRowColor;
                    }

                    outputColumns.forEach((colConfig, colOutIdx) => {
                        const masterIdx = getColIndex(colConfig.name);
                        let val = (masterIdx !== -1) ? rowObj.values[masterIdx] : (colConfig.riskIndex ? lookupRiskCount(riskIndexMap, sIds) : "");
                        let form = (masterIdx !== -1) ? rowObj.formulas[masterIdx] : null;

                        if (colConfig.name === 'Gradebook' && form && String(form).startsWith('=HYPERLINK')) {
                            // Keep formula
                        } else if (colConfig.name === 'Gradebook' && val && String(val).startsWith('http')) {
                            form = `=HYPERLINK("${val}", "Link")`;
                            val = "Link";
                        }

                        // --- Apply Retention Highlight (Partial Row) ---
                        // Only apply up to outreach column if it exists and this column is within range
                        if (isRetentionActive && !suppressRetentionHighlight && outreachColIndex !== -1 && colOutIdx <= outreachColIndex) {
                            cellHighlights.push({
                                colIndex: colOutIdx,
                                color: partialRowColor
                            });
                        }

                        // --- Value-Based Color Mapping (Overrides Retention Highlight) ---
                        if (masterIdx !== -1 && val) {
                             const colMap = columnColorMaps.get(masterIdx);
                             if (colMap && colMap.has(String(val))) {
                                 cellHighlights.push({
                                     colIndex: colOutIdx,
                                     color: colMap.get(String(val))
                                 });
                             }
                        }

                        // --- Outreach Message Injection ---
                        // --- Advisor Auto-Assignment Override ---
                        if (colConfig.name === 'Assigned' && advisorAssignmentMap.size > 0) {
                            const advisorInfo = advisorAssignmentMap.get(rowObj.originalIndex);
                            if (advisorInfo) {
                                val = advisorInfo.name;
                                form = null;
                                cellHighlights.push({ colIndex: colOutIdx, color: advisorInfo.color });
                            }
                        }
                        if (colConfig.name === 'Outreach') {
                            if (retentionMsg && !isGradeBookFlag) {
                                val = retentionMsg;
                            }
                            if (isGradeBookFlag) {
                                comments.push({ colIndex: colOutIdx, text: retentionMsg });
                            }
                            if (courseStartMsg) {
                                comments.push({ colIndex: colOutIdx, text: courseStartMsg });
                            }
                            // Expanded DNC fills only the Outreach cell (no row highlight).
                            if (isDncOutreach) {
                                cellHighlights.push({ colIndex: colOutIdx, color: settings.dncColor });
                            }
                        }

                        // --- DNC Highlight (Highest Priority - Phone Columns) ---
                        // Only strikethrough the specific phone column matching the DNC type.
                        // General "DNC" strikes both. "DNC - Phone" only strikes Phone, etc.
                        if (settings.includeDNCTag && dncTagText) {
                            const dncTags = dncTagText.split(',').map(t => t.trim());
                            const hasGeneralDnc = dncTags.some(t => t === 'dnc');
                            const hasPhoneDnc = dncTags.some(t => t === 'dnc - phone');
                            const hasOtherPhoneDnc = dncTags.some(t => t === 'dnc - other phone');
                            // Expanded DNC: strikethrough fill follows the DNC color with default
                            // text. Legacy keeps the original red fill + dark-red text.
                            const strikeColor = settings.expandedDNC ? settings.dncColor : "#FFC7CE";
                            const strikeFontColor = settings.expandedDNC ? null : "#9C0006";

                            if (colConfig.name === 'Phone' && (hasGeneralDnc || hasPhoneDnc)) {
                                cellHighlights.push({ colIndex: colOutIdx, color: strikeColor, strikethrough: true, fontColor: strikeFontColor });
                            }
                            if (colConfig.name === 'Other Phone' && (hasGeneralDnc || hasOtherPhoneDnc)) {
                                cellHighlights.push({ colIndex: colOutIdx, color: strikeColor, strikethrough: true, fontColor: strikeFontColor });
                            }
                        }

                        cells.push(val);
                        formulas.push(form);
                    });

                    return { cells, formulas, rowColor, cellHighlights, comments };
                };

                // Build processed-row arrays up front so all phases can share them.
                const tablePlan = [];

                if (dataRows.length > 0) {
                    tablePlan.push({
                        tableName: 'LDA_Table',
                        startRow: 0,
                        rows: dataRows.map(r => buildOutputRow(r, 'LDA_Table')),
                        title: null,
                    });
                } else {
                    // No data rows — still write headers so the sheet isn't empty.
                    newSheet.getRangeByIndexes(0, 0, 1, outputColumns.length).values = [outputColumns.map(c => c.name)];
                }

                let nextRow = dataRows.length + 4;
                if (settings.includeFailingList && failingRows.length > 0) {
                    tablePlan.push({
                        tableName: 'Failing_Table',
                        startRow: nextRow,
                        rows: failingRows.map(r => buildOutputRow(r, 'Failing_Table')),
                        title: { row: nextRow - 1, text: 'Failing Students (Active)' },
                    });
                    nextRow += failingRows.length + 4;
                }

                if (settings.includeAttendanceList && attendanceRows.length > 0) {
                    tablePlan.push({
                        tableName: 'Attendance_Table',
                        startRow: nextRow,
                        rows: attendanceRows.map(r => buildOutputRow(r, 'Attendance_Table')),
                        title: { row: nextRow - 1, text: 'Low Attendance Students' },
                    });
                    nextRow += attendanceRows.length + 4;
                }

                if (settings.includeSapList && sapRows.length > 0) {
                    tablePlan.push({
                        tableName: 'SAP_Table',
                        startRow: nextRow,
                        rows: sapRows.map(r => buildOutputRow(r, 'SAP_Table')),
                        title: { row: nextRow - 1, text: 'SAP Students (Not Met)' },
                    });
                    nextRow += sapRows.length + 4;
                }

                // Write any titles + table data (phase: writing). Titles need to exist
                // before the tables below them are created.
                if (onProgress) onProgress('writing', 'active');
                for (const plan of tablePlan) {
                    if (plan.title) {
                        const titleRange = newSheet.getRangeByIndexes(plan.title.row, 0, 1, 1);
                        titleRange.values = [[plan.title.text]];
                        titleRange.format.font.bold = true;
                    }
                    const wrapProgress = onBatchProgress
                        ? (current, total, phase) => onBatchProgress(current, total, phase, plan.tableName)
                        : null;
                    plan.ctx = await writeTableData(
                        context, newSheet, plan.startRow, plan.tableName,
                        outputColumns, plan.rows, masterSheet, getColIndex,
                        dateColumnIndices, wrapProgress
                    );
                }
                if (onProgress) onProgress('writing', 'completed');

                // Highlights phase across all tables.
                if (onProgress) onProgress('highlights', 'active');
                for (const plan of tablePlan) {
                    const wrapProgress = onBatchProgress
                        ? (current, total, phase) => onBatchProgress(current, total, phase, plan.tableName)
                        : null;
                    await applyTableHighlights(context, newSheet, plan.ctx, plan.rows, wrapProgress);
                }
                if (onProgress) onProgress('highlights', 'completed');

                // Hyperlinks phase across all tables (2D-array batched per column-run).
                if (onProgress) onProgress('hyperlinks', 'active');
                for (const plan of tablePlan) {
                    const wrapProgress = onBatchProgress
                        ? (current, total, phase) => onBatchProgress(current, total, phase, plan.tableName)
                        : null;
                    await applyTableFormulas(context, newSheet, plan.ctx, plan.rows, wrapProgress);
                }
                if (onProgress) onProgress('hyperlinks', 'completed');

                // Comments (silent; rare & fast).
                for (const plan of tablePlan) {
                    await applyTableComments(context, newSheet, plan.ctx, plan.rows);
                }

                // Autofit
                newSheet.getUsedRange().getEntireColumn().format.autofitColumns();

                // Pin configured columns to fixed widths (Outreach also wraps) after autofit.
                if (settings.compactColumnWidth) applyFixedColumnWidths(newSheet);

                // --- Apply hidden columns (must be LAST after autofit) ---
                const HIDE_BATCH = 50;
                const hiddenIndices = outputColumns
                    .map((colConfig, idx) => colConfig.hidden ? idx : -1)
                    .filter(idx => idx !== -1);
                for (let i = 0; i < hiddenIndices.length; i += HIDE_BATCH) {
                    const batch = hiddenIndices.slice(i, i + HIDE_BATCH);
                    batch.forEach(idx => {
                        newSheet.getRangeByIndexes(0, idx, 1, 1).getEntireColumn().columnHidden = true;
                    });
                    await context.sync();
                }

                // --- STEP 8: Finalize ---
                if (onProgress) onProgress('finalize', 'active');
                await context.sync();
                if (onProgress) onProgress('finalize', 'completed');
            }

        });

    } catch (error) {
        console.error("LDA Generation Error:", error);
        throw error;
    } finally {
        // Defer clearing the Outreach handler suppression flag. Excel can
        // dispatch onChanged events for the writes we just made AFTER this
        // function returns; if we cleared the flag synchronously, those
        // late events would run StudentView's Outreach handler, which
        // would interpret each LDA row write as a user comment and call
        // addComment for every row — which crashes the add-in. Matches
        // the 2s grace period the Chrome extension already uses.
        if (hadWindow) {
            setTimeout(() => {
                window.__srkSuppressOutreachHandler = previousSuppress || false;
            }, 2000);
        }
    }
}

/**
 * Writes the table's headers + data + creates the table + copies CF from the
 * Master List. Returns a context object the later phases can reuse.
 */
async function writeTableData(context, sheet, startRow, tableName, outputColumns, processedRows, masterSheet, getColIndex, dateColumnNames, onBatchProgress = null) {
    if (processedRows.length === 0) return null;

    const rowCount = processedRows.length;
    const colCount = outputColumns.length;
    const headers = outputColumns.map(c => c.name);

    // --- STEP 1: Headers ---
    sheet.getRangeByIndexes(startRow, 0, 1, colCount).values = [headers];
    await context.sync();

    // --- STEP 2: Data rows (chunked) ---
    const totalDataBatches = Math.ceil(rowCount / BATCH_SIZE);
    let currentBatch = 0;

    for (let batchStart = 0; batchStart < rowCount; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, rowCount);
        const batchValues = processedRows.slice(batchStart, batchEnd).map(r => r.cells);

        sheet.getRangeByIndexes(
            startRow + 1 + batchStart, 0, batchValues.length, colCount
        ).values = batchValues;

        await context.sync();

        currentBatch++;
        if (onBatchProgress) onBatchProgress(currentBatch, totalDataBatches, 'writing');
    }

    // --- STEP 3: Create table ---
    const fullRange = sheet.getRangeByIndexes(startRow, 0, rowCount + 1, colCount);
    const table = sheet.tables.add(fullRange, true);
    // Excel table names must be unique across the whole workbook and may only
    // contain letters, numbers, periods, and underscores. Date + time keeps
    // names readable while guaranteeing uniqueness across runs.
    const now = new Date();
    const datePart = `${now.getMonth() + 1}_${now.getDate()}_${now.getFullYear()}`;
    const timePart = `${now.getHours()}_${now.getMinutes()}_${now.getSeconds()}_${now.getMilliseconds()}`;
    table.name = `${tableName}_${datePart}_${timePart}`;
    table.style = "TableStyleLight9";
    try {
        await context.sync();
    } catch (e) {
        console.error(`Table creation sync failed for "${tableName}" (startRow=${startRow}, rows=${rowCount}, cols=${colCount})`, e);
        throw e;
    }

    // --- STEP 4: Copy CF from Master List ---
    const cfChecks = [];
    outputColumns.forEach((colConfig, idx) => {
        const masterIdx = getColIndex(colConfig.name);
        if (masterIdx !== -1) {
            cfChecks.push({
                sourceCell: masterSheet.getCell(1, masterIdx),
                targetIndex: idx,
                colName: colConfig.name,
            });
        }
    });

    for (let i = 0; i < cfChecks.length; i += BATCH_SIZE) {
        const batch = cfChecks.slice(i, i + BATCH_SIZE);
        batch.forEach(check => {
            const targetColRange = table.columns.getItemAt(check.targetIndex).getDataBodyRange();
            targetColRange.copyFrom(check.sourceCell, Excel.RangeCopyType.formats);
            targetColRange.format.fill.clear();
            if (dateColumnNames.has(check.colName)) {
                targetColRange.numberFormat = [["mm-dd-yy;@"]];
            }
        });
        try {
            await context.sync();
        } catch (e) {
            const colNames = batch.map(c => c.colName).join(', ');
            console.error(`CF copy sync failed on columns: [${colNames}]`, e);
            throw e;
        }
    }

    return { table, startRow, rowCount, colCount, tableName };
}

/**
 * Applies row colors and cell highlights. Uses merged range calls and
 * MAX_OPS_PER_SYNC-aware early flushing.
 */
async function applyTableHighlights(context, sheet, tableCtx, processedRows, onBatchProgress = null) {
    if (!tableCtx) return;
    const { startRow, rowCount, colCount } = tableCtx;
    const totalFormatBatches = Math.ceil(rowCount / FORMAT_BATCH_SIZE);
    let formatBatch = 0;

    for (let batchStart = 0; batchStart < rowCount; batchStart += FORMAT_BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + FORMAT_BATCH_SIZE, rowCount);

        const rowColorOps = [];
        const cellColorOps = [];

        for (let idx = batchStart; idx < batchEnd; idx++) {
            const r = processedRows[idx];
            if (r.rowColor) rowColorOps.push({ rowIdx: idx, color: r.rowColor });
            if (r.cellHighlights.length > 0) {
                const sorted = [...r.cellHighlights].sort((a, b) => a.colIndex - b.colIndex);
                let current = null;
                for (const h of sorted) {
                    if (current && current.color === h.color && current.endCol === h.colIndex - 1 && !h.strikethrough && !current.strikethrough) {
                        current.endCol = h.colIndex;
                    } else {
                        if (current) cellColorOps.push(current);
                        current = {
                            rowIdx: idx,
                            startCol: h.colIndex,
                            endCol: h.colIndex,
                            color: h.color,
                            strikethrough: h.strikethrough || false,
                            fontColor: h.fontColor || null,
                        };
                    }
                }
                if (current) cellColorOps.push(current);
            }
        }

        let queuedOps = 0;
        const flushIfFull = async () => {
            if (queuedOps >= MAX_OPS_PER_SYNC) {
                await context.sync();
                queuedOps = 0;
            }
        };

        if (rowColorOps.length > 0) {
            rowColorOps.sort((a, b) => a.rowIdx - b.rowIdx);
            let runStart = rowColorOps[0].rowIdx;
            let runEnd = runStart;
            let runColor = rowColorOps[0].color;
            const flushRun = async () => {
                sheet.getRangeByIndexes(startRow + 1 + runStart, 0, runEnd - runStart + 1, colCount)
                    .format.fill.color = runColor;
                queuedOps++;
                await flushIfFull();
            };
            for (let i = 1; i < rowColorOps.length; i++) {
                const op = rowColorOps[i];
                if (op.color === runColor && op.rowIdx === runEnd + 1) {
                    runEnd = op.rowIdx;
                } else {
                    await flushRun();
                    runStart = op.rowIdx;
                    runEnd = op.rowIdx;
                    runColor = op.color;
                }
            }
            await flushRun();
        }

        for (const op of cellColorOps) {
            const range = sheet.getRangeByIndexes(
                startRow + 1 + op.rowIdx, op.startCol, 1, op.endCol - op.startCol + 1
            );
            range.format.fill.color = op.color;
            queuedOps++;
            if (op.strikethrough) {
                range.format.font.strikethrough = true;
                queuedOps += 1;
                // Only override font color when specified; otherwise leave Excel's default.
                if (op.fontColor) {
                    range.format.font.color = op.fontColor;
                    queuedOps += 1;
                }
            }
            await flushIfFull();
        }

        try {
            await context.sync();
        } catch (e) {
            console.error(`Highlights sync failed on batch ${formatBatch + 1}/${totalFormatBatches} (rows ${batchStart}-${batchEnd - 1}). ` +
                `Ops: ${rowColorOps.length} rowColor, ${cellColorOps.length} cellColor.`, e);
            throw e;
        }

        formatBatch++;
        if (onBatchProgress) onBatchProgress(formatBatch, totalFormatBatches, 'highlights');
    }
}

/**
 * Applies cell formulas (usually =HYPERLINK). Groups by column and writes
 * contiguous runs with a single 2D-array assignment per run, collapsing
 * per-cell API ops into per-run ops.
 */
async function applyTableFormulas(context, sheet, tableCtx, processedRows, onBatchProgress = null) {
    if (!tableCtx) return;
    const { startRow, rowCount } = tableCtx;

    // Group formula entries by column.
    const byCol = new Map();
    for (let idx = 0; idx < rowCount; idx++) {
        const formulas = processedRows[idx].formulas;
        for (let cIdx = 0; cIdx < formulas.length; cIdx++) {
            const f = formulas[cIdx];
            if (!f) continue;
            if (!byCol.has(cIdx)) byCol.set(cIdx, []);
            byCol.get(cIdx).push({ rowIdx: idx, formula: f });
        }
    }

    if (byCol.size === 0) return;

    // Build the list of contiguous runs across all columns.
    const runs = []; // {colIdx, startRowIdx, formulas: string[][]}
    for (const [colIdx, entries] of byCol) {
        entries.sort((a, b) => a.rowIdx - b.rowIdx);
        let runStart = entries[0].rowIdx;
        let runFormulas = [[entries[0].formula]];
        for (let i = 1; i < entries.length; i++) {
            const e = entries[i];
            if (e.rowIdx === runStart + runFormulas.length) {
                runFormulas.push([e.formula]);
            } else {
                runs.push({ colIdx, startRowIdx: runStart, formulas: runFormulas });
                runStart = e.rowIdx;
                runFormulas = [[e.formula]];
            }
        }
        runs.push({ colIdx, startRowIdx: runStart, formulas: runFormulas });
    }

    const totalRuns = runs.length;
    let queuedOps = 0;
    let flushedRuns = 0;
    let lastReportedRun = 0;

    const flushIfFull = async (forceProgress = false) => {
        if (queuedOps >= MAX_OPS_PER_SYNC || forceProgress) {
            try {
                await context.sync();
            } catch (e) {
                console.error(`Hyperlinks sync failed (runs flushed=${flushedRuns}/${totalRuns}, queuedOps=${queuedOps})`, e);
                throw e;
            }
            queuedOps = 0;
            if (onBatchProgress && flushedRuns > lastReportedRun) {
                onBatchProgress(flushedRuns, totalRuns, 'hyperlinks');
                lastReportedRun = flushedRuns;
            }
        }
    };

    for (const run of runs) {
        const range = sheet.getRangeByIndexes(
            startRow + 1 + run.startRowIdx, run.colIdx, run.formulas.length, 1
        );
        range.formulas = run.formulas;
        queuedOps++;
        flushedRuns++;
        await flushIfFull();
    }

    // Final flush for remainder.
    if (queuedOps > 0) {
        try {
            await context.sync();
        } catch (e) {
            console.error(`Hyperlinks final sync failed (runs=${totalRuns}, queuedOps=${queuedOps})`, e);
            throw e;
        }
    }
    if (onBatchProgress) onBatchProgress(totalRuns, totalRuns, 'hyperlinks');
}

/**
 * Applies Excel comments (review threads) for course-start anomalies.
 */
async function applyTableComments(context, sheet, tableCtx, processedRows) {
    if (!tableCtx) return;
    const { startRow, rowCount } = tableCtx;

    const commentOps = [];
    for (let idx = 0; idx < rowCount; idx++) {
        const r = processedRows[idx];
        if (r.comments && r.comments.length > 0) {
            for (const c of r.comments) {
                commentOps.push({ row: startRow + 1 + idx, col: c.colIndex, text: c.text });
            }
        }
    }
    if (commentOps.length === 0) return;

    try {
        for (let i = 0; i < commentOps.length; i += FORMAT_BATCH_SIZE) {
            const batch = commentOps.slice(i, i + FORMAT_BATCH_SIZE);
            for (const op of batch) {
                const cellRange = sheet.getRangeByIndexes(op.row, op.col, 1, 1);
                context.workbook.comments.add(cellRange, op.text);
            }
            await context.sync();
        }
    } catch (e) {
        console.warn('Excel comments API not supported or failed, skipping comments:', e.message);
    }
}

/**
 * Convenience wrapper: writes a single table end-to-end (data, highlights,
 * hyperlinks, comments). Used by the multi-campus flow where each campus
 * owns its sheet and interleaved phases are fine.
 */
async function writeTable(context, sheet, startRow, tableName, outputColumns, processedRows, masterSheet, getColIndex, dateColumnNames, onBatchProgress = null) {
    const tableCtx = await writeTableData(
        context, sheet, startRow, tableName, outputColumns, processedRows,
        masterSheet, getColIndex, dateColumnNames, onBatchProgress
    );
    if (!tableCtx) return;
    await applyTableHighlights(context, sheet, tableCtx, processedRows, onBatchProgress);
    await applyTableFormulas(context, sheet, tableCtx, processedRows, onBatchProgress);
    await applyTableComments(context, sheet, tableCtx, processedRows);
}

/**
 * Checks whether the Master List sheet exists in the active workbook.
 * Used by LDAManager to short-circuit to a "missing master list" status
 * page before showing the settings form, since LDA cannot be generated
 * without a Master List.
 * @returns {Promise<boolean>} True if the Master List sheet is present.
 */
export async function checkMasterListExists() {
    try {
        let exists = false;
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            exists = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
        });
        return exists;
    } catch (e) {
        console.warn('checkMasterListExists failed, assuming present:', e);
        // Fail open: don't trap users behind the status page if the check itself errors.
        return true;
    }
}

/**
 * Checks if key LDA columns (Outreach, Assigned) are missing from the Master List headers.
 * Uses the same space-insensitive matching as the main LDA processor.
 * @returns {Promise<{outreach: boolean, assigned: boolean}>} Object indicating which columns are missing.
 */
export async function checkMissingLDAColumns() {
    const result = { outreach: false, assigned: false };
    try {
        const workbookSettings = getWorkbookSettings(defaultColumns);
        const columns = workbookSettings.columns;

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) return;

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const headerRange = masterSheet.getRangeByIndexes(0, 0, 1, 100);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0].filter(h => h !== null && h !== "");
            const stripStr = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, '');

            const findCol = (settingName) => {
                const targetStripped = stripStr(settingName);
                const colConfig = columns.find(c => stripStr(c.name) === targetStripped);
                let aliases = [];
                if (colConfig && Array.isArray(colConfig.alias)) {
                    aliases = colConfig.alias;
                }
                const candidates = [settingName, ...aliases];
                return candidates.some(cand => headers.some(h => stripStr(h) === stripStr(cand)));
            };

            result.outreach = !findCol('Outreach');
            result.assigned = !findCol('Assigned');
        });
    } catch (error) {
        console.error("checkMissingLDAColumns: Error checking columns:", error);
    }
    return result;
}

/**
 * Adds missing columns to the Master List by appending them after the last used column.
 * @param {string[]} columnNames - Array of column header names to add (e.g., ['Outreach', 'Assigned']).
 * @returns {Promise<void>}
 */
export async function addColumnsToMasterList(columnNames) {
    if (!columnNames || columnNames.length === 0) return;
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.MASTER_LIST);
            const usedRange = sheet.getUsedRange();
            usedRange.load("columnCount");
            await context.sync();

            const startCol = usedRange.columnCount;
            const headerRange = sheet.getRangeByIndexes(0, startCol, 1, columnNames.length);
            headerRange.values = [columnNames];
            await context.sync();

            console.log(`addColumnsToMasterList: Added columns [${columnNames.join(', ')}] at index ${startCol}`);
        });
    } catch (error) {
        console.error("addColumnsToMasterList: Error adding columns:", error);
        throw error;
    }
}

/**
 * Detect unique campus values from the Master List.
 * Reads only the Campus column in batches for efficiency.
 * @returns {Promise<string[]>} Array of unique campus names, sorted alphabetically
 */
export async function detectCampuses() {
    try {
        return await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) return [];

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("rowCount, columnCount");
            await context.sync();

            const totalRows = masterRange.rowCount;
            const totalCols = masterRange.columnCount;

            // Read headers only
            const headerRange = masterSheet.getRangeByIndexes(0, 0, 1, totalCols);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0];

            // Find Campus column with case/space insensitive matching
            const campusIdx = headers.findIndex(h => {
                const stripped = String(h || '').trim().toLowerCase().replace(/\s+/g, '');
                return stripped === 'campus';
            });

            if (campusIdx === -1) return [];

            // Read just the campus column in batches
            const campusSet = new Set();
            for (let startRow = 1; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = masterSheet.getRangeByIndexes(startRow, campusIdx, rowsToRead, 1);
                batchRange.load("values");
                await context.sync();

                for (const row of batchRange.values) {
                    const val = String(row[0] || '').trim();
                    if (val) campusSet.add(val);
                }
            }

            return Array.from(campusSet).sort();
        });
    } catch (error) {
        console.error("Campus detection error:", error);
        return [];
    }
}

/**
 * Scans the Master List to return all unique ProgramVersion values.
 * @returns {Promise<string[]>} Sorted array of unique ProgramVersion values.
 */
export async function detectProgramVersions() {
    try {
        return await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) return [];

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("rowCount, columnCount");
            await context.sync();

            const totalRows = masterRange.rowCount;
            const totalCols = masterRange.columnCount;

            const headerRange = masterSheet.getRangeByIndexes(0, 0, 1, totalCols);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0];
            const pvAliases = ['programversion', 'program', 'progversdescrip'];
            const pvIdx = headers.findIndex(h => {
                const stripped = String(h || '').trim().toLowerCase().replace(/\s+/g, '');
                return pvAliases.includes(stripped);
            });

            if (pvIdx === -1) return [];

            const pvSet = new Set();
            for (let startRow = 1; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = masterSheet.getRangeByIndexes(startRow, pvIdx, rowsToRead, 1);
                batchRange.load("values");
                await context.sync();

                for (const row of batchRange.values) {
                    const val = String(row[0] || '').trim();
                    if (val) pvSet.add(val);
                }
            }

            return Array.from(pvSet).sort();
        });
    } catch (error) {
        console.error("ProgramVersion detection error:", error);
        return [];
    }
}

/**
 * Predicts how students would be distributed among advisors on the LDA list.
 * Runs a lightweight version of the Days Out + Failing + Attendance filters,
 * then applies the advisor assignment algorithm.
 * @param {object} ldaSettings - The current LDA settings
 * @param {Array} advisors - Array of advisor config objects
 * @returns {Promise<Array>} Array of { id, name, color, count } per advisor
 */
export async function predictAdvisorDistribution(ldaSettings, advisors) {
    if (!advisors || advisors.length === 0) return [];

    try {
        return await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const hasMasterList = sheets.items.some(s => s.name === SHEET_NAMES.MASTER_LIST);
            if (!hasMasterList) return [];

            const masterSheet = sheets.getItem(SHEET_NAMES.MASTER_LIST);
            const masterRange = masterSheet.getUsedRange();
            masterRange.load("rowCount, columnCount");
            await context.sync();

            const totalRows = masterRange.rowCount;
            const totalCols = masterRange.columnCount;

            const headerRange = masterSheet.getRangeByIndexes(0, 0, 1, totalCols);
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0];
            const stripStr = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, '');

            // Find column indices
            const daysOutIdx = headers.findIndex(h => {
                const s = stripStr(h);
                return s === 'daysout' || s === 'days out';
            });
            const pvAliases = ['programversion', 'program', 'progversdescrip'];
            const pvIdx = headers.findIndex(h => pvAliases.includes(stripStr(h)));
            const gradeIdx = headers.findIndex(h => {
                const s = stripStr(h);
                return s === 'grade' || s === 'coursegrade' || s === 'currentscore' || s === 'currentgrade';
            });
            const attIdx = headers.findIndex(h => {
                const s = stripStr(h);
                return s === 'attendance%' || s === 'attendancepercent' || s === 'attendance';
            });
            const sapIdx = headers.findIndex(h => SAP_STATUS_ALIASES.some(a => stripStr(h) === stripStr(a)));
            const studentIdIdx = headers.findIndex(h => {
                const s = stripStr(h);
                return s === 'studentnumber' || s === 'studentid';
            });

            if (daysOutIdx === -1) return [];

            // Read all rows in batches and categorize
            const students = []; // { programVersion, daysOut, listType }
            const ldaStudentIds = new Set();

            // First pass: collect all rows
            const allRows = [];
            for (let startRow = 1; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = masterSheet.getRangeByIndexes(startRow, 0, rowsToRead, totalCols);
                batchRange.load("values");
                await context.sync();
                for (const row of batchRange.values) allRows.push(row);
            }

            // LDA list (Days Out >= threshold)
            for (const row of allRows) {
                const daysOutVal = row[daysOutIdx];
                if (typeof daysOutVal === 'number' && daysOutVal >= (ldaSettings.daysOut ?? 5)) {
                    students.push({
                        programVersion: pvIdx !== -1 ? String(row[pvIdx] || '').trim() : '',
                        daysOut: daysOutVal,
                        listType: 'lda'
                    });
                    const sId = studentIdIdx !== -1 ? row[studentIdIdx] : null;
                    if (sId) ldaStudentIds.add(sId);
                }
            }

            // Failing list (grade < 60% AND days out <= 4). SAP takes priority over
            // Failing, so a SAP-flagged student is counted on the SAP list, not here.
            if (ldaSettings.includeFailingList && gradeIdx !== -1) {
                for (const row of allRows) {
                    if (ldaSettings.includeSapList && sapIdx !== -1 && isSapFlagged(row[sapIdx])) continue;
                    const gradeVal = row[gradeIdx];
                    const daysOutVal = row[daysOutIdx];
                    const isFailing = (typeof gradeVal === 'number') && (gradeVal < 0.60 || (gradeVal >= 1 && gradeVal < 60));
                    const isRecent = (typeof daysOutVal === 'number') && (daysOutVal <= 4);
                    if (isFailing && isRecent) {
                        students.push({
                            programVersion: pvIdx !== -1 ? String(row[pvIdx] || '').trim() : '',
                            daysOut: typeof daysOutVal === 'number' ? daysOutVal : 0,
                            listType: 'failing'
                        });
                    }
                }
            }

            // SAP list (AdSAPStatus not "SAP Met"/blank AND not already on the LDA table)
            if (ldaSettings.includeSapList && sapIdx !== -1) {
                for (const row of allRows) {
                    if (!isSapFlagged(row[sapIdx])) continue;
                    const sId = studentIdIdx !== -1 ? row[studentIdIdx] : null;
                    if (sId && ldaStudentIds.has(sId)) continue;
                    const daysOutVal = row[daysOutIdx];
                    students.push({
                        programVersion: pvIdx !== -1 ? String(row[pvIdx] || '').trim() : '',
                        daysOut: typeof daysOutVal === 'number' ? daysOutVal : 0,
                        listType: 'sap'
                    });
                }
            }

            // Attendance list (attendance < 60% AND not already on LDA)
            if (ldaSettings.includeAttendanceList && attIdx !== -1) {
                for (const row of allRows) {
                    const sId = studentIdIdx !== -1 ? row[studentIdIdx] : null;
                    if (sId && ldaStudentIds.has(sId)) continue;
                    const attVal = row[attIdx];
                    let attPercent = null;
                    if (typeof attVal === 'number') {
                        attPercent = attVal <= 1 ? attVal * 100 : attVal;
                    }
                    if (attPercent !== null && attPercent < 60) {
                        const daysOutVal = row[daysOutIdx];
                        students.push({
                            programVersion: pvIdx !== -1 ? String(row[pvIdx] || '').trim() : '',
                            daysOut: typeof daysOutVal === 'number' ? daysOutVal : 0,
                            listType: 'attendance'
                        });
                    }
                }
            }

            const evenSplit = ldaSettings.advisorAssignment?.evenSplit ?? false;
            return assignStudentsToAdvisors(students, advisors, { evenSplit });
        });
    } catch (error) {
        console.error("Advisor distribution prediction error:", error);
        return [];
    }
}

/**
 * Core assignment algorithm: assigns students to advisors.
 * Filters: programVersions, listPreference (lda/failing/attendance), daysOut range.
 *
 * Phase 1: For each student, go through advisors top-to-bottom.
 *   An advisor "matches" if ALL of their active filters pass:
 *     - programVersions (non-empty): student PV must be in the set
 *     - listPreference (non-empty): student listType must be in the set
 *     - daysOutMin/daysOutMax (set): student daysOut must be in range
 *   If multiple advisors share the exact same filter signature, they share
 *   students via round-robin within that group.
 *   An advisor with NO filters set is skipped in Phase 1.
 *
 * Phase 2: Remaining unmatched students distributed round-robin across ALL advisors.
 *
 * @param {Array} students - Array of { programVersion, daysOut, listType }
 * @param {Array} advisors - Array of advisor config objects
 * @param {Object} [options] - Optional settings
 * @param {boolean} [options.evenSplit=false] - Redistribute to even out advisor counts
 * @returns {Array} Array of { id, name, color, count, students? }
 */
function assignStudentsToAdvisors(students, advisors, options = {}) {
    const counts = new Map();
    advisors.forEach(a => counts.set(a.id, 0));

    const assigned = new Array(students.length).fill(false);

    // Build filter signature for each advisor to detect shared filters
    const getFilterSig = (a) => {
        const parts = [];
        if (a.programVersions?.length) parts.push('pv:' + [...a.programVersions].sort().join(',').toLowerCase());
        if (a.listPreference?.length) parts.push('list:' + [...a.listPreference].sort().join(','));
        if (a.daysOutMin != null) parts.push('doMin:' + a.daysOutMin);
        if (a.daysOutMax != null) parts.push('doMax:' + a.daysOutMax);
        return parts.join('|');
    };

    const hasAnyFilter = (a) => {
        return (a.programVersions?.length > 0) ||
               (a.listPreference?.length > 0) ||
               (a.daysOutMin != null) ||
               (a.daysOutMax != null);
    };

    const advisorMatches = (advisor, student) => {
        if (advisor.programVersions?.length > 0) {
            const pvSet = new Set(advisor.programVersions.map(pv => pv.toLowerCase()));
            if (!pvSet.has((student.programVersion || '').toLowerCase())) return false;
        }
        if (advisor.listPreference?.length > 0) {
            if (!advisor.listPreference.includes(student.listType)) return false;
        }
        if (advisor.daysOutMin != null && typeof student.daysOut === 'number') {
            if (student.daysOut < advisor.daysOutMin) return false;
        }
        if (advisor.daysOutMax != null && typeof student.daysOut === 'number') {
            if (student.daysOut > advisor.daysOutMax) return false;
        }
        return true;
    };

    // Group advisors with identical filter signatures for shared round-robin
    const sigGroups = new Map(); // sig -> [advisor, ...]
    for (const advisor of advisors) {
        if (!hasAnyFilter(advisor)) continue;
        const sig = getFilterSig(advisor);
        if (!sigGroups.has(sig)) sigGroups.set(sig, []);
        sigGroups.get(sig).push(advisor);
    }

    // Phase 1: Filtered advisors claim matching students
    // Process groups in advisor-list order (by first advisor in group)
    const processedSigs = new Set();
    const rrCounters = new Map(); // sig -> round-robin counter

    for (const advisor of advisors) {
        if (!hasAnyFilter(advisor)) continue;
        const sig = getFilterSig(advisor);
        if (processedSigs.has(sig)) continue;
        processedSigs.add(sig);

        const group = sigGroups.get(sig);
        let rr = 0;

        for (let i = 0; i < students.length; i++) {
            if (assigned[i]) continue;
            // Check if the first advisor in this group matches (they all share the same filters)
            if (advisorMatches(group[0], students[i])) {
                const target = group[rr % group.length];
                counts.set(target.id, counts.get(target.id) + 1);
                assigned[i] = true;
                rr++;
            }
        }
    }

    // Phase 2: Remaining students distributed round-robin across ALL advisors
    // Track which advisor each student is assigned to for even split redistribution
    const studentAdvisorMap = new Array(students.length).fill(null);

    // Replay Phase 1 to build studentAdvisorMap
    const processedSigs2 = new Set();
    for (const advisor of advisors) {
        if (!hasAnyFilter(advisor)) continue;
        const sig = getFilterSig(advisor);
        if (processedSigs2.has(sig)) continue;
        processedSigs2.add(sig);
        const group = sigGroups.get(sig);
        let rr2 = 0;
        for (let i = 0; i < students.length; i++) {
            if (studentAdvisorMap[i]) continue;
            if (assigned[i] && advisorMatches(group[0], students[i])) {
                studentAdvisorMap[i] = group[rr2 % group.length].id;
                rr2++;
            }
        }
    }

    let rrIndex = 0;
    for (let i = 0; i < students.length; i++) {
        if (!assigned[i]) {
            // Try each advisor in round-robin order, preferring advisors whose
            // filters the student matches (or who have no filters).
            let placed = false;
            for (let attempt = 0; attempt < advisors.length; attempt++) {
                const advisor = advisors[(rrIndex + attempt) % advisors.length];
                if (!hasAnyFilter(advisor) || advisorMatches(advisor, students[i])) {
                    counts.set(advisor.id, counts.get(advisor.id) + 1);
                    studentAdvisorMap[i] = advisor.id;
                    rrIndex = (rrIndex + attempt + 1);
                    placed = true;
                    break;
                }
            }
            // If no advisor's filters match, assign round-robin so no student is left out
            if (!placed) {
                const advisor = advisors[rrIndex % advisors.length];
                counts.set(advisor.id, counts.get(advisor.id) + 1);
                studentAdvisorMap[i] = advisor.id;
                rrIndex++;
            }
        }
    }

    // Phase 3 (Even Split): Redistribute students to balance advisor counts
    if (options.evenSplit && advisors.length > 1) {
        const total = students.length;
        const base = Math.floor(total / advisors.length);
        const extra = total % advisors.length;
        // Target: first `extra` advisors (by current order) get base+1, rest get base
        const targets = new Map();
        advisors.forEach((a, idx) => targets.set(a.id, base + (idx < extra ? 1 : 0)));

        // Build per-advisor student index lists
        const advisorStudents = new Map();
        advisors.forEach(a => advisorStudents.set(a.id, []));
        for (let i = 0; i < students.length; i++) {
            advisorStudents.get(studentAdvisorMap[i]).push(i);
        }

        // Helper: move a student from one advisor to another
        const moveStudent = (stuIdx, fromId, toId) => {
            studentAdvisorMap[stuIdx] = toId;
            counts.set(fromId, counts.get(fromId) - 1);
            counts.set(toId, counts.get(toId) + 1);
            const fromList = advisorStudents.get(fromId);
            fromList.splice(fromList.indexOf(stuIdx), 1);
            advisorStudents.get(toId).push(stuIdx);
        };

        // Pass 1: Move students to under-allocated advisors that MATCH filters
        //         (best-fit by tightest daysOutMax first)
        for (const overAdv of advisors.filter(a => counts.get(a.id) > targets.get(a.id))) {
            const surplus = counts.get(overAdv.id) - targets.get(overAdv.id);
            if (surplus <= 0) continue;

            const indices = [...advisorStudents.get(overAdv.id)];
            indices.sort((a, b) => (students[a].daysOut || 0) - (students[b].daysOut || 0));

            let moved = 0;
            for (let si = 0; si < indices.length && moved < surplus; si++) {
                const stuIdx = indices[si];

                const eligible = advisors.filter(ua =>
                    ua.id !== overAdv.id &&
                    counts.get(ua.id) < targets.get(ua.id) &&
                    (!hasAnyFilter(ua) || advisorMatches(ua, students[stuIdx]))
                );
                if (eligible.length === 0) continue;

                eligible.sort((a, b) => {
                    const aMax = a.daysOutMax != null ? a.daysOutMax : Infinity;
                    const bMax = b.daysOutMax != null ? b.daysOutMax : Infinity;
                    if (aMax !== Infinity && bMax !== Infinity) return aMax - bMax;
                    if (aMax !== Infinity) return -1;
                    if (bMax !== Infinity) return 1;
                    return 0;
                });

                moveStudent(stuIdx, overAdv.id, eligible[0].id);
                moved++;
            }
        }

        // Pass 2: Force-balance remaining surplus by ignoring filters.
        //         Pick the most under-allocated advisor for each student.
        for (const overAdv of advisors.filter(a => counts.get(a.id) > targets.get(a.id))) {
            let surplus = counts.get(overAdv.id) - targets.get(overAdv.id);
            if (surplus <= 0) continue;

            const indices = [...advisorStudents.get(overAdv.id)];
            // Move students with the highest daysOut first (they're furthest
            // from the over-allocated advisor's likely specialty)
            indices.sort((a, b) => (students[b].daysOut || 0) - (students[a].daysOut || 0));

            for (let si = 0; si < indices.length && surplus > 0; si++) {
                const stuIdx = indices[si];

                // Find the most under-allocated advisor (regardless of filters)
                let bestTarget = null;
                let bestDeficit = 0;
                for (const ua of advisors) {
                    if (ua.id === overAdv.id) continue;
                    const deficit = targets.get(ua.id) - counts.get(ua.id);
                    if (deficit > bestDeficit) {
                        bestDeficit = deficit;
                        bestTarget = ua;
                    }
                }
                if (!bestTarget) break; // everyone is at or above target

                moveStudent(stuIdx, overAdv.id, bestTarget.id);
                surplus--;
            }
        }
    }

    // Build per-advisor student detail lists for debug/inspection
    const advisorStudentDetails = new Map();
    advisors.forEach(a => advisorStudentDetails.set(a.id, []));
    for (let i = 0; i < students.length; i++) {
        if (studentAdvisorMap[i] != null) {
            advisorStudentDetails.get(studentAdvisorMap[i]).push(students[i]);
        }
    }

    return advisors.map(a => ({
        id: a.id,
        name: a.name,
        color: a.color,
        count: counts.get(a.id),
        students: advisorStudentDetails.get(a.id)
    }));
}