/*
 * Risk Index — daily LDA/Days Out history and drop-episode counting.
 *
 * Every Create LDA run stores a snapshot of the whole Master List (SyStudentId,
 * LDA, Days Out) in the document settings under the `LDAHistory` key, keyed by
 * calendar date. Re-running on the same day overwrites that day's snapshot, so
 * there is at most one entry per day.
 *
 * A student's Risk Index is the number of separate drop episodes in which they
 * reached the risk threshold (default 14 days out). Episodes are keyed by the
 * student's LDA value: sitting at 14, 15, 16 days out across consecutive
 * snapshots shares one LDA and counts once; attending again (new LDA) and
 * dropping back to the threshold counts as a new episode. This also keeps the
 * count correct across weekends/holidays when no LDA sheet is created.
 */

/* global Office, Excel */

import { findColumnIndex, normalizeHeader } from '../../../../shared/excel-helpers.js';
import {
    STUDENT_ID_ALIASES,
    STUDENT_NUMBER_ALIASES,
    LAST_LDA_ALIASES,
    DAYS_OUT_ALIASES,
    STUDENT_NAME_ALIASES,
    GENDER_ALIASES,
    PROFILE_PICTURE_ALIASES,
} from '../../../../shared/columnAliases.js';
import { saveDocumentSetting, byteSize } from '../utility/documentSettings.js';

export const LDA_HISTORY_KEY = 'LDAHistory';

export const DEFAULT_RISK_INDEX_SETTINGS = {
    enabled: false,
    threshold: 14,
    showColumn: true,
};

export function sanitizeRiskIndexSettings(raw) {
    const defaults = { ...DEFAULT_RISK_INDEX_SETTINGS };
    if (!raw || typeof raw !== 'object') return defaults;
    const threshold = Number(raw.threshold);
    return {
        enabled: (raw.enabled !== undefined) ? !!raw.enabled : defaults.enabled,
        threshold: (Number.isFinite(threshold) && threshold >= 1) ? Math.floor(threshold) : defaults.threshold,
        showColumn: (raw.showColumn !== undefined) ? !!raw.showColumn : defaults.showColumn,
    };
}

// LDAHistory auto-prune configuration. The whole document-settings blob shares
// Excel's ~256 KB cap, so LDAHistory is budgeted well under it (workbookSettings
// and the other keys need the remaining headroom). Defaults: prune on, 150 KB.
export const DEFAULT_PRUNE_CONFIG = {
    enabled: true,
    maxSizeKB: 150,
    maxDays: 0, // 0 = no day-count limit; size is the only bound
};

export function sanitizePruneConfig(raw) {
    const d = { ...DEFAULT_PRUNE_CONFIG };
    if (!raw || typeof raw !== 'object') return d;
    const maxSizeKB = Number(raw.maxSizeKB);
    const maxDays = Number(raw.maxDays);
    return {
        enabled: (raw.enabled !== undefined) ? !!raw.enabled : d.enabled,
        maxSizeKB: (Number.isFinite(maxSizeKB) && maxSizeKB >= 1) ? Math.floor(maxSizeKB) : d.maxSizeKB,
        maxDays: (Number.isFinite(maxDays) && maxDays >= 0) ? Math.floor(maxDays) : d.maxDays,
    };
}

/**
 * Prune LDA history to fit the configured limits, dropping the OLDEST days
 * first (day keys are 'YYYY-MM-DD', so a lexical sort is chronological).
 * - maxDays > 0: keep at most that many most-recent days.
 * - maxSizeKB: drop oldest days until the serialized history fits the byte
 *   budget. The single most-recent day is never dropped, so today's snapshot
 *   always survives even if it alone exceeds the budget.
 * Both bounds apply only when `enabled`; disabling turns off all pruning.
 * Pure: returns { history, removed } and never touches Office.
 * @returns {{ history: {version:number, days:Object}, removed: string[] }}
 */
export function pruneHistory(history, config) {
    const cfg = sanitizePruneConfig(config);
    const srcDays = (history && typeof history === 'object' && history.days && typeof history.days === 'object')
        ? history.days
        : {};
    let keys = Object.keys(srcDays).sort(); // ascending → oldest first
    const removed = [];

    if (!cfg.enabled) {
        return { history: { version: 1, days: { ...srcDays } }, removed };
    }

    // Day-count cap.
    if (cfg.maxDays > 0 && keys.length > cfg.maxDays) {
        const dropCount = keys.length - cfg.maxDays;
        removed.push(...keys.slice(0, dropCount));
        keys = keys.slice(dropCount);
    }

    // Size cap: keep dropping the oldest until under budget (but never the last day).
    if (cfg.maxSizeKB > 0) {
        const budget = cfg.maxSizeKB * 1024;
        const sizeOf = (ks) => {
            const days = {};
            ks.forEach(k => { days[k] = srcDays[k]; });
            return byteSize({ version: 1, days });
        };
        while (keys.length > 1 && sizeOf(keys) > budget) {
            removed.push(keys[0]);
            keys = keys.slice(1);
        }
    }

    const days = {};
    keys.forEach(k => { days[k] = srcDays[k]; });
    return { history: { version: 1, days }, removed };
}

const pad2 = (n) => String(n).padStart(2, '0');

/** Local calendar date key, e.g. '2026-07-21'. */
export function todayKey(date = new Date()) {
    return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

/**
 * Normalize an LDA cell value to a stable episode key ('YYYY-MM-DD' when the
 * value is a recognizable date). Excel serials use UTC math so the calendar
 * day never shifts with the local timezone. Unparseable strings are kept
 * verbatim — they still work as episode keys. Empty values return null.
 */
export function normalizeLdaValue(val) {
    if (val === null || val === undefined || val === '') return null;
    if (val instanceof Date) {
        // ExcelJS parses xlsx date cells as UTC-midnight Dates.
        if (isNaN(val.getTime())) return null;
        return `${val.getUTCFullYear()}-${pad2(val.getUTCMonth() + 1)}-${pad2(val.getUTCDate())}`;
    }
    if (typeof val === 'number') {
        if (val < 20000 || val > 80000) return String(val);
        const d = new Date(Math.round((val - 25569) * 86400000));
        return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
    }
    const str = String(val).trim();
    if (!str) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1990 && parsed.getFullYear() < 2100) {
        return `${parsed.getFullYear()}-${pad2(parsed.getMonth() + 1)}-${pad2(parsed.getDate())}`;
    }
    return str;
}

/**
 * Build one day's snapshot from the Master List value matrix (row 0 = headers).
 * Rows without a student id or a numeric Days Out are skipped.
 * @returns {Object} { [studentId]: { lda: string|null, daysOut: number } }
 */
export function buildDailySnapshot(masterValues, { idIdx, ldaIdx, daysOutIdx }) {
    const snapshot = {};
    if (!Array.isArray(masterValues) || idIdx === -1 || daysOutIdx === -1) return snapshot;
    for (let i = 1; i < masterValues.length; i++) {
        const row = masterValues[i];
        if (!row) continue;
        const rawId = row[idIdx];
        if (rawId === null || rawId === undefined || rawId === '') continue;
        const id = String(rawId).trim();
        if (!id) continue;
        const daysOut = row[daysOutIdx];
        if (typeof daysOut !== 'number') continue;
        snapshot[id] = {
            lda: (ldaIdx !== -1) ? normalizeLdaValue(row[ldaIdx]) : null,
            daysOut,
        };
    }
    return snapshot;
}

/** Set (or overwrite) one date's snapshot in the history object. */
export function mergeSnapshot(history, dateKey, snapshot) {
    const days = (history && typeof history === 'object' && history.days && typeof history.days === 'object')
        ? { ...history.days }
        : {};
    days[dateKey] = snapshot;
    return { version: 1, days };
}

// Episode key for a snapshot entry: the LDA when known, otherwise a shared
// sentinel so unknown-LDA hits collapse into a single episode instead of
// counting every snapshot day separately.
const episodeKey = (entry) => entry.lda || 'unknown';

/**
 * Count drop episodes per student across the whole history.
 * @returns {Map<string, number>} studentId -> number of distinct LDAs that
 *   reached `threshold`+ days out.
 */
export function countRiskEpisodes(history, threshold) {
    const episodes = new Map(); // id -> Set of episode keys
    const days = (history && history.days) || {};
    for (const dateKey of Object.keys(days)) {
        const snapshot = days[dateKey];
        if (!snapshot || typeof snapshot !== 'object') continue;
        for (const id of Object.keys(snapshot)) {
            const entry = snapshot[id];
            if (!entry || typeof entry.daysOut !== 'number' || entry.daysOut < threshold) continue;
            if (!episodes.has(id)) episodes.set(id, new Set());
            episodes.get(id).add(episodeKey(entry));
        }
    }
    const counts = new Map();
    for (const [id, keys] of episodes) counts.set(id, keys.size);
    return counts;
}

/** Same count for a single student without walking every other student. */
export function countRiskEpisodesForStudent(history, studentId, threshold) {
    if (!studentId) return 0;
    const id = String(studentId).trim();
    const keys = new Set();
    const days = (history && history.days) || {};
    for (const dateKey of Object.keys(days)) {
        const entry = days[dateKey] && days[dateKey][id];
        if (!entry || typeof entry.daysOut !== 'number' || entry.daysOut < threshold) continue;
        keys.add(episodeKey(entry));
    }
    return keys.size;
}

/**
 * Parse a date-mode LDA sheet name ("LDA 7-21-2026", including same-day
 * re-run suffixes like "LDA 7-21-2026 (2)") into a history date key.
 * @returns {string|null} 'YYYY-MM-DD', or null when the name doesn't match.
 */
export function parseLdaSheetName(name) {
    const match = String(name || '').trim().match(/^lda (\d{1,2})-(\d{1,2})-(\d{4})(?: \(\d+\))?$/i);
    if (!match) return null;
    const [, month, day, year] = match;
    if (Number(month) < 1 || Number(month) > 12 || Number(day) < 1 || Number(day) > 31) return null;
    return `${year}-${pad2(Number(month))}-${pad2(Number(day))}`;
}

/**
 * Resolve the snapshot column indices from a raw header row (LDA sheets carry
 * the Master List's columns, some hidden). Student id prefers SyStudentId and
 * falls back to Student Number for legacy workbooks.
 * @returns {{ idIdx: number, ldaIdx: number, daysOutIdx: number }}
 */
export function resolveSnapshotIndices(headers) {
    const normalized = (Array.isArray(headers) ? headers : []).map(normalizeHeader);
    const idIdx = (() => {
        const sy = findColumnIndex(normalized, STUDENT_ID_ALIASES);
        return sy !== -1 ? sy : findColumnIndex(normalized, STUDENT_NUMBER_ALIASES);
    })();
    return {
        idIdx,
        ldaIdx: findColumnIndex(normalized, LAST_LDA_ALIASES),
        daysOutIdx: findColumnIndex(normalized, DAYS_OUT_ALIASES),
    };
}

/**
 * Merge imported day snapshots into the history WITHOUT overwriting dates that
 * already exist — real Create LDA snapshots always win over backfilled data.
 * @returns {{ history: Object, added: string[], skipped: string[] }}
 */
export function mergeImportedDays(history, importedDays) {
    const days = (history && typeof history === 'object' && history.days && typeof history.days === 'object')
        ? { ...history.days }
        : {};
    const added = [];
    const skipped = [];
    for (const dateKey of Object.keys(importedDays || {})) {
        if (Object.prototype.hasOwnProperty.call(days, dateKey)) {
            skipped.push(dateKey);
        } else {
            days[dateKey] = importedDays[dateKey];
            added.push(dateKey);
        }
    }
    return { history: { version: 1, days }, added: added.sort(), skipped: skipped.sort() };
}

/**
 * Minimal quote-aware CSV parser (handles quoted fields, escaped quotes, and
 * newlines inside quotes). Returns an array of string rows.
 */
export function parseCsv(text) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    const src = String(text ?? '');
    for (let i = 0; i < src.length; i++) {
        const ch = src[i];
        if (inQuotes) {
            if (ch === '"') {
                if (src[i + 1] === '"') { field += '"'; i++; }
                else inQuotes = false;
            } else {
                field += ch;
            }
        } else if (ch === '"') {
            inQuotes = true;
        } else if (ch === ',') {
            row.push(field);
            field = '';
        } else if (ch === '\n' || ch === '\r') {
            if (ch === '\r' && src[i + 1] === '\n') i++;
            row.push(field);
            field = '';
            rows.push(row);
            row = [];
        } else {
            field += ch;
        }
    }
    if (field !== '' || row.length > 0) {
        row.push(field);
        rows.push(row);
    }
    // Drop fully-empty trailing rows (common with trailing newlines).
    return rows.filter(r => r.some(c => String(c).trim() !== ''));
}

/**
 * Unwrap an ExcelJS cell value (rich text, hyperlink, formula result, Date)
 * into a plain value usable by rowsToImportedDays.
 */
export function excelCellToValue(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
        if (Array.isArray(v.richText)) return v.richText.map(t => t && t.text ? t.text : '').join('');
        if (v.result !== undefined && v.result !== null) return excelCellToValue(v.result);
        if (v.text !== undefined) return excelCellToValue(v.text);
        if (v.hyperlink !== undefined) return String(v.hyperlink);
        return '';
    }
    return v;
}

const DATE_COLUMN_ALIASES = ['date', 'snapshot date', 'day', 'timestamp'];

/**
 * Convert a header + data row matrix (from an uploaded CSV or Excel file) into
 * per-day snapshots, matching the Download History layout: Date, SyStudentID,
 * LDA, Days Out — columns resolved by alias, extra columns ignored. Rows
 * without a parseable date, an id, or a numeric Days Out are counted as
 * skipped. Throws when the required columns can't be found at all.
 * @returns {{ days: Object, rows: number, skippedRows: number }}
 */
export function rowsToImportedDays(rows) {
    if (!Array.isArray(rows) || rows.length < 2) {
        throw new Error('File has no data rows.');
    }
    const headers = rows[0];
    const normalized = headers.map(normalizeHeader);
    const dateIdx = findColumnIndex(normalized, DATE_COLUMN_ALIASES);
    const { idIdx, ldaIdx, daysOutIdx } = resolveSnapshotIndices(headers);
    const missing = [];
    if (dateIdx === -1) missing.push('Date');
    if (idIdx === -1) missing.push('SyStudentID');
    if (daysOutIdx === -1) missing.push('Days Out');
    if (missing.length > 0) {
        throw new Error(`Missing required column${missing.length === 1 ? '' : 's'}: ${missing.join(', ')}. Expected the Download History layout (Date, SyStudentID, LDA, Days Out).`);
    }

    const days = {};
    let imported = 0;
    let skippedRows = 0;
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i] || [];
        const dateKey = normalizeLdaValue(row[dateIdx]);
        const rawId = row[idIdx];
        const id = (rawId === null || rawId === undefined) ? '' : String(rawId).trim();
        const daysOut = Number(row[daysOutIdx]);
        if (!dateKey || !/^\d{4}-\d{2}-\d{2}$/.test(dateKey) || !id || !Number.isFinite(daysOut)) {
            skippedRows++;
            continue;
        }
        (days[dateKey] ||= {})[id] = {
            lda: (ldaIdx !== -1) ? normalizeLdaValue(row[ldaIdx]) : null,
            daysOut,
        };
        imported++;
    }
    if (imported === 0) {
        throw new Error('No usable rows found — check the Date, SyStudentID, and Days Out values.');
    }
    return { days, rows: imported, skippedRows };
}

/**
 * Merge already-parsed day snapshots into the stored history (existing dates
 * always win) and persist when anything was added.
 * @returns {{ added: string[], skipped: string[] }}
 */
export function importDaysIntoHistory(importedDays) {
    const { history, added, skipped } = mergeImportedDays(loadLdaHistory(), importedDays);
    if (added.length > 0) saveLdaHistory(history);
    return { added, skipped };
}

/** Flatten the history into CSV (Date, SyStudentID, LDA, Days Out), sorted by date then id. */
export function historyToCsv(history) {
    const escape = (v) => {
        const str = v === null || v === undefined ? '' : String(v);
        return (/[",\n]/.test(str)) ? `"${str.replace(/"/g, '""')}"` : str;
    };
    const lines = ['Date,SyStudentID,LDA,Days Out'];
    const days = (history && history.days) || {};
    for (const dateKey of Object.keys(days).sort()) {
        const snapshot = days[dateKey];
        if (!snapshot || typeof snapshot !== 'object') continue;
        for (const id of Object.keys(snapshot).sort()) {
            const entry = snapshot[id] || {};
            lines.push([dateKey, id, entry.lda ?? '', entry.daysOut ?? ''].map(escape).join(','));
        }
    }
    return lines.join('\n');
}

/**
 * Backfill LDAHistory from the workbook's previous date-mode LDA sheets.
 * Reads every sheet named like "LDA M-D-YYYY", extracts each student's
 * LDA/Days Out, and merges the results into the stored history — never
 * overwriting dates that already have a snapshot. Non-LDA rows on those
 * sheets (titles, spacer rows, secondary table headers) are skipped by the
 * numeric Days Out requirement.
 * @returns {Promise<{ scanned: number, added: string[], skipped: string[], unreadable: string[] }>}
 */
export async function importHistoryFromLdaSheets() {
    if (typeof Excel === 'undefined' || !Excel.run) {
        throw new Error('Excel API not available');
    }
    return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();

        const targets = sheets.items
            .map(s => ({ name: s.name, dateKey: parseLdaSheetName(s.name) }))
            .filter(t => t.dateKey);

        const importedDays = {};
        const unreadable = [];
        for (const target of targets) {
            const used = sheets.getItem(target.name).getUsedRangeOrNullObject();
            used.load('values');
            await context.sync();
            if (used.isNullObject || !used.values || used.values.length < 2) {
                unreadable.push(target.name);
                continue;
            }
            const values = used.values;
            const indices = resolveSnapshotIndices(values[0]);
            if (indices.idIdx === -1 || indices.daysOutIdx === -1) {
                unreadable.push(target.name);
                continue;
            }
            const snapshot = buildDailySnapshot(values, indices);
            if (Object.keys(snapshot).length === 0) {
                unreadable.push(target.name);
                continue;
            }
            // Same-day duplicates ("LDA 7-21-2026 (2)") are later runs of that
            // day; iterate in sheet order so the last one wins within the import.
            importedDays[target.dateKey] = snapshot;
        }

        const { added, skipped } = importDaysIntoHistory(importedDays);
        return { scanned: targets.length, added, skipped, unreadable };
    });
}

/**
 * Join risk counts with the Master List into a ranked leaderboard.
 * Only students with a count > 0 AND a current Master List row are included —
 * ids missing from the master are treated as dropped and left off. Sorted by
 * count descending, then name.
 * @param {Map<string, number>} riskCounts - from countRiskEpisodes
 * @param {Array<Array>} masterValues - Master List matrix (row 0 = headers)
 * @returns {Array<{ id, name, gender, profilePicture, count }>}
 */
export function buildRiskLeaderboard(riskCounts, masterValues) {
    if (!(riskCounts instanceof Map) || riskCounts.size === 0) return [];
    if (!Array.isArray(masterValues) || masterValues.length < 2) return [];
    const normalized = masterValues[0].map(normalizeHeader);
    const sy = findColumnIndex(normalized, STUDENT_ID_ALIASES);
    const idIdx = sy !== -1 ? sy : findColumnIndex(normalized, STUDENT_NUMBER_ALIASES);
    if (idIdx === -1) return [];
    const nameIdx = findColumnIndex(normalized, STUDENT_NAME_ALIASES);
    const genderIdx = findColumnIndex(normalized, GENDER_ALIASES);
    const pfpIdx = findColumnIndex(normalized, PROFILE_PICTURE_ALIASES);

    const asUrl = (v) => {
        const str = (v === null || v === undefined) ? '' : String(v).trim();
        return (/^https?:\/\//i.test(str) || str.startsWith('data:image')) ? str : null;
    };

    const entries = [];
    const seen = new Set();
    for (let i = 1; i < masterValues.length; i++) {
        const row = masterValues[i];
        if (!row) continue;
        const rawId = row[idIdx];
        const id = (rawId === null || rawId === undefined) ? '' : String(rawId).trim();
        if (!id || seen.has(id)) continue;
        const count = riskCounts.get(id);
        if (!count) continue;
        seen.add(id);
        entries.push({
            id,
            count,
            name: (nameIdx !== -1) ? String(row[nameIdx] ?? '').trim() : '',
            gender: (genderIdx !== -1) ? String(row[genderIdx] ?? '').trim() : '',
            profilePicture: (pfpIdx !== -1) ? asUrl(row[pfpIdx]) : null,
        });
    }
    entries.sort((a, b) => (b.count - a.count) || a.name.localeCompare(b.name));
    return entries;
}

/** First matching count for any of the student's id variants (SyStudentId, Student Number). */
export function lookupRiskCount(riskIndexMap, ids) {
    for (const id of ids) {
        if (riskIndexMap.has(id)) return riskIndexMap.get(id);
    }
    return 0;
}

// --- Document settings IO (no-ops outside an Office host) ---

export function loadLdaHistory() {
    try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            const stored = Office.context.document.settings.get(LDA_HISTORY_KEY);
            if (stored && typeof stored === 'object') return stored;
        }
    } catch (e) {
        console.warn('riskIndex: failed to read LDAHistory setting', e);
    }
    return { version: 1, days: {} };
}

export function saveLdaHistory(history, config) {
    // LDAHistory grows one entry per day and is the most likely key to push the
    // document-settings blob over Excel's size cap. Auto-prune to the configured
    // budget before saving, then route through the diagnostic saver so any
    // remaining failure is logged with its size instead of being swallowed.
    const { history: pruned } = pruneHistory(history, config || getPruneConfig());
    return saveDocumentSetting(LDA_HISTORY_KEY, pruned, 'riskIndex.saveLdaHistory');
}

/**
 * Current auto-prune config from workbook settings (sanitized, with defaults).
 * Stored as flat keys (ldaHistoryPruneEnabled, ldaHistoryMaxSizeKB,
 * ldaHistoryMaxDays) to match the Settings page's schema-driven rows.
 */
export function getPruneConfig() {
    try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            const wb = Office.context.document.settings.get('workbookSettings') || {};
            return sanitizePruneConfig({
                enabled: wb.ldaHistoryPruneEnabled,
                maxSizeKB: wb.ldaHistoryMaxSizeKB,
                maxDays: wb.ldaHistoryMaxDays,
            });
        }
    } catch (e) {
        console.warn('riskIndex: failed to read prune config', e);
    }
    return { ...DEFAULT_PRUNE_CONFIG };
}

/**
 * Apply the current (or supplied) prune config to the saved history right now
 * and persist the result. Used by the "Prune now" button in the config modal.
 * @returns {Promise<{ ok: boolean, removed: string[], dayCount: number }>}
 */
export async function pruneLdaHistoryNow(config) {
    const cfg = config || getPruneConfig();
    const { history: pruned, removed } = pruneHistory(loadLdaHistory(), cfg);
    const ok = await saveDocumentSetting(LDA_HISTORY_KEY, pruned, 'riskIndex.pruneLdaHistoryNow');
    return { ok, removed, dayCount: Object.keys(pruned.days).length };
}

/**
 * Wipe all saved LDA history, resetting the store to an empty (but valid)
 * shape. Every student's Risk Index count derives from this history, so
 * clearing it resets those counts to zero. Also the fastest way to shrink an
 * oversized document-settings blob.
 * @returns {Promise<boolean>} true on a confirmed successful save.
 */
export function clearLdaHistory() {
    return saveDocumentSetting(LDA_HISTORY_KEY, { version: 1, days: {} }, 'riskIndex.clearLdaHistory');
}

/**
 * Current Risk Index config from workbook settings (sanitized, with defaults).
 * Stored as flat keys (riskIndexEnabled, riskIndexThreshold,
 * riskIndexShowColumn) to match the Settings page's schema-driven rows.
 */
export function getRiskIndexConfig() {
    try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            const wb = Office.context.document.settings.get('workbookSettings') || {};
            return sanitizeRiskIndexSettings({
                enabled: wb.riskIndexEnabled,
                threshold: wb.riskIndexThreshold,
                showColumn: wb.riskIndexShowColumn,
            });
        }
    } catch (e) {
        console.warn('riskIndex: failed to read workbookSettings', e);
    }
    return { ...DEFAULT_RISK_INDEX_SETTINGS };
}
