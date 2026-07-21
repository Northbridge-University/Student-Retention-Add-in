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

/* global Office */

export const LDA_HISTORY_KEY = 'LDAHistory';

export const DEFAULT_RISK_INDEX_SETTINGS = {
    enabled: true,
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

export function saveLdaHistory(history) {
    try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            Office.context.document.settings.set(LDA_HISTORY_KEY, history);
            Office.context.document.settings.saveAsync(() => {});
        }
    } catch (e) {
        console.warn('riskIndex: failed to save LDAHistory setting', e);
    }
}

/** Current Risk Index config from workbook settings (sanitized, with defaults). */
export function getRiskIndexConfig() {
    try {
        if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.settings) {
            const wb = Office.context.document.settings.get('workbookSettings');
            return sanitizeRiskIndexSettings(wb && wb.riskIndex);
        }
    } catch (e) {
        console.warn('riskIndex: failed to read workbookSettings', e);
    }
    return { ...DEFAULT_RISK_INDEX_SETTINGS };
}
