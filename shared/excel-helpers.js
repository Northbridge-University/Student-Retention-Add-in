/*
 * shared/excel-helpers.js
 *
 * Generic Excel/Office helpers shared by both runtimes.
 */

/**
 * Canonicalizes a column header / alias for matching. NFKC-normalizes
 * Unicode (so fullwidth and ligature characters collapse to their basic
 * forms), lowercases, and strips ALL whitespace (including non-breaking
 * space and other Unicode whitespace).
 *
 *   "Grade Book", "  GRADEBOOK  ", "grade book", "Ｇｒａｄｅ Ｂｏｏｋ"
 *   → all collapse to "gradebook"
 *
 * This means alias lists do NOT need to enumerate case, whitespace, or
 * Unicode-form variants.
 */
export function normalizeHeader(s) {
    const str = String(s ?? '');
    try {
        return str.normalize('NFKC').toLowerCase().replace(/\p{White_Space}+/gu, '');
    } catch (_) {
        // Fallback for runtimes without Unicode property escapes
        return str.toLowerCase().replace(/\s+/g, '');
    }
}

/**
 * Finds the index of a column by checking the headers array against a list
 * of possible alias names. Aliases are normalized via normalizeHeader before
 * matching; callers must pre-normalize headers the same way.
 *
 * @param {string[]} headers - Headers normalized via normalizeHeader.
 * @param {string[]} possibleNames - Aliases to try, in order.
 * @returns {number} Matching column index, or -1 if no alias matches.
 */
export function findColumnIndex(headers, possibleNames) {
    if (!Array.isArray(possibleNames)) {
        console.error("[DEBUG] findColumnIndex received non-array for possibleNames:", possibleNames);
        return -1;
    }
    for (const name of possibleNames) {
        const index = headers.indexOf(normalizeHeader(name));
        if (index !== -1) {
            return index;
        }
    }
    return -1;
}

/**
 * Generates equivalent spellings of a sheet name that ends in a date, so a
 * sheet can be found regardless of zero-padding or separator differences:
 *
 *   sheetNameDateVariants("LDA 6-10-2026")
 *     → ["LDA 6-10-2026", "LDA 06-10-2026", "LDA 06/10/2026", ...]
 *
 * Names that don't end in a M/D/YYYY-style date are returned as a
 * single-element array unchanged. The original name is always first.
 *
 * @param {string} name - Sheet name, possibly with a date suffix.
 * @returns {string[]} Unique name variants to try, original first.
 */
export function sheetNameDateVariants(name) {
    if (!name || typeof name !== 'string') return [name];

    // Extract any prefix (e.g. "LDA " from "LDA 01-11-2026")
    const datePattern = /^(.*?)(\d{1,2}[/-]\d{1,2}[/-]\d{4})$/;
    const match = name.match(datePattern);
    if (!match) return [name];

    const prefix = match[1];
    const datePart = match[2];
    const separator = datePart.includes('/') ? '/' : datePart.includes('-') ? '-' : null;
    if (!separator) return [name];

    const parts = datePart.split(separator);
    if (parts.length !== 3) return [name];
    const [month, day, year] = parts;

    const variations = new Set();
    variations.add(name);
    ['/', '-'].forEach(sep => {
        variations.add(`${prefix}${month.padStart(2, '0')}${sep}${day.padStart(2, '0')}${sep}${year}`);
        variations.add(`${prefix}${parseInt(month, 10)}${sep}${parseInt(day, 10)}${sep}${year}`);
        variations.add(`${prefix}${month.padStart(2, '0')}${sep}${parseInt(day, 10)}${sep}${year}`);
        variations.add(`${prefix}${parseInt(month, 10)}${sep}${day.padStart(2, '0')}${sep}${year}`);
    });
    return Array.from(variations);
}

/**
 * Parses an Excel HYPERLINK formula string and returns its url and display text.
 * Returns null if the string isn't a valid HYPERLINK formula.
 *
 *   parseHyperlinkFormula('=HYPERLINK("https://x.com", "Click")')
 *     → { url: 'https://x.com', text: 'Click' }
 *
 * @param {string} formula
 * @returns {{ url: string, text: string } | null}
 */
export function parseHyperlinkFormula(formula) {
    if (!formula || typeof formula !== 'string') return null;

    const hyperlinkRegex = /=HYPERLINK\("([^"]+)",\s*"([^"]+)"\)/i;
    const match = formula.match(hyperlinkRegex);

    if (match) {
        return {
            url: match[1],
            text: match[2]
        };
    }

    return null;
}
