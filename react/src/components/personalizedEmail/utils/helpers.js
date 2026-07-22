// Utility functions for personalized email feature
import { parseHyperlinkFormula } from '../../../../../shared/excel-helpers.js';
export { findColumnIndex, parseHyperlinkFormula, normalizeHeader } from '../../../../../shared/excel-helpers.js';

export function getTodaysLdaSheetName() {
    const now = new Date();
    return `LDA ${now.getMonth() + 1}-${now.getDate()}-${now.getFullYear()}`;
}

// Stamp for the "Last Email sent ..." footer: "at 1:54 pm" when the timestamp
// falls on today's date, otherwise "on 6/9/26". Empty string if unparsable.
export function formatLastSentStamp(isoTimestamp, now = new Date()) {
    const sent = new Date(isoTimestamp);
    if (isNaN(sent.getTime())) return '';

    const isToday = sent.getFullYear() === now.getFullYear()
        && sent.getMonth() === now.getMonth()
        && sent.getDate() === now.getDate();

    if (isToday) {
        const hours24 = sent.getHours();
        const hour = hours24 % 12 || 12;
        const minutes = String(sent.getMinutes()).padStart(2, '0');
        return `at ${hour}:${minutes} ${hours24 < 12 ? 'am' : 'pm'}`;
    }
    const yearTwoDigit = String(sent.getFullYear() % 100).padStart(2, '0');
    return `on ${sent.getMonth() + 1}/${sent.getDate()}/${yearTwoDigit}`;
}

export function getNameParts(fullName) {
    if (!fullName || typeof fullName !== 'string') {
        return { first: '', last: '' };
    }

    const name = fullName.trim();

    if (name.includes(',')) {
        const parts = name.split(',').map(p => p.trim());
        const lastName = parts[0];
        const firstName = parts[1] || '';
        return { first: firstName, last: lastName };
    } else {
        const parts = name.split(' ').filter(p => p);
        if (parts.length === 1) {
            return { first: parts[0], last: '' };
        }
        const lastName = parts.pop();
        const firstName = parts.join(' ');
        return { first: firstName, last: lastName };
    }
}

export function isValidEmail(email) {
    if (typeof email !== 'string' || !email.trim()) {
        return false;
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

export function isValidHttpUrl(string) {
    try {
        const url = new URL(string);
        return url.protocol === "http:" || url.protocol === "https:";
    } catch (_) {
        return false;
    }
}

export function evaluateMapping(cellValue, mapping) {
    const cellStr = String(cellValue).trim().toLowerCase();
    const conditionStr = String(mapping.if).trim().toLowerCase();
    const cellNum = parseFloat(cellValue);
    const conditionNum = parseFloat(mapping.if);
    const isNumeric = !isNaN(cellNum) && !isNaN(conditionNum);

    switch (mapping.operator) {
        case 'eq': return cellStr === conditionStr;
        case 'neq': return cellStr !== conditionStr;
        case 'contains': return cellStr.includes(conditionStr);
        case 'does_not_contain': return !cellStr.includes(conditionStr);
        case 'starts_with': return cellStr.startsWith(conditionStr);
        case 'ends_with': return cellStr.endsWith(conditionStr);
        case 'gt': return isNumeric && cellNum > conditionNum;
        case 'lt': return isNumeric && cellNum < conditionNum;
        case 'gte': return isNumeric && cellNum >= conditionNum;
        case 'lte': return isNumeric && cellNum <= conditionNum;
        default: return false;
    }
}

export const renderTemplate = (template, data) => {
    if (!template) return '';
    let result = template;
    for (let i = 0; i < 10 && /\{(\w+)\}/.test(result); i++) {
        result = result.replace(/\{(\w+)\}/g, (match, key) => {
            let valueToInsert = data.hasOwnProperty(key) ? data[key] : match;
            if (typeof valueToInsert === 'string') {
                const trimmedValue = valueToInsert.trim();
                if (trimmedValue.startsWith('<p>') && trimmedValue.endsWith('</p>')) {
                    const innerHtml = trimmedValue.substring(3, trimmedValue.length - 4);
                    if (!innerHtml.includes('<p>') && !innerHtml.includes('<div>')) {
                        valueToInsert = innerHtml;
                    }
                }
            }
            return valueToInsert;
        });
    }
    return result;
};

export const renderCCTemplate = (recipients, data) => {
    if (!recipients || recipients.length === 0) return '';
    return recipients.map(recipient => renderTemplate(recipient, data)).join(';');
};

// Normalizes a From value to a comparable key. Lowercases/trims and, when the
// value is in "Name <email>" form, uses just the email so signatures match
// regardless of display name.
export const normalizeEmailKey = (value) => {
    if (!value) return '';
    const str = String(value).trim().toLowerCase();
    const match = str.match(/<([^>]+)>/);
    return (match ? match[1] : str).trim();
};

// Looks up the signature HTML assigned to a given From address.
// `signatures` is an array of { email, signature }. Returns '' if none.
export const findSignature = (signatures, fromValue) => {
    if (!Array.isArray(signatures) || signatures.length === 0) return '';
    const key = normalizeEmailKey(fromValue);
    if (!key) return '';
    const match = signatures.find(s => normalizeEmailKey(s.email) === key);
    return match ? (match.signature || '') : '';
};

/**
 * Pre-loads the "Missing Assignments" sheet and builds a lookup map
 * of Grade Book URL → HTML assignment list.  Call once before the
 * student loop, then use the returned Map for instant per-student lookups.
 *
 * @param {Excel.RequestContext} context
 * @returns {Promise<Map<string, string>>}  gradeBookUrl → "<ul>…</ul>"
 */
export async function buildMissingAssignmentsCache(context) {
    const cache = new Map();
    try {
        const missingSheet = context.workbook.worksheets.getItemOrNullObject("Missing Assignments");
        await context.sync();
        if (missingSheet.isNullObject) return cache;

        const usedRange = missingSheet.getUsedRangeOrNullObject();
        usedRange.load("values, formulas, isNullObject");
        await context.sync();
        if (usedRange.isNullObject) return cache;

        const values = usedRange.values;
        const formulas = usedRange.formulas;
        const headers = values[0].map(h => String(h ?? '').toLowerCase());

        const gradeBookColIndex = headers.findIndex(h =>
            h.includes('grade book') || h.includes('gradebook')
        );
        const assignmentColIndex = headers.findIndex(h =>
            h.includes('assignment')
        );

        if (gradeBookColIndex === -1 || assignmentColIndex === -1) return cache;

        // Group assignments by Grade Book URL
        const grouped = new Map(); // url → [{url, title}]
        for (let i = 1; i < values.length; i++) {
            const rowGradeBookFormula = formulas[i][gradeBookColIndex];
            let rowGradeBookUrl = values[i][gradeBookColIndex];
            if (rowGradeBookFormula) {
                const parsed = parseHyperlinkFormula(rowGradeBookFormula);
                if (parsed) rowGradeBookUrl = parsed.url;
            }
            const key = String(rowGradeBookUrl ?? '').trim();
            if (!key) continue;

            const assignmentFormula = formulas[i][assignmentColIndex];
            const assignmentValue = values[i][assignmentColIndex];
            const parsed = parseHyperlinkFormula(assignmentFormula);
            let entry = null;
            if (parsed) {
                entry = { url: parsed.url, title: parsed.text };
            } else if (assignmentValue) {
                entry = { url: null, title: String(assignmentValue) };
            }
            if (entry) {
                if (!grouped.has(key)) grouped.set(key, []);
                grouped.get(key).push(entry);
            }
        }

        // Convert each group into its HTML string
        for (const [url, assignments] of grouped) {
            const listItems = assignments.map(a =>
                a.url
                    ? `<li><a href="${a.url}" target="_blank">${a.title}</a></li>`
                    : `<li>${a.title}</li>`
            ).join('');
            cache.set(url, `<ul>${listItems}</ul>`);
        }
    } catch (error) {
        console.error('Error building missing assignments cache:', error);
    }
    return cache;
}

/**
 * Generates an HTML list of missing assignments for a student
 * Returns empty string if no assignments found
 */
export async function generateMissingAssignmentsList(gradeBookValue, gradeBookFormula, context) {
    try {
        // Extract the Grade Book URL from the formula (this is the unique identifier)
        let gradeBookUrl = null;
        if (gradeBookFormula) {
            const parsed = parseHyperlinkFormula(gradeBookFormula);
            if (parsed) {
                gradeBookUrl = parsed.url; // Use the URL as the unique identifier
            }
        }

        // Fallback to value if no formula found
        if (!gradeBookUrl) {
            gradeBookUrl = gradeBookValue;
        }

        if (!gradeBookUrl) return '';

        // Access the Missing Assignments sheet
        const missingSheet = context.workbook.worksheets.getItem("Missing Assignments");
        const usedRange = missingSheet.getUsedRange();
        usedRange.load("values, formulas");
        await context.sync();

        const values = usedRange.values;
        const formulas = usedRange.formulas;
        const headers = values[0].map(h => String(h ?? '').toLowerCase());

        // Find column indices
        const gradeBookColIndex = headers.findIndex(h =>
            h.includes('grade book') || h.includes('gradebook')
        );
        const assignmentColIndex = headers.findIndex(h =>
            h.includes('assignment')
        );

        if (gradeBookColIndex === -1 || assignmentColIndex === -1) {
            return '';
        }

        // Find all matching assignments
        const assignments = [];
        for (let i = 1; i < values.length; i++) {
            const rowGradeBookValue = values[i][gradeBookColIndex];
            const rowGradeBookFormula = formulas[i][gradeBookColIndex];

            // Extract the URL if it's a HYPERLINK formula
            let rowGradeBookUrl = rowGradeBookValue;
            if (rowGradeBookFormula) {
                const parsed = parseHyperlinkFormula(rowGradeBookFormula);
                if (parsed) {
                    rowGradeBookUrl = parsed.url; // Use URL for comparison
                }
            }

            // Check if this row matches the student's grade book URL
            if (String(rowGradeBookUrl).trim() === String(gradeBookUrl).trim()) {
                const assignmentFormula = formulas[i][assignmentColIndex];
                const assignmentValue = values[i][assignmentColIndex];

                // Parse the assignment hyperlink
                const parsed = parseHyperlinkFormula(assignmentFormula);
                if (parsed) {
                    assignments.push({
                        url: parsed.url,
                        title: parsed.text
                    });
                } else if (assignmentValue) {
                    // Fallback if no hyperlink
                    assignments.push({
                        url: null,
                        title: String(assignmentValue)
                    });
                }
            }
        }

        if (assignments.length === 0) {
            return '';
        }

        // Generate HTML bullet list
        const listItems = assignments.map(assignment => {
            if (assignment.url) {
                return `<li><a href="${assignment.url}" target="_blank">${assignment.title}</a></li>`;
            } else {
                return `<li>${assignment.title}</li>`;
            }
        }).join('');

        return `<ul>${listItems}</ul>`;

    } catch (error) {
        console.error('Error generating missing assignments list:', error);
        return '';
    }
}
