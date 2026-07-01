/*
 * emergencyContacts.js
 *
 * Stores per-student emergency contacts in the workbook's document settings.
 *
 * Contacts live under the existing `workbookSettings` document setting in an
 * `emergencyContacts` map keyed by the student's SyStudentId (the canonical
 * `ID`). Each entry is an array of contacts:
 *
 *   workbookSettings.emergencyContacts = {
 *     "<SyStudentId>": [
 *       { id, name, number, relationship, dateAdded }
 *     ]
 *   }
 *
 * Keying by SyStudentId means the same contacts load automatically whenever
 * that student is viewed, on any machine that opens the workbook.
 */

const DOC_KEY = 'workbookSettings';
const CONTACTS_KEY = 'emergencyContacts';

function isOfficeReady() {
    return typeof window !== 'undefined'
        && window.Office
        && Office.context
        && Office.context.document
        && Office.context.document.settings;
}

function readWorkbookSettings() {
    if (!isOfficeReady()) return null;
    try {
        return Office.context.document.settings.get(DOC_KEY) || {};
    } catch (err) {
        console.warn('emergencyContacts: failed to read document settings', err);
        return null;
    }
}

function writeWorkbookSettings(mapping) {
    if (!isOfficeReady()) return Promise.resolve(false);
    return new Promise(resolve => {
        try {
            Office.context.document.settings.set(DOC_KEY, mapping);
            Office.context.document.settings.saveAsync(result => {
                const ok = result && result.status === Office.AsyncResultStatus.Succeeded;
                if (!ok && result && result.error) {
                    console.warn('emergencyContacts: saveAsync failed', result.error);
                }
                resolve(!!ok);
            });
        } catch (err) {
            console.warn('emergencyContacts: failed to set document settings', err);
            resolve(false);
        }
    });
}

// Normalize an identifier to a non-empty string key, or null when unusable.
function toKey(studentId) {
    if (studentId === null || studentId === undefined) return null;
    const key = String(studentId).trim();
    return key ? key : null;
}

function makeId() {
    return `ec_${Date.now()}_${Math.floor(Math.random() * 1e6)}`;
}

/**
 * Returns the emergency contacts saved for a student (always an array).
 * @param {string|number} studentId - The student's SyStudentId (canonical ID).
 * @returns {Array<{id: string, name: string, number: string, relationship: string, dateAdded: string}>}
 */
export function getEmergencyContacts(studentId) {
    const key = toKey(studentId);
    if (!key) return [];
    const settings = readWorkbookSettings();
    if (!settings) return [];
    const map = settings[CONTACTS_KEY];
    if (!map || typeof map !== 'object') return [];
    const list = map[key];
    return Array.isArray(list) ? list : [];
}

/**
 * Adds an emergency contact for a student and persists it.
 * @param {string|number} studentId - The student's SyStudentId.
 * @param {{name: string, number: string, relationship?: string}} contact
 * @returns {Promise<Array|null>} The updated contacts array, or null on failure.
 */
export async function addEmergencyContact(studentId, contact) {
    const key = toKey(studentId);
    if (!key) {
        console.warn('emergencyContacts: cannot add — missing student identifier');
        return null;
    }
    const name = (contact && typeof contact.name === 'string') ? contact.name.trim() : '';
    const number = (contact && typeof contact.number === 'string') ? contact.number.trim() : '';
    const relationship = (contact && typeof contact.relationship === 'string') ? contact.relationship.trim() : '';
    if (!name || !number) return null;

    const settings = readWorkbookSettings();
    if (!settings) return null;

    const map = (settings[CONTACTS_KEY] && typeof settings[CONTACTS_KEY] === 'object')
        ? settings[CONTACTS_KEY]
        : {};
    const existing = Array.isArray(map[key]) ? map[key] : [];

    const record = { id: makeId(), name, number, relationship, dateAdded: new Date().toISOString() };
    const nextList = [...existing, record];
    const next = { ...settings, [CONTACTS_KEY]: { ...map, [key]: nextList } };

    const ok = await writeWorkbookSettings(next);
    return ok ? nextList : null;
}

/**
 * Removes an emergency contact (by id) for a student and persists the change.
 * @param {string|number} studentId - The student's SyStudentId.
 * @param {string} contactId - The id of the contact to remove.
 * @returns {Promise<Array|null>} The updated contacts array, or null on failure.
 */
export async function removeEmergencyContact(studentId, contactId) {
    const key = toKey(studentId);
    if (!key || !contactId) return null;

    const settings = readWorkbookSettings();
    if (!settings) return null;

    const map = (settings[CONTACTS_KEY] && typeof settings[CONTACTS_KEY] === 'object')
        ? settings[CONTACTS_KEY]
        : {};
    const existing = Array.isArray(map[key]) ? map[key] : [];
    const nextList = existing.filter(c => c && c.id !== contactId);
    if (nextList.length === existing.length) return existing; // nothing removed

    const nextMap = { ...map };
    if (nextList.length === 0) {
        delete nextMap[key]; // don't leave empty arrays lying around
    } else {
        nextMap[key] = nextList;
    }
    const next = { ...settings, [CONTACTS_KEY]: nextMap };

    const ok = await writeWorkbookSettings(next);
    return ok ? nextList : null;
}
