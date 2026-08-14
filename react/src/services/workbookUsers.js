/*
 * workbookUsers.js
 *
 * Tracks which Student Retention Kit users have opened this workbook.
 *
 * On SSO login (or auto-login from a cached SSO_USER), App.jsx calls
 * `registerWorkbookUser({ name, email })` once per session. The user is
 * upserted into the workbook's `users` array under the existing
 * `workbookSettings` document setting, with a `dateJoined` ISO timestamp
 * on first sight.
 *
 * Dedupe priority is email (case-insensitive), falling back to name.
 * Existing entries preserve their original dateJoined; if a newer login
 * provides an email that was previously missing, the record is enriched
 * in place. The list is surfaced in Settings → Workbook → Users and can
 * be used for future automation (e.g. routing, distribution lists).
 */

import { saveDocumentSetting } from '../components/utility/documentSettings.js';

const DOC_KEY = 'workbookSettings';
const USERS_KEY = 'users';

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
        console.warn('workbookUsers: failed to read document settings', err);
        return null;
    }
}

function writeWorkbookSettings(mapping) {
    if (!isOfficeReady()) return Promise.resolve(false);
    // Diagnostic saver: on failure logs the real error code/message and a
    // per-key size breakdown (see documentSettings.js).
    return saveDocumentSetting(DOC_KEY, mapping, 'workbookUsers.writeWorkbookSettings');
}

/**
 * Returns the list of registered users (always an array; empty if missing).
 * @returns {Array<{name: string, email?: string, dateJoined: string}>}
 */
export function getWorkbookUsers() {
    const settings = readWorkbookSettings();
    if (!settings) return [];
    const users = settings[USERS_KEY];
    return Array.isArray(users) ? users : [];
}

/**
 * Upserts the given user into the workbook's users list. If the user is
 * already registered, the existing record is preserved (original dateJoined
 * stays intact); if the new info adds an email that was previously missing,
 * the record is enriched in place.
 *
 * Safe to call repeatedly — intended to run once per session in the background.
 * @param {{name: string, email?: string} | string} userInfo - User to register.
 *   Pass a string for the legacy "name only" form.
 * @returns {Promise<boolean>} True if any change was written.
 */
export async function registerWorkbookUser(userInfo) {
    const input = typeof userInfo === 'string' ? { name: userInfo } : (userInfo || {});
    const name = typeof input.name === 'string' ? input.name.trim() : '';
    const email = typeof input.email === 'string' ? input.email.trim() : '';
    if (!name) return false;
    if (!isOfficeReady()) return false;

    const settings = readWorkbookSettings();
    if (!settings) return false;

    const users = Array.isArray(settings[USERS_KEY]) ? settings[USERS_KEY] : [];

    // Email is the more reliable identifier (names can collide / change), so
    // try it first and fall back to a name match for legacy entries.
    const lcEmail = email.toLowerCase();
    const lcName = name.toLowerCase();
    let matchIdx = -1;
    if (lcEmail) {
        matchIdx = users.findIndex(u => u && typeof u.email === 'string' && u.email.toLowerCase() === lcEmail);
    }
    if (matchIdx === -1) {
        matchIdx = users.findIndex(u => u && typeof u.name === 'string' && u.name.toLowerCase() === lcName);
    }

    let nextUsers;
    if (matchIdx === -1) {
        const record = { name, dateJoined: new Date().toISOString() };
        if (email) record.email = email;
        nextUsers = [...users, record];
    } else {
        const existing = users[matchIdx];
        const merged = { ...existing };
        // Back-fill missing fields without overwriting set values.
        if (!merged.name && name) merged.name = name;
        if (!merged.email && email) merged.email = email;
        if (merged.name === existing.name && merged.email === existing.email) {
            return false; // already up-to-date — skip the saveAsync round-trip
        }
        nextUsers = users.map((u, i) => (i === matchIdx ? merged : u));
    }

    const next = { ...settings, [USERS_KEY]: nextUsers };
    const ok = await writeWorkbookSettings(next);
    if (ok) {
        const action = matchIdx === -1 ? 'registered' : 'enriched';
        console.log(`workbookUsers: ${action} "${name}"${email ? ` <${email}>` : ''} in workbook settings`);
    }
    return ok;
}
