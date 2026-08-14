// Centralized diagnostics + safe writes for Office document settings.
//
// Every persistent add-in setting is stored in the SAME document-settings
// blob (`Office.context.document.settings`). `saveAsync` serializes and
// writes that ENTIRE blob at once, so if any single key grows past Excel's
// size cap, EVERY save fails with a generic `OSF.DDA.Error` — including
// unrelated writes like the Settings columns. These helpers make that
// failure legible: they surface the real error code/name/message and log a
// per-key size breakdown so we can see which key blew the budget.
//
// All logs are prefixed with `[SRK settings]` so they're easy to filter in
// the browser console.

const LOG_PREFIX = '[SRK settings]';

// Top-level keys this add-in stores under Office.context.document.settings.
// Keep in sync with the writers:
//   'workbookSettings' — columns, Create-LDA toggles, emergencyContacts,
//                        registered users, lastKnownHeaders (Settings.jsx,
//                        LDAManager.jsx, emergencyContacts.js, workbookUsers.js)
//   'LDAHistory'       — Risk Index daily snapshots (riskIndex.js, LDA_HISTORY_KEY)
// (SSO/email data lives in the separate context.workbook.settings store and
// does NOT count toward this blob.)
export const KNOWN_DOC_SETTING_KEYS = ['workbookSettings', 'LDAHistory'];

/** True when the Office document-settings API is usable in this runtime. */
export function isDocumentSettingsAvailable() {
	return (
		typeof window !== 'undefined' &&
		window.Office &&
		Office.context &&
		Office.context.document &&
		Office.context.document.settings
	);
}

/**
 * Pull the useful fields out of an OSF.DDA.Error (or any thrown value).
 * Office error objects stringify to a bare "OSF.DDA.Error", hiding the code
 * and message — this flattens them into something loggable.
 * @returns {{ name: string, code: (number|string|null), message: string }}
 */
export function describeOfficeError(error) {
	if (!error) return { name: 'Unknown', code: null, message: 'No error object provided' };
	// Office AsyncResult.error exposes name/code/message; plain Errors have name/message.
	const name = error.name ?? error.constructor?.name ?? 'Error';
	const code = (error.code !== undefined && error.code !== null) ? error.code : null;
	const message = error.message ?? String(error);
	return { name, code, message };
}

/** UTF-8 byte length of a JSON-serialized value; -1 if it can't be serialized. */
export function byteSize(value) {
	let json;
	try {
		json = JSON.stringify(value);
	} catch {
		return -1;
	}
	if (json === undefined) return 0;
	try {
		if (typeof TextEncoder !== 'undefined') {
			return new TextEncoder().encode(json).length;
		}
	} catch {
		/* fall through to length estimate */
	}
	return json.length;
}

function formatBytes(n) {
	if (n < 0) return 'unserializable';
	if (n < 1024) return `${n} B`;
	if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
	return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

/**
 * Measure the size of each known settings key plus the combined total, and
 * log a breakdown. This is the key diagnostic for `OSF.DDA.Error` on save:
 * whichever key dominates the total is the one pushing the blob over Excel's
 * cap and breaking every write.
 * @param {string} contextLabel - where the measurement was taken from.
 * @returns {{ perKey: Object<string, number>, total: number } | null}
 */
export function measureDocumentSettings(contextLabel = '') {
	if (!isDocumentSettingsAvailable()) return null;
	const perKey = {};
	let total = 0;
	for (const key of KNOWN_DOC_SETTING_KEYS) {
		let size = 0;
		try {
			const val = Office.context.document.settings.get(key);
			size = val === undefined || val === null ? 0 : byteSize(val);
		} catch (e) {
			size = -1;
			console.warn(`${LOG_PREFIX} could not read key "${key}" while measuring`, e);
		}
		perKey[key] = size;
		if (size > 0) total += size;
	}
	const summary = KNOWN_DOC_SETTING_KEYS
		.map(k => `${k}=${formatBytes(perKey[k])}`)
		.join(', ');
	console.info(
		`${LOG_PREFIX} size breakdown${contextLabel ? ` (${contextLabel})` : ''}: ` +
		`total=${formatBytes(total)} | ${summary}`
	);
	return { perKey, total };
}

/**
 * Log the outcome of a settings `saveAsync`. On failure it dumps the full
 * error detail AND a size breakdown so the point of failure is unambiguous.
 * @param {string} label - identifies the call site (e.g. 'Settings.columns').
 * @param {object} result - the Office AsyncResult passed to the callback.
 * @returns {boolean} whether the save succeeded.
 */
export function logSettingsSaveResult(label, result) {
	const succeeded =
		result &&
		typeof Office !== 'undefined' &&
		Office.AsyncResultStatus &&
		result.status === Office.AsyncResultStatus.Succeeded;

	if (succeeded) {
		console.debug(`${LOG_PREFIX} save OK: ${label}`);
		return true;
	}

	const err = describeOfficeError(result && result.error);
	console.error(
		`${LOG_PREFIX} save FAILED: ${label} — ` +
		`name=${err.name} code=${err.code ?? 'n/a'} message="${err.message}"`,
		result && result.error
	);
	// Show which key is oversized — the usual reason a save is rejected.
	measureDocumentSettings(`after failed save: ${label}`);
	return false;
}

/**
 * Set one key and persist the whole settings blob, with full diagnostics on
 * both the synchronous `set` (can throw on non-serializable values) and the
 * async `saveAsync` (fails when the combined blob is too large).
 *
 * Fire-and-forget friendly, but returns a Promise<boolean> for callers that
 * want to await the outcome.
 *
 * @param {string} key   - document settings key to write.
 * @param {*}      value - JSON-serializable value to store.
 * @param {string} label - call-site label for logs.
 * @returns {Promise<boolean>} true on a confirmed successful save.
 */
export function saveDocumentSetting(key, value, label) {
	if (!isDocumentSettingsAvailable()) {
		console.warn(`${LOG_PREFIX} skip save (${label}): Office document settings unavailable`);
		return Promise.resolve(false);
	}

	// Warn early if the value we're about to store is itself huge — that's the
	// key most likely to be pushing the whole blob over the limit.
	const thisSize = byteSize(value);
	if (thisSize < 0) {
		console.error(`${LOG_PREFIX} value for "${key}" (${label}) is not JSON-serializable; save aborted`);
		return Promise.resolve(false);
	}
	if (thisSize > 512 * 1024) {
		console.warn(`${LOG_PREFIX} large value for "${key}" (${label}): ${formatBytes(thisSize)}`);
	}

	return new Promise(resolve => {
		try {
			Office.context.document.settings.set(key, value);
		} catch (err) {
			const info = describeOfficeError(err);
			console.error(
				`${LOG_PREFIX} set() threw for "${key}" (${label}) — ` +
				`name=${info.name} message="${info.message}"`,
				err
			);
			resolve(false);
			return;
		}

		try {
			Office.context.document.settings.saveAsync(result => {
				resolve(logSettingsSaveResult(`${label} [key=${key}]`, result));
			});
		} catch (err) {
			const info = describeOfficeError(err);
			console.error(
				`${LOG_PREFIX} saveAsync() threw for "${key}" (${label}) — ` +
				`name=${info.name} message="${info.message}"`,
				err
			);
			measureDocumentSettings(`after saveAsync throw: ${label}`);
			resolve(false);
		}
	});
}
