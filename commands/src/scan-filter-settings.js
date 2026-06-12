/*
 * scan-filter-settings.js
 *
 * Answers the Chrome extension's SRK_REQUEST_SCAN_FILTER_SETTINGS message
 * (sent when the user presses Start on the submission checker). Inspects
 * today's LDA sheet and replies with the scan-filter settings it implies:
 * the minimum Days Out on the main table and whether a failing list was
 * generated. When today's LDA sheet doesn't exist the reply says so and
 * the extension makes no change.
 */
import chromeExtensionService from '../../shared/chromeExtensionService.js';
import { pickTodaysLdaSheet, deriveScanFilterFromGrid } from './scan-filter-derivation.js';

/**
 * Reads today's LDA sheet and sends SRK_SCAN_FILTER_SETTINGS back to the
 * Chrome extension. Always replies — ldaFound:false when there is no sheet
 * for today — so the extension can report the outcome instead of waiting
 * for its timeout.
 * @returns {Promise<void>}
 */
export async function sendScanFilterSettingsToExtension() {
    const reply = (data) => {
        chromeExtensionService.sendMessage({
            type: 'SRK_SCAN_FILTER_SETTINGS',
            data: { ...data, timestamp: new Date().toISOString() }
        });
    };

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load('items/name');
            await context.sync();

            const sheetName = pickTodaysLdaSheet(sheets.items.map(s => s.name));
            if (!sheetName) {
                console.log("ScanFilter: No LDA sheet for today — telling extension to keep its settings.");
                reply({ ldaFound: false });
                return;
            }

            const usedRange = sheets.getItem(sheetName).getUsedRange();
            usedRange.load('values');
            await context.sync();

            const derived = deriveScanFilterFromGrid(usedRange.values);
            console.log(`ScanFilter: Derived from "${sheetName}":`, derived);
            reply({ ldaFound: true, sheetName, ...derived });
        });
    } catch (error) {
        console.error('ScanFilter: Error deriving scan filter settings:', error);
        if (typeof OfficeExtension !== 'undefined' && error instanceof OfficeExtension.Error) {
            console.error('ScanFilter: Debug info:', JSON.stringify(error.debugInfo));
        }
        // Report as not found so the extension proceeds with its settings.
        reply({ ldaFound: false, error: error.message });
    }
}
