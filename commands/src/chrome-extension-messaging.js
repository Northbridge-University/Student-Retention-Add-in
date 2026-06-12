/*
 * chrome-extension-messaging.js
 *
 * Bridges the commands runtime to the Chrome extension via window.postMessage
 * and chromeExtensionService.
 *
 * Exports:
 *   setupPingResponseListener   - replies to SRK_PING heartbeat from the extension
 *   createSheetFromExtension    - creates/activates a sheet on extension request
 *   sendSheetListToExtension    - returns the list of worksheet names to the extension
 *
 * The Power Automate document-property poller lives in ./power-automate-poller.js.
 */
import chromeExtensionService from '../../shared/chromeExtensionService.js';

/**
 * Sets up listener to respond to ping checks from the Chrome extension
 * Call this during add-in initialization to enable connectivity checks
 */
export function setupPingResponseListener() {
    window.addEventListener("message", (event) => {
        // Heartbeat pings from the extension side panel repeat every few
        // seconds for as long as the panel is open; answer them silently
        // so this console stays readable.
        const isHeartbeatPing = !!(event.data && event.data.type === "SRK_PING" && event.data.heartbeat);

        // Log all incoming messages for debugging
        if (event.data && event.data.type && event.data.type.startsWith('SRK_') && !isHeartbeatPing) {
            console.log(`📨 Background Service: Received ${event.data.type} message:`, event.data);
        }

        if (event.data && event.data.type === "SRK_PING") {
            if (!isHeartbeatPing) {
                console.log('🏓 Background Service: Received SRK_PING, sending pong...');
            }

            // Send pong response back to extension
            window.postMessage({
                type: "SRK_PONG",
                timestamp: new Date().toISOString(),
                source: "excel-addin-background"
            }, "*");

            if (!isHeartbeatPing) {
                console.log('✅ Background Service: SRK_PONG sent to Chrome extension');
            }
        }

        if (event.data && event.data.type === "SRK_LINKS") {
            console.log('🔗 Background Service: Received SRK_LINKS message with links:', event.data.links);
            console.log('ℹ️ Background Service: SRK_LINKS is forwarded to Chrome extension (not handled here)');
        }
    });

    console.log('🔔 Background Service: Ping response listener set up');
}

/**
 * Creates a new sheet with the given name and headers
 * @param {string} sheetName - The name of the sheet to create
 * @param {string[]} headers - Array of column headers
 * @returns {Promise<void>}
 */
export async function createSheetFromExtension(sheetName, headers) {
    try {
        console.log(`CreateSheet: Received request to create sheet "${sheetName}" with headers:`, headers);

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Check if sheet already exists
            let targetSheet = sheets.items.find(s => s.name === sheetName);

            if (!targetSheet) {
                console.log(`CreateSheet: Creating new sheet "${sheetName}"...`);
                targetSheet = sheets.add(sheetName);
                await context.sync();
            } else {
                console.log(`CreateSheet: Sheet "${sheetName}" already exists, getting reference...`);
                targetSheet = context.workbook.worksheets.getItem(sheetName);
            }

            // Write headers to the first row
            if (headers && Array.isArray(headers) && headers.length > 0) {
                const headerRange = targetSheet.getRangeByIndexes(0, 0, 1, headers.length);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                await context.sync();
                console.log(`CreateSheet: Headers written to sheet "${sheetName}"`);
            }

            // Activate the new sheet
            targetSheet.activate();
            await context.sync();

            console.log(`CreateSheet: Sheet "${sheetName}" created and activated successfully`);
        });
    } catch (error) {
        console.error(`CreateSheet: Error creating sheet "${sheetName}":`, error);
        if (error instanceof OfficeExtension.Error) {
            console.error("CreateSheet: Debug info:", JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Retrieves all sheet names from the workbook and sends them to the Chrome extension
 * @returns {Promise<void>}
 */
export async function sendSheetListToExtension() {
    try {
        console.log("SheetList: Retrieving all sheet names...");

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Extract sheet names
            const sheetNames = sheets.items.map(sheet => sheet.name);
            console.log(`SheetList: Found ${sheetNames.length} sheets:`, sheetNames);

            // Send response to Chrome extension
            chromeExtensionService.sendMessage({
                type: 'SRK_SHEET_LIST_RESPONSE',
                sheets: sheetNames,
                timestamp: new Date().toISOString()
            });

            console.log("SheetList: Sheet list sent to Chrome extension");
        });
    } catch (error) {
        console.error("SheetList: Error retrieving sheet names:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("SheetList: Debug info:", JSON.stringify(error.debugInfo));
        }

        // Send error response to extension
        chromeExtensionService.sendMessage({
            type: 'SRK_SHEET_LIST_RESPONSE',
            sheets: [],
            error: error.message,
            timestamp: new Date().toISOString()
        });
    }
}
