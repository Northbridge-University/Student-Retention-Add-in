/*
 * background-service.js
 *
 * Entry point for the Office Add-in commands runtime (loaded by commands.html
 * into a hidden long-lived runtime, lifetime="long" in manifest.xml).
 *
 * Responsibilities:
 *   - Register ribbon button handlers and the document-open lifecycle hook
 *   - On Office.onReady, start the Chrome Extension bridge and dispatch
 *     incoming extension messages to the appropriate module
 *
 * Module-level concerns live under ./src/:
 *   - master-list-import          (Master List ingestion from extension)
 *   - master-list-transfer        (sending Master List data back to extension)
 *   - conditional-formatting      (color scales, highlights for Master List)
 *   - chrome-extension-messaging  (ping/sheet/highlight bridge + PA poller)
 */
import { toggleHighlight, createSendToCallQueue } from './src/ribbon-actions.js';
import chromeExtensionService from '../shared/chromeExtensionService.js';
import { importMasterListFromExtension } from './src/master-list-import.js';
import { transferMasterList } from './src/master-list-transfer.js';
import {
    setupPingResponseListener,
    createSheetFromExtension,
    sendSheetListToExtension,
} from './src/chrome-extension-messaging.js';
import { startCommandPropertyPoller } from './src/power-automate-poller.js';

// Required for the Analytics ribbon button — manifest uses ShowTaskpane,
// but Office still expects an associated function.
function openAnalyticsPane(event) {
    event.completed();
}

// Called automatically when a new document is opened
function onDocumentOpen(event) {
    console.log("Document opened - background services are active");
    event.completed();
}

// Register ribbon button commands
Office.actions.associate("toggleHighlight", toggleHighlight);
Office.actions.associate("openAnalyticsPane", openAnalyticsPane);
Office.actions.associate("sendToCallQueue", createSendToCallQueue(chromeExtensionService));

// Register autoload event handler
Office.actions.associate("onDocumentOpen", onDocumentOpen);

// Initialize background services when Office is ready
Office.onReady(() => {
    console.log("Background script initialized - starting Chrome Extension Service");

    // Set up ping/pong listener for Chrome extension connectivity checks
    setupPingResponseListener();

    // Start extension detection (continuous pinging)
    chromeExtensionService.startPinging();

    // Keep-alive heartbeat to prevent extension from going dormant
    chromeExtensionService.startKeepAlive();

    // Poll custom document property for highlight commands from Power Automate
    startCommandPropertyPoller();

    // Dispatch incoming extension events
    chromeExtensionService.addListener((event) => {
        // Side panel heartbeat pings arrive every few seconds and are
        // answered by setupPingResponseListener; nothing below handles
        // them, so skip the per-event logging entirely.
        if (event.type === "message" && event.data?.type === "SRK_PING" && event.data.heartbeat) {
            return;
        }
        console.log("Background: Chrome Extension event:", event);

        if (event.type === "installed") {
            console.log("Background: Chrome extension detected and ready!");
            // When extension is detected, ask if it wants the Master List
            console.log("Background: Asking extension if it wants Master List data...");
            chromeExtensionService.sendMessage({
                type: 'SRK_REQUEST_MASTER_LIST',
                timestamp: new Date().toISOString()
            });
            return;
        }

        if (event.type !== "message") return;

        console.log("Background: Received message from extension:", event.data);

        switch (event.data.type) {
            case 'SRK_MASTER_LIST_RESPONSE':
                if (event.data.wantsData === true) {
                    console.log("Background: Extension wants Master List data, sending...");
                    transferMasterList();
                } else {
                    console.log("Background: Extension declined Master List data");
                }
                break;

            case 'SRK_IMPORT_MASTER_LIST':
                console.log("Background: Received Master List import request from extension");
                importMasterListFromExtension(event.data.data);
                break;

            case 'SRK_CREATE_SHEET': {
                console.log("Background: Received sheet creation request from extension");
                const { sheetName, headers } = event.data;
                if (sheetName && headers) {
                    createSheetFromExtension(sheetName, headers);
                } else {
                    console.error("Background: Invalid SRK_CREATE_SHEET request - missing sheetName or headers");
                }
                break;
            }

            case 'SRK_REQUEST_SHEET_LIST':
                console.log("Background: Received sheet list request from extension");
                sendSheetListToExtension();
                break;
        }
    });

    console.log("Chrome Extension Service is now running in the background");
});
