/*
 * master-list-import.js
 *
 * Imports master list data received from the Chrome extension into the
 * Master List sheet. Handles the full pipeline:
 *   - read existing data, build preservation maps for unmatched columns
 *   - merge incoming rows with existing students (preserve gradebook links,
 *     assigned values, cell colors)
 *   - write data in batches (Excel ~5MB payload limit)
 *   - apply conditional formatting and date conversions
 *   - import any nested missing-assignments payload to the side sheet
 *
 * Status updates are posted back to the Chrome extension via window.postMessage
 * so the extension UI can show progress.
 */
import {
    CONSTANTS,
    normalizeName,
    formatToLastFirst,
    parseDate,
    normalizeAttendanceValue,
} from './constants.js';
import { findColumnIndex, normalizeHeader } from '../../shared/excel-helpers.js';
import { BATCH_SIZE } from '../../shared/constants.js';
import {
    applyGradeConditionalFormatting,
    applyLastCourseGradeConditionalFormatting,
    applyMissingAssignmentsConditionalFormatting,
    applyHoldConditionalFormatting,
    applyAdSAPStatusConditionalFormatting,
    applyNextAssignmentDueFormatting,
    applyAttendanceConditionalFormatting,
    applyLetterGradeConditionalFormatting,
    applyEnrollGpaConditionalFormatting,
    applyCourseDateFormatting,
} from './conditional-formatting.js';

// Smaller batch size for formatting operations (cell colors).
const FORMAT_BATCH_SIZE = 50;

/**
 * Sends an import status update back to the Chrome extension via window.postMessage.
 * The extension's content script (excelConnector.js) picks this up and forwards it.
 * @param {string} status - Status key (e.g., 'received', 'writing_data', 'complete', 'error')
 * @param {string} message - Human-readable status message
 * @param {number} [progress] - Optional progress percentage (0-100)
 * @param {string} [error] - Optional error message (for error status)
 */
function sendImportStatus(status, message, progress, error) {
    try {
        window.postMessage({
            type: "SRK_IMPORT_STATUS",
            data: {
                status: status,
                message: message,
                ...(progress != null && { progress }),
                ...(error && { error })
            }
        }, "*");
    } catch (e) {
        console.warn("Failed to send import status:", e);
    }
}

/**
 * Imports missing assignments data to the Missing Assignments sheet
 * @param {Array} studentsWithAssignments - Array of student objects that have missingAssignments
 * @returns {Promise<void>}
 */
async function importMissingAssignments(studentsWithAssignments) {
    try {
        console.log("ImportMissingAssignments: Starting import of missing assignments...");

        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // Check if Missing Assignments sheet exists, create if not
            let maSheet = sheets.items.find(s => s.name === "Missing Assignments");
            if (!maSheet) {
                console.log("ImportMissingAssignments: Creating Missing Assignments sheet...");
                maSheet = sheets.add("Missing Assignments");
                await context.sync();
            } else {
                maSheet = context.workbook.worksheets.getItem("Missing Assignments");
            }

            // Define the column structure
            const headers = ["Student", "Grade", "Grade Book", "Assignment", "Due Date", "Score", "Submission"];

            // Clear existing content
            const usedRange = maSheet.getUsedRangeOrNullObject();
            usedRange.load("address");
            await context.sync();

            if (!usedRange.isNullObject) {
                usedRange.clear(Excel.ClearApplyTo.all);
            }

            // Write headers
            const headerRange = maSheet.getRangeByIndexes(0, 0, 1, headers.length);
            headerRange.values = [headers];
            headerRange.format.font.bold = true;
            await context.sync();

            // Prepare data rows
            const dataToWrite = [];
            const formulasToWrite = [];

            studentsWithAssignments.forEach(student => {
                const studentName = student["Student Name"] || student.StudentName || "";
                const grade = student.Grade ?? "";
                const gradeBookUrl = typeof student["Grade Book"] === 'object' ? student["Grade Book"].url : student["Grade Book"];
                const gradeBookText = typeof student["Grade Book"] === 'object' ? student["Grade Book"].text : "Grade Book";

                if (student.missingAssignments && Array.isArray(student.missingAssignments)) {
                    student.missingAssignments.forEach(assignment => {
                        const row = new Array(headers.length).fill("");
                        const formulaRow = new Array(headers.length).fill(null);

                        // Student
                        row[0] = studentName;

                        // Grade
                        row[1] = grade;

                        // Grade Book (HYPERLINK)
                        if (gradeBookUrl) {
                            formulaRow[2] = `=HYPERLINK("${gradeBookUrl}", "${gradeBookText}")`;
                            row[2] = gradeBookText;
                        }

                        // Assignment (HYPERLINK with assignment URL and title as friendly name)
                        const assignmentUrl = typeof assignment.Assignment === 'object' ? assignment.Assignment.url : assignment.assignmentLink || assignment.assignmentUrl;
                        const assignmentTitle = typeof assignment.Assignment === 'object' ? assignment.Assignment.text : assignment.assignmentTitle || assignment.assignmentName || "Assignment";
                        if (assignmentUrl) {
                            formulaRow[3] = `=HYPERLINK("${assignmentUrl}", "${assignmentTitle}")`;
                            row[3] = assignmentTitle;
                        } else {
                            row[3] = assignmentTitle;
                        }

                        // Due Date
                        row[4] = typeof assignment["Due Date"] === 'object' ? assignment["Due Date"].text : assignment.dueDate || assignment["Due Date"] || "";

                        // Score
                        row[5] = typeof assignment.Score === 'object' ? assignment.Score.text : assignment.score ?? assignment.Score ?? "";

                        // Submission (HYPERLINK with submission URL and "Missing" as friendly name)
                        const submissionUrl = typeof assignment.Submission === 'object' ? assignment.Submission.url : assignment.submissionLink || assignment.submissionUrl;
                        if (submissionUrl) {
                            formulaRow[6] = `=HYPERLINK("${submissionUrl}", "Missing")`;
                            row[6] = "Missing";
                        }

                        dataToWrite.push(row);
                        formulasToWrite.push(formulaRow);
                    });
                }
            });

            if (dataToWrite.length > 0) {
                const totalBatches = Math.ceil(dataToWrite.length / BATCH_SIZE);
                console.log(`ImportMissingAssignments: Writing ${dataToWrite.length} assignment rows in ${totalBatches} batch(es)...`);

                for (let batchStart = 0; batchStart < dataToWrite.length; batchStart += BATCH_SIZE) {
                    const batchEnd = Math.min(batchStart + BATCH_SIZE, dataToWrite.length);
                    const batchSize = batchEnd - batchStart;

                    const batchValues = dataToWrite.slice(batchStart, batchEnd);
                    const batchFormulas = formulasToWrite.slice(batchStart, batchEnd);

                    const batchRange = maSheet.getRangeByIndexes(1 + batchStart, 0, batchSize, headers.length);
                    batchRange.values = batchValues;
                    batchRange.formulas = batchFormulas;

                    await context.sync();

                    const currentBatch = Math.floor(batchStart / BATCH_SIZE) + 1;
                    console.log(`ImportMissingAssignments: Batch ${currentBatch}/${totalBatches} completed (rows ${batchStart + 1}-${batchEnd})`);
                }

                // Autofit columns after all data is written
                maSheet.getUsedRange().format.autofitColumns();
                await context.sync();

                console.log(`ImportMissingAssignments: Successfully imported ${dataToWrite.length} missing assignments`);
            } else {
                console.log("ImportMissingAssignments: No missing assignments to import");
            }
        });

    } catch (error) {
        console.error("ImportMissingAssignments: Error importing missing assignments:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("ImportMissingAssignments: Debug info:", JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Shows a dialog prompting the user to create a Master List sheet when one doesn't exist.
 * If the user confirms, creates the sheet and retries the import.
 * @param {object} payload - The original import payload to retry after sheet creation.
 */
function showMissingMasterListDialog(payload) {
    // Resolve relative to the runtime page (commands.html) so the dialog works
    // on both prod (Northbridge-University.github.io) and staging (Vercel) without needing a
    // host-specific URL — both AppDomains lists allow same-origin URLs.
    const dialogUrl = new URL('missing-masterlist-dialog.html', window.location.href).href;

    sendImportStatus('awaiting_confirmation', 'Master List sheet not found — waiting for your confirmation in Excel...');

    Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 40, width: 35, displayInIframe: true },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("MissingMasterList: Dialog failed to open:", asyncResult.error.message);
                sendImportStatus('error', 'Could not show confirmation dialog', 0, asyncResult.error.message);
                return;
            }
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
                const message = JSON.parse(arg.message);
                dialog.close();
                if (message.type === 'createMasterList') {
                    console.log("MissingMasterList: User chose to create Master List sheet.");
                    await createMasterListAndImport(payload);
                } else {
                    console.log("MissingMasterList: User cancelled. Import aborted.");
                    sendImportStatus('cancelled', 'Import cancelled — no Master List sheet was created');
                }
            });
        }
    );
}

/**
 * Creates a blank Master List sheet and then retries the import.
 * @param {object} payload - The original import payload.
 */
async function createMasterListAndImport(payload) {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.add(CONSTANTS.MASTER_LIST_SHEET);
            // Write the incoming headers as the first row so the import has something to work with
            if (payload.headers && payload.headers.length > 0) {
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, payload.headers.length);
                headerRange.values = [payload.headers];
            }
            sheet.activate();
            await context.sync();
            console.log("MissingMasterList: Master List sheet created successfully.");
        });

        // Retry the import now that the sheet exists
        await importMasterListFromExtension(payload);
    } catch (error) {
        console.error("MissingMasterList: Error creating Master List sheet:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("MissingMasterList: Debug info:", JSON.stringify(error.debugInfo));
        }
    }
}

/**
 * Imports master list data received from the Chrome extension
 * @param {object} payload - The data payload from the extension
 * @param {string[]} payload.headers - Array of column headers
 * @param {Array[]} payload.data - Array of data rows
 * @returns {Promise<void>}
 */
export async function importMasterListFromExtension(payload) {
    try {
        console.log("ImportFromExtension: Starting master list import from Chrome extension...");

        // Validate payload
        if (!payload || !payload.headers || !payload.data) {
            console.error("ImportFromExtension: Invalid payload - missing headers or data");
            return;
        }

        const { headers: incomingHeaders, data: incomingData } = payload;
        console.log(`ImportFromExtension: Received ${incomingData.length} rows with headers: [${incomingHeaders.join(', ')}]`);

        sendImportStatus('received', `Received ${incomingData.length} students, validating data...`);

        // Set when the user must first confirm sheet creation. The dialog drives
        // its own status updates and triggers a re-entrant import on confirm, so
        // the rest of this run must short-circuit instead of falsely reporting
        // 'complete' while the dialog is still open.
        let awaitingDialogConfirmation = false;

        await Excel.run(async (context) => {
            // Check if Master List sheet exists
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            const masterListSheet = sheets.items.find(s => s.name === CONSTANTS.MASTER_LIST_SHEET);
            if (!masterListSheet) {
                console.log("ImportFromExtension: Master List sheet not found. Prompting user to create one.");
                awaitingDialogConfirmation = true;
                showMissingMasterListDialog(payload);
                return;
            }

            console.log("ImportFromExtension: Master List sheet found, proceeding with import...");
            sendImportStatus('reading_data', 'Reading existing spreadsheet data...');

            // Get the Master List sheet
            const sheet = context.workbook.worksheets.getItem(CONSTANTS.MASTER_LIST_SHEET);
            const usedRange = sheet.getUsedRange();
            // Load only dimensions first to avoid payload size limit
            usedRange.load("rowCount, columnCount, rowIndex, columnIndex");
            await context.sync();

            const totalRows = usedRange.rowCount;
            const totalCols = usedRange.columnCount;

            // Read values and formulas in batches to avoid response payload limits
            let allValues = [];
            let allFormulas = [];

            for (let startRow = 0; startRow < totalRows; startRow += BATCH_SIZE) {
                const rowsToRead = Math.min(BATCH_SIZE, totalRows - startRow);
                const batchRange = sheet.getRangeByIndexes(usedRange.rowIndex + startRow, usedRange.columnIndex, rowsToRead, totalCols);
                batchRange.load("values, formulas");
                await context.sync();

                allValues = allValues.concat(batchRange.values);
                allFormulas = allFormulas.concat(batchRange.formulas);

                if (totalRows > BATCH_SIZE) {
                    const currentBatch = Math.floor(startRow / BATCH_SIZE) + 1;
                    const totalBatchCount = Math.ceil(totalRows / BATCH_SIZE);
                    console.log(`ImportFromExtension: Read batch ${currentBatch}/${totalBatchCount} (rows ${startRow + 1}-${startRow + rowsToRead})`);
                }
            }

            console.log(`ImportFromExtension: Read ${allValues.length} rows in ${Math.ceil(totalRows / BATCH_SIZE)} batch(es)`);

            // Get Master List headers (raw for display) and normalized forms for matching.
            // normalizeHeader trims, lowercases, and strips all whitespace, so
            // "Grade Book", "GradeBook", and "  grade book  " all collapse to the
            // same key — case and whitespace differences match automatically.
            const masterHeaders = allValues[0].map(h => String(h || ''));
            const normalizedMasterHeaders = masterHeaders.map(normalizeHeader);
            const normalizedIncomingHeaders = incomingHeaders.map(normalizeHeader);

            console.log(`ImportFromExtension: Master List headers: [${masterHeaders.join(', ')}]`);

            // Create column mapping (incoming column index -> master column index)
            const colMapping = normalizedIncomingHeaders.map(
                (incomingHeader) => normalizedMasterHeaders.indexOf(incomingHeader)
            );

            // Detect new columns from incoming data not yet on Master List
            const newColumns = [];
            for (let i = 0; i < incomingHeaders.length; i++) {
                if (colMapping[i] === -1) {
                    const newColIdx = masterHeaders.length;
                    masterHeaders.push(String(incomingHeaders[i] || ''));
                    normalizedMasterHeaders.push(normalizedIncomingHeaders[i]);
                    colMapping[i] = newColIdx;
                    newColumns.push(String(incomingHeaders[i] || ''));
                }
            }

            if (newColumns.length > 0) {
                console.log(`ImportFromExtension: Adding ${newColumns.length} new column(s) to Master List: [${newColumns.join(', ')}]`);
            }

            console.log(`ImportFromExtension: Column mapping: [${colMapping.join(', ')}]`);

            // Determine which master columns are NOT covered by the incoming data
            const matchedMasterCols = new Set(colMapping);
            const unmatchedMasterCols = [];
            for (let i = 0; i < masterHeaders.length; i++) {
                if (!matchedMasterCols.has(i)) {
                    unmatchedMasterCols.push(i);
                }
            }

            // Find the student name column in both incoming data and master list
            const incomingStudentNameCol = findColumnIndex(normalizedIncomingHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterStudentNameCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.STUDENT_NAME_COLS);
            const masterGradebookCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.gradeBook);
            const masterAssignedCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.assigned);
            const masterAttendanceCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.attendance);
            const masterMissingAssignmentsCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const incomingMissingAssignmentsCol = findColumnIndex(normalizedIncomingHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);
            const masterIdCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.STUDENT_ID_COLS);
            const incomingIdCol = findColumnIndex(normalizedIncomingHeaders, CONSTANTS.STUDENT_ID_COLS);

            if (incomingStudentNameCol === -1) {
                console.error("ImportFromExtension: Incoming data is missing a 'Student Name' column");
                return;
            }

            if (masterStudentNameCol === -1) {
                console.error("ImportFromExtension: Master List is missing a 'StudentName' column");
                return;
            }

            // Create a map of existing students to preserve certain data
            const masterDataMap = new Map();
            const valueToColorMap = new Map();

            // First pass: Get existing data to preserve
            for (let i = 1; i < allValues.length; i++) {
                const name = allValues[i][masterStudentNameCol];
                if (name) {
                    const normalizedName = normalizeName(name);
                    const gradebookFormula = (masterGradebookCol !== -1 && allFormulas[i][masterGradebookCol]) ? allFormulas[i][masterGradebookCol] : null;
                    const assignedValue = (masterAssignedCol !== -1) ? allValues[i][masterAssignedCol] : null;

                    masterDataMap.set(normalizedName, {
                        gradebookFormula: gradebookFormula,
                        assigned: assignedValue
                    });
                }
            }

            console.log(`ImportFromExtension: Created map of ${masterDataMap.size} existing students`);
            sendImportStatus('preserving_data', `Preserving data for ${masterDataMap.size} existing students...`);

            // Build preservation map for master columns not covered by the incoming data
            const preservationMap = new Map();
            const nonBlankUnmatchedCols = [];
            if (unmatchedMasterCols.length > 0) {
                // Check which unmatched columns have any non-blank data
                for (const colIdx of unmatchedMasterCols) {
                    let hasData = false;
                    for (let i = 1; i < allValues.length; i++) {
                        const val = allValues[i][colIdx];
                        if (val !== null && val !== undefined && String(val).trim() !== '') {
                            hasData = true;
                            break;
                        }
                    }
                    if (hasData) nonBlankUnmatchedCols.push(colIdx);
                }

                // Build preservation data per student (keyed by Student ID, fallback to name)
                if (nonBlankUnmatchedCols.length > 0) {
                    for (let i = 1; i < allValues.length; i++) {
                        let key;
                        if (masterIdCol !== -1) {
                            const idValue = allValues[i][masterIdCol];
                            if (idValue != null && String(idValue).trim() !== '') {
                                key = String(idValue).trim();
                            }
                        }
                        if (!key) {
                            const name = allValues[i][masterStudentNameCol];
                            if (name) key = normalizeName(name);
                        }
                        if (!key) continue;

                        const preservedValues = {};
                        const preservedFormulas = {};
                        for (const colIdx of nonBlankUnmatchedCols) {
                            const val = allValues[i][colIdx];
                            const formula = allFormulas[i][colIdx];
                            if (formula && typeof formula === 'string' && formula.startsWith('=')) {
                                preservedFormulas[colIdx] = formula;
                                preservedValues[colIdx] = val;
                            } else if (val !== null && val !== undefined && String(val).trim() !== '') {
                                preservedValues[colIdx] = val;
                            }
                        }
                        preservationMap.set(key, { values: preservedValues, formulas: preservedFormulas });
                    }
                    console.log(`ImportFromExtension: Built preservation map for ${nonBlankUnmatchedCols.length} unmatched column(s): [${nonBlankUnmatchedCols.map(i => masterHeaders[i]).join(', ')}]`);
                }
            }

            // Get colors for assigned column values
            if (masterAssignedCol !== -1) {
                const allAssignedValues = allValues.map(row => row[masterAssignedCol]);
                const uniqueValues = [...new Set(allAssignedValues.slice(1).filter(v => v && String(v).trim() !== ""))];

                if (uniqueValues.length > 0) {
                    console.log(`ImportFromExtension: Found ${uniqueValues.length} unique values in 'Assigned' column`);
                    const cellsToLoad = [];
                    uniqueValues.forEach(value => {
                        const firstInstanceIndex = allAssignedValues.indexOf(value);
                        if (firstInstanceIndex > 0) {
                            const absoluteRowIndex = usedRange.rowIndex + firstInstanceIndex;
                            const absoluteColIndex = usedRange.columnIndex + masterAssignedCol;
                            const cell = sheet.getCell(absoluteRowIndex, absoluteColIndex);
                            cell.load("format/fill/color");
                            cellsToLoad.push({ value: value, cell: cell });
                        }
                    });

                    await context.sync();

                    cellsToLoad.forEach(item => {
                        const color = item.cell.format.fill.color;
                        if (color && color !== '#ffffff' && color !== '#000000') {
                            valueToColorMap.set(item.value, color);
                        }
                    });
                    console.log(`ImportFromExtension: Cached colors for ${valueToColorMap.size} unique values`);
                }
            }

            // Categorize students as new or existing
            const newStudents = [];
            const existingStudents = [];

            for (const row of incomingData) {
                const studentName = row[incomingStudentNameCol];
                const normalizedName = normalizeName(studentName);

                if (masterDataMap.has(normalizedName)) {
                    existingStudents.push(row);
                } else {
                    newStudents.push(row);
                }
            }

            console.log(`ImportFromExtension: Found ${newStudents.length} new students and ${existingStudents.length} existing students`);

            // Clear the sheet (keeping headers)
            if (totalRows > 1) {
                const rangeToClear = sheet.getRangeByIndexes(1, 0, totalRows - 1, masterHeaders.length);
                rangeToClear.clear(Excel.ClearApplyTo.all);
                rangeToClear.getEntireRow().delete(Excel.DeleteShiftDirection.up);
                await context.sync();
                console.log("ImportFromExtension: Sheet cleared");
            }

            // Update header row to include any new columns from import
            if (newColumns.length > 0) {
                const headerRange = sheet.getRangeByIndexes(0, 0, 1, masterHeaders.length);
                headerRange.values = [masterHeaders];
                headerRange.format.font.bold = true;
                await context.sync();
                console.log(`ImportFromExtension: Updated header row with ${newColumns.length} new column(s)`);
            }

            // Combine students (new first, then existing)
            const allStudentsToWrite = [...newStudents, ...existingStudents];

            if (allStudentsToWrite.length === 0) {
                console.log("ImportFromExtension: No students to import");
                return;
            }

            console.log(`ImportFromExtension: Preparing to write ${allStudentsToWrite.length} students...`);
            const dataToWrite = [];
            const formulasToWrite = [];
            const cellsToColor = [];
            let gradebookLinksPreservedCount = 0;
            let assignedUsersPreservedCount = 0;

            allStudentsToWrite.forEach((incomingRow, index) => {
                const newRow = new Array(masterHeaders.length).fill("");
                const formulaRow = new Array(masterHeaders.length).fill(null);

                // Map incoming data to master list columns
                for (let incomingColIdx = 0; incomingColIdx < incomingRow.length; incomingColIdx++) {
                    const masterColIdx = colMapping[incomingColIdx];
                    if (masterColIdx !== -1) {
                        let cellValue = incomingRow[incomingColIdx] ?? "";

                        // Format student name to "Last, First"
                        if (masterColIdx === masterStudentNameCol) {
                            cellValue = formatToLastFirst(String(cellValue));
                        }

                        // Normalize attendance to a fraction (67 or "67%" -> 0.67)
                        // so the "0%" number format displays it as "67%" instead
                        // of "6700%".
                        if (masterColIdx === masterAttendanceCol && cellValue !== "") {
                            cellValue = normalizeAttendanceValue(cellValue);
                        }

                        // Wrap Gradebook URLs in HYPERLINK formula
                        if (masterColIdx === masterGradebookCol && cellValue) {
                            const urlString = String(cellValue).trim();
                            // Check if it looks like a URL
                            if (urlString.startsWith('http://') || urlString.startsWith('https://')) {
                                formulaRow[masterColIdx] = `=HYPERLINK("${urlString}", "Grade Book")`;
                                cellValue = "Grade Book"; // Display value
                            }
                        }

                        newRow[masterColIdx] = cellValue;
                    }
                }

                // Handle missing assignments count based on gradebook link
                if (masterMissingAssignmentsCol !== -1) {
                    const incomingMissingCount = incomingMissingAssignmentsCol !== -1 ? incomingRow[incomingMissingAssignmentsCol] : null;
                    const hasGradebookLink = masterGradebookCol !== -1 && newRow[masterGradebookCol];

                    if (hasGradebookLink) {
                        // If gradebook link exists, default to 0 if no count provided
                        const missingCount = (incomingMissingCount !== null &&
                                            incomingMissingCount !== undefined &&
                                            incomingMissingCount !== '')
                                            ? incomingMissingCount
                                            : 0;
                        newRow[masterMissingAssignmentsCol] = missingCount;
                    } else {
                        // If no gradebook link, set to blank
                        newRow[masterMissingAssignmentsCol] = '';
                    }
                }

                // Preserve existing data for this student if they already exist
                const studentName = incomingRow[incomingStudentNameCol];
                const normalizedName = normalizeName(studentName);
                if (masterDataMap.has(normalizedName)) {
                    const existingData = masterDataMap.get(normalizedName);

                    // Preserve Gradebook links
                    if (existingData.gradebookFormula) {
                        if (masterGradebookCol !== -1 && !newRow[masterGradebookCol]) {
                            formulaRow[masterGradebookCol] = existingData.gradebookFormula;
                            const match = existingData.gradebookFormula.match(/, *"([^"]+)"\)/i);
                            newRow[masterGradebookCol] = match ? match[1] : "Gradebook";
                            gradebookLinksPreservedCount++;
                        }
                    }

                    // Preserve Assigned values
                    if (existingData.assigned) {
                        if (masterAssignedCol !== -1 && !newRow[masterAssignedCol]) {
                            newRow[masterAssignedCol] = existingData.assigned;
                            assignedUsersPreservedCount++;
                        }
                    }
                }

                // Preserve values from unmatched master columns
                if (nonBlankUnmatchedCols.length > 0) {
                    let preserveKey;
                    if (incomingIdCol !== -1) {
                        const idValue = incomingRow[incomingIdCol];
                        if (idValue != null && String(idValue).trim() !== '') {
                            preserveKey = String(idValue).trim();
                        }
                    }
                    if (!preserveKey) {
                        preserveKey = normalizedName;
                    }

                    if (preserveKey && preservationMap.has(preserveKey)) {
                        const preserved = preservationMap.get(preserveKey);
                        for (const colIdx of nonBlankUnmatchedCols) {
                            if (preserved.formulas[colIdx]) {
                                formulaRow[colIdx] = preserved.formulas[colIdx];
                                newRow[colIdx] = preserved.values[colIdx] || "";
                            } else if (preserved.values[colIdx] !== undefined) {
                                newRow[colIdx] = preserved.values[colIdx];
                            }
                        }
                    }
                }

                // Check for preserved colors based on Assigned column value
                if (masterAssignedCol !== -1) {
                    const assignedValue = newRow[masterAssignedCol];
                    if (assignedValue && valueToColorMap.has(assignedValue)) {
                        cellsToColor.push({
                            rowIndex: index + 1,
                            colIndex: masterAssignedCol,
                            color: valueToColorMap.get(assignedValue)
                        });
                    }
                }

                dataToWrite.push(newRow);
                formulasToWrite.push(formulaRow);
            });

            if (gradebookLinksPreservedCount > 0) {
                console.log(`ImportFromExtension: Preserved ${gradebookLinksPreservedCount} Gradebook links`);
            }
            if (assignedUsersPreservedCount > 0) {
                console.log(`ImportFromExtension: Preserved ${assignedUsersPreservedCount} Assigned users`);
            }

            // Convert Excel serial date numbers in Course Start/End columns before writing
            await applyCourseDateFormatting(context, sheet, masterHeaders, dataToWrite);

            // Write data and formulas in batches to avoid payload size limits
            const totalBatches = Math.ceil(dataToWrite.length / BATCH_SIZE);
            console.log(`ImportFromExtension: Writing ${dataToWrite.length} rows in ${totalBatches} batch(es) of up to ${BATCH_SIZE} rows...`);
            sendImportStatus('writing_data', `Writing ${dataToWrite.length} students...`, 0);

            for (let batchStart = 0; batchStart < dataToWrite.length; batchStart += BATCH_SIZE) {
                const batchEnd = Math.min(batchStart + BATCH_SIZE, dataToWrite.length);
                const batchSize = batchEnd - batchStart;

                const batchValues = dataToWrite.slice(batchStart, batchEnd);
                const batchFormulas = formulasToWrite.slice(batchStart, batchEnd);

                const batchRange = sheet.getRangeByIndexes(1 + batchStart, 0, batchSize, masterHeaders.length);
                batchRange.values = batchValues;
                batchRange.formulas = batchFormulas;

                await context.sync();

                const currentBatch = Math.floor(batchStart / BATCH_SIZE) + 1;
                console.log(`ImportFromExtension: Data write batch ${currentBatch}/${totalBatches} completed (rows ${batchStart + 1}-${batchEnd})`);
                sendImportStatus('writing_data', `Writing batch ${currentBatch} of ${totalBatches}...`, Math.round((currentBatch / totalBatches) * 100));
            }
            console.log("ImportFromExtension: All data writes completed");

            sendImportStatus('highlighting', 'Highlighting new students and preserved data...');

            // Highlight preserved columns in light gray (applied first so cell-level colors can override)
            if (nonBlankUnmatchedCols.length > 0) {
                for (const colIdx of nonBlankUnmatchedCols) {
                    const colRange = sheet.getRangeByIndexes(1, colIdx, dataToWrite.length, 1);
                    colRange.format.fill.color = "#EDEDED"; // Lighter Gray
                }
                console.log(`ImportFromExtension: Highlighted ${nonBlankUnmatchedCols.length} preserved column(s) in light gray: [${nonBlankUnmatchedCols.map(i => masterHeaders[i]).join(', ')}]`);
            }

            // Highlight new students based on the latest ExpStartDate
            const masterExpStartDateCol = findColumnIndex(normalizedMasterHeaders, CONSTANTS.COLUMN_MAPPINGS.expectedStartDate);
            if (masterExpStartDateCol !== -1) {
                // Find the latest ExpStartDate across all students
                let latestDate = null;
                for (let i = 0; i < dataToWrite.length; i++) {
                    const dateVal = parseDate(dataToWrite[i][masterExpStartDateCol]);
                    if (dateVal && (!latestDate || dateVal > latestDate)) {
                        latestDate = dateVal;
                    }
                }

                if (latestDate) {
                    const latestDateStr = latestDate.toDateString();
                    let highlightedCount = 0;
                    for (let i = 0; i < dataToWrite.length; i++) {
                        const dateVal = parseDate(dataToWrite[i][masterExpStartDateCol]);
                        if (dateVal && dateVal.toDateString() === latestDateStr) {
                            const rowRange = sheet.getRangeByIndexes(i + 1, 0, 1, masterHeaders.length);
                            rowRange.format.fill.color = "#ADD8E6"; // Light Blue
                            highlightedCount++;
                        }
                    }
                    console.log(`ImportFromExtension: Highlighted ${highlightedCount} students with latest ExpStartDate (${latestDateStr})`);
                }
            } else if (newStudents.length > 0) {
                // Fallback: highlight new students if ExpStartDate column doesn't exist
                console.log(`ImportFromExtension: Highlighting ${newStudents.length} new students...`);
                const highlightRange = sheet.getRangeByIndexes(1, 0, newStudents.length, masterHeaders.length);
                highlightRange.format.fill.color = "#ADD8E6"; // Light Blue
            }

            // Apply preserved cell colors last so they are not overwritten by column/row fills
            if (cellsToColor.length > 0) {
                const colorBatches = Math.ceil(cellsToColor.length / FORMAT_BATCH_SIZE);
                console.log(`ImportFromExtension: Applying ${cellsToColor.length} preserved cell colors in ${colorBatches} batch(es)...`);

                for (let i = 0; i < cellsToColor.length; i += FORMAT_BATCH_SIZE) {
                    const batchEnd = Math.min(i + FORMAT_BATCH_SIZE, cellsToColor.length);
                    for (let j = i; j < batchEnd; j++) {
                        const cell = cellsToColor[j];
                        sheet.getCell(cell.rowIndex, cell.colIndex).format.fill.color = cell.color;
                    }
                    await context.sync();
                }
            }

            // Apply conditional formatting to columns
            sendImportStatus('formatting', 'Applying conditional formatting...');
            await applyGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyLastCourseGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyMissingAssignmentsConditionalFormatting(context, sheet, masterHeaders);
            await applyHoldConditionalFormatting(context, sheet, masterHeaders);
            await applyAdSAPStatusConditionalFormatting(context, sheet, masterHeaders);
            await applyNextAssignmentDueFormatting(context, sheet, masterHeaders);
            await applyAttendanceConditionalFormatting(context, sheet, masterHeaders);
            await applyLetterGradeConditionalFormatting(context, sheet, masterHeaders);
            await applyEnrollGpaConditionalFormatting(context, sheet, masterHeaders);

            // Update workbook settings with new columns for LDA column selector
            if (newColumns.length > 0) {
                try {
                    const docSettings = Office.context.document.settings.get('workbookSettings');
                    if (docSettings && Array.isArray(docSettings.columns)) {
                        const normalize = (s) => String(s || '').toLowerCase().replace(/\s+/g, '');
                        const existingSet = new Set(docSettings.columns.map(c => normalize(c.name)));
                        for (const colName of newColumns) {
                            if (!existingSet.has(normalize(colName))) {
                                docSettings.columns.push({ name: colName });
                            }
                        }
                        Office.context.document.settings.set('workbookSettings', docSettings);
                        Office.context.document.settings.saveAsync();
                        console.log('ImportFromExtension: Updated workbook settings with new columns:', newColumns);
                    }
                } catch (e) {
                    console.warn('ImportFromExtension: Failed to update workbook settings:', e);
                }
            }

            // Autofit columns
            sendImportStatus('autofitting', 'Auto-fitting columns...');
            console.log("ImportFromExtension: Autofitting columns...");
            sheet.getUsedRange().format.autofitColumns();

            await context.sync();
            console.log("ImportFromExtension: Master List import completed successfully");
        });

        if (awaitingDialogConfirmation) {
            // The dialog is still open; status updates resume from the dialog
            // handler (createMasterListAndImport recurses into this function on
            // confirm, or sends 'cancelled' on cancel).
            return;
        }

        // Check if the payload includes student objects with missing assignments
        if (payload.students && Array.isArray(payload.students)) {
            const studentsWithAssignments = payload.students.filter(
                student => student.missingAssignments && Array.isArray(student.missingAssignments) && student.missingAssignments.length > 0
            );

            if (studentsWithAssignments.length > 0) {
                console.log(`ImportFromExtension: Found ${studentsWithAssignments.length} students with missing assignments, importing...`);
                sendImportStatus('missing_assignments', `Importing missing assignments for ${studentsWithAssignments.length} students...`);
                await importMissingAssignments(studentsWithAssignments);
            }
        }

        sendImportStatus('complete', 'Master list imported successfully!');

    } catch (error) {
        console.error("ImportFromExtension: Error importing Master List:", error);
        if (error instanceof OfficeExtension.Error) {
            console.error("ImportFromExtension: Debug info:", JSON.stringify(error.debugInfo));
        }
        sendImportStatus('error', 'Import failed', 0, error.message || 'An unexpected error occurred');
    }
}
