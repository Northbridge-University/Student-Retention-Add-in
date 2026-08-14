/*
 * conditional-formatting.js
 *
 * Conditional formatting helpers for the Master List sheet, applied after
 * data is written by the import flow. Each function targets one column
 * (or a small group of related columns) and silently skips if the column
 * isn't present. Errors are logged but never thrown — formatting is
 * non-critical and must not fail the import.
 */
import { CONSTANTS } from './constants.js';
import { findColumnIndex, normalizeHeader } from '../../shared/excel-helpers.js';

/**
 * Applies a 3-color scale conditional formatting to the grade column
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyGradeConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Grade column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const gradeColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.grade);

        if (gradeColIdx === -1) {
            console.log("ImportFromExtension: Grade column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const gradeColumnRange = sheet.getRangeByIndexes(1, gradeColIdx, range.rowCount - 1, 1);

        // Determine if grades are 0-1 or 0-100 scale by checking the first few values
        let isPercentScale = false;
        for (let i = 1; i < Math.min(range.rowCount, 10); i++) {
            if (range.values[i] && typeof range.values[i][gradeColIdx] === 'number' && range.values[i][gradeColIdx] > 1) {
                isPercentScale = true;
                break;
            }
        }

        console.log(`ImportFromExtension: Detected grade scale: ${isPercentScale ? '0-100' : '0-1'}`);

        // Clear existing conditional formats on the column to avoid duplicates
        gradeColumnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Red (low) -> Yellow (mid) -> Green (high)
        const conditionalFormat = gradeColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" }, // Red
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: isPercentScale ? "70" : "0.7", color: "#FFEB84" }, // Yellow
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" } // Green
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Grade column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies a 3-color scale conditional formatting to the Last Course Grade column
 * (same Red-Yellow-Green scale as the Grade column)
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyLastCourseGradeConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Last Course Grade column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const colIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.lastCourseGrade);

        if (colIdx === -1) {
            console.log("ImportFromExtension: Last Course Grade column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const columnRange = sheet.getRangeByIndexes(1, colIdx, range.rowCount - 1, 1);

        // Determine if grades are 0-1 or 0-100 scale by checking the first few values
        let isPercentScale = false;
        for (let i = 1; i < Math.min(range.rowCount, 10); i++) {
            if (range.values[i] && typeof range.values[i][colIdx] === 'number' && range.values[i][colIdx] > 1) {
                isPercentScale = true;
                break;
            }
        }

        console.log(`ImportFromExtension: Detected Last Course Grade scale: ${isPercentScale ? '0-100' : '0-1'}`);

        // Clear existing conditional formats on the column to avoid duplicates
        columnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Red (low) -> Yellow (mid) -> Green (high)
        const conditionalFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: isPercentScale ? "70" : "0.7", color: "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Last Course Grade column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying Last Course Grade conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to the missing assignments column to highlight 0s in light green
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyMissingAssignmentsConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Missing Assignments column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const missingAssignmentsColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.courseMissingAssignments);

        if (missingAssignmentsColIdx === -1) {
            console.log("ImportFromExtension: Missing Assignments column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const missingAssignmentsColumnRange = sheet.getRangeByIndexes(1, missingAssignmentsColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        missingAssignmentsColumnRange.conditionalFormats.clearAll();

        // Apply conditional formatting: cells with value 0 get Green, Accent 6, Lighter 80% background
        const conditionalFormat = missingAssignmentsColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        conditionalFormat.cellValue.format.fill.color = "#E2EFDA"; // Green, Accent 6, Lighter 80%
        conditionalFormat.cellValue.rule = { formula1: "0", operator: "EqualTo" };

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Missing Assignments column (0s highlighted in Green, Accent 6, Lighter 80%)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying missing assignments conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to the Hold column to highlight "Yes" values in light red
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyHoldConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Hold column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const holdColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.hold);

        if (holdColIdx === -1) {
            console.log("ImportFromExtension: Hold column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const holdColumnRange = sheet.getRangeByIndexes(1, holdColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        holdColumnRange.conditionalFormats.clearAll();

        // Apply conditional formatting: cells with "Yes" get light red background
        const conditionalFormat = holdColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
        conditionalFormat.textComparison.format.fill.color = "#FFB6C1"; // Light red (light pink)
        conditionalFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Yes" };

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Hold column ('Yes' highlighted in light red)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying hold conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to the AdSAPStatus column: highlights cells that
 * are NOT blank and NOT "SAP Met" (case-insensitive, whitespace-trimmed) in light
 * red — the same rule as the Create LDA SAP List.
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyAdSAPStatusConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to AdSAPStatus column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const adsapStatusColIdx = normalizedHeaders.indexOf('adsapstatus');

        if (adsapStatusColIdx === -1) {
            console.log("ImportFromExtension: AdSAPStatus column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("values, rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const adsapStatusColumnRange = sheet.getRangeByIndexes(1, adsapStatusColIdx, range.rowCount - 1, 1);
        adsapStatusColumnRange.load("address");

        // Clear existing conditional formats on the column to avoid duplicates
        adsapStatusColumnRange.conditionalFormats.clearAll();
        await context.sync();

        // Top-left cell of the range, relative (no sheet name, no $ anchors), e.g. "M2".
        // A custom conditional format formula is written for this cell and evaluated
        // relative to every cell down the column.
        const topLeft = adsapStatusColumnRange.address.split('!').pop().split(':')[0].replace(/\$/g, '');

        // Highlight any cell that flags a SAP concern: NOT blank and NOT "SAP Met"
        // — the same rule as the Create LDA SAP List. TRIM handles surrounding
        // whitespace, and Excel's <> comparison is case-insensitive, so "sap met"
        // and " SAP Met " are treated as met (not highlighted).
        const conditionalFormat = adsapStatusColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
        conditionalFormat.custom.format.fill.color = "#FFB6C1"; // Light red (light pink)
        conditionalFormat.custom.rule.formula = `=AND(TRIM(${topLeft})<>"",TRIM(${topLeft})<>"SAP Met")`;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to AdSAPStatus column (non-blank, non-'SAP Met' cells highlighted in light red)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying AdSAPStatus conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies left text alignment to the Next Assignment Due column
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyNextAssignmentDueFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying left alignment to Next Assignment Due column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const nextAssignmentDueColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.nextAssignmentDue);

        if (nextAssignmentDueColIdx === -1) {
            console.log("ImportFromExtension: Next Assignment Due column not found, skipping formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        // Apply left alignment to the entire column (header + data)
        const columnRange = sheet.getRangeByIndexes(0, nextAssignmentDueColIdx, range.rowCount, 1);
        columnRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

        await context.sync();
        console.log("ImportFromExtension: Left alignment applied to Next Assignment Due column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying Next Assignment Due formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting and percentage number format to the Attendance % column.
 * 3-color scale: Red (lowest) -> Yellow (70%) -> Green (highest)
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyAttendanceConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Attendance % column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const attendanceColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.attendance);

        if (attendanceColIdx === -1) {
            console.log("ImportFromExtension: Attendance % column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const attendanceColumnRange = sheet.getRangeByIndexes(1, attendanceColIdx, range.rowCount - 1, 1);

        // Format as percentage so 0.38 displays as "38%"
        attendanceColumnRange.numberFormat = [["0%"]];

        // Clear existing conditional formats on the column to avoid duplicates
        attendanceColumnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Red (low) -> Yellow (70%) -> Green (high)
        const conditionalFormat = attendanceColumnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#F8696B" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "0.7", color: "#FFEB84" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#63BE7B" }
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting and percentage format applied to Attendance % column");
    } catch (error) {
        console.error("ImportFromExtension: Error applying attendance conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies conditional formatting to Letter Grade and Last Course Letter Grade columns.
 * Highlights cells beginning with "D" in light red and cells beginning with "F" in a darker red.
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyLetterGradeConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to letter grade columns...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const letterGradeColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.letterGrade);
        const lastCourseLetterGradeColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.lastCourseLetterGrade);

        const colIndices = [];
        if (letterGradeColIdx !== -1) colIndices.push({ idx: letterGradeColIdx, name: "Letter Grade" });
        if (lastCourseLetterGradeColIdx !== -1) colIndices.push({ idx: lastCourseLetterGradeColIdx, name: "Last Course Letter Grade" });

        if (colIndices.length === 0) {
            console.log("ImportFromExtension: No letter grade columns found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        for (const col of colIndices) {
            const columnRange = sheet.getRangeByIndexes(1, col.idx, range.rowCount - 1, 1);

            // Clear existing conditional formats on the column to avoid duplicates
            columnRange.conditionalFormats.clearAll();

            // Highlight cells beginning with "F" (darker red) - added first so it has lower priority
            const fFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
            fFormat.textComparison.format.fill.color = "#FF6B6B";
            fFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.beginsWith, text: "F" };

            // Highlight cells beginning with "D" (light red) - added second so it has higher priority
            const dFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
            dFormat.textComparison.format.fill.color = "#FFB6C1";
            dFormat.textComparison.rule = { operator: Excel.ConditionalTextOperator.beginsWith, text: "D" };

            console.log(`ImportFromExtension: Conditional formatting applied to ${col.name} column (D=light red, F=darker red)`);
        }

        await context.sync();
    } catch (error) {
        console.error("ImportFromExtension: Error applying letter grade conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Applies a 3-color scale conditional formatting to the Enroll GPA column.
 * Pink (0) -> Baby Blue (2) -> Light Green (4)
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 */
export async function applyEnrollGpaConditionalFormatting(context, sheet, headers) {
    try {
        console.log("ImportFromExtension: Applying conditional formatting to Enroll GPA column...");

        const normalizedHeaders = headers.map(normalizeHeader);
        const enrollGpaColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.enrollGpa);

        if (enrollGpaColIdx === -1) {
            console.log("ImportFromExtension: Enroll GPA column not found, skipping conditional formatting");
            return;
        }

        const range = sheet.getUsedRange();
        range.load("rowCount");
        await context.sync();

        if (range.rowCount <= 1) {
            console.log("ImportFromExtension: No data rows to format");
            return;
        }

        const columnRange = sheet.getRangeByIndexes(1, enrollGpaColIdx, range.rowCount - 1, 1);

        // Clear existing conditional formats on the column to avoid duplicates
        columnRange.conditionalFormats.clearAll();

        // Apply 3-color scale: Pink (0) -> Blue (2) -> Green (4)
        const conditionalFormat = columnRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const criteria = {
            minimum: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "0", color: "#FFC7CE" },
            midpoint: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "2", color: "#9BC2E6" },
            maximum: { type: Excel.ConditionalFormatColorCriterionType.number, formula: "4", color: "#C6EFCE" }
        };
        conditionalFormat.colorScale.criteria = criteria;

        await context.sync();
        console.log("ImportFromExtension: Conditional formatting applied to Enroll GPA column (Pink #FFC7CE -> Blue #9BC2E6 -> Green #C6EFCE)");
    } catch (error) {
        console.error("ImportFromExtension: Error applying Enroll GPA conditional formatting:", error);
        // Don't throw - formatting is not critical
    }
}

/**
 * Converts Excel serial date numbers to M/DD/YYYY formatted strings in the
 * Course Start and Course End columns, then applies a date number format.
 * @param {Excel.RequestContext} context The request context
 * @param {Excel.Worksheet} sheet The worksheet to format
 * @param {string[]} headers The header row values
 * @param {Array[]} dataToWrite The data rows (mutated in place)
 */
export async function applyCourseDateFormatting(context, sheet, headers, dataToWrite) {
    try {
        const normalizedHeaders = headers.map(normalizeHeader);
        const courseStartColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.courseStart);
        const courseEndColIdx = findColumnIndex(normalizedHeaders, CONSTANTS.COLUMN_MAPPINGS.courseEnd);

        const colIndices = [];
        if (courseStartColIdx !== -1) colIndices.push({ idx: courseStartColIdx, name: "Course Start" });
        if (courseEndColIdx !== -1) colIndices.push({ idx: courseEndColIdx, name: "Course End" });

        if (colIndices.length === 0) {
            return;
        }

        console.log(`ImportFromExtension: Formatting date columns: ${colIndices.map(c => c.name).join(', ')}`);

        // Convert Excel serial numbers to M/DD/YYYY strings in the data
        const excelSerialToDateString = (serial) => {
            if (typeof serial !== 'number' || serial < 1) return serial;
            // Excel serial: days since 1900-01-01 (with the 1900 leap year bug)
            const date = new Date((serial - 25569) * 86400 * 1000);
            if (isNaN(date.getTime())) return serial;
            const month = date.getUTCMonth() + 1;
            const day = date.getUTCDate();
            const year = date.getUTCFullYear();
            return `${month}/${String(day).padStart(2, '0')}/${year}`;
        };

        for (const col of colIndices) {
            let convertedCount = 0;
            for (let i = 0; i < dataToWrite.length; i++) {
                const val = dataToWrite[i][col.idx];
                if (typeof val === 'number' && val > 25569) {
                    dataToWrite[i][col.idx] = excelSerialToDateString(val);
                    convertedCount++;
                }
            }
            if (convertedCount > 0) {
                console.log(`ImportFromExtension: Converted ${convertedCount} serial dates in ${col.name} column`);
            }
        }
    } catch (error) {
        console.error("ImportFromExtension: Error formatting course date columns:", error);
    }
}
