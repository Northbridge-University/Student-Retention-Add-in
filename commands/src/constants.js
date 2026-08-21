/*
 * constants.js
 *
 * Commands-runtime constants (CONSTANTS object — column mappings,
 * sheet names, settings keys) plus the name/date helpers used across
 * commands modules. Cross-runtime values live in /shared/.
 */
import { MASTER_LIST_SHEET, HISTORY_SHEET } from '../../shared/constants.js';
import {
    STUDENT_NAME_ALIASES,
    STUDENT_ID_ALIASES,
    STUDENT_NUMBER_ALIASES,
    OUTREACH_ALIASES,
    HOLD_ALIASES,
    ASSIGNED_ALIASES,
    GRADE_ALIASES,
    CURRENT_SCORE_ALIASES,
    LAST_COURSE_GRADE_ALIASES,
    LETTER_GRADE_ALIASES,
    LAST_COURSE_LETTER_GRADE_ALIASES,
    ENROLL_GPA_ALIASES,
    STUDENT_EMAIL_ALIASES,
    PERSONAL_EMAIL_ALIASES,
    PRIMARY_PHONE_ALIASES,
    OTHER_PHONE_ALIASES,
    COURSE_ALIASES,
    COURSE_ID_ALIASES,
    COURSE_LAST_ACCESS_ALIASES,
    COURSE_START_ALIASES,
    COURSE_END_ALIASES,
    PROGRAM_VERSION_ALIASES,
    DAYS_OUT_ALIASES,
    LAST_LDA_ALIASES,
    ATTENDANCE_ALIASES,
    EXPECTED_START_DATE_ALIASES,
    NEXT_ASSIGNMENT_DUE_ALIASES,
    GRADE_BOOK_ALIASES,
    MISSING_ASSIGNMENTS_ALIASES,
    COURSE_ZERO_ASSIGNMENTS_ALIASES,
} from '../../shared/columnAliases.js';

// All alias arrays come from /shared/columnAliases.js — see that file to add
// or change a column alias. Matching is case- and whitespace-insensitive
// (handled by normalizeHeader in /shared/excel-helpers.js).
//
// SyStudentId vs Student Number are intentionally separate concepts:
//   STUDENT_ID_COLS    → SyStudentId  (school SIS id; preferred identifier)
//   STUDENT_NUMBER_COLS → Student Number (a different school-issued id)
export const CONSTANTS = {
    STUDENT_NAME_COLS: STUDENT_NAME_ALIASES,
    OUTREACH_COLS: OUTREACH_ALIASES,
    STUDENT_ID_COLS: STUDENT_ID_ALIASES,
    STUDENT_NUMBER_COLS: STUDENT_NUMBER_ALIASES,
    MASTER_LIST_SHEET,
    HISTORY_SHEET,
    SETTINGS_KEY: "studentRetentionSettings",
    COLUMN_MAPPINGS: {
        course: COURSE_ALIASES,
        courseId: COURSE_ID_ALIASES,
        courseLastAccess: COURSE_LAST_ACCESS_ALIASES,
        currentScore: CURRENT_SCORE_ALIASES,
        grade: GRADE_ALIASES,
        gradeBook: GRADE_BOOK_ALIASES,
        daysOut: DAYS_OUT_ALIASES,
        lastLda: LAST_LDA_ALIASES,
        assigned: ASSIGNED_ALIASES,
        programVersion: PROGRAM_VERSION_ALIASES,
        courseMissingAssignments: MISSING_ASSIGNMENTS_ALIASES,
        courseZeroAssignments: COURSE_ZERO_ASSIGNMENTS_ALIASES,
        hold: HOLD_ALIASES,
        primaryPhone: PRIMARY_PHONE_ALIASES,
        otherPhone: OTHER_PHONE_ALIASES,
        PersonalEmail: PERSONAL_EMAIL_ALIASES,
        StudentEmail: STUDENT_EMAIL_ALIASES,
        nextAssignmentDue: NEXT_ASSIGNMENT_DUE_ALIASES,
        lastCourseGrade: LAST_COURSE_GRADE_ALIASES,
        expectedStartDate: EXPECTED_START_DATE_ALIASES,
        attendance: ATTENDANCE_ALIASES,
        letterGrade: LETTER_GRADE_ALIASES,
        lastCourseLetterGrade: LAST_COURSE_LETTER_GRADE_ALIASES,
        enrollGpa: ENROLL_GPA_ALIASES,
        courseStart: COURSE_START_ALIASES,
        courseEnd: COURSE_END_ALIASES,
    }
};

/**
 * Parses a date value from various possible formats (Date object, string, Excel serial number).
 * @param {*} dateValue The value to parse.
 * @returns {Date|null} A valid Date object or null.
 */
export function parseDate(dateValue) {
    if (!dateValue) return null;
    if (dateValue instanceof Date) {
        // ExcelJS can return Date objects for .xlsx
        return dateValue;
    }
    if (typeof dateValue === 'number') {
        // Excel serial date number
        // Check for a reasonable range to avoid treating random numbers as dates
        if (dateValue > 25569) { // Corresponds to 1970-01-01
            return new Date((dateValue - 25569) * 86400 * 1000);
        }
    }
    if (typeof dateValue === 'string') {
        // Try parsing common date formats
        const parsed = new Date(dateValue);
        if (!isNaN(parsed.getTime())) {
            return parsed;
        }
    }
    return null;
}

/**
 * Helper to normalize names from "Last, First" or "First Last" to "first last"
 * for consistent matching.
 * @param {string} name The name to normalize.
 * @returns {string} The normalized name.
 */
export const normalizeName = (name) => {
    if (!name || typeof name !== 'string') return '';
    name = name.trim().toLowerCase();
    if (name.includes(',')) {
        const parts = name.split(',').map(part => part.trim());
        if (parts.length > 1) {
            return `${parts[1]} ${parts[0]}`;
        }
    }
    return name;
};

/**
 * Helper to format names to "Last, First" format.
 * @param {string} name The name to format.
 * @returns {string} The formatted name.
 */
export const formatToLastFirst = (name) => {
    if (!name || typeof name !== 'string') return '';
    name = name.trim();
    if (name.includes(',')) {
        // Already "Last, First"
        return name.split(',').map(p => p.trim()).join(', ');
    }
    const parts = name.split(' ').filter(p => p);
    if (parts.length > 1) {
        const lastName = parts.pop();
        const firstName = parts.join(' ');
        return `${lastName}, ${firstName}`;
    }
    return name;
};

