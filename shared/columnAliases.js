/*
 * shared/columnAliases.js
 *
 * The single source of truth for every column alias either runtime uses.
 * Each export is a list of human-readable alias strings; matching is done
 * via normalizeHeader (in shared/excel-helpers.js), which NFKC-normalizes,
 * lowercases, and strips all whitespace. As a result:
 *
 *   - You do NOT need to enumerate case variants ("Grade", "grade", "GRADE")
 *   - You do NOT need to enumerate whitespace variants ("Grade Book",
 *     "GradeBook", "grade book") — they all normalize the same
 *   - You do NOT need to enumerate Unicode variants (fullwidth, ligatures)
 *
 * Each alias appearing here is one unique normalized form. Both runtimes
 * (commands/, react/) compose their COLUMN_MAPPINGS-style maps from these.
 */

// ===== Identity =====
export const STUDENT_NAME_ALIASES = ['student name', 'student'];

// SyStudentId (school's SIS Student Information System id) and Student Number
// (a separate school-issued identifier) are intentionally distinct concepts.
// Commands' code uses them separately. React composes both for its
// "StudentIdentifier" lookup (SyStudentId is preferred — tried first).
export const STUDENT_ID_ALIASES = ['student id', 'systudentid', 'id'];
export const STUDENT_NUMBER_ALIASES = ['student number', 'student identifier'];

export const GENDER_ALIASES = ['gender'];

export const PROFILE_PICTURE_ALIASES = ['profile picture', 'photo', 'avatar'];

// ===== Outreach / Status =====
// "comments"/"notes"/"comment" are also Outreach aliases — workbooks vary on
// what they call this column. Note these overlap with COMMENT_ALIASES (used
// for the History sheet's comment column); the two concepts are
// context-specific (Master List vs History sheet).
export const OUTREACH_ALIASES = ['outreach', 'comments', 'notes', 'comment'];
export const HOLD_ALIASES = ['hold'];
// Broad: matches assigned/advisor/counselor/assigned-to.
export const ASSIGNED_ALIASES = ['assigned', 'advisor', 'counselor', 'assigned to'];
// "Created by" / "Author" / "Advisor" — overlaps deliberately with
// ASSIGNED_ALIASES on 'advisor' since the assigned advisor often is the
// author of comments.
export const CREATED_BY_ALIASES = ['created by', 'author', 'advisor'];

// ===== Grade =====
export const GRADE_ALIASES = ['grade', 'course grade', 'grades', 'current grade', 'grade %'];
// "current score" is a separate concept that some workbooks use as a
// synonym for grade — kept distinct since commands queries it separately.
export const CURRENT_SCORE_ALIASES = ['current score', 'grade', 'course grade'];
export const LAST_COURSE_GRADE_ALIASES = ['last course grade'];
export const LETTER_GRADE_ALIASES = ['letter grade'];
export const LAST_COURSE_LETTER_GRADE_ALIASES = ['last course letter grade'];
export const ENROLL_GPA_ALIASES = ['enroll gpa'];

// ===== Email / Phone =====
export const STUDENT_EMAIL_ALIASES = ['student email', 'school email', 'email'];
export const PERSONAL_EMAIL_ALIASES = ['personal email', 'other email'];
export const PRIMARY_PHONE_ALIASES = ['primary phone', 'phone', 'phone number', 'contact number', 'contact'];
export const OTHER_PHONE_ALIASES = ['other phone', 'alt phone'];

// ===== Course =====
export const COURSE_ALIASES = ['course'];
export const COURSE_ID_ALIASES = ['course id'];
export const COURSE_LAST_ACCESS_ALIASES = ['course last access'];
export const COURSE_START_ALIASES = ['course start'];
export const COURSE_END_ALIASES = ['course end'];
export const PROGRAM_VERSION_ALIASES = ['program version'];

// ===== Attendance / Schedule =====
export const DAYS_OUT_ALIASES = ['days out'];
export const LAST_LDA_ALIASES = ['lda', 'last lda', 'last date of attendance'];
// Excel column may have a literal '%' — kept as a separate normalized form
// from "attendance" alone since they encode different intent.
export const ATTENDANCE_ALIASES = ['attendance %', 'attendance percent', 'attendance'];
export const EXPECTED_START_DATE_ALIASES = ['expected start date', 'start date', 'expstartdate'];
export const NEXT_ASSIGNMENT_DUE_ALIASES = ['next assignment due'];

// ===== Assignments / Gradebook =====
export const GRADE_BOOK_ALIASES = ['grade book', 'gradebook link', 'gradelink'];
export const ASSIGNMENT_ALIASES = ['assignment'];
// Per-row assignment metadata (used in the "Missing Assignments" sheet)
export const TITLE_ALIASES = ['assignment title', 'title', 'assignment'];
export const DUE_DATE_ALIASES = ['due date', 'deadline'];
export const SCORE_ALIASES = ['score', 'points'];
export const SUBMISSION_LINK_ALIASES = ['submission link', 'submission', 'submit link'];
export const ASSIGNMENT_LINK_ALIASES = ['assignment link', 'assignment url', 'assignment page', 'link'];
// Specific-to-broad: "course missing assignments" tried first; bare "missing"
// last.
export const MISSING_ASSIGNMENTS_ALIASES = [
    'course missing assignments',
    'missing assignments',
    'missing',
];
export const COURSE_ZERO_ASSIGNMENTS_ALIASES = ['course zero assignments', 'zero assignments'];

// ===== Tags / History (Student History sheet) =====
// "category"/"type" are also tag aliases.
export const TAGS_ALIASES = ['tag', 'tags', 'history', 'category', 'type'];
// History-sheet-specific concepts:
export const TIMESTAMP_ALIASES = ['timestamp', 'date', 'time', 'created at'];
// "comment"/"notes"/"history"/"entry" — same surface as some Outreach
// aliases; the calling context (Master List vs History) decides which
// canonical concept the column maps to.
export const COMMENT_ALIASES = ['comment', 'notes', 'history', 'entry'];
export const COMMENT_ID_ALIASES = ['comment id', 'history id', 'entry id'];
