import React from 'react';
import createLdaIcon from '../../assets/icons/create-lda-icon.png';

export const defaultColumns =[
	{ name: 'Assigned', alias: ['advisor'], static: true },
	{ name: 'Student Name', alias: ['Student']},
	{ name: 'Gradebook', alias: ['Gradelink', 'gradeBookLink', 'Grade Book'], static: true, identifier: true, format: ['=HYPERLINK']},
	{ name: 'ProgramVersion', alias: ['Program','ProgVersDescrip'] },
	{ name: 'Shift', alias: ['ShiftDescrip'] },
	{ name: 'LDA', alias: ['Last Date of Attendance','Date of Attendance', 'CurrentLDA'], format: ['MM/DD/YYYY']},
	{ name: 'Days Out', alias: ['Days Out'], format: ['G-Y-R Color Scale']},
	{ name: 'Grade', alias: ['Course Grade','current score', 'current grade']  },
	{ name: 'Missing Assignments', alias: ['Total Missing'] },
	{ name: 'Outreach', alias: ['Comments','Comment'] },
	{ name: 'Phone', alias: ['Phone Number','Contact Number']},
	{ name: 'Other Phone', alias: ['Second Phone', 'Alt Phone']},
];

export const Options = [
	{ option: 'identifier', label: 'Identifier', type: 'boolean' },
	{ option: 'alias', label: 'Aliases', type: 'string' },
];

export const defaultWorkbookSettings = [
	
	{
		id: 'columns',
		label: 'Columns',
		type: 'array',
		choices: defaultColumns,
		options: Options,
		defaultValue: [],
		section: 'Create LDA',
		description: 'List of columns to format in the master list.'
	},
	{
		id: 'daysOut',
		label: 'Days Out',
		type: 'number',
		defaultValue: 5,
        section: 'Create LDA',
		description: 'Number of days out threshold.'
	},
	{
		id: 'includeFailingList',
		label: 'Failing List',
		type: 'boolean',
		defaultValue: false, // No
        section: 'Create LDA',
		description: 'Whether to include the failing student list.'
	},
	{
		id: 'includeAttendanceList',
		label: 'Attendance List',
		type: 'boolean',
		defaultValue: false,
		section: 'Create LDA',
		description: 'Whether to include the low attendance student list.'
	},
	{
		id: 'includeNextAssignmentDue',
		label: 'Next Assignment Due',
		type: 'boolean',
		defaultValue: true,
		section: 'Create LDA',
		description: "Whether to include the student's next assignment due date message in the Outreach column."
	},
	{
		id: 'includeLdatTag',
		label: 'LDA Tag',
		type: 'boolean',
		defaultValue: true, // Yes
        section: 'Create LDA',
		description: 'Whether to include the LDA tag.'
	},
		{
		id: 'includeDncTag',
		label: 'DNC Tag',
		type: 'boolean',
		defaultValue: true, // Yes
        section: 'Create LDA',
		description: 'Whether to include the Do Not Contact students.'
	},
	{
		id: 'advancedLda',
		label: 'Advanced',
		type: 'advanced',
		defaultValue: null,
		section: 'Create LDA',
		description: 'Expanded DNC messaging and column width options.'
	},
	{
		// Surfaced via the Create LDA "Advanced" modal; hidden from the normal list.
		id: 'expandedDNC',
		label: 'Expanded DNC',
		type: 'boolean',
		defaultValue: false,
		hidden: true,
		description: 'Fill the Outreach column with the DNC comment text and who left it, instead of a flat "Do not contact". Off restores the original DNC behavior.'
	},
	{
		// Surfaced via the Create LDA "Advanced" modal; hidden from the normal list.
		id: 'compactColumnWidth',
		label: 'Compact Column Width',
		type: 'boolean',
		defaultValue: true,
		hidden: true,
		description: 'Pin key columns (Outreach, ProgramVersion, etc.) to fixed widths and wrap the Outreach text. Off autosizes columns like before.'
	},
	{
		// Surfaced via the Create LDA "Advanced" modal; hidden from the normal list.
		id: 'authorFormat',
		label: 'Author',
		type: 'string',
		defaultValue: 'initials',
		hidden: true,
		description: 'How the note author appears in outreach messages: full name (Victor Blanco) or initials (vb).'
	},
	{
		// Surfaced via the Create LDA "Advanced" modal; hidden from the normal list.
		id: 'ldaTagColor',
		label: 'LDA Tag Color',
		type: 'string',
		defaultValue: '#FFFF00',
		hidden: true,
		description: 'Highlight color for LDA follow-up outreach comments.'
	},
	{
		// Surfaced via the Create LDA "Advanced" modal; hidden from the normal list.
		id: 'dncColor',
		label: 'DNC Color',
		type: 'string',
		defaultValue: '#f2bdbd',
		hidden: true,
		description: 'Fill color for DNC outreach comments and the phone-column strikethrough.'
	},
	{
		id: 'riskIndexEnabled',
		label: 'Track Risk Index',
		type: 'boolean',
		defaultValue: true,
		section: 'Risk Index',
		description: 'Save a daily snapshot of every student\'s LDA and Days Out (stored under LDAHistory in workbook settings) each time an LDA sheet is created — once per day.'
	},
	{
		id: 'riskIndexThreshold',
		label: 'Risk Threshold',
		type: 'number',
		defaultValue: 14,
		section: 'Risk Index',
		description: 'Days Out value that counts as a drop-risk hit. Reaching it counts once per drop (per LDA), no matter how many days the student stays out.'
	},
	{
		id: 'riskIndexShowColumn',
		label: 'RI Column',
		type: 'boolean',
		defaultValue: true,
		section: 'Risk Index',
		description: 'Add an RI (Risk Index) column next to Days Out on the LDA sheet showing how many times each student has hit the risk threshold.'
	},
	{
		id: 'importLdaHistory',
		label: 'Import Past Reports',
		type: 'action',
		defaultValue: null,
		section: 'Risk Index',
		description: 'Backfill the history by scanning previous "LDA M-D-YYYY" sheets in this workbook, or by uploading a CSV/Excel file. Dates that already have a snapshot are left untouched.'
	},
	{
		id: 'downloadLdaHistory',
		label: 'Download History',
		type: 'action',
		defaultValue: null,
		section: 'Risk Index',
		description: 'Download the saved LDA history (date, student ID, LDA, days out) as a CSV file.'
	},
	{
		id: 'riskLeaderboard',
		label: 'At-Risk Students',
		type: 'riskleaderboard',
		defaultValue: null,
		section: 'Risk Index',
		description: 'Students ranked by Risk Index, greatest first. Students no longer on the Master List are treated as dropped and hidden.'
	},
	{
		id: 'powerAutomateUrl',
		label: 'Power Automate',
		type: 'powerautomate',
		defaultValue: null,
		section: 'Send Emails',
		description: 'Configure Power Automate flow URL for sending personalized emails.'
	},
	{
		id: 'excludeDNCDefault',
		label: 'Exclude DNC',
		type: 'boolean',
		defaultValue: true,
		section: 'Send Emails',
		description: 'Default state for the "DNC-tagged students" exclusion toggle when composing personalized emails.'
	},
	{
		id: 'excludeFillColorDefault',
		label: 'Exclude Outreach Color',
		type: 'boolean',
		defaultValue: true,
		section: 'Send Emails',
		description: 'Default state for the "Outreach Color Filled students" exclusion toggle.'
	},
	{
		id: 'excludeNoMissingAssignmentsDefault',
		label: 'Exclude 0 Missing Assignments',
		type: 'boolean',
		defaultValue: true,
		section: 'Send Emails',
		description: 'Default state for the "No Missing Assignments" exclusion toggle (only excludes rows with an explicit 0).'
	},
	{
		id: 'emailTemplates',
		label: 'Templates',
		type: 'templates',
		defaultValue: null,
		section: 'Send Emails',
		description: 'Import or export saved email templates as JSON.'
	},
	{
		id: 'customParameters',
		label: 'Custom Parameters',
		type: 'customparameters',
		defaultValue: null,
		section: 'Send Emails',
		description: 'Import or export saved custom parameters as JSON.'
	},
	{
		id: 'createHistorySheet',
		label: 'Create Sheet',
		type: 'action',
		defaultValue: null,
		section: 'Student History',
		description: 'Create the Student History sheet with the standard column headers.'
	},
	{
		id: 'downloadHistoryCsv',
		label: 'Download Copy',
		type: 'action',
		defaultValue: null,
		section: 'Student History',
		description: 'Download the Student History sheet as a CSV file.'
	},
	{
		id: 'workbookUsers',
		label: 'Registered Users',
		type: 'userslist',
		defaultValue: null,
		section: 'Users',
		description: 'Users on file who have opened this workbook with the Student Retention Kit.'
	}
];

// create a small SVG avatar and encode as data URI for the default profile picture
const _defaultAvatarSvg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'>
	<rect fill='%23e5e7eb' width='100' height='100' rx='16'/>
	<circle fill='%239ca3af' cx='50' cy='36' r='18'/>
	<path fill='%239ca3af' d='M20 84c0-14 12-26 30-26s30 12 30 26' />
	</svg>`;
const defaultProfilePicture = `data:image/svg+xml;utf8,${encodeURIComponent(_defaultAvatarSvg)}`;

export const defaultUserSettings = [
	// Profile picture removed - now handled by UserAvatar component with SSO
	// Theme removed for production
];

// export lucide-react icons keyed by section name
export const sectionIcons = {
	'Create LDA': <img src={createLdaIcon} alt="Create LDA" style={{ width: 16, height: 16 }} aria-hidden="true" />,
	'Send Emails': (
		<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
			<path d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
		</svg>
	),
	'Risk Index': (
		<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
			<path d="M12 21a9 9 0 1 1 9-9" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
			<path d="M12 12l4.5-4.5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round"/>
			<circle cx="12" cy="12" r="1" fill="currentColor"/>
		</svg>
	),
	'Student History': (
		<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
			<path d="M12 8v4l3 3" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
			<circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.6"/>
		</svg>
	),
	'Users': (
		<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
			<path d="M17 21v-2a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v2" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
			<circle cx="10" cy="8" r="3.5" stroke="currentColor" strokeWidth="1.6"/>
			<path d="M21 21v-2a4 4 0 0 0-3-3.87" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
			<path d="M16.5 4.13a3.5 3.5 0 0 1 0 6.74" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
		</svg>
	)
};

// Optionally default export for convenience
export default {
	defaultWorkbookSettings,
	defaultUserSettings
};
