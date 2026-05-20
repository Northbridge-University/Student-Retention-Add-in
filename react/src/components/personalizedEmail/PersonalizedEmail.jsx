import React, { useState, useEffect, useRef, useCallback } from 'react';
import ReactQuill from 'react-quill-new';
import 'react-quill-new/dist/quill.snow.css';
import PillInput from './components/PillInput';
import { EMAIL_TEMPLATES_KEY, CUSTOM_PARAMS_KEY, standardParameters, specialParameters, QUILL_EDITOR_CONFIG, PARAMETER_BUTTON_STYLES, COLUMN_MAPPINGS } from './utils/constants';
import { findColumnIndex, normalizeHeader, getTodaysLdaSheetName, getNameParts, isValidEmail, isValidHttpUrl, evaluateMapping, renderTemplate, renderCCTemplate, generateMissingAssignmentsList, buildMissingAssignmentsCache } from './utils/helpers';
import { generatePdfReceipt } from './utils/receiptGenerator';
import { downloadMailMergeTemplate, extractFieldNames } from './utils/docxGenerator';
import { downloadRecipientsXlsx } from './utils/recipientsGenerator';
import DownloadModal from './modals/DownloadModal';
import ExampleModal from './modals/ExampleModal';
import TemplatesModal from './modals/TemplatesModal';
import CustomParamModal from './modals/CustomParamModal';
import RecipientModal from './modals/RecipientModal';
import ConfirmSendModal from './modals/ConfirmSendModal';
import SuccessModal from './modals/SuccessModal';
import { getWorkbookSettings } from '../utility/getSettings';
import { MASTER_LIST_SHEET, HISTORY_SHEET } from '../../../../shared/constants.js';

export default function PersonalizedEmail({ user, onReady }) {
    // Connection state
    const [powerAutomateConnection, setPowerAutomateConnection] = useState(null);
    const [isConnected, setIsConnected] = useState(false);
    const [setupUrl, setSetupUrl] = useState('');
    const [setupStatus, setSetupStatus] = useState('');
    const [mode, setMode] = useState(null); // null (loading), 'individual', or 'powerautomate'
    const [userEmail, setUserEmail] = useState('');

    // Email composer state
    const [fromPills, setFromPills] = useState([]);
    const [ccPills, setCcPills] = useState([]);
    const [subject, setSubject] = useState('');
    const [body, setBody] = useState('');
    const [status, setStatus] = useState('');

    // Student data state
    const [studentDataCache, setStudentDataCache] = useState([]);
    const [cachedSpecialParams, setCachedSpecialParams] = useState([]);
    const [customParameters, setCustomParameters] = useState([]);
    const [recipientSelection, setRecipientSelection] = useState({
        type: 'lda',
        customSheetName: '',
        excludeDNC: true,
        excludeFillColor: true,
        excludeNoMissingAssignments: true,
        hasBeenSet: false
    });
    const [recipientCount, setRecipientCount] = useState(0);
    const [recipientDataCache, setRecipientDataCache] = useState(new Map());
    const [worksheetDataCache, setWorksheetDataCache] = useState({});

    // UI state
    const [lastFocusedInput, setLastFocusedInput] = useState(null);
    const [showMoreParams, setShowMoreParams] = useState(false);
    const [showParameters, setShowParameters] = useState(false);
    const [showRecipientHighlight, setShowRecipientHighlight] = useState(false);
    const [showSendContextMenu, setShowSendContextMenu] = useState(false);
    const [showSendTooltip, setShowSendTooltip] = useState(false);
    const quillRef = useRef(null);
    const recipientButtonRef = useRef(null);
    const sendButtonRef = useRef(null);
    const tooltipRef = useRef(null);

    // Modal states
    const [showExampleModal, setShowExampleModal] = useState(false);
    const [showTemplatesModal, setShowTemplatesModal] = useState(false);
    const [showCustomParamModal, setShowCustomParamModal] = useState(false);
    const [showRecipientModal, setShowRecipientModal] = useState(false);
    const [showConfirmModal, setShowConfirmModal] = useState(false);
    const [showSuccessModal, setShowSuccessModal] = useState(false);
    const [showDownloadModal, setShowDownloadModal] = useState(false);
    const [lastSentPayload, setLastSentPayload] = useState([]);
    const [lastSentInfo, setLastSentInfo] = useState(() => {
        try {
            const stored = localStorage.getItem('lastEmailSent');
            return stored ? JSON.parse(stored) : null;
        } catch { return null; }
    });

    // Pre-loaded templates state
    const [templates, setTemplates] = useState(null); // null = loading, [] = loaded (empty or with data)

    // Check for existing connection on mount
    useEffect(() => {
        const initializeComponent = async () => {
            // Load user email from localStorage
            try {
                const userInfoStr = localStorage.getItem('SSO_USER_INFO');
                if (userInfoStr) {
                    const userInfo = JSON.parse(userInfoStr);
                    setUserEmail((userInfo.email || '').toLowerCase());
                }
            } catch (error) {
                console.error('Error loading user email:', error);
            }

            // Load connection, custom parameters, and templates in parallel
            await Promise.all([
                checkConnection(),
                loadCustomParameters(),
                loadTemplates()
            ]);
            // Call onReady after all initialization is complete
            if (onReady) onReady();
        };
        initializeComponent();
    }, [onReady]);

    // Close send context menu when clicking outside
    useEffect(() => {
        const handleClickOutside = (event) => {
            if (showSendContextMenu && sendButtonRef.current && !sendButtonRef.current.parentElement.contains(event.target)) {
                setShowSendContextMenu(false);
            }
        };

        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, [showSendContextMenu]);

    // Setup automatic parameter highlighting
    useEffect(() => {
        if (quillRef.current) {
            const editor = quillRef.current.getEditor();

            const handleTextChange = () => {
                const text = editor.getText();
                const paramRegex = /\{([a-zA-Z0-9_]+)\}/g;
                let match;
                const paramPositions = [];

                // Find all parameter positions
                while ((match = paramRegex.exec(text)) !== null) {
                    paramPositions.push({
                        start: match.index,
                        length: match[0].length,
                        name: match[1]
                    });
                }

                // Apply formatting to each parameter
                paramPositions.forEach(param => {
                    const colors = getParameterColor(param.name);
                    const currentFormat = editor.getFormat(param.start, param.length);

                    // Only update if background or text color is different
                    if (currentFormat.background !== colors.background || currentFormat.color !== colors.color) {
                        editor.formatText(param.start, param.length, {
                            background: colors.background,
                            color: colors.color
                        }, 'silent');
                    }
                });

                // Remove formatting from incomplete parameters (e.g., just typed '{' or removed '}')
                const allText = editor.getText();
                const contents = editor.getContents();
                let index = 0;

                contents.ops.forEach(op => {
                    if (op.insert && typeof op.insert === 'string') {
                        const opText = op.insert;

                        // Check if this text has a background but doesn't contain a complete parameter
                        if (op.attributes && (op.attributes.background || op.attributes.color)) {
                            const hasCompleteParam = /\{[a-zA-Z0-9_]+\}/.test(opText);
                            if (!hasCompleteParam) {
                                // Remove background and color from incomplete parameter text
                                editor.formatText(index, opText.length, {
                                    background: false,
                                    color: false
                                }, 'silent');
                            }
                        }
                        index += opText.length;
                    }
                });
            };

            editor.on('text-change', handleTextChange);

            return () => {
                editor.off('text-change', handleTextChange);
            };
        }
    }, [customParameters]); // Re-run when custom parameters change

    // Load exclusion defaults from workbook settings once on mount. Any value that hasn't
    // been explicitly set to `false` falls back to `true`, matching the old hardcoded behavior.
    useEffect(() => {
        const wb = getWorkbookSettings() || {};
        const excludeDNC = wb.excludeDNCDefault !== false;
        const excludeFillColor = wb.excludeFillColorDefault !== false;
        const excludeNoMissingAssignments = wb.excludeNoMissingAssignmentsDefault !== false;
        setRecipientSelection(prev => prev.hasBeenSet
            ? prev
            : { ...prev, excludeDNC, excludeFillColor, excludeNoMissingAssignments });
    }, []);

    // Pre-cache recipient data after connection is established
    useEffect(() => {
        if (isConnected) {
            preCacheRecipientData();
        }
    }, [isConnected]);

    // Auto-populate From field with user's email — only on initial load. Without
    // the ref guard, removing the pill (length → 0) would re-trigger this effect
    // and immediately re-add the email, making it impossible to clear.
    const fromAutofilledRef = useRef(false);
    useEffect(() => {
        if (userEmail && !fromAutofilledRef.current && fromPills.length === 0) {
            setFromPills([userEmail]);
            fromAutofilledRef.current = true;
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [userEmail]);

    const checkConnection = async () => {
        await Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const connectionsSetting = settings.getItemOrNullObject("connections");
            connectionsSetting.load("value");
            await context.sync();

            const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
            const connection = connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');

            if (connection && connection.enabled !== false) {
                setPowerAutomateConnection(connection);
                setIsConnected(true);
                setMode('powerautomate');
            } else {
                setPowerAutomateConnection(connection || null);
                setIsConnected(false);
                setMode('individual');
            }
        });
    };

    const createConnection = async () => {
        if (!isValidHttpUrl(setupUrl)) {
            setSetupStatus("Please enter a valid HTTP URL.");
            return;
        }

        setSetupStatus("Creating connection...");

        await Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const connectionsSetting = settings.getItemOrNullObject("connections");
            connectionsSetting.load("value");
            await context.sync();

            let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
            const newConnection = {
                id: 'pa-' + Math.random().toString(36).substr(2, 9),
                name: 'Send Personalized Email',
                type: 'power-automate',
                url: setupUrl,
                actions: [],
                history: []
            };
            connections.push(newConnection);
            settings.add("connections", JSON.stringify(connections));
            await context.sync();

            setSetupStatus("Connection created successfully!");
            setTimeout(() => {
                checkConnection();
            }, 1500);
        });
    };

    const loadCustomParameters = async () => {
        const params = await getCustomParameters();
        setCustomParameters(params);
    };

    const loadTemplates = async () => {
        try {
            const loadedTemplates = await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const templatesSetting = settings.getItemOrNullObject(EMAIL_TEMPLATES_KEY);
                templatesSetting.load("value");
                await context.sync();
                return templatesSetting.value ? JSON.parse(templatesSetting.value) : [];
            });
            setTemplates(loadedTemplates);
        } catch (error) {
            console.error('Error loading templates:', error);
            setTemplates([]); // Set to empty array on error so we don't block loading
        }
    };

    const getCustomParameters = async () => {
        return Excel.run(async (context) => {
            const settings = context.workbook.settings;
            const paramsSetting = settings.getItemOrNullObject(CUSTOM_PARAMS_KEY);
            paramsSetting.load("value");
            await context.sync();
            return paramsSetting.value ? JSON.parse(paramsSetting.value) : [];
        });
    };

    const saveCustomParameters = async (params) => {
        await Excel.run(async (context) => {
            context.workbook.settings.add(CUSTOM_PARAMS_KEY, JSON.stringify(params));
            await context.sync();
        });
        setCustomParameters(params);
    };

    const getWorksheetData = async (sheetNameToFetch) => {
        if (worksheetDataCache[sheetNameToFetch]) return worksheetDataCache[sheetNameToFetch];

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem(sheetNameToFetch);
                const range = sheet.getUsedRange();
                range.load("values");
                await context.sync();

                const values = range.values;
                const data = (values.length > 1)
                    ? { headers: values[0].map(normalizeHeader), values: values.slice(1) }
                    : { headers: [], values: [] };

                setWorksheetDataCache(prev => ({ ...prev, [sheetNameToFetch]: data }));
            });
            return worksheetDataCache[sheetNameToFetch];
        } catch (error) {
            if (error.code !== 'ItemNotFound') {
                console.error(`Error loading '${sheetNameToFetch}' sheet:`, error);
            }
            setWorksheetDataCache(prev => ({ ...prev, [sheetNameToFetch]: null }));
            return null;
        }
    };

    const getStudentDataCore = async (selection, skipSpecialParams = false, specialParamsToProcess = [], onProgress = null) => {
        const { type, customSheetName, excludeDNC, excludeFillColor, excludeNoMissingAssignments } = selection;
        let sheetName;
        let selectedRowSet = null;

        if (type === 'selection') {
            sheetName = selection.selectionSheetName;
            if (!sheetName || !selection.selectedRows || selection.selectedRows.length === 0) {
                const err = new Error('No rows selected. Please select rows in Excel first.');
                err.userFacing = true;
                throw err;
            }
            selectedRowSet = new Set(selection.selectedRows);
        } else if (type === 'custom') {
            sheetName = customSheetName.trim();
            if (!sheetName) {
                const err = new Error('Custom sheet name is required.');
                err.userFacing = true;
                throw err;
            }
        } else {
            sheetName = type === 'lda' ? getTodaysLdaSheetName() : MASTER_LIST_SHEET;
        }

        const includedStudents = [];
        const excludedStudents = [];
        setWorksheetDataCache({});

        try {
            await Excel.run(async (context) => {
                // Progress tracking for counting stages
                let totalStages = 2; // load sheet data + process students
                if (excludeDNC) totalStages++;
                if (excludeFillColor) totalStages++;
                let completedStages = 0;
                const reportProgress = () => {
                    completedStages++;
                    if (onProgress) onProgress(Math.round((completedStages / totalStages) * 100));
                };

                const dncStudentIdentifiers = new Set();

                if (excludeDNC) {
                    try {
                        const historySheet = context.workbook.worksheets.getItem(HISTORY_SHEET);
                        const historyRange = historySheet.getUsedRange();
                        historyRange.load("values");
                        await context.sync();

                        const historyValues = historyRange.values;
                        if (historyValues.length > 1) {
                            const historyHeaders = historyValues[0].map(normalizeHeader);
                            const identifierIndex = findColumnIndex(historyHeaders, COLUMN_MAPPINGS.StudentIdentifier);
                            const tagsIndex = findColumnIndex(historyHeaders, COLUMN_MAPPINGS.Tags);

                            if (identifierIndex === -1 || tagsIndex === -1) {
                                console.warn("DNC exclusion: Could not find required columns in 'Student History' sheet.",
                                    { identifierIndex, tagsIndex, historyHeaders });
                            }

                            if (identifierIndex !== -1 && tagsIndex !== -1) {
                                for (let i = 1; i < historyValues.length; i++) {
                                    const row = historyValues[i];
                                    const tagsString = String(row[tagsIndex] || '').toUpperCase();
                                    const individualTags = tagsString.split(',').map(t => t.trim());
                                    const hasExcludableDnc = individualTags.some(tag =>
                                        tag.includes('DNC') && !['DNC - PHONE', 'DNC - OTHER PHONE'].includes(tag)
                                    );

                                    if (hasExcludableDnc) {
                                        const studentIdentifier = row[identifierIndex];
                                        if (studentIdentifier) dncStudentIdentifiers.add(String(studentIdentifier));
                                    }
                                }
                            }
                        }
                    } catch (error) {
                        console.error("Could not process 'Student History' sheet for DNC exclusion.", error);
                    }
                    reportProgress();
                }

                const sheet = context.workbook.worksheets.getItem(sheetName);
                const usedRange = sheet.getUsedRangeOrNullObject();

                // Only load formulas if we need to process special parameters
                if (skipSpecialParams) {
                    usedRange.load("values, rowCount, columnCount, rowIndex, isNullObject");
                } else {
                    usedRange.load("values, formulas, rowCount, columnCount, rowIndex, isNullObject");
                }
                await context.sync();
                reportProgress();

                if (usedRange.isNullObject) {
                    const err = new Error(`Sheet "${sheetName}" is empty — no data found.`);
                    err.userFacing = true;
                    throw err;
                }

                const values = usedRange.values;
                const formulas = skipSpecialParams ? null : usedRange.formulas;
                const headers = values[0].map(normalizeHeader);

                // Load fill colors for ONLY the Outreach column (not entire sheet)
                // to avoid exceeding the response payload size limit
                let outreachColors = null;
                if (excludeFillColor) {
                    const outreachIdx = findColumnIndex(headers, COLUMN_MAPPINGS.Outreach);
                    if (outreachIdx !== -1) {
                        const outreachCol = sheet.getRangeByIndexes(0, outreachIdx, usedRange.rowCount, 1);
                        const colProps = outreachCol.getCellProperties({ format: { fill: { color: true } } });
                        await context.sync();
                        outreachColors = colProps.value; // array of [{format:{fill:{color}}}] per row
                    }
                    reportProgress();
                }

                const colIndices = {};
                for (const key in COLUMN_MAPPINGS) {
                    colIndices[key] = findColumnIndex(headers, COLUMN_MAPPINGS[key]);
                }

                const customParamIndices = {};
                customParameters.forEach(param => {
                    const headerIndex = headers.indexOf(normalizeHeader(param.sourceColumn));
                    if (headerIndex !== -1) customParamIndices[param.name] = headerIndex;
                });

                // Pre-load missing assignments cache once (instead of per-student)
                let missingAssignmentsCache = null;
                if (!skipSpecialParams && specialParamsToProcess.includes('MissingAssignmentsList')) {
                    missingAssignmentsCache = await buildMissingAssignmentsCache(context);
                }

                const usedRangeRowStart = usedRange.rowIndex || 0;

                for (let i = 1; i < values.length; i++) {
                    const row = values[i];
                    if (!row) continue;

                    // For row selection mode, skip rows not in the user's selection
                    if (selectedRowSet && !selectedRowSet.has(usedRangeRowStart + i)) continue;

                    const studentIdentifier = row[colIndices.StudentIdentifier];
                    const studentName = row[colIndices.StudentName];
                    const studentEmail = row[colIndices.StudentEmail] ?? '';

                    // Skip ghost/empty rows — no identifier, no name, no email
                    if (!studentIdentifier && !studentName && !studentEmail) continue;

                    const studentNameForRow = studentName || `ID: ${studentIdentifier || 'Unknown'}`;

                    if (!isValidEmail(studentEmail)) {
                        excludedStudents.push({ name: studentNameForRow, reason: 'Invalid Email' });
                        continue;
                    }

                    if (excludeDNC && colIndices.StudentIdentifier !== -1) {
                        if (studentIdentifier && dncStudentIdentifiers.has(String(studentIdentifier))) {
                            excludedStudents.push({ name: studentNameForRow, reason: 'DNC Tag' });
                            continue;
                        }
                    }

                    if (excludeFillColor && outreachColors) {
                        const cellColor = outreachColors[i]?.[0]?.format?.fill?.color;
                        if (cellColor && cellColor !== '#FFFFFF' && cellColor !== '#000000') {
                            excludedStudents.push({ name: studentNameForRow, reason: 'Fill Color', color: cellColor });
                            continue;
                        }
                    }

                    if (excludeNoMissingAssignments && colIndices.MissingAssignments !== -1) {
                        const rawMissing = row[colIndices.MissingAssignments];
                        // Only exclude on an explicit 0; blank/empty cells are ignored.
                        if (rawMissing !== null && rawMissing !== undefined && String(rawMissing).trim() !== '' && Number(rawMissing) === 0) {
                            excludedStudents.push({ name: studentNameForRow, reason: '0 Missing' });
                            continue;
                        }
                    }

                    const nameParts = getNameParts(studentName || '');
                    const student = {
                        StudentName: studentName || '',
                        FirstName: nameParts.first,
                        LastName: nameParts.last,
                        StudentEmail: studentEmail,
                        PersonalEmail: row[colIndices.PersonalEmail] ?? '',
                        Grade: row[colIndices.Grade] ?? '',
                        DaysOut: row[colIndices.DaysOut] ?? '',
                        Assigned: row[colIndices.Assigned] ?? '',
                        StudentIdentifier: studentIdentifier ?? ''
                    };

                    for (const param of customParameters) {
                        let value = '';
                        const colIndex = customParamIndices[param.name];
                        if (colIndex !== undefined) {
                            const cellValue = row[colIndex] ?? '';
                            // If no mappings configured, use the raw cell value
                            if (!param.mappings || param.mappings.length === 0) {
                                value = cellValue;
                            } else if (cellValue !== '') {
                                // If mappings exist, try to find a match
                                for (const mapping of param.mappings) {
                                    if (evaluateMapping(cellValue, mapping)) {
                                        value = mapping.then;
                                        break;
                                    }
                                }
                                // If no mapping matches, leave value blank (empty string)
                            }
                        }
                        student[param.name] = value;
                    }

                    // Process special parameters (only when generating actual emails, not during counting)
                    if (!skipSpecialParams && specialParamsToProcess.length > 0) {
                        for (const paramName of specialParamsToProcess) {
                            if (paramName === 'MissingAssignmentsList') {
                                const gradeBookIndex = colIndices.GradeBook;
                                if (gradeBookIndex !== -1 && missingAssignmentsCache) {
                                    // Look up from pre-built cache
                                    const gradeBookFormula = formulas[i][gradeBookIndex];
                                    const gradeBookValue = row[gradeBookIndex];
                                    let lookupUrl = gradeBookValue;
                                    if (gradeBookFormula) {
                                        const match = String(gradeBookFormula).match(/=HYPERLINK\("([^"]+)"/i);
                                        if (match) lookupUrl = match[1];
                                    }
                                    student.MissingAssignmentsList = missingAssignmentsCache.get(String(lookupUrl ?? '').trim()) || '';
                                } else {
                                    student.MissingAssignmentsList = '';
                                }
                            } else if (paramName === 'DaysLeft') {
                                // Calculate DaysLeft as 14 - DaysOut, defaulting to 0 if negative
                                const daysOut = parseInt(student.DaysOut, 10) || 0;
                                const daysLeft = Math.max(0, 14 - daysOut);
                                student.DaysLeft = daysLeft.toString();
                            } else if (paramName === 'Salutation') {
                                // Get time-based greeting
                                const hour = new Date().getHours();
                                let timeGreeting;
                                if (hour < 12) {
                                    timeGreeting = 'Good Morning';
                                } else if (hour < 17) {
                                    timeGreeting = 'Good Afternoon';
                                } else {
                                    timeGreeting = 'Good Evening';
                                }
                                // Randomly pick from greetings including time-based one
                                const greetings = ['Dear', 'Hello', 'Greetings', timeGreeting];
                                student.Salutation = greetings[Math.floor(Math.random() * greetings.length)];
                            }
                        }
                    }

                    includedStudents.push(student);
                }

                // Deduplicate by StudentIdentifier — if the same ID appears more than once,
                // keep only the first occurrence and exclude the rest to prevent sending
                // multiple emails to the same student.
                const seenIdentifiers = new Map(); // identifier → first index
                const duplicateIndices = new Set();
                for (let idx = 0; idx < includedStudents.length; idx++) {
                    const id = String(includedStudents[idx].StudentIdentifier || '').trim();
                    if (!id) continue; // skip students with no identifier
                    if (seenIdentifiers.has(id)) {
                        duplicateIndices.add(idx);
                    } else {
                        seenIdentifiers.set(id, idx);
                    }
                }
                if (duplicateIndices.size > 0) {
                    // Remove duplicates from included (iterate in reverse to preserve indices)
                    const removedStudents = [];
                    for (let idx = includedStudents.length - 1; idx >= 0; idx--) {
                        if (duplicateIndices.has(idx)) {
                            removedStudents.push(includedStudents.splice(idx, 1)[0]);
                        }
                    }
                    for (const s of removedStudents) {
                        excludedStudents.push({ name: s.StudentName || `ID: ${s.StudentIdentifier}`, reason: 'Duplicate' });
                    }
                }

                reportProgress();
            });
            return { included: includedStudents, excluded: excludedStudents };
        } catch (error) {
            if (error.code === 'ItemNotFound') {
                error.userFacingMessage = `Error: Sheet "${sheetName}" not found.`;
            } else if (error.userFacing) {
                error.userFacingMessage = error.message;
            } else if (error.code === 'InvalidArgument' || error.code === 'GeneralException') {
                error.userFacingMessage = `Error reading "${sheetName}": ${error.message || 'The sheet may be too large or empty.'}`;
            } else {
                error.userFacingMessage = `Error: ${error.message || 'An unknown error occurred while reading the sheet.'}`;
            }
            console.error(`[PersonalizedEmail] getStudentDataCore error on sheet "${sheetName}":`, error);
            throw error;
        }
    };

    const getStudentDataWithUI = async () => {
        try {
            // Determine which special parameters are actually used in the email template
            const specialParamsToProcess = specialParameters.filter(param =>
                isParameterUsedInTemplate(param)
            );

            // Set specific status message based on what's being fetched
            if (specialParamsToProcess.includes('MissingAssignmentsList')) {
                setStatus('Fetching students and missing assignments...');
            } else {
                setStatus('Fetching students...');
            }

            const result = await getStudentDataCore(recipientSelection, false, specialParamsToProcess);
            setStudentDataCache(result.included);
            setCachedSpecialParams(specialParamsToProcess);
            setStatus(`Found ${result.included.length} students.`);
            setTimeout(() => setStatus(''), 3000);
            return result.included;
        } catch (error) {
            const message = error.userFacingMessage || (error.userFacing ? error.message : 'An error occurred while fetching data.');
            setStatus(message);
            throw error;
        }
    };

    const preCacheRecipientData = async () => {
        try {
            const wb = getWorkbookSettings() || {};
            const excludeDNC = wb.excludeDNCDefault !== false;
            const excludeFillColor = wb.excludeFillColorDefault !== false;
            const excludeNoMissingAssignments = wb.excludeNoMissingAssignmentsDefault !== false;

            const ldaSelection = { type: 'lda', customSheetName: '', excludeDNC, excludeFillColor, excludeNoMissingAssignments };
            const ldaResult = await getStudentDataCore(ldaSelection, true); // Skip special params for faster caching
            setRecipientDataCache(prev => new Map(prev).set('lda', ldaResult));

            const masterSelection = { type: 'master', customSheetName: '', excludeDNC, excludeFillColor, excludeNoMissingAssignments };
            const masterResult = await getStudentDataCore(masterSelection, true); // Skip special params for faster caching
            setRecipientDataCache(prev => new Map(prev).set('master', masterResult));
        } catch (error) {
            console.warn("Pre-caching failed. This may happen if sheets are not yet created.", error);
        }
    };

    const handleRecipientUpdate = (newSelection, count) => {
        setRecipientSelection({ ...newSelection, hasBeenSet: true });
        setRecipientCount(count);
        // Clear cache to ensure fresh data is fetched when needed
        setStudentDataCache([]);
        setCachedSpecialParams([]);
    };

    const isParameterUsedInTemplate = (paramName) => {
        // Check if the parameter is used in any part of the email template
        const template = `${fromPills[0] || ''} ${ccPills.join(' ')} ${subject} ${body}`;
        return template.includes(`{${paramName}}`);
    };

    const ensureStudentDataLoaded = async () => {
        if (!recipientSelection.hasBeenSet) return studentDataCache;
        // Re-fetch when cache is empty, or when the template now references a
        // special parameter that wasn't resolved the last time we fetched.
        const needed = specialParameters.filter(p => isParameterUsedInTemplate(p));
        const missingSpecial = needed.some(p => !cachedSpecialParams.includes(p));
        if (studentDataCache.length === 0 || missingSpecial) {
            return await getStudentDataWithUI();
        }
        return studentDataCache;
    };

    const handleOpenExampleModal = async () => {
        await ensureStudentDataLoaded();
        setShowExampleModal(true);
    };

    const handleExampleButtonClick = async () => {
        // If no recipients selected, scroll to recipient button and highlight it
        if (recipientCount === 0) {
            recipientButtonRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
            setShowRecipientHighlight(true);
            setTimeout(() => setShowRecipientHighlight(false), 2000);
            return;
        }
        // Otherwise, open the example modal
        await handleOpenExampleModal();
    };

    const handleOpenConfirmModal = async () => {
        await ensureStudentDataLoaded();
        setShowConfirmModal(true);
    };

    const getParameterColor = (paramName) => {
        // Color palette is intentionally aligned with PARAMETER_BUTTON_STYLES so a
        // parameter looks the same in the button row, the body editor, and pill inputs.
        if (specialParameters.includes(paramName)) {
            return {
                background: '#ffedd5', // orange-100
                color: '#9a3412'       // orange-800
            };
        }

        const customParam = customParameters.find(p => p.name === paramName);
        if (customParam) {
            const hasMappings = customParam.mappings && customParam.mappings.length > 0;
            const hasNested = hasMappings && customParam.mappings.some(m => /\{(\w+)\}/.test(m.then));

            if (hasNested) {
                return {
                    background: '#ffe4e6', // rose-100
                    color: '#881337'       // rose-800
                };
            }
            if (hasMappings) {
                return {
                    background: '#f3e8ff', // purple-100
                    color: '#581c87'       // purple-800
                };
            }
            return {
                background: '#dbeafe', // blue-100
                color: '#1e3a8a'       // blue-800
            };
        }

        // Standard parameter
        return {
            background: '#e5e7eb', // gray-200
            color: '#374151'       // gray-700
        };
    };

    const insertParameter = (param) => {
        if (lastFocusedInput === 'quill' && quillRef.current) {
            const editor = quillRef.current.getEditor();
            const range = editor.getSelection(true);

            // Extract parameter name from {ParamName} format
            const paramName = param.replace(/[{}]/g, '');
            const colors = getParameterColor(paramName);

            // Insert with background and text color formatting
            editor.insertText(range.index, param, {
                background: colors.background,
                color: colors.color
            });

            // Insert a zero-width space after the parameter without formatting
            // This creates a buffer that prevents the cursor from picking up the highlight
            const bufferPosition = range.index + param.length;
            editor.insertText(bufferPosition, '\u200B', {
                background: false,
                color: false
            });

            // Move cursor after the buffer
            editor.setSelection(bufferPosition + 1);
        } else if (lastFocusedInput === 'from') {
            setFromPills([param]);
        } else if (lastFocusedInput === 'cc') {
            setCcPills(prev => [...prev, param]);
        } else if (lastFocusedInput === 'subject') {
            const input = document.getElementById('email-subject');
            if (input) {
                const start = input.selectionStart || 0;
                const end = input.selectionEnd || 0;
                const text = subject;
                const newValue = text.substring(0, start) + param + text.substring(end);
                setSubject(newValue);
                setTimeout(() => {
                    input.focus();
                    input.selectionStart = input.selectionEnd = start + param.length;
                }, 0);
            }
        } else if (quillRef.current) {
            const editor = quillRef.current.getEditor();
            editor.focus();

            const paramName = param.replace(/[{}]/g, '');
            const colors = getParameterColor(paramName);

            const insertPosition = editor.getLength() - 1;
            editor.insertText(insertPosition, param, {
                background: colors.background,
                color: colors.color
            });

            // Insert a zero-width space after the parameter without formatting
            const bufferPosition = insertPosition + param.length;
            editor.insertText(bufferPosition, '\u200B', {
                background: false,
                color: false
            });

            // Move cursor after the buffer
            editor.setSelection(bufferPosition + 1);
        }
    };

    const stripParameterBackgrounds = (html) => {
        // Create a temporary div to parse HTML
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;

        // Find all elements with background or color styling
        const styledElements = tempDiv.querySelectorAll('[style*="background"], [style*="color"]');

        styledElements.forEach(element => {
            const text = element.textContent || '';
            // Only strip background and color if this element contains a parameter pattern
            if (/\{[a-zA-Z0-9_]+\}/.test(text)) {
                // Remove background-color and color from inline style
                const style = element.getAttribute('style') || '';
                const newStyle = style
                    .replace(/background-color:\s*[^;]+;?/gi, '')
                    .replace(/background:\s*[^;]+;?/gi, '')
                    .replace(/color:\s*[^;]+;?/gi, '')
                    .trim();

                if (newStyle) {
                    element.setAttribute('style', newStyle);
                } else {
                    element.removeAttribute('style');
                }
            }
        });

        return tempDiv.innerHTML;
    };

    const generatePayload = () => {
        const fromTemplate = fromPills[0] || '';
        // Strip parameter backgrounds from body before rendering
        const cleanBodyHtml = stripParameterBackgrounds(body);

        const emails = studentDataCache.map(student => ({
            from: renderTemplate(fromTemplate, student),
            to: student.StudentEmail || '',
            cc: renderCCTemplate(ccPills, student),
            subject: renderTemplate(subject, student),
            body: renderTemplate(cleanBodyHtml, student)
        })).filter(email => email.to && email.from);

        // Calculate sender breakdown
        const senderCounts = emails.reduce((acc, email) => {
            const from = email.from || 'Unknown';
            acc[from] = (acc[from] || 0) + 1;
            return acc;
        }, {});

        const senders = Object.entries(senderCounts).map(([email, count]) => ({
            email,
            count
        }));

        return {
            byName: user || '',
            byEmail: userEmail || '',
            totalCount: emails.length,
            senders: senders,
            emails: emails
        };
    };

    const sendEmailsViaPowerAutomate = async () => {
        setShowConfirmModal(false);
        setStatus(`Sending ${studentDataCache.length} emails...`);

        const payload = generatePayload();

        if (payload.emails.length === 0) {
            setStatus('No students with valid "To" and "From" email addresses found.');
            return;
        }

        // Generate base64 PDF receipt to include in payload
        const initiator = { name: user, email: userEmail };
        const receiptBase64 = generatePdfReceipt(
            payload.emails,
            body,
            initiator,
            true // returnBase64
        );

        // Add receipt to payload
        const payloadWithReceipt = {
            ...payload,
            receipt: receiptBase64 || ''
        };

        setLastSentPayload(payloadWithReceipt);

        try {
            // text/plain keeps this a CORS-simple request — Power Automate's
            // new *.environment.api.powerplatform.com endpoint doesn't answer
            // OPTIONS preflight; the trigger still parses the body as JSON.
            const response = await fetch(powerAutomateConnection.url, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
                body: JSON.stringify(payloadWithReceipt)
            });
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            const info = { count: payload.emails.length, timestamp: new Date().toISOString() };
            setLastSentInfo(info);
            try { localStorage.setItem('lastEmailSent', JSON.stringify(info)); } catch {}
            setStatus(`Successfully sent ${payload.emails.length} emails!`);
            setShowSuccessModal(true);
        } catch (error) {
            setStatus(`Failed to send emails: ${error.message}`);
            console.error("Error sending emails:", error);
        }
    };

    const executeSend = async () => {
        if (mode === 'powerautomate') {
            await sendEmailsViaPowerAutomate();
        } else if (mode === 'individual') {
            openDownloadModal();
        } else {
            setStatus('Invalid sending mode. Please refresh and try again.');
        }
    };

    const openDownloadModal = () => {
        // Show the modal so the user can pick template, recipients, or both.
        // Validation is per-button: template needs a body, recipients needs students.
        setShowDownloadModal(true);
    };

    const downloadTemplateOnly = async () => {
        if (!body || !body.trim()) {
            setStatus('Please write an email body first.');
            return;
        }
        const cleanBodyHtml = stripParameterBackgrounds(body);
        const stamp = new Date().toISOString().slice(0, 10);
        try {
            await downloadMailMergeTemplate(cleanBodyHtml, `email-template-${stamp}.docx`);
            setStatus('Downloaded Word template.');
        } catch (error) {
            setStatus(`Failed to generate template: ${error.message}`);
            console.error('Error generating template:', error);
            throw error;
        }
    };

    const downloadRecipientsOnly = async () => {
        // Use the returned data directly — reading studentDataCache here would see
        // the stale closure value on the first call because setStudentDataCache is async.
        const students = await ensureStudentDataLoaded();
        const cleanBodyHtml = body ? stripParameterBackgrounds(body) : '';
        const fieldNames = extractFieldNames(`${cleanBodyHtml} ${subject || ''}`);
        const recipients = (students || []).filter(s => isValidEmail(s.StudentEmail));

        if (recipients.length === 0) {
            setStatus('No students with valid email addresses found.');
            return;
        }

        const stamp = new Date().toISOString().slice(0, 10);
        try {
            await downloadRecipientsXlsx(recipients, fieldNames, `email-recipients-${stamp}.xlsx`);
            const info = { count: recipients.length, timestamp: new Date().toISOString(), action: 'download' };
            setLastSentInfo(info);
            try { localStorage.setItem('lastEmailSent', JSON.stringify(info)); } catch {}
            setStatus(`Downloaded recipient list (${recipients.length} ${recipients.length === 1 ? 'row' : 'rows'}).`);
        } catch (error) {
            setStatus(`Failed to generate recipient list: ${error.message}`);
            console.error('Error generating recipient list:', error);
            throw error;
        }
    };

    const handleSendTestEmail = async () => {
        setShowSendContextMenu(false);

        // Validate we have minimum required fields (From, Subject, Body)
        const from = fromPills[0] || '';
        if (!from.trim() || !subject.trim() || !body.trim()) {
            setStatus('Please fill in From, Subject, and Body to send a test email.');
            return;
        }

        if (!userEmail) {
            setStatus('Unable to determine your email address. Please sign in again.');
            return;
        }

        // Ensure student data is loaded
        await ensureStudentDataLoaded();

        // Pick a random student for parameter replacement, or use a placeholder if no students
        let testStudent;
        if (studentDataCache.length > 0) {
            const randomIndex = Math.floor(Math.random() * studentDataCache.length);
            testStudent = studentDataCache[randomIndex];
        } else {
            // Create placeholder student data if no students selected
            testStudent = {
                StudentEmail: userEmail,
                FirstName: 'Test',
                LastName: 'Student',
                StudentName: 'Test Student'
            };
        }

        // Generate single test email payload with user's email as recipient
        const fromTemplate = fromPills[0] || '';
        const cleanBodyHtml = stripParameterBackgrounds(body);
        const testFromEmail = renderTemplate(fromTemplate, testStudent);

        const testPayload = {
            byName: user || '',
            byEmail: userEmail || '',
            totalCount: 1,
            senders: [{ email: testFromEmail, count: 1 }],
            emails: [{
                from: testFromEmail,
                to: userEmail, // Always send to the logged-in user
                cc: '', // Don't CC anyone on test emails
                subject: `[TEST] ${renderTemplate(subject, testStudent)}`,
                body: renderTemplate(cleanBodyHtml, testStudent)
            }]
        };

        setLastSentPayload(testPayload);
        setStatus('Sending test email to yourself...');

        try {
            if (mode === 'powerautomate') {
                const initiator = { name: user, email: userEmail };
                const receiptBase64 = generatePdfReceipt(
                    testPayload.emails,
                    body,
                    initiator,
                    true // returnBase64
                );

                const testPayloadWithReceipt = {
                    ...testPayload,
                    receipt: receiptBase64 || ''
                };

                const response = await fetch(powerAutomateConnection.url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
                    body: JSON.stringify(testPayloadWithReceipt)
                });

                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                setStatus(`Test email sent to ${userEmail}!`);
            } else {
                setStatus('Invalid sending mode. Please refresh and try again.');
            }
        } catch (error) {
            setStatus(`Failed to send test email: ${error.message}`);
            console.error("Error sending test email:", error);
        }
    };

    const isFormValid = () => {
        const from = fromPills[0] || '';
        const isFromValid = from && from.trim() !== '';
        const isSubjectValid = subject && subject.trim() !== '';
        const isBodyValid = body && body.trim() !== '';
        const areRecipientsValid = recipientSelection.hasBeenSet && recipientCount > 0;
        return isFromValid && isSubjectValid && isBodyValid && areRecipientsValid;
    };

    const formatLastSent = (iso) => {
        const diff = Date.now() - new Date(iso).getTime();
        const mins = Math.floor(diff / 60000);
        if (mins < 1) return 'just now';
        if (mins < 60) return `${mins}m ago`;
        const hrs = Math.floor(mins / 60);
        if (hrs < 24) return `${hrs}h ago`;
        const days = Math.floor(hrs / 24);
        if (days < 7) return `${days}d ago`;
        return new Date(iso).toLocaleDateString();
    };

    const getValidationMessage = () => {
        const missing = [];
        if (!fromPills[0] || !fromPills[0].trim()) missing.push('From address');
        if (!recipientSelection.hasBeenSet || recipientCount === 0) missing.push('Recipients');
        if (!subject || !subject.trim()) missing.push('Subject');
        if (!body || !body.trim()) missing.push('Body');
        return missing.length > 0 ? `Required: ${missing.join(', ')}.` : '';
    };

    // Position the Send-button tooltip so it stays within the visible taskpane
    useEffect(() => {
        const btn = sendButtonRef.current;
        const tip = tooltipRef.current;
        if (!showSendTooltip || !btn || !tip) return;

        const btnRect = btn.getBoundingClientRect();
        const tipRect = tip.getBoundingClientRect();
        const pad = 8;

        // Try above the button first
        let top = btnRect.top - tipRect.height - pad;
        // If clipped at the top, flip below the button
        if (top < pad) top = btnRect.bottom + pad;

        // Center horizontally on the button, then clamp to viewport
        let left = btnRect.left + btnRect.width / 2 - tipRect.width / 2;
        left = Math.max(pad, Math.min(left, window.innerWidth - tipRect.width - pad));

        tip.style.top = `${top}px`;
        tip.style.left = `${left}px`;
        tip.style.visibility = 'visible';
    }, [showSendTooltip]);

    const renderParameterButton = (param) => {
        const isCustom = typeof param === 'object';
        const paramName = isCustom ? param.name : param;

        let buttonClass = PARAMETER_BUTTON_STYLES.standard;

        // Check if this is a special parameter
        if (specialParameters.includes(paramName)) {
            buttonClass = PARAMETER_BUTTON_STYLES.special;
        } else if (isCustom) {
            const hasMappings = param.mappings && param.mappings.length > 0;
            const hasNested = hasMappings && param.mappings.some(m => /\{(\w+)\}/.test(m.then));

            if (hasNested) buttonClass = PARAMETER_BUTTON_STYLES.nested;
            else if (hasMappings) buttonClass = PARAMETER_BUTTON_STYLES.mapped;
            else buttonClass = PARAMETER_BUTTON_STYLES.custom;
        }

        return (
            <button
                key={paramName}
                className={buttonClass}
                onClick={() => insertParameter(`{${paramName}}`)}
            >
                {`{${paramName}}`}
            </button>
        );
    };

    // Show loading screen while checking connection/mode or loading templates
    if (mode === null || templates === null) {
        return (
            <div className="max-w-md mx-auto p-4 bg-gray-50 min-h-screen flex items-center justify-center">
                <div className="text-center">
                    <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
                    <p className="text-gray-600">Loading...</p>
                </div>
            </div>
        );
    }

    // Show work-in-progress screen when no Power Automate connection is configured.
    // Email Composer View — shared by powerautomate (send) and individual (download as .docx) modes
    return (
        <div className="max-w-md mx-auto p-4 bg-gray-50 select-none">
            <div className="flex justify-between items-center mb-4">
                <h1 className="text-xl font-bold text-gray-800">Send Emails</h1>
                <button
                    onClick={() => setShowTemplatesModal(true)}
                    className="px-3 py-1 bg-gray-200 text-gray-800 text-sm font-semibold rounded-md hover:bg-gray-300"
                >
                    Templates
                </button>
            </div>

            <div className="space-y-4">
                {/* From Field */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">
                        From {mode === 'individual' && <span className="text-xs text-gray-500">(Individual Mode)</span>}
                    </label>
                    <PillInput
                        pills={fromPills}
                        onPillsChange={setFromPills}
                        placeholder="sender@example.com"
                        singleValue={true}
                        onFocus={() => setLastFocusedInput('from')}
                        readOnly={mode === 'individual'}
                        noWrap={true}
                        getPillColor={getParameterColor}
                    />
                </div>

                {/* Recipient Selection */}
                <div>
                    <label className="block text-sm font-medium text-gray-700">Recipient</label>
                    <button
                        ref={recipientButtonRef}
                        onClick={() => setShowRecipientModal(true)}
                        className={`mt-1 w-full text-left px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all duration-500 ${
                            recipientSelection.hasBeenSet && recipientCount > 0
                                ? 'bg-green-100 text-green-800 font-semibold'
                                : 'bg-white'
                        } ${
                            showRecipientHighlight ? 'ring-2 ring-red-300 bg-red-50' : ''
                        }`}
                    >
                        {recipientSelection.hasBeenSet && recipientCount > 0
                            ? `${recipientCount} Student${recipientCount !== 1 ? 's' : ''} Selected`
                            : 'Select Students'}
                    </button>
                </div>

            </div>

            {/* Lower section */}
            <div>
            <div className="space-y-4 mt-4">
                {/* CC Field — hidden in individual/download mode because Word's mail-merge
                    Send Email dialog has no CC option, so CC values would be silently dropped. */}
                {mode !== 'individual' && (
                    <div>
                        <label className="block text-sm font-medium text-gray-700">CC</label>
                        <PillInput
                            pills={ccPills}
                            onPillsChange={setCcPills}
                            placeholder="Add an additional email"
                            onFocus={() => setLastFocusedInput('cc')}
                            noWrap={true}
                            getPillColor={getParameterColor}
                        />
                    </div>
                )}

                {/* Subject */}
                <div>
                    <label htmlFor="email-subject" className="block text-sm font-medium text-gray-700">
                        Subject
                    </label>
                    <input
                        type="text"
                        id="email-subject"
                        value={subject}
                        onChange={(e) => setSubject(e.target.value)}
                        onFocus={() => setLastFocusedInput('subject')}
                        className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                        placeholder="Email Subject"
                    />
                </div>

                {/* Body (Quill Editor) — selectable; the outer container disables selection elsewhere */}
                <div className="select-text">
                    <label className="block text-sm font-medium text-gray-700 select-none">Body</label>
                    <ReactQuill
                        ref={quillRef}
                        theme="snow"
                        value={body}
                        onChange={setBody}
                        onFocus={() => setLastFocusedInput('quill')}
                        modules={QUILL_EDITOR_CONFIG.modules}
                        className="mt-1 bg-white"
                        style={{ height: '192px', marginBottom: '48px' }}
                    />
                </div>

                {/* Parameters — collapsed by default so users with saved templates have more space */}
                <div className="mt-3">
                    <button
                        type="button"
                        onClick={() => setShowParameters(prev => !prev)}
                        aria-expanded={showParameters}
                        className="flex items-center w-full text-left text-sm font-medium text-gray-700 hover:text-gray-900"
                    >
                        <svg
                            className={`h-4 w-4 mr-1 transition-transform ${showParameters ? 'rotate-90' : ''}`}
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                        >
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                        </svg>
                        Insert Parameter
                    </button>

                    {showParameters && (
                        <>
                            <div className="mt-2 flex flex-wrap gap-2">
                                {standardParameters.map(param => renderParameterButton(param))}
                            </div>

                            {specialParameters.length > 0 && (
                                <div className="mt-3">
                                    <label className="block text-xs font-medium text-gray-600 mb-2">Special Parameters</label>
                                    <div className="flex flex-wrap gap-2">
                                        {specialParameters.map(param => renderParameterButton(param))}
                                    </div>
                                </div>
                            )}

                            <div className="mt-3">
                                <div className="flex items-center gap-1 mb-2">
                                    <label className="text-xs font-medium text-gray-600">Custom Parameters</label>
                                    <button
                                        type="button"
                                        onClick={() => setShowCustomParamModal(true)}
                                        aria-label="Create or edit custom parameters"
                                        title="Create or edit custom parameters"
                                        className="p-0.5 text-gray-400 hover:text-blue-600"
                                    >
                                        <svg className="h-3.5 w-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                                        </svg>
                                    </button>
                                </div>
                                {customParameters.length > 0 && (
                                    <>
                                        <div className="flex flex-wrap gap-2">
                                            {customParameters.slice(0, 5).map(param => renderParameterButton(param))}
                                        </div>
                                        {customParameters.length > 5 && (
                                            <>
                                                {showMoreParams && (
                                                    <div className="flex flex-wrap gap-2 mt-2">
                                                        {customParameters.slice(5).map(param => renderParameterButton(param))}
                                                    </div>
                                                )}
                                                <button
                                                    onClick={() => setShowMoreParams(!showMoreParams)}
                                                    className="mt-2 text-xs text-blue-600 hover:underline"
                                                >
                                                    {showMoreParams ? 'Show Less' : `Show ${customParameters.length - 5} More...`}
                                                </button>
                                            </>
                                        )}
                                    </>
                                )}
                            </div>
                        </>
                    )}
                </div>
            </div>

            {/* Action Buttons */}
            <div className="mt-6 flex space-x-2">
                <button
                    onClick={handleExampleButtonClick}
                    disabled={recipientCount === 0}
                    className={`${user === 'Guest' ? 'w-full' : 'w-1/2'} font-bold py-2 px-4 rounded-lg focus:outline-none focus:ring-2 focus:ring-opacity-50 transition-colors duration-200 ${
                        recipientCount === 0
                            ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                            : 'bg-gray-200 text-gray-800 hover:bg-gray-300 focus:ring-gray-400'
                    }`}
                >
                    Example
                </button>
                {user !== 'Guest' && (
                    <div className="relative w-1/2">
                        <button
                            ref={sendButtonRef}
                            onClick={mode === 'individual' ? openDownloadModal : handleOpenConfirmModal}
                            onContextMenu={(e) => {
                                e.preventDefault();
                                if (mode === 'powerautomate') setShowSendContextMenu(true);
                            }}
                            onMouseEnter={() => { if (!isFormValid()) setShowSendTooltip(true); }}
                            onMouseLeave={() => setShowSendTooltip(false)}
                            disabled={!isFormValid()}
                            className="w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors duration-200 disabled:bg-gray-400 disabled:cursor-not-allowed"
                        >
                            {mode === 'individual' ? 'Download' : 'Send Email'}
                        </button>
                        {!isFormValid() && showSendTooltip && (
                            <span
                                ref={tooltipRef}
                                style={{ position: 'fixed', visibility: 'hidden', zIndex: 9999 }}
                                className="w-56 bg-gray-800 text-white text-xs rounded-md p-2 text-center pointer-events-none"
                            >
                                {getValidationMessage()}
                            </span>
                        )}
                        {/* Context Menu for Send Button */}
                        {showSendContextMenu && (
                            <div className="absolute bottom-full left-0 mb-1 w-full bg-white border border-gray-200 rounded-lg shadow-lg z-50">
                                <button
                                    onClick={handleSendTestEmail}
                                    className="w-full px-4 py-2 text-left text-sm text-gray-700 hover:bg-gray-100 rounded-lg"
                                >
                                    Send Test Email to myself
                                </button>
                            </div>
                        )}
                    </div>
                )}
            </div>
            {mode === 'individual' && user !== 'Guest' && (
                <p className="text-xs text-gray-600 mt-2 text-center">
                    Downloads a Word template + recipient list so you can email them via Word&apos;s Mailings tab.
                </p>
            )}
            {mode === 'powerautomate' && user !== 'Guest' && (
                <p className="text-xs text-gray-600 mt-2 text-center">
                    Emails will be sent via Power Automate
                </p>
            )}
            <p className="text-xs text-gray-500 mt-2 h-4 text-center">{status}</p>
            </div>

            {/* Modals */}
            <ExampleModal
                isOpen={showExampleModal}
                onClose={() => setShowExampleModal(false)}
                studentData={studentDataCache}
                fromTemplate={fromPills[0] || ''}
                ccRecipients={ccPills}
                subjectTemplate={subject}
                bodyTemplate={body}
            />

            <TemplatesModal
                isOpen={showTemplatesModal}
                onClose={() => setShowTemplatesModal(false)}
                user={user}
                userEmail={userEmail}
                currentFrom={fromPills[0] || ''}
                currentSubject={subject}
                currentBody={body}
                currentCC={ccPills}
                templates={templates}
                onTemplatesChange={setTemplates}
                onLoadTemplate={(template) => {
                    setFromPills([template.from]);
                    setSubject(template.subject);
                    setBody(template.body);
                    setCcPills(template.cc || []);
                }}
            />

            <CustomParamModal
                isOpen={showCustomParamModal}
                onClose={() => setShowCustomParamModal(false)}
                customParameters={customParameters}
                onSave={saveCustomParameters}
            />

            <RecipientModal
                isOpen={showRecipientModal}
                onClose={() => setShowRecipientModal(false)}
                currentSelection={recipientSelection}
                onConfirm={handleRecipientUpdate}
                getStudentDataCore={getStudentDataCore}
                recipientDataCache={recipientDataCache}
            />

            <ConfirmSendModal
                isOpen={showConfirmModal}
                onClose={() => setShowConfirmModal(false)}
                onConfirm={executeSend}
                count={studentDataCache.length}
            />

            <SuccessModal
                isOpen={showSuccessModal}
                onClose={() => setShowSuccessModal(false)}
                count={lastSentPayload?.emails?.length || 0}
                payload={lastSentPayload}
                bodyTemplate={body}
                initiator={{ name: user, email: userEmail }}
            />

            <DownloadModal
                isOpen={showDownloadModal}
                onClose={() => setShowDownloadModal(false)}
                onDownloadTemplate={downloadTemplateOnly}
                onDownloadRecipients={downloadRecipientsOnly}
                recipientCount={recipientCount}
            />
        </div>
    );
}
