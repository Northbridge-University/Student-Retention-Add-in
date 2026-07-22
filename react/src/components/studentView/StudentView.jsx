// 2025-12-11 14:00 EST - Version 5.0.0 - SSO Logic Removed (Lifted to App)
import React, { useState, useEffect } from 'react';
import './Styling/StudentView.css';
import StudentHeader from './Parts/Header.jsx';
import StudentDetails from './Tabs/Details.jsx';
import StudentHistory, { setHistoryLoading } from './Tabs/History.jsx';
import StudentAssignments from './Tabs/Assignments.jsx';
import MultiStudentView from './MultiStudentView.jsx';
import { onSelectionChanged, setRowSegmentFill, loadSheet, getSelectedRange, onChanged } from '../utility/ExcelAPI.jsx';
import { loadCache, loadSheetCache } from '../utility/Cache.jsx';
import { isOutreachTrigger } from './Tag';
import { addComment, resolveStudentIdentity } from '../utility/EditStudentHistory.jsx';
import chromeExtensionService from '../../../../shared/chromeExtensionService.js';
import { HISTORY_SHEET } from '../../../../shared/constants.js';
import { ASSIGNED_ALIASES } from '../../../../shared/columnAliases.js';

/* global Excel */

const activeStudent = {};
export const COLUMN_ALIASES = {
  StudentName: ['Student Name', 'Student'],
  ID: ['Student ID', 'SyStudentID','Student identifier'],
  StudentNumber: ['Student Number'],
  Shift: ['Shift'],
  Instructor: ['Instructor', 'Teacher'],
  ProgVersDescrip: ['Program', 'Program Description', 'ProgVersDescrip'],
  Gender: ['Gender'],
  Phone: ['Phone Number', 'Contact'],
  CreatedBy: ['Created By', 'Author'],
  OtherPhone: ['Other Phone', 'Alt Phone'],
  StudentEmail: ['Email', 'Student Email'],
  PersonalEmail: ['Other Email'],
  Assigned: ['Advisor'],
  AdmissionsRep: ['Admissions Rep', 'Admissions','AdmRep','AmRep'],
  ExpectedStartDate: ['Expected Start Date', 'Start Date','ExpStartDate'],
  Grade: ['Current Grade', 'Grade %', 'Grade', 'Course Grade'],
  LDA: ['Last Date of Attendance', 'LDA', 'course lda', 'CurrentLDA'],
  DaysOut: ['Days Out'],
  Gradebook: ['Gradebook','gradeBookLink'],
  MissingAssignments: ['Missing Assignments', 'Missing'],
  Outreach: ['Outreach', 'Comments', 'Notes', 'Comment'],
  ProfilePicture: ['Profile Picture', 'Photo', 'Avatar']
};

const isHeaderRowAddress = (address) => {
  if (!address) return false;
  const cleanAddress = address.includes('!') ? address.split('!')[1] : address;
  if (cleanAddress === '1:1') return true;
  return /^[A-Z]+1(:[A-Z]+1)?$/.test(cleanAddress);
};

// Helper function to get value from object with case-insensitive key matching
const getCaseInsensitive = (obj, possibleKeys) => {
  if (!obj || typeof obj !== 'object') return null;

  // First try exact matches (fastest)
  for (const key of possibleKeys) {
    if (obj[key] !== undefined && obj[key] !== null) {
      return obj[key];
    }
  }

  // If no exact match, try case-insensitive lookup
  const objKeysLower = Object.keys(obj).reduce((acc, k) => {
    acc[k.toLowerCase()] = k;
    return acc;
  }, {});

  for (const key of possibleKeys) {
    const lowerKey = key.toLowerCase();
    if (objKeysLower[lowerKey]) {
      const actualKey = objKeysLower[lowerKey];
      const value = obj[actualKey];
      if (value !== undefined && value !== null) {
        return value;
      }
    }
  }

  return null;
};

// Now accepts 'user' prop from App.jsx
function StudentView({ onReady, user }) {
    const [activeTab, setActiveTab] = useState('details');
    const [historyData, setHistory] = useState([]);
    const [assignmentData, setAssignments] = useState([]);
    const [activeStudentState, setActiveStudentState] = useState(activeStudent);
    const [selectedRowCount, setSelectedRowCount] = useState(1);
    const [selectedStudents, setSelectedStudents] = useState([]);
    const [hiddenRowCount, setHiddenRowCount] = useState(0);
    const [selectedCellPhone, setSelectedCellPhone] = useState(null); // Direct phone from single cell selection

    // Track sheet version to force re-binding of listeners on sheet switch
    const [sheetVersion, setSheetVersion] = useState(0);

    const [availableTabs, setAvailableTabs] = useState({
      history: true,
      assignments: true
    });

    // 0. Global Sheet Switch Listener
    useEffect(() => {
      const registerSheetListener = async () => {
        try {
          await Excel.run(async (context) => {
             context.workbook.worksheets.onActivated.add(() => {
                 setSheetVersion(v => v + 1);
                 return Promise.resolve();
             });
             await context.sync();
          });
        } catch (e) {
          console.warn('Failed to register sheet activation listener', e);
        }
      };
      registerSheetListener();
    }, []);

    // 1. Check for Sheet Existence
    useEffect(() => {
      const checkSheets = async () => {
        try {
          await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            context.load(sheets, 'items/name');
            await context.sync();
            const sheetNames = sheets.items.map(s => s.name);
            setAvailableTabs({
              history: sheetNames.includes(HISTORY_SHEET),
              assignments: sheetNames.some(name => ['Missing Assignments', 'Assignments'].includes(name))
            });
          });
        } catch (error) {
          console.warn('Failed to check sheet existence', error);
        }
      };
      checkSheets();
    }, [sheetVersion]);

    // 2. Tab Auto-Switch Logic
    useEffect(() => {
      if (activeTab === 'history' && !availableTabs.history) setActiveTab('details');
      if (activeTab === 'assignments' && !availableTabs.assignments) setActiveTab('details');
    }, [availableTabs, activeTab]);

    const loadHistory = () => {
        if (!activeStudentState) return;
        if (!availableTabs.history) return;

        // Prefer SyStudentId (ID), fallback to StudentNumber for backwards compatibility
        const studentId = activeStudentState.ID;
        const studentNumber = activeStudentState.StudentNumber;

        console.log('Extracted studentId (ID):', studentId);
        console.log('Extracted studentNumber:', studentNumber);

        if (!studentId && !studentNumber) return;

        console.log('Loading history for', activeStudentState.StudentName);
        setHistoryLoading(true);

        // Try cache with StudentNumber first for backwards compatibility
        if (studentNumber) {
            loadSheetCache(studentNumber)
                .then((cached) => { if (cached) setHistory(cached); })
                .catch(() => {});
        }

        // Load all history and filter by either SyStudentId or StudentNumber
        loadSheet(HISTORY_SHEET)
            .then((res) => {
                if (res && res.data && Array.isArray(res.data)) {
                    // Filter to match either SyStudentId (preferred) or StudentNumber (fallback)
                    const filtered = res.data.filter(entry => {
                        // Get the identifier value from the history entry
                        // This could be in any of these columns: SyStudentId, Student Identifier, Student Number, etc.
                        const entryIdValue = getCaseInsensitive(entry, [
                            'ID',
                            'Student ID',
                            'SyStudentID',
                            'Student Identifier',
                            'StudentID',
                            'Student Number',  // Also check Student Number column for backwards compatibility
                            'StudentNumber'
                        ]);

                        if (!entryIdValue) return false;

                        // Match if the entry's identifier matches EITHER the student's SyStudentId OR StudentNumber
                        // This ensures we find entries regardless of which column name is used in the History sheet
                        return (studentId && entryIdValue == studentId) ||
                               (studentNumber && entryIdValue == studentNumber);
                    });

                    setHistory(filtered);
                }
            })
            .catch((err) => console.error('Failed to load Student History sheet:', err));
    };

    const loadAssignments = () => {
        if (!activeStudentState || !activeStudentState.Gradebook) return;
        if (!availableTabs.assignments) return;
        loadSheet('Missing Assignments', 'Gradebook', activeStudentState.Gradebook)
            .then((res) => { setAssignments(res.data); })
            .catch((err) => console.error(err));
    };

    // 3. AUTO-LOADER
    useEffect(() => {
        if (!activeStudentState?.ID && !activeStudentState?.StudentNumber) return;
        if (availableTabs.history) loadHistory();
        if (availableTabs.assignments) loadAssignments();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [activeStudentState?.ID, activeStudentState?.StudentNumber]);

    const renderActiveTab = () => {
        if (!activeStudentState || !activeStudentState.ID) {
            return (
                <div className="p-8 text-center text-gray-500">
                    <p>Please select a student row to view details.</p>
                </div>
            );
        }
        switch (activeTab) {
        case 'history':
            return availableTabs.history ? <StudentHistory history={historyData} student={activeStudentState} reload={loadHistory} /> : <StudentDetails student={activeStudentState} />;
        case 'assignments':
            return availableTabs.assignments ? <StudentAssignments assignments={assignmentData} reload={loadAssignments} /> : <StudentDetails student={activeStudentState} />;
        case 'details':
        default:
            return <StudentDetails student={activeStudentState} />
        }
    };

  // 4. Selection Handler & Initial Load
  useEffect(() => {
    if (sheetVersion > 0) {
      setActiveStudentState({});
      setSelectedRowCount(1);
      setSelectedStudents([]);
    } 

    let handlerRef = null;
    (async () => {
      try {
        handlerRef = await onSelectionChanged(({ address, values, data, rowCount, allRows, hiddenRowCount, selectedCellValue, selectedCellHeader }) => {
          if (isHeaderRowAddress(address)) {
            setActiveStudentState({});
            setSelectedRowCount(1);
            setSelectedStudents([]);
            setHiddenRowCount(0);
            setSelectedCellPhone(null);
            return;
          }
          setSelectedRowCount(rowCount || 1);
          setSelectedStudents(allRows || []);
          setHiddenRowCount(hiddenRowCount || 0);
          setActiveStudentState(prev => ({ ...prev, ...data }));

          // Fast-path: if user selected a single cell in a Phone column (or a cell
          // whose value looks like a phone number), expose it for quick dialing.
          const PHONE_HEADERS = ['Phone', 'OtherPhone'];
          const isPhoneColumn = selectedCellHeader && PHONE_HEADERS.includes(selectedCellHeader);
          const cellStr = selectedCellValue != null ? String(selectedCellValue) : '';
          const looksLikePhone = /^\+?[\d\s\-().]{7,}$/.test(cellStr.trim());
          setSelectedCellPhone((isPhoneColumn || looksLikePhone) && cellStr.trim() ? cellStr.trim() : null);
        }, COLUMN_ALIASES);

        try {
          const sel = await getSelectedRange(COLUMN_ALIASES);
          if (sel && sel.success) {
            setSelectedRowCount(sel.rowCount || 1);
            setSelectedStudents(sel.allRows || []);
            setHiddenRowCount(sel.hiddenRowCount || 0);
            const initialRow = sel.singleRow || (Array.isArray(sel.rows) && sel.rows[0]) || null;
            if (initialRow) {
                if (initialRow.ID === 'Student ID' || initialRow.StudentName === 'Student Name') {
                    setActiveStudentState({});
                } else {
                    setActiveStudentState(prev => ({ ...prev, ...initialRow }));
                    if (sheetVersion === 0) {
                        await loadCache(initialRow);
                    }
                }
            }
          }
        } catch (gErr) {
          console.warn('getSelectedRange failed to initialize selection:', gErr);
        } finally {
            // Signal Ready regardless of selection result
            if (onReady) onReady();
        }

      } catch (err) {
        console.error('Failed to register Excel selection handler:', err);
        if (onReady) onReady();
      }
    })();

    return () => {
      if (handlerRef && typeof handlerRef.remove === 'function') {
        handlerRef.remove();
      }
    };
  }, [sheetVersion]); 

  // 5. Outreach Handler
  useEffect(() => {
    let changeHandlerRef = null;
    (async () => {
      try {
        changeHandlerRef = await onChanged(
          (changeEvent) => {
            // Skip changes caused by the Chrome extension's submission highlight
            // writing into the Outreach column — otherwise it would create an
            // unintended history comment whenever the student view is open.
            if (typeof window !== 'undefined' && window.__srkSuppressOutreachHandler) {
              return false;
            }
            const changes = (changeEvent && Array.isArray(changeEvent.changes)) ? changeEvent.changes : [];
            const matches = changes.map(ch => {
              const rawVal = ch && ch.value;
              const text = (rawVal === undefined || rawVal === null) ? '' : String(rawVal);
              try {
                const out = isOutreachTrigger(text); 
                return { change: ch, text, match: Boolean(out && out.matched), tag: out && out.tag };
              } catch (e) {
                return { change: ch, text, match: false, tag: null };
              }
            });
            matches.forEach(({ change, text, match, tag }) => {
              try {
                const { studentId, studentName } = resolveStudentIdentity(change.otherValues);
                const tagString = match && tag ? `${tag}, Outreach` : 'Outreach';
                // Highlight (or clear) the row from column A through the
                // Outreach column, skipping the Assigned column so its own
                // formatting is preserved. On a non-matching edit we only clear
                // a row that's currently yellow (the Contacted highlight) so we
                // don't accidentally wipe highlights applied in other colors.
                const fillColor = match ? 'yellow' : null;
                setRowSegmentFill(change.rowIndex, change.colIndex, fillColor, [ASSIGNED_ALIASES], match ? {} : { clearOnlyIfColor: '#FFFF00' });
                // Auto Outreach handler opts into dedupe: a single advisor edit can
                // produce repeated change events, and co-editing advisors update the
                // same cell during the day. Manual comments do NOT pass this flag.
                addComment(String(text), tagString, undefined, studentId, studentName, { dedupeOutreach: true });
              } catch (e) {
                console.warn('setRowSegmentFill/addComment failed', e);
              }
            });
            const anyMatch = matches.some(m => m.match);
            return anyMatch;
          },
          null,
          'Outreach',
          COLUMN_ALIASES,
          ['StudentName','ID','StudentNumber']
        );
      } catch (err) {
        console.error('Failed to register Excel cell-change handler:', err);
      }
    })();

    return () => {
      if (changeHandlerRef && typeof changeHandlerRef.remove === 'function') {
        try { changeHandlerRef.remove(); } catch (err) {}
      }
    };
  }, [sheetVersion]);

  // 6. Sync selected students to Chrome Extension
  useEffect(() => {
    // Only send if we have valid student data
    if (selectedRowCount > 1 && selectedStudents.length > 0) {
      // Multiple students selected
      chromeExtensionService.sendSelectedStudents(selectedStudents);
    } else if (selectedRowCount === 1 && activeStudentState?.ID) {
      // Single student selected — include directPhone if user clicked a phone cell
      chromeExtensionService.sendSelectedStudents(activeStudentState, selectedCellPhone);
    }
  }, [selectedRowCount, selectedStudents, activeStudentState, selectedCellPhone]);

  // RENDER (User is guaranteed to be logged in by App.jsx)
  return (
        <div className="studentview-outer">
            <StudentHeader student={activeStudentState} selectedRowCount={selectedRowCount} />
            {selectedRowCount === 1 ? (
                <>
                    <div className="studentview-tabs">
                        <button type="button" className={`studentview-tab ${activeTab === 'details' ? 'active' : ''}`} onClick={() => setActiveTab('details')}>Details</button>
                        {availableTabs.history && ( <button type="button" className={`studentview-tab ${activeTab === 'history' ? 'active' : ''}`} onClick={() => { loadHistory(); setActiveTab('history'); }}>History</button> )}
                        {availableTabs.assignments && ( <button type="button" className={`studentview-tab ${activeTab === 'assignments' ? 'active' : ''}`} onClick={() => { loadAssignments(); setActiveTab('assignments'); }}>Assignments</button> )}
                    </div>
                    <div className="studentview-tab-content">
                        {renderActiveTab()}
                    </div>
                </>
            ) : (
                <MultiStudentView students={selectedStudents} hiddenRowCount={hiddenRowCount} />
            )}
        </div>
    );
}

StudentView.displayName = 'StudentView';
export default React.memo(StudentView);