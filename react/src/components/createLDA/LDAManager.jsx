/*
 * Timestamp: 2026-01-26 00:00:00
 * Description: Fully functional LDA Manager. Integrates with ldaProcessor.js to run the real Excel generation logic.
 * Update: Added batch progress display for large datasets (6000+ students). Shows progress bar under
 *         "Reading Master List" and "Formatting LDA Table" steps when multiple batches are being processed.
 */

import React, { useState, useEffect } from 'react';
import { Info, CheckCircle2, Circle, Loader2, ArrowLeft, AlertCircle, ChevronRight, Plus, Pencil, Trash2, X, ClipboardList } from 'lucide-react';
import { createLDA, detectCampuses, detectProgramVersions, predictAdvisorDistribution, checkMissingLDAColumns, addColumnsToMasterList, checkMasterListExists } from './ldaProcessor';

// --- CONFIGURATION: Steps matching the processor logic ---
const PROCESS_STEPS = [
  { id: 'validate', label: 'Validating Workbook Settings' },
  { id: 'read', label: 'Reading Master List' },
  { id: 'filter', label: 'Filtering Students' },
  { id: 'createSheet', label: 'Creating Sheet' },
  { id: 'tags', label: 'Applying LDA & DNC Tags' },
  { id: 'writing', label: 'Writing Tables' },
  { id: 'highlights', label: 'Applying Highlights' },
  { id: 'hyperlinks', label: 'Inserting Hyperlinks' },
  { id: 'finalize', label: 'Finalizing Report' },
];

export default function CreateLDAManager({ onReady } = {}) {
  // State for the settings
  const [ldaSettings, setLdaSettings] = useState({
    daysOut: 5,
    includeFailingList: false,
    includeAttendanceList: false,
    includeLDATag: true,
    includeDNCTag: true,
    includeNextAssignmentDue: true,
    expandedDNC: false,
    compactColumnWidth: true,
    sheetNameMode: 'date',
    advisorAssignment: {
      enabled: false,
      advisors: []
    }
  });

  // State for View Management: 'settings' | 'processing' | 'done' | 'error'
  const [view, setView] = useState('settings');
  const [settingsView, setSettingsView] = useState('main'); // 'main' | 'tags' | 'assigned' | 'inclusions'
  const [errorMessage, setErrorMessage] = useState('');

  // State for Progress Tracking: { [stepId]: 'pending' | 'active' | 'completed' }
  const [stepStatus, setStepStatus] = useState({});

  // State for Batch Progress (for large datasets)
  // { current, total, phase, tableName } where phase is 'writing' | 'formatting'
  const [batchProgress, setBatchProgress] = useState(null);

  // State for Campus Detection (when campus mode is active)
  const [detectedCampuses, setDetectedCampuses] = useState([]);
  const [campusLoading, setCampusLoading] = useState(false);

  // State for Multi-Campus Processing Progress
  // { campusName: 'pending' | 'active' | 'completed' }
  const [campusStatuses, setCampusStatuses] = useState({});
  const [isMultiCampus, setIsMultiCampus] = useState(false);

  // State for Missing Columns Modal
  // { show: boolean, outreach: boolean, assigned: boolean, pendingPayload: object|null }
  const [missingColumnsModal, setMissingColumnsModal] = useState({ show: false, outreach: false, assigned: false });

  // Track whether the initial settings load has resolved so we can hold the
  // app-level loading state until the toggles reflect saved values.
  const [initialSettingsLoaded, setInitialSettingsLoaded] = useState(false);

  // 'checking' until the Master List sheet check resolves; 'present' lets the
  // normal settings UI render; 'missing' swaps to the status page.
  const [masterListStatus, setMasterListStatus] = useState('checking');

  // Probe for the Master List sheet on mount. Without it there's nothing for
  // the LDA processor to read, so the status page replaces the settings form.
  // Importing data is the chrome extension's job, so this page is informational
  // only — no "create master list" button.
  useEffect(() => {
    let isMounted = true;
    checkMasterListExists()
      .then(exists => {
        if (!isMounted) return;
        setMasterListStatus(exists ? 'present' : 'missing');
      })
      .catch(() => {
        // Match the helper's fail-open behavior: don't block on a check failure.
        if (isMounted) setMasterListStatus('present');
      });
    return () => { isMounted = false; };
  }, []);

  // Load workbook settings on mount
  useEffect(() => {
    let isMounted = true;

    const applyWorkbookSettings = (wb) => {
      if (!isMounted) return;
      if (wb && typeof wb === 'object') {
        setLdaSettings(prev => ({
          daysOut: (wb.daysOut !== undefined && wb.daysOut !== null) ? Number(wb.daysOut) : prev.daysOut,
          includeFailingList: (wb.includeFailingList !== undefined) ? !!wb.includeFailingList : prev.includeFailingList,
          includeAttendanceList: (wb.includeAttendanceList !== undefined) ? !!wb.includeAttendanceList : prev.includeAttendanceList,
          includeLDATag: (wb.includeLDATag !== undefined) ? !!wb.includeLDATag : ((wb.includeLdatTag !== undefined) ? !!wb.includeLdatTag : prev.includeLDATag),
          includeDNCTag: (wb.includeDNCTag !== undefined) ? !!wb.includeDNCTag : ((wb.includeDncTag !== undefined) ? !!wb.includeDncTag : prev.includeDNCTag),
          includeNextAssignmentDue: (wb.includeNextAssignmentDue !== undefined) ? !!wb.includeNextAssignmentDue : prev.includeNextAssignmentDue,
          expandedDNC: (wb.expandedDNC !== undefined) ? !!wb.expandedDNC : prev.expandedDNC,
          compactColumnWidth: (wb.compactColumnWidth !== undefined) ? !!wb.compactColumnWidth : prev.compactColumnWidth,
          sheetNameMode: wb.sheetNameMode || prev.sheetNameMode,
          advisorAssignment: wb.advisorAssignment || prev.advisorAssignment,
        }));
      }
    };

    const loadSettings = (onDone) => {
      if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.refreshAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            applyWorkbookSettings(Office.context.document.settings.get('workbookSettings'));
          }
          if (onDone) onDone();
        });
      } else if (onDone) {
        onDone();
      }
    };

    // Initial load — mark ready only after refreshAsync resolves so the first
    // render the user sees already has saved values applied.
    loadSettings(() => {
      if (isMounted) setInitialSettingsLoaded(true);
    });

    const settingsChangedHandler = () => loadSettings();
    if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
      Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, settingsChangedHandler);
    }

    return () => {
      isMounted = false;
      if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, settingsChangedHandler);
      }
    };
  }, []);

  // Signal ready only after both the initial settings load and the Master List
  // check resolve, so the parent loading screen stays up until we know which
  // view to render (status page vs. settings form).
  useEffect(() => {
    if (initialSettingsLoaded && masterListStatus !== 'checking' && onReady) {
      onReady();
    }
  }, [initialSettingsLoaded, masterListStatus, onReady]);

  // Detect campuses when campus mode is toggled on
  useEffect(() => {
    if (ldaSettings.sheetNameMode === 'campus') {
      setCampusLoading(true);
      detectCampuses()
        .then(campuses => {
          setDetectedCampuses(campuses);
          setCampusLoading(false);
        })
        .catch(() => {
          setDetectedCampuses([]);
          setCampusLoading(false);
        });
    } else {
      setDetectedCampuses([]);
    }
  }, [ldaSettings.sheetNameMode]);

  const handleSettingChange = (key, value) => {
    setLdaSettings((prev) => {
      const next = { ...prev, [key]: value };
      // Persist advisorAssignment to workbook settings
      if (key === 'advisorAssignment') {
        saveAdvisorAssignmentToWorkbook(value);
      }
      return next;
    });
  };

  const saveAdvisorAssignmentToWorkbook = (advisorAssignment) => {
    try {
      if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
        const wb = Office.context.document.settings.get('workbookSettings') || {};
        wb.advisorAssignment = advisorAssignment;
        Office.context.document.settings.set('workbookSettings', wb);
        Office.context.document.settings.saveAsync(() => {});
      }
    } catch (e) {
      console.error('Failed to save advisor assignment:', e);
    }
  };

  // --- Pre-check for missing Outreach/Assigned columns ---
  const handlePreCheck = async () => {
    try {
      const missing = await checkMissingLDAColumns();
      if (missing.outreach || missing.assigned) {
        setMissingColumnsModal({ show: true, outreach: missing.outreach, assigned: missing.assigned });
        return; // Wait for user decision via modal
      }
    } catch (e) {
      console.warn('Column pre-check failed, proceeding anyway:', e);
    }
    handleCreateLDA();
  };

  const handleMissingColumnsConfirm = async () => {
    const columnsToAdd = [];
    if (missingColumnsModal.outreach) columnsToAdd.push('Outreach');
    if (missingColumnsModal.assigned) columnsToAdd.push('Assigned');
    setMissingColumnsModal({ show: false, outreach: false, assigned: false });
    try {
      await addColumnsToMasterList(columnsToAdd);
    } catch (e) {
      console.error('Failed to add columns:', e);
    }
    handleCreateLDA();
  };

  const handleMissingColumnsDismiss = () => {
    setMissingColumnsModal({ show: false, outreach: false, assigned: false });
    handleCreateLDA();
  };

  // --- REAL LOGIC: Trigger the Processor ---
  const handleCreateLDA = async () => {
    console.log('Starting LDA Creation Process with:', ldaSettings);

    const multiCampus = ldaSettings.sheetNameMode === 'campus' && detectedCampuses.length > 1;

    // 1. Switch View
    setView('processing');
    setErrorMessage('');
    setBatchProgress(null);
    setIsMultiCampus(multiCampus);

    // 2. Reset Steps
    const initialStatus = {};
    PROCESS_STEPS.forEach(s => initialStatus[s.id] = 'pending');
    setStepStatus(initialStatus);

    // Initialize campus statuses if multi-campus
    if (multiCampus) {
      const initial = {};
      detectedCampuses.forEach(c => initial[c] = 'pending');
      setCampusStatuses(initial);
    } else {
      setCampusStatuses({});
    }

    try {
        // 3. Call the imported processor function
        await createLDA(
            ldaSettings,
            // Step progress callback
            (stepId, status) => {
                setStepStatus(prev => ({
                    ...prev,
                    [stepId]: status
                }));
                // Clear batch progress when moving to a new step
                if (status === 'active') {
                    setBatchProgress(null);
                }
            },
            // Batch progress callback (for large datasets)
            (current, total, phase, tableName) => {
                setBatchProgress({ current, total, phase, tableName });
            },
            // Campus progress callback (for multi-campus mode)
            multiCampus ? (campusName, campusIndex, totalCampuses, status) => {
                setCampusStatuses(prev => ({
                    ...prev,
                    [campusName]: status
                }));
            } : null
        );

        // 4. Success
        setBatchProgress(null);
        setView('done');

    } catch (error) {
        console.error("LDA Failed:", error);
        setErrorMessage(error.message || "An unexpected error occurred.");
        setBatchProgress(null);
        setView('error');
    }
  };

  const handleReset = () => {
    setView('settings');
    setStepStatus({});
    setErrorMessage('');
    setBatchProgress(null);
    setCampusStatuses({});
    setIsMultiCampus(false);
  };

  if (masterListStatus === 'missing') {
    return <MissingMasterListStatus />;
  }

  return (
    <div className="w-full max-w-2xl mx-auto bg-white rounded-2xl shadow-xl shadow-slate-200/60 border border-white overflow-hidden p-6 transition-all duration-300 min-h-[400px]">

      {/* Header Area */}
      <div className="mb-6 flex items-center justify-between">
        <div>
            <h2 className={`text-2xl font-bold tracking-tight ${view === 'error' ? 'text-red-600' : 'text-slate-800'}`}>
              {view === 'settings' ? 'Create LDA' :
               view === 'done' ? 'Complete' :
               view === 'error' ? 'Error' :
               'Processing...'}
            </h2>
            <p className="text-slate-400 text-sm mt-1">
                {view === 'settings' 
                  ? 'Configure your settings below to generate the report' 
                  : (view === 'done' ? 'Your LDA report has been generated.' : 
                     view === 'error' ? 'Something went wrong.' :
                     'Please wait while we generate your report.')}
            </p>
        </div>
        {/* Back button (only show if not in settings view) */}
        {view !== 'settings' && (
             <button onClick={handleReset} className="text-slate-400 hover:text-slate-600 transition-colors" title="Back to Settings">
                <ArrowLeft className="w-5 h-5" />
             </button>
        )}
      </div>

      <div className="relative">
        {/* VIEW 1: Settings Form */}
        {view === 'settings' && (
          <section className="flex flex-col gap-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <LDASettings
              settings={ldaSettings}
              onSettingChange={handleSettingChange}
              settingsView={settingsView}
              setSettingsView={setSettingsView}
            />

            {settingsView === 'main' && (
              <div className="pt-4 border-t border-slate-100 mt-2">
                <button
                  type="button"
                  className="w-full sm:w-auto bg-[#145F82] hover:bg-[#0f4b66] text-white font-medium px-6 py-2.5 rounded-full shadow-lg shadow-[#145F82]/20 transition-all duration-200 hover:-translate-y-0.5 flex items-center justify-center gap-2"
                  onClick={handlePreCheck}
                >
                  Create LDA
                </button>

                {/* Campus Detection List - fades in when campus mode is active */}
                <div
                  className={`transition-all duration-500 ease-in-out overflow-hidden ${
                    ldaSettings.sheetNameMode === 'campus'
                      ? 'opacity-100 max-h-96 mt-4'
                      : 'opacity-0 max-h-0 mt-0'
                  }`}
                >
                  {campusLoading ? (
                    <div className="flex items-center gap-2 text-slate-400 text-sm p-3">
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>Scanning for campuses...</span>
                    </div>
                  ) : detectedCampuses.length > 1 ? (
                    <div className="bg-slate-50/80 rounded-xl border border-slate-100/80 p-4">
                      <div className="flex items-center gap-2 mb-3">
                        <div className="w-5 h-5 rounded-md bg-[#145F82]/10 flex items-center justify-center">
                          <span className="text-[#145F82] text-xs font-bold">{detectedCampuses.length}</span>
                        </div>
                        <span className="text-sm font-medium text-slate-700">
                          Campuses Detected
                        </span>
                      </div>
                      <div className="space-y-1.5">
                        {detectedCampuses.map(campus => (
                          <div key={campus} className="flex items-center gap-2.5 text-sm text-slate-600 pl-1">
                            <div className="w-1.5 h-1.5 rounded-full bg-[#145F82]/50" />
                            <span>{campus}</span>
                          </div>
                        ))}
                      </div>
                      <p className="text-xs text-slate-400 mt-3 pt-3 border-t border-slate-100">
                        {detectedCampuses.length} separate LDA sheets will be created
                      </p>
                    </div>
                  ) : detectedCampuses.length === 1 ? (
                    <div className="text-sm text-slate-500 p-3 bg-slate-50/80 rounded-xl border border-slate-100/80">
                      1 campus found: <span className="font-medium text-slate-700">{detectedCampuses[0]}</span>
                    </div>
                  ) : ldaSettings.sheetNameMode === 'campus' && !campusLoading ? (
                    <div className="text-sm text-slate-400 italic p-3">
                      No Campus column found in Master List
                    </div>
                  ) : null}
                </div>
              </div>
            )}
          </section>
        )}

        {/* VIEW 2: Processing Checklist */}
        {(view === 'processing' || view === 'done' || view === 'error') && (
           <section className="flex flex-col gap-4 animate-in fade-in slide-in-from-right-8 duration-500">
              <div className="bg-slate-50 rounded-xl border border-slate-100 p-4 space-y-3">
                  {/* Render steps - for multi-campus, show global steps then campus list */}
                  {(isMultiCampus
                    ? PROCESS_STEPS.filter(s => ['validate', 'read', 'filter', 'tags'].includes(s.id))
                    : PROCESS_STEPS
                  ).map((step) => {
                      const status = stepStatus[step.id] || 'pending';
                      const showBatchProgress = (
                          (step.id === 'writing' && batchProgress && batchProgress.phase === 'writing') ||
                          (step.id === 'highlights' && batchProgress && batchProgress.phase === 'highlights') ||
                          (step.id === 'hyperlinks' && batchProgress && batchProgress.phase === 'hyperlinks') ||
                          (step.id === 'read' && batchProgress && batchProgress.phase === 'reading')
                      ) && status === 'active' && batchProgress.total > 1;

                      const isErrored = view === 'error';
                      return (
                          <div key={step.id}>
                              <div className="flex items-center gap-3">
                                  <div className="w-6 flex justify-center shrink-0">
                                      {status === 'pending' && <Circle className="w-4 h-4 text-slate-300" />}
                                      {status === 'active' && (
                                          isErrored
                                              ? <Circle className="w-4 h-4 text-slate-300" />
                                              : <Loader2 className="w-5 h-5 text-[#145F82] animate-spin" />
                                      )}
                                      {status === 'completed' && (
                                          <CheckCircle2 className={`w-5 h-5 animate-in zoom-in duration-300 ${isErrored ? 'text-slate-300' : 'text-emerald-500'}`} />
                                      )}
                                  </div>
                                  <div className={`flex-1 text-sm font-medium transition-colors duration-300 ${
                                      isErrored ? 'text-slate-400' :
                                      status === 'pending' ? 'text-slate-400' :
                                      status === 'active' ? 'text-slate-800' : 'text-emerald-700'
                                  }`}>
                                      {step.label}
                                  </div>
                              </div>

                              {showBatchProgress && (
                                  <div className="ml-9 mt-2 animate-in fade-in slide-in-from-top-2 duration-300">
                                      <div className="flex items-center justify-between text-xs text-slate-500 mb-1">
                                          <span>
                                              {batchProgress.phase === 'reading' ? 'Reading' :
                                               batchProgress.phase === 'writing' ? 'Writing' :
                                               batchProgress.phase === 'highlights' ? 'Highlighting' :
                                               batchProgress.phase === 'hyperlinks' ? 'Hyperlinking' : 'Formatting'}{' '}
                                              {batchProgress.tableName === 'Master_List' ? 'Master List' :
                                               batchProgress.tableName === 'LDA_Table' ? 'LDA' :
                                               batchProgress.tableName === 'Attendance_Table' ? 'Attendance' : 'Failing'} data
                                          </span>
                                          <span className="font-medium">
                                              {batchProgress.current} / {batchProgress.total}
                                          </span>
                                      </div>
                                      <div className="h-1.5 bg-slate-200 rounded-full overflow-hidden">
                                          <div
                                              className="h-full bg-[#145F82] rounded-full transition-all duration-300 ease-out"
                                              style={{ width: `${(batchProgress.current / batchProgress.total) * 100}%` }}
                                          />
                                      </div>
                                  </div>
                              )}
                          </div>
                      );
                  })}

                  {/* Multi-Campus Progress Section */}
                  {isMultiCampus && Object.keys(campusStatuses).length > 0 && (
                      <div className="mt-1 pt-3 border-t border-slate-200/60">
                          <div className="flex items-center gap-2 mb-2.5">
                              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
                                  Campus Reports
                              </span>
                              <span className="text-xs text-slate-400">
                                  ({Object.values(campusStatuses).filter(s => s === 'completed').length} / {Object.keys(campusStatuses).length})
                              </span>
                          </div>
                          <div className="space-y-2">
                              {Object.entries(campusStatuses).map(([campusName, status]) => {
                                  const isErrored = view === 'error';
                                  return (
                                  <div key={campusName} className="flex items-center gap-3 animate-in fade-in duration-300">
                                      <div className="w-6 flex justify-center shrink-0">
                                          {status === 'pending' && <Circle className="w-3.5 h-3.5 text-slate-300" />}
                                          {status === 'active' && (
                                              isErrored
                                                  ? <Circle className="w-3.5 h-3.5 text-slate-300" />
                                                  : <Loader2 className="w-4 h-4 text-[#145F82] animate-spin" />
                                          )}
                                          {status === 'completed' && (
                                              <CheckCircle2 className={`w-4 h-4 animate-in zoom-in duration-300 ${isErrored ? 'text-slate-300' : 'text-emerald-500'}`} />
                                          )}
                                      </div>
                                      <div className={`flex-1 text-sm transition-colors duration-300 ${
                                          isErrored ? 'text-slate-400' :
                                          status === 'pending' ? 'text-slate-400' :
                                          status === 'active' ? 'text-slate-700 font-medium' : 'text-emerald-600'
                                      }`}>
                                          {campusName}
                                      </div>
                                  </div>
                                  );
                              })}
                          </div>
                      </div>
                  )}
              </div>

              {view === 'error' && (
                  <div className="pt-2 animate-in fade-in slide-in-from-bottom-2 duration-500">
                      <div className="bg-red-50 text-red-700 px-4 py-3 rounded-lg text-sm font-medium flex items-start gap-3 border border-red-100">
                          <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                          <span className="flex-1 min-w-0 break-words whitespace-pre-wrap max-h-40 overflow-y-auto">
                              {errorMessage}
                          </span>
                      </div>
                      <button
                        onClick={handleReset}
                        className="mt-4 w-full text-slate-500 hover:text-slate-700 text-sm font-medium hover:underline"
                      >
                        Try Again
                      </button>
                  </div>
              )}

              {view === 'done' && (
                  <div className="pt-2 animate-in fade-in slide-in-from-bottom-2 duration-500">
                      <div className="bg-emerald-50 text-emerald-700 px-4 py-3 rounded-lg text-sm font-medium flex items-center justify-center border border-emerald-100">
                          {isMultiCampus
                            ? `${Object.keys(campusStatuses).length} LDA Sheets Generated Successfully!`
                            : 'LDA Generated Successfully!'
                          }
                      </div>
                      <button
                        onClick={handleReset}
                        className="mt-4 w-full text-slate-500 hover:text-slate-700 text-sm font-medium hover:underline"
                      >
                        Start Over
                      </button>
                  </div>
              )}
           </section>
        )}
      </div>

      {/* Missing Columns Modal */}
      {missingColumnsModal.show && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-2xl shadow-2xl p-6 max-w-sm w-full mx-4">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center">
                <AlertCircle className="w-5 h-5 text-amber-600" />
              </div>
              <h3 className="text-lg font-semibold text-slate-800">Missing Columns</h3>
            </div>
            <p className="text-sm text-slate-600 mb-3">
              The following columns were not found on the Master List:
            </p>
            <ul className="text-sm text-slate-700 mb-4 space-y-1 pl-4">
              {missingColumnsModal.outreach && (
                <li className="flex items-center gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-amber-400 inline-block" />
                  <span><strong>Outreach</strong> — used for retention comments</span>
                </li>
              )}
              {missingColumnsModal.assigned && (
                <li className="flex items-center gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-amber-400 inline-block" />
                  <span><strong>Assigned</strong> — used for advisor assignments</span>
                </li>
              )}
            </ul>
            <p className="text-sm text-slate-500 mb-5">
              Would you like to add {missingColumnsModal.outreach && missingColumnsModal.assigned ? 'them' : 'it'} to the Master List?
            </p>
            <div className="flex gap-3 justify-end">
              <button
                type="button"
                className="px-4 py-2 text-sm font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
                onClick={handleMissingColumnsDismiss}
              >
                Skip
              </button>
              <button
                type="button"
                className="px-4 py-2 text-sm font-medium text-white bg-[#145F82] hover:bg-[#0f4b66] rounded-lg transition-colors"
                onClick={handleMissingColumnsConfirm}
              >
                Add & Continue
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// --- Sub-Components ---

function MissingMasterListStatus() {
  return (
    <div className="w-full max-w-2xl mx-auto bg-white rounded-2xl shadow-xl shadow-slate-200/60 border border-white overflow-hidden p-6 min-h-[400px] flex items-center justify-center">
      <div className="text-center max-w-sm animate-in fade-in slide-in-from-bottom-2 duration-500">
        <div className="w-16 h-16 mx-auto mb-5 rounded-2xl bg-amber-50 border border-amber-100 flex items-center justify-center">
          <ClipboardList className="w-8 h-8 text-amber-500" strokeWidth={1.75} />
        </div>
        <h2 className="text-xl font-bold tracking-tight text-slate-800 mb-2">
          Missing Master List Sheet
        </h2>
        <p className="text-sm text-slate-500 leading-relaxed mb-4">
          Create LDA reads from a <strong className="text-slate-700">Master List</strong> sheet
          in this workbook, and one wasn't found. Without it there's no student
          data to build the report from.
        </p>
        <div className="text-left text-xs text-slate-500 leading-relaxed bg-slate-50/80 border border-slate-100 rounded-xl px-4 py-3">
          Use the Student Retention Kit <strong className="text-slate-700">Chrome
          extension</strong> to send your roster into Excel, then reopen
          <strong className="text-slate-700"> Create LDA</strong> to continue.
        </div>
      </div>
    </div>
  );
}

function LDASettings({ settings, onSettingChange, settingsView, setSettingsView }) {

  const handleToggle = (key) => {
    onSettingChange(key, !settings[key]);
  };

  const handleInputChange = (key, value) => {
    onSettingChange(key, value);
  };

  if (settingsView === 'tags') {
    return (
      <div className="flex flex-col gap-4 w-full animate-in fade-in slide-in-from-right-4 duration-300">
        <button
          type="button"
          onClick={() => setSettingsView('main')}
          className="flex items-center gap-1 text-slate-400 hover:text-slate-600 text-sm font-medium transition-colors w-fit"
        >
          <ArrowLeft className="w-4 h-4" />
          Back
        </button>

        <ToggleRow
          key="toggle-lda-tag"
          label={
            <span className="inline-flex items-center gap-2">
              Include
              <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-orange-200 text-orange-800">LDA</span>
            </span>
          }
          tooltip="Mark students who already have an LDA follow-up logged in Student History."
          isOn={settings.includeLDATag}
          onToggle={() => handleToggle('includeLDATag')}
        />

        <ToggleRow
          key="toggle-dnc-tag"
          label={
            <span className="inline-flex items-center gap-2">
              Include
              <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-red-200 text-black">DNC</span>
            </span>
          }
          tooltip="Mark students flagged as Do Not Contact in Student History."
          isOn={settings.includeDNCTag}
          onToggle={() => handleToggle('includeDNCTag')}
        />

        <div className="mt-2">
          <span className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">
            Preview
          </span>
          <div className="mt-2 space-y-2">
            <CommentPreview
              tagLabel="LDA"
              tagClass="bg-orange-200 text-orange-800"
              bgClass="bg-orange-100"
              borderClass="border-orange-400"
              text="Student plans to submit by Friday."
            />
            <CommentPreview
              tagLabel="DNC"
              tagClass="bg-red-200 text-black"
              bgClass="bg-red-100"
              borderClass="border-red-600"
              text="Student requested not to be contacted."
            />
          </div>
        </div>
      </div>
    );
  }

  if (settingsView === 'assigned') {
    return (
      <AssignedSettings
        settings={settings}
        onSettingChange={onSettingChange}
        onBack={() => setSettingsView('main')}
      />
    );
  }

  if (settingsView === 'inclusions') {
    return (
      <div className="flex flex-col gap-4 w-full animate-in fade-in slide-in-from-right-4 duration-300">
        <button
          type="button"
          onClick={() => setSettingsView('main')}
          className="flex items-center gap-1 text-slate-400 hover:text-slate-600 text-sm font-medium transition-colors w-fit"
        >
          <ArrowLeft className="w-4 h-4" />
          Back
        </button>

        <ToggleRow
          key="toggle-failing-list"
          label="Include Failing List"
          tooltip="Add a second table below the LDA table listing students with failing grades."
          isOn={settings.includeFailingList}
          onToggle={() => handleToggle('includeFailingList')}
        />

        <ToggleRow
          key="toggle-attendance-list"
          label="Include Attendance List"
          tooltip="Add a third table listing students with attendance below 60%."
          isOn={settings.includeAttendanceList}
          onToggle={() => handleToggle('includeAttendanceList')}
        />

        <ToggleRow
          key="toggle-next-assignment-due"
          label="Include Next Assignment Due"
          tooltip="Use each student's next assignment due date in retention messaging."
          isOn={settings.includeNextAssignmentDue}
          onToggle={() => handleToggle('includeNextAssignmentDue')}
        />
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-4 w-full">
      <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
        <div className="flex items-center gap-2">
          <span className="text-slate-700 font-medium text-sm">Sheet Name</span>
          <Info
            title="Name the new sheet by today's date or by the student's campus."
            className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600"
          />
        </div>
        <div className="relative flex bg-slate-200 rounded-full p-0.5" style={{ width: 130 }}>
          <button
            type="button"
            onClick={() => handleInputChange('sheetNameMode', 'date')}
            className={`relative z-10 flex-1 py-1 text-xs font-medium rounded-full text-center transition-colors duration-200 ${
              settings.sheetNameMode === 'date'
                ? 'text-white'
                : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            Date
          </button>
          <button
            type="button"
            onClick={() => handleInputChange('sheetNameMode', 'campus')}
            className={`relative z-10 flex-1 py-1 text-xs font-medium rounded-full text-center transition-colors duration-200 ${
              settings.sheetNameMode === 'campus'
                ? 'text-white'
                : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            Campus
          </button>
          <span
            className="absolute top-0.5 bottom-0.5 rounded-full bg-[#145F82] transition-all duration-200"
            style={{
              width: 'calc(50% - 2px)',
              left: settings.sheetNameMode === 'date' ? '2px' : 'calc(50%)',
            }}
          />
        </div>
      </div>

      <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
        <div className="flex items-center gap-2">
          <label htmlFor="daysOut" className="text-slate-700 font-medium text-sm">
            Days Out
          </label>
          <div className="group relative">
             <Info
                title="Minimum days since last attendance before a student appears on the LDA table."
                className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600"
             />
          </div>
        </div>
        <input
          id="daysOut"
          type="number"
          value={settings.daysOut}
          onChange={(e) => handleInputChange('daysOut', Number(e.target.value))}
          className="w-24 border border-slate-200 bg-white rounded-lg px-3 py-1.5 text-right text-sm focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:border-[#145F82] transition-all"
        />
      </div>

      {(() => {
        const inclusionsOn = [
          settings.includeFailingList,
          settings.includeAttendanceList,
          settings.includeNextAssignmentDue,
        ].filter(Boolean).length;
        return (
          <button
            type="button"
            onClick={() => setSettingsView('inclusions')}
            className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors"
          >
            <div className="flex items-center gap-2">
              <span className="text-slate-700 font-medium text-sm">Inclusions</span>
              {inclusionsOn > 0 && (
                <span className="text-[10px] font-semibold text-white bg-[#145F82] px-1.5 py-0.5 rounded-full">
                  {inclusionsOn}
                </span>
              )}
            </div>
            <ChevronRight className="w-4 h-4 text-slate-400" />
          </button>
        );
      })()}

      <button
        type="button"
        onClick={() => setSettingsView('tags')}
        className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors"
      >
        <div className="flex items-center gap-2">
          <span className="text-slate-700 font-medium text-sm">Tags</span>
          <Info
            title="Configure LDA and Do Not Contact (DNC) tag columns."
            className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600"
          />
        </div>
        <ChevronRight className="w-4 h-4 text-slate-400" />
      </button>

      <button
        type="button"
        onClick={() => setSettingsView('assigned')}
        className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors"
      >
        <div className="flex items-center gap-2">
          <span className="text-slate-700 font-medium text-sm">Assigned</span>
          {settings.advisorAssignment?.enabled && (
            <span className="text-[10px] font-semibold text-white bg-[#145F82] px-1.5 py-0.5 rounded-full">ON</span>
          )}
        </div>
        <ChevronRight className="w-4 h-4 text-slate-400" />
      </button>
    </div>
  );
}

function CommentPreview({ tagLabel, tagClass, bgClass, borderClass, text }) {
  return (
    <div className={`${bgClass} ${borderClass} p-3 rounded-lg shadow-sm border-l-4`}>
      <div className="flex items-center gap-1 mb-1.5">
        <span className={`px-2 py-0.5 text-xs font-semibold rounded-full ${tagClass}`}>
          {tagLabel}
        </span>
      </div>
      <p className="text-sm text-gray-800">{text}</p>
      <div className="text-[10px] text-gray-500 mt-2 pt-2 border-t border-gray-200/70 flex justify-between">
        <span>Advisor</span>
        <span>Today</span>
      </div>
    </div>
  );
}

function ToggleRow({ label, isOn, onToggle, tooltip }) {
  return (
    <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
      <div className="flex items-center gap-2">
        <span className="text-slate-700 font-medium text-sm">{label}</span>
        {tooltip && <Info title={tooltip} className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />}
      </div>

      <button
        type="button"
        onClick={onToggle}
        className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:ring-offset-1 ${
          isOn ? 'bg-[#145F82]' : 'bg-slate-200'
        }`}
      >
        <span
          className={`${
            isOn ? 'translate-x-6' : 'translate-x-1'
          } inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform duration-200`}
        />
      </button>
    </div>
  );
}

// --- Assigned Settings (Advisor Auto-Assignment) ---

const DEFAULT_COLORS = ['#ADD8E6', '#FFDAB9', '#D4EDDA', '#E8D5F5', '#FFE0B2', '#B2DFDB', '#F8BBD0', '#C5CAE9'];
const DAY_KEYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
const DAY_LABELS = ['SU', 'M', 'T', 'W', 'TH', 'F', 'SA'];

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function getTodayDayKey() {
  return DAY_KEYS[new Date().getDay()];
}

function isAdvisorExcludedToday(advisor) {
  const todayKey = getTodayDayKey();
  return (advisor.excludeDays || []).includes(todayKey);
}

function AssignedSettings({ settings, onSettingChange, onBack }) {
  const assignment = settings.advisorAssignment || { enabled: false, advisors: [] };
  const advisors = assignment.advisors || [];

  const [programVersions, setProgramVersions] = useState([]);
  const [pvLoading, setPvLoading] = useState(false);
  const [distribution, setDistribution] = useState([]);
  const [distLoading, setDistLoading] = useState(false);
  const [filterModalAdvisor, setFilterModalAdvisor] = useState(null);
  const [editingName, setEditingName] = useState(null);
  const [editNameValue, setEditNameValue] = useState('');

  useEffect(() => {
    setPvLoading(true);
    detectProgramVersions()
      .then(pvs => { setProgramVersions(pvs); setPvLoading(false); })
      .catch(() => { setProgramVersions([]); setPvLoading(false); });
  }, []);

  // Filter out day-excluded advisors for prediction
  const activeAdvisors = advisors.filter(a => !isAdvisorExcludedToday(a));

  useEffect(() => {
    // Defer recalculation until the filter modal is closed
    if (filterModalAdvisor) return;
    if (!assignment.enabled || activeAdvisors.length === 0) {
      setDistribution([]);
      return;
    }
    setDistLoading(true);
    const timer = setTimeout(() => {
      predictAdvisorDistribution(settings, activeAdvisors)
        .then(d => {
          // Debug log: full LDA settings, advisor settings, and per-advisor student lists
          console.group('%c[LDA Auto-Assign Debug]', 'color: #6366f1; font-weight: bold');

          console.group('LDA Report Settings');
          console.log('Days Out Threshold:', settings.daysOut ?? 5);
          console.log('Include Failing List:', settings.includeFailingList ? 'YES' : 'NO');
          console.log('Include Attendance List:', settings.includeAttendanceList ? 'YES' : 'NO');
          console.log('Even Split:', assignment.evenSplit ? 'ON' : 'OFF');
          console.log('Total Advisors:', advisors.length, '| Active Today:', activeAdvisors.length);
          console.groupEnd();

          console.group('Advisor Filter Settings');
          console.table(activeAdvisors.map(a => ({
            Name: a.name,
            'Days Out Min': a.daysOutMin ?? '(none)',
            'Days Out Max': a.daysOutMax ?? '(none)',
            'Program Versions': a.programVersions?.length ? a.programVersions.join(', ') : '(any)',
            'List Preference': a.listPreference?.length ? a.listPreference.join(', ') : '(any)',
            'Exclude Days': a.excludeDays?.length ? a.excludeDays.join(', ') : '(none)',
          })));
          console.groupEnd();

          console.group('Predicted Distribution');
          console.table(d.map(r => ({ Name: r.name, Count: r.count })));
          const totalStudents = d.reduce((sum, r) => sum + r.count, 0);
          console.log('Total Students:', totalStudents);
          console.groupEnd();

          // Per-advisor student lists
          for (const r of d) {
            if (r.students && r.students.length > 0) {
              console.group(`${r.name} (${r.count} students)`);
              console.table(r.students.map(s => ({
                'Program Version': s.programVersion || '(empty)',
                'Days Out': s.daysOut,
                'List Type': s.listType,
              })));
              console.groupEnd();
            }
          }

          console.groupEnd();
          setDistribution(d); setDistLoading(false);
        })
        .catch(() => { setDistribution([]); setDistLoading(false); });
    }, 300);
    return () => clearTimeout(timer);
  }, [assignment.enabled, assignment.evenSplit, advisors, settings.daysOut, settings.includeFailingList, settings.includeAttendanceList, filterModalAdvisor]);

  const updateAssignment = (updates) => {
    onSettingChange('advisorAssignment', { ...assignment, ...updates });
  };

  const addAdvisor = () => {
    const colorIdx = advisors.length % DEFAULT_COLORS.length;
    const newAdvisor = {
      id: generateId(),
      name: `Advisor ${advisors.length + 1}`,
      color: DEFAULT_COLORS[colorIdx],
      programVersions: [],
      listPreference: [],
      daysOutMin: null,
      daysOutMax: null,
      excludeDays: []
    };
    updateAssignment({ advisors: [...advisors, newAdvisor] });
    setEditingName(newAdvisor.id);
    setEditNameValue(newAdvisor.name);
  };

  const removeAdvisor = (id) => {
    updateAssignment({ advisors: advisors.filter(a => a.id !== id) });
    if (filterModalAdvisor?.id === id) setFilterModalAdvisor(null);
  };

  const updateAdvisor = (id, updates) => {
    updateAssignment({
      advisors: advisors.map(a => a.id === id ? { ...a, ...updates } : a)
    });
  };

  const confirmEditName = (id) => {
    if (editNameValue.trim()) {
      updateAdvisor(id, { name: editNameValue.trim() });
    }
    setEditingName(null);
  };

  const totalStudents = distribution.reduce((sum, d) => sum + d.count, 0);
  const todayKey = getTodayDayKey();

  return (
    <div className="flex flex-col gap-4 w-full animate-in fade-in slide-in-from-right-4 duration-300">
      <button
        type="button"
        onClick={onBack}
        className="flex items-center gap-1 text-slate-400 hover:text-slate-600 text-sm font-medium transition-colors w-fit"
      >
        <ArrowLeft className="w-4 h-4" />
        Back
      </button>

      {/* Enable toggle */}
      <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
        <div className="flex items-center gap-2">
          <span className="text-slate-700 font-medium text-sm">Auto-Assign Advisors</span>
        </div>
        <button
          type="button"
          onClick={() => updateAssignment({ enabled: !assignment.enabled })}
          className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:ring-offset-1 ${
            assignment.enabled ? 'bg-[#145F82]' : 'bg-slate-200'
          }`}
        >
          <span
            className={`${
              assignment.enabled ? 'translate-x-6' : 'translate-x-1'
            } inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform duration-200`}
          />
        </button>
      </div>

      {/* Even Split toggle (only visible when enabled) */}
      <div className={`transition-all duration-300 overflow-hidden ${assignment.enabled ? 'opacity-100 max-h-20' : 'opacity-0 max-h-0'}`}>
        <div className="flex items-center justify-between p-3 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:border-slate-200 transition-colors">
          <div className="flex items-center gap-2">
            <span className="text-slate-700 font-medium text-sm">Even Split</span>
            <div className="group relative">
              <Info className="w-4 h-4 text-slate-400 cursor-help hover:text-slate-600" />
              <div className="invisible group-hover:visible absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-52 p-2 bg-slate-800 text-white text-[10px] rounded-lg shadow-lg z-50 leading-relaxed">
                Redistributes students so each advisor gets as close to the same number as possible. Moves students with the lowest days out first.
              </div>
            </div>
          </div>
          <button
            type="button"
            onClick={() => updateAssignment({ evenSplit: !assignment.evenSplit })}
            className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:ring-offset-1 ${
              assignment.evenSplit ? 'bg-[#145F82]' : 'bg-slate-200'
            }`}
          >
            <span
              className={`${
                assignment.evenSplit ? 'translate-x-6' : 'translate-x-1'
              } inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform duration-200`}
            />
          </button>
        </div>
      </div>

      {/* Advisor list (only visible when enabled) */}
      <div className={`transition-all duration-300 overflow-hidden ${assignment.enabled ? 'opacity-100 max-h-[2000px]' : 'opacity-0 max-h-0'}`}>
        <div className={`flex flex-col gap-3 ${advisors.length > 6 ? 'max-h-[360px] overflow-y-auto pr-1' : ''}`}>
          {advisors.map((advisor) => {
            const excluded = isAdvisorExcludedToday(advisor);
            return (
              <div
                key={advisor.id}
                className={`rounded-xl border overflow-hidden transition-all duration-200 ${
                  excluded ? 'border-red-200/60 opacity-60' : 'border-slate-100/80'
                }`}
              >
                <div
                  className={`flex items-center gap-2 p-2.5 transition-colors ${excluded ? 'bg-red-50/50' : 'hover:bg-slate-50/50'}`}
                  style={excluded ? {} : { backgroundColor: advisor.color + '33' }}
                >
                  <div
                    className={`w-6 h-6 rounded-md border shrink-0 relative overflow-hidden ${excluded ? 'border-red-200 grayscale' : 'border-black/10'}`}
                    style={{ backgroundColor: excluded ? '#e5e7eb' : advisor.color }}
                  >
                    {!excluded && (
                      <input
                        type="color"
                        value={advisor.color}
                        onChange={(e) => updateAdvisor(advisor.id, { color: e.target.value })}
                        className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                        title="Change color"
                      />
                    )}
                  </div>

                  {editingName === advisor.id ? (
                    <input
                      type="text"
                      value={editNameValue}
                      onChange={(e) => setEditNameValue(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === 'Escape') e.target.blur(); }}
                      onBlur={() => confirmEditName(advisor.id)}
                      autoFocus
                      className="flex-1 min-w-0 text-sm font-medium text-slate-700 bg-white border border-slate-200 rounded-md px-2 py-0.5 focus:outline-none focus:ring-1 focus:ring-[#145F82]/30"
                    />
                  ) : (
                    <span className={`text-sm font-medium truncate flex-1 min-w-0 ${excluded ? 'text-slate-400 line-through' : 'text-slate-700'}`}>{advisor.name}</span>
                  )}

                  {excluded && (
                    <span className="text-[10px] text-red-400 bg-red-50 px-1.5 py-0.5 rounded-full shrink-0 font-medium">
                      Off today
                    </span>
                  )}

                  {(() => {
                    const fc = (advisor.programVersions?.length || 0) + (advisor.listPreference?.length || 0) + (advisor.daysOutMin != null ? 1 : 0) + (advisor.daysOutMax != null ? 1 : 0) + (advisor.excludeDays?.length || 0);
                    return !excluded && fc > 0 ? (
                      <span className="text-[10px] text-slate-500 bg-white/60 px-1.5 py-0.5 rounded-full shrink-0">
                        {fc} filter{fc !== 1 ? 's' : ''}
                      </span>
                    ) : null;
                  })()}

                  <button
                    onClick={(e) => { e.stopPropagation(); setFilterModalAdvisor(advisor); }}
                    className="text-slate-400 hover:text-[#145F82] p-0.5 shrink-0 transition-colors"
                    title="Edit filters"
                  >
                    <Pencil className="w-3.5 h-3.5" />
                  </button>

                  <button
                    onClick={(e) => { e.stopPropagation(); removeAdvisor(advisor.id); }}
                    className="text-slate-300 hover:text-red-400 p-0.5 shrink-0 transition-colors"
                    title="Remove advisor"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            );
          })}

          {/* Add advisor button */}
          <button
            type="button"
            onClick={addAdvisor}
            className="flex items-center justify-center gap-1.5 p-2.5 rounded-xl border border-dashed border-slate-200 text-slate-400 hover:text-slate-600 hover:border-slate-300 transition-colors text-sm"
          >
            <Plus className="w-4 h-4" />
            Add Advisor
          </button>

          {/* Pie Chart Distribution Preview */}
          {advisors.length > 0 && (
            <div className="mt-2 p-4 bg-slate-50/50 rounded-xl border border-slate-100/50">
              <p className="text-xs font-medium text-slate-500 mb-3">Predicted Distribution</p>
              {distLoading ? (
                <div className="flex items-center justify-center gap-2 text-slate-400 text-xs py-4">
                  <Loader2 className="w-4 h-4 animate-spin" />
                  Calculating...
                </div>
              ) : totalStudents === 0 ? (
                <p className="text-xs text-slate-400 italic text-center py-4">
                  {activeAdvisors.length === 0 ? 'All advisors are excluded today' : 'No students match current filters'}
                </p>
              ) : (
                <div className="flex items-center gap-4">
                  <AdvisorPieChart distribution={distribution} size={120} />
                  <div className="flex flex-col gap-1.5 flex-1 min-w-0">
                    {distribution.map(d => (
                      <div key={d.id} className="flex items-center gap-2 text-xs">
                        <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: d.color }} />
                        <span className="text-slate-600 truncate flex-1">{d.name}</span>
                        <span className="text-slate-500 font-medium shrink-0">
                          {d.count} <span className="text-slate-400 font-normal">({totalStudents > 0 ? Math.round((d.count / totalStudents) * 100) : 0}%)</span>
                        </span>
                      </div>
                    ))}
                    <div className="border-t border-slate-100 pt-1 mt-0.5">
                      <div className="flex items-center justify-between text-xs">
                        <span className="text-slate-500 font-medium">Total</span>
                        <span className="text-slate-700 font-semibold">{totalStudents} students</span>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Filter Modal */}
      {filterModalAdvisor && (
        <AdvisorFilterModal
          advisor={filterModalAdvisor}
          programVersions={programVersions}
          pvLoading={pvLoading}
          onUpdate={(updates) => updateAdvisor(filterModalAdvisor.id, updates)}
          onClose={() => setFilterModalAdvisor(null)}
        />
      )}

    </div>
  );
}

// --- Advisor Filter Modal ---

function AdvisorFilterModal({ advisor: initialAdvisor, programVersions, pvLoading, onUpdate, onClose }) {
  const [localAdvisor, setLocalAdvisor] = useState(initialAdvisor);
  const [editingModalName, setEditingModalName] = useState(false);
  const [modalNameValue, setModalNameValue] = useState(initialAdvisor.name);

  const localUpdate = (updates) => {
    setLocalAdvisor(prev => ({ ...prev, ...updates }));
  };

  const handleClose = () => {
    // Compute diff and apply all changes at once on close
    const changes = {};
    for (const key of Object.keys(localAdvisor)) {
      if (key === 'id' || key === 'color') continue;
      if (JSON.stringify(localAdvisor[key]) !== JSON.stringify(initialAdvisor[key])) {
        changes[key] = localAdvisor[key];
      }
    }
    if (Object.keys(changes).length > 0) {
      onUpdate(changes);
    }
    onClose();
  };

  const confirmModalName = () => {
    if (modalNameValue.trim()) {
      localUpdate({ name: modalNameValue.trim() });
    }
    setEditingModalName(false);
  };

  const advisor = localAdvisor;

  const togglePV = (pv) => {
    const pvs = advisor.programVersions || [];
    localUpdate({ programVersions: pvs.includes(pv) ? pvs.filter(p => p !== pv) : [...pvs, pv] });
  };

  const toggleList = (listType) => {
    const prefs = advisor.listPreference || [];
    localUpdate({ listPreference: prefs.includes(listType) ? prefs.filter(p => p !== listType) : [...prefs, listType] });
  };

  const toggleDay = (dayKey) => {
    const days = advisor.excludeDays || [];
    localUpdate({ excludeDays: days.includes(dayKey) ? days.filter(d => d !== dayKey) : [...days, dayKey] });
  };

  const setRange = (field, value) => {
    const parsed = value === '' ? null : Number(value);
    localUpdate({ [field]: isNaN(parsed) ? null : parsed });
  };

  const todayKey = getTodayDayKey();

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={handleClose}>
      <div className="absolute inset-0 bg-black/30 backdrop-blur-sm" />
      <div
        className="relative bg-white rounded-2xl shadow-xl w-full max-w-sm max-h-[80vh] overflow-y-auto animate-in fade-in zoom-in-95 duration-200"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="sticky top-0 bg-white border-b border-slate-100 p-4 flex items-center justify-between rounded-t-2xl z-10">
          <div className="flex items-center gap-2 flex-1 min-w-0">
            <div className="w-4 h-4 rounded shrink-0" style={{ backgroundColor: advisor.color }} />
            {editingModalName ? (
              <input
                type="text"
                value={modalNameValue}
                onChange={(e) => setModalNameValue(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === 'Escape') e.target.blur(); }}
                onBlur={confirmModalName}
                autoFocus
                className="flex-1 min-w-0 text-sm font-semibold text-slate-800 bg-white border border-slate-200 rounded-md px-2 py-0.5 focus:outline-none focus:ring-1 focus:ring-[#145F82]/30"
              />
            ) : (
              <div className="flex items-center gap-1.5 min-w-0">
                <h3 className="text-sm font-semibold text-slate-800 truncate">{advisor.name}</h3>
                <button
                  onClick={() => { setModalNameValue(advisor.name); setEditingModalName(true); }}
                  className="text-slate-300 hover:text-slate-500 p-0.5 shrink-0"
                  title="Edit name"
                >
                  <Pencil className="w-3 h-3" />
                </button>
              </div>
            )}
          </div>
          <button onClick={handleClose} className="text-slate-400 hover:text-slate-600 transition-colors p-1 shrink-0">
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 space-y-5">

          {/* Exclude Days */}
          <div>
            <p className="text-xs font-medium text-slate-500 mb-2">Exclude Days</p>
            <p className="text-[10px] text-slate-400 mb-2">Select days this advisor is excluded from the LDA report</p>
            <div className="flex gap-1">
              {DAY_KEYS.map((dayKey, i) => {
                const isExcluded = (advisor.excludeDays || []).includes(dayKey);
                const isToday = dayKey === todayKey;
                return (
                  <button
                    key={dayKey}
                    type="button"
                    onClick={() => toggleDay(dayKey)}
                    className={`w-9 h-9 rounded-lg text-[11px] font-semibold border transition-all duration-150 ${
                      isExcluded
                        ? 'bg-red-50 border-red-200 text-red-500'
                        : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'
                    } ${isToday ? 'ring-2 ring-offset-1 ring-[#145F82]/30' : ''}`}
                  >
                    {DAY_LABELS[i]}
                  </button>
                );
              })}
            </div>
          </div>

          {/* List Preference */}
          <div>
            <p className="text-xs font-medium text-slate-500 mb-2">List Preference</p>
            <div className="flex gap-1.5">
              {[{ key: 'lda', label: 'LDA' }, { key: 'failing', label: 'Failing' }, { key: 'attendance', label: 'Attendance' }].map(({ key, label }) => {
                const isSelected = (advisor.listPreference || []).includes(key);
                return (
                  <button
                    key={key}
                    type="button"
                    onClick={() => toggleList(key)}
                    className={`text-xs px-2.5 py-1.5 rounded-lg border transition-all duration-150 ${
                      isSelected
                        ? 'border-[#145F82] bg-[#145F82]/10 text-[#145F82] font-medium'
                        : 'border-slate-200 text-slate-500 hover:border-slate-300'
                    }`}
                  >
                    {label}
                  </button>
                );
              })}
            </div>
            {(advisor.listPreference || []).length === 0 && (
              <p className="text-[10px] text-slate-400 mt-1.5 italic">No preference — receives from all lists</p>
            )}
          </div>

          {/* Days Out Range */}
          <div>
            <p className="text-xs font-medium text-slate-500 mb-2">Days Out Range</p>
            <div className="flex items-center gap-2">
              <input
                type="number"
                placeholder="Min"
                value={advisor.daysOutMin ?? ''}
                onChange={(e) => setRange('daysOutMin', e.target.value)}
                className="w-20 border border-slate-200 bg-white rounded-lg px-2 py-1.5 text-xs text-center focus:outline-none focus:ring-1 focus:ring-[#145F82]/30 focus:border-[#145F82]"
              />
              <span className="text-slate-400 text-xs">to</span>
              <input
                type="number"
                placeholder="Max"
                value={advisor.daysOutMax ?? ''}
                onChange={(e) => setRange('daysOutMax', e.target.value)}
                className="w-20 border border-slate-200 bg-white rounded-lg px-2 py-1.5 text-xs text-center focus:outline-none focus:ring-1 focus:ring-[#145F82]/30 focus:border-[#145F82]"
              />
            </div>
            {advisor.daysOutMin == null && advisor.daysOutMax == null && (
              <p className="text-[10px] text-slate-400 mt-1.5 italic">No range — all days out values</p>
            )}
          </div>

          {/* Program Version Filters */}
          <div>
            <p className="text-xs font-medium text-slate-500 mb-2">Program Versions</p>
            {pvLoading ? (
              <div className="flex items-center gap-2 text-slate-400 text-xs py-2">
                <Loader2 className="w-3 h-3 animate-spin" /> Loading...
              </div>
            ) : programVersions.length === 0 ? (
              <p className="text-xs text-slate-400 italic">No ProgramVersion column found</p>
            ) : (
              <div className="flex flex-wrap gap-1.5 max-h-44 overflow-y-auto">
                {programVersions.map(pv => {
                  const isSelected = (advisor.programVersions || []).includes(pv);
                  return (
                    <button
                      key={pv}
                      type="button"
                      onClick={() => togglePV(pv)}
                      className={`text-xs px-2 py-1 rounded-full border transition-all duration-150 ${
                        isSelected
                          ? 'border-[#145F82] bg-[#145F82]/10 text-[#145F82] font-medium'
                          : 'border-slate-200 text-slate-500 hover:border-slate-300'
                      }`}
                    >
                      {pv}
                    </button>
                  );
                })}
              </div>
            )}
            {!pvLoading && programVersions.length > 0 && (advisor.programVersions || []).length === 0 && (
              <p className="text-[10px] text-slate-400 mt-1.5 italic">No filter — all program versions</p>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="sticky bottom-0 bg-white border-t border-slate-100 p-4 rounded-b-2xl">
          <button
            onClick={handleClose}
            className="w-full bg-[#145F82] hover:bg-[#0f4b66] text-white font-medium px-4 py-2 rounded-lg text-sm transition-colors"
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
}

// --- SVG Pie Chart ---

function AdvisorPieChart({ distribution, size = 120 }) {
  const total = distribution.reduce((sum, d) => sum + d.count, 0);
  if (total === 0) return null;

  const radius = size / 2;
  const center = radius;
  let cumulativeAngle = -Math.PI / 2; // Start from top

  const slices = distribution.filter(d => d.count > 0).map(d => {
    const angle = (d.count / total) * 2 * Math.PI;
    const startAngle = cumulativeAngle;
    cumulativeAngle += angle;
    const endAngle = cumulativeAngle;

    const x1 = center + radius * Math.cos(startAngle);
    const y1 = center + radius * Math.sin(startAngle);
    const x2 = center + radius * Math.cos(endAngle);
    const y2 = center + radius * Math.sin(endAngle);
    const largeArc = angle > Math.PI ? 1 : 0;

    const path = `M ${center} ${center} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;

    return { ...d, path };
  });

  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} className="shrink-0">
      {slices.length === 1 ? (
        <circle cx={center} cy={center} r={radius} fill={slices[0].color} />
      ) : (
        slices.map((s, i) => (
          <path key={i} d={s.path} fill={s.color} stroke="white" strokeWidth="1.5" />
        ))
      )}
      {/* Center hole for donut effect */}
      <circle cx={center} cy={center} r={radius * 0.45} fill="white" />
      <text x={center} y={center - 4} textAnchor="middle" className="text-sm font-bold fill-slate-700">
        {total}
      </text>
      <text x={center} y={center + 10} textAnchor="middle" className="text-[8px] fill-slate-400">
        students
      </text>
    </svg>
  );
}