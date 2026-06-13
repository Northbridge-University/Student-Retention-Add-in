import React, { useState, useEffect, useRef } from 'react';
import Comment, { COMMENT_TAGS, CommentSkeleton } from '../Parts/Comment';
import NewComment from '../Modal/NewCommentModal';
import { formatExcelDate, normalizeKeys, formatTimestamp } from '../../utility/Conversion';
import { /* insertRow, editRow, */ } from '../../utility/ExcelAPI'; // removed direct insert/edit usage
import { addComment, deleteComment, generateCommentID, resolveStudentIdentity } from '../../utility/EditStudentHistory';
import { Folder, FolderOpen } from 'lucide-react';
import { toast } from 'react-toastify';

// Module-level global loading flag + event bus so other modules can trigger History to show skeletons.
let _globalHistoryLoading = false;
const _globalHistoryLoadingBus = new EventTarget();
export function setHistoryLoading(enable = true) {
  _globalHistoryLoading = !!enable;
  _globalHistoryLoadingBus.dispatchEvent(new CustomEvent('change', { detail: _globalHistoryLoading }));
}
export function getHistoryLoading() {
  return _globalHistoryLoading;
}

// Add styles constant
const styles = `
  @keyframes fadeInDrop {
    from {
      opacity: 0;
      transform: translateY(-24px);
    }
    to {
      opacity: 1;
      transform: translateY(0);
    }
  }
  .animate-fadein {
    animation: fadeInDrop 0.4s cubic-bezier(0.4,0,0.2,1);
  }
  /* See-through scrollbar styles */
  #history-content {
    scrollbar-width: thin; /* Firefox */
    scrollbar-color: rgba(0,0,0,0.15) rgba(0,0,0,0.03);
    position: relative; /* allow overlay absolute positioning */
  }
  #history-content::-webkit-scrollbar {
    width: 8px;
    background: transparent;
  }
  #history-content::-webkit-scrollbar-thumb {
    background: rgba(0,0,0,0.15);
    border-radius: 8px;
  }
  #history-content::-webkit-scrollbar-track {
    background: rgba(0,0,0,0.03);
  }

  /* Skeleton overlay that sits above comments and fades out
     Remove padding so layout/spacing exactly matches the normal comment list */
  .skeleton-overlay {
    position: absolute;
    inset: 0;
    z-index: 20;
    background: rgba(255,255,255,0.9);
    display: flex;
    flex-direction: column;
    padding: 0; /* keep items aligned with comments */
    /* Allow skeleton content to overfill the overlay; do NOT show overlay scrollbars */
    overflow: visible;
    transition: opacity 400ms cubic-bezier(0.4,0,0.2,1), visibility 400ms;
  }
`;

function getMonthFromTimestamp(ts) {
  // Try to parse Excel date or ISO string
  if (!ts) return null;
  // If ts is a number, treat as Excel serial
  if (!isNaN(ts)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(excelEpoch.getTime() + ts * 86400000);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }
  // If ts is ISO string
  const date = new Date(ts);
  if (!isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }
  return null;
}

function getCurrentMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function getCurrentYear() {
  return new Date().getFullYear();
}

function getMonthYearLabel(monthStr) {
  // monthStr: "YYYY-MM"
  const [year, month] = monthStr.split('-');
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const monthIdx = parseInt(month, 10) - 1;
  return monthNames[monthIdx];
}

// add helper to parse timestamp to milliseconds
function parseTimestampMs(ts) {
  if (ts == null) return 0;
  // numeric (Excel serial) or numeric string
  if (!isNaN(ts)) {
    const n = Number(ts);
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return excelEpoch.getTime() + n * 86400000;
  }
  const d = new Date(ts);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function StudentHistory({ history, student, reload }) {
  console.log('Rendering StudentHistory with history:', history, 'and student:', student);
  // Local state for history to allow adding new comments
  const [localHistory, setLocalHistory] = useState(Array.isArray(history) ? history : []);
  // External/global loading flag: initialized from the exported global state and updated via bus
  const [externalLoading, setExternalLoading] = useState(getHistoryLoading());
  // Background processing ready flag: remains false while normalization/sorting/grouping runs
  const [ready, setReady] = useState(false);
  // Processed data (prepared in background)
  const [processed, setProcessed] = useState({
    pinnedComments: [],
    currentMonthComments: [],
    currentYearMonths: [],       // months from current year (not in year folders)
    yearGroups: {},               // { year: { monthGroups: {}, sortedMonths: [] } }
    sortedYears: []               // sorted list of previous years
  });
  // isLoading true while parent hasn't provided history OR external loading was triggered
  const isLoading = externalLoading || history == null || !ready;

  // When loading starts, immediately clear any previously-processed comments/folders
  // so the old list cannot flash while skeletons are shown.
  useEffect(() => {
    if (isLoading) {
      // NOTE: keep processed comments intact so they render behind the skeleton overlay.
      // Previously this cleared processed state which removed comments while skeletons displayed.
      // We preserve collapsedFolders so the UI behind the overlay remains consistent.
      // ...no-op to intentionally keep comments visible behind skeletons...
    }
  }, [isLoading]);

  // UI state (declare early so hooks order is stable)
  const [showSearch, setShowSearch] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showNewComment, setShowNewComment] = useState(false);
  // status message removed in favor of react-toastify toasts

  // Collapsed state for folders — initialize from computed months below (stable hook call)
  const [collapsedFolders, setCollapsedFolders] = useState({});
  const collapsedRef = useRef(collapsedFolders);
  useEffect(() => { collapsedRef.current = collapsedFolders; }, [collapsedFolders]);

  // Collapsed state for year folders
  const [collapsedYears, setCollapsedYears] = useState({});
  const collapsedYearsRef = useRef(collapsedYears);
  useEffect(() => { collapsedYearsRef.current = collapsedYears; }, [collapsedYears]);

  // Sync localHistory with history prop if it changes
  useEffect(() => {
    setLocalHistory(Array.isArray(history) ? history : []);
    // When a real history prop arrives, clear any external/global loading flag so component leaves skeleton state.
    if (history != null) {
      // update local external state and global flag
      setExternalLoading(false);
      setHistoryLoading(false);
    }
    // clear/leave toasts as-is when student changes
}, [history]);

// Subscribe to global loading bus so external callers can toggle loading
useEffect(() => {
  const handler = (e) => {
    setExternalLoading(!!e.detail);
  };
  _globalHistoryLoadingBus.addEventListener('change', handler);
  return () => _globalHistoryLoadingBus.removeEventListener('change', handler);
}, []);
 
   // Background processing: normalize, filter, split pinned/unpinned, sort and group by month.
   // This is done off-render so isLoading remains true until the work finishes and order is stable.
   useEffect(() => {
     // When there's no history (empty array), stop loading and show the "no comments" message.
     // This prevents skeletons from showing indefinitely when there truly are no comments.
     if (!Array.isArray(localHistory) || localHistory.length === 0) {
       // clear processed data and collapsed folders
       setProcessed({
         pinnedComments: [],
         currentMonthComments: [],
         currentYearMonths: [],
         yearGroups: {},
         sortedYears: []
       });
       setCollapsedFolders({});
       setCollapsedYears({});
       // mark ready so isLoading becomes false and skeleton overlay fades away
       setReady(true);
       // clear external/global loading flag as a courtesy so other modules know loading finished
       setExternalLoading(false);
       setHistoryLoading(false);
       return;
     }

     setReady(false);
     let idleId;
     const doWork = () => {
       try {
         const normalizedHistory = localHistory.map(entry => normalizeKeys(entry || {}));
         const filteredHistory = normalizedHistory.filter(
           entry =>
             !searchTerm ||
             (entry.comment && entry.comment.toLowerCase().includes(searchTerm.toLowerCase()))
         );

         const isEntryPinned = (entry) => {
           if (!entry.tag) return false;
           const entryTags = entry.tag.split(',').map(t => t.trim());
           return entryTags.some(tagLabel => {
             const tagObj = COMMENT_TAGS.find(t => t.label === tagLabel);
             if (tagObj) {
               if (tagObj.pinned) return true;
               if (Array.isArray(tagObj.subtags)) {
                 if (tagObj.subtags.some(subtag => subtag.pinned)) return true;
               }
             }
             for (const tag of COMMENT_TAGS) {
               if (Array.isArray(tag.subtags)) {
                 const subtagObj = tag.subtags.find(subtag => subtag.label === tagLabel);
                 if (subtagObj && subtagObj.pinned) return true;
               }
             }
             return false;
           });
         };

         const pinnedComments = filteredHistory
           .filter(isEntryPinned)
           .slice()
           .sort((a, b) => parseTimestampMs(b.timestamp) - parseTimestampMs(a.timestamp));

         const unpinnedComments = filteredHistory
           .filter(entry => !isEntryPinned(entry))
           .slice()
           .sort((a, b) => parseTimestampMs(b.timestamp) - parseTimestampMs(a.timestamp));

         const currentMonth = getCurrentMonth();
         const currentYear = getCurrentYear();
         const currentMonthComments = [];
         const currentYearMonthGroups = {};
         const yearGroups = {};

         // Group unpinned comments by year and month
         unpinnedComments.forEach(entry => {
           const month = getMonthFromTimestamp(entry.timestamp);
           if (!month) return;

           const [year] = month.split('-');
           const yearNum = parseInt(year, 10);

           if (month === currentMonth) {
             // Current month - no folder
             currentMonthComments.push(entry);
           } else if (yearNum === currentYear) {
             // Other months in current year - month folder at top level
             if (!currentYearMonthGroups[month]) currentYearMonthGroups[month] = [];
             currentYearMonthGroups[month].push(entry);
           } else {
             // Previous years - organize into year folders containing month folders
             if (!yearGroups[yearNum]) {
               yearGroups[yearNum] = {};
             }
             if (!yearGroups[yearNum][month]) yearGroups[yearNum][month] = [];
             yearGroups[yearNum][month].push(entry);
           }
         });

         // Sort current year months (most recent first)
         const sortedCurrentYearMonths = Object.keys(currentYearMonthGroups).sort((a, b) => b.localeCompare(a));

         // Sort years (most recent first)
         const sortedYears = Object.keys(yearGroups)
           .map(y => parseInt(y, 10))
           .sort((a, b) => b - a);

         // For each year, sort its months
         const processedYearGroups = {};
         sortedYears.forEach(year => {
           const months = Object.keys(yearGroups[year]).sort((a, b) => b.localeCompare(a));
           processedYearGroups[year] = {
             monthGroups: yearGroups[year],
             sortedMonths: months
           };
         });

         // Merge collapsed state for current year months
         const prevMonths = collapsedRef.current || {};
         const mergedMonths = { ...prevMonths };
         sortedCurrentYearMonths.forEach(month => {
           if (!(month in mergedMonths)) mergedMonths[month] = true;
         });
         // Also handle months within year folders
         sortedYears.forEach(year => {
           Object.keys(yearGroups[year]).forEach(month => {
             const key = `${year}-${month}`;
             if (!(key in mergedMonths)) mergedMonths[key] = true;
           });
         });
         // Remove months that no longer exist
         Object.keys(mergedMonths).forEach(m => {
           const isCurrentYearMonth = sortedCurrentYearMonths.includes(m);
           // Year-folder months are keyed by the composite `${year}-${month}`,
           // so compare against that form (not the bare month) or they'd be
           // dropped here and lose their collapsed-by-default state.
           const isYearMonth = sortedYears.some(year => {
             return yearGroups[year] && Object.keys(yearGroups[year]).some(month => `${year}-${month}` === m);
           });
           if (!isCurrentYearMonth && !isYearMonth) delete mergedMonths[m];
         });

         // Merge collapsed state for years
         const prevYears = collapsedYearsRef.current || {};
         const mergedYears = { ...prevYears };
         sortedYears.forEach(year => {
           if (!(year in mergedYears)) mergedYears[year] = true;
         });
         // Remove years that no longer exist
         Object.keys(mergedYears).forEach(y => {
           if (!sortedYears.includes(parseInt(y, 10))) delete mergedYears[y];
         });

         // Update state
         setCollapsedFolders(mergedMonths);
         setCollapsedYears(mergedYears);
         setProcessed({
           pinnedComments,
           currentMonthComments,
           currentYearMonths: sortedCurrentYearMonths.map(month => ({
             month,
             comments: currentYearMonthGroups[month]
           })),
           yearGroups: processedYearGroups,
           sortedYears
         });
         setReady(true);
       } catch (err) {
         // If processing fails for some reason, mark ready so UI can attempt rendering (and show error via toast if needed)
         setCollapsedFolders({});
         setCollapsedYears({});
         setProcessed({
           pinnedComments: [],
           currentMonthComments: [],
           currentYearMonths: [],
           yearGroups: {},
           sortedYears: []
         });
         setReady(true);
       }
     };

     if (typeof window !== 'undefined' && 'requestIdleCallback' in window) {
       idleId = window.requestIdleCallback(doWork, { timeout: 500 });
     } else {
       idleId = setTimeout(doWork, 0);
     }

     return () => {
       if (typeof window !== 'undefined' && 'cancelIdleCallback' in window && idleId && typeof idleId === 'number') {
         window.cancelIdleCallback(idleId);
       } else if (idleId) {
         clearTimeout(idleId);
       }
     };
   }, [localHistory, searchTerm]);
 
   function toggleFolder(month) {
     setCollapsedFolders(prev => ({
       ...prev,
       [month]: !prev[month]
     }));
   }

   function toggleYear(year) {
     setCollapsedYears(prev => ({
       ...prev,
       [year]: !prev[year]
     }));
   }

   // Wrapper that delegates to the shared addComment. Manual comments always
   // insert a new row — they never opt into Outreach dedupe, since silently
   // overwriting another advisor's text would be a data-loss bug.
   async function addCommentToHistory(comment, tag = '') {
     if (!comment) return false;
     try {
       const { studentId, studentName } = resolveStudentIdentity(student);
       await addComment(String(comment), tag, undefined, studentId, studentName);
       await reload();
       toast.success('Comment saved');
       return true;
     } catch (err) {
       toast.error('Failed to save comment');
       return false;
     }
   }

   async function deleteCommentFromHistory(commentID) {
     await deleteComment(commentID);
     await reload();
     toast.error('Comment deleted');
   }

   async function saveCommentFromHistory(commentID) {
     await reload();
     toast.success('Comment changes saved');
   }
 
   // add refs to support long-loading toast behavior
   const isLoadingRef = useRef(externalLoading || history == null || !ready);
   const longLoadTimerRef = useRef(null);
   const longLoadToastRef = useRef(null);

   // keep isLoadingRef up-to-date so the timeout callback can read current state
   useEffect(() => {
     isLoadingRef.current = externalLoading || history == null || !ready;
   }, [externalLoading, history, ready]);

   // Show a single toast if loading persists longer than 10 seconds
   useEffect(() => {
     const isLoadingNow = externalLoading || history == null || !ready;
     if (isLoadingNow) {
       // start timer if not already started
       if (!longLoadTimerRef.current) {
         longLoadTimerRef.current = setTimeout(() => {
           // check the ref for current loading state
           if (isLoadingRef.current && !longLoadToastRef.current) {
             longLoadToastRef.current = toast.warn('Uh oh. This is taking longer than usual', { autoClose: 6000 });
           }
         }, 3000); // 3 seconds
       }
     } else {
       // clear any pending timer and reset toast ref (do not dismiss existing toasts automatically)
       if (longLoadTimerRef.current) {
         clearTimeout(longLoadTimerRef.current);
         longLoadTimerRef.current = null;
       }
       longLoadToastRef.current = null;
     }

     // cleanup on unmount
     return () => {
       if (longLoadTimerRef.current) {
         clearTimeout(longLoadTimerRef.current);
         longLoadTimerRef.current = null;
       }
     };
   }, [externalLoading, history, ready]);

   // Contact rate: percentage of Contacted entries out of all Contacted + Outreach entries
   const contactRate = React.useMemo(() => {
     let contacted = 0, total = 0;
     for (const raw of (localHistory || [])) {
       const entry = normalizeKeys(raw || {});
       const tags = (entry.tag || '').toLowerCase();
       const hasContacted = tags.includes('contacted');
       const hasOutreach = tags.includes('outreach');
       if (hasContacted || hasOutreach) {
         total++;
         if (hasContacted) contacted++;
       }
     }
     console.log('[ContactRate]', { contacted, total, rate: total > 0 ? Math.round((contacted / total) * 100) : null });
     return total > 0 ? Math.round((contacted / total) * 100) : null;
   }, [localHistory]);

   return (
     <div>
       <style>{styles}</style>
       {/* History Header */}
       <div className="sticky-header space-y-4 pl-2">
         <div className="flex justify-between items-center">
           <h3 className="text-lg font-bold text-gray-800">History</h3>
           <div className="flex items-center space-x-2">
             {contactRate !== null && (
               <div
                 className={`w-8 h-8 rounded-full shadow-lg flex items-center justify-center text-[10px] font-bold ${
                   contactRate >= 50 ? 'bg-green-100 text-green-800' :
                   contactRate >= 25 ? 'bg-yellow-100 text-yellow-800' :
                   'bg-red-100 text-red-800'
                 }`}
                 title={`Contact rate: ${contactRate}%`}
               >
                 {contactRate}%
               </div>
             )}
             <button
               id="search-history-button"
               className="bg-gray-600 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-gray-700"
               onClick={() => setShowSearch(v => !v)}
               aria-label="Search history"
             >
               <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
             </button>
             <button
               id="add-comment-button"
               className="bg-blue-600 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-blue-700"
               onClick={() => setShowNewComment(v => !v)}
               aria-label="Add comment"
             >
               <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 6v6m0 0v6m0-6h6m-6 0H6"></path></svg>
             </button>
           </div>
         </div>

         <div
           id="search-container"
           className={`${showSearch ? '' : 'hidden'} space-y-2`}
         >
           <div id="tag-filter-container" className="flex flex-wrap items-center gap-2 pb-2 border-b border-gray-200">
             
           </div>
           <div className="relative">
             <input
               type="text"
               id="search-input"
               className="w-full p-2 pl-8 border rounded-md"
               placeholder="Search comments..."
               value={searchTerm}
               onChange={e => setSearchTerm(e.target.value)}
               autoFocus={showSearch}
             />
             <div className="absolute inset-y-0 left-0 pl-2 flex items-center pointer-events-none">
               <svg className="h-5 w-5 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
             </div>
             <button
               id="clear-search-button"
               className="absolute inset-y-0 right-0 pr-3 flex items-center text-gray-500 hover:text-gray-700"
               type="button"
               onClick={() => setSearchTerm('')}
               aria-label="Clear search"
             >
               ×
             </button>
           </div>
         </div>

         {/* Move new comment box to NewComment component */}
         <NewComment
           show={showNewComment}
           onClose={() => setShowNewComment(false)}
           addCommentToHistory={addCommentToHistory}
         />

         {/* status toasts are handled by the app-level ToastContainer in App.jsx */}
       </div>
       {/* End History Header */}
 
       <div
         id="history-content"
         className="overflow-y-auto"
         style={{
           height: showNewComment || showSearch
             ? 'calc(105vh - 460px)'
             : 'calc(100vh - 260px)'
         }}
       >
         {/* Always render comments (or empty message) — skeletons are shown in a top overlay */}
         <ul className="space-y-4">
           {(!Array.isArray(localHistory) || localHistory.length === 0) ? (
             <li className="p-3 bg-gray-100 rounded-lg shadow-sm relative">
               <p className="text-sm text-gray-800">Student has no comments yet</p>
             </li>
           ) : (
             <>
               {/* Pinned comments first (regardless of timestamp) */}
               {processed.pinnedComments.map((entry, index) => (
                 <Comment
                   key={`pinned-${index}`}
                   entry={entry}
                   searchTerm={searchTerm}
                   index={index}
                   delete={deleteCommentFromHistory}
                   save={saveCommentFromHistory}
                 />
               ))}
               {/* Current month comments (not in a folder, only unpinned) */}
               {[...processed.currentMonthComments].map((entry, idx) => (
                 <Comment
                   key={`currentmonth-${idx}`}
                   entry={entry}
                   searchTerm={searchTerm}
                   index={idx}
                   delete={deleteCommentFromHistory}
                   save={saveCommentFromHistory}
                 />
               ))}
               {/* Other months from current year in collapsible folders (only unpinned) */}
               {(processed.currentYearMonths || []).map(({ month, comments }) => (
                 <li key={month} className="bg-gray-50 rounded-lg shadow-sm p-2">
                   <div
                     className="flex items-center font-semibold text-gray-700 mb-2 cursor-pointer select-none"
                     onClick={() => toggleFolder(month)}
                     style={{ userSelect: 'none' }}
                     aria-label={collapsedFolders[month] ? "Expand" : "Collapse"}
                     tabIndex={0}
                     role="button"
                     onKeyDown={e => {
                       if (e.key === 'Enter' || e.key === ' ') toggleFolder(month);
                     }}
                   >
                     <span className="mr-2 text-gray-500" style={{ fontSize: '1.2em', lineHeight: '1' }}>
                       {collapsedFolders[month]
                         ? <Folder size={20} strokeWidth={2} />
                         : <FolderOpen size={20} strokeWidth={2} />
                       }
                     </span>
                     {getMonthYearLabel(month)}
                   </div>
                   {!collapsedFolders[month] && (
                     <ul className="space-y-2">
                       {[...(comments || [])].map((entry, idx) => (
                         <Comment
                           key={`month-${month}-${idx}`}
                           entry={entry}
                           searchTerm={searchTerm}
                           index={idx}
                           delete={deleteCommentFromHistory}
                           save={saveCommentFromHistory}
                         />
                       ))}
                     </ul>
                   )}
                 </li>
               ))}
               {/* Previous years in collapsible year folders */}
               {(processed.sortedYears || []).map(year => (
                 <li key={year} className="bg-gray-100 rounded-lg shadow-md p-2 mb-2">
                   <div
                     className="flex items-center font-bold text-gray-800 mb-2 cursor-pointer select-none"
                     onClick={() => toggleYear(year)}
                     style={{ userSelect: 'none' }}
                     aria-label={collapsedYears[year] ? "Expand" : "Collapse"}
                     tabIndex={0}
                     role="button"
                     onKeyDown={e => {
                       if (e.key === 'Enter' || e.key === ' ') toggleYear(year);
                     }}
                   >
                     <span className="mr-2 text-gray-600" style={{ fontSize: '1.2em', lineHeight: '1' }}>
                       {collapsedYears[year]
                         ? <Folder size={22} strokeWidth={2.5} />
                         : <FolderOpen size={22} strokeWidth={2.5} />
                       }
                     </span>
                     {year}
                   </div>
                   {!collapsedYears[year] && processed.yearGroups[year] && (
                     <ul className="space-y-2 ml-2">
                       {processed.yearGroups[year].sortedMonths.map(month => (
                         <li key={month} className="bg-gray-50 rounded-lg shadow-sm p-2">
                           <div
                             className="flex items-center font-semibold text-gray-700 mb-2 cursor-pointer select-none"
                             onClick={() => toggleFolder(`${year}-${month}`)}
                             style={{ userSelect: 'none' }}
                             aria-label={collapsedFolders[`${year}-${month}`] ? "Expand" : "Collapse"}
                             tabIndex={0}
                             role="button"
                             onKeyDown={e => {
                               if (e.key === 'Enter' || e.key === ' ') toggleFolder(`${year}-${month}`);
                             }}
                           >
                             <span className="mr-2 text-gray-500" style={{ fontSize: '1.2em', lineHeight: '1' }}>
                               {collapsedFolders[`${year}-${month}`]
                                 ? <Folder size={20} strokeWidth={2} />
                                 : <FolderOpen size={20} strokeWidth={2} />
                               }
                             </span>
                             {getMonthYearLabel(month)}
                           </div>
                           {!collapsedFolders[`${year}-${month}`] && (
                             <ul className="space-y-2">
                               {[...(processed.yearGroups[year].monthGroups[month] || [])].map((entry, idx) => (
                                 <Comment
                                   key={`year-${year}-month-${month}-${idx}`}
                                   entry={entry}
                                   searchTerm={searchTerm}
                                   index={idx}
                                   delete={deleteCommentFromHistory}
                                   save={saveCommentFromHistory}
                                 />
                               ))}
                             </ul>
                           )}
                         </li>
                       ))}
                     </ul>
                   )}
                 </li>
               ))}
             </>
           )}
         </ul>

         {/* Skeleton overlay sits above the comments and fades away when loading completes */}
         <div
           className="skeleton-overlay"
           style={{
             opacity: isLoading ? 1 : 0,
             pointerEvents: isLoading ? 'auto' : 'none',
             visibility: isLoading ? 'visible' : 'visible' /* keep in DOM for smooth fade; pointerEvents handles interaction */
           }}
           aria-hidden={!isLoading}
         >
           <ul className="space-y-4">
             {Array.from({ length: 5 }).map((_, i) => (
               <CommentSkeleton key={`skeleton-${i}`} />
             ))}
           </ul>
         </div>
       </div>
     </div>
   );
 }
 export default StudentHistory;


