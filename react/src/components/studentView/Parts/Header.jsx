// Timestamp: 2025-12-11 11:20 AM | Version: 1.2.2
import React from 'react';
import { User } from 'lucide-react';
import { formatName } from '../../utility/Conversion';
import BounceAnimation from '../../utility/BounceAnimation';
import DaysOutModal from '../Modal/DaysOutModal';
import { getRiskIndexConfig, loadLdaHistory, countRiskEpisodesForStudent } from '../../createLDA/riskIndex';

function StudentHeader({ student, selectedRowCount = 1 }) {
  // Use a fallback student object to prevent errors if the prop is null or undefined
  const safeStudent = student || {};

  // If multiple rows are selected, show a count instead of individual student details
  const isMultipleSelected = selectedRowCount > 1;

  // Use formatName only if name is "Last, First"
  const studentNameRaw = safeStudent.StudentName || 'Select a Student';
  const studentName = studentNameRaw.includes(',') ? formatName(studentNameRaw) : studentNameRaw;

  // Get assigned value; if missing, it remains undefined/null (so we can hide the pill)
  const assignedTo = safeStudent.Assigned;

  const daysOut = safeStudent.DaysOut ?? '--'; // Using nullish coalescing to allow 0

  // --- Risk Index: how many separate times this student hit the risk
  // threshold (default 14 days out), from the LDAHistory document setting
  // written by Create LDA. Shown as a badge on the Days Out tile when > 0.
  const riskStudentId = safeStudent.ID ?? safeStudent.StudentNumber ?? null;
  const risk = React.useMemo(() => {
    if (riskStudentId === null || riskStudentId === undefined || riskStudentId === '') return null;
    try {
      const cfg = getRiskIndexConfig();
      if (!cfg.enabled) return null;
      const count = countRiskEpisodesForStudent(loadLdaHistory(), String(riskStudentId).trim(), cfg.threshold);
      return count > 0 ? { count, threshold: cfg.threshold } : null;
    } catch {
      return null;
    }
  }, [riskStudentId]);
  // Determine background color for Days Out
  let daysOutBg = 'bg-gray-200';
  let daysOutText = 'text-gray-800';
  let daysOutLabelText = 'text-gray-500';
  if (typeof daysOut === 'number') {
    if (daysOut === 0) {
      daysOutBg = 'bg-green-300'; // Brighter green for 0 days out
      daysOutText = 'text-black';
      daysOutLabelText = 'text-black';
    } else if (daysOut >= 14) {
      daysOutBg = 'bg-red-400';
      daysOutText = 'text-black';
      daysOutLabelText = 'text-black';
    } else if (daysOut >= 10) {
      daysOutBg = 'bg-red-300';
      daysOutText = 'text-black';
      daysOutLabelText = 'text-black';
    } else if (daysOut >= 5) {
      daysOutBg = 'bg-yellow-200';
      daysOutText = 'text-gray-800';
      daysOutLabelText = 'text-gray-500';
    } else {
      daysOutBg = 'bg-green-200';
      daysOutText = 'text-gray-800';
      daysOutLabelText = 'text-gray-500';
    }
  }
  // Parse grade as a number if possible, handling "%" sign
  let grade = safeStudent.Grade ?? null; // use null when not provided
  let parsedGrade = grade;
  if (grade !== null && typeof grade === 'string' && grade.trim() !== '') {
    const match = grade.match(/^(\d+(?:\.\d+)?)\s*%?$/);
    if (match) {
      parsedGrade = Number(match[1]);
    }
  }
  // Determine background color for Grade
  let gradeBg = 'bg-gray-200';
  let gradeText = 'text-gray-800';
  let gradeTextLabel = 'text-gray-500';
  let gradeDisplay = grade;
  if (typeof parsedGrade === 'number' && !isNaN(parsedGrade)) {
    // Round to nearest integer
    const roundedGrade = Math.round(parsedGrade);
    if (roundedGrade >= 90) {
      gradeBg = 'bg-green-300';
      gradeText = 'text-black';
      gradeTextLabel = 'text-black';
    } else if (roundedGrade >= 70) {
      gradeBg = 'bg-green-200';
      gradeText = 'text-gray-800';
      gradeTextLabel = 'text-gray-500';
    } else if (roundedGrade >= 60) {
      gradeBg = 'bg-yellow-200';
      gradeText = 'text-gray-800';
      gradeTextLabel = 'text-gray-500';
    } else {
      gradeBg = 'bg-red-300';
      gradeText = 'text-black';
      gradeTextLabel = 'text-black';
    }
    gradeDisplay = `${roundedGrade}%`;
  }
  // If no grade was provided, show 'No' so the UI reads "No" + "Grade"
  if (grade === null || (typeof grade === 'string' && grade.trim() === '')) {
    gradeDisplay = '--';
    // keep muted styling for missing grade
    gradeBg = 'bg-gray-200';
    gradeText = 'text-gray-800';
    gradeTextLabel = 'text-gray-500';
  }
  
  // --- Gender-based avatar background ---
  const gender = (safeStudent.Gender || '').toLowerCase();
  const avatarBg =
    gender === 'male'
      ? { background: '#1278FF', color: '#FFFFFF' } // light blue, dark text
      : gender === 'female'
        ? { background: '#ed72b5', color: '#FFFFFF' } // pink, dark pink text
        : { background: '#6b7280', color: '#FFFFFF' };   // default gray, white text

  const gradebookUrl = safeStudent.Gradebook;

  // Helper to determine if a student is "new" based on ExpectedStartDate.
  // Returns true if ExpectedStartDate (or common variants) is a valid date and is within the past 31 days.
  const IsStudentNew = (studentObj) => {
    if (!studentObj) return false;
    // Accept several possible property names
    const dateVal =
      studentObj.ExpectedStartDate ??
      studentObj.expectedStartDate ??
      studentObj.ExpectedStart ??
      studentObj.startDate;
    if (!dateVal) return false;
    const parsed = Date.parse(dateVal);
    if (isNaN(parsed)) return false;
    const now = Date.now();
    const msInDay = 1000 * 60 * 60 * 24;
    const daysOld = Math.floor((now - parsed) / msInDay);
    // Consider "new" if started within the last 31 days (0..31)
    return daysOld >= 0 && daysOld <= 31;
  };

  // Compute a user-friendly formatted ExpectedStartDate (Month Day, Year) for the tooltip.
  const expectedStartRaw =
    safeStudent.ExpectedStartDate ??
    safeStudent.expectedStartDate ??
    safeStudent.ExpectedStart ??
    safeStudent.startDate ??
    null;
  let expectedStartDisplay = null;
  if (expectedStartRaw) {
    const parsed = Date.parse(expectedStartRaw);
    if (!isNaN(parsed)) {
      const d = new Date(parsed);
      expectedStartDisplay = d.toLocaleDateString('en-US', {
        month: 'long',
        day: 'numeric',
        year: 'numeric'
      }); // "October 17, 2025"
    }
  }
  
  // Determine if the student should show the "NEW" tag.
  // Use ExpectedStartDate-based detection first, fall back to explicit flags.
  const isNew =
    IsStudentNew(safeStudent) ||
    !!(
      safeStudent.IsNew === true ||
      safeStudent.isNew === true ||
      safeStudent.New === true ||
      safeStudent.newStudent === true
    );

  // --- NEW: determine if the student is on Hold (string "Yes", case-insensitive) ---
  const isHold = String(safeStudent.Hold ?? safeStudent.hold ?? '').toLowerCase() === 'yes';

  // --- Determine profile picture URL (accept string or object with Value/value) ---
  const getProfilePicUrl = (s) => {
    const val = s?.ProfilePicture ?? s?.profilePicture ?? s?.ProfileImage ?? s?.ProfilePic;
    if (!val) return null;
    if (typeof val === 'string' && val.trim()) return val.trim();
    if (typeof val === 'object' && val !== null) {
      const candidate = val.Value ?? val.value ?? val.url ?? val.Url;
      return (typeof candidate === 'string' && candidate.trim()) ? candidate.trim() : null;
    }
    return null;
  };
  const profilePicUrl = getProfilePicUrl(safeStudent);

  // Track if the profile image failed to load so we can show initials instead
  const [imageError, setImageError] = React.useState(false);

  // Helper to check if gradebookUrl is a valid URL
  const isValidGradebookUrl = typeof gradebookUrl === "string" && /^https?:\/\/\S+$/i.test(gradebookUrl);

  // --- Handler to open gradebook ---
  const openGradebook = () => {
    if (!isValidGradebookUrl) return;
    console.log("Gradebook URL clicked:", gradebookUrl);
    if (window.Office && window.Office.context && window.Office.context.ui && window.Office.context.ui.openBrowserWindow) {
      window.Office.context.ui.openBrowserWindow(gradebookUrl);
    } else {
      window.open(gradebookUrl, '_blank');
    }
  };

  const [showNoLink, setShowNoLink] = React.useState(false);
  const [showDaysModal, setShowDaysModal] = React.useState(false);
  const [bounce, setBounce] = React.useState(false);
  const [showEnlargedPhoto, setShowEnlargedPhoto] = React.useState(false);

  // --- Dynamic Text Resizing Logic ---
  const nameRef = React.useRef(null);
  const [nameTextSize, setNameTextSize] = React.useState('text-lg');

  // Reset to default large size whenever the student name changes
  React.useLayoutEffect(() => {
    setNameTextSize('text-lg');
  }, [studentName]);

  // Check for overflow after render. If overflowing at text-lg, shrink to text-sm.
  React.useLayoutEffect(() => {
    if (nameTextSize === 'text-lg' && nameRef.current) {
      // scrollHeight > clientHeight indicates content is being truncated/clamped
      if (nameRef.current.scrollHeight > nameRef.current.clientHeight) {
        setNameTextSize('text-sm');
      }
    }
  }, [nameTextSize, studentName]);

  // Render simplified view for multiple selections
  if (isMultipleSelected) {
    return (
      <div className="p-4 bg-white border-b border-gray-200">
        <div className="flex items-center justify-center space-x-2">
          <div className="w-12 h-12 rounded-full bg-blue-500 flex items-center justify-center text-white text-xl font-bold shrink-0">
            {selectedRowCount}
          </div>
          <h2 className="text-lg font-bold text-gray-800">
            Students Selected
          </h2>
        </div>
      </div>
    );
  }

  return (
    <div className="p-4 bg-white border-b border-gray-200">
      {/* BounceAnimation injects bounce CSS */}
      <BounceAnimation />
      <div className="flex items-center justify-between space-x-4 min-w-0">
        {/* Left side: Avatar and Name */}
        <div className="flex items-center space-x-4 min-w-0">
          <button
            type="button"
            className={`relative w-12 h-12 rounded-full flex items-center justify-center text-xl font-bold shrink-0 focus:outline-none${bounce ? " bounce" : ""}`}
            style={avatarBg}
            onClick={() => {
              if (profilePicUrl && !imageError) {
                setShowEnlargedPhoto(true);
              } else {
                setBounce(true);
                setTimeout(() => setBounce(false), 500);
              }
            }}
            aria-label={profilePicUrl && !imageError ? "View profile photo" : "Bounce avatar"}
            // Only show a generic title on the avatar if there is no expected start date.
            title={isNew && !expectedStartDisplay ? "New student" : undefined}
          >
            {/* Render profile image if available and hasn't errored; otherwise show a generic avatar icon */}
            {profilePicUrl && !imageError ? (
              <img
                src={profilePicUrl}
                alt={studentName ? `${studentName} profile` : 'Student profile'}
                onError={() => setImageError(true)}
                className="w-full h-full object-cover rounded-full"
                // Prevent the image from stealing pointer events for overlays
                style={{ display: 'block' }}
              />
            ) : (
              <User className="w-7 h-7" strokeWidth={2} aria-hidden="true" />
            )}

            {isNew && (
              <span
                aria-label={`New student${expectedStartDisplay ? `: ${expectedStartDisplay}` : ''}`}
                role="status"
                title={expectedStartDisplay || 'New student'}
                // Allow the badge itself to receive pointer events so its title is shown on hover.
                style={{
                  position: 'absolute',
                  bottom: 0,
                  right: 0,
                  transform: 'translate(25%, 25%)',
                  background: '#10B981', // emerald-500
                  color: '#ffffff',
                  borderRadius: 9999,
                  fontSize: 10,
                  padding: '2px 6px',
                  fontWeight: 700,
                  boxShadow: '0 1px 2px rgba(0,0,0,0.12)',
                  pointerEvents: 'auto',
                  lineHeight: 1,
                  zIndex: 10
                }}
                tabIndex={0}
               >
                 NEW
               </span>
             )}
            {isHold && (
              <span
                aria-label="Hold"
                role="status"
                title="Hold"
                style={{
                  position: 'absolute',
                  bottom: 0,
                  left: 0,
                  transform: 'translate(-25%, 25%)',
                  background: '#ef4444', // red-500
                  color: '#ffffff',
                  borderRadius: 9999,
                  fontSize: 10,
                  padding: '2px 6px',
                  fontWeight: 700,
                  boxShadow: '0 1px 2px rgba(0,0,0,0.12)',
                  pointerEvents: 'auto',
                  lineHeight: 1,
                  zIndex: 10
                }}
                tabIndex={0}
              >
                HOLD
              </span>
            )}
           </button>
          <div className="min-w-0">
            <h2
              ref={nameRef}
              className={`${nameTextSize} font-bold text-gray-800 break-words`}
              style={{
                display: '-webkit-box',
                WebkitLineClamp: 2,
                WebkitBoxOrient: 'vertical',
                overflow: 'hidden'
              }}
            >
              {studentName}
            </h2>
            {/* Only render the Assigned pill if assignedTo has a value */}
            {assignedTo && (
              <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-gray-200 text-gray-800 mt-1 inline-block truncate max-w-[120px]">
                {assignedTo}
              </span>
            )}
          </div>
        </div>
        {/* Right side: Stats */}
        <div className="flex space-x-2 flex-shrink-0">
          <button
            type="button"
            className={`p-2 text-center rounded-lg ${daysOutBg} ${daysOutText} w-20 focus:outline-none`}
            style={{ position: 'relative' }}
            onClick={() => setShowDaysModal(true)}
            aria-label="Show days out details"
            title="Show days out details"
          >
            <div className="text-xl font-bold">{daysOut}</div>
            <div className={`text-xs font-medium uppercase ${daysOutLabelText}`}>
              {daysOut === 0 ? 'Engaged' : daysOut === 1 ? 'Day Out' : 'Days Out'}
            </div>
            {risk && (
              <span
                aria-label={`Risk Index: hit ${risk.threshold}+ days out ${risk.count} ${risk.count === 1 ? 'time' : 'times'}`}
                role="status"
                title={`Risk Index: hit ${risk.threshold}+ days out ${risk.count} ${risk.count === 1 ? 'time' : 'times'}`}
                style={{
                  position: 'absolute',
                  top: 0,
                  right: 0,
                  transform: 'translate(30%, -30%)',
                  background: '#991b1b', // red-800
                  color: '#ffffff',
                  borderRadius: 9999,
                  fontSize: 10,
                  minWidth: 18,
                  padding: '2px 5px',
                  fontWeight: 700,
                  boxShadow: '0 1px 2px rgba(0,0,0,0.25)',
                  pointerEvents: 'auto',
                  lineHeight: 1.2,
                  zIndex: 10
                }}
                tabIndex={0}
              >
                {risk.count}
              </span>
            )}
          </button>
          <button
            type="button"
            className={`p-2 text-center rounded-lg ${gradeBg} ${gradeText} w-20 transition-colors duration-150 border border-gray-300 hover:border-blue-400`}
            style={{
              outline: 'none',
              position: 'relative',
              cursor: isValidGradebookUrl ? 'pointer' : 'not-allowed'
            }}
            onClick={openGradebook}
            disabled={!isValidGradebookUrl}
            title={isValidGradebookUrl ? "Open Gradebook" : "No Gradebook link"}
            onMouseEnter={() => { if (!isValidGradebookUrl) setShowNoLink(true); }}
            onMouseLeave={() => setShowNoLink(false)}
          >
            <div className="text-xl font-bold">{gradeDisplay}</div>
            <div className={`text-xs font-medium uppercase ${gradeTextLabel}`}>Grade</div>
            {!isValidGradebookUrl && showNoLink && (
              <span
                style={{
                  position: 'absolute',
                  top: 2,
                  right: 2,
                  background: '#f87171',
                  color: '#fff',
                  borderRadius: '6px',
                  fontSize: '10px',
                  padding: '2px 6px',
                  fontWeight: 'bold',
                  zIndex: 2,
                  opacity: showNoLink ? 1 : 0,
                  pointerEvents: 'none',
                  transition: 'opacity 0.15s'
                }}
                aria-label="No gradebook link"
              >
                No Link
              </span>
            )}
          </button>
        </div>
      </div>

      {/* Enlarged Profile Photo Modal */}
      {showEnlargedPhoto && profilePicUrl && !imageError && (
        <div
          onClick={() => setShowEnlargedPhoto(false)}
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.8)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 9999,
            cursor: 'pointer'
          }}
          aria-label="Close enlarged photo"
          role="button"
          tabIndex={0}
          onKeyDown={(e) => {
            if (e.key === 'Escape' || e.key === 'Enter' || e.key === ' ') {
              setShowEnlargedPhoto(false);
            }
          }}
        >
          <img
            src={profilePicUrl}
            alt={studentName ? `${studentName} profile photo enlarged` : 'Student profile photo enlarged'}
            style={{
              maxWidth: '90vw',
              maxHeight: '90vh',
              objectFit: 'contain',
              borderRadius: '8px',
              boxShadow: '0 10px 40px rgba(0, 0, 0, 0.5)'
            }}
            onClick={(e) => e.stopPropagation()}
          />
        </div>
      )}

      {/* Days Out Modal */}
      <DaysOutModal
        isOpen={showDaysModal}
        onClose={() => setShowDaysModal(false)}
        daysOut={typeof daysOut === 'number' ? daysOut : (Number.isFinite(Number(daysOut)) ? Number(daysOut) : null)}
      />
    </div>
  );
}

export default StudentHeader;