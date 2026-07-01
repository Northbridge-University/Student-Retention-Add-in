// 2025-12-11 11:28 EST - Version 3.8.0 - Updated Program bolding: 'in'/'with' remain normal, text AFTER them is bolded
import React, { useState } from 'react';
import {
  IdCardLanyard,
  Mail,
  CalendarDays,
  Briefcase,
  User,
  Clock,
  GraduationCap,
  LogIn,
  UserCheck,
  Lock,
  Activity,
  Phone,
  Plus,
  Trash2
} from 'lucide-react';
import callIcon from '../../../assets/icons/call-icon.png';
import { ChevronLeft, ChevronRight } from 'lucide-react';
import { formatExcelDate } from '../../utility/Conversion';

// A small reusable component for displaying a single detail item.
const DetailItem = ({ label, value }) => {
  if (!value) return null; // Omit if no value

  const itemStyles = {
    marginBottom: '10px'
  };
  const labelStyles = {
    fontWeight: '600',
    color: '#34495e',
    display: 'block',
    fontSize: '14px'
  };
  const valueStyles = {
    color: '#7f8c8d',
    fontSize: '14px',
    marginLeft: '5px'
  };

  return (
    <div style={itemStyles}>
      <span style={labelStyles}>{label}:</span>
      <span style={valueStyles}>{value}</span>
    </div>
  );
};

// Utility for copy-to-clipboard with Fallback mechanism
const copyToClipboard = async (text) => {
  if (!text) return;

  // 1. Try Modern Async API
  if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch (err) {
      console.warn('Clipboard API failed, attempting fallback...', err);
    }
  }

  // 2. Fallback: document.execCommand('copy')
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.style.top = '0';
    ta.setAttribute('readonly', '');
    
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    
    const successful = document.execCommand('copy');
    document.body.removeChild(ta);
    
    if (!successful) {
        console.error('Fallback copy failed.');
    }
  } catch (err) {
    console.error('All copy methods failed', err);
  }
};

// Helper to format ISO strings (e.g. 2025-11-16 17:47:43+00:00)
const formatIsoDateTime = (isoString) => {
  if (!isoString) return '';
  try {
    const date = new Date(isoString);
    // Return formatted string with date and time
    return date.toLocaleString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit'
    });
  } catch (error) {
    return isoString; // Fallback to original string if parsing fails
  }
};

// Helper to ensure phone numbers contain actual digits
const filterPhone = (phoneStr) => {
  if (!phoneStr) return null;
  // If no digits are found in the string (e.g. "N/A", "Unknown", "-"), treat as empty
  if (!/\d/.test(phoneStr)) return null;
  return phoneStr;
};

// Helper to detect input type (Excel Serial Number vs ISO String) and format accordingly
const formatDynamicDate = (dateValue) => {
  if (!dateValue) return '';

  // Check if purely numeric (Excel Serial Date)
  const isNumeric = !isNaN(dateValue) && !isNaN(parseFloat(dateValue));

  if (isNumeric) {
    return formatExcelDate(dateValue, "long");
  }

  // Otherwise treat as ISO/Standard string
  return formatIsoDateTime(dateValue);
};

// Helper to filter Program Name (Removes Year and Preceding Text)
const filterProgramName = (programStr) => {
  if (!programStr) return null;

  // Regex looks for 19xx or 20xx surrounded by word boundaries
  const yearRegex = /\b(19|20)\d{2}\b/;
  const match = programStr.match(yearRegex);

  if (match) {
    // match.index is where the year starts
    // match[0].length is the length of the year (4)
    // We want the substring starting AFTER the year
    const cutIndex = match.index + match[0].length;
    const result = programStr.substring(cutIndex).trim();
    // If result is empty string, return null so the field is hidden, otherwise return result
    return result || null;
  }

  return programStr;
};

// ** NEW ** Helper to Format Program Display with Bolding
// 1. "Spanish" -> Bolded (unless it's part of the suffix, where it's already bold)
// 2. "in ..." / "with ..." -> The words "in" and "with" are NORMAL. Everything AFTER them is bolded.
const formatProgramDisplay = (text) => {
  if (!text) return text;

  // Find the split point for "in" or "with"
  // Regex: Find ' in ' or ' with ' (case insensitive, word boundaries)
  const splitRegex = /\b(in|with)\b/i;
  const match = text.match(splitRegex);

  let prefix = text;
  let suffixNode = null;

  // If "in" or "with" is found, split the string
  if (match) {
    // We want to keep "in"/"with" in the prefix (non-bold) part.
    // match.index is the start of the word.
    // match[0].length is the length of the word.
    // So the split point is index + length.
    const splitIndex = match.index + match[0].length;
    
    prefix = text.substring(0, splitIndex);
    const suffix = text.substring(splitIndex); // This part is BOLD
    
    // Only render suffix if it actually contains text
    if (suffix) {
      suffixNode = <strong key="suffix">{suffix}</strong>;
    }
  }

  // Now process "Spanish" inside the prefix part (which is still plain text)
  const parts = [];
  const spanishRegex = /(\bSpanish\b)/gi; // case insensitive match for Spanish
  let lastIndex = 0;
  let spanMatch;

  while ((spanMatch = spanishRegex.exec(prefix)) !== null) {
    // Push text before "Spanish"
    if (spanMatch.index > lastIndex) {
      parts.push(prefix.substring(lastIndex, spanMatch.index));
    }
    // Push "Spanish" bolded
    parts.push(<strong key={`span-${spanMatch.index}`}>{spanMatch[0]}</strong>);
    
    lastIndex = spanishRegex.lastIndex;
  }
  // Push remaining text in prefix
  if (lastIndex < prefix.length) {
    parts.push(prefix.substring(lastIndex));
  }

  return (
    <span>
      {parts}
      {suffixNode}
    </span>
  );
};

const style = `
.hover-bg:hover {
  background: #e5e7eb;
}
`;

// Updated CopyField to accept 'displayValue' (React Node) separate from 'value' (Clipboard Text)
function CopyField({ label, value, displayValue, id, Icon, iconSrc }) {
  const [copied, setCopied] = useState(false);

  // If there is no value, omit the component entirely from the view
  if (!value) return null;

  const handleCopy = async () => {
    await copyToClipboard(value);
    setCopied(true);
    setTimeout(() => setCopied(false), 1200);
  };

  return (
    <>
      <style>{style}</style>
      <div
        id={id}
        style={{
          padding: '0.5rem',
          borderRadius: '0.5rem',
          cursor: 'pointer',
          position: 'relative',
          transition: 'background 0.15s',
          display: 'flex',
          alignItems: 'center',
          gap: '0.5rem',
          overflow: 'hidden'
        }}
        className="hover-bg"
        onClick={handleCopy}
      >
        <div
          style={{
            position: 'absolute',
            inset: 0,
            background: 'rgba(34,197,94,0.25)',
            opacity: copied ? 1 : 0,
            pointerEvents: 'none',
            transition: 'opacity 0.7s ease'
          }}
        />
        {iconSrc ? (
          <img src={iconSrc} alt="" style={{ width: 18, height: 18, flexShrink: 0 }} />
        ) : Icon ? (
          <Icon size={18} color="#6b7280" style={{ flexShrink: 0 }} />
        ) : null}
        <div style={{ flex: 1 }}>
          <label style={{ fontSize: '0.75rem', color: '#6b7280' }}>{label}</label>
          <p style={{ fontWeight: 600, color: '#1f2937', margin: 0 }}>
            {/* Render displayValue if present, otherwise render raw value */}
            {displayValue || value}
          </p>
        </div>
      </div>
    </>
  );
}

function StudentDetails({ student }) {
  // State for pagination (0 = Page 1, 1 = Page 2, 2 = Page 3)
  const [page, setPage] = useState(0);
  const totalPages = 3;

  // State for the Emergency Contact view toggle
  const [showEmergency, setShowEmergency] = useState(false);

  const handlePrev = () => setPage((prev) => Math.max(0, prev - 1));
  const handleNext = () => setPage((prev) => Math.min(totalPages - 1, prev + 1));

  const toggleEmergency = () => setShowEmergency((prev) => !prev);

  // Placeholder handlers for adding / removing emergency numbers (logic TBD)
  const handleAddEmergency = () => {
    // TODO: wire up add-emergency-number logic
  };
  const handleRemoveEmergency = () => {
    // TODO: wire up remove-emergency-number logic
  };

  // Helper to determine header text
  const getHeaderText = () => {
    if (showEmergency) return 'Emergency Contact';
    if (page === 0) return 'Details';
    return `Details - ${page + 1}`;
  };

  // Prepare Program Data
  const rawProgram = student.ProgVersDescrip;
  const cleanProgram = filterProgramName(rawProgram);
  const displayProgram = formatProgramDisplay(cleanProgram);

  return (
    <>
      <div className="sticky-header space-y-4 mb-1 pl-2">
        <div className="flex justify-between items-center">
          <h3 className="text-lg font-bold text-gray-800">
            {getHeaderText()}
          </h3>
          <div className="flex items-center gap-2">
            <div className="relative">
              <button
                id="emergency-contact-button"
                className={`w-8 h-8 rounded-full shadow-lg flex items-center justify-center transition-colors ${
                  showEmergency
                    ? 'bg-red-300 text-red-800 hover:bg-red-400'
                    : 'bg-red-200 text-red-700 hover:bg-red-300'
                }`}
                aria-label="Emergency Contact"
                title="Emergency Contact"
                type="button"
                aria-pressed={showEmergency}
                onClick={toggleEmergency}
              >
                <Phone className="h-4 w-4" />
              </button>
            </div>
            {showEmergency && (
              <>
                <div className="relative">
                  <button
                    id="add-emergency-button"
                    className="bg-green-500 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-green-600 transition-colors"
                    aria-label="Add Emergency Number"
                    title="Add Emergency Number"
                    type="button"
                    onClick={handleAddEmergency}
                  >
                    <Plus className="h-4 w-4" />
                  </button>
                </div>
                <div className="relative">
                  <button
                    id="remove-emergency-button"
                    className="bg-red-500 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-red-600 transition-colors"
                    aria-label="Remove Emergency Number"
                    title="Remove Emergency Number"
                    type="button"
                    onClick={handleRemoveEmergency}
                  >
                    <Trash2 className="h-4 w-4" />
                  </button>
                </div>
              </>
            )}
            {!showEmergency && (
              <>
                <div className="relative">
                  <button
                    id="nav-left-details-button"
                    className={`bg-gray-500 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-gray-600 transition-opacity ${
                      page === 0 ? 'opacity-50 cursor-not-allowed' : 'opacity-100'
                    }`}
                    aria-label="Previous"
                    title="Previous"
                    type="button"
                    onClick={handlePrev}
                    disabled={page === 0}
                  >
                    <ChevronLeft className="h-4 w-4" />
                  </button>
                </div>
                <div className="relative">
                  <button
                    id="nav-right-details-button"
                    className={`bg-gray-500 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-gray-600 transition-opacity ${
                      page === totalPages - 1 ? 'opacity-50 cursor-not-allowed' : 'opacity-100'
                    }`}
                    aria-label="Next"
                    title="Next"
                    type="button"
                    onClick={handleNext}
                    disabled={page === totalPages - 1}
                  >
                    <ChevronRight className="h-4 w-4" />
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      <div
        id="panel-details"
        style={{
          padding: '0.25rem 1rem 1rem 0.25rem',
          display: 'flex',
          flexDirection: 'column',
          gap: '1rem'
        }}
      >
        <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>

          {/* EMERGENCY CONTACT: Emergency Numbers saved */}
          {showEmergency && (
            <>
              <CopyField
                label="Emergency Contact"
                value={student.EmergencyContact}
                id="copy-emergency-contact"
                Icon={User}
              />
              <CopyField
                label="Emergency Phone"
                value={filterPhone(student.EmergencyPhone)}
                id="copy-emergency-phone"
                Icon={Phone}
              />
            </>
          )}

          {/* PAGE 1: Contact Information */}
          {!showEmergency && page === 0 && (
            <>
              <CopyField
                label="Student Number"
                value={student.StudentNumber}
                id="copy-student-id"
                Icon={IdCardLanyard}
              />
              <CopyField
                label="Primary Phone"
                value={filterPhone(student.Phone)}
                id="copy-primary-phone"
                iconSrc={callIcon}
              />
              <CopyField
                label="Other Phone"
                value={filterPhone(student.OtherPhone)}
                id="copy-other-phone"
                iconSrc={callIcon}
              />
              <CopyField
                label="Student Email"
                value={student.StudentEmail}
                id="copy-student-email"
                Icon={Mail}
              />
              <CopyField
                label="Personal Email"
                value={student.PersonalEmail}
                id="copy-personal-email"
                Icon={Mail}
              />
              <CopyField
                label="Last LDA"
                value={formatDynamicDate(student["LDA"])}
                id="copy-last-lda"
                Icon={CalendarDays}
              />
            </>
          )}

          {/* PAGE 2: Academic Information */}
          {!showEmergency && page === 1 && (
            <>
              <CopyField
                label="System Student ID"
                value={student.ID}
                id="copy-systudent-id"
                Icon={IdCardLanyard}
              />
              {/* Updated Program Field using both value (for copy) and displayValue (for visuals) */}
              <CopyField
                label="Program"
                value={cleanProgram}
                displayValue={displayProgram}
                id="copy-program"
                Icon={Briefcase}
              />
              <CopyField
                label="Instructor"
                value={student.Instructor}
                id="copy-instructor"
                Icon={User}
              />
              <CopyField
                label="Shift"
                value={student.Shift}
                id="copy-shift"
                Icon={Clock}
              />
              <CopyField
                label="Expected Start Date"
                value={formatExcelDate(student.ExpectedStartDate, "long")}
                id="copy-exp-start-date"
                Icon={CalendarDays}
              />
              <CopyField
                label="Graduation Date"
                value={formatExcelDate(student.GradDate, "long")}
                id="copy-grad-date"
                Icon={GraduationCap}
              />
            </>
          )}

          {/* PAGE 3: Status Information */}
          {!showEmergency && page === 2 && (
            <>
              <CopyField
                label="Last Login"
                value={formatIsoDateTime(student.last_login)}
                id="copy-last-login"
                Icon={LogIn}
              />
              <CopyField
                label="Admissions Rep"
                value={student.AdmissionsRep}
                id="copy-adm-rep"
                Icon={UserCheck}
              />
              <CopyField
                label="Hold Status"
                value={student.Hold}
                id="copy-hold"
                Icon={Lock}
              />
              <CopyField
                label="SAP Status"
                value={student.AdSAPStatus}
                id="copy-sap-status"
                Icon={Activity}
              />
            </>
          )}

        </div>
      </div>
    </>
  );
}

export default StudentDetails;