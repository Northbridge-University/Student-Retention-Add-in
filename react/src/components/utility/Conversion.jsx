// Conversion.jsx

// Convert Excel serial date to JavaScript Date object
// Excel uses a date system where 1 corresponds to 1900-01-01
// This function converts an Excel serial date to a JavaScript Date object

export function formatExcelDate(serial, format = "default") {
  // Parse numeric value
  const num = Number(serial);
  if (isNaN(num)) return "N/A";

  // Separate whole days and fractional day (time)
  const days = Math.floor(num);
  const fraction = num - days;

  // Adjust for Excel's bogus 1900 leap-year (Excel treats 1900 as leap year).
  // For serials >= 60 Excel includes a non-existent 1900-02-29, so subtract 1 day for serials > 59.
  let adjDays = days;
  if (adjDays > 59) adjDays -= 1;

  // Compute UTC milliseconds from epoch 1899-12-31 so that serial 1 => 1900-01-01
  const epochUtc = Date.UTC(1899, 11, 31);
  const ms = epochUtc + adjDays * 86400000 + Math.round(fraction * 86400000);

  const date = new Date(ms);
  if (isNaN(date.getTime())) return "N/A";

  // Helper getters use UTC to avoid local timezone shifts
  if (format === "long") {
    return date.toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      timeZone: 'UTC'
    });
  }

  const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(date.getUTCDate()).padStart(2, '0');
  const yy = String(date.getUTCFullYear()).slice(-2);

  if (format === "short") {
    return `${mm}/${dd}/${yy}`;
  }

  let hours = date.getUTCHours();
  const minutes = String(date.getUTCMinutes()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // 0 -> 12

  return `${mm}/${dd}/${yy} ${hours}:${minutes} ${ampm}`;
}

// Converts "First Last" <-> "Last, First"
export function formatName(name) {
  if (!name || typeof name !== 'string') return name;

  // Detect "Last, First"
  if (name.includes(',')) {
    const [last, first] = name.split(',').map(s => s.trim());
    if (first && last) {
      return `${first} ${last}`;
    }
  } else {
    // Assume "First Last"
    const parts = name.trim().split(/\s+/);
    if (parts.length === 2) {
      return `${parts[1]}, ${parts[0]}`;
    }
  }
  // If format not recognized, return as is
  return name;
}

/**
 * Normalizes the keys of an object:
 * - Converts keys to lowercase
 * - Removes all whitespace
 * * Returns a new object.
 */
export function normalizeKeys(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  const normalized = {};
  Object.keys(obj).forEach(key => {
    const normKey = String(key).toLowerCase().replace(/\s+/g, '');
    normalized[normKey] = obj[key];
  });
  return normalized;
}

// helper: format date as "MM/DD/YY HH:MM AM/PM"
export function formatTimestamp(date = new Date()) {
  try {
    const d = new Date(date);

    // Use Intl to get Eastern Time components (handles EST/EDT automatically)
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone: 'America/New_York',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', hour12: true,
    }).formatToParts(d);

    const get = (type) => parts.find(p => p.type === type)?.value || '';
    const mm = get('month');
    const dd = get('day');
    const yy = get('year').slice(-2);
    const hh = get('hour').padStart(2, '0');
    const mins = get('minute');
    const ampm = get('dayPeriod');

    return `${mm}/${dd}/${yy} ${hh}:${mins} ${ampm}`;
  } catch (_) {
    return String(date);
  }
}

// Formats a phone number string into (XXX)-XXX-XXXX where possible.
// - 10 digits            -> (786)-234-8749
// - 11 digits (US, "1")  -> 1 (786)-234-8749
// Anything else is returned unchanged so partial/invalid input isn't mangled.
export function formatPhoneNumber(value) {
  if (value === null || value === undefined) return '';
  const str = String(value);
  const digits = str.replace(/\D/g, '');

  if (digits.length === 10) {
    return `(${digits.slice(0, 3)})-${digits.slice(3, 6)}-${digits.slice(6)}`;
  }
  if (digits.length === 11 && digits[0] === '1') {
    const d = digits.slice(1);
    return `1 (${d.slice(0, 3)})-${d.slice(3, 6)}-${d.slice(6)}`;
  }
  return str;
}

// Usage example:
// const date = formatExcelDate(44561); // Converts Excel date serial to JavaScript Date
// console.log(date); // Outputs: "10/07/25 12:00 AM" (depending on the input serial)