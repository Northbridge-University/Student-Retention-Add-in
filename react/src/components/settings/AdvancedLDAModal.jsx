import React from 'react';
import { X } from 'lucide-react';

// Toggles surfaced inside the Create LDA "Advanced" modal. Each id maps to a
// workbookSettings key the LDA processor reads (see DefaultSettings.jsx / ldaProcessor.js).
const ADVANCED_TOGGLES = [
  {
    id: 'expandedDNC',
    label: 'Expanded DNC',
    description: 'Fill the Outreach column with the DNC comment text and who left it, instead of a flat "Do not contact". Off restores the original DNC behavior.'
  },
  {
    id: 'compactColumnWidth',
    label: 'Compact Column Width',
    description: 'Pin key columns (Outreach, ProgramVersion, etc.) to fixed widths and wrap the Outreach text. Off autosizes columns like before.'
  }
];

function Toggle({ checked, onChange, label }) {
  return (
    <button
      type="button"
      role="switch"
      aria-checked={checked}
      aria-label={label}
      onClick={() => onChange(!checked)}
      className={`relative inline-flex h-6 w-11 shrink-0 items-center rounded-full transition-colors duration-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20 focus:ring-offset-1 ${
        checked ? 'bg-[#145F82]' : 'bg-slate-200'
      }`}
    >
      <span
        className={`${checked ? 'translate-x-6' : 'translate-x-1'} inline-block h-4 w-4 transform rounded-full bg-white shadow-sm transition-transform duration-200`}
      />
    </button>
  );
}

export default function AdvancedLDAModal({ isOpen, onClose, values, onChange }) {
  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-100">
          <h3 className="text-lg font-semibold text-slate-800">Advanced</h3>
          <button
            type="button"
            onClick={onClose}
            className="text-slate-400 hover:text-slate-600 transition-colors"
            aria-label="Close"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-4 flex flex-col gap-3">
          {ADVANCED_TOGGLES.map((t) => (
            <div
              key={t.id}
              className="flex items-start justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100"
            >
              <div className="min-w-0">
                <div className="text-sm font-medium text-slate-700">{t.label}</div>
                <p className="text-xs text-slate-400 mt-0.5 leading-relaxed">{t.description}</p>
              </div>
              <Toggle
                checked={!!values[t.id]}
                onChange={(v) => onChange(t.id, v)}
                label={t.label}
              />
            </div>
          ))}
        </div>

        <div className="px-4 pb-4 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-sm font-medium text-white bg-[#145F82] hover:bg-[#0f4b66] rounded-lg transition-colors"
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
}
