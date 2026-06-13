import React from 'react';
import { X } from 'lucide-react';

// Toggles surfaced inside the Create LDA "Advanced" modal. Each id maps to a
// workbookSettings key the LDA processor reads (see DefaultSettings.jsx / ldaProcessor.js).
const ADVANCED_TOGGLES = [
  { id: 'expandedDNC', label: 'Expanded DNC' },
  { id: 'compactColumnWidth', label: 'Compact Column Width' }
];

function Segmented({ value, onChange, options, label }) {
  return (
    <div className="relative flex bg-slate-200 rounded-full p-0.5" role="group" aria-label={label}>
      {options.map((opt) => (
        <button
          key={opt.value}
          type="button"
          onClick={() => onChange(opt.value)}
          aria-pressed={value === opt.value}
          className={`relative z-10 px-3 py-1 text-xs font-medium rounded-full transition-colors duration-200 ${
            value === opt.value ? 'bg-[#145F82] text-white' : 'text-slate-500 hover:text-slate-700'
          }`}
        >
          {opt.label}
        </button>
      ))}
    </div>
  );
}

function ColorRow({ label, value, onChange }) {
  return (
    <div className="flex items-center justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100">
      <div className="text-sm font-medium text-slate-700 min-w-0">{label}</div>
      <div className="flex items-center gap-2 shrink-0">
        <span className="text-[11px] font-mono text-slate-400 uppercase">{value}</span>
        <input
          type="color"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          aria-label={label}
          className="w-8 h-8 rounded-md border border-slate-200 cursor-pointer bg-white p-0.5"
        />
      </div>
    </div>
  );
}

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
              className="flex items-center justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100"
            >
              <div className="text-sm font-medium text-slate-700 min-w-0">{t.label}</div>
              <Toggle
                checked={!!values[t.id]}
                onChange={(v) => onChange(t.id, v)}
                label={t.label}
              />
            </div>
          ))}

          <div className="flex items-center justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100">
            <div className="text-sm font-medium text-slate-700 min-w-0">Author</div>
            <Segmented
              label="Author format"
              value={values.authorFormat === 'full' ? 'full' : 'initials'}
              onChange={(v) => onChange('authorFormat', v)}
              options={[
                { value: 'full', label: 'Full Name' },
                { value: 'initials', label: 'Initials' },
              ]}
            />
          </div>

          <ColorRow
            label="LDA Tag color"
            value={values.ldaTagColor || '#FFFF00'}
            onChange={(v) => onChange('ldaTagColor', v)}
          />
          <ColorRow
            label="DNC color"
            value={values.dncColor || '#f2bdbd'}
            onChange={(v) => onChange('dncColor', v)}
          />
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
