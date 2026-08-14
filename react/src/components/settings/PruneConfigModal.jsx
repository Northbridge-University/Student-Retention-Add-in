import React, { useEffect, useState } from 'react';
import { X, Loader2 } from 'lucide-react';
import { loadLdaHistory, pruneLdaHistoryNow, DEFAULT_PRUNE_CONFIG } from '../createLDA/riskIndex.js';
import { byteSize } from '../utility/documentSettings.js';

// workbookSettings keys the prune config is stored under (see DefaultSettings.jsx
// / riskIndex.js getPruneConfig). Excel caps the whole settings blob at ~256 KB,
// shared across every key — so LDAHistory is budgeted well below that.
const KEY_ENABLED = 'ldaHistoryPruneEnabled';
const KEY_MAX_SIZE = 'ldaHistoryMaxSizeKB';
const KEY_MAX_DAYS = 'ldaHistoryMaxDays';
const SETTINGS_CAP_KB = 256;

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

function formatKB(bytes) {
	if (typeof bytes !== 'number' || bytes < 0) return '—';
	return `${(bytes / 1024).toFixed(1)} KB`;
}

/**
 * Configure automatic pruning of the saved LDA history so the document-settings
 * blob stays under Excel's ~256 KB cap. `values` carries the flat workbook
 * settings; `onChange(key, value)` persists them (Settings.updateWorkbookSetting).
 */
export default function PruneConfigModal({ isOpen, onClose, values = {}, onChange = () => {} }) {
	const enabled = values[KEY_ENABLED] ?? DEFAULT_PRUNE_CONFIG.enabled;
	const maxSizeKB = values[KEY_MAX_SIZE] ?? DEFAULT_PRUNE_CONFIG.maxSizeKB;
	const maxDays = values[KEY_MAX_DAYS] ?? DEFAULT_PRUNE_CONFIG.maxDays;

	// Local text state for the number fields so the user can clear/edit freely;
	// committed back to workbook settings (as sanitized numbers) on change.
	const [sizeText, setSizeText] = useState(String(maxSizeKB));
	const [daysText, setDaysText] = useState(String(maxDays));

	const [usage, setUsage] = useState(null); // { bytes, dayCount }
	const [pruning, setPruning] = useState(false);
	const [result, setResult] = useState(null); // { removed, dayCount, bytes }

	const refreshUsage = () => {
		try {
			const history = loadLdaHistory();
			const days = (history && history.days) || {};
			setUsage({ bytes: byteSize(history), dayCount: Object.keys(days).length });
		} catch {
			setUsage(null);
		}
	};

	useEffect(() => {
		if (!isOpen) return;
		setSizeText(String(maxSizeKB));
		setDaysText(String(maxDays));
		setResult(null);
		refreshUsage();
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [isOpen]);

	if (!isOpen) return null;

	const commitSize = (text) => {
		setSizeText(text);
		const n = parseInt(text, 10);
		if (Number.isFinite(n) && n >= 1) onChange(KEY_MAX_SIZE, n);
	};
	const commitDays = (text) => {
		setDaysText(text);
		const n = parseInt(text, 10);
		if (Number.isFinite(n) && n >= 0) onChange(KEY_MAX_DAYS, n);
	};

	const handlePruneNow = async () => {
		if (pruning) return;
		setPruning(true);
		setResult(null);
		try {
			const cfg = {
				enabled: true, // "Prune now" always applies, regardless of the auto toggle
				maxSizeKB: parseInt(sizeText, 10) || DEFAULT_PRUNE_CONFIG.maxSizeKB,
				maxDays: parseInt(daysText, 10) || 0,
			};
			const res = await pruneLdaHistoryNow(cfg);
			refreshUsage();
			const history = loadLdaHistory();
			setResult({ removed: res.removed.length, dayCount: res.dayCount, bytes: byteSize(history), ok: res.ok });
		} catch {
			setResult({ error: true });
		} finally {
			setPruning(false);
		}
	};

	const overCap = usage && usage.bytes > SETTINGS_CAP_KB * 1024 * 0.9;

	return (
		<div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
			<div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden" onClick={(e) => e.stopPropagation()}>
				<div className="flex items-center justify-between px-5 py-4 border-b border-slate-100">
					<h3 className="text-lg font-semibold text-slate-800">History Pruning</h3>
					<button type="button" onClick={onClose} className="text-slate-400 hover:text-slate-600 transition-colors" aria-label="Close">
						<X className="w-5 h-5" />
					</button>
				</div>

				<div className="p-4 flex flex-col gap-3">
					{/* Live usage */}
					<div className={`p-3 rounded-xl border text-sm ${overCap ? 'bg-red-50 border-red-200 text-red-700' : 'bg-slate-50/60 border-slate-100 text-slate-600'}`}>
						<div className="flex items-center justify-between">
							<span className="font-medium">Current history</span>
							<span className="font-mono">{usage ? formatKB(usage.bytes) : '—'}</span>
						</div>
						<div className="flex items-center justify-between mt-1 text-[13px]">
							<span>{usage ? `${usage.dayCount} day${usage.dayCount === 1 ? '' : 's'} stored` : ''}</span>
							<span className="text-slate-400">Excel cap ~{SETTINGS_CAP_KB} KB (all settings)</span>
						</div>
					</div>

					{/* Auto-prune toggle */}
					<div className="flex items-center justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100">
						<div className="min-w-0">
							<div className="text-sm font-medium text-slate-700">Auto-prune</div>
							<div className="text-[12px] text-slate-500">Trim history on every LDA save.</div>
						</div>
						<Toggle checked={!!enabled} onChange={(v) => onChange(KEY_ENABLED, v)} label="Auto-prune" />
					</div>

					{/* Max size */}
					<div className="flex items-center justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100">
						<div className="min-w-0">
							<div className="text-sm font-medium text-slate-700">Max size</div>
							<div className="text-[12px] text-slate-500">Drop oldest days past this. 150 KB recommended.</div>
						</div>
						<div className="flex items-center gap-1 shrink-0">
							<input
								type="number"
								min="1"
								value={sizeText}
								onChange={(e) => commitSize(e.target.value)}
								className="w-16 text-right px-2 py-1 text-sm rounded-md border border-slate-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20"
								aria-label="Max history size in KB"
							/>
							<span className="text-xs text-slate-400">KB</span>
						</div>
					</div>

					{/* Max days */}
					<div className="flex items-center justify-between gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-100">
						<div className="min-w-0">
							<div className="text-sm font-medium text-slate-700">Max days</div>
							<div className="text-[12px] text-slate-500">Keep at most this many recent days. 0 = no limit.</div>
						</div>
						<input
							type="number"
							min="0"
							value={daysText}
							onChange={(e) => commitDays(e.target.value)}
							className="w-16 text-right px-2 py-1 text-sm rounded-md border border-slate-200 focus:outline-none focus:ring-2 focus:ring-[#145F82]/20"
							aria-label="Max history days"
						/>
					</div>

					{result && (
						<div className={`text-[13px] px-1 ${result.error || result.ok === false ? 'text-red-600' : 'text-slate-600'}`}>
							{result.error || result.ok === false
								? 'Prune failed — see the browser console.'
								: `Removed ${result.removed} old day${result.removed === 1 ? '' : 's'}. Now ${result.dayCount} day${result.dayCount === 1 ? '' : 's'}, ${formatKB(result.bytes)}.`}
						</div>
					)}
				</div>

				<div className="px-4 pb-4 flex justify-between gap-2">
					<button
						type="button"
						onClick={handlePruneNow}
						disabled={pruning}
						className="px-4 py-2 text-sm font-medium text-[#145F82] bg-[#145F82]/10 hover:bg-[#145F82]/15 rounded-lg transition-colors inline-flex items-center gap-2 disabled:opacity-60"
					>
						{pruning && <Loader2 className="w-4 h-4 animate-spin" />}
						{pruning ? 'Pruning...' : 'Prune now'}
					</button>
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
