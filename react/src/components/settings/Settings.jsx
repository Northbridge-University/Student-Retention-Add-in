import React, { useState, useEffect } from 'react';
import { Info, ChevronRight, ArrowLeft, Layers, HardDrive, MessageSquare } from 'lucide-react'; // removed X + Plus imports (moved to SettingsModal)
import '../studentView/Styling/StudentView.css'; // add StudentView tab styles
import { defaultUserSettings, defaultWorkbookSettings, defaultColumns, sectionIcons } from './DefaultSettings'; // added: import defaults + defaultColumns
import SettingsModal from './SettingsModal'; // new: modal component
import WorkbookSettingsModal from './WorkbookSettingsModal'; // <-- ADDED
import PowerAutomateConfigModal from './PowerAutomateConfigModal'; // <-- ADDED: Power Automate config modal
import TemplateImportExportModal from './TemplateImportExportModal'; // <-- ADDED: Template import/export modal
import CustomParamImportExportModal from './CustomParamImportExportModal'; // <-- ADDED: Custom parameters import/export modal
import AdvancedLDAModal from './AdvancedLDAModal'; // <-- ADDED: Create LDA advanced options modal
import PruneConfigModal from './PruneConfigModal'; // <-- ADDED: LDA history prune config modal
import { saveDocumentSetting, measureDocumentSettings } from '../utility/documentSettings'; // <-- settings write diagnostics
import LicenseChecker from '../utility/LicenseChecker'; // <-- License checker (requires Graph API)
import UserInfoDisplay from '../utility/UserInfoDisplay'; // <-- User info from token (no API needed)
import About from '../about/About'; // <-- ADDED: Import About component for Help tab
import { HISTORY_SHEET, MASTER_LIST_SHEET } from '../../../../shared/constants.js';
import { getWorkbookUsers } from '../../services/workbookUsers.js';
import { loadLdaHistory, historyToCsv, clearLdaHistory as clearLdaHistoryStore } from '../createLDA/riskIndex.js';
import DeleteConfirmModal from './DeleteConfirmModal';
import RiskIndexImportModal from './RiskIndexImportModal'; // <-- ADDED: Risk Index history import modal
import RiskIndexLeaderboard from './RiskIndexLeaderboard'; // <-- ADDED: ranked at-risk students card

const SUMMARY_BRAND = '#145F82';

function formatBytes(bytes) {
	if (bytes == null || Number.isNaN(bytes)) return '—';
	if (bytes < 1024) return `${bytes} B`;
	if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
	if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
	return `${(bytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

// Format an ISO timestamp as a short, readable join date.
function formatJoinDate(iso) {
	if (!iso) return '';
	const d = new Date(iso);
	if (isNaN(d.getTime())) return '';
	return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

// Renders the workbook's registered-users roster on the Users sub-page.
// Reads directly from document settings (not the parent Settings state) so a
// fresh registration written by App.jsx still shows up if it lands while this
// component is mounted, and re-reads on Office's SettingsChanged event.
function UsersList() {
	const [users, setUsers] = useState(() => getWorkbookUsers());

	useEffect(() => {
		const refresh = () => setUsers(getWorkbookUsers());
		refresh();

		const settings = (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document)
			? Office.context.document.settings
			: null;
		if (!settings || typeof settings.addHandlerAsync !== 'function') return;

		const handler = () => refresh();
		settings.addHandlerAsync(Office.EventType.SettingsChanged, handler);
		return () => {
			try { settings.removeHandlerAsync(Office.EventType.SettingsChanged, handler); } catch { /* ignore */ }
		};
	}, []);

	// Newest joins first so a user can spot themselves quickly after registering.
	const sorted = [...users].sort((a, b) => {
		const aTime = a && a.dateJoined ? Date.parse(a.dateJoined) : 0;
		const bTime = b && b.dateJoined ? Date.parse(b.dateJoined) : 0;
		return bTime - aTime;
	});

	return (
		<div style={{ display: 'grid', gap: 8, width: '100%' }}>
			<div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', padding: '0 2px' }}>
				<div style={{ fontSize: 14, fontWeight: 600 }}>Registered Users</div>
				<div style={{ fontSize: 12, color: '#6b7280' }}>
					{sorted.length} {sorted.length === 1 ? 'user' : 'users'}
				</div>
			</div>
			{sorted.length === 0 ? (
				<div style={{
					padding: '12px 14px',
					border: '1px dashed #e5e7eb',
					borderRadius: 8,
					background: '#fff',
					color: '#6b7280',
					fontSize: 13,
					textAlign: 'center',
				}}>
					No users registered yet. Users are added automatically when they open this workbook with the Student Retention Kit signed in.
				</div>
			) : (
				<ul style={{ listStyle: 'none', margin: 0, padding: 0, display: 'grid', gap: 6 }}>
					{sorted.map((u, i) => (
						<li
							key={`${u.email || u.name}-${u.dateJoined || i}`}
							style={{
								display: 'grid',
								gridTemplateColumns: '28px 1fr auto',
								alignItems: 'center',
								gap: 10,
								padding: '8px 12px',
								background: '#fff',
								border: '1px solid #e5e7eb',
								borderRadius: 8,
							}}
						>
							<div style={{
								width: 28, height: 28, borderRadius: '50%',
								background: '#eef2ff', color: '#4338ca',
								display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
								fontSize: 12, fontWeight: 700,
							}}>
								{(u.name || u.email || '?').slice(0, 1).toUpperCase()}
							</div>
							<div style={{ minWidth: 0, display: 'grid', gap: 1 }}>
								<div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: 13, fontWeight: 500 }}>
									{u.name || 'Unknown'}
								</div>
								{u.email && (
									<div
										style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: 11, color: '#6b7280' }}
										title={u.email}
									>
										{u.email}
									</div>
								)}
							</div>
							<div style={{ fontSize: 11, color: '#6b7280', whiteSpace: 'nowrap' }} title={u.dateJoined ? new Date(u.dateJoined).toLocaleString() : undefined}>
								Joined {formatJoinDate(u.dateJoined) || '—'}
							</div>
						</li>
					))}
				</ul>
			)}
		</div>
	);
}

function SummaryCard({ icon, title, value, hint, accent = SUMMARY_BRAND }) {
	return (
		<div style={{
			background: '#fff',
			border: '1px solid #e5e7eb',
			borderRadius: 10,
			padding: '14px 16px',
			display: 'flex',
			flexDirection: 'column',
			gap: 6,
			minWidth: 0,
			boxShadow: '0 1px 2px rgba(15,23,42,0.04)',
		}}>
			<div style={{
				display: 'flex',
				alignItems: 'center',
				gap: 8,
				fontSize: 11,
				fontWeight: 600,
				textTransform: 'uppercase',
				letterSpacing: '0.04em',
				color: '#6b7280',
			}}>
				<span style={{
					width: 24,
					height: 24,
					borderRadius: 6,
					background: `${accent}14`,
					color: accent,
					display: 'inline-flex',
					alignItems: 'center',
					justifyContent: 'center',
					flexShrink: 0,
				}}>
					{icon}
				</span>
				<span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{title}</span>
			</div>
			<div style={{ fontSize: 24, fontWeight: 700, color: '#0f172a', lineHeight: 1.1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
				{value}
			</div>
			{hint && <div style={{ fontSize: 12, color: '#6b7280' }}>{hint}</div>}
		</div>
	);
}

const Settings = ({ user, accessToken, onReady }) => { // <-- ADDED accessToken and onReady props
	const [activeTab, setActiveTab] = useState('workbook');
	// 'main' | '<section name>' — controls sub-page navigation for workbook settings
	const [workbookSectionView, setWorkbookSectionView] = useState('main');

	// initialize user settings state from defaults
	const [userSettingsState, setUserSettingsState] = useState(() =>
		defaultUserSettings.reduce((acc, s) => {
			acc[s.id] = s.defaultValue;
			return acc;
		}, {})
	);

	// initialize workbook settings state from defaults
	const [workbookSettingsState, setWorkbookSettingsState] = useState(() =>
		defaultWorkbookSettings.reduce((acc, s) => {
			acc[s.id] = s.defaultValue;
			return acc;
		}, {})
	);

	// modal state to edit array-type settings
	const [modalOpen, setModalOpen] = useState(false);
	const [modalSetting, setModalSetting] = useState(null);
	const [modalArray, setModalArray] = useState([]);
	const [modalUpdater, setModalUpdater] = useState(null);

	// new: workbook inspector modal state
	const [workbookModalOpen, setWorkbookModalOpen] = useState(false);

	// workbook summary card state (sheet count, file size, comment count)
	const [sheetCount, setSheetCount] = useState(null);
	const [historyCommentCount, setHistoryCommentCount] = useState(null);
	const [fileSizeBytes, setFileSizeBytes] = useState(null);
	const [summaryLoading, setSummaryLoading] = useState(false);
	const [fileSizeLoading, setFileSizeLoading] = useState(false);

	// Power Automate config modal state
	const [powerAutomateModalOpen, setPowerAutomateModalOpen] = useState(false);

	// Template import/export modal state
	const [templateModalOpen, setTemplateModalOpen] = useState(false);

	// Custom parameters import/export modal state
	const [customParamModalOpen, setCustomParamModalOpen] = useState(false);

	// Create LDA "Advanced" options modal state
	const [advancedModalOpen, setAdvancedModalOpen] = useState(false);

	// Download CSV state
	const [downloadingCsv, setDownloadingCsv] = useState(false);

	// Create Student History sheet state
	const [creatingHistorySheet, setCreatingHistorySheet] = useState(false);

	// Risk Index: LDA history import modal + download state
	const [riskImportModalOpen, setRiskImportModalOpen] = useState(false);
	const [downloadingLdaHistory, setDownloadingLdaHistory] = useState(false);
	const [clearHistoryModalOpen, setClearHistoryModalOpen] = useState(false);
	const [clearingLdaHistory, setClearingLdaHistory] = useState(false);
	const [pruneModalOpen, setPruneModalOpen] = useState(false);
	// Bumped to force the At-Risk leaderboard to reload after a history clear.
	const [leaderboardRefresh, setLeaderboardRefresh] = useState(0);

	// Default headers for a freshly-created Student History sheet. Each name is
	// a canonical alias from shared/columnAliases.js so the existing add/edit/
	// delete-comment flows resolve them via canonicalHeaderMap.
	const HISTORY_SHEET_HEADERS = ['SyStudentId', 'Student', 'Comment', 'Timestamp', 'Created By', 'Tag', 'Comment ID'];

	// master list headers read from the active workbook (populated when columns modal opens)
	const [masterListHeaders, setMasterListHeaders] = useState(null);

	// set of Master List headers that are NEW (not seen when the user last saved columns)
	const [newMasterListHeaders, setNewMasterListHeaders] = useState(null);

	// set of Master List headers whose data columns are entirely empty/blank
	const [blankColumns, setBlankColumns] = useState(null);

	// read headers from the "Master List" worksheet via Excel API
	// Returns { headers, blankColumns } where blankColumns is a Set of header names whose data columns are entirely empty
	const readMasterListHeaders = async () => {
		try {
			if (typeof window !== 'undefined' && window.Excel && Excel.run) {
				const result = await Excel.run(async (context) => {
					const sheet = context.workbook.worksheets.getItemOrNullObject(MASTER_LIST_SHEET);
					await context.sync();
					if (sheet.isNullObject) return null;
					const used = sheet.getUsedRangeOrNullObject();
					used.load('values');
					await context.sync();
					if (used.isNullObject || !used.values || used.values.length === 0) return null;
					const allValues = used.values;
					const headers = allValues[0].map(v => (v == null ? '' : String(v).trim())).filter(Boolean);
					// detect blank columns: columns where every data row (rows 1+) is empty/null
					const blankSet = new Set();
					headers.forEach((header, colIdx) => {
						let allEmpty = true;
						for (let row = 1; row < allValues.length; row++) {
							const cell = allValues[row][colIdx];
							if (cell != null && String(cell).trim() !== '') {
								allEmpty = false;
								break;
							}
						}
						// only mark as blank if there are data rows; a header-only sheet is not "blank"
						if (allEmpty && allValues.length > 1) {
							blankSet.add(header);
						}
					});
					return { headers, blankColumns: blankSet };
				});
				return result;
			}
		} catch (err) {
			console.warn('Failed to read Master List headers', err);
		}
		return null;
	};

	// load workbook summary stats: sheet count + Student History row count
	const loadWorkbookSummary = async () => {
		if (typeof window === 'undefined' || !window.Excel || !Excel.run) {
			setSheetCount(null);
			setHistoryCommentCount(null);
			return;
		}
		setSummaryLoading(true);
		try {
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				sheets.load('items/name');
				await context.sync();

				const names = sheets.items.map(s => s.name);
				setSheetCount(names.length);

				if (names.includes(HISTORY_SHEET)) {
					const histSheet = context.workbook.worksheets.getItem(HISTORY_SHEET);
					const histRange = histSheet.getUsedRangeOrNullObject();
					histRange.load('rowCount, isNullObject');
					await context.sync();
					setHistoryCommentCount(histRange.isNullObject ? 0 : Math.max(0, (histRange.rowCount || 0) - 1));
				} else {
					setHistoryCommentCount(null);
				}
			});
		} catch (err) {
			console.warn('Failed to load workbook summary', err);
		} finally {
			setSummaryLoading(false);
		}
	};

	// load workbook file size via Office.js (compressed xlsx); closes immediately so we don't hold slices
	const loadFileSize = () => {
		if (!(typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && typeof Office.context.document.getFileAsync === 'function')) {
			setFileSizeBytes(null);
			return;
		}
		setFileSizeLoading(true);
		try {
			Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (asyncResult) => {
				try {
					if (asyncResult && asyncResult.status === Office.AsyncResultStatus.Succeeded) {
						const file = asyncResult.value;
						setFileSizeBytes(typeof file?.size === 'number' ? file.size : null);
						try { file && typeof file.closeAsync === 'function' && file.closeAsync(() => {}); } catch { /* ignore */ }
					} else {
						setFileSizeBytes(null);
					}
				} finally {
					setFileSizeLoading(false);
				}
			});
		} catch {
			setFileSizeLoading(false);
			setFileSizeBytes(null);
		}
	};

	// refresh summaries when the workbook tab becomes active and on initial mount
	useEffect(() => {
		if (activeTab !== 'workbook' || workbookSectionView !== 'main') return;
		loadWorkbookSummary();
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [activeTab, workbookSectionView]);

	// file size is expensive (Office.getFileAsync streams the workbook), so load it once when Settings mounts
	useEffect(() => {
		loadFileSize();
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	// Download "Student History" sheet as CSV
	const downloadHistoryCsv = async () => {
		if (downloadingCsv) return;
		setDownloadingCsv(true);
		try {
			if (typeof window === 'undefined' || !window.Excel || !Excel.run) {
				throw new Error('Excel API not available');
			}
			const csvContent = await Excel.run(async (context) => {
				const sheet = context.workbook.worksheets.getItemOrNullObject(HISTORY_SHEET);
				await context.sync();
				if (sheet.isNullObject) throw new Error(`${HISTORY_SHEET} sheet not found`);
				const used = sheet.getUsedRangeOrNullObject();
				used.load('values');
				await context.sync();
				if (used.isNullObject || !used.values || used.values.length === 0) {
					throw new Error('Student History sheet is empty');
				}
				// Build CSV from the 2D values array
				return used.values.map(row =>
					row.map(cell => {
						const val = cell == null ? '' : String(cell);
						// Escape cells containing commas, quotes, or newlines
						if (val.includes(',') || val.includes('"') || val.includes('\n')) {
							return `"${val.replace(/"/g, '""')}"`;
						}
						return val;
					}).join(',')
				).join('\n');
			});
			// Trigger download
			const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
			const url = URL.createObjectURL(blob);
			const a = document.createElement('a');
			a.href = url;
			a.download = 'Student History.csv';
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			URL.revokeObjectURL(url);
		} catch (err) {
			console.error('Failed to download Student History CSV:', err);
			alert(err.message || 'Failed to download Student History');
		} finally {
			setDownloadingCsv(false);
		}
	};

	// Create the "Student History" sheet with the default header row.
	// Headers come from shared/columnAliases.js — see HISTORY_SHEET_HEADERS above.
	const createHistorySheet = async () => {
		if (creatingHistorySheet) return;
		setCreatingHistorySheet(true);
		try {
			if (typeof window === 'undefined' || !window.Excel || !Excel.run) {
				throw new Error('Excel API not available');
			}
			await Excel.run(async (context) => {
				const sheets = context.workbook.worksheets;
				const existing = sheets.getItemOrNullObject(HISTORY_SHEET);
				await context.sync();
				if (!existing.isNullObject) {
					// Concurrent create raced us — bail without touching the existing sheet.
					return;
				}
				const sheet = sheets.add(HISTORY_SHEET);
				const headerRange = sheet.getRangeByIndexes(0, 0, 1, HISTORY_SHEET_HEADERS.length);
				headerRange.values = [HISTORY_SHEET_HEADERS];
				headerRange.format.font.bold = true;
				sheet.getUsedRange().format.autofitColumns();
				await context.sync();
			});
			// Optimistically reflect the new (header-only) sheet in the summary so the
			// Create button greys out and the Download button enables without a refresh.
			setHistoryCommentCount(0);
			setSheetCount(prev => (typeof prev === 'number' ? prev + 1 : prev));
		} catch (err) {
			console.error('Failed to create Student History sheet:', err);
			alert(err.message || 'Failed to create Student History sheet');
		} finally {
			setCreatingHistorySheet(false);
		}
	};

	// Risk Index: download the saved LDA history as CSV
	const downloadLdaHistory = () => {
		if (downloadingLdaHistory) return;
		setDownloadingLdaHistory(true);
		try {
			const history = loadLdaHistory();
			const dayCount = Object.keys((history && history.days) || {}).length;
			if (dayCount === 0) {
				alert('No LDA history saved yet. It builds up as LDA sheets are created, or use Import Past Reports.');
				return;
			}
			const blob = new Blob([historyToCsv(history)], { type: 'text/csv;charset=utf-8;' });
			const url = URL.createObjectURL(blob);
			const a = document.createElement('a');
			a.href = url;
			a.download = 'LDA History.csv';
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			URL.revokeObjectURL(url);
		} catch (err) {
			console.error('Failed to download LDA history:', err);
			alert(err.message || 'Failed to download LDA history');
		} finally {
			setDownloadingLdaHistory(false);
		}
	};

	// Risk Index: wipe the saved LDA history (confirmed via modal). Resets every
	// student's Risk Index count to zero; also shrinks the document-settings blob.
	const handleClearLdaHistory = async () => {
		if (clearingLdaHistory) return;
		setClearingLdaHistory(true);
		try {
			const ok = await clearLdaHistoryStore();
			if (!ok) {
				alert('Failed to clear LDA history. See the browser console for details.');
				return;
			}
			setClearHistoryModalOpen(false);
			// Reflect the reset in the At-Risk leaderboard without a manual refresh.
			setLeaderboardRefresh(n => n + 1);
		} catch (err) {
			console.error('Failed to clear LDA history:', err);
			alert(err.message || 'Failed to clear LDA history');
		} finally {
			setClearingLdaHistory(false);
		}
	};

	const openArrayModal = async (setting, currentValue, updater) => {
		// if this is the columns setting, read master list headers first
		if (setting.id === 'columns') {
			const result = await readMasterListHeaders();
			const headers = result ? result.headers : null;
			setMasterListHeaders(headers); // null means no master list found (show all)
			setBlankColumns(result ? result.blankColumns : null);

			// detect new headers by comparing current ML headers against lastKnownHeaders
			const lastKnown = workbookSettingsState.lastKnownHeaders;
			if (headers && Array.isArray(lastKnown) && lastKnown.length > 0) {
				const stripStr = s => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
				const knownSet = new Set(lastKnown.map(h => stripStr(h)));
				const newSet = new Set();
				headers.forEach(h => {
					if (!knownSet.has(stripStr(h))) newSet.add(h);
				});
				setNewMasterListHeaders(newSet.size > 0 ? newSet : null);
			} else {
				// first time opening or no lastKnownHeaders — nothing is "new"
				setNewMasterListHeaders(null);
			}
		} else {
			setMasterListHeaders(null);
			setNewMasterListHeaders(null);
			setBlankColumns(null);
		}

		setModalSetting(setting);
		// normalize items to objects: { name: string, alias: string[], edit: 'name'|'alias', static: boolean }
		const normalized = Array.isArray(currentValue)
			? currentValue.map(it => {
					if (typeof it === 'string') return { name: it, alias: [], edit: 'name', static: false };
					if (it && typeof it === 'object') {
						return {
							name: it.name ?? '',
							alias: Array.isArray(it.alias)
								? it.alias
								: it.alias
									? String(it.alias).split(',').map(s => s.trim()).filter(Boolean)
									: [],
							edit: 'name',
							static: !!it.static
						};
					}
					return { name: '', alias: [], edit: 'name', static: false };
			  })
			: [];
		setModalArray(normalized);
		// store updater as a function reference
		setModalUpdater(() => updater);
		setModalOpen(true);
	};

	const closeModal = () => {
		setModalOpen(false);
		setModalSetting(null);
		setModalArray([]);
		setModalUpdater(null);
	};

	// modal save: accept optional newArray provided by the modal to avoid reading stale state
	const saveModal = (providedArray) => {
		if (modalSetting && typeof modalUpdater === 'function') {
			// prefer the array passed from the modal (freshly computed) otherwise use parent's modalArray state
			const sourceArray = Array.isArray(providedArray) ? providedArray : modalArray;
			// ensure we send a clean array of objects: { name, alias: string[], static: boolean, ...options }
			const cleaned = (Array.isArray(sourceArray) ? sourceArray : []).map(it => ({
				name: String(it?.name ?? it?.column ?? '').trim(),
				alias: Array.isArray(it?.alias) ? it.alias.map(a => String(a).trim()).filter(Boolean) : [],
				static: !!it?.static,
				// preserve any other option fields the editor wrote under `options`
				...((it && it.options && typeof it.options === 'object') ? it.options : {})
			}));
			modalUpdater(modalSetting.id, cleaned);

			// when saving columns, snapshot current Master List headers so we can detect new ones later
			if (modalSetting.id === 'columns' && Array.isArray(masterListHeaders)) {
				updateWorkbookSetting('lastKnownHeaders', masterListHeaders);
				setNewMasterListHeaders(null); // clear "new" badges after save
			}
		}
		// close modal via parent helper
		closeModal();
	};

	// Persist/load helpers for Office document (Workbook) settings
	const DOC_KEY = 'workbookSettings';

	const loadWorkbookSettingsFromDocument = () => {
		try {
			// ensure Office.js is available and document settings exist
			if (typeof window !== 'undefined' && window.Office && Office.context && Office.context.document && Office.context.document.settings) {
				const docSettings = Office.context.document.settings.get(DOC_KEY);
				return docSettings || null;
			}
		} catch (err) {
			console.warn('Failed to read document settings', err);
		}
		return null;
	};

	const saveWorkbookSettingsToDocument = (mapping) => {
		// Routed through saveDocumentSetting so any OSF.DDA.Error surfaces the
		// real code/message plus a per-key size breakdown (see documentSettings.js).
		return saveDocumentSetting(DOC_KEY, mapping, 'Settings.saveWorkbookSettingsToDocument');
	};

	// one-time effect: ensure document has a workbook settings key; load it if present
	useEffect(() => {
		// Baseline size log on every Settings open, so an approaching size-cap
		// problem is visible in the console before a save actually fails.
		measureDocumentSettings('Settings mount');
		const existing = loadWorkbookSettingsFromDocument();
		if (existing && typeof existing === 'object') {
			// if columns are missing or empty, inject defaultColumns (one-time import)
			const needsColumns = !Array.isArray(existing.columns) || existing.columns.length === 0;
			const merged = { ...existing };
			if (needsColumns) {
				merged.columns = [...defaultColumns];
				// persist the merged mapping back to document settings
				saveWorkbookSettingsToDocument(merged);
			}
			setWorkbookSettingsState(prev => ({ ...prev, ...merged }));
			return;
		}

		// no settings in document: create mapping from defaultWorkbookSettings and persist once
		const initial = defaultWorkbookSettings.reduce((acc, s) => {
			acc[s.id] = s.defaultValue;
			return acc;
		}, {});
		// ensure columns are populated from defaultColumns on initial import
		if (!Array.isArray(initial.columns) || initial.columns.length === 0) {
			initial.columns = [...defaultColumns];
		}
		setWorkbookSettingsState(initial);
		saveWorkbookSettingsToDocument(initial);
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []); // run once on mount

	// Signal that Settings is ready
	useEffect(() => {
		if (onReady) {
			onReady();
		}
	}, [onReady]);

	// helper to update a workbook setting (also persist to document)
	const updateWorkbookSetting = (id, value) => {
		setWorkbookSettingsState(prev => {
			const next = { ...prev, [id]: value };
			// persist mapping to document settings
			saveWorkbookSettingsToDocument(next);
			return next;
		});
	};

	// helper to update a setting
	const updateSetting = (id, value) => {
		setUserSettingsState(prev => ({ ...prev, [id]: value }));
	};

	// render settings grouped by optional `section` property
	const renderSettingsControls = (settings, state, updater, idPrefix = '') => {
		// group by section
		const sections = {};
		const unsectioned = [];
		settings.forEach(s => {
			// Hidden settings still seed defaults/persist but aren't rendered as rows
			// (e.g. surfaced through a dedicated modal instead).
			if (s.hidden) return;
			if (s.section) {
				(sections[s.section] ||= []).push(s);
			} else {
				unsectioned.push(s);
			}
		});

		const renderRow = setting => {
			const cur = state[setting.id];
			const inputId = `${idPrefix}${setting.id}`;
			// Custom full-width row for the workbook user roster.
			if (setting.type === 'userslist') {
				return <UsersList key={setting.id} />;
			}
			// Custom full-width row: students ranked by Risk Index. Reloads when the
			// import modal closes or the history is cleared (either changes the token).
			if (setting.type === 'riskleaderboard') {
				return <RiskIndexLeaderboard key={setting.id} refreshToken={`${riskImportModalOpen}|${leaderboardRefresh}`} />;
			}
			return (
				<div key={setting.id} style={{ display: 'grid', gridTemplateColumns: '1fr auto', alignItems: 'center', gap: 8, width: '100%', boxSizing: 'border-box' }}>
					{/* label area: stays in the left column and truncates if too long */}
					<div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0, overflow: 'hidden' }}>
						<div style={{ fontSize: 14, fontWeight: 600, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
							{setting.label}
						</div>
						{/* info icon kept compact and closer to the label */}
						{setting.description && (
							<button
								type="button"
								title={setting.description}
								aria-label={`${setting.label} description`}
								style={{
									marginLeft: 1,
									padding: 1,
									border: 'none',
									background: 'transparent',
									color: '#6b7280',
									cursor: 'pointer',
									display: 'inline-flex',
									alignItems: 'center',
								}}
							>
								<Info size={14} />
							</button>
						)}
					</div>

					{/* input/control container: reduced gap so control sits nearer the label */}
					<div style={{ display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'flex-start' }}>
						{(() => {
							if (setting.type === 'boolean') {
								// Tailwind-like toggle: accessible hidden checkbox + visual track & knob
								return (
									<label style={{ display: 'inline-flex', alignItems: 'center', cursor: 'pointer' }}>
										<span style={{ position: 'relative', width: 44, height: 24, display: 'inline-block' }}>
											<input
												type="checkbox"
												checked={!!cur}
												onChange={e => updater(setting.id, e.target.checked)}
												aria-label={setting.label}
												style={{ position: 'absolute', opacity: 0, width: 0, height: 0 }}
											/>
											{/* track */}
											<span style={{
												position: 'absolute',
												inset: 0,
												borderRadius: 9999,
												background: !!cur ? '#145F82' : '#e5e7eb',
												transition: 'background-color 160ms linear',
												padding: 2,
												boxSizing: 'border-box'
											}} />
											{/* knob */}
											<span style={{
												position: 'absolute',
												top: 2,
												left: !!cur ? 22 : 2,
												width: 20,
												height: 20,
												borderRadius: '50%',
												background: '#fff',
												boxShadow: '0 1px 2px rgba(0,0,0,0.2)',
												transition: 'left 160ms linear',
											}} />
										</span>
									</label>
								);
							}

							if (setting.type === 'select' && Array.isArray(setting.options)) {
								return (
									<select value={cur ?? ''} onChange={e => updater(setting.id, e.target.value)} style={{ padding: '6px 10px', borderRadius: 6, minWidth: 120 }} aria-label={setting.label}>
										{setting.options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
									</select>
								);
							}

							if (setting.type === 'array') {
								// show disabled Configure button with lock indicator when locked
								if (setting.locked) {
									return (
										<button
											onClick={() => {}}
											style={{
												padding: '6px 10px',
												borderRadius: 6,
												background: '#f9fafb',
												border: '1px solid #e6e7eb',
												cursor: 'not-allowed',
												display: 'inline-flex',
												alignItems: 'center',
												gap: 8,
												color: '#6b7280'
											}}
											aria-label={`Configure`}
											aria-disabled="true"
											title="This setting is locked"
										>
											{/* small lock icon */}
											<svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" style={{ display: 'block', color: '#9ca3af' }}>
												<rect x="3" y="11" width="18" height="10" rx="2" stroke="currentColor" strokeWidth="1.2" />
												<path d="M7 11V8a5 5 0 0 1 10 0v3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
											</svg>
											<span style={{ fontSize: 13 }}>Locked</span>
										</button>
									);
								}
								return (
									<button
										onClick={() => openArrayModal(setting, cur, updater)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'powerautomate') {
								return (
									<button
										onClick={() => setPowerAutomateModalOpen(true)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure Power Automate`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'templates') {
								return (
									<button
										onClick={() => setTemplateModalOpen(true)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure email templates import/export`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'customparameters') {
								return (
									<button
										onClick={() => setCustomParamModalOpen(true)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure custom parameters import/export`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'advanced') {
								return (
									<button
										onClick={() => setAdvancedModalOpen(true)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure advanced Create LDA options`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'pruneconfig') {
								return (
									<button
										onClick={() => setPruneModalOpen(true)}
										style={{ padding: '6px 10px', borderRadius: 6, background: '#f3f4f6', border: '1px solid #e6e7eb', cursor: 'pointer' }}
										aria-label={`Configure LDA history pruning`}
									>
										Configure
									</button>
								);
							}

							if (setting.type === 'action') {
								const plusIcon = (
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
										<line x1="12" y1="5" x2="12" y2="19" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
										<line x1="5" y1="12" x2="19" y2="12" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
									</svg>
								);
								const downloadIcon = (
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
										<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
										<polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
										<line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
									</svg>
								);
								const importIcon = (
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
										<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
										<polyline points="7 8 12 3 17 8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
										<line x1="12" y1="3" x2="12" y2="15" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
									</svg>
								);
								const trashIcon = (
									<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
										<polyline points="3 6 5 6 21 6" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
										<path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
									</svg>
								);

								// historyCommentCount is null both before the workbook summary loads and
								// when the History sheet doesn't exist — either way, there's nothing to
								// download yet, so disable the button. Inverse for Create: greyed out when
								// the sheet already exists, or while the workbook summary is still
								// resolving so we don't race a duplicate.
								const noHistorySheet = setting.id === 'downloadHistoryCsv' && historyCommentCount == null;
								const historySheetExists = setting.id === 'createHistorySheet' && (historyCommentCount != null || summaryLoading);
								const actions = {
									createHistorySheet: {
										run: createHistorySheet,
										busy: creatingHistorySheet,
										icon: plusIcon,
										idle: 'Create',
										busyLabel: 'Creating...',
										greyed: historySheetExists,
										greyedReason: historyCommentCount != null ? `"${HISTORY_SHEET}" sheet already exists` : undefined,
									},
									downloadHistoryCsv: {
										run: downloadHistoryCsv,
										busy: downloadingCsv,
										icon: downloadIcon,
										idle: 'Download',
										busyLabel: 'Downloading...',
										greyed: noHistorySheet,
										greyedReason: noHistorySheet ? `No "${HISTORY_SHEET}" sheet found — nothing to download` : undefined,
									},
									importLdaHistory: {
										run: () => setRiskImportModalOpen(true),
										busy: false,
										icon: importIcon,
										idle: 'Import',
										busyLabel: 'Importing...',
									},
									downloadLdaHistory: {
										run: downloadLdaHistory,
										busy: downloadingLdaHistory,
										icon: downloadIcon,
										idle: 'Download',
										busyLabel: 'Downloading...',
									},
									clearLdaHistory: {
										run: () => setClearHistoryModalOpen(true),
										busy: clearingLdaHistory,
										icon: trashIcon,
										idle: 'Clear',
										busyLabel: 'Clearing...',
									},
								};
								const action = actions[setting.id];
								if (!action) return null;
								const isDisabled = action.busy || !!action.greyed;
								return (
									<button
										onClick={() => { if (!isDisabled) action.run(); }}
										disabled={isDisabled}
										title={action.greyedReason}
										style={{
											padding: '6px 10px',
											borderRadius: 6,
											background: isDisabled ? '#e5e7eb' : '#f3f4f6',
											border: '1px solid #e6e7eb',
											cursor: isDisabled ? 'not-allowed' : 'pointer',
											color: action.greyed ? '#9ca3af' : undefined,
											display: 'inline-flex',
											alignItems: 'center',
											gap: 6,
											fontSize: 13,
										}}
										aria-label={setting.label}
									>
										{action.icon}
										{action.busy ? action.busyLabel : action.idle}
									</button>
								);
							}

							if (setting.type === 'image') {
								const fileId = `file-${inputId}`;
								return (
									<div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
										<input id={fileId} type="file" accept="image/*" style={{ display: 'none' }} onChange={e => {
											const f = e.target.files && e.target.files[0];
											if (f) {
												const r = new FileReader();
												r.onload = () => updater(setting.id, r.result);
												r.readAsDataURL(f);
											}
										}} aria-label={`Upload ${setting.label}`} />
										<label htmlFor={fileId} title="Upload" style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: 32, height: 32, borderRadius: 6, background: '#f3f4f6', cursor: 'pointer', border: '1px solid #e5e7eb', color: '#374151' }}>
											<svg width="18" height="18" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
												<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
												<polyline points="17 8 12 3 7 8" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
												<line x1="12" y1="3" x2="12" y2="15" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
											</svg>
										</label>
										<button onClick={() => updater(setting.id, setting.defaultValue)} style={{ padding: '6px 8px', borderRadius: 6 }} title="Reset to default">Reset</button>
									</div>
								);
							}

							if (setting.type === 'number') {
								return (
									<input
										type="number"
										value={cur === undefined || cur === null ? '' : cur}
										onChange={e => updater(setting.id, e.target.value === '' ? '' : Number(e.target.value))}
										style={{
											padding: '6px 8px',
											borderRadius: 6,
											width: 72,
											boxSizing: 'border-box',
											border: '1px solid #e6e7eb', /* very subtle outline */
											background: '#fff',
											boxShadow: 'inset 0 0 0 1px rgba(0,0,0,0.02)',
											transition: 'box-shadow 120ms ease, border-color 120ms ease'
										}}
										placeholder={setting.placeholderValue ?? ''}
										aria-label={setting.label}
									/>
								);
							}

							return (
								<input
									type="text"
									value={cur ?? ''}
									onChange={e => updater(setting.id, e.target.value)}
									style={{
										padding: '6px 10px',
										borderRadius: 6,
										border: '1px solid #e6e7eb', /* very subtle outline */
										background: '#fff',
										boxShadow: 'inset 0 0 0 1px rgba(0,0,0,0.02)',
										transition: 'box-shadow 120ms ease, border-color 120ms ease',
										minWidth: 0
									}}
									placeholder={setting.placeholderValue ?? ''}
									aria-label={setting.label}
								/>
							);
						})()}
					</div>
				</div>
 			);
 		};

		// Sub-page: render Back button + rows for the active section.
		if (workbookSectionView !== 'main' && sections[workbookSectionView]) {
			const name = workbookSectionView;
			return (
				<div style={{ display: 'grid', gap: 8 }}>
					<button
						type="button"
						onClick={() => setWorkbookSectionView('main')}
						style={{
							display: 'inline-flex', alignItems: 'center', gap: 4,
							background: 'transparent', border: 'none', padding: '2px 4px',
							color: '#94a3b8', cursor: 'pointer', fontSize: 13, fontWeight: 500,
							width: 'fit-content'
						}}
						aria-label="Back to Settings"
					>
						<ArrowLeft size={16} />
						Back
					</button>
					<h3 style={{ margin: '2px 0', fontSize: 15, fontWeight: 700, backgroundColor: '#eaeaea', padding: '2px 6px', borderRadius: 6 }}>
						<span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
							{sectionIcons[name] && (
								<span aria-hidden="true" style={{ display: 'inline-flex', width: 16, height: 16, alignItems: 'center' }}>
									{sectionIcons[name]}
								</span>
							)}
							<span>{name}</span>
						</span>
					</h3>
					<div style={{ padding: 8, border: '1px solid #f3f4f6', borderRadius: 6, background: '#fafafa', display: 'grid', gap: 8 }}>
						{sections[name].map(renderRow)}
					</div>
				</div>
			);
		}

		// Main: unsectioned rows first, then one clickable header button per section.
		return (
			<div style={{ display: 'grid', gap: 8 }}>
				{unsectioned.map(renderRow)}
				{Object.keys(sections).map(name => (
					<button
						key={name}
						type="button"
						onClick={() => setWorkbookSectionView(name)}
						style={{
							display: 'flex', alignItems: 'center', justifyContent: 'space-between',
							margin: 0, padding: '6px 8px',
							backgroundColor: '#eaeaea', border: 'none', borderRadius: 6,
							cursor: 'pointer', width: '100%', textAlign: 'left',
							fontSize: 15, fontWeight: 700, color: 'inherit'
						}}
					>
						<span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
							{sectionIcons[name] && (
								<span aria-hidden="true" style={{ display: 'inline-flex', width: 16, height: 16, alignItems: 'center' }}>
									{sectionIcons[name]}
								</span>
							)}
							<span>{name}</span>
						</span>
						<ChevronRight size={16} style={{ color: '#94a3b8' }} />
					</button>
				))}
			</div>
		);
	};

	// handle button actions per setting type
	const handleSettingButton = (setting) => {
		const cur = userSettingsState[setting.id];
		if (setting.type === 'boolean') {
			updateSetting(setting.id, !cur);
		} else if (setting.type === 'select' && Array.isArray(setting.options)) {
			const idx = setting.options.indexOf(cur);
			const next = setting.options[(idx + 1) % setting.options.length];
			updateSetting(setting.id, next);
		} else if (setting.type === 'image') {
			// reset to default image on click
			updateSetting(setting.id, setting.defaultValue);
		} else if (setting.type === 'number') {
			// simple increment for demo
			updateSetting(setting.id, (typeof cur === 'number' ? cur + 1 : setting.defaultValue));
		} else {
			// fallback: toggle between default and empty
			updateSetting(setting.id, cur === setting.defaultValue ? '' : setting.defaultValue);
		}
	};

	return (
		<div
			className="settings-placeholder"
			role="region"
			aria-label="Settings"
			style={{ position: 'relative' }} // added: anchor for absolute avatar
		>
			<div
				className="settings-header"
				style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}
			>
				<div style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
					<h2 className="text-2xl text-slate-800 font-bold tracking-tight" style={{ margin: 0, paddingLeft: 12 }}>Settings</h2>
					<p className="text-slate-400 text-sm mt-1" style={{ paddingLeft: 12 }}>Manage your workbook and user preferences</p>

					{/* Tabs - use same classes as StudentView for identical styling */}
					<div className="studentview-tabs" role="tablist" aria-label="Settings tabs" style={{ marginBottom: 8 }}>
						<button
							role="tab"
							aria-selected={activeTab === 'workbook'}
							onClick={() => setActiveTab('workbook')}
							className={`studentview-tab ${activeTab === 'workbook' ? 'active' : ''}`}
						>
							Workbook
						</button>
						<button
							role="tab"
							aria-selected={activeTab === 'user'}
							onClick={() => setActiveTab('user')}
							className={`studentview-tab ${activeTab === 'user' ? 'active' : ''}`}
						>
							User
						</button>
					</div>

					{/* Tab panels */}
					<div
						role="tabpanel"
						aria-labelledby={activeTab}
						style={{ marginTop: 12, padding: 12, border: '1px solid #e5e7eb', borderRadius: 8, background: '#fff' }}
					>
						{activeTab === 'workbook' ? (
							<div style={{ display: 'grid', gap: 8 }}>
								{renderSettingsControls(defaultWorkbookSettings, workbookSettingsState, updateWorkbookSetting, 'workbook-')}
								{workbookSectionView === 'main' && (
									<>
										<button
											type="button"
											onClick={() => setWorkbookModalOpen(true)}
											title="Open workbook debug view"
											aria-label="Open workbook debug view"
											style={{
												display: 'flex', alignItems: 'center', justifyContent: 'space-between',
												margin: 0, padding: '6px 8px',
												backgroundColor: '#eaeaea', border: 'none', borderRadius: 6,
												cursor: 'pointer', width: '100%', textAlign: 'left',
												fontSize: 15, fontWeight: 700, color: 'inherit'
											}}
										>
											<span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
												<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
													<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
													<path d="M14 2v6h6" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round"/>
												</svg>
												<span>Workbook Debug</span>
											</span>
											<ChevronRight size={16} style={{ color: '#94a3b8' }} />
										</button>

										<div style={{
											display: 'grid',
											gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
											gap: 12,
											marginTop: 4,
										}}>
											<SummaryCard
												icon={<Layers size={14} />}
												title="Sheets in Workbook"
												value={summaryLoading && sheetCount == null ? '…' : (sheetCount == null ? '—' : sheetCount.toLocaleString())}
												hint={sheetCount === 1 ? 'sheet' : 'sheets'}
											/>
											<SummaryCard
												icon={<HardDrive size={14} />}
												title="File Size"
												value={fileSizeLoading ? '…' : formatBytes(fileSizeBytes)}
												hint={fileSizeLoading ? 'Loading…' : (fileSizeBytes == null ? 'Unavailable' : 'Compressed (.xlsx)')}
											/>
											<SummaryCard
												icon={<MessageSquare size={14} />}
												title="Student History Comments"
												value={summaryLoading && historyCommentCount == null ? '…' : (historyCommentCount == null ? '—' : historyCommentCount.toLocaleString())}
												hint={historyCommentCount == null ? `No "${HISTORY_SHEET}" sheet` : 'Rows excluding header'}
											/>
										</div>
									</>
								)}
							</div>
						) : activeTab === 'user' ? (
							<div>
								{/* User Information from SSO Token */}
								<UserInfoDisplay accessToken={accessToken} />

								{/* Power Automate License Section - Currently disabled due to Graph API requirements */}
								{/* Uncomment this section once you implement backend OBO flow */}
								{/*
								<div style={{ marginBottom: 16 }}>
									<h3 style={{ margin: '0 0 8px 0', fontSize: 14, fontWeight: 600, color: '#374151' }}>Power Automate License</h3>
									<LicenseChecker accessToken={accessToken} />
								</div>
								*/}
							</div>
						) : null}
					</div>
				</div>
			</div>

			{/* Render the modals via the new component */}
			<SettingsModal
				// array modal
				modalOpen={modalOpen}
				modalSetting={modalSetting}
				modalArray={modalArray}
				setModalArray={setModalArray}
				closeModal={closeModal}
				saveModal={saveModal}
				// provide current workbook columns so modal uses document columns instead of defaults
				workbookColumns={workbookSettingsState.columns}
				// master list headers to filter which columns are shown (null = show all)
				masterListHeaders={masterListHeaders}
				// set of ML headers that are new since last save (for "New" badge in Add picker)
				newMasterListHeaders={newMasterListHeaders}
				// set of ML headers whose data columns are entirely blank
				blankColumns={blankColumns}
			/>
			{/* pass the document key constant so the modal can read the workbook-specific mapping */}
			<WorkbookSettingsModal isOpen={workbookModalOpen} onClose={() => setWorkbookModalOpen(false)} docKey={DOC_KEY} />

			{/* Power Automate configuration modal */}
			<PowerAutomateConfigModal isOpen={powerAutomateModalOpen} onClose={() => setPowerAutomateModalOpen(false)} />

			{/* Email Templates import/export modal */}
			<TemplateImportExportModal isOpen={templateModalOpen} onClose={() => setTemplateModalOpen(false)} />

			{/* Custom Parameters import/export modal */}
			<CustomParamImportExportModal isOpen={customParamModalOpen} onClose={() => setCustomParamModalOpen(false)} />

			{/* Create LDA Advanced options modal */}
			<AdvancedLDAModal
				isOpen={advancedModalOpen}
				onClose={() => setAdvancedModalOpen(false)}
				values={{
					expandedDNC: workbookSettingsState.expandedDNC ?? false,
					compactColumnWidth: workbookSettingsState.compactColumnWidth ?? true,
					authorFormat: workbookSettingsState.authorFormat ?? 'initials',
					ldaTagColor: workbookSettingsState.ldaTagColor ?? '#FFFF00',
					dncColor: workbookSettingsState.dncColor ?? '#f2bdbd',
				}}
				onChange={updateWorkbookSetting}
			/>

			<RiskIndexImportModal
				isOpen={riskImportModalOpen}
				onClose={() => setRiskImportModalOpen(false)}
			/>

			<PruneConfigModal
				isOpen={pruneModalOpen}
				onClose={() => setPruneModalOpen(false)}
				values={workbookSettingsState}
				onChange={updateWorkbookSetting}
			/>

			<DeleteConfirmModal
				isOpen={clearHistoryModalOpen}
				title="Clear LDA history?"
				message="This permanently deletes all saved LDA history and resets every student's Risk Index to zero. This cannot be undone. Consider using Download History first if you want a backup."
				confirmLabel={clearingLdaHistory ? 'Clearing...' : 'Clear History'}
				onConfirm={handleClearLdaHistory}
				onCancel={() => { if (!clearingLdaHistory) setClearHistoryModalOpen(false); }}
			/>
		</div>
	);
};

export default Settings;
