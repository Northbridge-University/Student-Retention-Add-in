import React, { useState, useEffect } from 'react';
import { FileSearch, Upload, Loader2 } from 'lucide-react';
import ExcelJS from 'exceljs';
import {
    importHistoryFromLdaSheets,
    parseCsv,
    excelCellToValue,
    rowsToImportedDays,
    importDaysIntoHistory,
} from '../createLDA/riskIndex.js';

// Import LDA history either by scanning the workbook's previous "LDA M-D-YYYY"
// sheets or by uploading a CSV/Excel file in the Download History layout
// (Date, SyStudentID, LDA, Days Out). Both paths merge at the day level and
// never overwrite dates that already have a snapshot.
export default function RiskIndexImportModal({ isOpen, onClose }) {
    const [busy, setBusy] = useState(null); // null | 'scan' | 'file'
    const [result, setResult] = useState(null); // { lines: string[] }
    const [error, setError] = useState('');

    useEffect(() => {
        if (!isOpen) return;
        setBusy(null);
        setResult(null);
        setError('');
    }, [isOpen]);

    const describeMerge = ({ added, skipped }) => {
        const lines = [];
        lines.push(added.length > 0
            ? `Imported ${added.length} day${added.length === 1 ? '' : 's'} of history (${added[0]} to ${added[added.length - 1]}).`
            : 'Nothing new to import.');
        if (skipped.length > 0) {
            lines.push(`${skipped.length} day${skipped.length === 1 ? '' : 's'} already in history (left untouched).`);
        }
        return lines;
    };

    const handleScan = async () => {
        if (busy) return;
        setBusy('scan');
        setError('');
        setResult(null);
        try {
            const res = await importHistoryFromLdaSheets();
            if (res.scanned === 0) {
                setError('No previous LDA sheets found. This looks for sheets named like "LDA 7-21-2026".');
                return;
            }
            const lines = [`Scanned ${res.scanned} LDA sheet${res.scanned === 1 ? '' : 's'}.`, ...describeMerge(res)];
            if (res.unreadable.length > 0) lines.push(`Skipped unreadable: ${res.unreadable.join(', ')}.`);
            setResult({ lines });
        } catch (err) {
            console.error('LDA history workbook scan failed:', err);
            setError(err.message || 'Scan failed.');
        } finally {
            setBusy(null);
        }
    };

    const readFileRows = async (file) => {
        const name = String(file.name || '').toLowerCase();
        if (name.endsWith('.csv')) {
            return parseCsv(await file.text());
        }
        if (name.endsWith('.xlsx') || name.endsWith('.xlsm') || name.endsWith('.xls')) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(await file.arrayBuffer());
            const sheet = workbook.worksheets[0];
            if (!sheet) throw new Error('Workbook has no sheets.');
            const rows = [];
            sheet.eachRow({ includeEmpty: false }, (row) => {
                // row.values is 1-indexed; slot 0 is always empty.
                rows.push(row.values.slice(1).map(excelCellToValue));
            });
            return rows;
        }
        throw new Error('Unsupported file type — upload a .csv or .xlsx file.');
    };

    const handleFileSelect = async (e) => {
        const file = e.target.files && e.target.files[0];
        e.target.value = '';
        if (!file || busy) return;
        setBusy('file');
        setError('');
        setResult(null);
        try {
            const parsed = rowsToImportedDays(await readFileRows(file));
            const merge = importDaysIntoHistory(parsed.days);
            const lines = [`Read ${parsed.rows} row${parsed.rows === 1 ? '' : 's'} from "${file.name}".`, ...describeMerge(merge)];
            if (parsed.skippedRows > 0) lines.push(`${parsed.skippedRows} row${parsed.skippedRows === 1 ? '' : 's'} skipped (missing date, ID, or Days Out).`);
            setResult({ lines });
        } catch (err) {
            console.error('LDA history file import failed:', err);
            setError(err.message || 'File import failed.');
        } finally {
            setBusy(null);
        }
    };

    if (!isOpen) return null;

    const optionClass = 'w-full flex items-start gap-3 p-4 rounded-xl border border-gray-200 hover:border-[#145F82]/60 hover:bg-slate-50 transition-colors text-left disabled:opacity-60 disabled:cursor-not-allowed';

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md mx-4">
                <div className="flex justify-between items-start mb-4">
                    <div>
                        <h3 className="text-lg font-semibold text-gray-800">Import LDA History</h3>
                        <p className="text-xs text-gray-500 mt-0.5">
                            Backfill the Risk Index history. Days that already have a snapshot are never overwritten.
                        </p>
                    </div>
                    <button
                        onClick={onClose}
                        aria-label="Close"
                        className="text-gray-400 hover:text-gray-600 text-xl leading-none"
                    >
                        &times;
                    </button>
                </div>

                <div className="space-y-3">
                    <button
                        type="button"
                        onClick={handleScan}
                        disabled={!!busy}
                        className={optionClass}
                    >
                        <span className="w-9 h-9 rounded-lg bg-[#145F82]/10 text-[#145F82] flex items-center justify-center shrink-0">
                            {busy === 'scan' ? <Loader2 className="w-5 h-5 animate-spin" /> : <FileSearch className="w-5 h-5" />}
                        </span>
                        <span className="min-w-0">
                            <span className="block text-sm font-semibold text-gray-800">
                                {busy === 'scan' ? 'Scanning workbook…' : 'Scan Workbook'}
                            </span>
                            <span className="block text-xs text-gray-500 mt-0.5">
                                Read previous LDA sheets in this workbook (named like "LDA 7-21-2026").
                            </span>
                        </span>
                    </button>

                    <label className={`${optionClass} ${busy ? 'opacity-60 cursor-not-allowed' : 'cursor-pointer'}`}>
                        <input
                            type="file"
                            accept=".csv,.xlsx,.xlsm,.xls"
                            onChange={handleFileSelect}
                            disabled={!!busy}
                            className="hidden"
                        />
                        <span className="w-9 h-9 rounded-lg bg-[#145F82]/10 text-[#145F82] flex items-center justify-center shrink-0">
                            {busy === 'file' ? <Loader2 className="w-5 h-5 animate-spin" /> : <Upload className="w-5 h-5" />}
                        </span>
                        <span className="min-w-0">
                            <span className="block text-sm font-semibold text-gray-800">
                                {busy === 'file' ? 'Importing file…' : 'Upload File'}
                            </span>
                            <span className="block text-xs text-gray-500 mt-0.5">
                                Import a CSV or Excel file with Date, SyStudentID, LDA, and Days Out columns
                                (the Download History format).
                            </span>
                        </span>
                    </label>
                </div>

                {error && (
                    <div className="mt-4 text-sm text-red-600 bg-red-50 border border-red-100 rounded-lg px-3 py-2">
                        {error}
                    </div>
                )}
                {result && (
                    <div className="mt-4 text-sm text-emerald-800 bg-emerald-50 border border-emerald-100 rounded-lg px-3 py-2 space-y-1">
                        {result.lines.map((line, i) => <div key={i}>{line}</div>)}
                    </div>
                )}

                <div className="flex justify-end mt-5">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300 text-sm"
                    >
                        Close
                    </button>
                </div>
            </div>
        </div>
    );
}
