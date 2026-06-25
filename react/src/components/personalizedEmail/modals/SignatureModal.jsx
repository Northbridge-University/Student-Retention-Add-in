import React, { useState, useEffect } from 'react';
import ReactQuill from 'react-quill-new';
import 'react-quill-new/dist/quill.snow.css';
import { QUILL_EDITOR_CONFIG } from '../utils/constants';
import { normalizeEmailKey } from '../utils/helpers';
import { compressImagesInHtml } from '../utils/imageCompression';
import EmailAutocompleteInput from '../components/EmailAutocompleteInput';

// Manages the per-From-address signatures that power the {Signature} special
// parameter. Each entry maps a From email to a rich-text signature; at send
// time {Signature} resolves to the signature assigned to that email's From.
export default function SignatureModal({ isOpen, onClose, signatures, onSave, currentFrom, suggestions = [] }) {
    const [entries, setEntries] = useState([]);
    const [saveStatus, setSaveStatus] = useState('');

    useEffect(() => {
        if (isOpen) {
            // Work on a copy so Cancel discards unsaved edits.
            const copy = (signatures || []).map(s => ({ email: s.email || '', signature: s.signature || '' }));
            // Offer the current From address as a starter row when nothing is set up yet.
            if (copy.length === 0 && currentFrom && currentFrom.trim()) {
                copy.push({ email: currentFrom.trim(), signature: '' });
            }
            setEntries(copy);
            setSaveStatus('');
        }
    }, [isOpen, signatures, currentFrom]);

    const handleEntryChange = (index, field, value) => {
        setEntries(prev => prev.map((entry, i) => (i === index ? { ...entry, [field]: value } : entry)));
    };

    const handleAddEntry = () => {
        setEntries(prev => [...prev, { email: '', signature: '' }]);
    };

    const handleRemoveEntry = (index) => {
        setEntries(prev => prev.filter((_, i) => i !== index));
    };

    const handleSave = async () => {
        const base = entries
            .map(e => ({ email: (e.email || '').trim(), signature: e.signature || '' }))
            .filter(e => e.email);

        // Block duplicate From addresses (same normalized email).
        const keys = base.map(e => normalizeEmailKey(e.email));
        if (new Set(keys).size !== keys.length) {
            setSaveStatus('Two signatures use the same From address. Remove the duplicate.');
            return;
        }

        setSaveStatus('Saving…');
        // Compress oversized images so a signature (repeated across every
        // recipient) doesn't bloat the email payload.
        const cleaned = await Promise.all(
            base.map(async (e) => ({ email: e.email, signature: await compressImagesInHtml(e.signature) }))
        );

        await onSave(cleaned);
        setSaveStatus('Signatures saved!');
        setTimeout(() => {
            setSaveStatus('');
            onClose();
        }, 1000);
    };

    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-start justify-center z-50 overflow-y-auto p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-2xl my-8 max-h-[calc(100vh-4rem)] overflow-y-auto">
                <div className="flex justify-between items-center mb-2">
                    <h3 className="text-lg font-semibold text-gray-800">Email Signatures</h3>
                    <button onClick={onClose} className="text-gray-400 hover:text-gray-600" aria-label="Close">
                        <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                    </button>
                </div>
                <p className="text-xs text-gray-500 mb-4">
                    Assign a signature to each From address. Insert <span className="font-mono text-orange-700">{'{Signature}'}</span> in
                    your email and it will be replaced with the signature for whichever From address is used. Emails can&apos;t be sent
                    from a From address that has no signature.
                </p>

                <div className="space-y-4">
                    {entries.length === 0 ? (
                        <p className="text-center text-gray-500 text-sm py-4">No signatures yet. Add one below.</p>
                    ) : (
                        entries.map((entry, index) => (
                            <div key={index} className="border border-gray-200 rounded-md p-3 bg-gray-50">
                                <div className="flex items-center justify-between mb-2">
                                    <span className="text-xs font-medium text-gray-600">Signature {index + 1}</span>
                                    <button
                                        onClick={() => handleRemoveEntry(index)}
                                        aria-label={`Remove signature ${index + 1}`}
                                        title="Remove"
                                        className="text-red-500 hover:text-red-700 text-xl leading-none"
                                    >
                                        ×
                                    </button>
                                </div>
                                <label className="block text-sm font-medium text-gray-700 mb-1">From address</label>
                                <EmailAutocompleteInput
                                    value={entry.email}
                                    onChange={(value) => handleEntryChange(index, 'email', value)}
                                    suggestions={suggestions}
                                    placeholder="Type a name or email, e.g. advisor@school.edu"
                                    className="block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm text-sm bg-white"
                                />
                                <label className="block text-sm font-medium text-gray-700 mt-3 mb-1">Signature</label>
                                <div className="bg-white rounded-md">
                                    <ReactQuill
                                        theme="snow"
                                        value={entry.signature}
                                        onChange={(value) => handleEntryChange(index, 'signature', value)}
                                        modules={QUILL_EDITOR_CONFIG.modules}
                                        placeholder="Your signature (formatting and pasted images supported)..."
                                    />
                                </div>
                            </div>
                        ))
                    )}
                </div>

                <button
                    onClick={handleAddEntry}
                    className="mt-3 text-sm text-blue-600 hover:underline"
                >
                    + Add Signature
                </button>

                <p className="text-xs mt-2 h-4 text-center" style={{ color: saveStatus.includes('saved') ? '#16a34a' : '#dc2626' }}>
                    {saveStatus}
                </p>

                <div className="flex justify-end gap-2 mt-4 border-t pt-4">
                    <button
                        onClick={onClose}
                        className="px-3 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300 text-sm"
                    >
                        Cancel
                    </button>
                    <button
                        onClick={handleSave}
                        className="px-3 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 text-sm"
                    >
                        Save
                    </button>
                </div>
            </div>
        </div>
    );
}
