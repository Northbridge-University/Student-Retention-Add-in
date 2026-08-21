import React, { useState } from 'react';
import { EMAIL_TEMPLATES_KEY } from '../utils/constants';

// Addresses permitted to set a custom author name when saving a template.
// Everyone else gets the Author field fixed to their own display name.
//
// This was previously hardcoded to a single address on the pre-rebrand
// ftccollege.edu domain, which meant the capability had already stopped
// working before it was replaced here. Entries must be lowercase.
// Set to [] to remove the custom-author capability entirely.
const AUTHOR_OVERRIDE_ALLOWLIST = ['dvann@northbridge.edu'];

export default function TemplatesModal({ isOpen, onClose, onLoadTemplate, user, userEmail, currentFrom, currentSubject, currentBody, currentCC, templates, onTemplatesChange }) {
    const [expandedAuthors, setExpandedAuthors] = useState(new Set());
    const [showSaveModal, setShowSaveModal] = useState(false);
    const [editingTemplate, setEditingTemplate] = useState(null);
    const [templateName, setTemplateName] = useState('');
    const [customAuthor, setCustomAuthor] = useState('');
    const [saveStatus, setSaveStatus] = useState('');
    const [confirmOverwrite, setConfirmOverwrite] = useState(null);
    const [lastLoadedName, setLastLoadedName] = useState('');

    const isGuest = user === 'Guest';
    const canEditAuthor = AUTHOR_OVERRIDE_ALLOWLIST.includes(userEmail?.toLowerCase() ?? '');

    const saveTemplates = async (templatesArray) => {
        await Excel.run(async (context) => {
            context.workbook.settings.add(EMAIL_TEMPLATES_KEY, JSON.stringify(templatesArray));
            await context.sync();
        });
        onTemplatesChange(templatesArray);
    };

    const toggleAuthor = (author) => {
        const newExpanded = new Set(expandedAuthors);
        if (newExpanded.has(author)) {
            newExpanded.delete(author);
        } else {
            newExpanded.add(author);
        }
        setExpandedAuthors(newExpanded);
    };

    const handleLoadTemplate = (template) => {
        setLastLoadedName(template.name);
        onLoadTemplate(template);
        onClose();
    };

    const handleEditTemplate = (template) => {
        setEditingTemplate(template);
        setTemplateName(template.name);
        setCustomAuthor(template.author || user || 'Unknown');
        setShowSaveModal(true);
    };

    const handleDeleteTemplate = async (template) => {
        const updatedTemplates = templates.filter(t => t.id !== template.id);
        await saveTemplates(updatedTemplates);
        setShowSaveModal(false);
    };

    const handleSaveTemplate = async (forceOverwrite = false) => {
        if (!templateName.trim()) {
            setSaveStatus('Template name is required.');
            return;
        }

        const authorToUse = canEditAuthor ? (customAuthor.trim() || user || 'Unknown') : (user || 'Unknown');

        // Check for existing template with same name + author (only for new saves, not edits)
        if (!editingTemplate && !forceOverwrite) {
            const existing = templates.find(
                t => t.name.toLowerCase() === templateName.trim().toLowerCase() && t.author === authorToUse
            );
            if (existing) {
                setConfirmOverwrite(existing);
                return;
            }
        }

        const templateData = {
            name: templateName.trim(),
            author: authorToUse,
            from: currentFrom,
            subject: currentSubject,
            body: currentBody,
            cc: currentCC || []
        };

        let updatedTemplates = [...templates];
        if (editingTemplate) {
            const index = updatedTemplates.findIndex(t => t.id === editingTemplate.id);
            if (index !== -1) {
                updatedTemplates[index] = { ...updatedTemplates[index], ...templateData };
            }
        } else if (forceOverwrite && confirmOverwrite) {
            const index = updatedTemplates.findIndex(t => t.id === confirmOverwrite.id);
            if (index !== -1) {
                updatedTemplates[index] = { ...updatedTemplates[index], ...templateData };
            }
        } else {
            updatedTemplates.push({
                ...templateData,
                id: 'tpl-' + Date.now(),
                createdAt: new Date().toISOString()
            });
        }

        await saveTemplates(updatedTemplates);
        setConfirmOverwrite(null);
        setSaveStatus('Template saved successfully!');
        setTimeout(() => {
            setShowSaveModal(false);
            setSaveStatus('');
        }, 1500);
    };

    if (!isOpen) return null;

    // Group templates by author
    const groupedTemplates = templates.reduce((acc, template) => {
        const author = template.author || 'Uncategorized';
        if (!acc[author]) acc[author] = [];
        acc[author].push(template);
        return acc;
    }, {});

    return (
        <>
            <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
                <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-xl">
                    <h3 className="text-lg font-semibold text-gray-800 mb-4">Email Templates</h3>

                    <div className="space-y-2 max-h-80 overflow-y-auto border-t border-b py-4">
                        {templates.length === 0 ? (
                            <p className="text-center text-gray-500 text-sm">No saved templates yet.</p>
                        ) : (
                            Object.keys(groupedTemplates).sort().map(author => (
                                <div key={author}>
                                    <div
                                        onClick={() => toggleAuthor(author)}
                                        className="flex items-center justify-between p-2 rounded-md hover:bg-gray-100 cursor-pointer"
                                    >
                                        <div className="flex items-center space-x-2">
                                            <svg className="h-5 w-5 text-gray-500" viewBox="0 0 20 20" fill="currentColor">
                                                <path d="M2 6a2 2 0 012-2h5l2 2h5a2 2 0 012 2v6a2 2 0 01-2 2H4a2 2 0 01-2-2V6z"></path>
                                            </svg>
                                            <span className="font-semibold text-gray-700">{author}</span>
                                        </div>
                                        <svg
                                            className={`h-5 w-5 text-gray-500 transition-transform ${
                                                expandedAuthors.has(author) ? 'rotate-90' : ''
                                            }`}
                                            viewBox="0 0 20 20"
                                            fill="currentColor"
                                        >
                                            <path d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z"></path>
                                        </svg>
                                    </div>

                                    {expandedAuthors.has(author) && (
                                        <div className="pl-6 border-l-2 border-gray-200 ml-2">
                                            {groupedTemplates[author].map(template => (
                                                <div
                                                    key={template.id}
                                                    className="flex items-center justify-between p-2 my-1 rounded-md hover:bg-gray-50"
                                                >
                                                    <div className="text-sm font-medium text-gray-800">{template.name}</div>
                                                    <div className="flex space-x-2">
                                                        <button
                                                            onClick={() => handleLoadTemplate(template)}
                                                            className="px-2 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-md hover:bg-blue-200"
                                                        >
                                                            Load
                                                        </button>
                                                        {!isGuest && (
                                                            <button
                                                                onClick={() => handleEditTemplate(template)}
                                                                className="p-1 text-gray-400 hover:text-gray-600 rounded hover:bg-gray-200"
                                                                title="Edit template"
                                                            >
                                                                <svg className="h-4 w-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor">
                                                                    <path strokeLinecap="round" strokeLinejoin="round" d="M16.862 4.487l1.687-1.688a1.875 1.875 0 112.652 2.652L6.832 19.82a4.5 4.5 0 01-1.897 1.13l-2.685.8.8-2.685a4.5 4.5 0 011.13-1.897L16.863 4.487zm0 0L19.5 7.125" />
                                                                </svg>
                                                            </button>
                                                        )}
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                </div>
                            ))
                        )}
                    </div>

                    <div className={`flex ${isGuest || !currentBody?.trim() ? 'justify-end' : 'justify-between'} mt-4`}>
                        {!isGuest && currentBody?.trim() && (
                            <button
                                onClick={() => {
                                    setEditingTemplate(null);
                                    setTemplateName(lastLoadedName);
                                    setCustomAuthor(user || 'Unknown');
                                    setShowSaveModal(true);
                                }}
                                className="px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
                            >
                                Save
                            </button>
                        )}
                        <button
                            onClick={onClose}
                            className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                        >
                            Close
                        </button>
                    </div>
                </div>
            </div>

            {/* Save Template Modal */}
            {showSaveModal && (
                <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-60">
                    <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                        {confirmOverwrite ? (
                            <>
                                <h3 className="text-lg font-semibold text-gray-800 mb-3">Overwrite Template?</h3>
                                <p className="text-sm text-gray-600">
                                    A template named "<span className="font-medium">{confirmOverwrite.name}</span>" already exists. Do you want to replace it with your current email?
                                </p>
                                <div className="flex justify-end gap-2 mt-5">
                                    <button
                                        onClick={() => setConfirmOverwrite(null)}
                                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                                    >
                                        Cancel
                                    </button>
                                    <button
                                        onClick={() => handleSaveTemplate(true)}
                                        className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700"
                                    >
                                        Overwrite
                                    </button>
                                </div>
                            </>
                        ) : (
                            <>
                                <h3 className="text-lg font-semibold text-gray-800 mb-4">
                                    {editingTemplate ? 'Edit Template' : 'Save Template'}
                                </h3>
                                <div className="space-y-4">
                                    <div>
                                        <label htmlFor="template-name" className="block text-sm font-medium text-gray-700">
                                            Template Name
                                        </label>
                                        <input
                                            type="text"
                                            id="template-name"
                                            value={templateName}
                                            onChange={(e) => setTemplateName(e.target.value)}
                                            className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                                            placeholder="e.g., Mid-term Check-in"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-gray-700">
                                            Author
                                        </label>
                                        {canEditAuthor ? (
                                            <input
                                                type="text"
                                                value={customAuthor}
                                                onChange={(e) => setCustomAuthor(e.target.value)}
                                                className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                                                placeholder="Enter author name"
                                            />
                                        ) : (
                                            <div className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm bg-gray-100 text-gray-700">
                                                {user || 'Unknown'}
                                            </div>
                                        )}
                                    </div>
                                </div>
                                <p className="text-xs mt-2 h-4 text-center text-red-600">{saveStatus}</p>
                                <div className="flex justify-between items-center mt-4">
                                    {editingTemplate && (
                                        <button
                                            onClick={() => handleDeleteTemplate(editingTemplate)}
                                            aria-label="Delete template"
                                            title="Delete template"
                                            className="p-2 text-red-600 rounded-md hover:bg-red-50"
                                        >
                                            <svg
                                                className="h-5 w-5"
                                                fill="none"
                                                stroke="currentColor"
                                                viewBox="0 0 24 24"
                                            >
                                                <path
                                                    strokeLinecap="round"
                                                    strokeLinejoin="round"
                                                    strokeWidth={2}
                                                    d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6M1 7h22M9 7V4a1 1 0 011-1h4a1 1 0 011 1v3"
                                                />
                                            </svg>
                                        </button>
                                    )}
                                    <div className="flex gap-2 ml-auto">
                                        <button
                                            onClick={() => {
                                                setShowSaveModal(false);
                                                setSaveStatus('');
                                                setConfirmOverwrite(null);
                                            }}
                                            className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                                        >
                                            Cancel
                                        </button>
                                        <button
                                            onClick={() => handleSaveTemplate()}
                                            className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700"
                                        >
                                            Save
                                        </button>
                                    </div>
                                </div>
                            </>
                        )}
                    </div>
                </div>
            )}
        </>
    );
}
