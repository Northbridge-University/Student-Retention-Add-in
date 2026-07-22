import React, { useState, useEffect } from 'react';

export default function PowerAutomateConfigModal({ isOpen, onClose }) {
    const [url, setUrl] = useState('');
    const [status, setStatus] = useState('');
    const [currentConnection, setCurrentConnection] = useState(null);
    const [enabled, setEnabled] = useState(true);
    const [includeReceipt, setIncludeReceipt] = useState(false);
    const [isLoading, setIsLoading] = useState(false);

    useEffect(() => {
        if (isOpen) {
            loadCurrentConnection();
        }
    }, [isOpen]);

    const loadCurrentConnection = async () => {
        setIsLoading(true);
        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
                const connection = connections.find(c => c.type === 'power-automate' && c.name === 'Send Personalized Email');

                if (connection) {
                    setCurrentConnection(connection);
                    setEnabled(connection.enabled !== false);
                    setIncludeReceipt(connection.includeReceipt === true);
                    setUrl(''); // Don't expose the URL for security
                    setStatus('Connection configured');
                } else {
                    setCurrentConnection(null);
                    setEnabled(true);
                    setIncludeReceipt(false);
                    setUrl('');
                    setStatus('No connection configured');
                }
            });
        } catch (error) {
            console.error('Error loading connection:', error);
            setStatus('Error loading connection');
        } finally {
            setIsLoading(false);
        }
    };

    const isValidHttpUrl = (string) => {
        try {
            const url = new URL(string);
            return url.protocol === "http:" || url.protocol === "https:";
        } catch {
            return false;
        }
    };

    const handleToggleEnabled = async (newEnabled) => {
        if (!currentConnection) return;
        setEnabled(newEnabled);

        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
                const updated = connections.map(c =>
                    (c.type === 'power-automate' && c.name === 'Send Personalized Email')
                        ? { ...c, enabled: newEnabled }
                        : c
                );
                settings.add("connections", JSON.stringify(updated));
                await context.sync();

                setCurrentConnection(prev => prev ? { ...prev, enabled: newEnabled } : prev);
                setStatus(newEnabled
                    ? 'Power Automate enabled — Send mode active.'
                    : 'Power Automate paused — Download mode active.');
            });
        } catch (error) {
            console.error('Error toggling enabled state:', error);
            setStatus('Error saving toggle');
            setEnabled(!newEnabled);
        }
    };

    const handleToggleIncludeReceipt = async (newIncludeReceipt) => {
        if (!currentConnection) return;
        setIncludeReceipt(newIncludeReceipt);

        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                const connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
                const updated = connections.map(c =>
                    (c.type === 'power-automate' && c.name === 'Send Personalized Email')
                        ? { ...c, includeReceipt: newIncludeReceipt }
                        : c
                );
                settings.add("connections", JSON.stringify(updated));
                await context.sync();

                setCurrentConnection(prev => prev ? { ...prev, includeReceipt: newIncludeReceipt } : prev);
                setStatus(newIncludeReceipt
                    ? 'Receipts will be included in the payload.'
                    : 'Receipts will not be included in the payload.');
            });
        } catch (error) {
            console.error('Error toggling include receipt:', error);
            setStatus('Error saving toggle');
            setIncludeReceipt(!newIncludeReceipt);
        }
    };

    const handleSave = async () => {
        // If no new URL entered but connection exists, just close (no changes)
        if (!url.trim() && currentConnection) {
            onClose();
            return;
        }

        if (!isValidHttpUrl(url)) {
            setStatus('Please enter a valid HTTP URL.');
            return;
        }

        setStatus('Saving connection...');

        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];

                // Remove existing connection if any
                connections = connections.filter(c => !(c.type === 'power-automate' && c.name === 'Send Personalized Email'));

                // Add new connection (preserve enabled flag from current state)
                const newConnection = {
                    id: currentConnection?.id || ('pa-' + Math.random().toString(36).substr(2, 9)),
                    name: 'Send Personalized Email',
                    type: 'power-automate',
                    url: url,
                    enabled: enabled,
                    includeReceipt: includeReceipt,
                    actions: [],
                    history: []
                };
                connections.push(newConnection);

                settings.add("connections", JSON.stringify(connections));
                await context.sync();

                setCurrentConnection(newConnection);
                setStatus('Connection saved successfully!');
                setTimeout(() => {
                    onClose();
                }, 1500);
            });
        } catch (error) {
            console.error('Error saving connection:', error);
            setStatus('Error saving connection');
        }
    };

    const handleDelete = async () => {
        if (!currentConnection) return;

        setStatus('Deleting connection...');

        try {
            await Excel.run(async (context) => {
                const settings = context.workbook.settings;
                const connectionsSetting = settings.getItemOrNullObject("connections");
                connectionsSetting.load("value");
                await context.sync();

                let connections = connectionsSetting.value ? JSON.parse(connectionsSetting.value) : [];
                connections = connections.filter(c => !(c.type === 'power-automate' && c.name === 'Send Personalized Email'));

                settings.add("connections", JSON.stringify(connections));
                await context.sync();

                setCurrentConnection(null);
                setUrl('');
                setStatus('Connection deleted successfully!');
                setTimeout(() => {
                    onClose();
                }, 1500);
            });
        } catch (error) {
            console.error('Error deleting connection:', error);
            setStatus('Error deleting connection');
        }
    };

    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <div className="flex justify-between items-center mb-4">
                    <h3 className="text-lg font-semibold text-gray-800">Configure Power Automate</h3>
                    <button
                        onClick={onClose}
                        className="text-gray-400 hover:text-gray-600"
                        aria-label="Close"
                    >
                        <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                    </button>
                </div>

                {isLoading ? (
                    <div className="flex justify-center items-center py-8">
                        <div className="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-blue-600"></div>
                    </div>
                ) : (
                    <>
                        {currentConnection && (
                            <div className="mb-4 p-3 rounded-md bg-gray-50 border border-gray-200">
                                <label htmlFor="pa-enabled-toggle" className="flex items-center justify-between cursor-pointer">
                                    <span className="text-sm">
                                        <span className="font-medium text-gray-700 block">Use Power Automate</span>
                                        <span className="text-xs text-gray-500">
                                            {enabled
                                                ? 'Send Email button is active.'
                                                : 'Paused — the composer shows the Download button instead.'}
                                        </span>
                                    </span>
                                    <div className="relative inline-flex items-center flex-shrink-0 ml-3">
                                        <input
                                            id="pa-enabled-toggle"
                                            type="checkbox"
                                            checked={enabled}
                                            onChange={(e) => handleToggleEnabled(e.target.checked)}
                                            className="sr-only peer"
                                        />
                                        <div className="w-11 h-6 bg-gray-200 rounded-full peer peer-focus:ring-4 peer-focus:ring-blue-300 peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-0.5 after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
                                    </div>
                                </label>
                            </div>
                        )}

                        {currentConnection && (
                            <div className="mb-4 p-3 rounded-md bg-gray-50 border border-gray-200">
                                <label htmlFor="pa-receipt-toggle" className="flex items-center justify-between cursor-pointer">
                                    <span className="text-sm">
                                        <span className="font-medium text-gray-700 block">Include Receipts</span>
                                        <span className="text-xs text-gray-500">
                                            {includeReceipt
                                                ? 'A PDF receipt is attached to the payload sent to Power Automate.'
                                                : 'No receipt is included in the payload.'}
                                        </span>
                                    </span>
                                    <div className="relative inline-flex items-center flex-shrink-0 ml-3">
                                        <input
                                            id="pa-receipt-toggle"
                                            type="checkbox"
                                            checked={includeReceipt}
                                            onChange={(e) => handleToggleIncludeReceipt(e.target.checked)}
                                            className="sr-only peer"
                                        />
                                        <div className="w-11 h-6 bg-gray-200 rounded-full peer peer-focus:ring-4 peer-focus:ring-blue-300 peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-0.5 after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-blue-600"></div>
                                    </div>
                                </label>
                            </div>
                        )}

                        <div className={`mb-4 transition-opacity ${currentConnection && !enabled ? 'opacity-60' : ''}`}>
                            <label htmlFor="pa-url" className="block text-sm font-medium text-gray-700 mb-2">
                                Power Automate HTTP URL
                            </label>
                            <input
                                type="password"
                                id="pa-url"
                                value={url}
                                onChange={(e) => setUrl(e.target.value)}
                                className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-blue-500 focus:border-blue-500"
                                placeholder={currentConnection ? "••••••••••••••••••••••••••••••••" : "https://prod-xx.eastus.logic.azure.com..."}
                                autoComplete="off"
                            />
                            <p className="mt-2 text-xs text-gray-500">
                                {currentConnection
                                    ? "A URL is configured. Enter a new URL to replace it."
                                    : "This URL is generated by the \"When a HTTP request is received\" trigger in your Power Automate flow."}
                            </p>
                        </div>

                        <p className="text-xs text-center h-4 mb-4" style={{ color: status.includes('Error') ? '#dc2626' : '#16a34a' }}>
                            {status}
                        </p>

                        <div className="flex justify-between gap-2">
                            {currentConnection && (
                                <button
                                    onClick={handleDelete}
                                    className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700"
                                >
                                    Delete
                                </button>
                            )}
                            <div className={`flex gap-2 ${currentConnection ? '' : 'ml-auto'}`}>
                                <button
                                    onClick={onClose}
                                    className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                                >
                                    Cancel
                                </button>
                                <button
                                    onClick={handleSave}
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
    );
}
