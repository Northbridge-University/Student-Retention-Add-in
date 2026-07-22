import React from 'react';
import { generatePdfReceipt } from '../utils/receiptGenerator';

export default function SuccessModal({ isOpen, onClose, count, payload, bodyTemplate, initiator }) {
    if (!isOpen) return null;

    const handleDownloadReceipt = async () => {
        // payload now has structure { byName, byEmail, emails }, so pass emails array
        const emails = payload?.emails || payload;
        await generatePdfReceipt(emails, bodyTemplate, initiator || {});
    };

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md text-center">
                <div className="mx-auto flex items-center justify-center h-12 w-12 rounded-full bg-green-100">
                    <svg className="h-6 w-6 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" />
                    </svg>
                </div>
                <h3 className="text-lg font-semibold text-gray-800 mt-4">Emails Sent Successfully!</h3>
                <p className="text-sm text-gray-600 mt-2">
                    Your batch of {count} {count === 1 ? 'email' : 'emails'} has been successfully sent to Power Automate.
                </p>
                <div className="mt-6 flex flex-col space-y-2">
                    <button
                        onClick={handleDownloadReceipt}
                        className="w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-lg hover:bg-blue-700"
                    >
                        Download PDF Receipt
                    </button>
                    <button
                        onClick={onClose}
                        className="w-full bg-gray-200 text-gray-700 py-2 px-4 rounded-lg hover:bg-gray-300"
                    >
                        Close
                    </button>
                </div>
            </div>
        </div>
    );
}
