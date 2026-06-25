import React, { useState, useEffect } from 'react';
import { renderTemplate, renderCCTemplate, findSignature } from '../utils/helpers';

export default function ExampleModal({ isOpen, onClose, studentData, fromTemplate, ccRecipients, subjectTemplate, bodyTemplate, signatures }) {
    const [currentIndex, setCurrentIndex] = useState(0);
    const [showSearch, setShowSearch] = useState(false);
    const [searchTerm, setSearchTerm] = useState('');
    const [searchResults, setSearchResults] = useState([]);

    useEffect(() => {
        if (isOpen) {
            setCurrentIndex(0);
            setShowSearch(false);
            setSearchTerm('');
            setSearchResults([]);
        }
    }, [isOpen]);

    if (!isOpen || !studentData || studentData.length === 0) return null;

    const student = studentData[currentIndex];
    const from = renderTemplate(fromTemplate, student);
    // Inject the From-specific signature so {Signature} previews correctly.
    const studentWithSignature = { ...student, Signature: findSignature(signatures, from) };
    const to = student.StudentEmail || '[No Email]';
    const cc = renderCCTemplate(ccRecipients, studentWithSignature);
    const subject = renderTemplate(subjectTemplate, studentWithSignature);
    const body = renderTemplate(bodyTemplate, studentWithSignature);

    const handlePrevious = () => {
        if (currentIndex > 0) setCurrentIndex(currentIndex - 1);
    };

    const handleNext = () => {
        if (currentIndex < studentData.length - 1) setCurrentIndex(currentIndex + 1);
    };

    const handleRandom = () => {
        const randomIndex = Math.floor(Math.random() * studentData.length);
        setCurrentIndex(randomIndex);
    };

    const handleSearch = (value) => {
        setSearchTerm(value);
        if (!value.trim()) {
            setSearchResults([]);
            return;
        }

        const matches = studentData
            .map((student, index) => ({ student, index }))
            .filter(item => item.student.StudentName && item.student.StudentName.toLowerCase().includes(value.toLowerCase()))
            .slice(0, 10);
        setSearchResults(matches);
    };

    const handleSelectStudent = (index) => {
        setCurrentIndex(index);
        setShowSearch(false);
        setSearchTerm('');
        setSearchResults([]);
    };

    return (
        <div className="fixed inset-0 bg-gray-600 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-lg">
                <div className="flex justify-between items-center mb-4">
                    <h3 className="text-lg font-semibold text-gray-800">Email Preview</h3>
                    <div className="flex items-center gap-2">
                        <button
                            onClick={handlePrevious}
                            disabled={currentIndex === 0}
                            className="p-1 rounded-full hover:bg-gray-200 disabled:opacity-50 disabled:cursor-not-allowed"
                            title="Previous Student"
                        >
                            <svg className="h-5 w-5 text-gray-600" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor">
                                <path fillRule="evenodd" d="M12.707 5.293a1 1 0 010 1.414L9.414 10l3.293 3.293a1 1 0 01-1.414 1.414l-4-4a1 1 0 010-1.414l4-4a1 1 0 011.414 0z" clipRule="evenodd" />
                            </svg>
                        </button>
                        <span className="text-sm text-gray-600 font-medium w-36 text-center">
                            Student: {currentIndex + 1} / {studentData.length}
                        </span>
                        <button
                            onClick={handleNext}
                            disabled={currentIndex === studentData.length - 1}
                            className="p-1 rounded-full hover:bg-gray-200 disabled:opacity-50 disabled:cursor-not-allowed"
                            title="Next Student"
                        >
                            <svg className="h-5 w-5 text-gray-600" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor">
                                <path fillRule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clipRule="evenodd" />
                            </svg>
                        </button>
                        <button
                            onClick={handleRandom}
                            className="p-1 rounded-full hover:bg-gray-200"
                            title="Random Student"
                        >
                            <svg className="h-5 w-5 text-gray-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" d="M16.023 9.348h4.992v-.001M2.985 19.644v-4.992m0 0h4.992m-4.993 0l3.181 3.183a8.25 8.25 0 0011.667 0l3.182-3.182m0-11.667l-3.182 3.182a8.25 8.25 0 00-11.667 0L2.985 19.644z" />
                            </svg>
                        </button>
                        <button
                            onClick={() => setShowSearch(!showSearch)}
                            className="p-1 rounded-full hover:bg-gray-200"
                            title="Search Student"
                        >
                            <svg className="h-5 w-5 text-gray-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-5.197-5.197m0 0A7.5 7.5 0 105.196 5.196a7.5 7.5 0 0010.607 10.607z" />
                            </svg>
                        </button>
                    </div>
                </div>

                {showSearch && (
                    <div className="relative mb-2">
                        <input
                            type="text"
                            value={searchTerm}
                            onChange={(e) => handleSearch(e.target.value)}
                            className="w-full pl-3 pr-10 py-2 border border-gray-300 rounded-md shadow-sm"
                            placeholder="Search by student name..."
                        />
                        {searchResults.length > 0 && (
                            <div className="absolute z-10 w-full mt-1 bg-white border border-gray-300 rounded-md shadow-lg max-h-40 overflow-y-auto">
                                {searchResults.map((result) => (
                                    <div
                                        key={result.index}
                                        onClick={() => handleSelectStudent(result.index)}
                                        className="px-3 py-2 text-sm text-gray-700 cursor-pointer hover:bg-gray-100"
                                    >
                                        {result.student.StudentName}
                                    </div>
                                ))}
                            </div>
                        )}
                    </div>
                )}

                <div className="space-y-3 text-sm border-t border-b py-4">
                    <div className="flex">
                        <label className="font-semibold text-gray-600 w-20">From:</label>
                        <p className="text-gray-800">{from}</p>
                    </div>
                    <div className="flex">
                        <label className="font-semibold text-gray-600 w-20">Recipient:</label>
                        <p className="text-gray-800">{to}</p>
                    </div>
                    <div className="flex">
                        <label className="font-semibold text-gray-600 w-20">CC:</label>
                        <p className="text-gray-800">{cc}</p>
                    </div>
                    <div className="flex">
                        <label className="font-semibold text-gray-600 w-20">Subject:</label>
                        <p className="text-gray-800 font-medium">{subject}</p>
                    </div>
                    <div>
                        <label className="font-semibold text-gray-600">Body:</label>
                        <div
                            className="text-gray-800 border rounded-md p-2 bg-gray-50 max-h-64 overflow-y-auto w-full mt-1"
                            dangerouslySetInnerHTML={{ __html: body }}
                        />
                    </div>
                </div>

                <div className="flex justify-end mt-4">
                    <button
                        onClick={onClose}
                        className="px-4 py-2 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300"
                    >
                        Close
                    </button>
                </div>
            </div>
        </div>
    );
}
