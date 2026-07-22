import React, { useRef, useState } from 'react';

export default function PillInput({
    pills,
    onPillsChange,
    placeholder,
    singleValue = false,
    onFocus,
    readOnly = false,
    noWrap = false,
    getPillColor,
    suggestions = []
}) {
    const inputRef = useRef(null);
    const containerRef = useRef(null);
    const [inputValue, setInputValue] = useState('');
    const [showSuggestions, setShowSuggestions] = useState(false);
    const [activeIndex, setActiveIndex] = useState(-1);

    // Accept either { name, email } objects or plain email strings.
    const normalizedSuggestions = (suggestions || [])
        .map(s => (typeof s === 'string' ? { name: '', email: s } : { name: s.name || '', email: s.email || '' }))
        .filter(s => s.email);

    const pillSet = new Set(pills.map(p => p.toLowerCase()));
    const query = inputValue.trim().toLowerCase();
    // Match on email OR name, so typing someone's name surfaces their email.
    const matches = query
        ? normalizedSuggestions
            .filter(s => !pillSet.has(s.email.toLowerCase()))
            .filter(s => s.email.toLowerCase().includes(query) || s.name.toLowerCase().includes(query))
            .slice(0, 8)
        : [];
    const isOpen = showSuggestions && matches.length > 0;

    const addPill = (text) => {
        if (!text || !text.trim()) return;
        const value = text.trim();
        if (singleValue) {
            onPillsChange([value]);
        } else if (!pills.some(p => p.toLowerCase() === value.toLowerCase())) {
            onPillsChange([...pills, value]);
        }
        setInputValue('');
        setActiveIndex(-1);
        setShowSuggestions(false);
    };

    const removePill = (index) => {
        const newPills = pills.filter((_, i) => i !== index);
        onPillsChange(newPills);
    };

    const handleKeyDown = (e) => {
        if (isOpen && e.key === 'ArrowDown') {
            e.preventDefault();
            setActiveIndex(i => Math.min(i + 1, matches.length - 1));
            return;
        }
        if (isOpen && e.key === 'ArrowUp') {
            e.preventDefault();
            setActiveIndex(i => Math.max(i - 1, 0));
            return;
        }
        if (e.key === 'Escape') {
            setShowSuggestions(false);
            setActiveIndex(-1);
            return;
        }

        if (e.key === ',' || e.key === 'Enter' || e.key === ';') {
            e.preventDefault();
            // If a suggestion is highlighted, take it; otherwise use what was typed.
            if (isOpen && activeIndex >= 0 && matches[activeIndex]) {
                addPill(matches[activeIndex].email);
            } else {
                addPill(inputValue);
            }
        } else if (e.key === 'Backspace' && inputValue === '' && pills.length > 0) {
            removePill(pills.length - 1);
        }
    };

    const handleChange = (e) => {
        setInputValue(e.target.value);
        setShowSuggestions(true);
        setActiveIndex(-1);
    };

    const handleBlur = () => {
        // A suggestion click uses onMouseDown (preventDefault) so it runs before
        // this; clicking elsewhere commits whatever was typed, matching the old
        // behavior.
        if (inputValue.trim()) addPill(inputValue);
        setShowSuggestions(false);
        setActiveIndex(-1);
    };

    const handleFocus = (e) => {
        if (onFocus) onFocus(e);
        setShowSuggestions(true);
    };

    const handleContainerClick = () => {
        inputRef.current?.focus();
    };

    return (
        <div className="relative" ref={containerRef}>
            <div
                onClick={readOnly ? undefined : handleContainerClick}
                className={`flex items-center gap-1.5 p-1.5 border border-gray-300 rounded-md ${
                    noWrap ? 'overflow-x-auto' : 'flex-wrap'
                } ${readOnly ? 'bg-gray-100 cursor-default' : 'bg-white cursor-text'}`}
            >
                {pills.map((pill, index) => {
                    const isParam = pill.startsWith('{') && pill.endsWith('}');
                    const customColor = isParam && getPillColor ? getPillColor(pill.slice(1, -1)) : null;
                    return (
                        <span
                            key={index}
                            style={customColor ? { backgroundColor: customColor.background, color: customColor.color } : undefined}
                            className={`flex items-center rounded-full px-2 py-0.5 text-sm ${
                                customColor
                                    ? ''
                                    : isParam
                                        ? 'bg-blue-100 text-blue-800'
                                        : 'bg-gray-200 text-gray-700'
                            }`}
                        >
                            {pill}
                            {!readOnly && (
                                <span
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        removePill(index);
                                    }}
                                    className="ml-1.5 cursor-pointer font-bold hover:text-red-600"
                                >
                                    ×
                                </span>
                            )}
                        </span>
                    );
                })}
                {!readOnly && (
                    <input
                        ref={inputRef}
                        type="text"
                        value={inputValue}
                        className={`flex-grow border-none outline-none p-1 bg-transparent ${
                            noWrap ? 'min-w-[80px]' : 'min-w-[120px]'
                        }`}
                        placeholder={pills.length === 0 ? placeholder : ''}
                        onChange={handleChange}
                        onKeyDown={handleKeyDown}
                        onBlur={handleBlur}
                        onFocus={handleFocus}
                        autoComplete="off"
                        role="combobox"
                        aria-expanded={isOpen}
                        aria-autocomplete="list"
                    />
                )}
            </div>

            {isOpen && (
                <ul
                    className="absolute left-0 right-0 top-full mt-1 z-50 bg-white border border-gray-200 rounded-md shadow-lg max-h-48 overflow-y-auto py-1"
                    role="listbox"
                >
                    {matches.map((s, idx) => (
                        <li
                            key={s.email}
                            role="option"
                            aria-selected={idx === activeIndex}
                            // Keep focus so onBlur doesn't fire before the click selects.
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => {
                                addPill(s.email);
                                inputRef.current?.focus();
                            }}
                            onMouseEnter={() => setActiveIndex(idx)}
                            className={`px-3 py-1.5 cursor-pointer ${idx === activeIndex ? 'bg-blue-50' : 'hover:bg-gray-100'}`}
                        >
                            {s.name
                                ? (
                                    <div className="flex flex-col leading-tight">
                                        <span className="text-sm text-gray-800">{s.name}</span>
                                        <span className="text-xs text-gray-500">{s.email}</span>
                                    </div>
                                )
                                : <span className="text-sm text-gray-800">{s.email}</span>}
                        </li>
                    ))}
                </ul>
            )}
        </div>
    );
}
