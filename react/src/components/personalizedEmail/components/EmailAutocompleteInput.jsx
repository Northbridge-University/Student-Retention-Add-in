import React, { useRef, useState } from 'react';

// A single-value text input with an email-suggestion dropdown. Matches on email
// OR name (typing a name surfaces the email) and writes the selected email back
// through onChange. Used where a plain email field needs autocomplete (e.g. the
// signature modal's From address) rather than the pill-based PillInput.
export default function EmailAutocompleteInput({
    value,
    onChange,
    suggestions = [],
    placeholder,
    className,
    id,
}) {
    const inputRef = useRef(null);
    const [showSuggestions, setShowSuggestions] = useState(false);
    const [activeIndex, setActiveIndex] = useState(-1);

    const normalized = (suggestions || [])
        .map(s => (typeof s === 'string' ? { name: '', email: s } : { name: s.name || '', email: s.email || '' }))
        .filter(s => s.email);

    const query = (value || '').trim().toLowerCase();
    const matches = query
        ? normalized
            .filter(s => s.email.toLowerCase() !== query) // already entered exactly
            .filter(s => s.email.toLowerCase().includes(query) || s.name.toLowerCase().includes(query))
            .slice(0, 8)
        : [];
    const isOpen = showSuggestions && matches.length > 0;

    const select = (email) => {
        onChange(email);
        setShowSuggestions(false);
        setActiveIndex(-1);
        inputRef.current?.focus();
    };

    const handleKeyDown = (e) => {
        if (!isOpen) return;
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            setActiveIndex(i => Math.min(i + 1, matches.length - 1));
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            setActiveIndex(i => Math.max(i - 1, 0));
        } else if (e.key === 'Enter' && activeIndex >= 0 && matches[activeIndex]) {
            e.preventDefault();
            select(matches[activeIndex].email);
        } else if (e.key === 'Escape') {
            setShowSuggestions(false);
            setActiveIndex(-1);
        }
    };

    return (
        <div className="relative">
            <input
                ref={inputRef}
                id={id}
                type="text"
                value={value}
                onChange={(e) => { onChange(e.target.value); setShowSuggestions(true); setActiveIndex(-1); }}
                onFocus={() => setShowSuggestions(true)}
                onBlur={() => { setShowSuggestions(false); setActiveIndex(-1); }}
                onKeyDown={handleKeyDown}
                placeholder={placeholder}
                className={className}
                autoComplete="off"
                role="combobox"
                aria-expanded={isOpen}
                aria-autocomplete="list"
            />
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
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => select(s.email)}
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
