import React, { useState, useEffect } from 'react';
import { User, RefreshCw } from 'lucide-react';
import { formatName } from '../utility/Conversion';
import { MASTER_LIST_SHEET } from '../../../../shared/constants.js';
import {
    getRiskIndexConfig,
    loadLdaHistory,
    countRiskEpisodes,
    buildRiskLeaderboard,
} from '../createLDA/riskIndex.js';

/* global Excel */

// Same gender-based avatar treatment as the student view header.
const avatarBg = (gender) => {
    const g = String(gender || '').toLowerCase();
    if (g === 'male') return { background: '#1278FF', color: '#FFFFFF' };
    if (g === 'female') return { background: '#ed72b5', color: '#FFFFFF' };
    return { background: '#6b7280', color: '#FFFFFF' };
};

function RowAvatar({ student }) {
    const [imageError, setImageError] = useState(false);
    useEffect(() => { setImageError(false); }, [student.profilePicture]);

    if (student.profilePicture && !imageError) {
        return (
            <img
                src={student.profilePicture}
                alt={student.name ? `${student.name} profile` : 'Student profile'}
                onError={() => setImageError(true)}
                style={{ width: 28, height: 28, borderRadius: '50%', objectFit: 'cover', display: 'block' }}
            />
        );
    }
    return (
        <div style={{
            width: 28, height: 28, borderRadius: '50%',
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            ...avatarBg(student.gender),
        }}>
            <User size={16} strokeWidth={2} aria-hidden="true" />
        </div>
    );
}

/**
 * Ranked list of students by Risk Index (times they hit the risk threshold),
 * greatest first. Students with no hits are excluded, and so are ids no longer
 * on the Master List — those are assumed dropped.
 */
export default function RiskIndexLeaderboard({ refreshToken }) {
    const [students, setStudents] = useState(null); // null = loading
    const [threshold, setThreshold] = useState(null);
    const [error, setError] = useState('');

    const load = async () => {
        setError('');
        try {
            const cfg = getRiskIndexConfig();
            setThreshold(cfg.threshold);
            const counts = countRiskEpisodes(loadLdaHistory(), cfg.threshold);
            if (counts.size === 0) {
                setStudents([]);
                return;
            }
            if (typeof window === 'undefined' || !window.Excel || !Excel.run) {
                throw new Error('Excel API not available');
            }
            const masterValues = await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItemOrNullObject(MASTER_LIST_SHEET);
                await context.sync();
                if (sheet.isNullObject) return null;
                const used = sheet.getUsedRangeOrNullObject();
                used.load('values');
                await context.sync();
                return (used.isNullObject || !used.values) ? null : used.values;
            });
            if (!masterValues) {
                throw new Error(`No "${MASTER_LIST_SHEET}" sheet found — names can't be looked up.`);
            }
            setStudents(buildRiskLeaderboard(counts, masterValues));
        } catch (err) {
            console.error('Failed to load Risk Index leaderboard:', err);
            setError(err.message || 'Failed to load at-risk students');
            setStudents([]);
        }
    };

    useEffect(() => {
        setStudents(null);
        load();
        // Reload on mount and whenever the parent signals fresh history
        // (e.g. the import modal just closed).
    }, [refreshToken]);

    const loading = students === null;

    return (
        <div style={{ display: 'grid', gap: 8, width: '100%' }}>
            <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', padding: '0 2px' }}>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: 6 }}>
                    <div style={{ fontSize: 14, fontWeight: 600 }}>At-Risk Students</div>
                    {threshold != null && (
                        <div style={{ fontSize: 11, color: '#9ca3af' }} title={`Times each student hit ${threshold}+ days out`}>
                            hit {threshold}+ days out
                        </div>
                    )}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    {!loading && students.length > 0 && (
                        <div style={{ fontSize: 12, color: '#6b7280' }}>
                            {students.length} {students.length === 1 ? 'student' : 'students'}
                        </div>
                    )}
                    <button
                        type="button"
                        onClick={() => { if (!loading) { setStudents(null); load(); } }}
                        title="Refresh"
                        aria-label="Refresh at-risk students"
                        style={{
                            border: 'none', background: 'transparent', color: '#6b7280',
                            cursor: 'pointer', padding: 2, display: 'inline-flex', alignItems: 'center',
                        }}
                    >
                        <RefreshCw size={13} className={loading ? 'animate-spin' : undefined} />
                    </button>
                </div>
            </div>

            {loading ? (
                <div style={{
                    padding: '12px 14px', border: '1px dashed #e5e7eb', borderRadius: 8,
                    background: '#fff', color: '#6b7280', fontSize: 13, textAlign: 'center',
                }}>
                    Loading…
                </div>
            ) : error ? (
                <div style={{
                    padding: '12px 14px', border: '1px solid #fecaca', borderRadius: 8,
                    background: '#fef2f2', color: '#b91c1c', fontSize: 13,
                }}>
                    {error}
                </div>
            ) : students.length === 0 ? (
                <div style={{
                    padding: '12px 14px', border: '1px dashed #e5e7eb', borderRadius: 8,
                    background: '#fff', color: '#6b7280', fontSize: 13, textAlign: 'center',
                }}>
                    No students have hit the risk threshold yet. History builds as LDA sheets are created,
                    or use Import Past Reports.
                </div>
            ) : (
                <ul style={{
                    listStyle: 'none', margin: 0, padding: 0, display: 'grid', gap: 6,
                    maxHeight: 320, overflowY: 'auto',
                }}>
                    {students.map((s, i) => {
                        const displayName = s.name
                            ? (s.name.includes(',') ? formatName(s.name) : s.name)
                            : `Student ${s.id}`;
                        return (
                            <li
                                key={s.id}
                                style={{
                                    display: 'grid',
                                    gridTemplateColumns: '20px 28px 1fr auto',
                                    alignItems: 'center',
                                    gap: 10,
                                    padding: '8px 12px',
                                    background: '#fff',
                                    border: '1px solid #e5e7eb',
                                    borderRadius: 8,
                                }}
                            >
                                <div style={{ fontSize: 11, fontWeight: 600, color: '#9ca3af', textAlign: 'right' }}>
                                    {i + 1}
                                </div>
                                <RowAvatar student={s} />
                                <div style={{ minWidth: 0, display: 'grid', gap: 1 }}>
                                    <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: 13, fontWeight: 500 }}>
                                        {displayName}
                                    </div>
                                    <div style={{ fontSize: 11, color: '#6b7280' }}>
                                        ID {s.id}
                                    </div>
                                </div>
                                <div
                                    title={`Hit ${threshold}+ days out ${s.count} ${s.count === 1 ? 'time' : 'times'}`}
                                    style={{
                                        minWidth: 26, height: 22, padding: '0 8px', borderRadius: 9999,
                                        background: s.count >= 3 ? '#991b1b' : s.count === 2 ? '#dc2626' : '#fecaca',
                                        color: s.count >= 2 ? '#ffffff' : '#991b1b',
                                        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                        fontSize: 12, fontWeight: 700, justifySelf: 'end',
                                    }}
                                >
                                    {s.count}
                                </div>
                            </li>
                        );
                    })}
                </ul>
            )}
        </div>
    );
}
