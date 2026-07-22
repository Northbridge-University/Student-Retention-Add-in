import { describe, it, expect } from 'vitest';
import {
    dataUriByteLength,
    formatBytes,
    IMAGE_COMPRESS_THRESHOLD_BYTES,
} from '../imageCompression';

describe('dataUriByteLength', () => {
    it('returns 0 for empty / invalid input', () => {
        expect(dataUriByteLength('')).toBe(0);
        expect(dataUriByteLength(null)).toBe(0);
        expect(dataUriByteLength(undefined)).toBe(0);
        expect(dataUriByteLength(12345)).toBe(0);
    });

    it('computes decoded byte length from a data-URI, accounting for padding', () => {
        // "AAAA" (no padding) -> 3 bytes
        expect(dataUriByteLength('data:image/png;base64,AAAA')).toBe(3);
        // "AAA=" -> 2 bytes, "AA==" -> 1 byte
        expect(dataUriByteLength('data:image/png;base64,AAA=')).toBe(2);
        expect(dataUriByteLength('data:image/png;base64,AA==')).toBe(1);
    });

    it('accepts a bare base64 string (no data-URI prefix)', () => {
        expect(dataUriByteLength('AAAA')).toBe(3);
    });

    it('scales with payload length', () => {
        // 1000 groups of 4 base64 chars, no padding -> 3000 bytes
        expect(dataUriByteLength(`data:image/jpeg;base64,${'AAAA'.repeat(1000)}`)).toBe(3000);
    });
});

describe('formatBytes', () => {
    it('formats common sizes', () => {
        expect(formatBytes(0)).toBe('0 B');
        expect(formatBytes(-5)).toBe('0 B');
        expect(formatBytes(512)).toBe('512 B');
        expect(formatBytes(1024)).toBe('1 KB');
        expect(formatBytes(1536)).toBe('1.5 KB');
        expect(formatBytes(5 * 1024 * 1024)).toBe('5 MB');
    });
});

describe('IMAGE_COMPRESS_THRESHOLD_BYTES', () => {
    it('is a sensible positive threshold', () => {
        expect(IMAGE_COMPRESS_THRESHOLD_BYTES).toBeGreaterThan(0);
    });
});
