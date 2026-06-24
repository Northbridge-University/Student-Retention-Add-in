import { describe, it, expect } from 'vitest';
import { isEmojiCluster, tokenizeRichText } from '../emojiText';

describe('isEmojiCluster', () => {
    it('detects pictographic emoji', () => {
        expect(isEmojiCluster('🎉')).toBe(true);
        expect(isEmojiCluster('👍')).toBe(true);
        expect(isEmojiCluster('❤️')).toBe(true);
    });

    it('rejects plain text, digits, and whitespace', () => {
        expect(isEmojiCluster('a')).toBe(false);
        expect(isEmojiCluster('5')).toBe(false);
        expect(isEmojiCluster('#')).toBe(false);
        expect(isEmojiCluster(' ')).toBe(false);
        expect(isEmojiCluster('')).toBe(false);
    });
});

describe('tokenizeRichText', () => {
    it('splits words and spaces', () => {
        expect(tokenizeRichText('hello world')).toEqual([
            { type: 'text', value: 'hello' },
            { type: 'space', value: ' ' },
            { type: 'text', value: 'world' },
        ]);
    });

    it('separates an emoji from adjacent text', () => {
        expect(tokenizeRichText('hi👋there')).toEqual([
            { type: 'text', value: 'hi' },
            { type: 'emoji', value: '👋' },
            { type: 'text', value: 'there' },
        ]);
    });

    it('keeps a ZWJ sequence as a single emoji token', () => {
        const tokens = tokenizeRichText('👨‍👩‍👧');
        const emojiTokens = tokens.filter((t) => t.type === 'emoji');
        expect(emojiTokens).toHaveLength(1);
        expect(emojiTokens[0].value).toBe('👨‍👩‍👧');
    });

    it('handles empty input', () => {
        expect(tokenizeRichText('')).toEqual([]);
        expect(tokenizeRichText(null)).toEqual([]);
    });
});
