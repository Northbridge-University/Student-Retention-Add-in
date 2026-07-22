// Emoji-aware text tokenization for the PDF receipt renderer.
//
// jsPDF's built-in fonts can't draw emoji, so the receipt renderer rasterizes
// each emoji into a small image and places it inline. To do that it first needs
// to split a text run into plain-text words, spaces, and individual emoji
// grapheme clusters. These helpers are pure (no DOM/canvas) so they can be
// unit-tested outside the Office host.

// Extended_Pictographic covers the vast majority of emoji; Regional_Indicator
// covers flag pairs. Plain ASCII — including digits and '#'/'*' — is excluded.
const EMOJI_PATTERN = /(?:\p{Extended_Pictographic}|\p{Regional_Indicator})/u;

/**
 * True if a grapheme cluster is (or begins with) an emoji.
 * @param {string} cluster
 * @returns {boolean}
 */
export function isEmojiCluster(cluster) {
    if (!cluster) return false;
    return EMOJI_PATTERN.test(cluster);
}

function segmentGraphemes(str) {
    if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
        const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
        return Array.from(segmenter.segment(str), (s) => s.segment);
    }
    // Fallback: split on code points (keeps surrogate pairs intact, but won't
    // group ZWJ sequences — acceptable degradation on old runtimes).
    return Array.from(str);
}

/**
 * Split a string into ordered tokens: { type: 'text'|'space'|'emoji', value }.
 * Runs of non-space, non-emoji graphemes are merged into one 'text' token;
 * each emoji grapheme cluster (including ZWJ sequences) is its own token.
 * @param {string} str
 * @returns {Array<{type: string, value: string}>}
 */
export function tokenizeRichText(str) {
    const tokens = [];
    let textBuffer = '';
    const flushText = () => {
        if (textBuffer) {
            tokens.push({ type: 'text', value: textBuffer });
            textBuffer = '';
        }
    };

    for (const grapheme of segmentGraphemes(str || '')) {
        if (grapheme === ' ') {
            flushText();
            tokens.push({ type: 'space', value: ' ' });
        } else if (isEmojiCluster(grapheme)) {
            flushText();
            tokens.push({ type: 'emoji', value: grapheme });
        } else {
            textBuffer += grapheme;
        }
    }
    flushText();
    return tokens;
}
