import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { getTodaysLdaSheetName } from './helpers';
import { tokenizeRichText } from './emojiText';

// Cache of rasterized emoji so the same glyph isn't re-rendered to a canvas
// repeatedly within one receipt. Keyed by `${emoji}@${size}`.
const emojiImageCache = new Map();

/**
 * Rasterize a single emoji to a PNG data-URI using the canvas. In the Office
 * task pane (a Chromium/WebView2 webview) canvas text uses the OS color-emoji
 * font (Segoe UI Emoji / Apple Color Emoji), so this yields full-color emoji,
 * falling back to whatever monochrome glyph the platform provides. Returns null
 * if no canvas is available (e.g. unit tests).
 * @param {string} emoji
 * @param {number} sizePx
 * @returns {string|null}
 */
function emojiToDataUrl(emoji, sizePx) {
    if (typeof document === 'undefined') return null;
    const key = `${emoji}@${sizePx}`;
    if (emojiImageCache.has(key)) return emojiImageCache.get(key);

    try {
        const scale = 3; // render larger then downscale for crisp glyphs
        const dim = Math.max(8, Math.round(sizePx)) * scale;
        const canvas = document.createElement('canvas');
        canvas.width = dim;
        canvas.height = dim;
        const ctx = canvas.getContext('2d');
        ctx.clearRect(0, 0, dim, dim);
        ctx.textBaseline = 'alphabetic';
        ctx.textAlign = 'left';
        ctx.font = `${Math.round(dim * 0.82)}px "Segoe UI Emoji","Apple Color Emoji","Noto Color Emoji","Twemoji Mozilla","EmojiOne Color",sans-serif`;
        ctx.fillText(emoji, dim * 0.04, dim * 0.86);
        const url = canvas.toDataURL('image/png');
        emojiImageCache.set(key, url);
        return url;
    } catch (err) {
        emojiImageCache.set(key, null);
        return null;
    }
}

/**
 * Load an image source (data-URI or URL) into an HTMLImageElement.
 * Resolves to the element once decoded, or null if it can't be loaded.
 * @param {string} src
 * @returns {Promise<HTMLImageElement|null>}
 */
function loadImageElement(src) {
    return new Promise((resolve) => {
        if (typeof Image === 'undefined') {
            resolve(null);
            return;
        }
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = src;
    });
}

/**
 * Pre-load every <img> referenced by the given HTML strings so the (synchronous)
 * renderer can draw them with known dimensions. Returns a Map keyed by src.
 * @param {Array<string>} htmlStrings
 * @returns {Promise<Map<string, {img: HTMLImageElement, dataUri: string, format: string, width: number, height: number}>>}
 */
async function prepareReceiptImageMap(htmlStrings) {
    const map = new Map();
    if (typeof DOMParser === 'undefined') return map;

    const srcSet = new Set();
    (htmlStrings || []).forEach((html) => {
        if (!html || typeof html !== 'string' || !html.includes('<img')) return;
        const parsed = new DOMParser().parseFromString(html, 'text/html');
        parsed.querySelectorAll('img').forEach((img) => {
            const src = img.getAttribute('src');
            if (src) srcSet.add(src);
        });
    });

    const srcList = Array.from(srcSet);
    await Promise.all(srcList.map(async (src, index) => {
        const img = await loadImageElement(src);
        const width = img ? (img.naturalWidth || img.width) : 0;
        const height = img ? (img.naturalHeight || img.height) : 0;
        if (img && width && height) {
            const formatMatch = src.match(/^data:image\/(\w+)/i);
            const format = (formatMatch ? formatMatch[1] : 'png').toUpperCase().replace('JPG', 'JPEG');
            // Short, stable alias so jsPDF embeds each image once even when it
            // appears in both the template body and the example body.
            map.set(src, { img, dataUri: src, format, width, height, alias: `receipt-img-${index}` });
        }
    }));

    return map;
}

/**
 * Renders an HTML string with basic formatting into a jsPDF document,

 * with automatic page breaks (no truncation).
 * @returns {number} The final Y position after rendering
 */
function renderHtmlInPdf(doc, html, options) {
    let { startX, startY, maxWidth, margin, pageHeight, imageMap } = options;
    let currentY = startY;
    const lineHeight = 12;
    const paragraphSpacing = 18;
    const emojiSize = Math.round(lineHeight * 0.9);

    const tempDiv = document.createElement('div');
    tempDiv.style.display = 'none';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    const checkPageBreak = (neededHeight = lineHeight) => {
        if (currentY + neededHeight > pageHeight - margin) {
            doc.addPage();
            currentY = margin;
        }
    };

    // Draw an <img> as its own block, scaled to fit the content width (and page
    // height). Advances currentY past the image; returns the reset currentX.
    const drawImage = (node, currentX) => {
        const src = node.getAttribute('src');
        const entry = src && imageMap ? imageMap.get(src) : null;
        if (!entry || !entry.width || !entry.height) return currentX;

        if (currentX > startX) {
            currentY += lineHeight;
            currentX = startX;
        }

        const attrW = parseInt(node.getAttribute('width'), 10);
        let dispW = attrW && attrW > 0 ? attrW : entry.width;
        if (dispW > maxWidth) dispW = maxWidth;
        let dispH = entry.height * (dispW / entry.width);

        const maxImgHeight = pageHeight - margin * 2;
        if (dispH > maxImgHeight) {
            dispW = dispW * (maxImgHeight / dispH);
            dispH = maxImgHeight;
        }

        // Push a tall image onto a fresh page rather than clipping it.
        if (currentY + dispH > pageHeight - margin) {
            doc.addPage();
            currentY = margin;
        }

        try {
            doc.addImage(entry.img, entry.format, startX, currentY, dispW, dispH, entry.alias);
        } catch (err) {
            try {
                doc.addImage(entry.dataUri, entry.format, startX, currentY, dispW, dispH, entry.alias);
            } catch (err2) {
                console.warn('Failed to render image in receipt:', err2);
            }
        }

        currentY += dispH + 4;
        return startX;
    };

    // Render a plain-text run, drawing emoji as inline images and wrapping words.
    const drawTextRun = (text, currentX) => {
        const tokens = tokenizeRichText(text);
        for (const token of tokens) {
            if (token.type === 'space') {
                if (currentX > startX) currentX += doc.getTextWidth(' ');
            } else if (token.type === 'emoji') {
                if (currentX > startX && currentX + emojiSize > startX + maxWidth) {
                    currentY += lineHeight;
                    currentX = startX;
                    checkPageBreak();
                }
                const url = emojiToDataUrl(token.value, emojiSize);
                if (url) {
                    try {
                        // Alias keyed by glyph+size so a repeated emoji is embedded once.
                        doc.addImage(url, 'PNG', currentX, currentY - emojiSize * 0.85, emojiSize, emojiSize, `emoji-${token.value}-${emojiSize}`);
                    } catch (err) {
                        // Skip an emoji we can't place rather than aborting the receipt.
                    }
                }
                currentX += emojiSize;
            } else {
                const wordWidth = doc.getTextWidth(token.value);
                if (currentX > startX && currentX + wordWidth > startX + maxWidth) {
                    currentY += lineHeight;
                    currentX = startX;
                    checkPageBreak();
                }
                doc.text(token.value, currentX, currentY);
                currentX += wordWidth;
            }
        }
        return currentX;
    };

    const processNode = (node, currentX, styles) => {
        if (node.nodeType === 1 && node.tagName === 'IMG') {
            return drawImage(node, currentX);
        }

        const isBold = styles.isBold || node.tagName === 'STRONG' || node.tagName === 'B';
        const isItalic = styles.isItalic || node.tagName === 'EM' || node.tagName === 'I';
        let fontStyle = 'normal';
        if (isBold && isItalic) fontStyle = 'bolditalic';
        else if (isBold) fontStyle = 'bold';
        else if (isItalic) fontStyle = 'italic';

        if (node.nodeType === 3) {
            const textContent = (node.textContent || '').replace(/\s+/g, ' ');

            // Split text by parameter patterns {ParameterName}, keeping delimiters
            const paramRegex = /(\{[^}]+\})/g;
            const parts = textContent.split(paramRegex);

            for (let partIndex = 0; partIndex < parts.length; partIndex++) {
                const part = parts[partIndex];
                if (!part) continue;

                const isParameter = /^\{[^}]+\}$/.test(part);

                doc.setFont(undefined, fontStyle);
                if (isParameter) {
                    doc.setTextColor(234, 88, 12); // Orange color for parameters

                    // Render parameter as a single unit (no trailing space)
                    const partWidth = doc.getTextWidth(part);
                    if (currentX > startX && currentX + partWidth > startX + maxWidth) {
                        currentY += lineHeight;
                        currentX = startX;
                        checkPageBreak();
                    }
                    doc.text(part, currentX, currentY);
                    currentX += partWidth;
                } else {
                    doc.setTextColor(0); // Black for regular text
                    currentX = drawTextRun(part, currentX);
                }
            }
        } else {
            for (const child of Array.from(node.childNodes)) {
                currentX = processNode(child, currentX, { isBold, isItalic });
            }
        }
        return currentX;
    };

    Array.from(tempDiv.children).forEach(element => {
        checkPageBreak(paragraphSpacing);

        switch (element.tagName) {
            case 'P':
                processNode(element, startX, {});
                currentY += paragraphSpacing;
                break;
            case 'UL':
            case 'OL':
                Array.from(element.children).forEach((li, index) => {
                    checkPageBreak(paragraphSpacing);
                    const bullet = (element.tagName === 'OL') ? `${index + 1}. ` : '• ';
                    doc.text(bullet, startX, currentY);
                    processNode(li, startX + 15, {});
                    currentY += paragraphSpacing;
                });
                break;
            case 'IMG':
                drawImage(element, startX);
                currentY += 4;
                break;
            default:
                processNode(element, startX, {});
                currentY += paragraphSpacing;
        }
    });

    document.body.removeChild(tempDiv);
    return currentY;
}

/**
 * Estimates the height needed to render HTML content
 */
function estimateHtmlHeight(doc, html, maxWidth, imageMap) {
    const lineHeight = 12;
    const paragraphSpacing = 18;
    let estimatedHeight = 0;

    const tempDiv = document.createElement('div');
    tempDiv.style.display = 'none';
    tempDiv.innerHTML = html;
    document.body.appendChild(tempDiv);

    const estimateNodeHeight = (node, currentX) => {
        if (node.nodeType === 3) {
            let textContent = (node.textContent || '').replace(/\s+/g, ' ');
            const words = textContent.split(' ');
            for (const word of words) {
                if (!word) continue;
                const wordWithSpace = word + ' ';
                const wordWidth = doc.getTextWidth(wordWithSpace);
                if (currentX + wordWidth > maxWidth) {
                    estimatedHeight += lineHeight;
                    currentX = 0;
                }
                currentX += wordWidth;
            }
        } else if (node.nodeType === 1 && node.tagName === 'IMG') {
            const entry = imageMap ? imageMap.get(node.getAttribute('src')) : null;
            if (entry && entry.width && entry.height) {
                const dispW = Math.min(entry.width, maxWidth);
                estimatedHeight += entry.height * (dispW / entry.width) + 4;
            }
            currentX = 0;
        } else {
            for (const child of Array.from(node.childNodes)) {
                currentX = estimateNodeHeight(child, currentX);
            }
        }
        return currentX;
    };

    Array.from(tempDiv.children).forEach(element => {
        estimateNodeHeight(element, 0);
        estimatedHeight += paragraphSpacing;
    });

    document.body.removeChild(tempDiv);
    return estimatedHeight + 20; // Add some padding
}

/**
 * Generates a PDF receipt from the email payload using jsPDF and jsPDF-AutoTable.
 * @param {Array} emails - Array of email objects
 * @param {string} bodyTemplate - The email body template
 * @param {Object} initiator - Object with name and email of who initiated the send
 * @param {boolean} returnBase64 - If true, returns base64 string instead of saving
 * @returns {Promise<string|undefined>} - Base64 string if returnBase64 is true, undefined otherwise
 */
export async function generatePdfReceipt(emails, bodyTemplate, initiator = {}, returnBase64 = false) {
    if (!emails || emails.length === 0) {
        console.error("Emails array is empty. Cannot generate PDF receipt.");
        return;
    }

    // Pre-load any embedded images (template + a rendered body share the same
    // <img> sources) so the synchronous renderer can draw them with real sizes.
    const imageMap = await prepareReceiptImageMap([bodyTemplate, emails[0] && emails[0].body]);

    try {
        const doc = new jsPDF({ orientation: "portrait", unit: "px", format: "letter" });

        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 30;
        const contentWidth = pageWidth - (margin * 2);
        let currentY = 0;

        // Header
        doc.setFontSize(18);
        doc.text("Email Sending Receipt", pageWidth / 2, currentY + 40, { align: "center" });
        doc.setFontSize(10);
        doc.text(`Sent on: ${new Date().toLocaleString()}`, pageWidth / 2, currentY + 55, { align: "center" });

        // Add initiator info
        if (initiator.name || initiator.email) {
            doc.text(`Initiated by: ${initiator.name || 'Unknown'}${initiator.email ? ` (${initiator.email})` : ''}`, pageWidth / 2, currentY + 68, { align: "center" });
            currentY = 88;
        } else {
            currentY = 75;
        }

        // Summary section
        doc.setFontSize(12);
        doc.text("Summary", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 15;

        doc.setFontSize(10);
        doc.text(`Total Emails Sent: ${emails.length}`, margin, currentY);
        currentY += 12;

        const senderCounts = emails.reduce((acc, email) => {
            const from = email.from || "N/A";
            acc[from] = (acc[from] || 0) + 1;
            return acc;
        }, {});

        const uniqueSenders = Object.keys(senderCounts);

        if (uniqueSenders.length === 1) {
            doc.text(`Sent From: ${uniqueSenders[0]}`, margin, currentY);
            currentY += 12;
        } else {
            doc.setFont(undefined, 'bold');
            doc.text(`Sent From (Breakdown):`, margin, currentY);
            doc.setFont(undefined, 'normal');
            currentY += 12;

            uniqueSenders.forEach(sender => {
                const count = senderCounts[sender];
                doc.text(`- ${sender}: ${count} email(s)`, margin + 10, currentY);
                currentY += 12;
            });
        }
        currentY += 20;

        // Message Body section
        doc.setFontSize(12);
        doc.text("Message Body", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 20;

        const containsParameters = /\{(\w+)\}/.test(bodyTemplate);

        doc.setFontSize(10);
        doc.setFont(undefined, 'bold');
        const beforeTitle = containsParameters ? "Template Format:" : "Email Body:";
        doc.text(beforeTitle, margin, currentY);
        doc.setFont(undefined, 'normal');
        currentY += 15;

        // Render template body (full content, no truncation)
        currentY = renderHtmlInPdf(doc, bodyTemplate, {
            startX: margin,
            startY: currentY,
            maxWidth: contentWidth,
            margin: margin,
            pageHeight: pageHeight,
            imageMap: imageMap
        });

        currentY += 10;

        // Example section (if template has parameters)
        if (containsParameters) {
            const randomStudentPayload = emails[Math.floor(Math.random() * emails.length)];

            // Estimate height needed for example section
            const estimatedExampleHeight = estimateHtmlHeight(doc, randomStudentPayload.body, contentWidth, imageMap);
            const spaceRemaining = pageHeight - margin - currentY;

            // If example won't fit on current page, start new page
            if (estimatedExampleHeight > spaceRemaining) {
                doc.addPage();
                currentY = margin;
            }

            doc.setFont(undefined, 'bold');
            doc.text("Example:", margin, currentY);
            doc.setFont(undefined, 'normal');
            currentY += 15;

            // Render example body (full content, no truncation)
            currentY = renderHtmlInPdf(doc, randomStudentPayload.body, {
                startX: margin,
                startY: currentY,
                maxWidth: contentWidth,
                margin: margin,
                pageHeight: pageHeight,
                imageMap: imageMap
            });

            currentY += 10;
        }

        // Recipient list on a new page
        doc.addPage();

        // Add header for recipients page
        doc.setFontSize(12);
        doc.text("Recipient List", margin, margin);
        doc.line(margin, margin + 2, pageWidth - margin, margin + 2);

        const tableColumn = ["#", "Recipient Email", "Subject"];
        const tableRows = emails.map((email, index) => [
            index + 1,
            email.to,
            email.subject.substring(0, 45) + (email.subject.length > 45 ? '...' : '')
        ]);

        autoTable(doc, {
            head: [tableColumn],
            body: tableRows,
            startY: margin + 15,
            theme: 'grid',
            headStyles: { fillColor: [41, 128, 185] },
            styles: { fontSize: 8, cellPadding: 2 },
            columnStyles: {
                0: { cellWidth: 'auto' },
                1: { cellWidth: 150 },
                2: { cellWidth: 'auto' }
            }
        });

        if (returnBase64) {
            // Return as base64 string (without the data:application/pdf;base64, prefix)
            return doc.output('datauristring').split(',')[1];
        } else {
            const fileName = `Email_Receipt_${getTodaysLdaSheetName().replace("LDA ", "")}.pdf`;
            doc.save(fileName);
        }

    } catch (error) {
        console.error("Failed to generate PDF receipt:", error);
        return undefined;
    }
}
