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
    } catch {
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

function escapeHtml(value) {
    return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// Wait for every <img> in a container to finish decoding so html2canvas
// captures them (data-URIs decode quickly but still asynchronously).
function waitForImages(container) {
    const imgs = Array.from(container.querySelectorAll('img'));
    return Promise.all(imgs.map((img) => {
        if (img.complete && img.naturalWidth) return Promise.resolve();
        return new Promise((resolve) => {
            img.onload = () => resolve();
            img.onerror = () => resolve();
        });
    }));
}

/**
 * Render one email (subject/from/to + body) as an email-client-style card and
 * rasterize it with html2canvas, so the receipt mirrors how the message looks
 * in a real email/web view (true fonts, colors, layout, images, color emoji).
 * Returns { dataUrl, width, height } or null if capture isn't possible.
 * @param {{subject?: string, from?: string, to?: string, cc?: string, body?: string}} email
 * @param {number} cssWidth - layout width of the preview card in CSS px
 * @returns {Promise<{dataUrl: string, width: number, height: number}|null>}
 */
async function captureEmailPreview(email, cssWidth) {
    if (typeof document === 'undefined') return null;

    // Loaded on demand so html2canvas (~200KB) stays out of the main bundle.
    let html2canvas;
    try {
        html2canvas = (await import('html2canvas')).default;
    } catch (err) {
        console.warn('html2canvas unavailable; falling back to text rendering:', err);
        return null;
    }
    if (typeof html2canvas !== 'function') return null;

    const container = document.createElement('div');
    container.style.position = 'absolute';
    container.style.left = '-10000px';
    container.style.top = '0';
    container.style.width = `${cssWidth}px`;
    container.style.background = '#ffffff';
    container.style.boxSizing = 'border-box';
    container.style.pointerEvents = 'none';
    container.style.fontFamily = 'Arial, Helvetica, sans-serif';

    const ccRow = email.cc
        ? `<div style="font-size:12px; color:#6b7280; margin-top:2px;"><span style="color:#374151; font-weight:600;">Cc:</span> ${escapeHtml(email.cc)}</div>`
        : '';

    container.innerHTML = `
        <div style="border:1px solid #e5e7eb; border-radius:8px; overflow:hidden; background:#ffffff;">
            <div style="background:#f9fafb; padding:14px 18px; border-bottom:1px solid #e5e7eb;">
                <div style="font-size:18px; font-weight:600; color:#111827; line-height:1.3;">${escapeHtml(email.subject) || '(no subject)'}</div>
                <div style="font-size:12px; color:#6b7280; margin-top:8px;"><span style="color:#374151; font-weight:600;">From:</span> ${escapeHtml(email.from)}</div>
                <div style="font-size:12px; color:#6b7280; margin-top:2px;"><span style="color:#374151; font-weight:600;">To:</span> ${escapeHtml(email.to)}</div>
                ${ccRow}
            </div>
            <div style="padding:18px; font-size:14px; line-height:1.5; color:#111827; word-wrap:break-word; overflow-wrap:break-word;">
                ${email.body || ''}
            </div>
        </div>
    `;

    // Constrain any embedded images to the card width.
    container.querySelectorAll('img').forEach((img) => {
        img.style.maxWidth = '100%';
        img.style.height = 'auto';
    });

    document.body.appendChild(container);
    try {
        await waitForImages(container);
        const canvas = await html2canvas(container, {
            backgroundColor: '#ffffff',
            scale: 2,
            useCORS: true,
            logging: false,
            windowWidth: cssWidth,
        });
        return { dataUrl: canvas.toDataURL('image/jpeg', 0.92), width: canvas.width, height: canvas.height };
    } catch (err) {
        console.warn('Email preview capture failed:', err);
        return null;
    } finally {
        document.body.removeChild(container);
    }
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
        } catch {
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
                    } catch {
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
 * Generates a PDF receipt from the email payload using jsPDF and jsPDF-AutoTable.
 * The message preview is rendered from a representative entry in `emails`, so it
 * mirrors how the sent message looks in an email/web view.
 * @param {Array} emails - Array of email objects ({ from, to, cc, subject, body })
 * @param {string} [bodyTemplate] - Retained for backward compatibility; no longer used
 * @param {Object} initiator - Object with name and email of who initiated the send
 * @param {boolean} returnBase64 - If true, returns base64 string instead of saving
 * @returns {Promise<string|undefined>} - Base64 string if returnBase64 is true, undefined otherwise
 */
export async function generatePdfReceipt(emails, bodyTemplate, initiator = {}, returnBase64 = false) {
    if (!emails || emails.length === 0) {
        console.error("Emails array is empty. Cannot generate PDF receipt.");
        return;
    }

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
        const sentOn = new Date().toLocaleString(undefined, {
            year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: '2-digit'
        });
        doc.text(`Sent on: ${sentOn}`, pageWidth / 2, currentY + 55, { align: "center" });

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

        // Message Preview section — a representative rendered email shown
        // web-view style (mirrors how it looks in an email client) and scaled
        // to fit the rest of this page.
        doc.setFontSize(12);
        doc.text("Message Preview", margin, currentY);
        doc.line(margin, currentY + 2, pageWidth - margin, currentY + 2);
        currentY += 14;

        const previewEmail = emails[Math.floor(Math.random() * emails.length)];
        const availableHeight = (pageHeight - margin) - currentY;
        const preview = await captureEmailPreview(previewEmail, 640);

        if (preview && preview.width && preview.height) {
            let drawW = contentWidth;
            let drawH = contentWidth * (preview.height / preview.width);
            // Shrink-to-fit so the whole message stays on this one page.
            if (drawH > availableHeight) {
                drawW = drawW * (availableHeight / drawH);
                drawH = availableHeight;
            }
            const drawX = margin + (contentWidth - drawW) / 2; // center if narrowed
            try {
                doc.addImage(preview.dataUrl, 'JPEG', drawX, currentY, drawW, drawH);
            } catch (err) {
                console.warn('Failed to place email preview:', err);
            }
        } else {
            // Fallback: text-based rendering (still includes images + emoji).
            const imageMap = await prepareReceiptImageMap([previewEmail.body]);
            doc.setFontSize(10);
            renderHtmlInPdf(doc, previewEmail.body, {
                startX: margin,
                startY: currentY,
                maxWidth: contentWidth,
                margin: margin,
                pageHeight: pageHeight,
                imageMap: imageMap
            });
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
