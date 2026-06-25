// Image compression helpers for the email body editor.
//
// Pasted/dropped images land in the Quill body as base64 data-URIs. The body is
// duplicated into every recipient's payload entry, so a single large image gets
// copied N times and can push JSON.stringify past V8's max string length
// (~512MB) — the whole send then fails with "RangeError: Invalid string
// length". To keep payloads small, images larger than
// IMAGE_COMPRESS_THRESHOLD_BYTES are downscaled and re-encoded as JPEG before
// they're embedded.

// Images larger than this (decoded byte size) are compressed on insert.
export const IMAGE_COMPRESS_THRESHOLD_BYTES = 500 * 1024; // 500 KB

/**
 * Approximate decoded byte length of a base64 data-URI (or bare base64 string).
 * @param {string} dataUri
 * @returns {number}
 */
export function dataUriByteLength(dataUri) {
    if (!dataUri || typeof dataUri !== 'string') return 0;
    const commaIndex = dataUri.indexOf(',');
    const base64 = commaIndex >= 0 ? dataUri.slice(commaIndex + 1) : dataUri;
    const length = base64.length;
    if (length === 0) return 0;
    let padding = 0;
    if (base64.endsWith('==')) padding = 2;
    else if (base64.endsWith('=')) padding = 1;
    return Math.max(0, Math.floor((length * 3) / 4) - padding);
}

/**
 * Human-readable byte size, e.g. "4.2 MB".
 * @param {number} bytes
 * @returns {string}
 */
export function formatBytes(bytes) {
    if (!bytes || bytes <= 0) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / Math.pow(1024, exponent);
    const rounded = exponent === 0 ? Math.round(value) : Math.round(value * 10) / 10;
    return `${rounded} ${units[exponent]}`;
}

/**
 * Compress any oversized data-URI images inside an HTML string and return the
 * updated HTML. Used for content authored outside the main editor (e.g. saved
 * signatures), which would otherwise embed full-size images that repeat across
 * every recipient and bloat the payload. Non-image and remote sources are left
 * untouched. Falls back to the original HTML if parsing/compression isn't
 * possible (e.g. outside a browser).
 * @param {string} html
 * @returns {Promise<string>}
 */
export async function compressImagesInHtml(html) {
    if (!html || typeof html !== 'string' || !html.includes('<img')) return html;
    if (typeof DOMParser === 'undefined') return html;

    const doc = new DOMParser().parseFromString(html, 'text/html');
    const imgs = Array.from(doc.querySelectorAll('img'));
    let changed = false;

    await Promise.all(imgs.map(async (img) => {
        const src = img.getAttribute('src') || '';
        if (src.startsWith('data:image') && dataUriByteLength(src) > IMAGE_COMPRESS_THRESHOLD_BYTES) {
            try {
                const compressed = await compressDataUriImage(src);
                if (compressed) {
                    img.setAttribute('src', compressed);
                    changed = true;
                }
            } catch {
                // Keep the original image if compression fails.
            }
        }
    }));

    return changed ? doc.body.innerHTML : html;
}

/**
 * Downscale + re-encode a base64 image data-URI so it's small enough to embed in
 * the email payload. Resolves to a new JPEG data-URI, or rejects if the image
 * can't be loaded/encoded (e.g. a tainted or unsupported source).
 *
 * Uses the browser canvas, so it only runs inside the Office task pane.
 * @param {string} dataUri
 * @param {{ maxDimension?: number, quality?: number }} [options]
 * @returns {Promise<string>}
 */
export function compressDataUriImage(dataUri, { maxDimension = 1280, quality = 0.72 } = {}) {
    return new Promise((resolve, reject) => {
        if (typeof document === 'undefined' || typeof Image === 'undefined') {
            reject(new Error('Image compression is not available in this environment.'));
            return;
        }

        const img = new Image();
        img.onload = () => {
            try {
                let width = img.naturalWidth || img.width;
                let height = img.naturalHeight || img.height;
                if (!width || !height) {
                    reject(new Error('Image has no intrinsic size.'));
                    return;
                }

                // Scale the longest edge down to maxDimension, preserving aspect ratio.
                if (width > maxDimension || height > maxDimension) {
                    if (width >= height) {
                        height = Math.round((height * maxDimension) / width);
                        width = maxDimension;
                    } else {
                        width = Math.round((width * maxDimension) / height);
                        height = maxDimension;
                    }
                }

                const canvas = document.createElement('canvas');
                canvas.width = width;
                canvas.height = height;
                const ctx = canvas.getContext('2d');
                // JPEG has no alpha channel; flatten transparency onto white so
                // transparent PNGs don't render with a black background.
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, width, height);
                ctx.drawImage(img, 0, 0, width, height);

                const result = canvas.toDataURL('image/jpeg', quality);
                if (!result || !result.startsWith('data:image/jpeg')) {
                    reject(new Error('Failed to encode compressed image.'));
                    return;
                }
                resolve(result);
            } catch (error) {
                reject(error);
            }
        };
        img.onerror = () => reject(new Error('Failed to load image for compression.'));
        img.src = dataUri;
    });
}
