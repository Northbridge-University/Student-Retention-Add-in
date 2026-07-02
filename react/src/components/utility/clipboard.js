// Utility for copy-to-clipboard with fallback mechanism
export const copyToClipboard = async (text) => {
  if (!text) return;

  // 1. Try Modern Async API
  if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch (err) {
      console.warn('Clipboard API failed, attempting fallback...', err);
    }
  }

  // 2. Fallback: document.execCommand('copy')
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.style.top = '0';
    ta.setAttribute('readonly', '');

    document.body.appendChild(ta);
    ta.focus();
    ta.select();

    const successful = document.execCommand('copy');
    document.body.removeChild(ta);

    if (!successful) {
      console.error('Fallback copy failed.');
    }
  } catch (err) {
    console.error('All copy methods failed', err);
  }
};
