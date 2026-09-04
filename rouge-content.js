/* ROUGE Site Content layer — applies "Site Content" sheet overrides to any
   element carrying a data-ck="key" attribute. If a key has no row in the
   sheet, the element's existing hard-coded HTML is left untouched, so the
   site always looks correct even before you've edited anything. */

let ROUGE_CONTENT = {};
let ROUGE_CONTENT_READY = false;

function onRougeContentReady(callback) {
  if (ROUGE_CONTENT_READY) callback();
  else window.addEventListener('rougeContentReady', callback, { once: true });
}

function applyRougeContent() {
  document.querySelectorAll('[data-ck]').forEach(el => {
    const key = el.getAttribute('data-ck');
    const entry = ROUGE_CONTENT[key];
    if (!entry || !entry.value) return; // nothing saved yet -> keep default markup

    const type = (entry.type || el.getAttribute('data-ck-type') || 'text').toLowerCase();

    if (type === 'image') {
      if (el.tagName === 'IMG') {
        el.src = normalizeImageUrl(entry.value);
      } else {
        el.style.backgroundImage = `url("${normalizeImageUrl(entry.value).replace(/"/g, '\\"')}")`;
      }
    } else if (type === 'html') {
      el.innerHTML = entry.value;
    } else {
      // Plain text, but we still allow the author to use literal \n for line breaks
      // on elements that were already using multi-line titles (white-space:pre-line).
      el.textContent = entry.value.replace(/\\n/g, '\n');
    }
  });
}

function loadRougeContent() {
  return new Promise((resolve) => {
    const cb = '__rougeContent_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    const cleanup = () => { try { delete window[cb]; } catch (e) { window[cb] = undefined; } script.remove(); };
    const timeout = setTimeout(() => { cleanup(); finish({}); }, 12000);

    function finish(map) {
      ROUGE_CONTENT = map || {};
      ROUGE_CONTENT_READY = true;
      applyRougeContent();
      window.dispatchEvent(new Event('rougeContentReady'));
      resolve(ROUGE_CONTENT);
    }

    window[cb] = (data) => {
      clearTimeout(timeout);
      cleanup();
      if (data && data.success) finish(data.content || {});
      else finish({});
    };

    script.onerror = () => { clearTimeout(timeout); cleanup(); finish({}); };
    script.src = `${ROUGE_API_URL}?action=siteContent&callback=${encodeURIComponent(cb)}&_=${Date.now()}`;
    document.head.appendChild(script);
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', loadRougeContent);
} else {
  loadRougeContent();
}
