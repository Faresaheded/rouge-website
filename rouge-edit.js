/* ROUGE visual editor — activate on any page with ?edit=1
   Click any highlighted block on the actual site to edit its text or image
   in place. Saves go straight to the "Site Content" tab in Google Sheets
   through the same Apps Script Web App used for products. */
(function () {
  const params = new URLSearchParams(location.search);
  if (params.get('edit') !== '1') return;

  const PW_KEY = 'rougeEditPassword';
  let password = (sessionStorage.getItem(PW_KEY) || '').trim();

  function pageKeyFromPath() {
    const file = (location.pathname.split('/').pop() || 'index.html').replace('.html', '');
    return file === '' ? 'index' : file;
  }

  function toast(msg, isError) {
    let t = document.getElementById('rougeEditToast');
    if (!t) {
      t = document.createElement('div');
      t.id = 'rougeEditToast';
      Object.assign(t.style, {
        position: 'fixed', left: '50%', bottom: '90px', transform: 'translateX(-50%)',
        background: '#1a1a1a', color: '#fff', padding: '10px 18px', borderRadius: '4px',
        font: '500 12px/1.4 Manrope,sans-serif', letterSpacing: '.04em', zIndex: 999999,
        boxShadow: '0 8px 24px rgba(0,0,0,.4)', transition: 'opacity .25s', opacity: '0'
      });
      document.body.appendChild(t);
    }
    t.style.background = isError ? '#7a1a1a' : '#1a1a1a';
    t.textContent = msg;
    t.style.opacity = '1';
    clearTimeout(t._timer);
    t._timer = setTimeout(() => { t.style.opacity = '0'; }, 2200);
  }

  async function saveKey(key, type, value) {
    if (!password) {
      password = (prompt('Enter the ROUGE edit password:') || '').trim();
      sessionStorage.setItem(PW_KEY, password);
    }
    return new Promise((resolve) => {
      const cb = '__rougeEditSave_' + Date.now() + '_' + Math.random().toString(36).slice(2);
      const script = document.createElement('script');
      let settled = false;

      const finish = (data) => {
        if (settled) return;
        settled = true;
        try { delete window[cb]; } catch (e) { window[cb] = undefined; }
        script.remove();
        if (!data.success) {
          if (data.error === 'Invalid password') {
            sessionStorage.removeItem(PW_KEY);
            password = '';
          }
          toast('Not saved — ' + (data.error || 'unknown error'), true);
          resolve(false);
          return;
        }
        ROUGE_CONTENT[key] = { value, type };
        toast('Saved ✓');
        resolve(true);
      };

      const timeout = setTimeout(() => {
        finish({ success:false, error:'request timed out' });
      }, 25000);

      window[cb] = (data) => { clearTimeout(timeout); finish(data); };
      script.onerror = () => { clearTimeout(timeout); finish({ success:false, error:'could not reach the ROUGE API' }); };

      const params = new URLSearchParams({
        action: 'updateContent', key, type, value,
        page: pageKeyFromPath(), password, callback: cb, _: Date.now()
      });
      script.src = `${ROUGE_API_URL}?${params.toString()}`;
      document.head.appendChild(script);
    });
  }

  function injectStyles() {
    const s = document.createElement('style');
    s.textContent = `
      [data-ck]{outline:1px dashed transparent;outline-offset:3px;cursor:pointer;transition:outline-color .15s}
      [data-ck]:hover{outline-color:#ff3b3b}
      .rouge-edit-badge{position:fixed;left:20px;bottom:20px;z-index:999999;background:#c02032;color:#fff;
        font:600 11px/1 Manrope,sans-serif;letter-spacing:.12em;padding:10px 16px;border-radius:3px;
        box-shadow:0 8px 24px rgba(0,0,0,.35)}
      .rouge-edit-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:999998}
      .rouge-edit-panel{position:fixed;z-index:999999;top:50%;left:50%;transform:translate(-50%,-50%);
        background:#fff;color:#1a1a1a;border-radius:6px;box-shadow:0 12px 40px rgba(0,0,0,.35);
        padding:18px;width:min(420px,90vw);max-height:80vh;overflow-y:auto;
        font:400 13px/1.5 Manrope,sans-serif}
      .rouge-edit-panel textarea, .rouge-edit-panel input{width:100%;box-sizing:border-box;padding:10px;
        border:1px solid #ddd;border-radius:4px;font:inherit;margin-top:8px;resize:vertical}
      .rouge-edit-panel .rouge-edit-row{display:flex;gap:8px;margin-top:14px;position:sticky;bottom:0;background:#fff;padding-top:6px}
      .rouge-edit-panel button{flex:1;padding:10px;border:none;border-radius:4px;font:600 11px/1 Manrope,sans-serif;
        letter-spacing:.06em;cursor:pointer}
      .rouge-edit-save{background:#c02032;color:#fff}
      .rouge-edit-cancel{background:#eee;color:#333}
      .rouge-edit-preview{width:100%;max-height:220px;object-fit:cover;border-radius:4px;margin-top:8px;background:#f2f2f2}
      .rouge-edit-label{font:600 10px/1 Manrope,sans-serif;letter-spacing:.1em;color:#888;text-transform:uppercase}
    `;
    document.head.appendChild(s);
  }

  function closePanel() {
    document.getElementById('rougeEditPanel')?.remove();
    document.getElementById('rougeEditBackdrop')?.remove();
  }

  function openBackdrop() {
    const b = document.createElement('div');
    b.id = 'rougeEditBackdrop';
    b.className = 'rouge-edit-backdrop';
    b.onclick = closePanel;
    document.body.appendChild(b);
  }

  function openTextPanel(el, key) {
    closePanel();
    openBackdrop();
    const panel = document.createElement('div');
    panel.id = 'rougeEditPanel';
    panel.className = 'rouge-edit-panel';
    const current = (ROUGE_CONTENT[key] && ROUGE_CONTENT[key].value) || el.textContent.trim();
    panel.innerHTML = `
      <div class="rouge-edit-label">EDITING: ${key}</div>
      <textarea rows="6">${current.replace(/</g, '&lt;')}</textarea>
      <div class="rouge-edit-row">
        <button class="rouge-edit-save">SAVE</button>
        <button class="rouge-edit-cancel">CANCEL</button>
      </div>`;
    document.body.appendChild(panel);
    const textarea = panel.querySelector('textarea');
    textarea.focus();
    panel.querySelector('.rouge-edit-cancel').onclick = closePanel;
    panel.querySelector('.rouge-edit-save').onclick = async () => {
      const ok = await saveKey(key, 'text', textarea.value.trim());
      if (ok) { el.textContent = textarea.value.trim(); closePanel(); }
    };
  }

  function openImagePanel(el, key) {
    closePanel();
    openBackdrop();
    const panel = document.createElement('div');
    panel.id = 'rougeEditPanel';
    panel.className = 'rouge-edit-panel';
    const current = (ROUGE_CONTENT[key] && ROUGE_CONTENT[key].value) || '';
    panel.innerHTML = `
      <div class="rouge-edit-label">EDITING IMAGE: ${key}</div>
      <input type="text" placeholder="Google Drive share link or image URL" value="${current.replace(/"/g, '&quot;')}">
      <img class="rouge-edit-preview" src="${normalizeImageUrl(current)}">
      <div class="rouge-edit-row">
        <button class="rouge-edit-save">SAVE</button>
        <button class="rouge-edit-cancel">CANCEL</button>
      </div>`;
    document.body.appendChild(panel);
    const input = panel.querySelector('input');
    const preview = panel.querySelector('.rouge-edit-preview');
    input.focus();
    input.addEventListener('input', () => { preview.src = normalizeImageUrl(input.value.trim()); });
    panel.querySelector('.rouge-edit-cancel').onclick = closePanel;
    panel.querySelector('.rouge-edit-save').onclick = async () => {
      const url = input.value.trim();
      const ok = await saveKey(key, 'image', url);
      if (ok) {
        if (el.tagName === 'IMG') el.src = normalizeImageUrl(url);
        else el.style.backgroundImage = `url("${normalizeImageUrl(url).replace(/"/g, '\\"')}")`;
        closePanel();
      }
    };
  }

  function isImageTarget(el) {
    const t = (el.getAttribute('data-ck-type') || '').toLowerCase();
    if (t) return t === 'image';
    return el.tagName === 'IMG' || getComputedStyle(el).backgroundImage !== 'none';
  }

  function start() {
    injectStyles();
    const badge = document.createElement('div');
    badge.className = 'rouge-edit-badge';
    badge.textContent = 'ROUGE EDIT MODE — click any dashed block';
    document.body.appendChild(badge);

    document.addEventListener('click', (e) => {
      const el = e.target.closest('[data-ck]');
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      const key = el.getAttribute('data-ck');
      if (isImageTarget(el)) openImagePanel(el, key);
      else openTextPanel(el, key);
    }, true);

    document.addEventListener('click', (e) => {
      if (e.target.id === 'rougeEditBackdrop') return; // backdrop has its own handler
      if (!e.target.closest('.rouge-edit-panel') && !e.target.closest('[data-ck]')) closePanel();
    });
  }

  if (window.ROUGE_CONTENT_READY) start();
  else onRougeContentReady(start);
})();
