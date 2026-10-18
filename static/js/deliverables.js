// static/js/deliverables.js
// Deliverables modal — generate Word, Excel, PowerPoint, PDF directly
// without relying on the AI agent or tool-calling.

const API = '/api/deliverable';

// ── helpers ──────────────────────────────────────────────────────────────────

function _el(id) { return document.getElementById(id); }

function _showResult(elId, msg, isError = false) {
  const el = _el(elId);
  if (!el) return;
  el.style.color = isError ? 'var(--color-error, #e74c3c)' : 'var(--brand-color, #0a9396)';
  el.innerHTML = msg;
}

function _downloadLink(url, filename) {
  return `<a href="${url}" download="${filename}"
    style="display:inline-flex;align-items:center;gap:6px;padding:7px 14px;
    border-radius:6px;background:var(--brand-color,#0a9396);color:#fff;
    font-size:12px;font-weight:600;text-decoration:none;margin-top:6px;">
    ⬇ Download ${filename}
  </a>`;
}

async function _generate(endpoint, body, resultElId) {
  const btn = _el(resultElId.replace('-result', '-generate-btn'));
  if (btn) { btn.disabled = true; btn.textContent = 'Generating…'; }
  _showResult(resultElId, 'Building your file…');

  try {
    const res = await fetch(`${API}/${endpoint}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'same-origin',
      body: JSON.stringify(body),
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      _showResult(resultElId, `❌ ${data.error || 'Failed'}`, true);
      return;
    }

    _showResult(resultElId,
      `✅ ${data.message}<br>${_downloadLink(data.download_url, data.filename)}`
    );
  } catch (e) {
    _showResult(resultElId, `❌ Network error: ${e.message}`, true);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = btn.dataset.label || 'Generate'; }
  }
}

// ── Word (.docx) ─────────────────────────────────────────────────────────────

function _addDocxSection(container) {
  const idx = container.querySelectorAll('.docx-section').length;
  const div = document.createElement('div');
  div.className = 'docx-section';
  div.style.cssText = 'border:1px solid color-mix(in srgb,var(--fg) 12%,transparent);border-radius:8px;padding:10px;margin-bottom:8px;';
  div.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
      <span style="font-size:11px;font-weight:600;opacity:0.6;">Section ${idx + 1}</span>
      <button type="button" class="docx-remove-section" style="background:none;border:none;cursor:pointer;opacity:0.4;font-size:14px;color:var(--fg);">×</button>
    </div>
    <input type="text" placeholder="Section heading (optional)" class="docx-sec-title"
      style="width:100%;padding:6px 8px;border-radius:5px;border:1px solid color-mix(in srgb,var(--fg) 15%,transparent);background:color-mix(in srgb,var(--fg) 5%,transparent);color:var(--fg);font-size:12px;box-sizing:border-box;margin-bottom:6px;">
    <textarea rows="4" placeholder="Section content / paragraphs…" class="docx-sec-body"
      style="width:100%;padding:6px 8px;border-radius:5px;border:1px solid color-mix(in srgb,var(--fg) 15%,transparent);background:color-mix(in srgb,var(--fg) 5%,transparent);color:var(--fg);font-size:12px;box-sizing:border-box;resize:vertical;font-family:inherit;"></textarea>
  `;
  div.querySelector('.docx-remove-section').addEventListener('click', () => div.remove());
  container.appendChild(div);
}

function _initDocx() {
  const list = _el('docx-sections-list');
  if (!list) return;

  // Start with 3 default sections
  _addDocxSection(list);
  _addDocxSection(list);
  _addDocxSection(list);

  _el('docx-add-section')?.addEventListener('click', () => _addDocxSection(list));

  const btn = _el('docx-generate-btn');
  if (btn) {
    btn.dataset.label = 'Generate Word Document';
    btn.addEventListener('click', async () => {
      const title = (_el('docx-title')?.value || '').trim() || 'Untitled Document';
      const sections = [];
      list.querySelectorAll('.docx-section').forEach(sec => {
        sections.push({
          title: sec.querySelector('.docx-sec-title')?.value?.trim() || '',
          body: sec.querySelector('.docx-sec-body')?.value?.trim() || '',
        });
      });
      if (sections.every(s => !s.title && !s.body)) {
        _showResult('docx-result', '❌ Add at least one section with content.', true);
        return;
      }
      await _generate('docx', { title, sections }, 'docx-result');
    });
  }
}

// ── Excel (.xlsx) ─────────────────────────────────────────────────────────────

function _initXlsx() {
  const btn = _el('xlsx-generate-btn');
  if (!btn) return;
  btn.dataset.label = 'Generate Excel Spreadsheet';
  btn.addEventListener('click', async () => {
    const title = (_el('xlsx-title')?.value || '').trim() || 'Spreadsheet';
    const sheetName = (_el('xlsx-sheet')?.value || '').trim() || 'Sheet1';
    const headersRaw = (_el('xlsx-headers')?.value || '').trim();
    const rowsRaw = (_el('xlsx-rows')?.value || '').trim();

    if (!headersRaw) {
      _showResult('xlsx-result', '❌ Enter at least the headers row.', true);
      return;
    }

    const headers = headersRaw.split(',').map(h => h.trim());
    const rows = [headers];
    if (rowsRaw) {
      rowsRaw.split('\n').forEach(line => {
        const cells = line.split(',').map(c => c.trim());
        if (cells.some(c => c)) rows.push(cells);
      });
    }

    await _generate('xlsx', { title, sheets: { [sheetName]: rows } }, 'xlsx-result');
  });
}

// ── PowerPoint (.pptx) ────────────────────────────────────────────────────────

function _addPptxSlide(container) {
  const idx = container.querySelectorAll('.pptx-slide').length;
  const div = document.createElement('div');
  div.className = 'pptx-slide';
  div.style.cssText = 'border:1px solid color-mix(in srgb,var(--fg) 12%,transparent);border-radius:8px;padding:10px;margin-bottom:8px;';
  div.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
      <span style="font-size:11px;font-weight:600;opacity:0.6;">Slide ${idx + 1}</span>
      <button type="button" class="pptx-remove-slide" style="background:none;border:none;cursor:pointer;opacity:0.4;font-size:14px;color:var(--fg);">×</button>
    </div>
    <input type="text" placeholder="Slide title" class="pptx-slide-title"
      style="width:100%;padding:6px 8px;border-radius:5px;border:1px solid color-mix(in srgb,var(--fg) 15%,transparent);background:color-mix(in srgb,var(--fg) 5%,transparent);color:var(--fg);font-size:12px;box-sizing:border-box;margin-bottom:6px;">
    <textarea rows="4" placeholder="Bullet points — one per line" class="pptx-slide-points"
      style="width:100%;padding:6px 8px;border-radius:5px;border:1px solid color-mix(in srgb,var(--fg) 15%,transparent);background:color-mix(in srgb,var(--fg) 5%,transparent);color:var(--fg);font-size:12px;box-sizing:border-box;resize:vertical;font-family:inherit;"></textarea>
  `;
  div.querySelector('.pptx-remove-slide').addEventListener('click', () => div.remove());
  container.appendChild(div);
}

function _initPptx() {
  const list = _el('pptx-slides-list');
  if (!list) return;

  _addPptxSlide(list);
  _addPptxSlide(list);
  _addPptxSlide(list);

  _el('pptx-add-slide')?.addEventListener('click', () => _addPptxSlide(list));

  const btn = _el('pptx-generate-btn');
  if (btn) {
    btn.dataset.label = 'Generate PowerPoint';
    btn.addEventListener('click', async () => {
      const title = (_el('pptx-title')?.value || '').trim() || 'Presentation';
      const slides = [];
      list.querySelectorAll('.pptx-slide').forEach(s => {
        const slideTitle = s.querySelector('.pptx-slide-title')?.value?.trim() || '';
        const points = (s.querySelector('.pptx-slide-points')?.value || '')
          .split('\n').map(l => l.trim()).filter(Boolean);
        if (slideTitle || points.length) slides.push({ title: slideTitle, points });
      });
      if (!slides.length) {
        _showResult('pptx-result', '❌ Add at least one slide.', true);
        return;
      }
      await _generate('pptx', { title, slides }, 'pptx-result');
    });
  }
}

// ── PDF ───────────────────────────────────────────────────────────────────────

function _initPdf() {
  const btn = _el('pdf-generate-btn');
  if (!btn) return;
  btn.dataset.label = 'Generate PDF';
  btn.addEventListener('click', async () => {
    const title = (_el('pdf-title')?.value || '').trim() || 'Document';
    const content = (_el('pdf-content')?.value || '').trim();
    if (!content) {
      _showResult('pdf-result', '❌ Enter some content.', true);
      return;
    }
    await _generate('pdf', { title, content }, 'pdf-result');
  });
}

// ── Files list ─────────────────────────────────────────────────────────────────

async function _loadFileList() {
  const container = _el('deliverables-file-list');
  if (!container) return;
  container.innerHTML = '<div style="opacity:0.5;font-size:12px;">Loading…</div>';
  try {
    const res = await fetch(`${API}/list`, { credentials: 'same-origin' });
    const data = await res.json();
    if (!data.files || data.files.length === 0) {
      container.innerHTML = '<div style="opacity:0.5;font-size:12px;">No files yet. Generate one from the other tabs.</div>';
      return;
    }
    const extIcons = { docx: '📄', xlsx: '📊', pptx: '📑', pdf: '📋' };
    container.innerHTML = data.files.map(f => `
      <div style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-radius:8px;background:color-mix(in srgb,var(--fg) 5%,transparent);">
        <span style="font-size:18px;">${extIcons[f.type] || '📁'}</span>
        <div style="flex:1;min-width:0;">
          <div style="font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${f.filename}</div>
          <div style="font-size:11px;opacity:0.5;">${(f.size/1024).toFixed(1)} KB · ${f.type.toUpperCase()}</div>
        </div>
        <a href="${f.download_url}" download="${f.filename}"
          style="padding:5px 12px;border-radius:6px;background:var(--brand-color,#0a9396);color:#fff;font-size:11px;font-weight:600;text-decoration:none;white-space:nowrap;">
          ⬇ Download
        </a>
      </div>
    `).join('');
  } catch (e) {
    container.innerHTML = `<div style="opacity:0.5;font-size:12px;">Error loading files: ${e.message}</div>`;
  }
}

// ── Tab switching ──────────────────────────────────────────────────────────────

function _initTabs() {
  const tabs = document.querySelectorAll('.deliv-tab');
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      // Style all tabs
      tabs.forEach(t => {
        t.style.background = 'none';
        t.style.opacity = '0.6';
      });
      tab.style.background = 'color-mix(in srgb,var(--fg) 10%,transparent)';
      tab.style.opacity = '1';

      // Show/hide panels
      document.querySelectorAll('.deliv-tab-panel').forEach(p => p.style.display = 'none');
      const panel = _el(`deliv-tab-${tab.dataset.tab}`);
      if (panel) panel.style.display = '';

      if (tab.dataset.tab === 'files') _loadFileList();
    });
  });
}

// ── Modal open/close ───────────────────────────────────────────────────────────

let _initialized = false;

export function openDeliverables() {
  const modal = _el('deliverables-modal');
  if (!modal) return;
  modal.classList.remove('hidden');

  if (!_initialized) {
    _initialized = true;
    _initTabs();
    _initDocx();
    _initXlsx();
    _initPptx();
    _initPdf();

    _el('deliverables-modal-close')?.addEventListener('click', closeDeliverables);
    modal.addEventListener('click', e => { if (e.target === modal) closeDeliverables(); });
    _el('deliverables-refresh-btn')?.addEventListener('click', _loadFileList);
  }
}

export function closeDeliverables() {
  _el('deliverables-modal')?.classList.add('hidden');
}

export default { open: openDeliverables, close: closeDeliverables };
