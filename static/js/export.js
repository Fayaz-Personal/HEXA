// static/js/export.js
/**
 * HEXA Enterprise Document Generation & Export System - Frontend Controller
 */

const API_BASE = '/api/export';

let _activeExportContent = '';
let _activeExportTitle = '';
let _activeSourceType = 'current_response';

function _el(id) {
  return document.getElementById(id);
}

function _showStatus(msg, isError = false, isSuccess = false) {
  const box = _el('export-status-box');
  if (!box) return;
  box.style.display = 'block';
  box.className = `export-status-banner ${isError ? 'status-error' : isSuccess ? 'status-success' : 'status-info'}`;
  box.innerHTML = msg;
}

function _hideStatus() {
  const box = _el('export-status-box');
  if (box) box.style.display = 'none';
}

function _updateFormatOptionsUI() {
  const format = document.querySelector('input[name="export-format"]:checked')?.value || 'pdf';

  const panels = ['pdf', 'docx', 'pptx', 'xlsx'];
  panels.forEach(p => {
    const el = _el(`export-opts-${p}`);
    if (el) el.style.display = (format === p || format === 'all') ? 'block' : 'none';
  });
}

export function openExportModal(options = {}) {
  const modal = _el('export-modal');
  if (!modal) return;

  _activeExportContent = options.content || window.chatModule?.getLastAIMessageText?.() || '';
  _activeExportTitle = options.title || _extractTitleFromContent(_activeExportContent) || 'HEXA AI Generated Report';
  _activeSourceType = options.sourceType || 'current_response';

  const titleInput = _el('export-title');
  if (titleInput) titleInput.value = _activeExportTitle;

  const authorInput = _el('export-author');
  if (authorInput && !authorInput.value) authorInput.value = 'HEXA AI';

  const sourceSelect = _el('export-source-type');
  if (sourceSelect) sourceSelect.value = _activeSourceType;

  const textInput = _el('export-content-text');
  if (textInput) textInput.value = _activeExportContent;

  _hideStatus();
  _updateFormatOptionsUI();

  modal.classList.remove('hidden');
}

export function closeExportModal() {
  const modal = _el('export-modal');
  if (modal) modal.classList.add('hidden');
}

function _extractTitleFromContent(content) {
  if (!content) return '';
  const firstLine = content.split('\n').find(l => l.trim().length > 0) || '';
  const clean = firstLine.replace(/^[#\*\-\s]+/, '').strip ? firstLine.replace(/^[#\*\-\s]+/, '').strip() : firstLine.replace(/^[#\*\-\s]+/, '');
  return clean.slice(0, 60);
}

export async function handleGenerateExport() {
  const btn = _el('export-submit-btn');
  const format = document.querySelector('input[name="export-format"]:checked')?.value || 'pdf';

  const title = (_el('export-title')?.value || '').trim() || 'HEXA AI Generated Report';
  const author = (_el('export-author')?.value || '').trim() || 'HEXA AI';
  const sourceType = _el('export-source-type')?.value || 'current_response';
  const content = (_el('export-content-text')?.value || _activeExportContent).trim();

  if (!content) {
    _showStatus('<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg> Please enter or select content to export.', true);
    return;
  }

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin"><line x1="12" y1="2" x2="12" y2="6"/><line x1="12" y1="18" x2="12" y2="22"/><line x1="4.93" y1="4.93" x2="7.76" y2="7.76"/><line x1="16.24" y1="16.24" x2="19.07" y2="19.07"/><line x1="2" y1="12" x2="6" y2="12"/><line x1="18" y1="12" x2="22" y2="12"/></svg> Generating…`;
  }

  _showStatus('Preparing content…<br><span style="opacity:0.7;font-size:11px;">Parsing markdown sections and tables…</span>');

  const payload = {
    title,
    author,
    content,
    source_type: sourceType,
    pdf_options: {
      include_header: _el('export-pdf-header')?.checked ?? true,
      include_datetime: _el('export-pdf-datetime')?.checked ?? true,
      include_page_numbers: _el('export-pdf-pagenums')?.checked ?? true,
      include_sources: _el('export-pdf-sources')?.checked ?? true,
    },
    word_options: {
      include_title_page: _el('export-docx-titlepage')?.checked ?? true,
      include_toc: _el('export-docx-toc')?.checked ?? true,
      include_headers_footers: _el('export-docx-hf')?.checked ?? true,
      include_sources: _el('export-docx-sources')?.checked ?? true,
    },
    pptx_options: {
      max_slides: parseInt(_el('export-pptx-maxslides')?.value || '20', 10),
      include_title_slide: _el('export-pptx-titleslide')?.checked ?? true,
      include_summary_slide: _el('export-pptx-summaryslide')?.checked ?? true,
      include_source_slide: _el('export-pptx-sourceslide')?.checked ?? true,
    },
    excel_options: {
      auto_worksheets: _el('export-xlsx-worksheets')?.checked ?? true,
      include_source_info: _el('export-xlsx-sources')?.checked ?? true,
      auto_format_tables: _el('export-xlsx-formattables')?.checked ?? true,
      freeze_header: _el('export-xlsx-freeze')?.checked ?? true,
      auto_size_columns: _el('export-xlsx-autosize')?.checked ?? true,
    },
  };

  try {
    _showStatus('Generating document…<br><span style="opacity:0.7;font-size:11px;">Formatting tables and layout…</span>');

    const res = await fetch(`${API_BASE}/${format}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'same-origin',
      body: JSON.stringify(payload),
    });

    const data = await res.json();

    if (!res.ok || data.error) {
      _showStatus(`Export failed: ${data.error || 'Unknown error'}`, true);
      return;
    }

    _showStatus('Finalizing file…', false, false);

    if (format === 'all' && data.files) {
      let downloadHTML = `<div style="font-weight:600;margin-bottom:6px;">Export Complete — ZIP Package Generated!</div>`;
      downloadHTML += `<a href="${data.download_url}" download="${data.filename}" class="export-dl-btn export-dl-main">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Download All (ZIP Archive - ${(data.file_size / 1024).toFixed(1)} KB)
      </a>`;
      downloadHTML += `<div style="margin-top:10px;font-size:11px;opacity:0.8;display:grid;grid-template-columns:1fr 1fr;gap:6px;">`;
      data.files.forEach(f => {
        downloadHTML += `<a href="${f.download_url}" download="${f.filename}" class="export-dl-sub">⬇ ${f.format.toUpperCase()}: ${f.filename}</a>`;
      });
      downloadHTML += `</div>`;
      _showStatus(downloadHTML, false, true);
    } else {
      const downloadHTML = `<div style="font-weight:600;margin-bottom:6px;">Export Complete!</div>
      <a href="${data.download_url}" download="${data.filename}" class="export-dl-btn export-dl-main">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Download ${data.filename} (${(data.file_size / 1024).toFixed(1)} KB)
      </a>`;
      _showStatus(downloadHTML, false, true);
    }
  } catch (err) {
    _showStatus(`Export failed: ${err.message}. Your original content is unchanged.`, true);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg> Generate Document`;
    }
  }
}

export function initExportSystem() {
  const modalClose = _el('export-modal-close');
  if (modalClose) modalClose.addEventListener('click', closeExportModal);

  const modalBackdrop = _el('export-modal');
  if (modalBackdrop) {
    modalBackdrop.addEventListener('click', e => {
      if (e.target === modalBackdrop) closeExportModal();
    });
  }

  const formatRadios = document.querySelectorAll('input[name="export-format"]');
  formatRadios.forEach(r => r.addEventListener('change', _updateFormatOptionsUI));

  const submitBtn = _el('export-submit-btn');
  if (submitBtn) submitBtn.addEventListener('click', handleGenerateExport);

  const cancelBtn = _el('export-cancel-btn');
  if (cancelBtn) cancelBtn.addEventListener('click', closeExportModal);
}

document.addEventListener('DOMContentLoaded', () => {
  initExportSystem();
});

window.hexaExport = {
  open: openExportModal,
  close: closeExportModal,
  exportContent: openExportModal,
};

export default window.hexaExport;
