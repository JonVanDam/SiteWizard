const els = {
  loadExcelBtn: document.getElementById('loadExcelBtn'),
  sheetSelect: document.getElementById('sheetSelect'),
  urlCol: document.getElementById('urlCol'),
  statusCol: document.getElementById('statusCol'),
  commentCol: document.getElementById('commentCol'),
  gevolgCol: document.getElementById('gevolgCol'),
  gevolgSheet: document.getElementById('gevolgSheet'),
  gevolgToggle: document.getElementById('gevolgToggle'),
  gevolgCaret: document.getElementById('gevolgCaret'),
  gevolgSummary: document.getElementById('gevolgSummary'),
  gevolgClearBtn: document.getElementById('gevolgClearBtn'),
  gevolgPanel: document.getElementById('gevolgPanel'),
  applyBtn: document.getElementById('applyBtn'),
  templateBtn: document.getElementById('templateBtn'),
  templatePath: document.getElementById('templatePath'),
  outputBtn: document.getElementById('outputBtn'),
  outputPath: document.getElementById('outputPath'),
  currentInfo: document.getElementById('currentInfo'),
  addressBar: document.getElementById('addressBar'),
  backBtn: document.getElementById('backBtn'),
  forwardBtn: document.getElementById('forwardBtn'),
  reloadBtn: document.getElementById('reloadBtn'),
  commentField: document.getElementById('commentField'),
  undoBtn: document.getElementById('undoBtn'),
  redoBtn: document.getElementById('redoBtn'),
  generateBtn: document.getElementById('generateBtn'),
  deleteReportBtn: document.getElementById('deleteReportBtn'),
  falsePositiveBtn: document.getElementById('falsePositiveBtn'),
  notSureBtn: document.getElementById('notSureBtn'),
  prevBtn: document.getElementById('prevBtn'),
  skipBtn: document.getElementById('skipBtn'),
  log: document.getElementById('log'),
  bvContainer: document.getElementById('bvContainer'),
};

let lastLoadedComment = '';

function appendLog(msg) {
  const line = document.createElement('div');
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  els.log.appendChild(line);
  els.log.scrollTop = els.log.scrollHeight;
}

window.api.onLog((msg) => appendLog(msg));
window.api.onNav((navState) => {
  els.addressBar.value = navState.url;
});

function updateBounds() {
  const rect = els.bvContainer.getBoundingClientRect();
  // DOMRect isn't structured-cloneable across the contextBridge, so pull out
  // plain numbers or the IPC message silently fails to arrive.
  window.api.setBrowserViewBounds({
    x: rect.x,
    y: rect.y,
    width: rect.width,
    height: rect.height,
  });
}
window.addEventListener('resize', updateBounds);
window.addEventListener('DOMContentLoaded', () => setTimeout(updateBounds, 50));

function setEntryActionButtonsEnabled(enabled) {
  [els.generateBtn, els.falsePositiveBtn, els.notSureBtn, els.skipBtn].forEach((b) => (b.disabled = !enabled));
}

// Saves the comment field if its content differs from what was last loaded
// from/saved to the sheet. Called on blur and defensively before any action
// that moves to a different entry, so nothing typed is lost.
async function flushComment() {
  if (els.commentField.disabled) return;
  const value = els.commentField.value;
  if (value === lastLoadedComment) return;
  await window.api.saveComment(value);
  lastLoadedComment = value;
}
els.commentField.addEventListener('blur', flushComment);

// ---- Gevolg (follow-up measures) ------------------------------------
// Options come from a sheet in the workbook, so the panel is rebuilt from
// whatever the main process hands back rather than from a fixed list here.

let gevolgOptions = [];
let gevolgSelection = []; // [{label, text}]
let gevolgEnabled = false;
// Whether the user changed the panel on *this* entry. A selection that was
// only carried over from the previous entry is a suggestion, so Skip leaves
// the file untouched unless it was actually edited here.
let gevolgDirty = false;

function summariseGevolg() {
  if (!gevolgEnabled) {
    els.gevolgSummary.textContent = '(select a gevolg column to enable)';
    els.gevolgSummary.className = 'empty';
    return;
  }
  if (gevolgSelection.length === 0) {
    els.gevolgSummary.textContent = 'No measures selected.';
    els.gevolgSummary.className = 'empty';
    return;
  }
  const labels = gevolgSelection.map((s) => (s.text ? `${s.label} ${s.text}`.trim() : s.label));
  els.gevolgSummary.textContent = `${labels.length} selected: ${labels.join(' | ')}`;
  els.gevolgSummary.className = gevolgDirty ? '' : 'prefilled';
}

function readGevolgFromPanel() {
  const rows = els.gevolgPanel.querySelectorAll('.gevolgRow');
  const out = [];
  rows.forEach((row) => {
    const box = row.querySelector('input[type="checkbox"]');
    if (!box || !box.checked) return;
    const textInput = row.querySelector('.freeText');
    out.push({ label: box.dataset.label, text: textInput ? textInput.value.trim() : '' });
  });
  return out;
}

function onGevolgChanged() {
  gevolgSelection = readGevolgFromPanel();
  gevolgDirty = true;
  // Free-text boxes only make sense while their measure is ticked.
  els.gevolgPanel.querySelectorAll('.gevolgRow').forEach((row) => {
    const box = row.querySelector('input[type="checkbox"]');
    const textInput = row.querySelector('.freeText');
    if (textInput && box) textInput.disabled = !box.checked;
  });
  els.gevolgClearBtn.disabled = gevolgSelection.length === 0;
  summariseGevolg();
}

function buildGevolgPanel() {
  els.gevolgPanel.innerHTML = '';
  const chosen = new Map(gevolgSelection.map((s) => [s.label, s.text || '']));

  // Anything stored in the sheet that is no longer in the options list is
  // still shown (ticked) so editing the options sheet can't hide a decision.
  const known = new Set(gevolgOptions.map((o) => o.label));
  const extras = gevolgSelection
    .filter((s) => !known.has(s.label))
    .map((s) => ({ label: s.label, needsText: false }));

  [...gevolgOptions, ...extras].forEach((opt, i) => {
    const row = document.createElement('div');
    row.className = 'gevolgRow';

    const label = document.createElement('label');
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.id = `gevolgOpt${i}`;
    box.dataset.label = opt.label;
    box.checked = chosen.has(opt.label);
    box.addEventListener('change', onGevolgChanged);

    const span = document.createElement('span');
    span.textContent = opt.label;
    label.appendChild(box);
    label.appendChild(span);
    row.appendChild(label);

    if (opt.needsText) {
      const textInput = document.createElement('input');
      textInput.type = 'text';
      textInput.className = 'freeText';
      textInput.placeholder = 'Klik of tik om tekst in te voeren.';
      textInput.value = chosen.get(opt.label) || '';
      textInput.disabled = !box.checked;
      textInput.addEventListener('input', () => {
        gevolgDirty = true;
      });
      textInput.addEventListener('change', onGevolgChanged);
      row.appendChild(textInput);
    }

    els.gevolgPanel.appendChild(row);
  });

  els.gevolgClearBtn.disabled = gevolgSelection.length === 0;
  summariseGevolg();
}

function setGevolgPanelOpen(open) {
  els.gevolgPanel.hidden = !open;
  els.gevolgToggle.setAttribute('aria-expanded', String(open));
  els.gevolgCaret.innerHTML = open ? '&#9652;' : '&#9662;';
  // The panel is in normal flow, so opening it resizes #bvContainer and the
  // BrowserView has to be told about its new box.
  updateBounds();
}

// force: record the current selection even if the user didn't touch the panel
// (i.e. they accepted the carried-over measures by making a decision).
async function flushGevolg(force) {
  if (!gevolgEnabled) return;
  if (!gevolgDirty && !force) return;
  await window.api.saveGevolg(gevolgSelection);
  gevolgDirty = false;
  summariseGevolg();
}

els.gevolgToggle.addEventListener('click', () => {
  setGevolgPanelOpen(els.gevolgPanel.hidden);
});

els.gevolgClearBtn.addEventListener('click', () => {
  gevolgSelection = [];
  gevolgDirty = true;
  buildGevolgPanel();
});

function setEntry(result) {
  if (!result) return;

  els.undoBtn.disabled = !result.canUndo;
  els.redoBtn.disabled = !result.canRedo;
  els.prevBtn.disabled = !result.canPrevious;

  if (result.done) {
    els.currentInfo.textContent = 'No more entries.';
    els.addressBar.value = '';
    setEntryActionButtonsEnabled(false);
    els.commentField.value = '';
    els.commentField.disabled = true;
    els.deleteReportBtn.disabled = true;
    els.generateBtn.textContent = 'Generate Report';
    lastLoadedComment = '';
    els.gevolgToggle.disabled = true;
    els.gevolgClearBtn.disabled = true;
    return;
  }

  els.currentInfo.textContent = `Row ${result.row + 1}: ${result.url}`;
  els.addressBar.value = result.url;
  setEntryActionButtonsEnabled(true);

  els.commentField.disabled = !result.commentEnabled;
  els.commentField.value = result.comment || '';
  lastLoadedComment = result.comment || '';

  els.generateBtn.textContent = result.hasReport ? 'Add Screenshot' : 'Generate Report';
  els.deleteReportBtn.disabled = !result.hasReport;

  gevolgEnabled = !!result.gevolgEnabled;
  gevolgOptions = result.gevolgOptions || [];
  gevolgSelection = result.gevolg || [];
  // Carried-over measures count as untouched, so Skip won't write them.
  gevolgDirty = false;
  els.gevolgToggle.disabled = !gevolgEnabled;
  buildGevolgPanel();

  updateBounds();
}

setEntryActionButtonsEnabled(false);

(async () => {
  const saved = await window.api.loadSettings();
  if (!saved) return;
  if (saved.urlCol) els.urlCol.value = saved.urlCol;
  if (saved.statusCol) els.statusCol.value = saved.statusCol;
  if (saved.commentCol) els.commentCol.value = saved.commentCol;
  if (saved.gevolgCol) els.gevolgCol.value = saved.gevolgCol;
  if (saved.gevolgSheet) els.gevolgSheet.value = saved.gevolgSheet;
  if (saved.templatePath) els.templatePath.textContent = saved.templatePath;
  if (saved.outputFolder) els.outputPath.textContent = saved.outputFolder;
})();

els.loadExcelBtn.addEventListener('click', async () => {
  const result = await window.api.openExcel();
  if (!result) return;
  els.sheetSelect.innerHTML = '';
  result.sheetNames.forEach((name) => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    els.sheetSelect.appendChild(opt);
  });
  appendLog('Loaded: ' + result.filePath);
  await window.api.selectSheet(els.sheetSelect.value);
});

els.sheetSelect.addEventListener('change', async () => {
  await window.api.selectSheet(els.sheetSelect.value);
});

els.applyBtn.addEventListener('click', async () => {
  try {
    const result = await window.api.setColumns({
      urlCol: els.urlCol.value,
      statusCol: els.statusCol.value,
      commentCol: els.commentCol.value,
      gevolgCol: els.gevolgCol.value,
      gevolgSheet: els.gevolgSheet.value,
    });
    appendLog('Columns applied. Jumping to first unprocessed row.');
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.templateBtn.addEventListener('click', async () => {
  const p = await window.api.openTemplate();
  if (p) els.templatePath.textContent = p;
});

els.outputBtn.addEventListener('click', async () => {
  const p = await window.api.selectOutputFolder();
  if (p) els.outputPath.textContent = p;
});

els.backBtn.addEventListener('click', () => window.api.navBack());
els.forwardBtn.addEventListener('click', () => window.api.navForward());
els.reloadBtn.addEventListener('click', () => window.api.navReload());

els.generateBtn.addEventListener('click', async () => {
  els.generateBtn.disabled = true;
  // Record the measures before capturing so the sheet and the report agree.
  await flushGevolg(true);
  try {
    const result = await window.api.generateReport();
    const n = result.screenshotCount;
    appendLog(`Report saved with ${n} screenshot${n === 1 ? '' : 's'}: ${result.outPath}`);
    els.generateBtn.textContent = 'Add Screenshot';
    els.deleteReportBtn.disabled = false;
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  } finally {
    els.generateBtn.disabled = false;
  }
});

els.deleteReportBtn.addEventListener('click', async () => {
  const confirmed = confirm('Delete the report for this entry? This cannot be undone.');
  if (!confirmed) return;
  try {
    await window.api.deleteReport();
    els.generateBtn.textContent = 'Generate Report';
    els.deleteReportBtn.disabled = true;
    appendLog('Report deleted.');
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.falsePositiveBtn.addEventListener('click', async () => {
  await flushComment();
  await flushGevolg(true);
  try {
    const result = await window.api.markStatus('False Positive');
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.notSureBtn.addEventListener('click', async () => {
  await flushComment();
  await flushGevolg(true);
  try {
    const result = await window.api.markStatus('Not Sure');
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.skipBtn.addEventListener('click', async () => {
  await flushComment();
  await flushGevolg(false);
  try {
    const result = await window.api.skip();
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.prevBtn.addEventListener('click', async () => {
  await flushComment();
  await flushGevolg(false);
  try {
    const result = await window.api.previous();
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.undoBtn.addEventListener('click', async () => {
  try {
    const result = await window.api.undo();
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.redoBtn.addEventListener('click', async () => {
  try {
    const result = await window.api.redo();
    setEntry(result);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});
