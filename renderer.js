const els = {
  sidebar: document.getElementById('sidebar'),
  collapseBtn: document.getElementById('collapseBtn'),
  sidebarShowBtn: document.getElementById('sidebarShowBtn'),
  loadExcelBtn: document.getElementById('loadExcelBtn'),
  excelPath: document.getElementById('excelPath'),
  sheetSelect: document.getElementById('sheetSelect'),
  columnPool: document.getElementById('columnPool'),
  gevolgSheet: document.getElementById('gevolgSheet'),
  applyBtn: document.getElementById('applyBtn'),
  autoAdvance: document.getElementById('autoAdvance'),
  templateBtn: document.getElementById('templateBtn'),
  templatePath: document.getElementById('templatePath'),
  outputBtn: document.getElementById('outputBtn'),
  outputPath: document.getElementById('outputPath'),
  currentInfo: document.getElementById('currentInfo'),
  addressBar: document.getElementById('addressBar'),
  backBtn: document.getElementById('backBtn'),
  forwardBtn: document.getElementById('forwardBtn'),
  reloadBtn: document.getElementById('reloadBtn'),
  externalBtn: document.getElementById('externalBtn'),
  loadSpinner: document.getElementById('loadSpinner'),
  commentField: document.getElementById('commentField'),
  undoBtn: document.getElementById('undoBtn'),
  redoBtn: document.getElementById('redoBtn'),
  generateBtn: document.getElementById('generateBtn'),
  deleteReportBtn: document.getElementById('deleteReportBtn'),
  falsePositiveBtn: document.getElementById('falsePositiveBtn'),
  notSureBtn: document.getElementById('notSureBtn'),
  noAccessBtn: document.getElementById('noAccessBtn'),
  prevBtn: document.getElementById('prevBtn'),
  skipBtn: document.getElementById('skipBtn'),
  jobBar: document.getElementById('jobBar'),
  jobBarHead: document.getElementById('jobBarHead'),
  jobSummary: document.getElementById('jobSummary'),
  jobToggle: document.getElementById('jobToggle'),
  jobList: document.getElementById('jobList'),
  log: document.getElementById('log'),
  logToggle: document.getElementById('logToggle'),
  bvContainer: document.getElementById('bvContainer'),
};

let lastLoadedComment = '';
let headerNames = [];

// slot id -> column name currently docked there
const slotValues = {
  urlCol: '',
  statusCol: '',
  commentCol: '',
  gevolgCol: '',
  refCol: '',
  inbreukCol: '',
};

function appendLog(msg) {
  const line = document.createElement('div');
  const ts = document.createElement('span');
  ts.className = 'ts';
  ts.textContent = `${new Date().toLocaleTimeString()}  `;
  line.appendChild(ts);
  line.appendChild(document.createTextNode(msg));
  els.log.appendChild(line);
  els.log.scrollTop = els.log.scrollHeight;
}

window.api.onLog((msg) => appendLog(msg));
window.api.onNav((navState) => {
  els.addressBar.value = navState.url;
});

// A crawl-heavy site can take a while; show that something is happening.
window.api.onLoadState((s) => {
  els.loadSpinner.hidden = !s.loading;
  els.addressBar.classList.toggle('loading', !!s.loading);
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

// ---- sidebar / panels -------------------------------------------------

function setSidebarCollapsed(collapsed) {
  els.sidebar.classList.toggle('collapsed', collapsed);
  els.sidebarShowBtn.hidden = !collapsed;
  // The BrowserView is positioned in screen coordinates, so it has to be told
  // about the new content box once the CSS transition has finished.
  setTimeout(updateBounds, 200);
}
els.collapseBtn.addEventListener('click', () => setSidebarCollapsed(true));
els.sidebarShowBtn.addEventListener('click', () => setSidebarCollapsed(false));

els.logToggle.addEventListener('click', () => {
  const hidden = els.log.classList.toggle('hidden');
  els.logToggle.innerHTML = hidden ? '&#9652;' : '&#9662;';
  setTimeout(updateBounds, 50);
});

els.jobToggle.addEventListener('click', () => {
  const hidden = els.jobList.classList.toggle('hidden');
  els.jobToggle.innerHTML = hidden ? '&#9652;' : '&#9662;';
  setTimeout(updateBounds, 50);
});

// ---- background capture jobs ------------------------------------------

const JOB_LABELS = {
  queued: 'Waiting',
  running: 'Capturing',
  done: 'Ready to review',
  failed: 'Failed',
  cancelled: 'Cancelled',
};

function renderJobs(jobs) {
  els.jobBar.hidden = jobs.length === 0;
  els.jobList.innerHTML = '';

  const running = jobs.filter((j) => j.status === 'running').length;
  const queued = jobs.filter((j) => j.status === 'queued').length;
  const done = jobs.filter((j) => j.status === 'done').length;
  const parts = [];
  if (running) parts.push(`${running} running`);
  if (queued) parts.push(`${queued} waiting`);
  if (done) parts.push(`${done} ready`);
  els.jobSummary.textContent = parts.join(' · ');

  for (const job of jobs) {
    const row = document.createElement('div');
    row.className = 'job ' + job.status;

    const dot = document.createElement('span');
    dot.className = 'jobDot';

    const text = document.createElement('div');
    text.className = 'jobText';
    const url = document.createElement('div');
    url.className = 'jobUrl';
    url.textContent = `Row ${job.row + 1} — ${job.url}`;
    url.title = job.url;
    const msg = document.createElement('div');
    msg.className = 'jobMsg';
    msg.textContent = `${JOB_LABELS[job.status] || job.status}${job.message ? ' · ' + job.message : ''}`;
    text.appendChild(url);
    text.appendChild(msg);

    row.appendChild(dot);
    row.appendChild(text);

    if (job.status === 'done') {
      const review = document.createElement('button');
      review.className = 'primary';
      review.textContent = 'Review';
      review.addEventListener('click', async () => {
        try {
          await window.api.reviewJob(job.id);
        } catch (err) {
          appendLog('Error: ' + err.message);
          alert(err.message);
        }
      });
      row.appendChild(review);
    }

    if (job.status === 'queued' || job.status === 'running') {
      const cancel = document.createElement('button');
      cancel.className = 'ghost';
      cancel.textContent = 'Cancel';
      cancel.addEventListener('click', () => window.api.cancelJob(job.id));
      row.appendChild(cancel);
    } else {
      const dismiss = document.createElement('button');
      dismiss.className = 'ghost iconBtn';
      dismiss.innerHTML = '&times;';
      dismiss.title = 'Dismiss';
      dismiss.addEventListener('click', () => window.api.dismissJob(job.id));
      row.appendChild(dismiss);
    }

    els.jobList.appendChild(row);
  }
  setTimeout(updateBounds, 30);
}

window.api.onCaptureJobs((jobs) => renderJobs(jobs));

// ---- column docking --------------------------------------------------

function slotEl(slot) {
  return document.querySelector(`.slot[data-slot="${slot}"]`);
}

function refreshSlot(slot) {
  const el = slotEl(slot);
  const value = slotValues[slot];
  const valueEl = el.querySelector('.slotValue');
  const required = el.dataset.required === '1';
  if (value) {
    valueEl.textContent = value;
    valueEl.classList.remove('empty');
  } else {
    valueEl.textContent = required ? 'drop here' : 'optional';
    valueEl.classList.add('empty');
  }
  el.classList.toggle('filled', !!value);
  el.classList.toggle('unfilled', required && !value);
}

function refreshChips() {
  const used = new Set(Object.values(slotValues).filter(Boolean));
  els.columnPool.querySelectorAll('.chip').forEach((chip) => {
    chip.classList.toggle('used', used.has(chip.dataset.name));
  });
}

function refreshApplyState() {
  els.applyBtn.disabled = !(slotValues.urlCol && slotValues.statusCol && headerNames.length > 0);
}

function setSlot(slot, value) {
  slotValues[slot] = value;
  refreshSlot(slot);
  refreshChips();
  refreshApplyState();
}

function buildPool(names) {
  headerNames = names.filter((n) => n && String(n).trim()).map((n) => String(n).trim());
  els.columnPool.innerHTML = '';
  if (headerNames.length === 0) {
    const span = document.createElement('span');
    span.className = 'poolEmpty';
    span.textContent = 'This sheet has no header row.';
    els.columnPool.appendChild(span);
    return;
  }
  headerNames.forEach((name) => {
    const chip = document.createElement('div');
    chip.className = 'chip';
    chip.textContent = name;
    chip.title = name;
    chip.draggable = true;
    chip.dataset.name = name;
    chip.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', name);
      e.dataTransfer.effectAllowed = 'copy';
      chip.classList.add('dragging');
    });
    chip.addEventListener('dragend', () => chip.classList.remove('dragging'));
    els.columnPool.appendChild(chip);
  });
  refreshChips();
  refreshApplyState();
}

document.querySelectorAll('.slot').forEach((el) => {
  const slot = el.dataset.slot;
  el.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
    el.classList.add('over');
  });
  el.addEventListener('dragleave', () => el.classList.remove('over'));
  el.addEventListener('drop', (e) => {
    e.preventDefault();
    el.classList.remove('over');
    const name = e.dataTransfer.getData('text/plain');
    if (name) setSlot(slot, name);
  });
  el.querySelector('.clearSlot').addEventListener('click', () => setSlot(slot, ''));
  refreshSlot(slot);
});

// ---- entry state -----------------------------------------------------

function setEntryActionButtonsEnabled(enabled) {
  [
    els.generateBtn,
    els.falsePositiveBtn,
    els.notSureBtn,
    els.noAccessBtn,
    els.skipBtn,
  ].forEach((b) => (b.disabled = !enabled));
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
    return;
  }

  els.currentInfo.textContent = `Row ${result.row + 1}`;
  els.addressBar.value = result.url;
  setEntryActionButtonsEnabled(true);

  els.commentField.disabled = !result.commentEnabled;
  els.commentField.value = result.comment || '';
  lastLoadedComment = result.comment || '';

  els.generateBtn.textContent = result.hasReport ? 'Regenerate Report' : 'Generate Report';
  els.deleteReportBtn.disabled = !result.hasReport;

  updateBounds();
}

setEntryActionButtonsEnabled(false);

// ---- sheet pickers ----------------------------------------------------

function fillSheetSelects(sheetNames, selected) {
  const wanted = selected || els.gevolgSheet.value;
  els.sheetSelect.innerHTML = '';
  els.gevolgSheet.innerHTML = '';
  sheetNames.forEach((name) => {
    const a = document.createElement('option');
    a.value = name;
    a.textContent = name;
    els.sheetSelect.appendChild(a);
    const b = document.createElement('option');
    b.value = name;
    b.textContent = name;
    els.gevolgSheet.appendChild(b);
  });
  // Offer the default even when the workbook doesn't have it yet; the main
  // process creates it on demand.
  if (!sheetNames.includes('Gevolg opties')) {
    const c = document.createElement('option');
    c.value = 'Gevolg opties';
    c.textContent = 'Gevolg opties (will be created)';
    els.gevolgSheet.appendChild(c);
  }
  if (wanted && [...els.gevolgSheet.options].some((o) => o.value === wanted)) {
    els.gevolgSheet.value = wanted;
  } else if (sheetNames.includes('Gevolg opties')) {
    els.gevolgSheet.value = 'Gevolg opties';
  }
  els.sheetSelect.disabled = false;
  els.gevolgSheet.disabled = false;
}

// ---- startup ----------------------------------------------------------

(async () => {
  const saved = (await window.api.loadSettings()) || {};
  ['urlCol', 'statusCol', 'commentCol', 'gevolgCol', 'refCol', 'inbreukCol'].forEach((slot) => {
    if (saved[slot]) setSlot(slot, saved[slot]);
  });
  if (saved.templatePath) els.templatePath.textContent = saved.templatePath;
  if (saved.outputFolder) els.outputPath.textContent = saved.outputFolder;
  els.autoAdvance.checked = saved.autoAdvance !== false;

  // Reopen last session's workbook, sheet and columns.
  try {
    const restored = await window.api.restoreSession();
    if (restored) {
      els.excelPath.textContent = restored.filePath;
      fillSheetSelects(restored.sheetNames, saved.gevolgSheet);
      els.sheetSelect.value = restored.sheetName;
      buildPool(restored.header || []);
      if (restored.entry) setEntry(restored.entry);
    } else if (saved.gevolgSheet) {
      const opt = document.createElement('option');
      opt.value = saved.gevolgSheet;
      opt.textContent = saved.gevolgSheet;
      els.gevolgSheet.appendChild(opt);
    }
  } catch (err) {
    appendLog('Could not restore the last session: ' + err.message);
  }

  renderJobs(await window.api.captureJobs());
})();

els.autoAdvance.addEventListener('change', () => {
  window.api.saveSettings({ autoAdvance: els.autoAdvance.checked });
});

// ---- actions ---------------------------------------------------------

els.loadExcelBtn.addEventListener('click', async () => {
  const result = await window.api.openExcel();
  if (!result) return;
  els.excelPath.textContent = result.filePath;
  fillSheetSelects(result.sheetNames);
  appendLog('Loaded: ' + result.filePath);
  const info = await window.api.selectSheet(els.sheetSelect.value);
  buildPool(info.header || []);
});

els.sheetSelect.addEventListener('change', async () => {
  const info = await window.api.selectSheet(els.sheetSelect.value);
  buildPool(info.header || []);
});

els.applyBtn.addEventListener('click', async () => {
  try {
    const result = await window.api.setColumns({
      urlCol: slotValues.urlCol,
      statusCol: slotValues.statusCol,
      commentCol: slotValues.commentCol,
      gevolgCol: slotValues.gevolgCol,
      refCol: slotValues.refCol,
      inbreukCol: slotValues.inbreukCol,
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
els.externalBtn.addEventListener('click', async () => {
  try {
    await window.api.openExternal();
  } catch (err) {
    appendLog('Error: ' + err.message);
  }
});

els.generateBtn.addEventListener('click', async () => {
  await flushComment();
  try {
    const job = await window.api.queueCapture();
    appendLog(`Capture queued for row ${job.row + 1}. You can keep working.`);
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

// A status set from the error page's buttons still has to move the UI on.
window.api.onEntryUpdated((info) => setEntry(info));

window.api.onReportCreated((info) => {
  appendLog(`Report ${info.reference} created: ${info.outPath}`);
});

els.deleteReportBtn.addEventListener('click', async () => {
  if (!confirm('Delete the report for this entry? This cannot be undone.')) return;
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

async function markAndAdvance(status) {
  await flushComment();
  try {
    setEntry(await window.api.markStatus(status));
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
}
els.falsePositiveBtn.addEventListener('click', () => markAndAdvance('False Positive'));
els.notSureBtn.addEventListener('click', () => markAndAdvance('Not Sure'));
els.noAccessBtn.addEventListener('click', () => markAndAdvance('No access'));

els.skipBtn.addEventListener('click', async () => {
  await flushComment();
  try {
    setEntry(await window.api.skip());
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.prevBtn.addEventListener('click', async () => {
  await flushComment();
  try {
    setEntry(await window.api.previous());
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.undoBtn.addEventListener('click', async () => {
  try {
    setEntry(await window.api.undo());
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});

els.redoBtn.addEventListener('click', async () => {
  try {
    setEntry(await window.api.redo());
  } catch (err) {
    appendLog('Error: ' + err.message);
    alert(err.message);
  }
});
