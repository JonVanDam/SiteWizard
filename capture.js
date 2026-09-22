const els = {
  site: document.getElementById('capSite'),
  ref: document.getElementById('refField'),
  refBadge: document.getElementById('refBadge'),
  spinner: document.getElementById('spinner'),
  progressText: document.getElementById('progressText'),
  shotCount: document.getElementById('shotCount'),
  shotList: document.getElementById('shotList'),
  shotEmpty: document.getElementById('shotEmpty'),
  selectAllBtn: document.getElementById('selectAllBtn'),
  selectNoneBtn: document.getElementById('selectNoneBtn'),
  recaptureBtn: document.getElementById('recaptureBtn'),
  agentSelect: document.getElementById('agentSelect'),
  datum: document.getElementById('datumField'),
  bron: document.getElementById('bronField'),
  gevolgPanel: document.getElementById('gevolgPanel'),
  gevolgSummary: document.getElementById('gevolgSummary'),
  cancelBtn: document.getElementById('cancelBtn'),
  generateBtn: document.getElementById('generateBtn'),
  viewer: document.getElementById('viewer'),
  viewerImage: document.getElementById('viewerImage'),
  viewerCaption: document.getElementById('viewerCaption'),
  viewerClose: document.getElementById('viewerClose'),
};

let ctx = null;
let shots = [];
const selected = new Set();
let gevolgOptions = [];
let gevolgSelection = [];

// ---- gevolg ----------------------------------------------------------

function summariseGevolg() {
  if (gevolgSelection.length === 0) {
    els.gevolgSummary.textContent = 'No measures selected.';
    return;
  }
  els.gevolgSummary.textContent = gevolgSelection
    .map((s) => (s.text ? `${s.label} ${s.text}`.trim() : s.label))
    .join(' | ');
}

function readGevolg() {
  const out = [];
  els.gevolgPanel.querySelectorAll('.gevolgRow').forEach((row) => {
    const box = row.querySelector('input[type="checkbox"]');
    if (!box || !box.checked) return;
    const text = row.querySelector('.freeText');
    out.push({ label: box.dataset.label, text: text ? text.value.trim() : '' });
  });
  return out;
}

function onGevolgChanged() {
  gevolgSelection = readGevolg();
  els.gevolgPanel.querySelectorAll('.gevolgRow').forEach((row) => {
    const box = row.querySelector('input[type="checkbox"]');
    const text = row.querySelector('.freeText');
    if (text && box) text.disabled = !box.checked;
  });
  summariseGevolg();
}

function buildGevolg() {
  els.gevolgPanel.innerHTML = '';
  const chosen = new Map(gevolgSelection.map((s) => [s.label, s.text || '']));
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
    box.id = 'gv' + i;
    box.dataset.label = opt.label;
    box.checked = chosen.has(opt.label);
    box.addEventListener('change', onGevolgChanged);

    const span = document.createElement('span');
    span.textContent = opt.label;
    label.appendChild(box);
    label.appendChild(span);
    row.appendChild(label);

    if (opt.needsText) {
      const text = document.createElement('input');
      text.type = 'text';
      text.className = 'freeText';
      text.placeholder = 'Klik of tik om tekst in te voeren.';
      text.value = chosen.get(opt.label) || '';
      text.disabled = !box.checked;
      text.addEventListener('change', onGevolgChanged);
      row.appendChild(text);
    }
    els.gevolgPanel.appendChild(row);
  });
  summariseGevolg();
}

// ---- screenshots -----------------------------------------------------

function updateCounts() {
  els.shotCount.textContent = shots.length
    ? `${selected.size} of ${shots.length} kept`
    : '';
  els.generateBtn.disabled = selected.size === 0;
  els.selectAllBtn.disabled = shots.length === 0;
  els.selectNoneBtn.disabled = shots.length === 0;
}

function renderShots() {
  els.shotList.innerHTML = '';
  els.shotEmpty.hidden = shots.length > 0;

  for (const shot of shots) {
    const row = document.createElement('div');
    row.className = 'shot' + (selected.has(shot.id) ? ' selected' : '');

    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = selected.has(shot.id);
    box.addEventListener('change', () => {
      if (box.checked) selected.add(shot.id);
      else selected.delete(shot.id);
      row.classList.toggle('selected', box.checked);
      updateCounts();
    });

    const wrap = document.createElement('div');
    wrap.className = 'thumbWrap';
    const img = document.createElement('img');
    img.src = shot.thumb;
    img.alt = shot.pageUrl;
    wrap.appendChild(img);
    wrap.addEventListener('click', () => openViewer(shot));
    wrap.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      openViewer(shot);
    });

    const meta = document.createElement('div');
    meta.className = 'meta';
    const url = document.createElement('div');
    url.className = 'url';
    url.textContent = shot.pageUrl;
    const sub = document.createElement('div');
    sub.className = 'sub';
    sub.innerHTML =
      `<span>Part ${shot.slice} of ${shot.sliceCount}</span>` +
      `<span>${shot.width}×${shot.height}</span>` +
      '<span>Right-click to enlarge</span>';
    meta.appendChild(url);
    meta.appendChild(sub);

    row.appendChild(box);
    row.appendChild(wrap);
    row.appendChild(meta);
    els.shotList.appendChild(row);
  }
  updateCounts();
}

async function openViewer(shot) {
  els.viewerCaption.textContent = `${shot.pageUrl} — part ${shot.slice} of ${shot.sliceCount}`;
  els.viewerImage.src = shot.thumb;
  els.viewer.hidden = false;
  const full = await window.api.captureImage(shot.id);
  if (full) els.viewerImage.src = full;
}

function closeViewer() {
  els.viewer.hidden = true;
  els.viewerImage.src = '';
}
els.viewerClose.addEventListener('click', closeViewer);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !els.viewer.hidden) closeViewer();
});

// ---- capture run -----------------------------------------------------

window.api.onCaptureProgress((p) => {
  if (p.phase === 'loading') {
    els.progressText.textContent = p.total
      ? `Loading page ${p.index} of ${p.total}: ${p.page}`
      : `Loading ${p.page}`;
  } else if (p.phase === 'queued') {
    els.progressText.textContent = `Found ${p.total} page(s) to capture`;
  } else if (p.phase === 'capturing') {
    els.progressText.textContent = `Captured ${p.shots} screenshot(s)`;
  } else if (p.phase === 'warning') {
    els.progressText.textContent = p.message;
  } else if (p.phase === 'done') {
    els.progressText.textContent = `Finished — ${p.shots} screenshot(s)`;
    els.spinner.classList.add('done');
  }
});

async function runCapture() {
  els.spinner.classList.remove('done');
  els.recaptureBtn.disabled = true;
  els.progressText.textContent = 'Starting…';
  shots = [];
  selected.clear();
  renderShots();
  try {
    shots = await window.api.captureRun();
    // Everything is kept by default; untick what you don't want.
    shots.forEach((s) => selected.add(s.id));
  } catch (err) {
    els.progressText.textContent = 'Capture failed: ' + err.message;
  }
  renderShots();
  els.recaptureBtn.disabled = false;
}

// ---- wiring ----------------------------------------------------------

els.selectAllBtn.addEventListener('click', () => {
  shots.forEach((s) => selected.add(s.id));
  renderShots();
});
els.selectNoneBtn.addEventListener('click', () => {
  selected.clear();
  renderShots();
});
els.recaptureBtn.addEventListener('click', runCapture);
els.cancelBtn.addEventListener('click', () => window.api.captureCancel());

els.generateBtn.addEventListener('click', async () => {
  els.generateBtn.disabled = true;
  els.generateBtn.textContent = 'Generating…';
  try {
    await window.api.captureGenerate({
      selectedIds: shots.filter((s) => selected.has(s.id)).map((s) => s.id),
      reference: els.ref.value.trim(),
      agentName: els.agentSelect.value,
      datum: els.datum.value.trim(),
      bron: els.bron.value.trim(),
      gevolg: gevolgSelection,
    });
  } catch (err) {
    alert(err.message);
    els.generateBtn.disabled = false;
    els.generateBtn.textContent = 'Generate Report';
  }
});

(async () => {
  ctx = await window.api.captureContext();
  els.site.textContent = ctx.rawUrl && ctx.rawUrl !== ctx.url
    ? `${ctx.url}   (cell: ${ctx.rawUrl})`
    : ctx.url;
  els.ref.value = ctx.reference;
  els.refBadge.textContent = ctx.referenceIsNew ? 'new' : 'existing';
  els.refBadge.className = 'badge' + (ctx.referenceIsNew ? ' new' : '');

  ctx.agentNames.forEach((n) => {
    const opt = document.createElement('option');
    opt.value = n;
    opt.textContent = n;
    els.agentSelect.appendChild(opt);
  });
  if (ctx.agentName && ctx.agentNames.includes(ctx.agentName)) {
    els.agentSelect.value = ctx.agentName;
  }

  els.datum.value = new Date().toLocaleDateString('nl-BE');
  els.bron.value = ctx.bron || '';

  gevolgOptions = ctx.gevolgOptions || [];
  gevolgSelection = ctx.gevolg || [];
  buildGevolg();

  runCapture();
})();
