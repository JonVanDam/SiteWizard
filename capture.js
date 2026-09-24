const els = {
  tabStrip: document.getElementById('tabStrip'),
  site: document.getElementById('capSite'),
  ref: document.getElementById('refField'),
  refBadge: document.getElementById('refBadge'),
  shotCount: document.getElementById('shotCount'),
  shotList: document.getElementById('shotList'),
  shotEmpty: document.getElementById('shotEmpty'),
  selectAllBtn: document.getElementById('selectAllBtn'),
  selectNoneBtn: document.getElementById('selectNoneBtn'),
  agentSelect: document.getElementById('agentSelect'),
  datum: document.getElementById('datumField'),
  bron: document.getElementById('bronField'),
  inbreukPanel: document.getElementById('inbreukPanel'),
  gevolgPanel: document.getElementById('gevolgPanel'),
  gevolgSummary: document.getElementById('gevolgSummary'),
  dismissBtn: document.getElementById('dismissBtn'),
  domainPrompt: document.getElementById('domainPrompt'),
  domainList: document.getElementById('domainList'),
  promptHint: document.getElementById('promptHint'),
  domainYes: document.getElementById('domainYes'),
  domainNo: document.getElementById('domainNo'),
  domainNone: document.getElementById('domainNone'),
  generateBtn: document.getElementById('generateBtn'),
  viewer: document.getElementById('viewer'),
  viewerImage: document.getElementById('viewerImage'),
  viewerCaption: document.getElementById('viewerCaption'),
  viewerClose: document.getElementById('viewerClose'),
};

let shots = [];
const selected = new Set();
let gevolgOptions = [];
let gevolgSelection = [];
let inbreukSelection = [];
let newDomains = [];

// ---- gevolg / infractions --------------------------------------------

function summariseGevolg() {
  els.gevolgSummary.textContent = gevolgSelection.length
    ? gevolgSelection.map((s) => (s.text ? `${s.label} ${s.text}`.trim() : s.label)).join(' | ')
    : 'No measures selected.';
}

function readGevolg() {
  const out = [];
  els.gevolgPanel.querySelectorAll('.checkRow').forEach((row) => {
    const box = row.querySelector('input[type="checkbox"]');
    if (!box || !box.checked) return;
    const text = row.querySelector('.freeText');
    out.push({ label: box.dataset.label, text: text ? text.value.trim() : '' });
  });
  return out;
}

function onGevolgChanged() {
  gevolgSelection = readGevolg();
  els.gevolgPanel.querySelectorAll('.checkRow').forEach((row) => {
    const box = row.querySelector('input[type="checkbox"]');
    const text = row.querySelector('.freeText');
    if (text && box) text.disabled = !box.checked;
  });
  summariseGevolg();
}

function buildCheckRow(label, checked, onChange, needsText, textValue) {
  const row = document.createElement('div');
  row.className = 'checkRow';

  const wrap = document.createElement('label');
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.dataset.label = label;
  box.checked = checked;
  box.addEventListener('change', onChange);

  const span = document.createElement('span');
  span.textContent = label;
  wrap.appendChild(box);
  wrap.appendChild(span);
  row.appendChild(wrap);

  if (needsText) {
    const text = document.createElement('input');
    text.type = 'text';
    text.className = 'freeText';
    text.placeholder = 'Klik of tik om tekst in te voeren.';
    text.value = textValue || '';
    text.disabled = !checked;
    text.addEventListener('change', onChange);
    row.appendChild(text);
  }
  return row;
}

function buildGevolg() {
  els.gevolgPanel.innerHTML = '';
  const chosen = new Map(gevolgSelection.map((s) => [s.label, s.text || '']));
  const known = new Set(gevolgOptions.map((o) => o.label));
  const extras = gevolgSelection
    .filter((s) => !known.has(s.label))
    .map((s) => ({ label: s.label, needsText: false }));

  [...gevolgOptions, ...extras].forEach((opt) => {
    els.gevolgPanel.appendChild(
      buildCheckRow(opt.label, chosen.has(opt.label), onGevolgChanged, opt.needsText, chosen.get(opt.label))
    );
  });
  summariseGevolg();
}

function onInbreukChanged() {
  inbreukSelection = [];
  els.inbreukPanel.querySelectorAll('input[type="checkbox"]').forEach((box) => {
    if (box.checked) inbreukSelection.push(box.dataset.label);
  });
}

function buildInbreuk(options) {
  els.inbreukPanel.innerHTML = '';
  const chosen = new Set(inbreukSelection);
  options.forEach((label) => {
    els.inbreukPanel.appendChild(buildCheckRow(label, chosen.has(label), onInbreukChanged, false));
  });
}

// ---- screenshots -----------------------------------------------------

function updateCounts() {
  els.shotCount.textContent = shots.length ? `${selected.size} of ${shots.length} kept` : '';
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
    meta.appendChild(url);

    // shot.via is the link text that led here; null on the homepage itself.
    const via = document.createElement('div');
    via.className = 'via';
    if (shot.via) {
      via.textContent = `Reached by clicking “${shot.via}”`;
      via.title = via.textContent;
    } else {
      via.textContent = 'Homepage';
    }
    meta.appendChild(via);

    const sub = document.createElement('div');
    sub.className = 'sub';
    sub.innerHTML =
      `<span>Part ${shot.slice} of ${shot.sliceCount}</span>` +
      `<span>${shot.width}×${shot.height}</span>` +
      '<span>Right-click to enlarge</span>';
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

// ---- wiring ----------------------------------------------------------

els.selectAllBtn.addEventListener('click', () => {
  shots.forEach((s) => selected.add(s.id));
  renderShots();
});
els.selectNoneBtn.addEventListener('click', () => {
  selected.clear();
  renderShots();
});
els.dismissBtn.addEventListener('click', async () => {
  const result = await window.api.dismissCurrent();
  // The window stays open while other captures are waiting.
  if (result && result.next) await load();
});

// Resolves to the domains to add: [] when the agent says No.
function askAboutDomains() {
  return new Promise((resolve) => {
    els.domainList.innerHTML = '';
    els.promptHint.textContent =
      `Linked from this site and not yet in the sheet (${newDomains.length} found).`;

    newDomains.forEach((d, i) => {
      const row = document.createElement('div');
      row.className = 'domainRow';
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.id = 'dom' + i;
      box.checked = true;
      box.dataset.domain = d.domain;

      const info = document.createElement('label');
      info.className = 'dInfo';
      info.setAttribute('for', box.id);
      const name = document.createElement('div');
      name.className = 'dName';
      name.textContent = d.domain;
      const meta = document.createElement('div');
      meta.className = 'dMeta';
      meta.textContent = d.text
        ? `${d.count} link(s) — e.g. “${d.text}”`
        : `${d.count} link(s)`;
      info.appendChild(name);
      info.appendChild(meta);

      row.appendChild(box);
      row.appendChild(info);
      els.domainList.appendChild(row);
    });

    const finish = (accept) => {
      els.domainPrompt.hidden = true;
      els.domainYes.onclick = null;
      els.domainNo.onclick = null;
      els.domainNone.onclick = null;
      if (!accept) return resolve([]);
      resolve(
        [...els.domainList.querySelectorAll('input[type="checkbox"]')]
          .filter((b) => b.checked)
          .map((b) => b.dataset.domain)
      );
    };

    els.domainYes.onclick = () => finish(true);
    els.domainNo.onclick = () => finish(false);
    els.domainNone.onclick = () => {
      els.domainList.querySelectorAll('input[type="checkbox"]').forEach((b) => (b.checked = false));
    };
    els.domainPrompt.hidden = false;
  });
}

els.generateBtn.addEventListener('click', async () => {
  let addDomains = [];
  if (newDomains.length > 0) {
    addDomains = await askAboutDomains();
  }
  els.generateBtn.disabled = true;
  els.generateBtn.textContent = 'Generating…';
  try {
    const result = await window.api.captureGenerate({
      selectedIds: shots.filter((s) => selected.has(s.id)).map((s) => s.id),
      reference: els.ref.value.trim(),
      agentName: els.agentSelect.value,
      datum: els.datum.value.trim(),
      bron: els.bron.value.trim(),
      gevolg: gevolgSelection,
      inbreuk: inbreukSelection,
      addDomains,
    });
    // Another capture is waiting; stay open and move on to it.
    if (result && result.next) await load();
  } catch (err) {
    alert(err.message);
    els.generateBtn.disabled = false;
    els.generateBtn.textContent = 'Generate Report';
  }
});

let activeJobId = null;

function renderTabs(tabs) {
  els.tabStrip.innerHTML = '';
  if (!tabs || tabs.length < 2) return; // a single capture needs no tab strip
  tabs.forEach((t) => {
    const btn = document.createElement('button');
    btn.className = 'tab' + (t.id === activeJobId ? ' active' : '');
    btn.title = `${t.url} — ${t.shotCount} screenshot(s)`;
    btn.innerHTML = `<span class="tabRow">Row ${t.row + 1}</span>`;
    btn.appendChild(document.createTextNode(t.url.replace(/^https?:\/\//, '')));
    btn.addEventListener('click', async () => {
      if (t.id === activeJobId) return;
      await window.api.selectJob(t.id);
      await load();
    });
    els.tabStrip.appendChild(btn);
  });
}

async function load() {
  const ctx = await window.api.captureContext();
  activeJobId = ctx.jobId;

  // Reset per-capture state before rebuilding, so nothing leaks across tabs.
  shots = [];
  selected.clear();
  gevolgSelection = [];
  inbreukSelection = [];
  newDomains = [];

  renderTabs(ctx.tabs);

  els.site.textContent =
    ctx.rawUrl && ctx.rawUrl !== ctx.url ? `${ctx.url}   (cell: ${ctx.rawUrl})` : ctx.url;
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

  newDomains = ctx.newDomains || [];
  inbreukSelection = ctx.inbreuk || [];
  buildInbreuk(ctx.inbreukOptions || []);

  gevolgOptions = ctx.gevolgOptions || [];
  gevolgSelection = ctx.gevolg || [];
  buildGevolg();

  // The crawl already ran in the background; everything is kept by default.
  shots = ctx.shots || [];
  shots.forEach((s) => selected.add(s.id));
  renderShots();

  els.generateBtn.textContent = 'Generate Report';
}

// A capture finishing elsewhere adds a tab without disturbing this one.
window.api.onCaptureJobs((jobs) => {
  renderTabs(
    jobs.filter((j) => j.status === 'done').map((j) => ({
      id: j.id,
      row: j.row,
      url: j.url,
      shotCount: j.shotCount,
    }))
  );
});

// Clicking Review in the main window while this is open switches tab.
window.api.onReviewSwitch(() => load());

load();
