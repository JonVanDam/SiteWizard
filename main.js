const { app, BrowserWindow, BrowserView, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

let mainWindow;
let browserView;

// Seeded into the Gevolg options sheet the first time it has to be created.
// After that the sheet is the source of truth and can be edited in Excel.
const DEFAULT_GEVOLG_OPTIONS = [
  'Geen verder gevolg',
  'Toevoeging zwarte lijst (website + mirrors) + kennisgeving ISP’s',
  'Melding aan DNS BELGIUM',
  'Opstellen PV en doorsturen aan parket',
  'Doorsturen naar dienst controle & compliance voor verwijdering reclame (META)',
  'Doorsturen naar andere dienst :',
  'Andere :',
];
const DEFAULT_GEVOLG_SHEET = 'Gevolg opties';
const GEVOLG_SEPARATOR = ' | ';

// Names offered in the "Agent Name" dropdown when the settings file has none
// yet; taken from the combo box that used to live in the report template.
const DEFAULT_AGENT_NAMES = ['Eline GUERBAOUI', 'Barbara MASQUELIER'];

// Report references look like I085-2026: a letter, a zero-padded sequence and
// the year. New ones continue the highest sequence found anywhere in the file.
const REFERENCE_PATTERN = /\b([A-Z])(\d{2,4})-(\d{4})\b/g;
const REFERENCE_PREFIX = 'I';

// Capture geometry. A4 at 150 dpi, which keeps body text legible in the
// report without producing enormous PNGs.
const A4_WIDTH = 1240;
const A4_HEIGHT = 1754;
const MAX_PAGES = 20; // homepage plus same-domain links found on it
const MAX_SLICES_PER_PAGE = 4;
const PAGE_LOAD_TIMEOUT_MS = 15000;
const PAGE_SETTLE_MS = 1200;
// Word's "click here to enter text" prompt, carried over when options are
// pasted out of the report template. Stripped so only the label remains.
const GEVOLG_PROMPT = /klik of tik om tekst in te voeren\.?\s*$/i;

// Everything SiteWizard currently has loaded / is working through.
const state = {
  excelPath: null,
  workbook: null,
  sheetName: null,
  range: null, // XLSX.utils.decode_range() of the active sheet; range.s.r is treated as the header row
  urlColIdx: null,
  statusColIdx: null,
  commentColIdx: null, // null if no comment column configured
  gevolgColIdx: null, // null if no gevolg column configured
  gevolgOptions: [], // [{label, needsText}] read from the options sheet
  gevolgSheetName: null,
  // Last selection the user confirmed, carried to the next entry as a
  // pre-fill so a run of sites getting the same measures is quick to mark.
  gevolgSticky: [],
  refColIdx: null, // column holding the report reference, null if not configured
  currentRow: null, // 0-based sheet row index of the entry currently on screen

  // The capture window and everything it has collected for one entry.
  capture: null, // { row, url, shots: [{id, pageUrl, slice, buffer, width, height}] }

  // <extLst> blocks captured from the source workbook, re-applied after
  // every save (see restoreSheetExtensions).
  preservedExtensions: {},

  templatePath: null,
  outputFolder: null,

  // Linear "previously opened entry" history, independent of undo/redo.
  history: [],
  historyPos: -1,

  // Undo/redo stack of reversible Excel writes (status marks + comments).
  actionLog: [],
  actionPos: -1, // index of the last applied action; -1 means nothing applied

  // Per-row in-memory cache of captured screenshots ({buffer, width, height}[]),
  // lazily reloaded from an existing report's embedded images if not yet in memory.
  screenshotsByRow: {},
};

// ---- Remembered settings (columns, template + output paths) --------------

function settingsFilePath() {
  return path.join(app.getPath('userData'), 'sitewizard-settings.json');
}

function loadSettings() {
  try {
    return JSON.parse(fs.readFileSync(settingsFilePath(), 'utf8'));
  } catch {
    return {};
  }
}

function saveSettings(partial) {
  const merged = { ...loadSettings(), ...partial };
  fs.writeFileSync(settingsFilePath(), JSON.stringify(merged, null, 2));
}

function createWindow() {
  const saved = loadSettings();
  state.templatePath = saved.templatePath || null;
  state.outputFolder = saved.outputFolder || null;

  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile('index.html');

  browserView = new BrowserView();
  mainWindow.setBrowserView(browserView);
  browserView.setBounds({ x: 0, y: 0, width: 0, height: 0 });
  browserView.webContents.loadURL('about:blank');

  browserView.webContents.on('did-navigate', (e, url) => {
    mainWindow.webContents.send('nav-state', { url });
  });
  browserView.webContents.on('did-navigate-in-page', (e, url) => {
    mainWindow.webContents.send('nav-state', { url });
  });
  browserView.webContents.on('did-fail-load', (e, errorCode, errorDescription, validatedURL) => {
    if (errorCode === -3) return; // ERR_ABORTED, usually just a superseded navigation
    log(`Failed to load ${validatedURL || '(unknown URL)'}: ${errorDescription} (${errorCode})`);
  });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

function log(msg) {
  if (mainWindow) mainWindow.webContents.send('log', msg);
}

// ---- Preserving what SheetJS's writer drops ------------------------------
// The community writer rebuilds each worksheet and loses the <extLst> block,
// which is where Excel keeps x14 data validations — the dropdown lists whose
// source range lives on another sheet. Losing those silently guts a workbook
// that relies on them, so the blocks are captured when the file is opened and
// put back after every save.

function sheetFileMap(zip) {
  const wbXml = zip.files['xl/workbook.xml'] && zip.files['xl/workbook.xml'].asText();
  const relsXml =
    zip.files['xl/_rels/workbook.xml.rels'] && zip.files['xl/_rels/workbook.xml.rels'].asText();
  if (!wbXml || !relsXml) return {};

  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0]);
    const target = /Target="([^"]+)"/.exec(m[0]);
    if (id && target) rels[id[1]] = target[1].replace(/^\/?(xl\/)?/, 'xl/');
  }

  const map = {};
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /name="([^"]+)"/.exec(m[0]);
    const rid = /r:id="([^"]+)"/.exec(m[0]);
    if (name && rid && rels[rid[1]]) {
      map[decodeXml(name[1])] = rels[rid[1]];
    }
  }
  return map;
}

function decodeXml(s) {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// Returns { sheetName: '<extLst>…</extLst>' } for sheets that have one.
function captureSheetExtensions(buffer) {
  const out = {};
  try {
    const zip = new PizZip(buffer);
    const map = sheetFileMap(zip);
    for (const [name, file] of Object.entries(map)) {
      const entry = zip.files[file];
      if (!entry) continue;
      const xml = entry.asText();
      const end = xml.lastIndexOf('</extLst>');
      if (end === -1) continue;
      const start = xml.lastIndexOf('<extLst>', end);
      if (start === -1) continue;
      // Only keep it if it's the worksheet-level block (nothing but whitespace
      // and the closing tag after it).
      const tail = xml.slice(end + '</extLst>'.length);
      if (!/^\s*<\/worksheet>\s*$/.test(tail)) continue;
      const block = xml.slice(start, end + '</extLst>'.length);
      if (!block.includes('x14:dataValidation')) continue;
      // xr:uid attributes reference a namespace SheetJS's <worksheet> element
      // won't declare; they're only revision ids, so drop them.
      out[name] = block.replace(/\sxr:uid="[^"]*"/g, '');
    }
  } catch (err) {
    log('Could not read workbook extensions: ' + err.message);
  }
  return out;
}

function restoreSheetExtensions(filePath, preserved) {
  const names = Object.keys(preserved || {});
  if (names.length === 0) return;
  try {
    const zip = new PizZip(fs.readFileSync(filePath));
    const map = sheetFileMap(zip);
    let changed = 0;
    for (const name of names) {
      const file = map[name];
      const entry = file && zip.files[file];
      if (!entry) continue;
      let xml = entry.asText();
      if (xml.includes('x14:dataValidation')) continue; // already there
      if (!xml.includes('</worksheet>')) continue;
      xml = xml.replace('</worksheet>', preserved[name] + '</worksheet>');
      zip.file(file, xml);
      changed += 1;
    }
    if (changed > 0) {
      fs.writeFileSync(filePath, zip.generate({ type: 'nodebuffer' }));
    }
  } catch (err) {
    log('Could not restore workbook extensions: ' + err.message);
  }
}

// Every write to the source workbook goes through here.
function saveWorkbook() {
  XLSX.writeFile(state.workbook, state.excelPath);
  restoreSheetExtensions(state.excelPath, state.preservedExtensions);
}

// ---- Excel helpers -------------------------------------------------------
// We read/write individual cells directly on the original sheet object
// (rather than round-tripping through XLSX.utils.aoa_to_sheet) so that
// untouched cells, formulas, and formatting elsewhere in the workbook are
// left alone as much as SheetJS's community edition allows.

function colSpecToIndex(spec) {
  const trimmed = spec.trim();
  if (/^\d+$/.test(trimmed)) {
    // Plain number: treat as a 1-based column number (1 = A).
    return parseInt(trimmed, 10) - 1;
  }
  const letter = trimmed.toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(letter)) return null;
  let idx = 0;
  for (let i = 0; i < letter.length; i++) {
    idx = idx * 26 + (letter.charCodeAt(i) - 64);
  }
  return idx - 1;
}

function cellValue(sheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
  return cell ? cell.v : undefined;
}

function setCell(sheet, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c });
  sheet[addr] = { t: 's', v: String(value) };
}

function resolveColumn(sheet, range, spec, createIfMissing) {
  const target = spec.trim().toLowerCase();
  for (let c = range.s.c; c <= range.e.c; c++) {
    const v = cellValue(sheet, range.s.r, c);
    if (v !== undefined && v.toString().trim().toLowerCase() === target) {
      return c;
    }
  }
  const letterIdx = colSpecToIndex(spec);
  if (letterIdx !== null) return letterIdx;
  if (createIfMissing) {
    const newCol = range.e.c + 1;
    setCell(sheet, range.s.r, newCol, spec.trim());
    range.e.c = newCol;
    sheet['!ref'] = XLSX.utils.encode_range(range);
    return newCol;
  }
  return -1;
}

// Excel represents a row filtered out by an AutoFilter as a hidden row, so
// honouring the sheet's filter is just a matter of skipping hidden rows.
function isRowHidden(sheet, r) {
  const rows = sheet['!rows'];
  return !!(rows && rows[r] && rows[r].hidden);
}

function findNextRow(sheet, range, urlCol, statusCol, startRow) {
  for (let r = startRow; r <= range.e.r; r++) {
    if (isRowHidden(sheet, r)) continue;
    const url = cellValue(sheet, r, urlCol);
    const status = cellValue(sheet, r, statusCol);
    if (url !== undefined && String(url).trim() && (status === undefined || String(status).trim() === '')) {
      return r;
    }
  }
  return -1;
}

// Counts what's left to work through, for the message shown after Apply Columns.
function countRows(sheet, range, urlCol, statusCol) {
  let visible = 0;
  let hidden = 0;
  let pending = 0;
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const url = cellValue(sheet, r, urlCol);
    if (url === undefined || !String(url).trim()) continue;
    if (isRowHidden(sheet, r)) {
      hidden += 1;
      continue;
    }
    visible += 1;
    const status = cellValue(sheet, r, statusCol);
    if (status === undefined || String(status).trim() === '') pending += 1;
  }
  return { visible, hidden, pending };
}

// ---- Gevolg (follow-up measures) --------------------------------------
// The option list lives in its own sheet in the loaded workbook so it can be
// maintained in Excel rather than in code. If that sheet isn't there yet it
// is created and seeded with DEFAULT_GEVOLG_OPTIONS, so a workbook that has
// never been used with SiteWizard still works on first run.

// An option whose label ends in ":" expects free text after it, e.g.
// "Andere :" becomes "Andere : <what the user typed>".
function parseGevolgOption(raw) {
  const label = String(raw).replace(GEVOLG_PROMPT, '').trim();
  if (!label) return null;
  return { label, needsText: /:\s*$/.test(label) };
}

function loadGevolgOptions(sheetName) {
  const name = (sheetName || '').trim() || DEFAULT_GEVOLG_SHEET;
  let sheet = state.workbook.Sheets[name];

  if (!sheet) {
    const rows = [[name], ...DEFAULT_GEVOLG_OPTIONS.map((o) => [o])];
    sheet = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(state.workbook, sheet, name);
    saveWorkbook();
    log(`Created sheet "${name}" with ${DEFAULT_GEVOLG_OPTIONS.length} default options. Edit it in Excel to change the list.`);
  }

  const options = [];
  if (sheet['!ref']) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    // Row 1 is the header; options run down the first column from row 2.
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const v = cellValue(sheet, r, range.s.c);
      if (v === undefined || !String(v).trim()) continue;
      const opt = parseGevolgOption(v);
      if (opt) options.push(opt);
    }
  }
  state.gevolgSheetName = name;
  state.gevolgOptions = options;
  return options;
}

// Turns a stored cell back into a selection. Values that no longer match any
// option are kept as-is rather than dropped, so editing the options sheet
// can't silently erase decisions already recorded in the sheet.
function parseGevolgCell(value) {
  if (value === undefined || value === null) return [];
  return String(value)
    .split('|')
    .map((s) => s.trim())
    .filter(Boolean)
    .map((piece) => {
      for (const opt of state.gevolgOptions) {
        if (opt.needsText && piece.startsWith(opt.label)) {
          return { label: opt.label, text: piece.slice(opt.label.length).trim() };
        }
        if (!opt.needsText && piece === opt.label) {
          return { label: opt.label, text: '' };
        }
      }
      return { label: piece, text: '' };
    });
}

function formatGevolg(selection) {
  return (selection || [])
    .map((s) => (s.text ? `${s.label} ${s.text}`.trim() : s.label))
    .join(GEVOLG_SEPARATOR);
}

// A Url cell often lists a site together with its mirrors, e.g.
// "winbeast.com + winbeast1.com". Only the first entry is browsed and
// reported on; the rest are left in the sheet untouched.
function primaryUrl(raw) {
  if (raw === undefined || raw === null) return '';
  const parts = String(raw)
    .split(/[+,;\n\r]| {2,}/)
    .map((s) => s.trim())
    .filter(Boolean);
  return parts[0] || '';
}

function rowUrl(sheet, row) {
  return primaryUrl(cellValue(sheet, row, state.urlColIdx));
}

function navigateBrowserView(url) {
  let target = url;
  if (!/^https?:\/\//i.test(target)) {
    target = 'https://' + target;
  }
  browserView.webContents.loadURL(target).catch((err) => {
    log('Failed to load ' + target + ': ' + err.message);
  });
}

function resetWorkflowState() {
  state.history = [];
  state.historyPos = -1;
  state.actionLog = [];
  state.actionPos = -1;
  state.screenshotsByRow = {};
  state.gevolgSticky = [];
}

function recordAction(action) {
  state.actionLog = state.actionLog.slice(0, state.actionPos + 1);
  state.actionLog.push(action);
  state.actionPos = state.actionLog.length - 1;
}

// ---- Report helpers --------------------------------------------------

function reportPathFor(url) {
  const safeName = String(url).replace(/[^a-z0-9]/gi, '_').slice(0, 80);
  return path.join(state.outputFolder, `report_${safeName}.docx`);
}

function writeReport(context, screenshots) {
  const content = fs.readFileSync(state.templatePath, 'binary');
  const zip = new PizZip(content);
  const maxWidth = 600;

  // Tag values are 1-based: the image module treats a falsy tag value (0 is
  // falsy in JS) as "no image" and silently drops it, so index 0 can't be used.
  const imageModule = new ImageModule({
    getImage: (tagValue) => screenshots[tagValue - 1].buffer,
    getSize: (img, tagValue) => {
      const shot = screenshots[tagValue - 1];
      const scale = shot.width > maxWidth ? maxWidth / shot.width : 1;
      return [Math.round(shot.width * scale), Math.round(shot.height * scale)];
    },
  });

  const doc = new Docxtemplater(zip, {
    modules: [imageModule],
    paragraphLoop: true,
    linebreaks: true,
  });

  try {
    doc.render({
      site: context.site,
      refnr: context.refnr || '',
      datum: context.datum || '',
      controleur: context.controleur || '',
      bron: context.bron || '',
      screenshots: screenshots.map((_, i) => i + 1),
      ...gevolgTagValues(context.gevolg),
    });
  } catch (err) {
    throw new Error(
      'Failed to fill in the report template. If your template still has a single ' +
        '{%screenshots} tag from before, change it to a loop: {#screenshots}{%.}{/screenshots}. ' +
        'Original error: ' + err.message
    );
  }
  return doc.getZip().generate({ type: 'nodebuffer' });
}

// ---- Report references ------------------------------------------------
// A row that already carries a reference keeps it, so regenerating a report
// doesn't renumber it. Otherwise the next free sequence is taken from the
// highest reference anywhere in the workbook.

function scanHighestReference(year) {
  let highest = 0;
  if (!state.workbook) return highest;
  for (const name of state.workbook.SheetNames) {
    const sheet = state.workbook.Sheets[name];
    if (!sheet || !sheet['!ref']) continue;
    for (const addr of Object.keys(sheet)) {
      if (addr[0] === '!') continue;
      const cell = sheet[addr];
      if (!cell || cell.v === undefined) continue;
      const text = String(cell.v);
      REFERENCE_PATTERN.lastIndex = 0;
      let m;
      while ((m = REFERENCE_PATTERN.exec(text)) !== null) {
        if (m[1] !== REFERENCE_PREFIX) continue;
        if (Number(m[3]) !== year) continue;
        highest = Math.max(highest, Number(m[2]));
      }
    }
  }
  return highest;
}

function formatReference(seq, year) {
  return `${REFERENCE_PREFIX}${String(seq).padStart(3, '0')}-${year}`;
}

// Returns { reference, isNew } for a row.
function referenceForRow(row) {
  const year = new Date().getFullYear();
  if (state.refColIdx != null && state.workbook) {
    const sheet = state.workbook.Sheets[state.sheetName];
    const existing = cellValue(sheet, row, state.refColIdx);
    if (existing !== undefined && String(existing).trim()) {
      return { reference: String(existing).trim(), isNew: false };
    }
  }
  return { reference: formatReference(scanHighestReference(year) + 1, year), isNew: true };
}

// The sheet's own Bron value for a row, if the sheet has such a column.
function bronForRow(row) {
  if (!state.workbook || !state.sheetName) return '';
  const sheet = state.workbook.Sheets[state.sheetName];
  const idx = resolveColumn(sheet, state.range, 'Bron', false);
  if (idx === -1) return '';
  const v = cellValue(sheet, row, idx);
  return v === undefined ? '' : String(v).trim();
}

// ---- Crawling and capturing --------------------------------------------

function absoluteUrl(url) {
  return /^https?:\/\//i.test(url) ? url : 'https://' + url;
}

function sameSite(a, b) {
  try {
    const strip = (h) => h.replace(/^www\./i, '').toLowerCase();
    return strip(new URL(a).hostname) === strip(new URL(b).hostname);
  } catch {
    return false;
  }
}

function loadInWindow(win, url) {
  return new Promise((resolve) => {
    let settled = false;
    const done = (ok) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(ok);
    };
    const timer = setTimeout(() => done(false), PAGE_LOAD_TIMEOUT_MS);
    win.webContents.once('did-finish-load', () => done(true));
    win.webContents.once('did-fail-load', (e, code) => done(code === -3));
    win.webContents.loadURL(url).catch(() => done(false));
  });
}

// Captures a page top to bottom in viewport-height slices.
// Pages are rendered at A4 width so they lay out like a printed page rather
// than a wide desktop window. The viewport height is whatever the display
// allows -- a window cannot be taller than the screen work area, and neither
// offscreen rendering nor the DevTools protocol lifts that limit -- so tall
// pages are captured as several stacked slices instead of one A4 sheet.

function viewportHeightFor(win) {
  return win.getContentBounds().height;
}

async function documentHeight(win) {
  try {
    return await win.webContents.executeJavaScript(
      'Math.max(document.body ? document.body.scrollHeight : 0,' +
        ' document.documentElement ? document.documentElement.scrollHeight : 0,' +
        ' window.innerHeight)'
    );
  } catch {
    return viewportHeightFor(win);
  }
}

async function captureSlices(win, pageUrl, onShot) {
  await new Promise((r) => setTimeout(r, PAGE_SETTLE_MS));
  const viewport = Math.max(1, viewportHeightFor(win));
  const height = await documentHeight(win);
  const slices = Math.max(1, Math.min(MAX_SLICES_PER_PAGE, Math.ceil(height / viewport)));

  for (let i = 0; i < slices; i++) {
    try {
      await win.webContents.executeJavaScript(`window.scrollTo(0, ${i * viewport})`);
      await new Promise((r) => setTimeout(r, 350));
      const image = await win.webContents.capturePage();
      const size = image.getSize();
      if (size.width === 0 || size.height === 0) continue;
      onShot({
        pageUrl,
        slice: i + 1,
        sliceCount: slices,
        buffer: image.toPNG(),
        thumb: image.resize({ width: 320 }).toDataURL(),
        width: size.width,
        height: size.height,
      });
    } catch (err) {
      log(`Capture failed for ${pageUrl} (slice ${i + 1}): ${err.message}`);
    }
  }
}

async function crawlAndCapture(startUrl, progress) {
  const root = absoluteUrl(startUrl);
  const win = new BrowserWindow({
    show: false,
    // Without useContentSize these are the outer frame dimensions, and the
    // page would render shorter than A4 with the wrong aspect ratio.
    useContentSize: true,
    width: A4_WIDTH,
    height: A4_HEIGHT, // clamped to the work area by the OS; that is fine
    frame: false,
    webPreferences: { offscreen: false, javascript: true, images: true, sandbox: true },
  });
  win.webContents.setAudioMuted(true);

  const shots = [];
  let nextId = 1;
  const collect = (shot) => {
    shots.push({ id: String(nextId++), ...shot });
    progress({ phase: 'capturing', shots: shots.length, page: shot.pageUrl });
  };

  try {
    progress({ phase: 'loading', page: root });
    const ok = await loadInWindow(win, root);
    if (!ok) {
      progress({ phase: 'warning', message: `Homepage did not load cleanly: ${root}` });
    }
    const landed = win.webContents.getURL() || root;
    await captureSlices(win, landed, collect);

    let links = [];
    try {
      links = await win.webContents.executeJavaScript(
        "Array.from(document.querySelectorAll('a[href]')).map(a => a.href)"
      );
    } catch {
      /* no links is fine */
    }

    const seen = new Set([landed.split('#')[0]]);
    const queue = [];
    for (const href of links) {
      const clean = String(href).split('#')[0];
      if (!/^https?:\/\//i.test(clean)) continue;
      if (!sameSite(clean, landed)) continue;
      if (seen.has(clean)) continue;
      seen.add(clean);
      queue.push(clean);
      if (queue.length >= MAX_PAGES - 1) break;
    }
    progress({ phase: 'queued', total: queue.length + 1 });

    for (let i = 0; i < queue.length; i++) {
      if (!state.capture || state.capture.cancelled) break;
      progress({ phase: 'loading', page: queue[i], index: i + 2, total: queue.length + 1 });
      const loaded = await loadInWindow(win, queue[i]);
      if (!loaded) continue;
      await captureSlices(win, queue[i], collect);
    }
  } finally {
    if (!win.isDestroyed()) win.destroy();
  }
  return shots;
}

// ---- Entry navigation ------------------------------------------------

function buildEntryInfo(row) {
  const canUndo = state.actionPos >= 0;
  const canRedo = state.actionPos < state.actionLog.length - 1;
  const canPrevious = state.historyPos > 0;

  if (row == null || row === -1) {
    return { done: true, canUndo, canRedo, canPrevious };
  }
  const sheet = state.workbook.Sheets[state.sheetName];
  const url = rowUrl(sheet, row);
  const commentEnabled = state.commentColIdx != null;
  const comment = commentEnabled ? cellValue(sheet, row, state.commentColIdx) : undefined;
  const hasReport = !!(state.outputFolder && fs.existsSync(reportPathFor(url)));

  // A row that already has measures recorded always wins; only a blank one
  // inherits the previous entry's selection as a pre-fill.
  const gevolgEnabled = state.gevolgColIdx != null;
  const stored = gevolgEnabled ? cellValue(sheet, row, state.gevolgColIdx) : undefined;
  const hasStored = stored !== undefined && String(stored).trim() !== '';
  const gevolg = hasStored ? parseGevolgCell(stored) : state.gevolgSticky.slice();

  return {
    done: false,
    row,
    url,
    comment: comment == null ? '' : String(comment),
    commentEnabled,
    hasReport,
    gevolgEnabled,
    gevolgOptions: state.gevolgOptions,
    gevolg,
    gevolgPrefilled: !hasStored && gevolg.length > 0,
    canUndo,
    canRedo,
    canPrevious,
  };
}

// Moves to a *new* row reached by normal forward progress (initial load,
// False Positive / Not Sure / Skip, or Redo), pushing it onto the "previous
// entry" history and truncating any stale forward history.
function moveToRow(row) {
  const isValid = row !== null && row !== -1;
  if (isValid) {
    state.history = state.history.slice(0, state.historyPos + 1);
    state.history.push(row);
    state.historyPos = state.history.length - 1;
  }
  state.currentRow = isValid ? row : null;
  navigateBrowserView(isValid ? rowUrl(state.workbook.Sheets[state.sheetName], row) : 'about:blank');
  return buildEntryInfo(state.currentRow);
}

// ---- IPC: file dialogs / setup -------------------------------------------

ipcMain.handle('dialog:openExcel', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const filePath = result.filePaths[0];
  // cellStyles is what makes SheetJS parse row properties, which is how a
  // filtered-out row is represented (hidden="1"). Without it '!rows' is
  // undefined and SiteWizard can't tell which rows the filter excludes.
  const workbook = XLSX.readFile(filePath, { cellStyles: true });
  state.excelPath = filePath;
  state.workbook = workbook;
  state.preservedExtensions = captureSheetExtensions(fs.readFileSync(filePath));
  const extSheets = Object.keys(state.preservedExtensions);
  if (extSheets.length > 0) {
    log(`Preserving dropdown lists on sheet(s): ${extSheets.join(', ')}.`);
  }
  return { filePath, sheetNames: workbook.SheetNames };
});

ipcMain.handle('excel:selectSheet', async (event, sheetName) => {
  if (!state.workbook) throw new Error('No workbook loaded');
  const sheet = state.workbook.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) throw new Error('Sheet not found or empty: ' + sheetName);
  state.sheetName = sheetName;
  state.range = XLSX.utils.decode_range(sheet['!ref']);
  resetWorkflowState();
  const header = [];
  for (let c = state.range.s.c; c <= state.range.e.c; c++) {
    header.push(cellValue(sheet, state.range.s.r, c));
  }
  return { header };
});

ipcMain.handle('excel:setColumns', async (event, cfg) => {
  if (!state.workbook || !state.sheetName) throw new Error('No sheet loaded');
  const sheet = state.workbook.Sheets[state.sheetName];
  const urlIdx = resolveColumn(sheet, state.range, cfg.urlCol, false);
  if (urlIdx === -1) throw new Error('URL column not found: ' + cfg.urlCol);
  const statusIdx = resolveColumn(sheet, state.range, cfg.statusCol, true);
  let commentIdx = null;
  if (cfg.commentCol && cfg.commentCol.trim()) {
    commentIdx = resolveColumn(sheet, state.range, cfg.commentCol, true);
  }
  let gevolgIdx = null;
  if (cfg.gevolgCol && cfg.gevolgCol.trim()) {
    gevolgIdx = resolveColumn(sheet, state.range, cfg.gevolgCol, true);
  }
  let refIdx = null;
  if (cfg.refCol && cfg.refCol.trim()) {
    refIdx = resolveColumn(sheet, state.range, cfg.refCol, true);
  }
  state.urlColIdx = urlIdx;
  state.statusColIdx = statusIdx;
  state.commentColIdx = commentIdx;
  state.gevolgColIdx = gevolgIdx;
  state.refColIdx = refIdx;
  resetWorkflowState();

  // Only touch the options sheet when a gevolg column is actually in use, so
  // workbooks that don't need the feature aren't modified at all.
  state.gevolgOptions = [];
  state.gevolgSheetName = null;
  if (gevolgIdx != null) {
    const opts = loadGevolgOptions(cfg.gevolgSheet);
    if (opts.length === 0) {
      log(`Sheet "${state.gevolgSheetName}" has no options listed under its header row — the Gevolg panel will be empty.`);
    }
  }

  const counts = countRows(sheet, state.range, urlIdx, statusIdx);
  if (counts.hidden > 0) {
    log(`Sheet filter active: ${counts.hidden} row(s) filtered out and will be skipped, ${counts.visible} visible (${counts.pending} still without a status).`);
  } else {
    log(`${counts.visible} row(s) with a URL, ${counts.pending} still without a status.`);
  }

  const startRow = findNextRow(sheet, state.range, urlIdx, statusIdx, state.range.s.r + 1);
  saveSettings({
    urlCol: cfg.urlCol,
    statusCol: cfg.statusCol,
    commentCol: cfg.commentCol || '',
    gevolgCol: cfg.gevolgCol || '',
    gevolgSheet: cfg.gevolgSheet || '',
    refCol: cfg.refCol || '',
  });
  return moveToRow(startRow);
});

ipcMain.handle('settings:load', () => loadSettings());

ipcMain.handle('dialog:openTemplate', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'Word Template', extensions: ['docx'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  state.templatePath = result.filePaths[0];
  saveSettings({ templatePath: state.templatePath });
  return state.templatePath;
});

ipcMain.handle('dialog:selectOutputFolder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  state.outputFolder = result.filePaths[0];
  saveSettings({ outputFolder: state.outputFolder });
  return state.outputFolder;
});

// ---- IPC: browsing ---------------------------------------------------

ipcMain.on('browserview:bounds', (event, rect) => {
  if (!browserView) return;
  browserView.setBounds({
    x: Math.round(rect.x),
    y: Math.round(rect.y),
    width: Math.max(0, Math.round(rect.width)),
    height: Math.max(0, Math.round(rect.height)),
  });
});

ipcMain.handle('browserview:back', () => {
  if (browserView.webContents.canGoBack()) browserView.webContents.goBack();
});
ipcMain.handle('browserview:forward', () => {
  if (browserView.webContents.canGoForward()) browserView.webContents.goForward();
});
ipcMain.handle('browserview:reload', () => {
  browserView.webContents.reload();
});

// ---- IPC: workflow actions ---------------------------------------------

ipcMain.handle('entries:markStatus', async (event, status) => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const oldValue = cellValue(sheet, row, state.statusColIdx);
  const oldStr = oldValue === undefined ? '' : String(oldValue);
  setCell(sheet, row, state.statusColIdx, status);
  saveWorkbook();
  recordAction({ type: 'status', row, colIdx: state.statusColIdx, oldValue: oldStr, newValue: status });
  log(`Row ${row + 1}: marked "${status}" and saved.`);
  const nextRow = findNextRow(sheet, state.range, state.urlColIdx, state.statusColIdx, row + 1);
  return moveToRow(nextRow);
});

ipcMain.handle('entries:skip', async () => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  log(`Row ${row + 1}: skipped (source file not modified).`);
  const nextRow = findNextRow(sheet, state.range, state.urlColIdx, state.statusColIdx, row + 1);
  return moveToRow(nextRow);
});

ipcMain.handle('entries:previous', async () => {
  if (state.historyPos <= 0) throw new Error('No previous entry');
  state.historyPos -= 1;
  const row = state.history[state.historyPos];
  state.currentRow = row;
  const sheet = state.workbook.Sheets[state.sheetName];
  navigateBrowserView(rowUrl(sheet, row));
  log(`Row ${row + 1}: viewing previous entry.`);
  return buildEntryInfo(row);
});

ipcMain.handle('entries:undo', async () => {
  if (state.actionPos < 0) throw new Error('Nothing to undo');
  const action = state.actionLog[state.actionPos];
  const sheet = state.workbook.Sheets[state.sheetName];
  setCell(sheet, action.row, action.colIdx, action.oldValue);
  saveWorkbook();
  state.actionPos -= 1;
  state.currentRow = action.row;
  navigateBrowserView(rowUrl(sheet, action.row));
  log(`Row ${action.row + 1}: undid ${action.type} change.`);
  return buildEntryInfo(action.row);
});

ipcMain.handle('entries:redo', async () => {
  if (state.actionPos >= state.actionLog.length - 1) throw new Error('Nothing to redo');
  state.actionPos += 1;
  const action = state.actionLog[state.actionPos];
  const sheet = state.workbook.Sheets[state.sheetName];
  setCell(sheet, action.row, action.colIdx, action.newValue);
  saveWorkbook();
  state.currentRow = action.row;
  navigateBrowserView(rowUrl(sheet, action.row));
  log(`Row ${action.row + 1}: redid ${action.type} change.`);
  return buildEntryInfo(action.row);
});

ipcMain.handle('entries:saveComment', async (event, value) => {
  if (state.currentRow == null || state.currentRow === -1) return false;
  if (state.commentColIdx == null) return false;
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const oldValue = cellValue(sheet, row, state.commentColIdx);
  const oldStr = oldValue === undefined ? '' : String(oldValue);
  if (oldStr === value) return false;
  setCell(sheet, row, state.commentColIdx, value);
  saveWorkbook();
  recordAction({ type: 'comment', row, colIdx: state.commentColIdx, oldValue: oldStr, newValue: value });
  log(`Row ${row + 1}: comment saved.`);
  return true;
});

// ---- IPC: capture window ---------------------------------------------
// "Generate Report" no longer writes anything on its own. It opens a capture
// window which crawls the site, offers the screenshots for selection, and
// collects the agent name and gevolg measures. Only when that window's own
// Generate button is pressed is the report written and the sheet updated.

let captureWindow = null;

function captureProgress(payload) {
  if (captureWindow && !captureWindow.isDestroyed()) {
    captureWindow.webContents.send('capture-progress', payload);
  }
}

ipcMain.handle('capture:open', async () => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry to report on');
  if (!state.templatePath) throw new Error('Select a report template first');
  if (!state.outputFolder) throw new Error('Select an output folder first');
  if (captureWindow && !captureWindow.isDestroyed()) {
    captureWindow.focus();
    return true;
  }

  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  state.capture = { row, url: rowUrl(sheet, row), shots: [], cancelled: false };

  captureWindow = new BrowserWindow({
    width: 1280,
    height: 900,
    parent: mainWindow,
    title: 'Generate Report',
    backgroundColor: '#1b1b1d',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  captureWindow.setMenu(null);
  captureWindow.loadFile('capture.html');
  captureWindow.on('closed', () => {
    captureWindow = null;
    if (state.capture) state.capture.cancelled = true;
  });
  return true;
});

// The capture window asks for its own context once it has loaded.
ipcMain.handle('capture:context', async () => {
  if (!state.capture) throw new Error('No capture in progress');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.capture.row;
  const { reference, isNew } = referenceForRow(row);
  const saved = loadSettings();

  const stored = state.gevolgColIdx != null ? cellValue(sheet, row, state.gevolgColIdx) : undefined;
  const gevolg =
    stored !== undefined && String(stored).trim() !== ''
      ? parseGevolgCell(stored)
      : state.gevolgSticky.slice();

  return {
    row,
    url: state.capture.url,
    rawUrl: String(cellValue(sheet, row, state.urlColIdx) ?? ''),
    reference,
    referenceIsNew: isNew,
    gevolgEnabled: state.gevolgColIdx != null,
    gevolgOptions: state.gevolgOptions,
    gevolg,
    agentNames: saved.agentNames && saved.agentNames.length ? saved.agentNames : DEFAULT_AGENT_NAMES,
    agentName: saved.agentName || '',
    bron: bronForRow(row),
  };
});

ipcMain.handle('capture:run', async () => {
  if (!state.capture) throw new Error('No capture in progress');
  state.capture.shots = [];
  const shots = await crawlAndCapture(state.capture.url, captureProgress);
  if (!state.capture) return [];
  state.capture.shots = shots;
  captureProgress({ phase: 'done', shots: shots.length });
  // Buffers stay in the main process; the window only needs the thumbnails.
  return shots.map((s) => ({
    id: s.id,
    pageUrl: s.pageUrl,
    slice: s.slice,
    sliceCount: s.sliceCount,
    thumb: s.thumb,
    width: s.width,
    height: s.height,
  }));
});

// Full-resolution image for the fullscreen viewer.
ipcMain.handle('capture:image', async (event, id) => {
  if (!state.capture) return null;
  const shot = state.capture.shots.find((s) => s.id === id);
  return shot ? 'data:image/png;base64,' + shot.buffer.toString('base64') : null;
});

ipcMain.handle('capture:cancel', async () => {
  if (state.capture) state.capture.cancelled = true;
  if (captureWindow && !captureWindow.isDestroyed()) captureWindow.close();
  return true;
});

// This is the only place a report is written and the sheet updated.
ipcMain.handle('capture:generate', async (event, payload) => {
  if (!state.capture) throw new Error('No capture in progress');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.capture.row;
  const url = state.capture.url;

  const chosen = (payload.selectedIds || [])
    .map((id) => state.capture.shots.find((s) => s.id === id))
    .filter(Boolean);
  if (chosen.length === 0) throw new Error('Select at least one screenshot');

  const context = {
    site: url,
    refnr: payload.reference || '',
    datum: payload.datum || new Date().toLocaleDateString('nl-BE'),
    controleur: payload.agentName || '',
    bron: payload.bron || '',
    gevolg: payload.gevolg || [],
  };

  const outBuffer = writeReport(context, chosen);
  const outPath = reportPathFor(url);
  fs.writeFileSync(outPath, outBuffer);
  state.screenshotsByRow[row] = chosen.map((s) => ({
    buffer: s.buffer,
    width: s.width,
    height: s.height,
  }));

  // Sheet writes happen only now, never while the capture window is open.
  let wrote = [];
  if (state.gevolgColIdx != null) {
    const value = formatGevolg(payload.gevolg || []);
    const old = cellValue(sheet, row, state.gevolgColIdx);
    const oldStr = old === undefined ? '' : String(old);
    if (oldStr !== value) {
      setCell(sheet, row, state.gevolgColIdx, value);
      recordAction({ type: 'gevolg', row, colIdx: state.gevolgColIdx, oldValue: oldStr, newValue: value });
      wrote.push('gevolg');
    }
    state.gevolgSticky = (payload.gevolg || []).map((s) => ({ label: s.label, text: s.text || '' }));
  }
  if (state.refColIdx != null && context.refnr) {
    const old = cellValue(sheet, row, state.refColIdx);
    const oldStr = old === undefined ? '' : String(old);
    if (oldStr !== context.refnr) {
      setCell(sheet, row, state.refColIdx, context.refnr);
      recordAction({ type: 'reference', row, colIdx: state.refColIdx, oldValue: oldStr, newValue: context.refnr });
      wrote.push('reference');
    }
  }
  if (wrote.length > 0) saveWorkbook();

  saveSettings({ agentName: payload.agentName || '' });
  log(`Row ${row + 1}: report ${context.refnr} saved with ${chosen.length} screenshot(s): ${outPath}`);
  if (wrote.length > 0) log(`Row ${row + 1}: wrote ${wrote.join(' and ')} to the sheet.`);

  if (captureWindow && !captureWindow.isDestroyed()) captureWindow.close();
  if (mainWindow) mainWindow.webContents.send('report-created', { row, outPath, reference: context.refnr });
  return { outPath, screenshotCount: chosen.length, reference: context.refnr };
});

ipcMain.handle('report:delete', async () => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry');
  if (!state.outputFolder) throw new Error('No output folder selected');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const url = rowUrl(sheet, row);
  const outPath = reportPathFor(url);
  if (fs.existsSync(outPath)) {
    fs.unlinkSync(outPath);
    log(`Row ${row + 1}: deleted report ${outPath}`);
  }
  delete state.screenshotsByRow[row];
  return true;
});
