const { app, BrowserWindow, BrowserView, ipcMain, dialog, shell, screen } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

let mainWindow;
let browserView;
let reviewWindow = null;

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

// Infringement articles offered when a report is generated. Fixed rather than
// read from a sheet: these are statutory references, not a working list.
const INBREUK_OPTIONS = [
  'Art. 4 §1 (illegale exploitatie)',
  'Art. 4 §2 (faciliteren / reclame)',
];
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
  inbreukColIdx: null, // column holding the infringement articles
  currentRow: null, // 0-based sheet row index of the entry currently on screen

  // What the address bar should show, which is not always what the view has
  // loaded: a failed page is replaced by a local error document.
  intendedUrl: '',
  showingError: false,

  // Background capture jobs. Screenshots live on disk; a job only holds
  // metadata and thumbnails.
  captureJobs: [],
  captureRunning: 0,
  activeReviewJobId: null,

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

  browserView = new BrowserView({ webPreferences: { backgroundColor: '#1b1b1d' } });
  mainWindow.setBrowserView(browserView);
  browserView.setBounds({ x: 0, y: 0, width: 0, height: 0 });
  // about:blank is white and stays on screen whenever there is no entry, which
  // is hard on the eyes next to the rest of the dark UI.
  browserView.setBackgroundColor('#1b1b1d');
  showIdlePage();

  const sendNav = (url) => {
    if (state.showingError) return; // keep the failed address on screen
    state.intendedUrl = url;
    mainWindow.webContents.send('nav-state', { url });
  };

  browserView.webContents.on('did-navigate', (e, url) => sendNav(url));
  browserView.webContents.on('did-navigate-in-page', (e, url) => sendNav(url));

  browserView.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith(ACTION_ORIGIN)) return;
    e.preventDefault();
    const showing = browserView.webContents.getURL();
    if (!state.showingError || !showing.startsWith('file://')) return;
    handleErrorPageAction(url);
  });

  browserView.webContents.on('did-start-loading', () => {
    mainWindow.webContents.send('load-state', { loading: true });
  });
  browserView.webContents.on('did-stop-loading', () => {
    mainWindow.webContents.send('load-state', { loading: false });
  });

  browserView.webContents.on('did-fail-load', (e, errorCode, errorDescription, validatedURL, isMainFrame) => {
    if (errorCode === -3) return; // ERR_ABORTED, usually just a superseded navigation
    if (!isMainFrame) return; // a failed image or iframe is not a failed page
    const url = validatedURL || state.intendedUrl || '(unknown URL)';
    log(`Failed to load ${url}: ${errorDescription} (${errorCode})`);
    showErrorPage(url, errorDescription, errorCode);
  });

  // Certificate problems are common on these sites and otherwise surface as a
  // bare ERR_CERT_* with no explanation in the view.
  browserView.webContents.on('certificate-error', (e, url, error) => {
    log(`Certificate problem on ${url}: ${error}`);
  });
}

// ---- idle page ---------------------------------------------------------

function idlePagePath() {
  return path.join(app.getPath('temp'), 'sitewizard-idle.html');
}

function showIdlePage() {
  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  * { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
  body { margin: 0; height: 100vh; display: flex; align-items: center; justify-content: center;
         background: #1b1b1d; color: #6f6d6a; }
  .idle { text-align: center; font-size: 13px; line-height: 1.7; }
  .idle strong { color: #9c9a96; font-weight: 600; display: block; margin-bottom: 4px; font-size: 14px; }
</style></head>
<body><div class="idle"><strong>No site loaded</strong>Load a workbook and apply the columns to begin.</div></body></html>`;
  try {
    fs.writeFileSync(idlePagePath(), html, 'utf8');
    state.showingError = false;
    state.intendedUrl = '';
    browserView.webContents.loadFile(idlePagePath()).catch(() => {});
  } catch {
    browserView.webContents.loadURL('about:blank');
  }
}

// ---- in-view error page ------------------------------------------------
// Rendered into the BrowserView itself so a failure is visible where the page
// would have been, not only in the activity log. Written to a temp file rather
// than a data: URL because Chromium blocks top-level data: navigations.

function errorPagePath() {
  return path.join(app.getPath('temp'), 'sitewizard-error.html');
}

// The error page needs buttons, but it lives in the same BrowserView that
// loads the sites under investigation, so it gets no preload and no IPC
// bridge. Instead its buttons are links to a host that can never resolve
// (.invalid is reserved for exactly this), and the navigation is intercepted
// before it goes anywhere. Two guards make sure only our own page can fire
// one: the view must currently be showing the error document, and the
// document doing the navigating must be that file.
const ACTION_ORIGIN = 'https://sitewizard.invalid';

function handleErrorPageAction(url) {
  const action = url.slice(ACTION_ORIGIN.length).replace(/^\/+/, '').split(/[?#]/)[0];

  if (action === 'retry') {
    if (state.intendedUrl) navigateBrowserView(state.intendedUrl);
    return;
  }
  if (action === 'open-external') {
    const target = state.intendedUrl;
    if (!target || !/^https?:\/\//i.test(target)) return;
    shell.openExternal(target);
    log('Opened in the default browser: ' + target);
    return;
  }
  if (action === 'no-access') {
    try {
      const info = applyStatus('No access');
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('entry-updated', info);
      }
    } catch (err) {
      log('Could not mark No access: ' + err.message);
    }
  }
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Plain-language causes for the codes these sites actually produce.
function explainLoadError(code) {
  const reasons = {
    '-2': 'The request failed. The server may be refusing connections.',
    '-6': 'The file or page could not be found.',
    '-7': 'The site took too long to respond and the request timed out.',
    '-21': 'The network changed while the page was loading.',
    '-102': 'The connection was refused — nothing is listening on that address.',
    '-105': 'The domain name could not be resolved. It may no longer exist.',
    '-106': 'The machine appears to be offline.',
    '-109': 'The host is unreachable.',
    '-118': 'The connection timed out before the site responded.',
    '-130': 'A proxy refused the connection.',
    '-137': 'The domain name could not be resolved.',
    '-200': 'The site presented a certificate that is not valid.',
    '-201': 'The certificate has expired or is not yet valid.',
    '-202': 'The certificate was issued for a different domain.',
    '-501': 'The site is not sending a usable response.',
  };
  return reasons[String(code)] || 'The page could not be displayed.';
}

function showErrorPage(url, description, code) {
  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  * { box-sizing: border-box; font-family: "Segoe UI", system-ui, sans-serif; }
  body { margin: 0; height: 100vh; display: flex; align-items: center; justify-content: center;
         background: #1b1b1d; color: #e9e7e4; padding: 32px; }
  .card { max-width: 560px; text-align: center; }
  .icon { font-size: 34px; color: #d9a441; margin-bottom: 14px; }
  h1 { font-size: 17px; font-weight: 600; margin: 0 0 10px; }
  p { font-size: 13.5px; line-height: 1.6; color: #9c9a96; margin: 0 0 14px; }
  .url { font-size: 12.5px; color: #4a9cbe; word-break: break-all; margin-bottom: 18px; }
  .actions { display: flex; gap: 8px; justify-content: center; margin-bottom: 20px; flex-wrap: wrap; }
  .btn { font-size: 13px; padding: 8px 14px; border-radius: 6px; text-decoration: none;
         border: 1px solid #38383f; background: #2c2c31; color: #e9e7e4; }
  .btn:hover { background: #303036; border-color: #46464e; }
  .btn.primary { background: #4a9cbe; border-color: #4a9cbe; color: #1b1b1d; font-weight: 600; }
  .btn.primary:hover { background: #5fb0d0; border-color: #5fb0d0; }
  .detail { font-family: Consolas, monospace; font-size: 11.5px; color: #6f6d6a;
            border-top: 1px solid #2e2e34; padding-top: 12px; }
</style></head>
<body><div class="card">
  <div class="icon">&#9888;</div>
  <h1>This site could not be opened</h1>
  <div class="url">${escapeHtml(url)}</div>
  <p>${escapeHtml(explainLoadError(code))}</p>
  <div class="actions">
    <a class="btn primary" href="${ACTION_ORIGIN}/open-external">Open in your browser</a>
    <a class="btn" href="${ACTION_ORIGIN}/no-access">Mark as No access</a>
    <a class="btn" href="${ACTION_ORIGIN}/retry">Try again</a>
  </div>
  <div class="detail">${escapeHtml(description)} (${escapeHtml(code)})</div>
</div></body></html>`;
  try {
    fs.writeFileSync(errorPagePath(), html, 'utf8');
    state.showingError = true;
    browserView.webContents.loadFile(errorPagePath()).catch(() => {});
  } catch (err) {
    log('Could not display the error page: ' + err.message);
  }
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
  state.showingError = false;
  state.intendedUrl = target;
  if (mainWindow) mainWindow.webContents.send('nav-state', { url: target });
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

// Builds the gevolg-related tag values for the template. Three shapes are
// offered so a template can show the measures as a list, as one line, or as a
// full tick-box checklist mirroring the paper form:
//   {#gevolg}{.}{/gevolg}                  -> only the selected measures
//   {gevolgText}                           -> selected measures on one line
//   {#gevolgAll}{mark} {label}{/gevolgAll} -> every option, ticked or not
function gevolgTagValues(selection) {
  const chosen = (selection || []).map((s) => (s.text ? `${s.label} ${s.text}`.trim() : s.label));
  const chosenLabels = new Set((selection || []).map((s) => s.label));
  const all = state.gevolgOptions.map((opt) => {
    const hit = (selection || []).find((s) => s.label === opt.label);
    const checked = chosenLabels.has(opt.label);
    return {
      label: hit && hit.text ? `${opt.label} ${hit.text}`.trim() : opt.label,
      checked,
      mark: checked ? '☒' : '☐',
    };
  });
  return { gevolg: chosen, gevolgText: chosen.join(GEVOLG_SEPARATOR), gevolgAll: all };
}

// Same three shapes as gevolg, for the INBREUK section:
//   {#inbreukAll}{mark} {label}{/inbreukAll}, {#inbreuk}{.}{/inbreuk}, {inbreukText}
function inbreukTagValues(selection) {
  const chosen = (selection || []).slice();
  const set = new Set(chosen);
  const all = INBREUK_OPTIONS.map((label) => ({
    label,
    checked: set.has(label),
    mark: set.has(label) ? '☒' : '☐',
  }));
  return { inbreuk: chosen, inbreukText: chosen.join(GEVOLG_SEPARATOR), inbreukAll: all };
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
      // Each iteration carries the page it came from so the template can
      // print the URL above the image.
      screenshots: screenshots.map((shot, i) => ({
        image: i + 1,
        pageUrl: shot.pageUrl || context.site,
      })),
      ...gevolgTagValues(context.gevolg),
      ...inbreukTagValues(context.inbreuk),
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

// Whatever articles are already recorded for a row, as a plain list.
function inbreukForRow(row) {
  if (state.inbreukColIdx == null) return [];
  const sheet = state.workbook.Sheets[state.sheetName];
  const v = cellValue(sheet, row, state.inbreukColIdx);
  if (v === undefined || !String(v).trim()) return [];
  return String(v)
    .split('|')
    .map((x) => x.trim())
    .filter(Boolean);
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

async function crawlAndCapture(startUrl, progress, opts) {
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
  // Full-resolution PNGs go straight to disk. A single site is several
  // megabytes, and with a queue there can be many in flight, so only the
  // small thumbnails are kept in memory.
  const collect = (shot) => {
    const id = String(nextId++);
    const file = path.join(opts.dir, `shot-${id}.png`);
    try {
      fs.writeFileSync(file, shot.buffer);
    } catch (err) {
      log(`Could not store a screenshot: ${err.message}`);
      return;
    }
    shots.push({
      id,
      file,
      pageUrl: shot.pageUrl,
      slice: shot.slice,
      sliceCount: shot.sliceCount,
      thumb: shot.thumb,
      width: shot.width,
      height: shot.height,
    });
    progress({ phase: 'capturing', shots: shots.length, page: shot.pageUrl });
  };

  try {
    progress({ phase: 'loading', page: root, index: 1 });
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
      if (opts.isCancelled()) break;
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

// ---- Capture queue -----------------------------------------------------
// Captures run in the background so triage never has to wait on a crawl.
// Concurrency is capped because each job is a full Chromium window.

const MAX_CONCURRENT_CAPTURES = 2;
let nextJobId = 1;

function jobRoot() {
  return path.join(app.getPath('temp'), 'sitewizard-captures');
}

function publicJob(job) {
  return {
    id: job.id,
    row: job.row,
    url: job.url,
    status: job.status,
    message: job.message || '',
    shotCount: job.shots ? job.shots.length : 0,
    error: job.error || '',
  };
}

function broadcastJobs() {
  const payload = state.captureJobs.map(publicJob);
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('capture-jobs', payload);
  }
  // The review window shows a tab per finished capture, so it needs these too.
  if (reviewWindow && !reviewWindow.isDestroyed()) {
    reviewWindow.webContents.send('capture-jobs', payload);
  }
}

function describeProgress(p) {
  if (p.phase === 'loading') {
    return p.total ? `Loading page ${p.index} of ${p.total}` : 'Loading homepage';
  }
  if (p.phase === 'queued') return `Found ${p.total} page(s)`;
  if (p.phase === 'capturing') return `${p.shots} screenshot(s)`;
  if (p.phase === 'warning') return p.message;
  return '';
}

function discardJobFiles(job) {
  try {
    fs.rmSync(path.join(jobRoot(), job.id), { recursive: true, force: true });
  } catch {
    /* a leftover temp folder is harmless */
  }
}

async function runJob(job) {
  job.status = 'running';
  job.message = 'Starting…';
  broadcastJobs();

  const dir = path.join(jobRoot(), job.id);
  try {
    fs.mkdirSync(dir, { recursive: true });
    job.shots = await crawlAndCapture(
      job.url,
      (p) => {
        const text = describeProgress(p);
        if (text) {
          job.message = text;
          broadcastJobs();
        }
      },
      { dir, isCancelled: () => job.cancelled }
    );
    if (job.cancelled) {
      job.status = 'cancelled';
      job.message = 'Cancelled';
      discardJobFiles(job);
    } else if (job.shots.length === 0) {
      job.status = 'failed';
      job.error = 'Nothing could be captured';
      job.message = job.error;
    } else {
      job.status = 'done';
      job.message = `${job.shots.length} screenshot(s) ready`;
      log(`Row ${job.row + 1}: capture finished — ${job.shots.length} screenshot(s) for ${job.url}`);
    }
  } catch (err) {
    job.status = 'failed';
    job.error = err.message;
    job.message = 'Failed: ' + err.message;
    log(`Row ${job.row + 1}: capture failed — ${err.message}`);
  } finally {
    state.captureRunning -= 1;
    broadcastJobs();
    pumpQueue();
  }
}

function pumpQueue() {
  while (state.captureRunning < MAX_CONCURRENT_CAPTURES) {
    const next = state.captureJobs.find((j) => j.status === 'queued' && !j.cancelled);
    if (!next) return;
    state.captureRunning += 1;
    runJob(next);
  }
}

function queueCapture(row, url) {
  const existing = state.captureJobs.find(
    (j) => j.row === row && (j.status === 'queued' || j.status === 'running')
  );
  if (existing) return existing;
  const job = {
    id: String(nextJobId++),
    row,
    url,
    status: 'queued',
    message: 'Waiting…',
    shots: [],
    cancelled: false,
  };
  state.captureJobs.push(job);
  broadcastJobs();
  pumpQueue();
  return job;
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
  if (isValid) {
    navigateBrowserView(rowUrl(state.workbook.Sheets[state.sheetName], row));
  } else {
    showIdlePage();
  }
  return buildEntryInfo(state.currentRow);
}

// ---- IPC: file dialogs / setup -------------------------------------------

// The three setup steps are factored out of their IPC handlers so that
// restoring the last session can run exactly the same code path.

function openWorkbook(filePath) {
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
  saveSettings({ excelPath: filePath });
  return { filePath, sheetNames: workbook.SheetNames };
}

function selectSheet(sheetName) {
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
  saveSettings({ sheetName });
  return { header };
}

function applyColumns(cfg) {
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
  let inbreukIdx = null;
  if (cfg.inbreukCol && cfg.inbreukCol.trim()) {
    inbreukIdx = resolveColumn(sheet, state.range, cfg.inbreukCol, true);
  }
  state.urlColIdx = urlIdx;
  state.statusColIdx = statusIdx;
  state.commentColIdx = commentIdx;
  state.gevolgColIdx = gevolgIdx;
  state.refColIdx = refIdx;
  state.inbreukColIdx = inbreukIdx;
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
    inbreukCol: cfg.inbreukCol || '',
  });
  return moveToRow(startRow);
}

ipcMain.handle('dialog:openExcel', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return openWorkbook(result.filePaths[0]);
});

ipcMain.handle('excel:selectSheet', async (event, sheetName) => selectSheet(sheetName));

ipcMain.handle('excel:setColumns', async (event, cfg) => applyColumns(cfg));

// Reopens last session's workbook, sheet and columns on startup. Anything
// that has moved or been renamed since is reported and skipped rather than
// treated as an error, so the app still opens cleanly.
ipcMain.handle('excel:restoreSession', async () => {
  const saved = loadSettings();
  if (!saved.excelPath) return null;
  if (!fs.existsSync(saved.excelPath)) {
    log(`Last workbook is no longer at ${saved.excelPath} — load it again.`);
    return null;
  }
  try {
    const opened = openWorkbook(saved.excelPath);
    const sheetName = opened.sheetNames.includes(saved.sheetName)
      ? saved.sheetName
      : opened.sheetNames[0];
    const { header } = selectSheet(sheetName);
    log(`Reopened ${path.basename(saved.excelPath)} (${sheetName}).`);

    let entry = null;
    if (saved.urlCol && saved.statusCol) {
      try {
        entry = applyColumns({
          urlCol: saved.urlCol,
          statusCol: saved.statusCol,
          commentCol: saved.commentCol || '',
          gevolgCol: saved.gevolgCol || '',
          gevolgSheet: saved.gevolgSheet || '',
          refCol: saved.refCol || '',
          inbreukCol: saved.inbreukCol || '',
        });
      } catch (err) {
        log('Could not reapply the saved columns: ' + err.message);
      }
    }
    return { ...opened, sheetName, header, entry };
  } catch (err) {
    log('Could not reopen the last workbook: ' + err.message);
    return null;
  }
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
  // After a failure the view holds the local error document, so reloading it
  // would just redisplay the error.
  if (state.showingError && state.intendedUrl) {
    navigateBrowserView(state.intendedUrl);
    return;
  }
  browserView.webContents.reload();
});

ipcMain.handle('browserview:openExternal', async () => {
  const url = state.intendedUrl || browserView.webContents.getURL();
  if (!url || !/^https?:\/\//i.test(url)) throw new Error('No site to open');
  await shell.openExternal(url);
  log('Opened in the default browser: ' + url);
  return url;
});

// ---- IPC: workflow actions ---------------------------------------------

function applyStatus(status) {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const oldValue = cellValue(sheet, row, state.statusColIdx);
  const oldStr = oldValue === undefined ? '' : String(oldValue);
  setCell(sheet, row, state.statusColIdx, status);
  saveWorkbook();
  recordAction({ type: 'status', row, colIdx: state.statusColIdx, oldValue: oldStr, newValue: status });
  log(`Row ${row + 1}: marked "${status}" and saved.`);

  // Staying put lets an agent mark a status and still add a comment or queue a
  // capture for the same entry before moving on.
  const saved = loadSettings();
  if (saved.autoAdvance === false) return buildEntryInfo(row);

  const nextRow = findNextRow(sheet, state.range, state.urlColIdx, state.statusColIdx, row + 1);
  return moveToRow(nextRow);
}

ipcMain.handle('entries:markStatus', async (event, status) => applyStatus(status));

ipcMain.handle('settings:save', async (event, partial) => {
  saveSettings(partial || {});
  return loadSettings();
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

// ---- IPC: capture queue and review window -----------------------------
// Generate Report queues a background capture and returns immediately. When a
// job finishes, its review window is where the screenshots are chosen and the
// report is finally written; nothing touches the sheet before that.

ipcMain.handle('capture:jobs', async () => state.captureJobs.map(publicJob));

ipcMain.handle('capture:queue', async () => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry to report on');
  if (!state.templatePath) throw new Error('Select a report template first');
  if (!state.outputFolder) throw new Error('Select an output folder first');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const job = queueCapture(row, rowUrl(sheet, row));
  log('Row ' + (row + 1) + ': capture queued for ' + job.url);
  return publicJob(job);
});

ipcMain.handle('capture:cancelJob', async (event, id) => {
  const job = state.captureJobs.find((j) => j.id === id);
  if (!job) return false;
  job.cancelled = true;
  if (job.status === 'queued') {
    job.status = 'cancelled';
    job.message = 'Cancelled';
    broadcastJobs();
  }
  return true;
});

ipcMain.handle('capture:dismissJob', async (event, id) => {
  const job = state.captureJobs.find((j) => j.id === id);
  if (job) {
    job.cancelled = true;
    discardJobFiles(job);
  }
  state.captureJobs = state.captureJobs.filter((j) => j.id !== id);
  broadcastJobs();
  return true;
});

// Puts the two windows side by side across the work area rather than leaving
// the review window floating over the browser.
function dockReviewWindow() {
  if (!mainWindow || mainWindow.isDestroyed() || !reviewWindow || reviewWindow.isDestroyed()) return;
  try {
    const display = screen.getDisplayMatching(mainWindow.getBounds());
    const wa = display.workArea;
    const reviewWidth = Math.max(560, Math.min(760, Math.round(wa.width * 0.42)));
    const mainWidth = wa.width - reviewWidth;
    if (mainWidth < 640) return; // too narrow to split usefully; leave both alone
    if (mainWindow.isMaximized()) mainWindow.unmaximize();
    mainWindow.setBounds({ x: wa.x, y: wa.y, width: mainWidth, height: wa.height });
    reviewWindow.setBounds({ x: wa.x + mainWidth, y: wa.y, width: reviewWidth, height: wa.height });
  } catch (err) {
    log('Could not dock the report window: ' + err.message);
  }
}

ipcMain.handle('capture:review', async (event, id) => {
  const job = state.captureJobs.find((j) => j.id === id);
  if (!job) throw new Error('That capture is no longer available');
  if (job.status !== 'done') throw new Error('That capture has not finished yet');
  state.activeReviewJobId = id;

  if (reviewWindow && !reviewWindow.isDestroyed()) {
    reviewWindow.webContents.send('review-switch', id);
    reviewWindow.focus();
    return true;
  }
  reviewWindow = new BrowserWindow({
    width: 1040,
    height: 820,
    minWidth: 720,
    minHeight: 560,
    parent: mainWindow,
    title: 'Generate Report',
    backgroundColor: '#1b1b1d',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  reviewWindow.setMenu(null);
  reviewWindow.loadFile('capture.html');
  reviewWindow.once('ready-to-show', dockReviewWindow);
  reviewWindow.on('closed', () => {
    reviewWindow = null;
    state.activeReviewJobId = null;
  });
  return true;
});

function activeJob() {
  const job = state.captureJobs.find((j) => j.id === state.activeReviewJobId);
  if (!job) throw new Error('No capture is open for review');
  return job;
}

ipcMain.handle('capture:context', async () => {
  const job = activeJob();
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = job.row;
  const ref = referenceForRow(row);
  const saved = loadSettings();

  const stored = state.gevolgColIdx != null ? cellValue(sheet, row, state.gevolgColIdx) : undefined;
  const gevolg =
    stored !== undefined && String(stored).trim() !== ''
      ? parseGevolgCell(stored)
      : state.gevolgSticky.slice();

  const raw = cellValue(sheet, row, state.urlColIdx);

  return {
    jobId: job.id,
    // One tab per finished capture, so they can be worked through as they land.
    tabs: state.captureJobs
      .filter((j) => j.status === 'done')
      .map((j) => ({ id: j.id, row: j.row, url: j.url, shotCount: j.shots.length })),
    row,
    url: job.url,
    rawUrl: raw === undefined ? '' : String(raw),
    reference: ref.reference,
    referenceIsNew: ref.isNew,
    gevolgEnabled: state.gevolgColIdx != null,
    gevolgOptions: state.gevolgOptions,
    gevolg,
    agentNames: saved.agentNames && saved.agentNames.length ? saved.agentNames : DEFAULT_AGENT_NAMES,
    agentName: saved.agentName || '',
    bron: bronForRow(row),
    inbreukEnabled: state.inbreukColIdx != null,
    inbreukOptions: INBREUK_OPTIONS,
    inbreuk: inbreukForRow(row),
    shots: job.shots.map((s) => ({
      id: s.id,
      pageUrl: s.pageUrl,
      slice: s.slice,
      sliceCount: s.sliceCount,
      thumb: s.thumb,
      width: s.width,
      height: s.height,
    })),
  };
});

// Full-resolution image for the fullscreen viewer, read back off disk.
ipcMain.handle('capture:image', async (event, id) => {
  const job = activeJob();
  const shot = job.shots.find((s) => s.id === id);
  if (!shot) return null;
  try {
    return 'data:image/png;base64,' + fs.readFileSync(shot.file).toString('base64');
  } catch (err) {
    log('Could not read a captured screenshot: ' + err.message);
    return null;
  }
});

ipcMain.handle('capture:selectJob', async (event, id) => {
  const job = state.captureJobs.find((j) => j.id === id && j.status === 'done');
  if (!job) throw new Error('That capture is no longer available');
  state.activeReviewJobId = id;
  return true;
});

ipcMain.handle('capture:cancel', async () => {
  if (reviewWindow && !reviewWindow.isDestroyed()) reviewWindow.close();
  return true;
});

// The only place a report is written and the sheet updated.
ipcMain.handle('capture:generate', async (event, payload) => {
  const job = activeJob();
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = job.row;
  const url = job.url;

  const chosen = (payload.selectedIds || [])
    .map((id) => job.shots.find((s) => s.id === id))
    .filter(Boolean)
    .map((s) => ({
      buffer: fs.readFileSync(s.file),
      width: s.width,
      height: s.height,
      pageUrl: s.pageUrl,
    }));
  if (chosen.length === 0) throw new Error('Select at least one screenshot');

  const context = {
    site: url,
    refnr: payload.reference || '',
    datum: payload.datum || new Date().toLocaleDateString('nl-BE'),
    controleur: payload.agentName || '',
    bron: payload.bron || '',
    gevolg: payload.gevolg || [],
    inbreuk: payload.inbreuk || [],
  };

  const outBuffer = writeReport(context, chosen);
  const outPath = reportPathFor(url);
  fs.writeFileSync(outPath, outBuffer);

  // Sheet writes happen only now, never while the capture is running.
  const wrote = [];
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
  if (state.inbreukColIdx != null) {
    const value = (payload.inbreuk || []).join(GEVOLG_SEPARATOR);
    const old = cellValue(sheet, row, state.inbreukColIdx);
    const oldStr = old === undefined ? '' : String(old);
    if (oldStr !== value) {
      setCell(sheet, row, state.inbreukColIdx, value);
      recordAction({ type: 'inbreuk', row, colIdx: state.inbreukColIdx, oldValue: oldStr, newValue: value });
      wrote.push('infractions');
    }
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
  log('Row ' + (row + 1) + ': report ' + context.refnr + ' saved with ' + chosen.length + ' screenshot(s): ' + outPath);
  if (wrote.length > 0) log('Row ' + (row + 1) + ': wrote ' + wrote.join(' and ') + ' to the sheet.');

  // The job has served its purpose; drop it and its files.
  discardJobFiles(job);
  state.captureJobs = state.captureJobs.filter((j) => j.id !== job.id);
  state.activeReviewJobId = null;
  broadcastJobs();

  // Keep the window open while there are other captures waiting to be dealt with.
  const next = state.captureJobs.find((j) => j.status === 'done');
  if (next) {
    state.activeReviewJobId = next.id;
  } else if (reviewWindow && !reviewWindow.isDestroyed()) {
    reviewWindow.close();
  }

  if (mainWindow) mainWindow.webContents.send('report-created', { row: row, outPath: outPath, reference: context.refnr });
  return {
    outPath: outPath,
    screenshotCount: chosen.length,
    reference: context.refnr,
    next: next ? next.id : null,
  };
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
