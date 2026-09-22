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
  currentRow: null, // 0-based sheet row index of the entry currently on screen

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
    XLSX.writeFile(state.workbook, state.excelPath);
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

function pngDimensions(buffer) {
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}

// Loads (and caches) the screenshots captured so far for a row. If none are
// in memory yet but a report already exists on disk for this URL, its
// embedded images are pulled back out so "Add Screenshot" can append to them
// even after an app restart.
function getScreenshotsForRow(row, url) {
  if (!state.screenshotsByRow[row]) {
    const list = [];
    if (state.outputFolder) {
      const reportPath = reportPathFor(url);
      if (fs.existsSync(reportPath)) {
        try {
          const content = fs.readFileSync(reportPath, 'binary');
          const zip = new PizZip(content);
          const mediaNames = Object.keys(zip.files)
            .filter((name) => /^word\/media\/image\d+\.png$/i.test(name))
            .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));
          for (const name of mediaNames) {
            const buffer = zip.files[name].asNodeBuffer();
            const { width, height } = pngDimensions(buffer);
            list.push({ buffer, width, height });
          }
        } catch (err) {
          log('Could not read images from existing report: ' + err.message);
        }
      }
    }
    state.screenshotsByRow[row] = list;
  }
  return state.screenshotsByRow[row];
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

function writeReport(url, screenshots, selection) {
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
      site: url,
      screenshots: screenshots.map((_, i) => i + 1),
      ...gevolgTagValues(selection),
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

// ---- Entry navigation ------------------------------------------------

function buildEntryInfo(row) {
  const canUndo = state.actionPos >= 0;
  const canRedo = state.actionPos < state.actionLog.length - 1;
  const canPrevious = state.historyPos > 0;

  if (row == null || row === -1) {
    return { done: true, canUndo, canRedo, canPrevious };
  }
  const sheet = state.workbook.Sheets[state.sheetName];
  const url = String(cellValue(sheet, row, state.urlColIdx));
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
  navigateBrowserView(isValid ? String(cellValue(state.workbook.Sheets[state.sheetName], row, state.urlColIdx)) : 'about:blank');
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
  state.urlColIdx = urlIdx;
  state.statusColIdx = statusIdx;
  state.commentColIdx = commentIdx;
  state.gevolgColIdx = gevolgIdx;
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
  XLSX.writeFile(state.workbook, state.excelPath);
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
  navigateBrowserView(String(cellValue(sheet, row, state.urlColIdx)));
  log(`Row ${row + 1}: viewing previous entry.`);
  return buildEntryInfo(row);
});

ipcMain.handle('entries:undo', async () => {
  if (state.actionPos < 0) throw new Error('Nothing to undo');
  const action = state.actionLog[state.actionPos];
  const sheet = state.workbook.Sheets[state.sheetName];
  setCell(sheet, action.row, action.colIdx, action.oldValue);
  XLSX.writeFile(state.workbook, state.excelPath);
  state.actionPos -= 1;
  state.currentRow = action.row;
  navigateBrowserView(String(cellValue(sheet, action.row, state.urlColIdx)));
  log(`Row ${action.row + 1}: undid ${action.type} change.`);
  return buildEntryInfo(action.row);
});

ipcMain.handle('entries:redo', async () => {
  if (state.actionPos >= state.actionLog.length - 1) throw new Error('Nothing to redo');
  state.actionPos += 1;
  const action = state.actionLog[state.actionPos];
  const sheet = state.workbook.Sheets[state.sheetName];
  setCell(sheet, action.row, action.colIdx, action.newValue);
  XLSX.writeFile(state.workbook, state.excelPath);
  state.currentRow = action.row;
  navigateBrowserView(String(cellValue(sheet, action.row, state.urlColIdx)));
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
  XLSX.writeFile(state.workbook, state.excelPath);
  recordAction({ type: 'comment', row, colIdx: state.commentColIdx, oldValue: oldStr, newValue: value });
  log(`Row ${row + 1}: comment saved.`);
  return true;
});

ipcMain.handle('entries:saveGevolg', async (event, selection) => {
  // The selection is remembered even when there's nothing to write to, so the
  // panel still carries over between entries if no column is configured.
  state.gevolgSticky = (selection || []).map((s) => ({ label: s.label, text: s.text || '' }));
  if (state.currentRow == null || state.currentRow === -1) return false;
  if (state.gevolgColIdx == null) return false;

  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const value = formatGevolg(state.gevolgSticky);
  const oldValue = cellValue(sheet, row, state.gevolgColIdx);
  const oldStr = oldValue === undefined ? '' : String(oldValue);
  if (oldStr === value) return false;

  setCell(sheet, row, state.gevolgColIdx, value);
  XLSX.writeFile(state.workbook, state.excelPath);
  recordAction({ type: 'gevolg', row, colIdx: state.gevolgColIdx, oldValue: oldStr, newValue: value });
  log(`Row ${row + 1}: gevolg saved (${state.gevolgSticky.length} measure(s)).`);
  return true;
});

ipcMain.handle('report:generate', async () => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry to report on');
  if (!state.templatePath) throw new Error('Select a report template first');
  if (!state.outputFolder) throw new Error('Select an output folder first');

  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const url = String(cellValue(sheet, row, state.urlColIdx));

  const image = await browserView.webContents.capturePage();
  const size = image.getSize();
  const screenshots = getScreenshotsForRow(row, url);
  screenshots.push({ buffer: image.toPNG(), width: size.width, height: size.height });

  // Whatever the sheet holds for this row, falling back to the selection
  // carried over from the previous entry if this one hasn't been saved yet.
  const stored = state.gevolgColIdx != null ? cellValue(sheet, row, state.gevolgColIdx) : undefined;
  const selection =
    stored !== undefined && String(stored).trim() !== ''
      ? parseGevolgCell(stored)
      : state.gevolgSticky;

  const outBuffer = writeReport(url, screenshots, selection);
  const outPath = reportPathFor(url);
  fs.writeFileSync(outPath, outBuffer);
  log(`Row ${row + 1}: report saved with ${screenshots.length} screenshot(s): ${outPath}`);
  return { outPath, screenshotCount: screenshots.length };
});

ipcMain.handle('report:delete', async () => {
  if (state.currentRow == null || state.currentRow === -1) throw new Error('No current entry');
  if (!state.outputFolder) throw new Error('No output folder selected');
  const sheet = state.workbook.Sheets[state.sheetName];
  const row = state.currentRow;
  const url = String(cellValue(sheet, row, state.urlColIdx));
  const outPath = reportPathFor(url);
  if (fs.existsSync(outPath)) {
    fs.unlinkSync(outPath);
    log(`Row ${row + 1}: deleted report ${outPath}`);
  }
  delete state.screenshotsByRow[row];
  return true;
});
