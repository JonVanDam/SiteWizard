const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openExcel: () => ipcRenderer.invoke('dialog:openExcel'),
  selectSheet: (sheetName) => ipcRenderer.invoke('excel:selectSheet', sheetName),
  setColumns: (cfg) => ipcRenderer.invoke('excel:setColumns', cfg),
  restoreSession: () => ipcRenderer.invoke('excel:restoreSession'),

  openTemplate: () => ipcRenderer.invoke('dialog:openTemplate'),
  selectOutputFolder: () => ipcRenderer.invoke('dialog:selectOutputFolder'),
  loadSettings: () => ipcRenderer.invoke('settings:load'),
  saveSettings: (partial) => ipcRenderer.invoke('settings:save', partial),

  setBrowserViewBounds: (rect) => ipcRenderer.send('browserview:bounds', rect),
  navBack: () => ipcRenderer.invoke('browserview:back'),
  navForward: () => ipcRenderer.invoke('browserview:forward'),
  navReload: () => ipcRenderer.invoke('browserview:reload'),
  openExternal: () => ipcRenderer.invoke('browserview:openExternal'),

  markStatus: (status) => ipcRenderer.invoke('entries:markStatus', status),
  skip: () => ipcRenderer.invoke('entries:skip'),
  previous: () => ipcRenderer.invoke('entries:previous'),
  undo: () => ipcRenderer.invoke('entries:undo'),
  redo: () => ipcRenderer.invoke('entries:redo'),
  saveComment: (value) => ipcRenderer.invoke('entries:saveComment', value),

  // Captures run in the background; the review window opens when one is done.
  captureJobs: () => ipcRenderer.invoke('capture:jobs'),
  queueCapture: () => ipcRenderer.invoke('capture:queue'),
  cancelJob: (id) => ipcRenderer.invoke('capture:cancelJob', id),
  dismissJob: (id) => ipcRenderer.invoke('capture:dismissJob', id),
  reviewJob: (id) => ipcRenderer.invoke('capture:review', id),
  onCaptureJobs: (cb) => ipcRenderer.on('capture-jobs', (event, jobs) => cb(jobs)),

  captureContext: () => ipcRenderer.invoke('capture:context'),
  selectJob: (id) => ipcRenderer.invoke('capture:selectJob', id),
  onReviewSwitch: (cb) => ipcRenderer.on('review-switch', (event, id) => cb(id)),
  captureImage: (id) => ipcRenderer.invoke('capture:image', id),
  captureCancel: () => ipcRenderer.invoke('capture:cancel'),
  dismissCurrent: () => ipcRenderer.invoke('capture:dismissCurrent'),
  captureGenerate: (payload) => ipcRenderer.invoke('capture:generate', payload),

  deleteReport: () => ipcRenderer.invoke('report:delete'),

  onLog: (cb) => ipcRenderer.on('log', (event, msg) => cb(msg)),
  onNav: (cb) => ipcRenderer.on('nav-state', (event, navState) => cb(navState)),
  onLoadState: (cb) => ipcRenderer.on('load-state', (event, s) => cb(s)),
  onReportCreated: (cb) => ipcRenderer.on('report-created', (event, info) => cb(info)),
  onEntryUpdated: (cb) => ipcRenderer.on('entry-updated', (event, info) => cb(info)),
});
