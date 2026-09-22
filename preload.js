const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openExcel: () => ipcRenderer.invoke('dialog:openExcel'),
  selectSheet: (sheetName) => ipcRenderer.invoke('excel:selectSheet', sheetName),
  setColumns: (cfg) => ipcRenderer.invoke('excel:setColumns', cfg),

  openTemplate: () => ipcRenderer.invoke('dialog:openTemplate'),
  selectOutputFolder: () => ipcRenderer.invoke('dialog:selectOutputFolder'),
  loadSettings: () => ipcRenderer.invoke('settings:load'),

  setBrowserViewBounds: (rect) => ipcRenderer.send('browserview:bounds', rect),
  navBack: () => ipcRenderer.invoke('browserview:back'),
  navForward: () => ipcRenderer.invoke('browserview:forward'),
  navReload: () => ipcRenderer.invoke('browserview:reload'),

  markStatus: (status) => ipcRenderer.invoke('entries:markStatus', status),
  skip: () => ipcRenderer.invoke('entries:skip'),
  previous: () => ipcRenderer.invoke('entries:previous'),
  undo: () => ipcRenderer.invoke('entries:undo'),
  redo: () => ipcRenderer.invoke('entries:redo'),
  saveComment: (value) => ipcRenderer.invoke('entries:saveComment', value),
  saveGevolg: (selection) => ipcRenderer.invoke('entries:saveGevolg', selection),

  generateReport: () => ipcRenderer.invoke('report:generate'),
  deleteReport: () => ipcRenderer.invoke('report:delete'),

  onLog: (cb) => ipcRenderer.on('log', (event, msg) => cb(msg)),
  onNav: (cb) => ipcRenderer.on('nav-state', (event, navState) => cb(navState)),
});
