const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('presentApi', {
  deckGetCurrent: () => ipcRenderer.invoke('deck:getCurrent'),
  deckCreate: () => ipcRenderer.invoke('deck:create'),
  deckSave: (presentation) => ipcRenderer.invoke('deck:save', presentation),
  deckImport: () => ipcRenderer.invoke('deck:import'),
  deckExport: (deckId) => ipcRenderer.invoke('deck:export', deckId),
  pickSlideFiles: () => ipcRenderer.invoke('deck:pickSlideFiles'),
  pickSlideDirectory: () => ipcRenderer.invoke('deck:pickSlideDirectory'),
  importSlides: (params) => ipcRenderer.invoke('deck:importSlides', params),
  resolveSlideUrl: (params) => ipcRenderer.invoke('deck:resolveSlideUrl', params),
  resolveSlideDataUrl: (params) => ipcRenderer.invoke('deck:resolveSlideDataUrl', params),
  pickPptxFile: () => ipcRenderer.invoke('deck:pickPptxFile'),
  importPptx: (params) => ipcRenderer.invoke('deck:importPptx', params),
  startPresentation: (params) => ipcRenderer.invoke('presentation:start', params),
  stopPresentation: () => ipcRenderer.invoke('presentation:stop'),
  getDisplays: () => ipcRenderer.invoke('system:getDisplays'),
  openExternal: (url) => ipcRenderer.invoke('system:openExternal', url),
});
