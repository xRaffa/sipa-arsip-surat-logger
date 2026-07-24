'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // Database operations
  getLetters: () => ipcRenderer.invoke('db:getLetters'),
  addLetter: (table, entry) => ipcRenderer.invoke('db:addLetter', table, entry),
  deleteLetters: (table, ids) => ipcRenderer.invoke('db:deleteLetters', table, ids),

  // File operations
  uploadFile: (srcPath, originalName) => ipcRenderer.invoke('file:upload', srcPath, originalName),
  openFile: (filePath) => ipcRenderer.invoke('file:open', filePath),
  openFileDialog: () => ipcRenderer.invoke('dialog:openFile'),

  // App paths (for Settings display)
  getLocalPaths: () => ipcRenderer.invoke('app:getPaths'),
});
