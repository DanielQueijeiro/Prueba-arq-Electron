const { contextBridge, ipcRenderer } = require('electron/renderer')

contextBridge.exposeInMainWorld('electronAPI', {
  openFileDialog: () => ipcRenderer.send('open-file-dialog'),
  onFileSelected: (callback) => ipcRenderer.on('file-selected', (event, filePath) => callback(filePath)),
  savePDF: (callback) => {
    ipcRenderer.send('save-pdf-dialog');
    ipcRenderer.once('save-pdf-path', (event, filePath) => callback(filePath));
  },
  writePDFFile: (filePath, arrayBuffer) => {
    const buffer = new Uint8Array(arrayBuffer);
    ipcRenderer.send('write-pdf-file', filePath, buffer);
  },
  showError: (message) => ipcRenderer.send('show-error', message)
})