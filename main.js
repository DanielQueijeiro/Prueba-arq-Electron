const { app, BrowserWindow, ipcMain, dialog } = require('electron/main')
const path = require('node:path')
const XLSX = require('xlsx')
const fs = require('fs')

function createWindow () {
  const mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  })
 

  ipcMain.on('open-file-dialog', () => {
    dialog.showOpenDialog(mainWindow, {
      properties: ['openFile'],
      filters: [
        {name: 'Archivos de Excel', extensions: ['xlsx']}
      ],
    }).then(result => {
      if (!result.canceled && result.filePaths.length > 0) {
        const filePath = result.filePaths[0];
        const fileName = path.basename(filePath);
        const data = readExcelFile(filePath);

        mainWindow.webContents.send('file-selected', {
          fileName: fileName,
          data: data,
        });
        console.log('Data from Excel:', data);
      }
    }).catch(err => {
      console.log(err)
    })
  })

  ipcMain.on('save-pdf-dialog', async (event) => {
    try {
      const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: 'reporte.pdf',
        filters: [{ name: 'PDF', extensions: ['pdf'] }]
      });
      
      if (!result.canceled) {
        event.reply('save-pdf-path', result.filePath);
      } else {
        event.reply('save-pdf-path', null);
      }
    } catch (error) {
      console.error('Error al abrir el diálogo de guardar:', error);
      event.reply('save-pdf-path', null);
    }
  });
  
  ipcMain.on('write-pdf-file', (event, filePath, bufferData) => {
    try {
      // Convert the received Uint8Array to Buffer
      const buffer = Buffer.from(bufferData);
      fs.writeFileSync(filePath, buffer);
      dialog.showMessageBox(mainWindow, {
        type: 'info',
        title: 'PDF Guardado',
        message: 'El PDF se ha guardado correctamente',
        buttons: ['OK']
      });
    } catch (error) {
      console.error('Error al guardar el PDF:', error);
      dialog.showErrorBox('Error', `No se pudo guardar el PDF: ${error.message}`);
    }
  });
  
  ipcMain.on('show-error', (event, message) => {
    dialog.showErrorBox('Error', message);
  });

  mainWindow.loadFile('index.html')
}

app.whenReady().then(() => {
  createWindow()

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})

function readExcelFile(filePath) {
  try {
      // Read the workbook
      console.log('Reading Excel file...', filePath);
      const workbook = XLSX.readFile(filePath);
      
      // Convert first sheet to 2D array
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      return data;
  } catch (err) {
      console.error('Error reading Excel file:', err);
  }
}