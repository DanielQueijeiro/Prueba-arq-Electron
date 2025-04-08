/**
 * Módulo principal de la aplicación Electron
 * Maneja la creación de ventanas, la comunicación con el sistema de archivos
 * y la interacción con los procesos de renderizado
 */
const { app, BrowserWindow, ipcMain, dialog } = require('electron/main');
const path = require('node:path');
const XLSX = require('xlsx');
const fs = require('fs');

/**
 * Crea la ventana principal de la aplicación
 */
function crearVentanaPrincipal() {
  // Configuración de la ventana principal
  const ventanaPrincipal = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });
 
  // Configurar manejadores de eventos IPC
  configurarEventosIPC(ventanaPrincipal);

  // Cargar el archivo HTML de la aplicación
  ventanaPrincipal.loadFile('index.html');
}

/**
 * Configura los manejadores de eventos para la comunicación entre procesos
 */
function configurarEventosIPC(ventana) {
  // Evento para abrir diálogo de selección de archivo Excel
  ipcMain.on('abrir-dialogo-archivo', () => {
    manejarSeleccionArchivo(ventana);
  });

  // Evento para abrir diálogo de guardado de PDF
  ipcMain.on('dialogo-guardar-pdf', async (event) => {
    manejarGuardadoPDF(event, ventana);
  });
  
  // Evento para escribir el archivo PDF en el sistema
  ipcMain.on('escribir-archivo-pdf', (event, rutaArchivo, datoBuffer) => {
    guardarArchivoPDF(rutaArchivo, datoBuffer, ventana);
  });
  
  // Evento para mostrar mensajes de error
  ipcMain.on('mostrar-error', (event, mensaje) => {
    mostrarVentanaError(mensaje, ventana);
  });
}

/**
 * Maneja el proceso de selección de archivo Excel
 */
function manejarSeleccionArchivo(ventana) {
  dialog.showOpenDialog(ventana, {
    properties: ['openFile'],
    filters: [
      {name: 'Archivos de Excel', extensions: ['xlsx']}
    ],
  }).then(resultado => {
    if (!resultado.canceled && resultado.filePaths.length > 0) {
      const rutaArchivo = resultado.filePaths[0];
      const nombreArchivo = path.basename(rutaArchivo);
      
      try {
        const datos = leerArchivoExcel(rutaArchivo);
        
        // Validar que se obtuvieron datos correctamente
        if (!datos || !Array.isArray(datos)) {
          throw new Error('No se pudieron leer datos válidos del archivo');
        }
        
        ventana.webContents.send('archivo-seleccionado', {
          nombreArchivo: nombreArchivo,
          datos: datos,
        });
        console.log('Datos leídos del archivo Excel:', datos.length, 'filas');
      } catch (error) {
        console.error('Error procesando el archivo Excel:', error);
        mostrarVentanaError(`Error al procesar el archivo: ${error.message}`, ventana);
      }
    }
  }).catch(err => {
    console.error('Error al seleccionar archivo:', err);
    mostrarVentanaError(`Error al seleccionar archivo: ${err.message}`, ventana);
  });
}

/**
 * Maneja el proceso de guardado de PDF
 */
async function manejarGuardadoPDF(event, ventana) {
  try {
    const resultado = await dialog.showSaveDialog(ventana, {
      defaultPath: 'reporte.pdf',
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    });
    
    if (!resultado.canceled) {
      event.reply('ruta-guardar-pdf', resultado.filePath);
    } else {
      event.reply('ruta-guardar-pdf', null);
    }
  } catch (error) {
    console.error('Error al abrir el diálogo de guardar:', error);
    event.reply('ruta-guardar-pdf', null);
    mostrarVentanaError(`Error al abrir el diálogo para guardar: ${error.message}`, ventana);
  }
}

/**
 * Guarda el archivo PDF en el sistema
 */
function guardarArchivoPDF(rutaArchivo, datoBuffer, ventana) {
  try {
    // Convertir los datos recibidos a Buffer
    const buffer = Buffer.from(datoBuffer);
    
    // Validar la ruta del archivo
    if (!rutaArchivo || typeof rutaArchivo !== 'string') {
      throw new Error('Ruta de archivo inválida');
    }
    
    // Escribir el archivo
    fs.writeFileSync(rutaArchivo, buffer);
    
    // Mostrar mensaje de éxito
    dialog.showMessageBox(ventana, {
      type: 'info',
      title: 'PDF Guardado',
      message: 'El PDF se ha guardado correctamente',
      buttons: ['OK']
    });
  } catch (error) {
    console.error('Error al guardar el PDF:', error);
    mostrarVentanaError(`No se pudo guardar el PDF: ${error.message}`, ventana);
  }
}

/**
 * Muestra una ventana de error
 */
function mostrarVentanaError(mensaje, ventana) {
  dialog.showErrorBox('Error', mensaje);
}

/**
 * Lee un archivo Excel y lo convierte a un array bidimensional
 */
function leerArchivoExcel(rutaArchivo) {
  try {
    console.log('Leyendo archivo Excel...', rutaArchivo);
    
    // Validar que el archivo existe
    if (!fs.existsSync(rutaArchivo)) {
      throw new Error('El archivo no existe');
    }
    
    // Leer el archivo
    const libro = XLSX.readFile(rutaArchivo);
    
    // Verificar que tenga al menos una hoja
    if (!libro.SheetNames || libro.SheetNames.length === 0) {
      throw new Error('El archivo Excel no contiene hojas');
    }
    
    // Convertir la primera hoja a un array bidimensional
    const nombreHoja = libro.SheetNames[0];
    const hoja = libro.Sheets[nombreHoja];
    const datos = XLSX.utils.sheet_to_json(hoja, { header: 1 });
    
    // Validar que hay datos
    if (!datos || datos.length === 0) {
      throw new Error('No se encontraron datos en el archivo Excel');
    }
    
    return datos;
  } catch (err) {
    console.error('Error al leer archivo Excel:', err);
    throw err;
  }
}

// Iniciar la aplicación cuando Electron esté listo
app.whenReady().then(() => {
  crearVentanaPrincipal();

  // En macOS, recrear la ventana cuando se haga clic en el icono del dock
  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) {
      crearVentanaPrincipal();
    }
  });
});

// Cerrar la aplicación cuando todas las ventanas estén cerradas (excepto en macOS)
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});