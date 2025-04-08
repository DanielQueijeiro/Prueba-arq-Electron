/**
 * Módulo de precarga: Establece el puente seguro entre el proceso principal y el renderizador
 * Expone funcionalidades específicas del sistema a través de contextBridge
 */
const { contextBridge, ipcRenderer } = require('electron/renderer');

// Exposición de APIs seguras al proceso de renderizado
contextBridge.exposeInMainWorld('electronAPI', {
  /**
   * Abre el diálogo de selección de archivos Excel
   */
  abrirDialogoArchivo: () => ipcRenderer.send('abrir-dialogo-archivo'),
  
  /**
   * Escucha cuando un archivo ha sido seleccionado
   */
  enArchivoSeleccionado: (callback) => 
    ipcRenderer.on('archivo-seleccionado', (event, rutaArchivo) => callback(rutaArchivo)),
  
  /**
   * Inicia el proceso de guardado de PDF
   */
  guardarPDF: (callback) => {
    ipcRenderer.send('dialogo-guardar-pdf');
    ipcRenderer.once('ruta-guardar-pdf', (event, rutaArchivo) => callback(rutaArchivo));
  },
  
  /**
   * Guarda un archivo PDF en el sistema
   */
  escribirArchivoPDF: (rutaArchivo, arrayBuffer) => {
    const buffer = new Uint8Array(arrayBuffer);
    ipcRenderer.send('escribir-archivo-pdf', rutaArchivo, buffer);
  },
  
  /**
   * Muestra un mensaje de error al usuario
   */
  mostrarError: (mensaje) => ipcRenderer.send('mostrar-error', mensaje)
});