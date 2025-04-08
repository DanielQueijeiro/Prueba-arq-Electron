/**
 * Script del proceso de renderizado
 * Maneja la interfaz de usuario y la interacción con los archivos Excel y PDF
 */

// Elementos DOM
const botonSubirExcel = document.getElementById('excelBtn');
const elementoNombreArchivo = document.getElementById('fileName');
const botonPDF = document.getElementById('pdfBtn');

// Variables globales
let grafico = null;
let datosExcel = null;

/**
 * Valida que las librerías necesarias estén disponibles
 */
function validarDependencias() {
  const dependencias = [
    { nombre: 'Chart.js', objeto: typeof Chart, mensaje: 'Chart.js no está cargado correctamente' },
    { nombre: 'jsPDF', objeto: typeof window.jspdf !== 'undefined' || typeof window.jsPDF !== 'undefined', 
      mensaje: 'jsPDF no está cargado correctamente' }
  ];

  dependencias.forEach(dep => {
    if (!dep.objeto) {
      console.error(dep.mensaje);
      window.electronAPI.mostrarError(dep.mensaje);
    } else {
      console.log(`${dep.nombre} cargado correctamente`);
    }
  });
}

/**
 * Muestra una vista previa de los datos cargados
 */
function mostrarVistaPrevia(datos) {
  if (!datos || datos.length === 0) {
    console.warn('No hay datos para mostrar');
    return;
  }

  // Eliminar vista previa anterior si existe
  const vistaExistente = document.getElementById('dataPreview');
  if (vistaExistente) {
    vistaExistente.remove();
  }

  const contenedorVista = document.createElement('div');
  contenedorVista.id = 'dataPreview';
  contenedorVista.className = 'data-preview';

  const tabla = document.createElement('table');
  tabla.className = 'excel-data-table';

  // Límites para la vista previa
  const maxFilas = Math.min(10, datos.length);
  const maxColumnas = datos[0] ? Math.min(5, datos[0].length) : 0;

  // Crear filas y celdas
  for (let i = 0; i < maxFilas; i++) {
    const fila = document.createElement('tr');
    
    for (let j = 0; j < maxColumnas; j++) {
      const celda = i === 0 ? document.createElement('th') : document.createElement('td');
      
      // Validar que el dato exista antes de asignarlo
      if (datos[i] && datos[i][j] !== undefined) {
        celda.textContent = datos[i][j];
      } else {
        celda.textContent = '';
      }
      
      fila.appendChild(celda);
    }
    
    tabla.appendChild(fila);
  }

  // Información del archivo
  const textoInfo = document.createElement('p');
  textoInfo.textContent = `Archivo completo: ${datos.length} filas, ${datos[0] ? datos[0].length : 0} columnas`;

  contenedorVista.appendChild(textoInfo);
  contenedorVista.appendChild(tabla);

  elementoNombreArchivo.parentNode.appendChild(contenedorVista);
}

/**
 * Crea un gráfico con los datos cargados
 */
function crearGrafico(datos) {
  console.log('Creando gráfico con los datos cargados');

  // Validaciones previas
  if (!datos || !datos.length) {
    console.error('No hay datos para crear el gráfico');
    window.electronAPI.mostrarError('El archivo no contiene datos válidos');
    return;
  }

  const canvas = document.getElementById('myChart');
  if (!canvas) {
    console.error('No se encuentra el elemento canvas con id "myChart"');
    return;
  }

  if (typeof Chart === 'undefined') {
    console.error('Chart.js no está disponible');
    return;
  }

  // Destruir gráfico existente si hay uno
  if (grafico) {
    grafico.destroy();
  }
  
  try {
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      console.error('No se puede obtener el contexto 2d del canvas');
      return;
    }
    
    // Validar estructura de datos
    if (!datos[0] || datos[0].length < 2) {
      console.error('Los datos no tienen el formato esperado (se requieren al menos 2 columnas)');
      window.electronAPI.mostrarError('El formato de los datos no es compatible con el gráfico');
      return;
    }
    
    const encabezados = datos[0];
    const datosGrafico = datos.slice(1);

    // Validar que existan datos además de los encabezados
    if (datosGrafico.length === 0) {
      console.error('No hay datos para graficar más allá de los encabezados');
      window.electronAPI.mostrarError('El archivo no contiene datos para graficar');
      return;
    }

    const etiquetas = datosGrafico.map(fila => fila[1]);
    const valores = datosGrafico.map(fila => {
      // Asegurar que los valores sean numéricos
      const valor = parseFloat(fila[0]);
      return isNaN(valor) ? 0 : valor;
    });
    
    grafico = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: etiquetas,
        datasets: [{
          label: encabezados[0],
          data: valores,
          backgroundColor: 'rgba(54, 162, 235, 0.6)',
          borderColor: 'rgba(54, 162, 235, 1)',
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: true,
            title: {
              display: true,
              text: encabezados[0]
            }
          },
          x: {
            title: {
              display: true,
              text: encabezados[1]
            }
          }
        },
        plugins: {
          title: {
            display: true,
            text: `${encabezados[0]} por ${encabezados[1]}`
          }
        }
      }
    });
    
    console.log('Gráfico creado correctamente');
  } catch (error) {
    console.error('Error al crear el gráfico:', error);
    window.electronAPI.mostrarError(`Error al crear el gráfico: ${error.message}`);
  }
}

/**
 * Genera un PDF con los datos y el gráfico
 */
function generarPDF(datos) {
  try {
    // Obtener la clase jsPDF
    const jsPDFClass = window.jspdf?.jsPDF || window.jsPDF;
    if (!jsPDFClass) {
      console.error('jsPDF no está disponible');
      window.electronAPI.mostrarError('jsPDF no está disponible. Verifica la instalación.');
      return;
    }
    
    // Crear documento PDF
    const doc = new jsPDFClass({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });
    
    // Añadir encabezado
    doc.setFontSize(16);
    doc.text("Reporte de Datos", 14, 15);
    
    // Añadir fecha
    const fechaActual = new Date().toLocaleDateString();
    doc.setFontSize(10);
    doc.text(`Fecha: ${fechaActual}`, 14, 22);
    
    // Validar datos para la tabla
    if (!datos || !datos.length || !datos[0]) {
      window.electronAPI.mostrarError('No hay datos válidos para generar el PDF');
      return;
    }

    const encabezados = datos[0];
    const datosTabla = datos.slice(1);

    // Generar tabla
    doc.autoTable({
      head: [encabezados],
      body: datosTabla,
      startY: 30,
      theme: 'grid',
      styles: {
        fontSize: 8,
        cellPadding: 2,
        overflow: 'linebreak',
      },
      headStyles: {
        fillColor: [54, 162, 235],
        textColor: 255
      }
    });
    
    // Nueva página para el gráfico
    doc.addPage();

    doc.setFontSize(16);
    doc.text(`Gráfico: ${encabezados[0]} por ${encabezados[1]}`, 14, 15);

    // Obtener imagen del gráfico
    const canvas = document.getElementById('myChart');
    if (!canvas) {
      console.error('No se encuentra el elemento canvas');
      window.electronAPI.mostrarError('No se pudo encontrar el gráfico para incluirlo en el PDF');
      return;
    }

    const imagenGrafico = canvas.toDataURL('image/png', 1.0);

    // Calcular dimensiones para mantener proporciones
    const anchoHoja = doc.internal.pageSize.getWidth();
    const altoHoja = doc.internal.pageSize.getHeight();
    const anchoGrafico = anchoHoja - 30;
    const altoGrafico = anchoGrafico * (canvas.height / canvas.width);

    // Añadir imagen del gráfico al PDF
    doc.addImage(imagenGrafico, 'PNG', 15, 25, anchoGrafico, altoGrafico);

    // Guardar el PDF
    window.electronAPI.guardarPDF((rutaArchivo) => {
      if (rutaArchivo) {
        const pdfBlob = doc.output('blob');
        const lector = new FileReader();
        lector.onload = () => {
          window.electronAPI.escribirArchivoPDF(rutaArchivo, lector.result);
        };
        lector.readAsArrayBuffer(pdfBlob);
      }
    });
    
  } catch (error) {
    console.error('Error al generar el PDF:', error);
    window.electronAPI.mostrarError(`Error al generar el PDF: ${error.message}`);
  }
}

// Event Listeners
window.addEventListener('DOMContentLoaded', validarDependencias);

botonSubirExcel.addEventListener('click', () => {
  window.electronAPI.abrirDialogoArchivo();
});

botonPDF.addEventListener('click', () => {
  if (grafico && datosExcel) {
    generarPDF(datosExcel);
  } else {
    window.electronAPI.mostrarError('No hay datos o gráfico para generar el PDF');
  }
});

// Escuchar el evento de archivo seleccionado
window.electronAPI.enArchivoSeleccionado((datosArchivo) => {
  console.log('Archivo seleccionado:', datosArchivo.nombreArchivo);
  
  // Validar datos recibidos
  if (!datosArchivo || !datosArchivo.datos || !Array.isArray(datosArchivo.datos)) {
    window.electronAPI.mostrarError('El archivo seleccionado no contiene datos válidos');
    return;
  }
  
  datosExcel = datosArchivo.datos;

  // Actualizar interfaz
  elementoNombreArchivo.textContent = `Archivo seleccionado: ${datosArchivo.nombreArchivo}`;
  botonPDF.disabled = false;

  // Mostrar datos y crear gráfico
  mostrarVistaPrevia(datosArchivo.datos);

  // Usar setTimeout para asegurar que el DOM esté actualizado
  setTimeout(() => {
    crearGrafico(datosArchivo.datos);
  }, 100);
});