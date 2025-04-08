const subirExcel = document.getElementById('excelBtn');
const fileNameElement = document.getElementById('fileName');
const pdfButton = document.getElementById('pdfBtn');
let chart = null;
let excelData = null;

window.addEventListener('DOMContentLoaded', () => {
  console.log('DOM completamente cargado');
  if (typeof Chart === 'undefined') {
    console.error('Chart.js no está cargado correctamente');
  } else {
    console.log('Chart.js cargado correctamente');
  }
  
  if (typeof window.jspdf === 'undefined' && typeof window.jsPDF === 'undefined') {
    console.error('jsPDF no está cargado correctamente');
  } else {
    console.log('jsPDF cargado correctamente');
  }
});

subirExcel.addEventListener('click', () => {
  window.electronAPI.openFileDialog();
});

pdfButton.addEventListener('click', () => {
  if (chart && excelData) {
    generatePDF(excelData);
  }
});

window.electronAPI.onFileSelected((fileData) => {
  console.log('Archivo seleccionado:', fileData);
  
  excelData = fileData.data;

  fileNameElement.textContent = `Archivo seleccionado: ${fileData.fileName}`;

  pdfButton.disabled = false;

  displayDataPreview(fileData.data);

  setTimeout(() => {
    createChart(fileData.data);
  }, 100);
});

function displayDataPreview(data) {
  if (!data || data.length === 0) return;

  const existingPreview = document.getElementById('dataPreview');
  if (existingPreview) {
    existingPreview.remove();
  }

  const previewContainer = document.createElement('div');
  previewContainer.id = 'dataPreview';
  previewContainer.className = 'data-preview';

  const table = document.createElement('table');
  table.className = 'excel-data-table';

  const maxRows = Math.min(10, data.length);
  const maxCols = data[0] ? Math.min(5, data[0].length) : 0;

  for (let i = 0; i < maxRows; i++) {
    const row = document.createElement('tr');
    
    for (let j = 0; j < maxCols; j++) {
      const cell = i === 0 ? document.createElement('th') : document.createElement('td');
      if (data[i] && data[i][j] !== undefined) {
        cell.textContent = data[i][j];
      } else {
        cell.textContent = '';
      }
      row.appendChild(cell);
    }
    
    table.appendChild(row);
  }

  const infoText = document.createElement('p');
  infoText.textContent = `Archivo completo: ${data.length} filas, ${data[0] ? data[0].length : 0} columnas`;

  previewContainer.appendChild(infoText);
  previewContainer.appendChild(table);

  fileNameElement.parentNode.appendChild(previewContainer);
}

function createChart(data) {
  console.log('Intentando crear gráfico con datos:', data);

  const canvas = document.getElementById('myChart');
  if (!canvas) {
    console.error('No se encuentra el elemento canvas con id "myChart"');
    return;
  }

  if (typeof Chart === 'undefined') {
    console.error('Chart.js no está disponible. Verifica la importación.');
    return;
  }

  if (chart) {
    chart.destroy();
  }
  
  try {
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      console.error('No se puede obtener el contexto 2d del canvas');
      return;
    }
    
    const headers = data[0];

    const chartData = data.slice(1);

    const labels = chartData.map(row => row[1]);
    const values = chartData.map(row => row[0]);
    
    chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: headers[0],
          data: values,
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
              text: headers[0]
            }
          },
          x: {
            title: {
              display: true,
              text: headers[1]
            }
          }
        },
        plugins: {
          title: {
            display: true,
            text: `${headers[0]} por ${headers[1]}`
          }
        }
      }
    });
    
    console.log('Gráfico creado correctamente');
  } catch (error) {
    console.error('Error al crear el gráfico:', error);
  }
}

function generatePDF(data) {
  try {
    const jsPDFClass = window.jspdf?.jsPDF || window.jsPDF;
    if (!jsPDFClass) {
      console.error('jsPDF no está disponible');
      window.electronAPI.showError('jsPDF no está disponible. Verifica la instalación.');
      return;
    }
    
    const doc = new jsPDFClass({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });
    
    doc.setFontSize(16);
    doc.text("Reporte de Datos", 14, 15);
    
    const currentDate = new Date().toLocaleDateString();
    doc.setFontSize(10);
    doc.text(`Fecha: ${currentDate}`, 14, 22);
    
    const headers = data[0];
    const tableData = data.slice(1);

    doc.autoTable({
      head: [headers],
      body: tableData,
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
    
    doc.addPage();

    doc.setFontSize(16);
    doc.text(`Gráfico: ${headers[0]} por ${headers[1]}`, 14, 15);

    const canvas = document.getElementById('myChart');
    if (!canvas) {
      console.error('No se encuentra el elemento canvas');
      return;
    }

    const chartImage = canvas.toDataURL('image/png', 1.0);

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const chartWidth = pageWidth - 30;
    const chartHeight = chartWidth * (canvas.height / canvas.width);

    doc.addImage(chartImage, 'PNG', 15, 25, chartWidth, chartHeight);

    window.electronAPI.savePDF((filePath) => {
      if (filePath) {
        const pdfBlob = doc.output('blob');
        const reader = new FileReader();
        reader.onload = () => {
          window.electronAPI.writePDFFile(filePath, reader.result);
        };
        reader.readAsArrayBuffer(pdfBlob);
      }
    });
    
  } catch (error) {
    console.error('Error al generar el PDF:', error);
    window.electronAPI.showError(`Error al generar el PDF: ${error.message}`);
  }
}