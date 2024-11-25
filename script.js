document.getElementById('csv-file').addEventListener('change', handleFileSelect);

function handleFileSelect(event) {
  const files = event.target.files;
  if (files.length === 0) return;

  const filePromises = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    filePromises.push(processFile(file));
  }

  Promise.all(filePromises).then(results => {
    const combinedData = combineData(results);
    const table = createTable(combinedData);
    displayTable(table);
  }).catch(error => {
    console.error('Error processing files:', error);
  });
}

function processFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(event) {
      const contents = event.target.result;
      const config = {
        delimiter: ";",
        header: false,
        skipEmptyLines: true,
        complete: function(results) {
          const project = results.data[1][0]; // Valor en la celda A2
          const filteredData = results.data.slice(4) // Excluir las primeras 4 filas
            .filter(row => row[2].trim().toLowerCase() === "instalar"); // Filtrar "Instalar" en columna C

          resolve({ project: project, data: filteredData });
        },
        error: function(error) {
          reject(error);
        }
      };
      Papa.parse(contents, config);
    };
    reader.onerror = function(error) {
      reject(error);
    };
    reader.readAsText(file, 'ISO-8859-1');
  });
}

function combineData(results) {
  const combinedData = {};

  results.forEach(result => {
    const project = result.project;
    result.data.forEach(row => {
      const service = row[0];
      const ctdAprobada = parseFloat(row[3]) || 0;

      if (!combinedData[service]) {
        combinedData[service] = [];
      }

      combinedData[service].push({
        service: service,
        description: row[1],
        tipoServicio: row[2],
        ctdAprobada: ctdAprobada,
        project: project
      });
    });
  });

  const finalData = [];
  for (const service in combinedData) {
    const serviceData = combinedData[service];
    let subtotal = 0;

    serviceData.forEach(item => {
      finalData.push(item);
      subtotal += item.ctdAprobada;
    });

    const lastRow = serviceData[serviceData.length - 1];

    // Agregar la fila de subtotal copiando 'servicio' y 'descripción' de la fila superior en negrita
    finalData.push({
      service: lastRow.service,
      description: lastRow.description,
      tipoServicio: 'SUBTOTAL',
      ctdAprobada: subtotal,
      project: '' // puedes ajustar si necesitas el proyecto aquí también
    });
  }

  return finalData;
}

function createTable(combinedData) {
  const table = document.createElement('table');
  table.setAttribute('lang', 'es');

  const headerRow = document.createElement('tr');
  const thService = document.createElement('th');
  thService.textContent = 'SERVICIO';
  const thDescription = document.createElement('th');
  thDescription.textContent = 'DESCRIPCION';
  const thTipoServicio = document.createElement('th');
  thTipoServicio.textContent = 'TIPO DE SERVICIO';
  const thAprobada = document.createElement('th');
  thAprobada.textContent = 'CTD APROBADA';
  const thProject = document.createElement('th');
  thProject.textContent = 'PROYECTO';
  headerRow.appendChild(thService);
  headerRow.appendChild(thDescription);
  headerRow.appendChild(thTipoServicio);
  headerRow.appendChild(thAprobada);
  headerRow.appendChild(thProject);
  table.appendChild(headerRow);

  combinedData.forEach(rowData => {
    const tableRow = document.createElement('tr');
    const tdService = document.createElement('td');
    tdService.textContent = rowData.service;
    const tdDescription = document.createElement('td');
    tdDescription.textContent = rowData.description;
    const tdTipoServicio = document.createElement('td');
    tdTipoServicio.textContent = rowData.tipoServicio;
    const tdAprobada = document.createElement('td');
    tdAprobada.textContent = rowData.ctdAprobada ? rowData.ctdAprobada.toFixed(2).replace('.', ',') : ''; // Convertir punto a coma
    const tdProject = document.createElement('td');
    tdProject.textContent = rowData.project;

    if (rowData.tipoServicio === 'SUBTOTAL') {
      tdService.style.fontWeight = 'bold';
      tdDescription.style.fontWeight = 'bold';
      tdTipoServicio.style.fontWeight = 'bold';
      tdAprobada.style.fontWeight = 'bold';
    }

    tableRow.appendChild(tdService);
    tableRow.appendChild(tdDescription);
    tableRow.appendChild(tdTipoServicio);
    tableRow.appendChild(tdAprobada);
    tableRow.appendChild(tdProject);

    table.appendChild(tableRow);
  });

  return table;
}

function displayTable(table) {
  const tableContainer = document.getElementById('table-container');
  tableContainer.innerHTML = '';
  tableContainer.appendChild(table);
}

document.getElementById('export-button').addEventListener('click', exportToCSV);

function exportToCSV() {
  const table = document.querySelector('table');
  const headers = Array.from(table.querySelectorAll('th')).map(header => header.textContent);
  const rows = table.querySelectorAll('tr');
  const csvData = [];

  csvData.push(headers.join(';'));

  rows.forEach(row => {
    const rowData = [];
    const cells = row.querySelectorAll('td');
    cells.forEach((cell, index) => {
      let textContent = cell.textContent;
      if (index === 3) {
        const numericValue = parseFloat(textContent.replace(',', '.'));
        if (!isNaN(numericValue)) {
          rowData.push(numericValue.toFixed(2).replace('.', ','));
        } else {
          rowData.push('');
        }
      } else {
        textContent = textContent.replace(/"/g, '""');
        rowData.push('"' + textContent + '"');
      }
    });
    csvData.push(rowData.join(';'));
  });

  const csvContent = csvData.join('\n');

  const blob = new Blob(["\uFEFF" + csvContent], { type: 'application/vnd.ms-excel;charset=ANSI;' });
  const link = document.createElement('a');
  link.setAttribute('href', window.URL.createObjectURL(blob));
  link.setAttribute('download', 'data.csv');

  link.click();
}

