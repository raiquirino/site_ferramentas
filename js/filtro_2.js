let allData = [];
let headers = [];
let selectedIndexes = [];
let rawSheet = null;
let columnFormats = {};
let filteredRows = [];

document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    rawSheet = workbook.Sheets[sheetName];

    allData = XLSX.utils.sheet_to_json(rawSheet, { header: 1, raw: true });
    headers = allData[0];
    selectedIndexes = headers.map((_, i) => i);
    detectColumnFormats();
    createCheckboxes(headers);
    updateTable();
  };

  reader.readAsArrayBuffer(file);
}

function detectColumnFormats() {
  columnFormats = {};
  const sampleSize = Math.min(20, allData.length - 1);

  selectedIndexes.forEach(colIdx => {
    let dateCount = 0;
    let numberCount = 0;

    for (let i = 1; i <= sampleSize; i++) {
      const cellAddress = XLSX.utils.encode_cell({ r: i, c: colIdx });
      const cell = rawSheet[cellAddress];
      if (!cell || cell.v == null) continue;

      if (cell.t === 'n') {
        const date = XLSX.SSF.parse_date_code(cell.v);
        if (date && date.y > 1900 && date.y < 2100) {
          dateCount++;
        } else {
          numberCount++;
        }
      }
    }

    if (dateCount > sampleSize / 2) {
      columnFormats[colIdx] = 'data';
    } else if (numberCount > sampleSize / 2) {
      columnFormats[colIdx] = 'valor';
    } else {
      columnFormats[colIdx] = 'original';
    }
  });
}

function createCheckboxes(headers) {
  const container = document.getElementById('checkboxes');
  container.innerHTML = '<strong>Selecione as colunas:</strong><br>';
  headers.forEach((header, index) => {
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.value = index;
    checkbox.addEventListener('change', () => {
      selectedIndexes = Array.from(document.querySelectorAll('#checkboxes input[type="checkbox"]'))
        .filter(cb => cb.checked)
        .map(cb => parseInt(cb.value));
      updateTable();
    });

    const label = document.createElement('label');
    label.textContent = header;
    label.style.marginRight = '10px';

    container.appendChild(checkbox);
    container.appendChild(label);
  });
}

function updateTable() {
  const selectedHeaders = selectedIndexes.map(i => headers[i]);
  renderTableWithFilters(allData, selectedHeaders, selectedIndexes);
}

function renderTableWithFilters(data, selectedHeaders, selectedIndexes, chunkSize = 500) {
  const table = document.getElementById('dataTable');
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const filterRow = document.createElement('tr');
  const headerRow = document.createElement('tr');

  selectedHeaders.forEach((header, idx) => {
    const colIndex = selectedIndexes[idx];

    const th = document.createElement('th');
    const wrapper = document.createElement('div');
    wrapper.style.display = 'flex';
    wrapper.style.gap = '4px';
    wrapper.style.alignItems = 'center';

    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = 'Filtrar...';
    input.dataset.index = colIndex;
    input.style.flex = '1';
    input.addEventListener('input', () => applyFilters());

    const select = document.createElement('select');
    select.dataset.index = colIndex;
    ['original', 'data', 'valor'].forEach(opt => {
      const option = document.createElement('option');
      option.value = opt;
      option.textContent = opt.charAt(0).toUpperCase() + opt.slice(1);
      select.appendChild(option);
    });
    select.value = columnFormats[colIndex];
    select.style.width = '80px';
    select.addEventListener('change', () => {
      columnFormats[colIndex] = select.value;
      applyFilters();
    });

    wrapper.appendChild(input);
    wrapper.appendChild(select);
    th.appendChild(wrapper);
    filterRow.appendChild(th);

    const thHeader = document.createElement('th');
    thHeader.textContent = header;
    headerRow.appendChild(thHeader);
  });

  thead.appendChild(filterRow);
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  table.appendChild(tbody);

  function applyFilters() {
    const inputs = document.querySelectorAll('thead input');
    const filters = Array.from(inputs).map(input => ({
      index: parseInt(input.dataset.index),
      value: normalize(input.value)
    }));

    filteredRows = data.slice(1).map((_, i) => i + 1).filter(rowIdx =>
      filters.every(f => {
        const formattedValue = normalize(formatCell(rowIdx, f.index));
        return formattedValue.includes(f.value);
      })
    );

    renderChunk();
  }

  function normalize(text) {
    return text
      .toString()
      .toLowerCase()
      .replace(/\s/g, '')
      .replace(/\./g, '')
      .replace(/,/g, '.');
  }

  function renderChunk() {
    tbody.innerHTML = '';
    let start = 0;

    function chunk() {
      const end = Math.min(start + chunkSize, filteredRows.length);
      const fragment = document.createDocumentFragment();

      for (let i = start; i < end; i++) {
        const rowIdx = filteredRows[i];
        const row = document.createElement('tr');
        selectedIndexes.forEach(colIdx => {
          const cell = document.createElement('td');
          cell.textContent = formatCell(rowIdx, colIdx);
          row.appendChild(cell);
        });
        fragment.appendChild(row);
      }

      tbody.appendChild(fragment);
      start = end;

      if (start < filteredRows.length) {
        setTimeout(chunk, 0);
      }
    }

    chunk();
  }

  function formatCell(rowIdx, colIdx) {
    const cellAddress = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
    const cell = rawSheet[cellAddress];
    if (!cell || cell.v == null) return '';

    const formatType = columnFormats[colIdx];

    if (formatType === 'data') {
      if (typeof cell.v === 'number') {
        const date = XLSX.SSF.parse_date_code(cell.v);
        if (date && date.y > 1900 && date.y < 2100) {
          const day = String(date.d).padStart(2, '0');
          const month = String(date.m).padStart(2, '0');
          const year = date.y;
          return `${day}/${month}/${year}`;
        }
      }
      if (typeof cell.v === 'string') {
        const parsed = new Date(cell.v);
        if (!isNaN(parsed)) {
          const day = String(parsed.getDate()).padStart(2, '0');
          const month = String(parsed.getMonth() + 1).padStart(2, '0');
          const year = parsed.getFullYear();
          return `${day}/${month}/${year}`;
        }
      }
    }

    if (formatType === 'valor') {
      if (typeof cell.v === 'number') {
        let formatted = XLSX.SSF.format('#,##0.00', cell.v);
        formatted = formatted.replace(/,/g, '#').replace(/\./g, ',').replace(/#/g, '.');
        return formatted;
      }
    }

    return cell.w || cell.v.toString();
  }

  applyFilters();
}