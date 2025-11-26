let allData = [];
let headers = [];
let selectedIndexes = [];
let rawSheet = null;
let columnFormats = {};
let filteredRows = [];
let formattedCache = {};   // <<< NOVO: cache de células formatadas

window.applyFiltersTrigger = null;

document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(ev) {
    const data = new Uint8Array(ev.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    rawSheet = workbook.Sheets[sheetName];

    allData = XLSX.utils.sheet_to_json(rawSheet, { header: 1, raw: true });
    headers = Array.isArray(allData[0]) ? allData[0] : [];
    selectedIndexes = headers.map((_, i) => i);

    detectColumnFormats();
    initFormatCache();
    createCheckboxes(headers);
    createGlobalFilter();
    updateTable(true);    // <<< primeira vez, recria cabeçalho
  };

  reader.readAsArrayBuffer(file);
}

/*--------------------------------------------------
  FORMATAÇÃO E CACHE
---------------------------------------------------*/

function detectColumnFormats() {
  columnFormats = {};
  selectedIndexes.forEach(i => columnFormats[i] = 'original');
}

function initFormatCache() {
  formattedCache = {};
  selectedIndexes.forEach(col => {
    formattedCache[col] = {};
  });
}

function invalidateColumnCache(colIdx) {
  formattedCache[colIdx] = {};
}

/*--------------------------------------------------
  CHECKBOXES E FILTROS
---------------------------------------------------*/

function createCheckboxes(headers) {
  const container = document.getElementById('checkboxes');
  container.innerHTML = '<strong>Selecionar colunas:</strong><br>';

  headers.forEach((header, index) => {
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.value = index;

    checkbox.addEventListener('change', () => {
      selectedIndexes = [...container.querySelectorAll('input[type="checkbox"]')]
        .filter(cb => cb.checked)
        .map(cb => parseInt(cb.value, 10));

      initFormatCache();
      updateTable(true);
    });

    const label = document.createElement('label');
    label.textContent = `${columnLetter(index)} - ${header ?? ''}`;
    label.style.marginRight = '10px';

    container.appendChild(checkbox);
    container.appendChild(label);
  });
}

function createGlobalFilter() {
  const container = document.getElementById('globalFilter');
  container.innerHTML = '';

  const wrapper = document.createElement('div');
  wrapper.style.display = 'flex';
  wrapper.style.gap = '12px';
  wrapper.style.flexWrap = 'wrap';

  const searchInput = document.createElement('input');
  searchInput.id = 'searchInput';
  searchInput.placeholder = 'Digite o termo...';

  const colInput = document.createElement('input');
  colInput.id = 'searchColumnInput';
  colInput.placeholder = 'Coluna...';

  const trigger = () => window.applyFiltersTrigger?.();
  searchInput.addEventListener('input', trigger);
  colInput.addEventListener('input', trigger);

  wrapper.append('Busca:', searchInput, 'Coluna:', colInput);
  container.appendChild(wrapper);
}

/*--------------------------------------------------
  UPDATE TABLE
---------------------------------------------------*/

function updateTable(rebuildHeader = false) {
  renderTableWithFilters(allData, selectedIndexes.map(i => headers[i] ?? ''), selectedIndexes, rebuildHeader);
}

/*--------------------------------------------------
  RENDER TABLE
---------------------------------------------------*/

function renderTableWithFilters(data, selectedHeaders, selectedIndexes, rebuildHeader, chunkSize = 500) {
  const table = document.getElementById('dataTable');

  if (rebuildHeader) renderHeader(table, selectedHeaders, selectedIndexes);

  const tbody = table.querySelector('tbody');
  if (!tbody) {
    const tb = document.createElement('tbody');
    table.appendChild(tb);
  }

  function applyFilters() {
    const searchValue = normalize(document.getElementById('searchInput')?.value || '');
    const columnSpec = document.getElementById('searchColumnInput')?.value || '';

    const targetIndex = resolveColumnIndex(columnSpec);
    const col = Number.isInteger(targetIndex) ? targetIndex : selectedIndexes[0] ?? null;

    if (!searchValue) {
      filteredRows = data.slice(1).map((_, i) => i + 1);
    } else {
      filteredRows = data.slice(1).map((_, i) => i + 1).filter(r =>
        normalize(formatCell(r, col)).includes(searchValue)
      );
    }

    renderRows();
  }

  function renderRows() {
    const tbody = table.querySelector('tbody');
    tbody.innerHTML = '';

    let start = 0;

    function chunk() {
      const end = Math.min(start + chunkSize, filteredRows.length);
      const frag = document.createDocumentFragment();

      for (let i = start; i < end; i++) {
        const row = document.createElement('tr');
        const rIdx = filteredRows[i];

        selectedIndexes.forEach(col => {
          const td = document.createElement('td');
          td.textContent = formatCell(rIdx, col);
          row.appendChild(td);
        });

        frag.appendChild(row);
      }

      tbody.appendChild(frag);
      start = end;

      if (start < filteredRows.length) setTimeout(chunk, 0);
    }

    chunk();
  }

  window.applyFiltersTrigger = applyFilters;
  applyFilters();
}

/*--------------------------------------------------
  HEADER RENDER
---------------------------------------------------*/

function renderHeader(table, selectedHeaders, selectedIndexes) {
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const excelRow = document.createElement('tr');
  const headerRow = document.createElement('tr');

  selectedHeaders.forEach((header, i) => {
    const col = selectedIndexes[i];

    const thExcel = document.createElement('th');
    thExcel.textContent = columnLetter(col);
    excelRow.appendChild(thExcel);

    const thHeader = document.createElement('th');
    thHeader.textContent = header;

    // select de formato
    const fmt = document.createElement('select');
    ['original', 'data', 'valor'].forEach(opt => {
      const o = document.createElement('option');
      o.value = opt;
      o.textContent = opt;
      fmt.appendChild(o);
    });

    fmt.value = columnFormats[col];
    fmt.style.marginLeft = '6px';

    fmt.addEventListener('change', () => {
      columnFormats[col] = fmt.value;
      invalidateColumnCache(col);
      updateTable(false);
    });

    thHeader.appendChild(fmt);
    headerRow.appendChild(thHeader);
  });

  thead.append(excelRow, headerRow);
  table.appendChild(thead);
}

/*--------------------------------------------------
  CELLS
---------------------------------------------------*/

function formatCell(rowIdx, colIdx) {
  if (formattedCache[colIdx]?.[rowIdx]) return formattedCache[colIdx][rowIdx];

  const cell = rawSheet[XLSX.utils.encode_cell({ r: rowIdx, c: colIdx })];
  if (!cell || cell.v == null) return '';

  let out = '';
  const fmt = columnFormats[colIdx];

  if (fmt === 'data') {
    if (typeof cell.v === 'number') {
      const d = XLSX.SSF.parse_date_code(cell.v);
      if (d) out = `${d.d.toString().padStart(2,'0')}/${d.m.toString().padStart(2,'0')}/${d.y}`;
    } else {
      const d = new Date(cell.v);
      if (!isNaN(d)) out = `${d.getDate().toString().padStart(2,'0')}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getFullYear()}`;
    }
  }

  else if (fmt === 'valor' && typeof cell.v === 'number') {
    out = XLSX.SSF.format('#,##0.00', cell.v)
      .replace(/,/g, '#').replace(/\./g, ',').replace(/#/g, '.');
  }

  else {
    out = cell.w ?? String(cell.v);
  }

  formattedCache[colIdx][rowIdx] = out;
  return out;
}

/*--------------------------------------------------
  HELPERS
---------------------------------------------------*/

const normalize = t => t.toLowerCase().replace(/[ .]/g, '').replace(/,/g, '.');

function columnLetter(index) {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

function resolveColumnIndex(input) {
  if (!input) return null;
  const raw = input.trim();

  if (/^\d+$/.test(raw)) return parseInt(raw, 10);

  const letters = raw.toUpperCase();
  if (/^[A-Z]+$/.test(letters)) return columnIndexFromLetters(letters);

  const norm = raw.toLowerCase().replace(/\s+/g, '');
  for (let i = 0; i < headers.length; i++) {
    if ((headers[i] + '').toLowerCase().replace(/\s+/g, '') === norm) return i;
  }

  return null;
}

function columnIndexFromLetters(letters) {
  let idx = 0;
  for (let i = 0; i < letters.length; i++) {
    idx = idx * 26 + (letters.charCodeAt(i) - 64);
  }
  return idx - 1;
}
