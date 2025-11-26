let allData = [];
let headers = [];
let selectedIndexes = [];
let rawSheet = null;
let columnFormats = {};
let filteredRows = [];

// ponte para disparar filtro sem recriar tabela
window.applyFiltersTrigger = null;

document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function (ev) {
    const data = new Uint8Array(ev.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    rawSheet = workbook.Sheets[sheetName];

    allData = XLSX.utils.sheet_to_json(rawSheet, { header: 1, raw: true });
    headers = Array.isArray(allData[0]) ? allData[0] : [];
    selectedIndexes = headers.map((_, i) => i);

    detectColumnFormats();
    createCheckboxes(headers);
    createGlobalFilter();
    updateTable();
  };

  reader.readAsArrayBuffer(file);
}

function detectColumnFormats() {
  columnFormats = {};
  const sampleSize = Math.min(20, Math.max(0, allData.length - 1));

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
  container.innerHTML = '<strong>Selecione as colunas para exibir:</strong><br>';
  headers.forEach((header, index) => {
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.value = index;
    checkbox.addEventListener('change', () => {
      selectedIndexes = Array.from(document.querySelectorAll('#checkboxes input[type="checkbox"]'))
        .filter(cb => cb.checked)
        .map(cb => parseInt(cb.value, 10));
      updateTable();
    });

    const label = document.createElement('label');
    label.textContent = `${columnLetter(index)} - ${safeHeaderText(header)}`;
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
  wrapper.style.alignItems = 'center';
  wrapper.style.flexWrap = 'wrap';

  // Label e campo de busca
  const searchLabel = document.createElement('label');
  searchLabel.textContent = 'Busca:';
  searchLabel.setAttribute('for', 'searchInput');

  const searchInput = document.createElement('input');
  searchInput.type = 'text';
  searchInput.id = 'searchInput';
  searchInput.placeholder = 'Digite o termo...';
  searchInput.style.width = '220px';

  // Label e campo de coluna (pequeno)
  const columnLabel = document.createElement('label');
  columnLabel.textContent = 'Coluna:';
  columnLabel.setAttribute('for', 'searchColumnInput');

  const columnInput = document.createElement('input');
  columnInput.type = 'text';
  columnInput.id = 'searchColumnInput';
  columnInput.placeholder = 'Ex: A ou Nome';
  columnInput.style.width = '80px'; // campo pequeno

  const trigger = () => {
    if (typeof window.applyFiltersTrigger === 'function') {
      window.applyFiltersTrigger();
    }
  };
  searchInput.addEventListener('input', trigger);
  columnInput.addEventListener('input', trigger);

  wrapper.appendChild(searchLabel);
  wrapper.appendChild(searchInput);
  wrapper.appendChild(columnLabel);
  wrapper.appendChild(columnInput);

  container.appendChild(wrapper);
}

function updateTable() {
  const selectedHeaders = selectedIndexes.map(i => safeHeaderText(headers[i]));
  renderTableWithFilters(allData, selectedHeaders, selectedIndexes);
}

function renderTableWithFilters(data, selectedHeaders, selectedIndexes, chunkSize = 500) {
  const table = document.getElementById('dataTable');
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const excelRow = document.createElement('tr');   // linha com letras A, B, C...
  const headerRow = document.createElement('tr');  // linha com nomes originais

  selectedHeaders.forEach((header, idx) => {
    const colIndex = selectedIndexes[idx];

    // Cabeçalho com letra da coluna
    const thExcel = document.createElement('th');
    thExcel.textContent = columnLetter(colIndex);
    excelRow.appendChild(thExcel);

    // Cabeçalho com nome original
    const thHeader = document.createElement('th');
    thHeader.textContent = header;
    headerRow.appendChild(thHeader);
  });

  thead.appendChild(excelRow);
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  table.appendChild(tbody);

  function applyFilters() {
    const searchValue = document.getElementById('searchInput')?.value || '';
    const columnSpec = document.getElementById('searchColumnInput')?.value || '';

    const normalizedSearch = normalize(searchValue);

    // sem busca: todas as linhas
    if (!normalizedSearch) {
      filteredRows = data.slice(1).map((_, i) => i + 1);
      renderChunk();
      return;
    }

    // coluna alvo
    const targetIndex = resolveColumnIndex(columnSpec);

    // fallback: primeira coluna selecionada; se não houver, sem filtro
    const colIdxToUse = Number.isInteger(targetIndex)
      ? targetIndex
      : (selectedIndexes[0] ?? null);

    if (colIdxToUse === null) {
      filteredRows = data.slice(1).map((_, i) => i + 1);
      renderChunk();
      return;
    }

    filteredRows = data.slice(1).map((_, i) => i + 1).filter(rowIdx => {
      const formattedValue = normalize(formatCell(rowIdx, colIdxToUse));
      return formattedValue.includes(normalizedSearch);
    });

    renderChunk();
  }

  function normalize(text) {
    if (text == null) return '';
    const s = String(text);
    return s
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

    // Formatação de data
    if (formatType === 'data') {
      // Excel serial
      if (typeof cell.v === 'number') {
        const date = XLSX.SSF.parse_date_code(cell.v);
        if (date && date.y > 1900 && date.y < 2100) {
          const day = String(date.d).padStart(2, '0');
          const month = String(date.m).padStart(2, '0');
          const year = date.y;
          return `${day}/${month}/${year}`;
        }
      }
      // String de data
      if (typeof cell.v === 'string') {
        const parsed = new Date(cell.v);
        if (!isNaN(parsed.getTime())) {
          const day = String(parsed.getDate()).padStart(2, '0');
          const month = String(parsed.getMonth() + 1).padStart(2, '0');
          const year = parsed.getFullYear();
          return `${day}/${month}/${year}`;
        }
      }
    }

    // Formatação de valor
    if (formatType === 'valor') {
      if (typeof cell.v === 'number') {
        let formatted = XLSX.SSF.format('#,##0.00', cell.v);
        // trocar separadores para padrão brasileiro
        formatted = formatted.replace(/,/g, '#').replace(/\./g, ',').replace(/#/g, '.');
        return formatted;
      }
    }

    // Padrão: texto exibido (.w) ou valor bruto
    return cell.w != null ? String(cell.w) : String(cell.v);
  }

  // expõe o gatilho de filtro para os inputs
  window.applyFiltersTrigger = () => applyFilters();

  // inicia sem filtro
  applyFilters();
}

// Função para converter índice em letra de coluna Excel (A, B, C...)
function columnLetter(index) {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

// Resolve coluna digitada: aceita letra (A, B...), nome da coluna (headers) ou número (0, 1...)
function resolveColumnIndex(input) {
  if (!input) return null;

  const raw = input.trim();

  // número (índice)
  if (/^\d+$/.test(raw)) {
    const idx = parseInt(raw, 10);
    return Number.isInteger(idx) && idx >= 0 ? idx : null;
  }

  // letras estilo Excel (A, B, AA)
  const letters = raw.toUpperCase().replace(/\s+/g, '');
  if (/^[A-Z]+$/.test(letters)) {
    return columnIndexFromLetters(letters);
  }

  // nome do cabeçalho (case-insensitive, ignorando espaços)
  const norm = normalizeName(raw);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeName(headers[i]) === norm) return i;
  }

  return null;
}

// Converte letras (A, B, AA) em índice 0-based
function columnIndexFromLetters(letters) {
  let idx = 0;
  for (let i = 0; i < letters.length; i++) {
    idx = idx * 26 + (letters.charCodeAt(i) - 64);
  }
  return idx - 1; // 0-based
}

function normalizeName(s) {
  return String(s).toLowerCase().replace(/\s+/g, '');
}

function safeHeaderText(h) {
  if (h == null) return '';
  return String(h);
}