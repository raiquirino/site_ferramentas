// ===== Estado global =====
let allData = [];        // Matriz [linhas][colunas] com valores brutos da planilha
let headers = [];        // Cabeçalhos (linha 0 de allData)
let filteredRows = [];   // Índices de linhas (na matriz allData) que passam no filtro
let columnFormats = {};  // Mapa {colIdx: 'original' | 'data' | 'valor'}

// ===== Inicialização: input de arquivo =====
document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Carrega como matriz, não objetos — garante alinhamento por coluna
    allData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
    headers = allData[0] || [];

    // Formatos padrão
    headers.forEach((_, colIdx) => {
      columnFormats[colIdx] = 'original';
    });

    createControls(headers);
    updateTable(); // Render inicial
  };
  reader.readAsArrayBuffer(file);
}

// ===== Utilidades =====
function columnIndexToLetter(index) {
  // 0->A, 1->B, ..., 25->Z, 26->AA ...
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode((index % 26) + 65) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

function normalize(text) {
  return String(text)
    .normalize('NFD')                // separa acentos
    .replace(/[\u0300-\u036f]/g, '') // remove acentos
    .toLowerCase()
    .trim();
}

function rawCellText(rowIdx, colIdx) {
  const v = allData[rowIdx]?.[colIdx];
  if (v == null) return '';
  return String(v);
}

// ===== Controles (filtro + select de coluna + select de formato por coluna) =====
function createControls(headers) {
  const container = document.getElementById('checkboxes');
  container.innerHTML = '';

  // Campo de busca
  const globalFilter = document.createElement('input');
  globalFilter.type = 'text';
  globalFilter.id = 'globalFilter';
  globalFilter.placeholder = 'Digite para filtrar...';
  globalFilter.style.width = '260px';
  globalFilter.style.marginRight = '10px';

  // Select da coluna alvo
  const columnSelect = document.createElement('select');
  columnSelect.id = 'filterColumn';
  headers.forEach((header, index) => {
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = `${columnIndexToLetter(index)} - ${header}`;
    columnSelect.appendChild(option);
  });
  columnSelect.value = columnSelect.options[0]?.value ?? '0';

  // Eventos: recalculam e renderizam
  globalFilter.addEventListener('input', updateTable);
  columnSelect.addEventListener('change', updateTable);

  container.appendChild(globalFilter);
  container.appendChild(columnSelect);
}

// ===== Atualiza filtro e render =====
function updateTable() {
  const globalFilter = document.getElementById('globalFilter');
  const columnSelect = document.getElementById('filterColumn');

  const filterValue = globalFilter ? globalFilter.value : '';
  let colIndex = columnSelect ? Number(columnSelect.value) : 0;
  if (Number.isNaN(colIndex) || colIndex < 0 || colIndex >= headers.length) colIndex = 0;

  // Calcula linhas filtradas (allData: linha 0 = cabeçalho, dados a partir da linha 1)
  if (!filterValue) {
    filteredRows = allData.slice(1).map((_, i) => i + 1);
  } else {
    const normalizedFilter = normalize(filterValue);
    filteredRows = allData.slice(1).map((_, i) => i + 1).filter(rowIdx => {
      const text = rawCellText(rowIdx, colIndex);
      return normalize(text).includes(normalizedFilter);
    });
  }

  renderTableWithFilters();
}

// ===== Renderização da tabela =====
function renderTableWithFilters(chunkSize = 500) {
  const table = document.getElementById('dataTable');
  table.innerHTML = '';

  // Cabeçalho
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  headers.forEach((header, colIndex) => {
    const thHeader = document.createElement('th');

    // Título
    const headerText = document.createElement('div');
    headerText.textContent = header;
    headerText.style.fontWeight = 'bold';

    // Select de formato
    const fmtSelect = document.createElement('select');
    fmtSelect.dataset.index = String(colIndex);
    ['original', 'data', 'valor'].forEach(opt => {
      const option = document.createElement('option');
      option.value = opt;
      option.textContent = opt.charAt(0).toUpperCase() + opt.slice(1);
      fmtSelect.appendChild(option);
    });
    fmtSelect.value = columnFormats[colIndex] || 'original';
    fmtSelect.style.marginTop = '4px';
    fmtSelect.style.fontSize = '12px';
    fmtSelect.style.width = '90px';
    fmtSelect.addEventListener('change', () => {
      columnFormats[colIndex] = fmtSelect.value;
      // Re-render para aplicar formatação (não muda filtro)
      renderTableWithFilters();
    });

    thHeader.appendChild(headerText);
    thHeader.appendChild(fmtSelect);
    headerRow.appendChild(thHeader);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Corpo
  const tbody = document.createElement('tbody');
  table.appendChild(tbody);

  tbody.innerHTML = '';
  let start = 0;

  function chunk() {
    const end = Math.min(start + chunkSize, filteredRows.length);
    const fragment = document.createDocumentFragment();

    for (let i = start; i < end; i++) {
      const rowIdx = filteredRows[i];
      const tr = document.createElement('tr');

      headers.forEach((_, colIdx) => {
        const td = document.createElement('td');
        td.textContent = displayCell(rowIdx, colIdx);
        tr.appendChild(td);
      });

      fragment.appendChild(tr);
    }

    tbody.appendChild(fragment);
    start = end;

    if (start < filteredRows.length) {
      setTimeout(chunk, 0);
    }
  }

  chunk();
}

// ===== Exibição (formatação) =====
function displayCell(rowIdx, colIdx) {
  const v = allData[rowIdx]?.[colIdx];
  if (v == null) return '';

  const fmt = columnFormats[colIdx] || 'original';

  if (fmt === 'data') {
    // Tenta tratar número como série de data e string como ISO/BR
    if (typeof v === 'number') {
      const date = XLSX.SSF.parse_date_code(v);
      if (date && date.y >= 1900 && date.y <= 2100) {
        const dd = String(date.d).padStart(2, '0');
        const mm = String(date.m).padStart(2, '0');
        const yyyy = date.y;
        return `${dd}/${mm}/${yyyy}`;
      }
    }
    const parsed = new Date(v);
    if (!isNaN(parsed)) {
      const dd = String(parsed.getDate()).padStart(2, '0');
      const mm = String(parsed.getMonth() + 1).padStart(2, '0');
      const yyyy = parsed.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }
    return String(v);
  }

  if (fmt === 'valor') {
    const num = Number(v);
    if (!Number.isNaN(num)) {
      return num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(v);
  }

  return String(v);
}