let originalData = [];
let currentRange = null;
let currentFormat = null;
let formatActive = false;
let concIdxAtual = null; // Índice da coluna de conciliação atual
let filtroAtual = "todos";

const controlsContainer = document.getElementById('controls-container');
const conciliaButton = document.getElementById('btn-concilia');

// ==================== CARREGAR EXCEL ====================
document.getElementById('input-excel').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        currentRange = range;

        originalData = [];
        for (let R = range.s.r; R <= range.e.r; ++R) {
            const row = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                const cell = worksheet[cell_ref];
                row.push(cell ? cell.v : '');
            }
            originalData.push(row);
        }

        renderTable(originalData, currentRange);
        controlsContainer.style.display = 'block';
    };
    reader.readAsArrayBuffer(file);
});

// ==================== RENDERIZAR TABELA ====================
function renderTable(data, range) {
    let table = '<table id="excel-table">';
    table += '<tr>';
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const colLetter = XLSX.utils.encode_col(C);
        table += `<th data-col="${C}"><div class="header-cell"><span>${colLetter}</span></div></th>`;
    }
    table += '</tr>';

    for (let R = 0; R < data.length; R++) {
        table += '<tr>';
        for (let C = 0; C < data[R].length; C++) {
            const cell_value = data[R][C];
            let align = 'left';
            if (typeof cell_value === 'number') align = 'right';
            table += `<td data-row="${R}" data-col="${C}" style="text-align:${align};">${cell_value}</td>`;
        }
        table += '</tr>';
    }

    table += '</table>';
    document.getElementById('tabela-container').innerHTML = table;

    document.querySelectorAll('#excel-table th').forEach(th => {
        th.addEventListener('click', () => {
            if (!formatActive) return;
            const colIndex = th.getAttribute('data-col');
            applyFormatToColumn(colIndex, currentFormat);
        });
    });
}

// ==================== FORMATAÇÃO ====================
function applyFormatToColumn(colIndex, format) {
    const table = document.getElementById('excel-table');
    for (let i = 1; i < table.rows.length; i++) {
        const cell = table.rows[i].cells[colIndex];
        let value = originalData[i-1][colIndex];

        if (format === 'original') {
            cell.textContent = value;
            cell.style.textAlign = (typeof value === 'number') ? 'right' : 'left';
        } else if (format === 'data') {
            const date = new Date(value);
            if (!isNaN(date)) {
                const day = String(date.getDate()).padStart(2,'0');
                const month = String(date.getMonth()+1).padStart(2,'0');
                const year = date.getFullYear();
                cell.textContent = `${day}/${month}/${year}`;
            } else { cell.textContent = value; }
            cell.style.textAlign = 'center';
        } else if (format === 'valor') {
            if (typeof value === 'number') {
                cell.textContent = value.toLocaleString('pt-BR', { minimumFractionDigits:2, maximumFractionDigits:2 });
            } else { cell.textContent = value; }
            cell.style.textAlign = 'right';
        }
    }
}

document.getElementById('btn-original').addEventListener('click', () => activateFormat('original'));
document.getElementById('btn-data').addEventListener('click', () => activateFormat('data'));
document.getElementById('btn-valor').addEventListener('click', () => activateFormat('valor'));

function activateFormat(formatType) { 
    currentFormat = formatType; 
    formatActive = true; 
}

document.addEventListener('click', (event) => {
    const tableContainer = document.getElementById('tabela-container');
    const clickedInsideTable = tableContainer.contains(event.target);
    const clickedButton = document.getElementById('format-buttons')?.contains(event.target);

    if (!clickedInsideTable && !clickedButton) formatActive = false;
});

// ==================== CONCILIAÇÃO ====================
conciliaButton.addEventListener('click', () => {
    const colRefLetter = document.getElementById('col-ref').value.toUpperCase();
    const colSearchLetter = document.getElementById('col-search').value.toUpperCase();
    const colConcLetter = document.getElementById('col-conc').value.toUpperCase();

    if (!colRefLetter || !colSearchLetter || !colConcLetter) {
        alert('Informe colunas de referência, procura e conciliação!');
        return;
    }

    const colRef = XLSX.utils.decode_col(colRefLetter);
    const colSearch = XLSX.utils.decode_col(colSearchLetter);
    const colConc = XLSX.utils.decode_col(colConcLetter);
    concIdxAtual = colConc;

    const table = document.getElementById('excel-table');
    const headerRow = table.rows[0];

    // ==================== CABEÇALHO CONCILIAÇÃO ====================
    if (!headerRow.cells[colConc]) {
        while (headerRow.cells.length <= colConc) {
            headerRow.insertCell();
        }
    }

    // Linha 0 = letra da coluna
    headerRow.cells[colConc].innerHTML = `<div class="header-cell"><span>${colConcLetter}</span></div>`;

    // Linha 1 = "Conciliado"
    if (!table.rows[1].cells[colConc]) {
        table.rows[1].insertCell(colConc);
    }
    table.rows[1].cells[colConc].innerHTML = `<div class="header-cell"><span>Conciliado</span></div>`;
    table.rows[1].cells[colConc].style.textAlign = 'center';

    // Inicializar células de conciliação das linhas restantes
    for (let i = 2; i < table.rows.length; i++) {
        if (!table.rows[i].cells[colConc]) {
            table.rows[i].insertCell(colConc);
        }
        table.rows[i].cells[colConc].textContent = '';
        table.rows[i].cells[colConc].style.textAlign = 'center';
    }

    const refUsed = new Array(table.rows.length).fill(false);
    const searchUsed = new Array(table.rows.length).fill(false);

    for (let i = 2; i < table.rows.length; i++) { // começar da linha 2
        if (refUsed[i]) continue;
        const valRef = originalData[i-1][colRef];
        if (valRef === undefined || valRef === null || valRef === '') continue;

        for (let j = 2; j < table.rows.length; j++) {
            if (searchUsed[j]) continue;
            const valSearch = originalData[j-1][colSearch];

            if (valRef === valSearch) {
                table.rows[i].cells[colConc].textContent = 'Sim';
                table.rows[j].cells[colConc].textContent = 'Sim';
                refUsed[i] = true;
                searchUsed[j] = true;
                break;
            }
        }
    }

    aplicarFiltro(filtroAtual);
});

// ==================== FILTRO ====================
document.querySelectorAll('#filtro-conc button[data-filtro]').forEach(btn => {
    btn.addEventListener('click', () => {
        filtroAtual = btn.dataset.filtro;
        aplicarFiltro(filtroAtual);
    });
});

function aplicarFiltro(filtro) {
    if (concIdxAtual === null) return;
    const table = document.getElementById('excel-table');
    for (let i = 2; i < table.rows.length; i++) { // Começar da linha 2, cabeçalhos sempre visíveis
        const valor = table.rows[i].cells[concIdxAtual].textContent.trim().toLowerCase();
        if (filtro === "todos") table.rows[i].style.display = "";
        else if (filtro === "sim" && valor === "sim") table.rows[i].style.display = "";
        else if (filtro === "nao" && valor !== "sim") table.rows[i].style.display = "";
        else table.rows[i].style.display = "none";
    }
}

// ==================== EXPORTAÇÃO ====================
document.getElementById('btn-export').addEventListener('click', () => {
    const table = document.getElementById('excel-table');
    const dataFiltrada = [];

    for (let i = 0; i < table.rows.length; i++) {
        // Sempre incluir as duas primeiras linhas
        if (i > 1 && table.rows[i].style.display === "none") continue; 
        const tds = table.rows[i].cells;
        const linhaArray = [];
        // Ignorar coluna 0
        for (let j = 1; j < tds.length; j++) {
            linhaArray.push(tds[j].textContent);
        }
        dataFiltrada.push(linhaArray);
    }

    const ws = XLSX.utils.aoa_to_sheet(dataFiltrada);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Conciliado');
    XLSX.writeFile(wb, 'planilha_filtrada.xlsx');
});
