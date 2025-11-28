let workbook, worksheet, data = [];

// ===============================
// 📌 CARREGA ARQUIVO EXCEL (preservando formatação)
// ===============================
document.getElementById('excelFile').addEventListener('change', handleFile);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function (event) {
    const dataBinary = new Uint8Array(event.target.result);
    workbook = XLSX.read(dataBinary, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];

    // Extrai os valores formatados do Excel (cell.w)
    const range = XLSX.utils.decode_range(ws['!ref']);
    worksheet = [];

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = { c: C, r: R };
        const cellRef = XLSX.utils.encode_cell(cellAddress);
        const cell = ws[cellRef];
        // Pega o valor formatado se disponível, caso contrário o valor cru, senão ""
        row.push(cell ? (cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : "")) : "");
      }
      worksheet.push(row);
    }

    if (worksheet.length === 0 || !worksheet[0]) {
      alert("A planilha está vazia ou mal formatada.");
      return;
    }

    data = worksheet;
    document.getElementById('colunasSelect').style.display = 'block';

    exibirTabela(data);
    aplicarFiltro();
    limparTotais();
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

// ===============================
// 📌 CONVERTE LETRAS PARA ÍNDICE
// ===============================
function letraParaIndice(letra) {
  letra = letra.toUpperCase().trim();
  let indice = 0;
  for (let i = 0; i < letra.length; i++) {
    indice *= 26;
    indice += letra.charCodeAt(i) - 64;
  }
  return indice - 1;
}

// ===============================
// 📌 CONVERTE ÍNDICE PARA LETRA
// ===============================
function indiceParaLetra(indice) {
  let letra = '';
  while (indice >= 0) {
    letra = String.fromCharCode((indice % 26) + 65) + letra;
    indice = Math.floor(indice / 26) - 1;
  }
  return letra;
}

// ===============================
// 📌 BOTÃO CONCILIAR
// ===============================
document.getElementById('conciliarBtn').addEventListener('click', () => {
  const baseLetra = document.getElementById('colunaBase').value;
  const alvoLetra = document.getElementById('colunaAlvo').value;
  const concLetra = document.getElementById('colunaConciliacao').value;

  if (!baseLetra || !alvoLetra || !concLetra) {
    alert("Preencha todas as colunas antes de conciliar!");
    return;
  }

  const baseIdx = letraParaIndice(baseLetra);
  const alvoIdx = letraParaIndice(alvoLetra);
  const concIdx = letraParaIndice(concLetra);

  const maxCols = Math.max(...data.map(row => row.length));

  if (concIdx >= maxCols) {
    data[0][concIdx] = "Conciliado";
    for (let i = 1; i < data.length; i++) data[i][concIdx] = "";
  }

  // 🔄 LÓGICA DE CONCILIAÇÃO
  for (let i = 1; i < data.length; i++) {
    const baseVal = data[i][baseIdx];
    const conciliadoBase = data[i][concIdx];

    if (conciliadoBase === 'Sim' || baseVal === undefined || baseVal === null || baseVal === '') continue;

    for (let j = 1; j < data.length; j++) {
      const alvoVal = data[j][alvoIdx];
      const conciliadoAlvo = data[j][concIdx];

      if (alvoVal === baseVal && conciliadoAlvo !== 'Sim') {
        data[i][concIdx] = 'Sim';
        data[j][concIdx] = 'Sim';
        break;
      }
    }
  }

  exibirTabela(data);
  aplicarFiltro();
  atualizarTotais(baseIdx, alvoIdx, concIdx);

  document.getElementById('baixarBtn').style.display = 'inline-block';
});

// ===============================
// 📌 EXIBE TABELA
// ===============================
function exibirTabela(data) {
  const container = document.getElementById('tabelaContainer');
  container.innerHTML = '';

  const table = document.createElement('table');
  table.classList.add('tabela-conciliada');

  const numCols = Math.max(...data.map(row => row.length));

  // Linha com letras das colunas
  const letrasRow = document.createElement('tr');
  for (let j = 0; j < numCols; j++) {
    const th = document.createElement('th');
    th.textContent = indiceParaLetra(j);
    letrasRow.appendChild(th);
  }
  table.appendChild(letrasRow);

  // Linhas da planilha
  data.forEach((row, i) => {
    const tr = document.createElement('tr');
    for (let j = 0; j < numCols; j++) {
      const td = document.createElement(i === 0 ? 'th' : 'td');
      td.textContent = row[j] !== undefined ? row[j] : '';
      tr.appendChild(td);
    }
    table.appendChild(tr);
  });

  container.appendChild(table);
}

// ===============================
// 📌 EXPORTAR PLANILHA RESPEITANDO FILTRO
// ===============================
document.getElementById('baixarBtn').addEventListener('click', () => {
  const table = document.querySelector('.tabela-conciliada');
  if (!table) return;

  const linhas = table.querySelectorAll('tr');
  const dataFiltrada = [];

  linhas.forEach((tr, i) => {
    if (i === 0) {
      const ths = tr.querySelectorAll('th');
      dataFiltrada.push(Array.from(ths).map(th => th.textContent));
    } else if (tr.style.display !== "none") {
      const tds = tr.querySelectorAll('td, th');
      dataFiltrada.push(Array.from(tds).map(td => td.textContent));
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(dataFiltrada);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Conciliado');
  XLSX.writeFile(wb, 'planilha_conciliada.xlsx');
});

// ===============================
// 📌 FILTRO DE CONCILIAÇÃO
// ===============================
function aplicarFiltro() {
  const select = document.getElementById('filtroConciliacao');
  if (!select) return;

  const filtro = select.value;
  const concLetra = document.getElementById('colunaConciliacao').value;
  if (!concLetra) return;

  const concIdx = letraParaIndice(concLetra);

  const table = document.querySelector('.tabela-conciliada');
  if (!table) return;

  const linhas = table.querySelectorAll('tr');

  for (let i = 2; i < linhas.length; i++) {
    const td = linhas[i].querySelectorAll('td')[concIdx];
    const valor = td ? td.textContent.trim().toLowerCase() : "";

    if (filtro === "todos") linhas[i].style.display = "";
    else if (filtro === "sim" && valor === "sim") linhas[i].style.display = "";
    else if (filtro === "nao" && valor !== "sim") linhas[i].style.display = "";
    else linhas[i].style.display = "none";
  }

  const baseIdx = letraParaIndice(document.getElementById('colunaBase').value);
  const alvoIdx = letraParaIndice(document.getElementById('colunaAlvo').value);
  const concIdx2 = letraParaIndice(document.getElementById('colunaConciliacao').value);

  atualizarTotais(baseIdx, alvoIdx, concIdx2);
}

document.getElementById('filtroConciliacao')
  .addEventListener('change', aplicarFiltro);

// ===============================
// 📌 TOTALIZAÇÕES (linhas visíveis)
// ===============================
function atualizarTotais(baseIdx, alvoIdx, concIdx) {
  const area = document.getElementById("totaisArea");
  if (!area) return;

  const table = document.querySelector('.tabela-conciliada');
  if (!table) return;

  const linhas = table.querySelectorAll("tr");

  let totalBase = 0;
  let totalAlvo = 0;
  let totalConc = 0;

  for (let i = 2; i < linhas.length; i++) {
    if (linhas[i].style.display === "none") continue;

    const tds = linhas[i].querySelectorAll("td");

    const baseValStr = tds[baseIdx]?.textContent.trim().replace(/\./g,'').replace(',', '.');
    const alvoValStr = tds[alvoIdx]?.textContent.trim().replace(/\./g,'').replace(',', '.');

    const baseVal = parseFloat(baseValStr);
    const alvoVal = parseFloat(alvoValStr);
    const concVal = tds[concIdx]?.textContent.trim();

    if (!isNaN(baseVal)) totalBase += baseVal;
    if (!isNaN(alvoVal)) totalAlvo += alvoVal;
    if (concVal === "Sim") totalConc++;
  }

  area.innerHTML = `
    <div style="margin-bottom:8px;">📌 <strong>Totalizações (Filtradas)</strong></div>
    <div>🟦 Total da coluna base (${indiceParaLetra(baseIdx)}): <strong>${totalBase.toLocaleString('pt-BR')}</strong></div>
    <div>🟩 Total da coluna alvo (${indiceParaLetra(alvoIdx)}): <strong>${totalAlvo.toLocaleString('pt-BR')}</strong></div>
    <div>🟨 Total conciliados (visíveis): <strong>${totalConc}</strong></div>
  `;
}

function limparTotais() {
  const area = document.getElementById("totaisArea");
  if (area) area.innerHTML = "";
}
