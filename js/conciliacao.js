let workbook, worksheet, data = [];

document.getElementById('excelFile').addEventListener('change', handleFile);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function (event) {
    const dataBinary = new Uint8Array(event.target.result);
    workbook = XLSX.read(dataBinary, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: ""
    });

    if (worksheet.length === 0 || !worksheet[0]) {
      alert("A planilha está vazia ou mal formatada.");
      return;
    }

    data = worksheet;
    document.getElementById('colunasSelect').style.display = 'block';
    exibirTabela(data);
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

function letraParaIndice(letra) {
  letra = letra.toUpperCase().trim();
  let indice = 0;
  for (let i = 0; i < letra.length; i++) {
    indice *= 26;
    indice += letra.charCodeAt(i) - 64;
  }
  return indice - 1;
}

function indiceParaLetra(indice) {
  let letra = '';
  while (indice >= 0) {
    letra = String.fromCharCode((indice % 26) + 65) + letra;
    indice = Math.floor(indice / 26) - 1;
  }
  return letra;
}

document.getElementById('conciliarBtn').addEventListener('click', () => {
  const baseLetra = document.getElementById('colunaBase').value;
  const alvoLetra = document.getElementById('colunaAlvo').value;
  const concLetra = document.getElementById('colunaConciliacao').value;

  const baseIdx = letraParaIndice(baseLetra);
  const alvoIdx = letraParaIndice(alvoLetra);
  const concIdx = letraParaIndice(concLetra);

  const maxCols = Math.max(...data.map(row => row.length));
  if (concIdx >= maxCols) {
    data[0][concIdx] = "Conciliado";
    for (let i = 1; i < data.length; i++) {
      data[i][concIdx] = "";
    }
  }

  for (let i = 1; i < data.length; i++) {
    const baseVal = data[i][baseIdx];
    const conciliadoBase = data[i][concIdx];

    if (
      conciliadoBase === 'Sim' ||
      baseVal === undefined ||
      baseVal === null ||
      baseVal === ''
    ) continue;

    for (let j = 1; j < data.length; j++) {
      const alvoVal = data[j][alvoIdx];
      const conciliadoAlvo = data[j][concIdx];

      if (
        alvoVal === baseVal &&
        conciliadoAlvo !== 'Sim'
      ) {
        data[i][concIdx] = 'Sim';
        data[j][concIdx] = 'Sim';
        break;
      }
    }
  }

  exibirTabela(data);
  document.getElementById('baixarBtn').style.display = 'inline-block';
});

function exibirTabela(data) {
  const container = document.getElementById('tabelaContainer');
  container.innerHTML = '';

  const table = document.createElement('table');
  table.classList.add('tabela-conciliada');

  const numCols = Math.max(...data.map(row => row.length));

  // Linha com letras das colunas (A, B, C...) — PRIMEIRA LINHA
  const letrasRow = document.createElement('tr');
  for (let j = 0; j < numCols; j++) {
    const th = document.createElement('th');
    th.textContent = indiceParaLetra(j);
    letrasRow.appendChild(th);
  }
  table.appendChild(letrasRow);

  // Linhas da planilha (incluindo cabeçalho original)
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

document.getElementById('baixarBtn').addEventListener('click', () => {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Conciliado');
  XLSX.writeFile(wb, 'planilha_conciliada.xlsx');
});