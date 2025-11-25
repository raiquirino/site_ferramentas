// Estado global simples
let workbook = null;
let originalData = null; // Array de arrays (AOA)
let groupedData = null;  // Array de arrays (AOA)
let firstSheetName = null;

// 1 - Ler planilha
document.getElementById("fileInput").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (evt) {
    const data = new Uint8Array(evt.target.result);
    workbook = XLSX.read(data, { type: "array" });
    firstSheetName = workbook.SheetNames[0];
    const firstSheet = workbook.Sheets[firstSheetName];

    // Converte para AOA (primeira linha: cabeçalhos)
    originalData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: true });

    // Exibe metadados
    document.getElementById("metaInfo").textContent =
      `Arquivo: ${file.name} | Aba: ${firstSheetName} | Linhas: ${originalData.length}`;

    // Mostra tabela
    showTable(originalData, "tableContainer");
  };
  reader.readAsArrayBuffer(file);
});

// 2 - Mostrar dados em tabela HTML estilo Excel
function showTable(data, containerId) {
  const container = document.getElementById(containerId);
  if (!data || !data.length) {
    container.innerHTML = "<div class='hint'>Nenhum dado para exibir.</div>";
    return;
  }

  let html = "<table>";
  const maxCols = Math.max(...data.map(r => r.length));

  // Cabeçalho de colunas (A, B, C…)
  html += "<thead><tr><th>#</th>";
  for (let c = 0; c < maxCols; c++) {
    html += `<th>${String.fromCharCode(65 + c)}</th>`;
  }
  html += "</tr></thead>";

  // Corpo da tabela
  html += "<tbody>";
  for (let r = 0; r < data.length; r++) {
    html += `<tr><th>${r + 1}</th>`;
    for (let c = 0; c < maxCols; c++) {
      const cell = data[r][c] ?? "";
      html += `<td>${cell}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table>";

  container.innerHTML = html;
}

// 3 - Processar agrupamento
document.getElementById("processBtn").addEventListener("click", function () {
  if (!originalData || originalData.length < 2) {
    alert("Carregue uma planilha válida primeiro.");
    return;
  }

  const colDataIdx = colLetterToIndex(document.getElementById("colData").value);
  const colDescIdx = colLetterToIndex(document.getElementById("colDesc").value);
  const colVal1Idx = colLetterToIndex(document.getElementById("colVal1").value);
  const colVal2Idx = colLetterToIndex(document.getElementById("colVal2").value);

  if ([colDataIdx, colDescIdx, colVal1Idx, colVal2Idx].some(v => v == null)) {
    alert("Digite corretamente as colunas (A, B, C...).");
    return;
  }

  const map = {};

  for (let i = 1; i < originalData.length; i++) {
    const row = originalData[i] || [];
    const dataVal = row[colDataIdx];
    const descVal = row[colDescIdx];
    if (dataVal == null || descVal == null || descVal === "") continue;

    const v1 = toNumberSafe(row[colVal1Idx]);
    const v2 = toNumberSafe(row[colVal2Idx]);

    const dateKey = normalizeDateKey(dataVal);
    const descKey = String(descVal).trim();
    const key = `${dateKey}||${descKey}`;

    if (!map[key]) {
      // copia todas as colunas originais
      map[key] = { row: [...row], data: dateKey, desc: descKey, v1: 0, v2: 0 };
    }
    map[key].v1 += v1;
    map[key].v2 += v2;
  }

  // Cabeçalho: mantém todas as colunas originais
  groupedData = [originalData[0].map(h => h || "")];

  Object.values(map).forEach(obj => {
    const newRow = [...obj.row];
    newRow[colDataIdx] = obj.data;
    newRow[colDescIdx] = obj.desc;
    newRow[colVal1Idx] = round2(obj.v1);
    newRow[colVal2Idx] = round2(obj.v2);
    groupedData.push(newRow);
  });

  showTable(groupedData, "resultContainer");
});

// Helpers
function colLetterToIndex(letter) {
  if (!letter) return null;
  const l = letter.trim().toUpperCase();
  let idx = 0;
  for (let i = 0; i < l.length; i++) {
    const c = l.charCodeAt(i) - 64;
    if (c < 1 || c > 26) return null;
    idx = idx * 26 + c;
  }
  return idx - 1;
}
function toNumberSafe(val) {
  if (val == null || val === "") return 0;
  const n = typeof val === "number" ? val : Number(String(val).replace(",", "."));
  return isNaN(n) ? 0 : n;
}
function round2(n) {
  return Math.round(n * 100) / 100;
}
function normalizeDateKey(d) {
  // Se já for objeto Date
  if (d instanceof Date) {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  // Se vier como número (serial Excel)
  if (typeof d === "number") {
    const date = XLSX.SSF.parse_date_code(d);
    if (date) {
      const dd = String(date.d).padStart(2, "0");
      const mm = String(date.m).padStart(2, "0");
      const yyyy = date.y;
      return `${dd}/${mm}/${yyyy}`;
    }
  }

  // Se vier como string já no formato DD-MM-AAAA ou DD/MM/AAAA
  const str = String(d).trim();
  const parts = str.split(/[-\/]/);
  if (parts.length === 3) {
    let [p1, p2, p3] = parts;
    if (p1.length === 2 && p2.length === 2 && p3.length === 4) {
      return `${p1}/${p2}/${p3}`;
    }
  }

  return str;
}

// 4 - Salvar Original
document.getElementById("saveOriginal").addEventListener("click", function () {
  if (!originalData) {
    alert("Nada para salvar.");
    return;
  }
  const ws = XLSX.utils.aoa_to_sheet(originalData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Original");
  XLSX.writeFile(wb, "planilha_original.xlsx");
});

// 5 - Salvar apenas Valor 1
document.getElementById("saveV1").addEventListener("click", function () {
  if (!groupedData || groupedData.length < 2) {
    alert("Primeiro gere o resultado agrupado.");
    return;
  }
  const header = groupedData[0].filter((_, idx) => idx !== colLetterToIndex(document.getElementById("colVal2").value));
  const body = groupedData.slice(1).map(row => row.filter((_, idx) => idx !== colLetterToIndex(document.getElementById("colVal2").value)));
  const aoa = [header, ...body];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Valor1");
  XLSX.writeFile(wb, "agrupado_valor1.xlsx");
});

// 6 - Salvar apenas Valor 2
document.getElementById("saveV2").addEventListener("click", function () {
  if (!groupedData || groupedData.length < 2) {
    alert("Primeiro gere o resultado agrupado.");
    return;
  }
  const header = groupedData[0].filter((_, idx) => idx !== colLetterToIndex(document.getElementById("colVal1").value));
  const body = groupedData.slice(1).map(row => row.filter((_, idx) => idx !== colLetterToIndex(document.getElementById("colVal1").value)));
  const aoa = [header, ...body];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Valor2");
  XLSX.writeFile(wb, "agrupado_valor2.xlsx");
});
