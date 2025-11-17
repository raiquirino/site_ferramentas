let sheetData = [];
let columnFormats = {};

document.getElementById("excelInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    showPreview(sheetData);
  };

  reader.readAsArrayBuffer(file);
});

function colLetterToIndex(letter) {
  return letter.toUpperCase().charCodeAt(0) - 65;
}

function formatValue(value, type) {
  if (type === "date") {
    if (value instanceof Date) return value.toLocaleDateString("pt-BR");
    if (typeof value === "number") {
      const date = XLSX.SSF?.parse_date_code?.(value);
      if (date) {
        return `${String(date.d).padStart(2, "0")}/${String(date.m).padStart(2, "0")}/${date.y}`;
      }
    }
    if (typeof value === "string") {
      const parts = value.split(/[\/\-]/);
      if (parts.length === 3) {
        let [d, m, y] = parts;
        if (y.length === 2) y = "20" + y;
        return `${d.padStart(2, "0")}/${m.padStart(2, "0")}/${y}`;
      }
    }
    return value;
  }

  if (type === "currency") {
    const num = parseFloat(value?.toString().replace(",", "."));
    if (!isNaN(num)) {
      return num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return value;
  }

  return value ?? "";
}

function showPreview(data) {
  const table = document.getElementById("previewTable");
  table.innerHTML = "";

  const colCount = Math.max(...data.map(row => row.length));
  const headerRow = document.createElement("tr");

  for (let i = 0; i < colCount; i++) {
    const th = document.createElement("th");
    const wrapper = document.createElement("div");
    wrapper.className = "col-header";

    const label = document.createElement("span");
    label.className = "col-label";
    label.textContent = String.fromCharCode(65 + i); // A, B, C...

    const select = document.createElement("select");
    select.innerHTML = `
      <option value="auto">Auto</option>
      <option value="text">Texto</option>
      <option value="date">Data</option>
      <option value="currency">Valor</option>
    `;
    select.dataset.col = i;
    select.value = columnFormats[i] || "auto";
    select.addEventListener("change", () => {
      columnFormats[i] = select.value;
      showPreview(sheetData); // Atualiza visualização
    });

    wrapper.appendChild(label);
    wrapper.appendChild(select);
    th.appendChild(wrapper);
    headerRow.appendChild(th);
  }

  table.appendChild(headerRow);

  const maxRows = Math.min(data.length, 30);
  for (let i = 0; i < maxRows; i++) {
    const row = document.createElement("tr");
    for (let j = 0; j < colCount; j++) {
      const cell = document.createElement("td");
      const raw = data[i]?.[j];
      const type = columnFormats[j] || "auto";
      const formatted = type === "auto" ? raw : formatValue(raw, type);
      cell.textContent = formatted ?? "";
      row.appendChild(cell);
    }
    table.appendChild(row);
  }

  document.getElementById("previewContainer").style.display = "block";
}

document.getElementById("btnProcess").addEventListener("click", () => {
  const col1 = document.getElementById("col1").value.trim();
  const col2 = document.getElementById("col2").value.trim();
  const dest = document.getElementById("destCol").value.trim();
  const progress = document.getElementById("progress");
  const link = document.getElementById("downloadExcel");

  if (!col1 || !col2 || !dest) {
    alert("Preencha todas as colunas com letras válidas.");
    return;
  }

  const col1Index = colLetterToIndex(col1);
  const col2Index = colLetterToIndex(col2);
  const destIndex = colLetterToIndex(dest);

  const contador = {};
  const resultado = [];

  sheetData.forEach((row, i) => {
    const novaLinha = [];

    for (let j = 0; j < row.length; j++) {
      const tipo = columnFormats[j] || "auto";
      const valor = tipo === "auto" ? row[j] : formatValue(row[j], tipo);
      novaLinha[j] = valor;
    }

    if (i === 0) {
      novaLinha[destIndex] = "Repetição";
    } else {
      const val1 = row[col1Index];
      const val2 = row[col2Index];

      if (val1 !== undefined && val1 !== "" && val2 !== undefined && val2 !== "") {
        const f1 = formatValue(val1, columnFormats[col1Index] || "auto");
        const f2 = formatValue(val2, columnFormats[col2Index] || "auto");
        const chave = `${f1}||${f2}`;
        contador[chave] = (contador[chave] || 0) + 1;
        novaLinha[destIndex] = contador[chave];
      } else {
        novaLinha[destIndex] = "";
      }
    }

    resultado.push(novaLinha);
  });

  showPreview(resultado);

  const worksheet = XLSX.utils.aoa_to_sheet(resultado);
  const workbookOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbookOut, worksheet, "Repetições");

  const wbout = XLSX.write(workbookOut, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);

  link.href = url;
  link.download = "planilha_repeticoes_formatada.xlsx";
  link.style.display = "inline";

  progress.textContent = `✅ Planilha processada com sucesso. Clique acima para baixar.`;
});