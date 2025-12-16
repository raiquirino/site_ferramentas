// Armazena os dados da planilha carregada
let sheetData = [];

// Guarda os formatos escolhidos para cada coluna (Texto, Data, Valor etc.)
let columnFormats = {};

// Guarda todos os nomes já gerados, para evitar sobrescrever arquivos
let generatedNames = new Set();

// Guarda o arquivo carregado pelo usuário (para usar o nome original)
let file = null;


// 🔹 Evento acionado quando o usuário escolhe um arquivo Excel
document.getElementById("excelInput").addEventListener("change", (e) => {
  file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  // Quando o arquivo estiver totalmente carregado em memória
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);

    // Lê o Excel com XLSX
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0]; // primeira aba
    const sheet = workbook.Sheets[sheetName];

    // Converte a planilha para array de arrays
    sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Exibe prévia na tela
    showPreview(sheetData);
  };

  reader.readAsArrayBuffer(file);
});


// 🔹 Converte letra de coluna (ex: “A”) para índice numérico
function colLetterToIndex(letter) {
  return letter.toUpperCase().charCodeAt(0) - 65;
}


// 🔹 Formata valores conforme o tipo escolhido na interface
function formatValue(value, type) {
  
  // ---- FORMATAÇÃO DE DATA ----
  if (type === "date") {

    // Se já for Date()
    if (value instanceof Date)
      return value.toLocaleDateString("pt-BR");

    // Se for número Excel (ex: 45212)
    if (typeof value === "number") {
      const date = XLSX.SSF?.parse_date_code?.(value);
      if (date) {
        return `${String(date.d).padStart(2,"0")}/${String(date.m).padStart(2,"0")}/${date.y}`;
      }
    }
    return value;
  }

  // ---- FORMATAÇÃO DE VALOR (Moeda) ----
  if (type === "currency") {
    const num = parseFloat(value?.toString().replace(",", "."));
    if (!isNaN(num)) {
      return num.toLocaleString("pt-BR", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    }
    return value;
  }

  // Caso seja texto ou auto
  return value ?? "";
}


// 🔹 Exibe uma pré-visualização da planilha com a formatação aplicada
function showPreview(data) {
  const table = document.getElementById("previewTable");
  table.innerHTML = "";

  // Quantidade de colunas (pega a linha mais longa)
  const colCount = Math.max(...data.map(row => row.length));
  const headerRow = document.createElement("tr");

  // ---- CRIAÇÃO DO CABEÇALHO ----
  for (let i = 0; i < colCount; i++) {
    const th = document.createElement("th");
    const wrapper = document.createElement("div");
    wrapper.className = "col-header";

    // Letra da coluna (A, B, C...)
    const label = document.createElement("span");
    label.className = "col-label";
    label.textContent = String.fromCharCode(65 + i);

    // Menu para escolher o formato da coluna
    const select = document.createElement("select");
    select.innerHTML = `
      <option value="auto">Auto</option>
      <option value="text">Texto</option>
      <option value="date">Data</option>
      <option value="currency">Valor</option>
    `;
    select.dataset.col = i;
    select.value = columnFormats[i] || "auto";

    // Atualiza a prévia ao mudar o formato
    select.addEventListener("change", () => {
      columnFormats[i] = select.value;
      showPreview(sheetData);
    });

    wrapper.appendChild(label);
    wrapper.appendChild(select);
    th.appendChild(wrapper);
    headerRow.appendChild(th);
  }

  table.appendChild(headerRow);

  // Máximo de linhas exibidas na prévia
  const maxRows = Math.min(data.length, 30);

  // ---- CRIAÇÃO DAS LINHAS ----
  for (let i = 0; i < maxRows; i++) {
    const row = document.createElement("tr");

    for (let j = 0; j < colCount; j++) {
      const cell = document.createElement("td");
      const raw = data[i]?.[j];
      const type = columnFormats[j] || "auto";

      // Aplica formatação
      const formatted = type === "auto" ? raw : formatValue(raw, type);
      cell.textContent = formatted ?? "";

      row.appendChild(cell);
    }
    table.appendChild(row);
  }

  document.getElementById("previewContainer").style.display = "block";
}


// 🔹 PROCESSAMENTO DO ARQUIVO (contagem de repetição)
document.getElementById("btnProcess").addEventListener("click", () => {

  // Lê colunas informadas pelo usuário
  const col1 = document.getElementById("col1").value.trim();
  const col2 = document.getElementById("col2").value.trim();
  const dest = document.getElementById("destCol").value.trim();

  const progress = document.getElementById("progress");
  const link = document.getElementById("downloadExcel");

  if (!col1 || !col2 || !dest) {
    alert("Preencha todas as colunas.");
    return;
  }

  // Converte letras para índices (A → 0, B → 1...)
  const col1Index = colLetterToIndex(col1);
  const col2Index = colLetterToIndex(col2);
  const destIndex = colLetterToIndex(dest);

  // Objeto que conta combinações repetidas
  const contador = {};

  // Resultado final da planilha processada
  const resultado = [];

  // ---- PROCESSA CADA LINHA ----
  sheetData.forEach((row, i) => {

    // Aplica formatação em todas as colunas
    const novaLinha = row.map((value, j) =>
      columnFormats[j] === "auto" ? value : formatValue(value, columnFormats[j])
    );

    // Primeira linha = cabeçalho
    if (i === 0) {
      novaLinha[destIndex] = "Repetição";
    } else {
      const val1 = row[col1Index];
      const val2 = row[col2Index];

      if (val1 !== undefined && val1 !== "" && val2 !== undefined && val2 !== "") {

        // Aplica formatação individual
        const f1 = formatValue(val1, columnFormats[col1Index] || "auto");
        const f2 = formatValue(val2, columnFormats[col2Index] || "auto");

        // Cria chave única da combinação
        const chave = `${f1}||${f2}`;

        // Incrementa o contador
        contador[chave] = (contador[chave] || 0) + 1;

        novaLinha[destIndex] = contador[chave];
      } else {
        novaLinha[destIndex] = "";
      }
    }

    resultado.push(novaLinha);
  });

  // Atualiza prévia com o resultado final
  showPreview(resultado);

  // Cria worksheet e workbook
  const worksheet = XLSX.utils.aoa_to_sheet(resultado);
  const workbookOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbookOut, worksheet, "Repetições");

  // Gera arquivo binário
  const wbout = XLSX.write(workbookOut, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);


  // -------------------------------
  // 🔹 NOVA LÓGICA DE NOME DO ARQUIVO
  // -------------------------------

  // Nome do arquivo sem extensão
  const originalName = file ? file.name.replace(/\.xlsx$/i, "") : "Arquivo";

  // Nome base padrão
  let finalName = `${originalName} Enumerado.xlsx`;

  // Se já foi usado, cria Enumerado (2), (3)...
  let index = 2;
  while (generatedNames.has(finalName)) {
    finalName = `${originalName} Enumerado (${index}).xlsx`;
    index++;
  }

  // Marca o nome como usado
  generatedNames.add(finalName);

  // Ajusta o link de download
  link.href = url;
  link.download = finalName;
  link.style.display = "inline";

  progress.textContent = "✅ Planilha processada com sucesso!";
});
