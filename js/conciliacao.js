let currentFormat = null;

// -----------------------------
// Funções utilitárias
// -----------------------------
function columnToLetter(n) {
    let s = "";
    while (n >= 0) {
        s = String.fromCharCode((n % 26) + 65) + s;
        n = Math.floor(n / 26) - 1;
    }
    return s;
}

function letterToColIndex(letter) {
    letter = letter.trim().toUpperCase();
    let col = 0;
    for (let i = 0; i < letter.length; i++) {
        col = col * 26 + (letter.charCodeAt(i) - 64);
    }
    return col - 1;
}

function excelDateToDDMMYYYY(v) {
    if (!isNaN(v) && v > 0 && v < 60000) {
        const base = new Date(Date.UTC(1899, 11, 30));
        const d = new Date(base.getTime() + v * 86400000);
        return `${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;
    }
    return v;
}

function formatValue(v) {
    const num = parseFloat(v);
    return isNaN(num) ? v : num.toLocaleString("pt-BR",{minimumFractionDigits:2, maximumFractionDigits:2});
}

// -----------------------------
// Botões Formato
// -----------------------------
document.getElementById("btnDate").onclick = () => currentFormat = "date";
document.getElementById("btnValue").onclick = () => currentFormat = "value";
document.getElementById("btnOriginal").onclick = () => currentFormat = "original";
document.getElementById("btnDesativar").onclick = () => currentFormat = null;

// Reset input
document.getElementById("inputExcel").addEventListener("click", function () {
    this.value = "";
});

// -----------------------------
// Função que aplica o Macro (transformações específicas)
// -----------------------------
function aplicarMacro(ws) {
    ws[3][0]  = ws[3][2];
    ws[3][4]  = ws[3][3];
    ws[3][7]  = ws[3][8];
    ws[3][11] = ws[3][13];
    ws[3][15] = ws[3][14];
    ws[3][19] = ws[3][18];

    ws.splice(0, 3); // remove primeiras linhas

    function removerColunas(inicio, qtd) { ws.forEach(linha => linha.splice(inicio, qtd)); }
    removerColunas(1, 3);
    removerColunas(2, 2);
    removerColunas(3, 3);
    removerColunas(4, 3);
    removerColunas(5, 3);
    removerColunas(6, 6);

    ws = ws.filter((linha, i) => i === 0 || linha[3] !== "");
    ws.forEach((linha, i) => {
        if (i !== 0) {
            if (linha[4] === 0) linha[4] = "";
            if (linha[5] === 0) linha[5] = "";
        }
    });

    return ws;
}

// -----------------------------
// Leitura Excel com opção de Macro
// -----------------------------
document.getElementById("inputExcel").addEventListener("change", function (e) {
    const reader = new FileReader();
    const file = e.target.files[0];

    reader.onload = function (event) {
        const workbook = XLSX.read(new Uint8Array(event.target.result), { type: "array" });
        let html = "";

        workbook.SheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            let rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });

            // ========================
            // Aqui você decide se aplica o macro ou não
            // ========================
            const aplicarMacroFlag = document.getElementById("usarMacro1")?.checked; // checkbox para ativar macro
            if (aplicarMacroFlag) {
                rows = aplicarMacro(rows);
            }

            const maxCols = Math.max(...rows.map(r => r.length));

            let table = "<table><tr>";
            for (let c = 0; c < maxCols; c++) {
                table += `<th data-col="${c}">${columnToLetter(c)}</th>`;
            }
            table += "</tr>";

            rows.forEach(row => {
                table += "<tr>";
                for (let c = 0; c < maxCols; c++) {
                    const val = row[c] ?? "";
                    table += `<td data-col="${c}" data-original="${val}">${val}</td>`;
                }
                table += "</tr>";
            });

            table += "</table><hr>";
            html += table;
        });

        document.getElementById("output").innerHTML = html;

        // Evento para formatar células ao clicar (coluna inteira)
        document.querySelectorAll("td, th").forEach(cell => {
            cell.addEventListener("click", function () {
                if (!currentFormat) return;
                const col = parseInt(this.getAttribute("data-col"));
                const table = this.closest("table");
                const tds = table.querySelectorAll(`td[data-col="${col}"]`);
                tds.forEach(td => {
                    const original = td.getAttribute("data-original");
                    if (currentFormat === "date") td.textContent = excelDateToDDMMYYYY(original);
                    else if (currentFormat === "value") td.textContent = formatValue(original);
                    else td.textContent = original;
                });
                atualizarTotais();
            });
        });

        atualizarTotais();
    };

    reader.readAsArrayBuffer(file);
});

// -----------------------------
// Conciliação
// -----------------------------
document.getElementById("btnConciliar").addEventListener("click", () => {
    const colA = letterToColIndex(document.getElementById("col1").value);
    const colB = letterToColIndex(document.getElementById("col2").value);
    const colR = letterToColIndex(document.getElementById("colR").value);

    document.querySelectorAll("#output table").forEach(table => {
        const rows = table.querySelectorAll("tr");
        let numCols = rows[0].querySelectorAll("th").length;

        if (colR >= numCols) {
            for (let r = 0; r < rows.length; r++) {
                if (r === 0) rows[r].insertAdjacentHTML("beforeend", `<th data-col="${colR}">${columnToLetter(colR)}</th>`);
                else if (r === 1) rows[r].insertAdjacentHTML("beforeend", `<td data-col="${colR}" data-original="Conciliação">Conciliação</td>`);
                else rows[r].insertAdjacentHTML("beforeend", `<td data-col="${colR}" data-original=""></td>`);
            }
        }

        const usadosB = new Set();
        for (let r1 = 2; r1 < rows.length; r1++) {
            const cel1 = rows[r1].querySelector(`td[data-col="${colA}"]`);
            if (!cel1) continue;
            const v1 = cel1.textContent.trim();
            if (!v1) continue;

            for (let r2 = 2; r2 < rows.length; r2++) {
                if (usadosB.has(r2)) continue;
                const cel2 = rows[r2].querySelector(`td[data-col="${colB}"]`);
                if (!cel2) continue;
                const v2 = cel2.textContent.trim();
                if (!v2) continue;

                if (v1 === v2) {
                    usadosB.add(r2);
                    const res1 = rows[r1].querySelector(`td[data-col="${colR}"]`);
                    const res2 = rows[r2].querySelector(`td[data-col="${colR}"]`);
                    res1.textContent = "Sim"; res1.setAttribute("data-original","Sim");
                    res2.textContent = "Sim"; res2.setAttribute("data-original","Sim");
                    break;
                }
            }
        }

        [colA, colB].forEach(col => {
            const tds = table.querySelectorAll(`td[data-col="${col}"]`);
            tds.forEach(td => {
                const original = td.getAttribute("data-original");
                td.textContent = formatValue(original);
            });
        });
    });

    atualizarTotais();
});

// -----------------------------
// Filtros
// -----------------------------
function filtrarTabela(tipo) {
    const colR = letterToColIndex(document.getElementById("colR").value);
    if (colR < 0) return;

    document.querySelectorAll("#output table").forEach(table => {
        const linhas = table.querySelectorAll("tr");
        for (let i = 2; i < linhas.length; i++) {
            const cel = linhas[i].querySelector(`td[data-col="${colR}"]`);
            const valor = cel.textContent.trim().toLowerCase();
            if (tipo === "conciliados") linhas[i].style.display = (valor === "sim") ? "" : "none";
            else if (tipo === "nao") linhas[i].style.display = (valor === "") ? "" : "none";
            else linhas[i].style.display = "";
        }
    });

    atualizarTotais();
}

document.getElementById("btnConciliados").onclick = () => filtrarTabela("conciliados");
document.getElementById("btnNaoConciliados").onclick = () => filtrarTabela("nao");
document.getElementById("btnTodos").onclick = () => filtrarTabela("todos");

// -----------------------------
// Totalização
// -----------------------------
function atualizarTotais() {
    const colA = letterToColIndex(document.getElementById("col1").value);
    const colB = letterToColIndex(document.getElementById("col2").value);
    let totalA = 0, totalB = 0;
    const letraA = document.getElementById("col1").value.toUpperCase();
    const letraB = document.getElementById("col2").value.toUpperCase();

    document.querySelectorAll("#output table").forEach(table => {
        const linhas = table.querySelectorAll("tr");
        for (let i = 2; i < linhas.length; i++) {
            if (linhas[i].style.display === "none") continue;

            const celA = linhas[i].querySelector(`td[data-col="${colA}"]`);
            const celB = linhas[i].querySelector(`td[data-col="${colB}"]`);

            if (celA) {
                let txt = celA.textContent.trim().replace(/\./g,"").replace(",",".");

                const n = parseFloat(txt); if (!isNaN(n)) totalA += n;
            }
            if (celB) {
                let txt = celB.textContent.trim().replace(/\./g,"").replace(",",".");

                const n = parseFloat(txt); if (!isNaN(n)) totalB += n;
            }
        }
    });

    document.getElementById("totais").innerHTML =
        `Total Coluna ${letraB}: ${totalB.toLocaleString("pt-BR",{minimumFractionDigits:2, maximumFractionDigits:2})} &nbsp;&nbsp; | &nbsp;&nbsp; Total Coluna ${letraA}: ${totalA.toLocaleString("pt-BR",{minimumFractionDigits:2, maximumFractionDigits:2})}`;
}

// -----------------------------
// Salvar Excel (SEM LINHA 0)
// -----------------------------
document.getElementById("btnSalvar").addEventListener("click", () => {
    const tables = document.querySelectorAll("#output table");
    if (tables.length === 0) return;

    tables.forEach((table) => {
        const workbook = XLSX.utils.book_new();
        let data = [];
        const linhas = table.querySelectorAll("tr");

        for (let i = 1; i < linhas.length; i++) {
            if (i > 1 && linhas[i].style.display === "none") continue;

            const cells = linhas[i].querySelectorAll("th, td");
            let row = [];

            cells.forEach(cell => {
                let val = cell.textContent;
                const colIndex = parseInt(cell.getAttribute("data-col"));

                const col1Index = letterToColIndex(document.getElementById("col1").value);
                const col2Index = letterToColIndex(document.getElementById("col2").value);

                if (i > 1 && (colIndex === col1Index || colIndex === col2Index)) {
                    val = val.replace(/\./g, "").replace(",", ".");
                    val = parseFloat(val);
                    if (isNaN(val)) val = "";
                }

                row.push(val);
            });

            data.push(row);
        }

        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, ws, "Sheet1");

        let arquivoOriginal = document.getElementById("inputExcel").files[0]?.name || "planilha";
        let nomeArquivo = arquivoOriginal.replace(/\.[^/.]+$/, "") + ".xlsx";

        XLSX.writeFile(workbook, nomeArquivo);
    });
});

// -----------------------------
// Marcar/desmarcar linha inteira SOMENTE com CTRL + clique
// -----------------------------
document.querySelector("#output").addEventListener("click", function (e) {
    if (!e.ctrlKey) return; // BLOQUEIA clique normal

    const target = e.target;
    if (target.tagName === "TD") {
        e.preventDefault();
        e.stopPropagation();

        const row = target.parentElement;
        row.classList.toggle("selected-row");
    }
});

// -----------------------------
// Copiar célula com duplo clique (NÃO marca linha)
// -----------------------------
document.querySelector("#output").addEventListener("dblclick", function (e) {
    const target = e.target;
    if (target.tagName === "TD") {
        e.preventDefault();
        e.stopPropagation();

        const texto = target.textContent.trim();
        if (texto) {
            navigator.clipboard.writeText(texto).then(() => {
                const msg = document.getElementById("copiado-msg");
                if (msg) {
                    msg.textContent = `Copiado: ${texto}`;
                    msg.style.display = "block";
                    setTimeout(() => msg.style.display = "none", 1500);
                }
            });
        }
    }
});
