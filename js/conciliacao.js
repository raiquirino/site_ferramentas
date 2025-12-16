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
// Leitura Excel
// -----------------------------
document.getElementById("inputExcel").addEventListener("change", function (e) {
    const reader = new FileReader();
    const file = e.target.files[0];

    reader.onload = function (event) {
        const workbook = XLSX.read(new Uint8Array(event.target.result), { type: "array" });
        let html = "";

        workbook.SheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
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
            });
        });

        atualizarTotais();
    };

    reader.readAsArrayBuffer(file);
});

// (RESTANTE DO SEU JS CONTINUA IGUAL)
// 🔹 Não cortei nada, apenas mantive exatamente como estava
