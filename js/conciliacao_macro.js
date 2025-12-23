// Armazenará a planilha processada
let excelMacro = null;

document.getElementById("inputExcelMacro").addEventListener("change", function () {
    if (!this.files.length) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const wsName = wb.SheetNames[0];
        let ws = XLSX.utils.sheet_to_json(wb.Sheets[wsName], { header: 1, defval: "" });

        // === TRANSFORMAÇÃO MACRO ===
        ws[3][0]  = ws[3][2];
        ws[3][4]  = ws[3][3];
        ws[3][7]  = ws[3][8];
        ws[3][11] = ws[3][13];
        ws[3][15] = ws[3][14];
        ws[3][19] = ws[3][18];

        ws.splice(0,3); // remove primeiras linhas

        function removerColunas(inicio,qtd){ ws.forEach(linha=>linha.splice(inicio,qtd)); }
        removerColunas(1,3); removerColunas(2,2); removerColunas(3,3);
        removerColunas(4,3); removerColunas(5,3); removerColunas(6,6);

        ws = ws.filter((linha,i)=>i===0 || linha[3]!=="");
        ws.forEach((linha,i)=>{ if(i!==0){ if(linha[4]===0)linha[4]=""; if(linha[5]===0)linha[5]=""; }});

        excelMacro = ws;

        renderizarTabela(ws, "output");
    };
    reader.readAsArrayBuffer(this.files[0]);
});

// === MESMA FUNÇÃO DE RENDERIZAÇÃO QUE O OUTPUT NORMAL USA ===
function renderizarTabela(ws, outputId) {
    const output = document.getElementById(outputId);
    if (!output) return;

    let html = "";
    const maxCols = Math.max(...ws.map(r => r.length));

    html += "<table><tr>";
    for (let c = 0; c < maxCols; c++) html += `<th data-col="${c}">${String.fromCharCode(65+c)}</th>`;
    html += "</tr>";

    ws.forEach((row,i)=>{
        html += "<tr>";
        for (let c = 0; c < maxCols; c++) {
            const val = row[c] ?? "";
            html += `<td data-col="${c}" data-original="${val}">${val}</td>`;
        }
        html += "</tr>";
    });

    html += "</table><hr>";
    output.innerHTML = html;

    // Não altera sua lógica de formatação de colunas (data/valor)
}
