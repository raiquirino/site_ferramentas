let dadosExcel = [];
let formatosColuna = {};
let formatoSelecionado = "original"; 
let colCombIndex = null; 
let filtroAtual = "todos"; 
let nomeArquivoOriginal = "";

function setFormato(tipo){
    formatoSelecionado = tipo;
}

function setFiltro(valor){
    filtroAtual = valor;
    exibirTabelaExcel(dadosExcel);
}

function parseNumero(valor){
    if(valor == null) return NaN;
    if(typeof valor === "number") return valor;
    valor = valor.toString().trim();
    if(valor === "") return NaN;
    if(valor.match(/^\d{1,3}(\.\d{3})*,\d+$/)) return parseFloat(valor.replace(/\./g,'').replace(',', '.'));
    if(valor.match(/^\d{1,3}(,\d{3})*\.\d+$/)) return parseFloat(valor.replace(/,/g,''));
    return parseFloat(valor.replace(',', '.'));
}

function numeroParaColuna(n){
    let coluna = "";
    while(n >= 0){
        coluna = String.fromCharCode((n % 26) + 65) + coluna;
        n = Math.floor(n/26) - 1;
    }
    return coluna;
}

function colunaParaNumero(letra){
    let col = 0;
    for(let i=0; i<letra.length; i++){
        col *= 26;
        col += letra.charCodeAt(i) - 65 + 1;
    }
    return col-1;
}

function formatarDataPTBR(valor){
    if(!(valor instanceof Date)) return valor;
    const dia = String(valor.getDate()).padStart(2,'0');
    const mes = String(valor.getMonth() + 1).padStart(2,'0');
    const ano = valor.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function formatarValor(valor, colIndex){
    const tipo = formatosColuna[colIndex] || "original";
    if(tipo === "numero"){
        const n = parseNumero(valor);
        return isNaN(n) ? valor : n.toLocaleString("pt-BR",{minimumFractionDigits:2, maximumFractionDigits:2});
    }
    if(tipo === "data"){
        return formatarDataPTBR(valor);
    }
    return valor;
}

function exibirTabelaExcel(dados){
    if(dados.length === 0){
        document.getElementById("conteudo").innerHTML = "<p>Nenhum dado encontrado.</p>";
        return;
    }

    const colunas = dados[0].length;
    let html = "<table border='1'><thead><tr>";

    for(let c=0; c<colunas; c++){
        html += `<th class='col-header' data-col-index='${c}'>${numeroParaColuna(c)}</th>`;
    }
    html += "</tr></thead><tbody>";

    for(let i=0;i<dados.length;i++){
        if(i===0){
            html += "<tr>";
        } else if(colCombIndex !== null){
            const valComb = (dados[i][colCombIndex] || "").toString().toLowerCase();
            if(filtroAtual==="sim" && valComb !== "sim") continue;
            if(filtroAtual==="nao" && valComb === "sim") continue;
            html += "<tr>";
        } else {
            html += "<tr>";
        }

        for(let j=0;j<colunas;j++){
            const val = dados[i][j] ?? "";
            const formato = (typeof val === "number") ? "num" : "text";
            html += `<td class='${formato}'>${formatarValor(val, j)}</td>`;
        }
        html += "</tr>";
    }

    html += "</tbody></table>";
    document.getElementById("conteudo").innerHTML = html;

    document.querySelectorAll(".col-header").forEach(th=>{
        th.addEventListener("click", function(){
            const colIndex = parseInt(this.getAttribute("data-col-index"));
            formatosColuna[colIndex] = formatoSelecionado;
            exibirTabelaExcel(dadosExcel);
        });
    });
}

document.getElementById("inputExcel").addEventListener("change", function(e){
    const arquivo = e.target.files[0];
    if(!arquivo) return;
    nomeArquivoOriginal = arquivo.name.replace(/\.xlsx$/i,''); // salva o nome sem extensão
    const reader = new FileReader();

    reader.onload = function(evt){
        const dados = evt.target.result;
        const workbook = XLSX.read(dados, {type: "binary"});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        dadosExcel = XLSX.utils.sheet_to_json(sheet, {header: 1, raw: true});

        // Corrigir datas e números
        for(let i=1;i<dadosExcel.length;i++){
            const valData = dadosExcel[i][0];
            if(typeof valData === "number"){ 
                const d = XLSX.SSF.parse_date_code(valData);
                if(d) dadosExcel[i][0] = new Date(d.y, d.m-1, d.d);
            }
            const valDebito = dadosExcel[i][3];
            if(typeof valDebito === "string"){ 
                const num = parseNumero(valDebito);
                if(!isNaN(num)) dadosExcel[i][3] = num;
            }
        }

        exibirTabelaExcel(dadosExcel);
    };
    reader.readAsBinaryString(arquivo);
});

document.getElementById("conciliar").addEventListener("click", function(){
    const valorAlvo = parseNumero(document.getElementById("valorAlvo").value);
    const colunaRef = document.getElementById("colunaReferencia").value.toUpperCase();
    const colunaComb = document.getElementById("colunaComb").value.toUpperCase();

    if(isNaN(valorAlvo) || !colunaRef || !colunaComb){ 
        alert("Preencha valor alvo, coluna referência e coluna de combinação!"); 
        return; 
    }

    const colIndex = colunaParaNumero(colunaRef);
    let combIndex = colunaParaNumero(colunaComb);

    if(dadosExcel[0].length <= combIndex){
        for(let i=0;i<dadosExcel.length;i++){
            while(dadosExcel[i].length <= combIndex){
                dadosExcel[i].push("");
            }
        }
    }
    dadosExcel[0][combIndex] = "Combinação";
    colCombIndex = combIndex;

    const valores = [];
    for(let i=1;i<dadosExcel.length;i++){
        const v = dadosExcel[i][colIndex];
        if(typeof v === "number") valores.push({valor: v, linha: i});
    }

    for(let i=1;i<dadosExcel.length;i++) dadosExcel[i][combIndex] = "";

    let encontrada = false;
    function buscarCombinacao(start, combo, soma){
        if(encontrada) return;
        if(Math.abs(soma - valorAlvo) < 0.0001){
            combo.forEach(linha => dadosExcel[linha][combIndex] = "SIM");
            encontrada = true;
            return;
        }
        if(soma > valorAlvo || start >= valores.length) return;

        for(let i=start; i<valores.length; i++){
            combo.push(valores[i].linha);
            buscarCombinacao(i+1, combo, soma + valores[i].valor);
            combo.pop();
        }
    }

    buscarCombinacao(0, [], 0);
    exibirTabelaExcel(dadosExcel);
});

document.getElementById("salvarExcel").addEventListener("click", function(){
    if(dadosExcel.length === 0){
        alert("Nenhum dado para exportar!");
        return;
    }

    const exportData = [];
    exportData.push(dadosExcel[0]);

    for(let i=1;i<dadosExcel.length;i++){
        if(colCombIndex !== null){
            const valComb = (dadosExcel[i][colCombIndex] || "").toString().toLowerCase();
            if(filtroAtual === "sim" && valComb !== "sim") continue;
            if(filtroAtual === "nao" && valComb === "sim") continue;
        }
        exportData.push(dadosExcel[i]);
    }

    const ws = XLSX.utils.aoa_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Filtrado");

    XLSX.writeFile(wb, `${nomeArquivoOriginal}_Filtrado.xlsx`);
});
