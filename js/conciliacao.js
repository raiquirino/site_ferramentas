let workbook, worksheet, data = [];
let formatoSelecionado = null; // Armazena o botão de formatação selecionado

// ===============================
// 📌 CARREGA ARQUIVO EXCEL
// ===============================
document.getElementById('excelFile').addEventListener('change', handleFile);

function handleFile(e) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const dataBinary = new Uint8Array(event.target.result);
        workbook = XLSX.read(dataBinary, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];

        const range = XLSX.utils.decode_range(ws['!ref']);
        worksheet = [];

        for (let R = range.s.r; R <= range.e.r; ++R) {
            const row = [];
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = { c: C, r: R };
                const cellRef = XLSX.utils.encode_cell(cellAddress);
                const cell = ws[cellRef];

                let valor = cell ? cell.v : "";
                row.push(valor);
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
        aplicarFiltro("todos");
        limparTotais();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
}

// ===============================
// 📌 CONVERTE LETRAS ↔ ÍNDICE
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
    const baseIdx = letraParaIndice(document.getElementById('colunaBase').value);
    const alvoIdx = letraParaIndice(document.getElementById('colunaAlvo').value);
    const concIdx = letraParaIndice(document.getElementById('colunaConciliacao').value);

    if (!baseIdx && baseIdx !== 0 || !alvoIdx && alvoIdx !== 0 || !concIdx && concIdx !== 0) {
        alert("Preencha todas as colunas antes de conciliar!");
        return;
    }

    const maxCols = Math.max(...data.map(row => row.length));
    if (concIdx >= maxCols) {
        data[0][concIdx] = "Conciliado";
        for (let i = 1; i < data.length; i++) data[i][concIdx] = "";
    }

    // 🔹 Marca quais linhas devem ser conciliadas
    const conciliados = Array(data.length).fill(false);

    for (let i = 1; i < data.length; i++) {
        if (conciliados[i]) continue;
        const baseVal = data[i][baseIdx];
        if (baseVal === undefined || baseVal === null || baseVal === '') continue;

        for (let j = i + 1; j < data.length; j++) {
            const alvoVal = data[j][alvoIdx];
            if (alvoVal === baseVal) {
                conciliados[i] = true;
                conciliados[j] = true;
                break;
            }
        }
    }

    // 🔹 Atualiza coluna de conciliação
    for (let i = 1; i < data.length; i++) {
        data[i][concIdx] = conciliados[i] ? 'Sim' : data[i][concIdx];
    }

    exibirTabela(data);
    aplicarFiltro(document.querySelector('#filtroBotoes .ativo')?.dataset.filtro || "todos");
    atualizarTotais(baseIdx, alvoIdx, concIdx);

    document.getElementById('baixarBtn').style.display = 'inline-block';
});

// ===============================
// 📌 EXIBE TABELA NO HTML
// ===============================
function exibirTabela(data) {
    const container = document.getElementById('tabelaContainer');
    container.innerHTML = '';

    const table = document.createElement('table');
    table.classList.add('tabela-conciliada');

    const numCols = Math.max(...data.map(row => row.length));

    // Cabeçalho com letras
    const letrasRow = document.createElement('tr');
    for (let j = 0; j < numCols; j++) {
        const th = document.createElement('th');
        th.textContent = indiceParaLetra(j);
        th.style.position = "relative";

        const btnContainer = document.createElement('div');
        btnContainer.style.position = "absolute";
        btnContainer.style.top = "100%";
        btnContainer.style.left = "0";
        btnContainer.style.display = "none";
        btnContainer.style.background = "#fff";
        btnContainer.style.border = "1px solid #ccc";
        btnContainer.style.zIndex = "10";
        btnContainer.style.padding = "2px";

        ['Original','Data','Valor'].forEach(tipo=>{
            const btn = document.createElement('button');
            btn.textContent = tipo;
            btn.style.margin="1px";
            btn.style.fontSize="10px";
            btn.addEventListener('click', e=>{
                e.stopPropagation();
                formatoSelecionado = tipo; // Salva formato selecionado
                alert(`Clique na coluna que deseja aplicar o formato "${tipo}"`);
            });
            btnContainer.appendChild(btn);
        });

        th.addEventListener('click', () => {
            if(formatoSelecionado) {
                atualizarTipoColuna(j, formatoSelecionado);
            }
        });

        th.appendChild(btnContainer);
        th.addEventListener('mouseenter', ()=>btnContainer.style.display="block");
        th.addEventListener('mouseleave', ()=>btnContainer.style.display="none");
        letrasRow.appendChild(th);
    }
    table.appendChild(letrasRow);

    // Linhas de dados
    data.forEach((row,i)=>{
        const tr = document.createElement('tr');
        for(let j=0;j<numCols;j++){
            const td = document.createElement(i===0?'th':'td');
            let valor = row[j]!==undefined ? row[j] : '';

            if(typeof valor === 'number') valor = valor.toLocaleString('pt-BR',{minimumFractionDigits:2, maximumFractionDigits:2});

            td.textContent = valor;
            td.dataset.original = row[j]; 
            tr.appendChild(td);
        }
        table.appendChild(tr);
    });

    container.appendChild(table);
}

// ===============================
// 📌 ALTERAR TIPO DE COLUNA
// ===============================
function atualizarTipoColuna(colIdx, tipo){
    const table = document.querySelector('.tabela-conciliada');
    if(!table) return;

    const linhas = table.querySelectorAll('tr');
    for(let i=1;i<linhas.length;i++){
        const td = linhas[i].querySelectorAll('td,th')[colIdx];
        if(!td) continue;

        const valorOriginal = td.dataset.original ?? td.textContent;

        if(tipo==='Original'){
            td.textContent = valorOriginal;
            td.dataset.tipo='original';
            data[i][colIdx]=valorOriginal;
        }
        else if(tipo==='Data'){
            let dataObj = new Date(valorOriginal);
            if(!isNaN(dataObj)){
                const dia=String(dataObj.getDate()).padStart(2,'0');
                const mes=String(dataObj.getMonth()+1).padStart(2,'0');
                const ano=dataObj.getFullYear();
                td.textContent=`${dia}/${mes}/${ano}`;
                td.dataset.tipo='data';
                data[i][colIdx]=dataObj;
            } else {
                td.textContent=valorOriginal;
                td.dataset.tipo='original';
                data[i][colIdx]=valorOriginal;
            }
        }
        else if(tipo==='Valor'){
            let numStr = valorOriginal.toString().trim();
            if(numStr.includes(',')) numStr = numStr.replace(/\./g,'').replace(',','.');
            let num=parseFloat(numStr);
            if(!isNaN(num)){
                td.textContent=num.toLocaleString('pt-BR',{minimumFractionDigits:2, maximumFractionDigits:2});
                td.dataset.originalFloat=num;
                td.dataset.tipo='valor';
                data[i][colIdx]=num;
            } else {
                td.textContent=valorOriginal;
                td.dataset.tipo='original';
                data[i][colIdx]=valorOriginal;
            }
        }
    }
}

// ===============================
// 📌 EXPORTAR PLANILHA
// ===============================
document.getElementById('baixarBtn').addEventListener('click', () => {
    const table = document.querySelector('.tabela-conciliada');
    if (!table) return;

    const linhas = table.querySelectorAll('tr');
    const dataFiltrada = [];

    linhas.forEach((tr, index) => {
        if (tr.style.display === "none") return; 
        if (index === 0) return; 

        const tds = tr.querySelectorAll('td,th');
        const linhaArray = Array.from(tds).map(td => td.textContent); 
        dataFiltrada.push(linhaArray);
    });

    const ws = XLSX.utils.aoa_to_sheet(dataFiltrada);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Conciliado');
    XLSX.writeFile(wb, 'planilha_conciliada.xlsx');
});

// ===============================
// 📌 FILTRO DE CONCILIAÇÃO
// ===============================
document.querySelectorAll('#filtroBotoes .filtro-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
        document.querySelectorAll('#filtroBotoes .filtro-btn').forEach(b=>b.classList.remove('ativo'));
        btn.classList.add('ativo');
        aplicarFiltro(btn.dataset.filtro);
    });
});

function aplicarFiltro(filtro){
    const concLetra = document.getElementById('colunaConciliacao').value;
    if(!concLetra) return;

    const concIdx = letraParaIndice(concLetra);
    const table = document.querySelector('.tabela-conciliada');
    if(!table) return;

    const linhas = table.querySelectorAll('tr');
    for(let i=2;i<linhas.length;i++){
        const td = linhas[i].querySelectorAll('td')[concIdx];
        const valor = td ? td.textContent.trim().toLowerCase() : "";
        if(filtro==="todos") linhas[i].style.display="";
        else if(filtro==="sim" && valor==="sim") linhas[i].style.display="";
        else if(filtro==="nao" && valor!=="sim") linhas[i].style.display="";
        else linhas[i].style.display="none";
    }

    const baseIdx = letraParaIndice(document.getElementById('colunaBase').value);
    const alvoIdx = letraParaIndice(document.getElementById('colunaAlvo').value);
    atualizarTotais(baseIdx, alvoIdx, concIdx);
}

// ===============================
// 📌 TOTALIZAÇÕES
// ===============================
function atualizarTotais(baseIdx, alvoIdx, concIdx){
    const area = document.getElementById("totaisArea");
    if(!area) return;

    const table = document.querySelector('.tabela-conciliada');
    if(!table) return;

    const linhas = table.querySelectorAll("tr");
    let totalBase=0, totalAlvo=0, totalConc=0;

    for(let i=2;i<linhas.length;i++){
        if(linhas[i].style.display==="none") continue;
        const tds = linhas[i].querySelectorAll("td");
        const baseValStr = tds[baseIdx]?.dataset.original ?? tds[baseIdx]?.textContent;
        const alvoValStr = tds[alvoIdx]?.dataset.original ?? tds[alvoIdx]?.textContent;
        const concVal = tds[concIdx]?.textContent.trim();

        const baseVal=parseFloat(baseValStr.toString().replace(/\./g,'').replace(',','.'));
        const alvoVal=parseFloat(alvoValStr.toString().replace(/\./g,'').replace(',','.'));

        if(!isNaN(baseVal)) totalBase+=baseVal;
        if(!isNaN(alvoVal)) totalAlvo+=alvoVal;
        if(concVal==="Sim") totalConc++;
    }

    area.innerHTML=`
        <div style="display:flex; gap:30px; font-weight:bold; flex-wrap: wrap;">
            <div>Total coluna base (${indiceParaLetra(baseIdx)}): ${totalBase.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
            <div>Total coluna alvo (${indiceParaLetra(alvoIdx)}): ${totalAlvo.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
            <div>Total conciliados: ${totalConc}</div>
        </div>
    `;
}

function limparTotais(){
    const area=document.getElementById("totaisArea");
    if(area) area.innerHTML="";
}
