let arquivos = [];
let total = {};

document.getElementById('fileInput').addEventListener('change', e => {
  arquivos = [...e.target.files];
});

document.getElementById("btnLimpar").addEventListener("click", () => {
  arquivos = [];
  total = {};
  document.getElementById('fileInput').value = "";
  document.getElementById("totais-gerais").innerHTML = "";
  document.getElementById("danfe-container").innerHTML = "";
});

document.getElementById("btnImprimir").addEventListener("click", () => window.print());

document.getElementById("btnRender").addEventListener("click", () => {
  document.getElementById("totais-gerais").innerHTML = "";
  document.getElementById("danfe-container").innerHTML = "";
  total = {};
  let processadas = 0;

  arquivos.forEach(file => {
    const reader = new FileReader();
    reader.onload = e => {
      const xml = new DOMParser().parseFromString(e.target.result, "application/xml");

      function getText(parent, tags) {
        if (!Array.isArray(tags)) tags = [tags];
        for (const tag of tags) {
          const node = parent.querySelector(tag);
          if (node && node.textContent.trim() !== '') return node.textContent.trim();
        }
        return '';
      }

      // ----------------- BLOCO IDE -----------------
      const ideNode = xml.querySelector('ide');
      const ideTags = ['cUF', 'cNF', 'natOp', 'mod', 'serie', 'nNF', 'dhEmi', 'dhSaiEnt', 'tpNF', 'idDest', 'cMunFG', 'tpImp', 'tpEmis', 'cDV', 'tpAmb', 'finNFe', 'indFinal', 'indPres', 'indIntermed', 'procEmi', 'verProc', 'NFref', 'refNFe'];
      const ide = {};
      ideTags.forEach(tag => ide[tag] = getText(ideNode, tag));

      // ----------------- BLOCO EMITENTE -----------------
      const emitNode = xml.querySelector('emit');
      const emitTags = ['CNPJ', 'xNome', 'xLgr', 'nro', 'xCpl', 'xBairro', 'cMun', 'xMun', 'UF', 'CEP', 'cPais', 'xPais', 'fone', 'IE', 'CRT'];
      const emitente = {};
      emitTags.forEach(tag => emitente[tag] = getText(emitNode, tag));

      // ----------------- BLOCO DESTINATÁRIO -----------------
      const destNode = xml.querySelector('dest');
      const destTags = ['CNPJ', 'xNome', 'xLgr', 'nro', 'xCpl', 'xBairro', 'cMun', 'xMun', 'UF', 'CEP', 'cPais', 'xPais', 'indIEDest', 'IE'];
      const destinatario = {};
      destTags.forEach(tag => destinatario[tag] = getText(destNode, tag));

      // ----------------- BLOCO PRODUTOS -----------------
      const detNodes = Array.from(xml.querySelectorAll('det'));
      const itensHTML = detNodes.map(det => {
        const prodNode = det.querySelector('prod');
        const prodTags = ['cProd', 'cEAN', 'xProd', 'NCM', 'CEST', 'CFOP', 'uCom', 'qCom', 'vUnCom', 'vProd', 'cEANTrib', 'uTrib', 'qTrib', 'vUnTrib', 'indTot', 'xPed'];
        const prod = {};
        prodTags.forEach(tag => prod[tag] = getText(prodNode, tag));

        // No detalhe do produto, não vamos extrair impostos porque vamos pegar do total da NF-e

        // HTML do produto
        let prodHTML = `<tr><td>${prod.xProd || '-'}</td><td>${prod.qCom || '-'}</td><td>${prod.vUnCom || '-'}</td><td>${prod.vProd || '-'}</td></tr>`;

        return `<h4>Produto: ${prod.xProd || '-'}</h4>
                <table>
                  <tr><th>Descrição</th><th>Qtde</th><th>Valor Unit.</th><th>Valor Total</th></tr>
                  ${prodHTML}
                </table>`;
      }).join('');

      // ----------------- BLOCO TOTAL (ICMSTot) -----------------
      const totalNode = xml.querySelector('ICMSTot');
      const impostoTags = [
        'vBC',        // Base de cálculo do ICMS
        'vICMS',      // Valor do ICMS
        'vBCST',      // Base de cálculo ICMS Substituição Tributária
        'vST',        // Valor do ICMS Substituição Tributária
        'vII',        // Valor do Imposto de Importação
        'vFCP',       // Valor do FCP UF Destino
        'vFrete',     // Valor do Frete
        'vSeg',       // Valor do Seguro
        'vDesc',      // Desconto
        'vOutro',     // Outras despesas acessórias
        'vIPI',       // Valor total do IPI
        'vPIS',       // Valor do PIS
        'vCOFINS',    // Valor da COFINS
        'vProd',      // Valor total dos produtos
        'vTotTrib',   // Valor total aproximado dos tributos
        'vNF'         // Valor total da nota fiscal
      ];
      const totais = {};

      if (totalNode) {
        impostoTags.forEach(tag => {
          const valStr = getText(totalNode, tag);
          const valNum = Number(valStr.replace(',', '.') || 0);
          totais[tag] = valStr;
          if (valNum > 0) {
            total[tag] = (total[tag] || 0) + valNum;
          }
        });
      }

      // ----------------- HTML da NF-e -----------------
      const html = `
        <div class="danfe">
          <h2>NF-e Nº ${ide.nNF || '-'}</h2>
          <p><strong>Chave:</strong> ${getText(xml, 'chNFe')}</p>
          <h3>Emitente</h3>
          <p>${emitente.xNome || '-'} | CNPJ: ${emitente.CNPJ || '-'} | UF: ${emitente.UF || '-'}</p>
          <h3>Destinatário</h3>
          <p>${destinatario.xNome || '-'} | CNPJ: ${destinatario.CNPJ || '-'} | UF: ${destinatario.UF || '-'}</p>
          <h3>Produtos</h3>
          ${itensHTML}
          <h3>Totais da Nota</h3>
          <table>
            <tr><th>Descrição</th><th>Valor</th></tr>
            ${impostoTags.map(tag => {
              if (!totais[tag]) return '';
              const valNum = Number(totais[tag].replace(',', '.'));
              if (valNum === 0) return '';
              return `<tr><td>${tag}</td><td>R$ ${Number(valNum).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td></tr>`;
            }).join('')}
          </table>
        </div>
      `;

      document.getElementById("danfe-container").insertAdjacentHTML("beforeend", html);

      processadas++;
      if (processadas === arquivos.length) renderTotaisGerais();
    };
    reader.readAsText(file);
  });
});

function renderTotaisGerais() {
  const impostoTags = [
    'vBC',
    'vICMS',
    'vBCST',
    'vST',
    'vII',
    'vFCP',
    'vFrete',
    'vSeg',
    'vDesc',
    'vOutro',
    'vIPI',
    'vPIS',
    'vCOFINS',
    'vProd',
    'vTotTrib',
    'vNF'
  ];

  let impostosHTML = impostoTags
    .filter(tag => total[tag] && total[tag] > 0)
    .map(tag => `<tr><td>${tag}</td><td>R$ ${fmt(total[tag])}</td></tr>`)
    .join('');

  const html = `<div class="total-block">
                  <h2>Totalização Geral de Todas as Notas</h2>
                  <table>
                    <tr><th>Imposto</th><th>Total</th></tr>
                    ${impostosHTML}
                  </table>
                </div>`;
  document.getElementById("totais-gerais").innerHTML = html;
}

function fmt(v) {
  return Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
