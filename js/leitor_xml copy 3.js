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

      function getTextFromElement(parent, tagNames) {
        for (const tag of tagNames) {
          const el = parent.querySelector(tag);
          if (el && el.textContent.trim() !== '') return el.textContent.trim();
        }
        return '';
      }

      const get = (tags) => {
        if(!Array.isArray(tags)) tags = [tags];
        for(const tag of tags){
          const nodes = xml.getElementsByTagName(tag);
          if(nodes.length && nodes[0].textContent.trim() !== ''){
            return nodes[0].textContent.trim();
          }
        }
        return '';
      };

      const ide = {
        numero: get(['nNF']),
        data: get(['dhEmi', 'dEmi']),
        natOp: get(['natOp']),
        chave: get(['chNFe', 'infNFe']),
        uf: get(['cUF'])
      };

      const emitNode = xml.querySelector('emit') || xml.querySelector('enderEmit') || null;
      const destNode = xml.querySelector('dest') || xml.querySelector('enderDest') || null;

      const emitente = {
        nome: emitNode ? getTextFromElement(emitNode, ['xNome']) : '',
        cnpj: emitNode ? getTextFromElement(emitNode, ['CNPJ']) : '',
        municipio: emitNode ? getTextFromElement(emitNode, ['enderEmit > xMun', 'xMun']) : ''
      };

      const destinatario = {
        nome: destNode ? getTextFromElement(destNode, ['xNome']) : '',
        cnpj: destNode ? getTextFromElement(destNode, ['CNPJ']) : '',
        municipio: destNode ? getTextFromElement(destNode, ['enderDest > xMun', 'xMun']) : ''
      };

      const itens = Array.from(xml.querySelectorAll('det')).map(det => {
        const xProd = det.querySelector('prod > xProd')?.textContent || '';
        const qCom = det.querySelector('prod > qCom')?.textContent || '';
        const vUnCom = det.querySelector('prod > vUnCom')?.textContent || '';
        const vProd = det.querySelector('prod > vProd')?.textContent || '';
        return `<tr><td>${xProd}</td><td>${qCom}</td><td>${fmt(vUnCom)}</td><td>${fmt(vProd)}</td></tr>`;
      }).join('');

      const impostosNota = {};
      const icmsNodes = xml.getElementsByTagName("ICMS");
      Array.from(icmsNodes).forEach(n => {
        const node = n.children[0];
        if(!node) return;
        const tipo = node.tagName;
        const vICMS = Number(node.getElementsByTagName("vICMS")[0]?.textContent || 0);
        if(vICMS > 0){
          impostosNota[tipo] = (impostosNota[tipo]||0)+vICMS;
          total[tipo] = (total[tipo]||0)+vICMS;
        }
      });

      const outrosTags = ['vICMSST','vFCP','vFCPST','vIPI','vPIS','vCOFINS','vISSQN'];
      const nomeMap = {vICMSST:'ICMS-ST', vFCP:'FCP', vFCPST:'FCP-ST', vIPI:'IPI', vPIS:'PIS', vCOFINS:'COFINS', vISSQN:'ISS'};
      outrosTags.forEach(tag => {
        const val = Number(get([tag]) || 0);
        if(val > 0){
          const nome = nomeMap[tag] || tag;
          impostosNota[nome] = val;
          total[nome] = (total[nome] || 0) + val;
        }
      });

      let htmlImpostos = Object.entries(impostosNota)
        .map(([k, v]) => `<tr><td>${k}</td><td>R$ ${fmt(v)}</td></tr>`)
        .join('');

      const html = `
        <div class="danfe">
          <h2>NF-e Nº ${ide.numero || '-'}</h2>
          <p><strong>Chave de Acesso:</strong> ${ide.chave || '-'}</p>
          <p><strong>Data Emissão:</strong> ${ide.data || '-'} | <strong>UF:</strong> ${ide.uf || '-'}</p>
          <p><strong>Natureza da Operação:</strong> ${ide.natOp || '-'}</p>

          <h3>Emitente</h3>
          <p>${emitente.nome || '-'} | CNPJ: ${emitente.cnpj || '-'} | Município: ${emitente.municipio || '-'}</p>

          <h3>Destinatário</h3>
          <p>${destinatario.nome || '-'} | CNPJ: ${destinatario.cnpj || '-'} | Município: ${destinatario.municipio || '-'}</p>

          <h3>Produtos</h3>
          <table>
            <tr><th>Descrição</th><th>Qtde</th><th>Valor Unitário</th><th>Valor Total</th></tr>
            ${itens}
          </table>

          <h3>Impostos</h3>
          <table>
            <tr><th>Imposto</th><th>Valor</th></tr>
            ${htmlImpostos}
          </table>
        </div>
      `;
      document.getElementById("danfe-container").insertAdjacentHTML("beforeend", html);

      processadas++;
      if(processadas === arquivos.length) renderTotaisGerais();
    };
    reader.readAsText(file);
  });
});

function renderTotaisGerais(){
  // Ordena os impostos pelo nome em ordem alfabética
  let impostosHTML = Object.entries(total)
    .filter(([_, v]) => v > 0)
    .sort(([aKey], [bKey]) => aKey.localeCompare(bKey))
    .map(([k, v]) => `<tr><td>${k}</td><td>R$ ${fmt(v)}</td></tr>`)
    .join('');

  // Calcula o total de todos os ICMS
  let totalICMS = Object.entries(total)
    .filter(([k, _]) => k.startsWith('ICMS'))
    .reduce((acc, [_, v]) => acc + v, 0);

  impostosHTML += `<tr><td><strong>Total ICMS</strong></td><td><strong>R$ ${fmt(totalICMS)}</strong></td></tr>`;

  const html = `
    <div class="total-block">
      <h2>Totalização Geral de Todas as Notas</h2>
      <table>
        <tr><th>Imposto</th><th>Total</th></tr>
        ${impostosHTML}
      </table>
    </div>
  `;
  document.getElementById("totais-gerais").innerHTML = html;
}

function fmt(v){
  return Number(v).toLocaleString("pt-BR",{minimumFractionDigits:2, maximumFractionDigits:2});
}
