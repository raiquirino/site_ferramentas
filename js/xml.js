document.addEventListener('DOMContentLoaded', function () {
  const input = document.getElementById('xmlInput');
  const output = document.getElementById('output');

  input.addEventListener('change', function (event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
      const xmlText = e.target.result;
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlText, 'text/xml');

      const format = (el, tag) => el?.querySelector(tag)?.textContent || '';

      const infNFe = xmlDoc.querySelector('infNFe');
      const ide = xmlDoc.querySelector('ide');
      const emit = xmlDoc.querySelector('emit');
      const dest = xmlDoc.querySelector('dest');
      const total = xmlDoc.querySelector('ICMSTot');
      const transp = xmlDoc.querySelector('transporta');
      const vol = xmlDoc.querySelector('vol');
      const infAdic = xmlDoc.querySelector('infAdic');
      const produtos = xmlDoc.querySelectorAll('det');

      const chave = infNFe?.getAttribute('Id')?.replace(/^NFe/, '') || '';

      const html = `
        <div class="danfe">
          <div class="danfe-header">
            <h2>DANFE - Documento Auxiliar da NF-e</h2>
            <p><strong>Chave de Acesso:</strong> ${chave}</p>
            <p><strong>Número:</strong> ${format(ide, 'nNF')} | <strong>Série:</strong> ${format(ide, 'serie')}</p>
            <p><strong>Data de Emissão:</strong> ${format(ide, 'dhEmi')} | <strong>Saída:</strong> ${format(ide, 'dhSaiEnt')}</p>
            <p><strong>Natureza da Operação:</strong> ${format(ide, 'natOp')}</p>
          </div>

          <div class="danfe-section">
            <h3>Emitente</h3>
            <p><strong>Nome:</strong> ${format(emit, 'xNome')}</p>
            <p><strong>CNPJ:</strong> ${format(emit, 'CNPJ')}</p>
            <p><strong>IE:</strong> ${format(emit, 'IE')}</p>
            <p><strong>Endereço:</strong> ${format(emit, 'xLgr')}, ${format(emit, 'nro')} - ${format(emit, 'xBairro')}, ${format(emit, 'xMun')} - ${format(emit, 'UF')}</p>
          </div>

          <div class="danfe-section">
            <h3>Destinatário</h3>
            <p><strong>Nome:</strong> ${format(dest, 'xNome')}</p>
            <p><strong>CNPJ:</strong> ${format(dest, 'CNPJ')}</p>
            <p><strong>IE:</strong> ${format(dest, 'IE')}</p>
            <p><strong>Endereço:</strong> ${format(dest, 'xLgr')}, ${format(dest, 'nro')} - ${format(dest, 'xBairro')}, ${format(dest, 'xMun')} - ${format(dest, 'UF')}</p>
          </div>

          <div class="danfe-section">
            <h3>Produtos</h3>
            <table>
              <thead>
                <tr><th>Descrição</th><th>Qtd</th><th>Unid</th><th>Valor Unit</th><th>Total</th></tr>
              </thead>
              <tbody>
                ${Array.from(produtos).map(prod => {
                  const prodEl = prod.querySelector('prod');
                  return `<tr>
                    <td>${format(prodEl, 'xProd')}</td>
                    <td>${format(prodEl, 'qCom')}</td>
                    <td>${format(prodEl, 'uCom')}</td>
                    <td>${format(prodEl, 'vUnCom')}</td>
                    <td>${format(prodEl, 'vProd')}</td>
                  </tr>`;
                }).join('')}
              </tbody>
            </table>
          </div>

          <div class="danfe-section">
            <h3>Totais</h3>
            <p><strong>Base ICMS:</strong> R$ ${format(total, 'vBC')}</p>
            <p><strong>ICMS:</strong> R$ ${format(total, 'vICMS')}</p>
            <p><strong>Valor Produtos:</strong> R$ ${format(total, 'vProd')}</p>
            <p><strong>Frete:</strong> R$ ${format(total, 'vFrete')}</p>
            <p><strong>Seguro:</strong> R$ ${format(total, 'vSeg')}</p>
            <p><strong>Desconto:</strong> R$ ${format(total, 'vDesc')}</p>
            <p><strong>Outras Despesas:</strong> R$ ${format(total, 'vOutro')}</p>
            <p><strong>Valor Total da Nota:</strong> R$ ${format(total, 'vNF')}</p>
            <p><strong>Valor Aproximado Tributos:</strong> R$ ${format(total, 'vTotTrib')}</p>
          </div>

          <div class="danfe-section">
            <h3>Transportador</h3>
            <p><strong>Nome:</strong> ${format(transp, 'xNome')}</p>
            <p><strong>CNPJ:</strong> ${format(transp, 'CNPJ')}</p>
            <p><strong>Endereço:</strong> ${format(transp, 'xEnd')}</p>
            <p><strong>Placa:</strong> ${format(transp, 'placa')} | <strong>UF:</strong> ${format(transp, 'UF')}</p>
            <p><strong>Peso Bruto:</strong> ${format(vol, 'pesoB')} | <strong>Peso Líquido:</strong> ${format(vol, 'pesoL')}</p>
          </div>

          <div class="danfe-section">
            <h3>Informações Adicionais</h3>
            <p>${format(infAdic, 'infCpl')}</p>
          </div>
        </div>
      `;

      output.innerHTML = html;
    };

    reader.readAsText(file);
  });
});