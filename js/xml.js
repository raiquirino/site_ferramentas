const xmlInput = document.getElementById('xmlFiles');
const outputDiv = document.getElementById('output');
const pdfBtn = document.getElementById('pdfBtn');
let allContent = '';

xmlInput.addEventListener('change', () => {
  allContent = '';
  outputDiv.innerHTML = '';
  const files = xmlInput.files;

  if (files.length === 0) return;

  Array.from(files).forEach((file, index) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      const parser = new DOMParser();
      const xml = parser.parseFromString(e.target.result, 'text/xml');
      const ns = 'http://www.abrasf.org.br/nfse.xsd';

      const get = (tag, parent) =>
        parent?.getElementsByTagNameNS(ns, tag)[0]?.textContent?.trim() || '';

      const infNfse = xml.getElementsByTagNameNS(ns, 'InfNfse')[0];
      const declaracao = xml.getElementsByTagNameNS(ns, 'InfDeclaracaoPrestacaoServico')[0];
      const servico = declaracao?.getElementsByTagNameNS(ns, 'Servico')[0];
      const valores = servico?.getElementsByTagNameNS(ns, 'Valores')[0];
      const prestador = declaracao?.getElementsByTagNameNS(ns, 'Prestador')[0];
      const tomador = declaracao?.getElementsByTagNameNS(ns, 'TomadorServico')[0];
      const prestadorServico = infNfse?.getElementsByTagNameNS(ns, 'PrestadorServico')[0];
      const tomadorEndereco = tomador?.getElementsByTagNameNS(ns, 'Endereco')[0];
      const prestadorEndereco = prestadorServico?.getElementsByTagNameNS(ns, 'Endereco')[0];

      const notaDiv = document.createElement('div');
      notaDiv.className = 'nota';

      const header = document.createElement('div');
      header.className = 'nota-header';
      header.textContent = `Arquivo ${index + 1}: ${file.name}`;
      notaDiv.appendChild(header);

      const createSection = (title, lines, twoColumns = false) => {
      const validLines = lines.filter(line => {
          if (!line) return false;
          const parts = line.split(':');
          if (parts.length < 2) return false;
          const value = parts[1].trim().toLowerCase();
          return value &&
          value !== '-' &&
          value !== '0' &&
          value !== '0.00' &&
          value !== 'r$ 0.00' &&
          value !== 'não' &&
          value !== 'não informado';
    });

      if (validLines.length === 0) return null;

      const section = document.createElement('div');
      section.className = 'nota-section';
      if (twoColumns) section.classList.add('two-columns');

      const h3 = document.createElement('h3');
      h3.textContent = title;
      section.appendChild(h3);

      validLines.forEach(line => {
          const p = document.createElement('p');
          p.textContent = line;
          section.appendChild(p);
    });

      return section;
    };

      const pis = get('ValorPis', valores);
      const cofins = get('ValorCofins', valores);
      const inss = get('ValorInss', valores);
      const ir = get('ValorIr', valores);
      const csll = get('ValorCsll', valores);

      const pisCofinsInssIrCsll = [pis, cofins, inss, ir, csll].every(v =>
        ['0', '0.00', 'R$ 0.00', '-', '', null].includes(v?.trim())
      )
        ? null
        : `Valor PIS / COFINS / INSS / IR / CSLL: R$ ${pis} / ${cofins} / ${inss} / ${ir} / ${csll}`;

      const descontoIncond = get('DescontoIncondicionado', valores);
      const descontoCond = get('DescontoCondicionado', valores);

      const descontoLinha = ['0', '0.00', 'R$ 0.00', '-', '', null].includes(descontoIncond?.trim()) &&
        ['0', '0.00', 'R$ 0.00', '-', '', null].includes(descontoCond?.trim())
        ? null
        : `Desconto Incondicionado / Condicionado: R$ ${descontoIncond} / ${descontoCond}`;




      const sections = [
        createSection('Prestador de Serviço', [
          `Razão Social: ${get('RazaoSocial', prestadorServico)}`,
          `CNPJ: ${get('Cnpj', prestador?.getElementsByTagNameNS(ns, 'CpfCnpj')[0])}`
        ], true),
        createSection('Tomador de Serviço', [
          `Razão Social: ${get('RazaoSocial', tomador)}`,
          `CNPJ: ${get('Cnpj', tomador?.getElementsByTagNameNS(ns, 'CpfCnpj')[0])}`
        ], true),
        createSection('Dados da Nota', [
          `Número: ${get('Numero', infNfse)}`,
          `Data de Emissão: ${get('DataEmissao', infNfse)}`
        ], true),
        createSection('Valores', [
          `Valor dos Serviços: R$ ${get('ValorServicos', valores)}`,
          `Valor Líquido da Nota: R$ ${get('ValorLiquidoNfse', infNfse)}`,
          `Base de Cálculo: R$ ${get('BaseCalculo', infNfse?.getElementsByTagNameNS(ns, 'ValoresNfse')[0])}`,
          `Valor ISS: R$ ${get('ValorIss', valores)}`,
          `Alíquota: ${get('Aliquota', valores)}%`,
          `Valor Crédito: R$ ${get('ValorCredito', infNfse)}`,
          `Valor Deduções: R$ ${get('ValorDeducoes', valores)}`,
        pisCofinsInssIrCsll,
          `Outras Retenções: R$ ${get('OutrasRetencoes', valores)}`,
        descontoLinha
        ], true),

        createSection('Serviço', [
          `Discriminação: ${get('Discriminacao', servico)}`,
          /*`Descrição Tributação Município: ${get('DescricaoCodigoTributacaoMunicípio', infNfse)}`,
          `Item Lista Serviço: ${get('ItemListaServico', servico)}`,
          `Código CNAE: ${get('CodigoCnae', servico)}`,
          `Código Tributação Município: ${get('CodigoTributacaoMunicipio', servico)}`,
          `Exigibilidade ISS: ${get('ExigibilidadeISS', servico)}`,
          `Município Incidência: ${get('MunicipioIncidencia', servico)}`,
          `ISS Retido: ${get('IssRetido', servico)}`*/
        ]),

        createSection('Endereço do Prestador', [
          /*`Logradouro: ${get('Endereco', prestadorEndereco)}`,
          `Bairro: ${get('Bairro', prestadorEndereco)}`,
          `Município: ${get('CodigoMunicipio', prestadorEndereco)}`,
          `UF: ${get('Uf', prestadorEndereco)}`,
          `CEP: ${get('Cep', prestadorEndereco)}`,*/
          `Telefone: ${get('Telefone', prestadorServico?.getElementsByTagNameNS(ns, 'Contato')[0])}`,
          `Email: ${get('Email', prestadorServico?.getElementsByTagNameNS(ns, 'Contato')[0])}`
        ], true),
        createSection('Endereço do Tomador', [
          /*`Logradouro: ${get('Endereco', tomadorEndereco)}`,
          `Bairro: ${get('Bairro', tomadorEndereco)}`,
          `Município: ${get('CodigoMunicipio', tomadorEndereco)}`,
          `UF: ${get('Uf', tomadorEndereco)}`,
          `CEP: ${get('Cep', tomadorEndereco)}`,*/
          `Email: ${get('Email', tomador?.getElementsByTagNameNS(ns, 'Contato')[0])}`
        ], true),
        /*createSection('Órgão Gerador', [
          `Código do Município: ${get('CodigoMunicipio', infNfse?.getElementsByTagNameNS(ns, 'OrgaoGerador')[0])}`,
          `UF: ${get('Uf', infNfse?.getElementsByTagNameNS(ns, 'OrgaoGerador')[0])}`
        ]),*/
        
        createSection('Outros', [
          `Optante Simples Nacional: ${get('OptanteSimplesNacional', declaracao) === '1' ? 'Sim' : 'Não'}`,
          `Incentivo Fiscal: ${get('IncentivoFiscal', declaracao) === '1' ? 'Sim' : 'Não'}`
        ],true)
      ];

      sections.forEach(section => {
        if (section) notaDiv.appendChild(section);
      });

      outputDiv.appendChild(notaDiv);
      pdfBtn.style.display = 'inline-block';
    };
    reader.readAsText(file);
  });
});

function generatePDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('p', 'mm', 'a4');

  html2canvas(outputDiv, {
    scale: 2,
    useCORS: true
  }).then(canvas => {
    const imgData = canvas.toDataURL('image/png');
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const imgWidth = pageWidth;
    const imgHeight = canvas.height * imgWidth / canvas.width;

    if (imgHeight > pageHeight) {
      const totalPages = Math.ceil(imgHeight / pageHeight);
      for (let i = 0; i < totalPages; i++) {
        const sourceY = i * canvas.height / totalPages;
        const sourceHeight = canvas.height / totalPages;

        const pageCanvas = document.createElement('canvas');
        pageCanvas.width = canvas.width;
        pageCanvas.height = sourceHeight;

        const ctx = pageCanvas.getContext('2d');
        ctx.drawImage(
          canvas,
          0,
          sourceY,
          canvas.width,
          sourceHeight,
          0,
          0,
          canvas.width,
          sourceHeight
        );

        const pageImgData = pageCanvas.toDataURL('image/png');
        if (i > 0) doc.addPage();
        doc.addImage(pageImgData, 'PNG', 0, 0, imgWidth, pageHeight);
      }
    } else {
      doc.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
    }

    doc.save('nfse_visual.pdf');
  });
}