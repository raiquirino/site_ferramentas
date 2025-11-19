// Carrega PDF.js dinamicamente
const pdfScript = document.createElement('script');
pdfScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js';
pdfScript.onload = () => {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
  iniciarTabelaPDF();
};
document.head.appendChild(pdfScript);

// Carrega SheetJS dinamicamente
const xlsxScript = document.createElement('script');
xlsxScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
document.head.appendChild(xlsxScript);

function iniciarTabelaPDF() {
  const fileInput = document.getElementById('fileInput');
  const tabela = document.getElementById('tabelaExtrato');
  const thead = tabela.querySelector('thead');
  const tbody = tabela.querySelector('tbody');
  const btnExport = document.getElementById('btnExport');

  fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function () {
      const typedArray = new Uint8Array(reader.result);
      const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;

      thead.innerHTML = '';
      tbody.innerHTML = '';

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();

        // Agrupar por linha (y)
        const linhas = {};
        content.items.forEach(item => {
          const y = Math.round(item.transform[5]);
          if (!linhas[y]) linhas[y] = [];
          linhas[y].push({ x: item.transform[4], str: item.str });
        });

        // Ordenar linhas por y (de cima para baixo)
        const linhasOrdenadas = Object.keys(linhas)
          .map(y => parseInt(y))
          .sort((a, b) => b - a);

        linhasOrdenadas.forEach(y => {
          const linha = linhas[y];
          linha.sort((a, b) => a.x - b.x); // ordenar da esquerda para a direita

          const tr = document.createElement('tr');
          linha.forEach(item => {
            const td = document.createElement('td');
            td.textContent = item.str;
            tr.appendChild(td);
          });
          tbody.appendChild(tr);
        });
      }
    };

    reader.readAsArrayBuffer(file);
  });

  btnExport.addEventListener('click', () => {
    if (typeof XLSX === 'undefined') {
      alert('A biblioteca XLSX ainda está carregando. Tente novamente em instantes.');
      return;
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(tabela);
    XLSX.utils.book_append_sheet(wb, ws, 'PDF');
    XLSX.writeFile(wb, 'conteudo_pdf.xlsx');
  });
}