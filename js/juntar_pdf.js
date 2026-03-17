async function mergePDFs() {
  const input = document.getElementById('pdfFiles');
  const status = document.getElementById('status');
  const files = Array.from(input.files);

  if (files.length < 2) {
    status.textContent = 'Selecione pelo menos dois arquivos PDF.';
    return;
  }

  try {
    status.textContent = 'Iniciando junção...';

    const { PDFDocument } = PDFLib;
    const mergedPdf = await PDFDocument.create();

    let totalPages = 0;

    for (let i = 0; i < files.length; i++) {
      status.textContent = `Processando ${i + 1} de ${files.length}...`;

      const buffer = await files[i].arrayBuffer();
      const pdf = await PDFDocument.load(buffer);

      const copiedPages = await mergedPdf.copyPages(
        pdf,
        pdf.getPageIndices()
      );

      copiedPages.forEach(page => mergedPdf.addPage(page));
      totalPages += copiedPages.length;

      // 🔥 Libera memória explicitamente
      await new Promise(resolve => setTimeout(resolve, 0));
    }

    status.textContent = 'Finalizando arquivo...';

    const mergedPdfBytes = await mergedPdf.save({
      useObjectStreams: true,
    });

    const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);

    const baseName = files[0].name.replace(/\.pdf$/i, '');
    const fileName = `${baseName}_Juntado.pdf`;

    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();

    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    status.textContent = `✅ Concluído! ${totalPages} páginas unidas.`;

  } catch (error) {
    console.error(error);
    status.textContent = '❌ Erro: arquivos muito grandes para o navegador.';
  }
}

async function splitPDF() {
  const input = document.getElementById('pdfSplitFile');
  const status = document.getElementById('splitStatus');
  const file = input.files[0];

  if (!file) {
    status.textContent = 'Selecione um arquivo PDF.';
    return;
  }

  try {
    status.textContent = 'Lendo PDF...';

    const { PDFDocument } = PDFLib;
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await PDFDocument.load(arrayBuffer);

    const zip = new JSZip();
    const baseName = file.name.replace(/\.pdf$/i, '');
    const totalPages = pdf.getPageCount();

    for (let i = 0; i < totalPages; i++) {
      status.textContent = `Separando página ${i + 1} de ${totalPages}...`;

      const newPdf = await PDFDocument.create();
      const [page] = await newPdf.copyPages(pdf, [i]);
      newPdf.addPage(page);

      const pdfBytes = await newPdf.save({
        useObjectStreams: true,
      });

      zip.file(`${baseName}_Pagina_${i + 1}.pdf`, pdfBytes);
    }

    status.textContent = 'Compactando arquivos...';

    const zipBlob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 }
    });

    const url = URL.createObjectURL(zipBlob);

    const link = document.createElement('a');
    link.href = url;
    link.download = `${baseName}_Separado.zip`;
    document.body.appendChild(link);
    link.click();

    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    status.textContent = `✅ PDF separado em ${totalPages} arquivos.`;

  } catch (error) {
    console.error(error);
    status.textContent = '❌ Erro ao separar PDF.';
  }
}