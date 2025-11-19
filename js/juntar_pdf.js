async function mergePDFs() {
  const input = document.getElementById('pdfFiles');
  const status = document.getElementById('status');
  const files = input.files;

  if (files.length < 2) {
    status.textContent = 'Selecione pelo menos dois arquivos PDF.';
    return;
  }

  const { PDFDocument } = PDFLib;
  const mergedPdf = await PDFDocument.create();

  for (const file of files) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await PDFDocument.load(arrayBuffer);
    const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    copiedPages.forEach(page => mergedPdf.addPage(page));
  }

  const mergedPdfBytes = await mergedPdf.save();
  const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);

  // Gera nome baseado no primeiro arquivo
  const baseName = files[0].name.replace(/\.pdf$/i, '');
  const fileName = `${baseName}_Juntado.pdf`;

  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();

  setTimeout(() => {
    URL.revokeObjectURL(url);
    document.body.removeChild(link);
  }, 100);

  status.textContent = `PDFs juntados com sucesso como "${fileName}"!`;
}