async function mergePDFs() {
  const input = document.getElementById('pdfFiles');
  const status = document.getElementById('status');
  const files = input.files;

  if (files.length < 2) {
    status.textContent = 'Selecione pelo menos dois arquivos PDF.';
    return;
  }

  const { PDFDocument } = window.pdfLib;
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

  const link = document.createElement('a');
  link.href = url;
  link.download = 'PDF_Juntado.pdf';
  document.body.appendChild(link); // necessário para Firefox
  link.click();
  document.body.removeChild(link); // remove o link após o clique
  URL.revokeObjectURL(url); // libera a memória

  status.textContent = 'PDFs juntados com sucesso!';
}