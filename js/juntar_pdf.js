/* =========================
   FUNÇÃO PARA JUNTAR PDFs
========================= */

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


/* =========================
   FUNÇÃO PARA SEPARAR PDF
========================= */

async function splitPDF() {
  const input = document.getElementById('pdfSplitFile');
  const status = document.getElementById('splitStatus');
  const file = input.files[0];

  if (!file) {
    status.textContent = 'Selecione um arquivo PDF.';
    return;
  }

  const { PDFDocument } = PDFLib;
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await PDFDocument.load(arrayBuffer);

  const zip = new JSZip();
  const baseName = file.name.replace(/\.pdf$/i, '');

  for (let i = 0; i < pdf.getPageCount(); i++) {
    const newPdf = await PDFDocument.create();
    const [copiedPage] = await newPdf.copyPages(pdf, [i]);
    newPdf.addPage(copiedPage);

    const pdfBytes = await newPdf.save();
    zip.file(`${baseName}_Pagina_${i + 1}.pdf`, pdfBytes);
  }

  const zipBlob = await zip.generateAsync({ type: "blob" });
  const url = URL.createObjectURL(zipBlob);

  const link = document.createElement('a');
  link.href = url;
  link.download = `${baseName}_Separado.zip`;
  document.body.appendChild(link);
  link.click();

  setTimeout(() => {
    URL.revokeObjectURL(url);
    document.body.removeChild(link);
  }, 100);

  status.textContent = 'PDF separado com sucesso!';
}
