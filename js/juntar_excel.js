async function mergeExcels() {
  const input = document.getElementById('excelFiles');
  const files = input.files;

  if (files.length === 0) {
    alert("Selecione pelo menos um arquivo Excel.");
    return;
  }

  let mergedData = [];

  for (let file of files) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    mergedData = mergedData.concat(json);
  }

  const newSheet = XLSX.utils.json_to_sheet(mergedData);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, "PlanilhaUnificada");

  XLSX.writeFile(newWorkbook, "Planilha_Unificada.xlsx");
}