const worker = new Worker("AsposePDFforJS.js");

worker.onerror = evt => {
  document.getElementById("output").textContent = `Erro: ${evt.message}`;
};

worker.onmessage = evt => {
  if (evt.data === 'ready') {
    console.log("Worker pronto");
  } else if (evt.data.json.errorCode === 0) {
    const link = document.createElement("a");
    link.href = evt.data.json.fileNameResult;
    link.download = "resultado.xlsx";
    link.textContent = "Clique para baixar o Excel";
    document.getElementById("output").appendChild(link);
  } else {
    document.getElementById("output").textContent = `Erro: ${evt.data.json.errorText}`;
  }
};

document.getElementById("convertBtn").addEventListener("click", () => {
  const files = document.getElementById("pdfInput").files;
  for (let i = 0; i < files.length; i++) {
    const reader = new FileReader();
    reader.onload = function(e) {
      const arrayBuffer = e.target.result;
      worker.postMessage({
        operation: 'ConvertPdfToXlsx',
        params: [arrayBuffer],
        json: { fileName: files[i].name }
      });
    };
    reader.readAsArrayBuffer(files[i]);
  }
});