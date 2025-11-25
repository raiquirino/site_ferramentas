// Função auxiliar: formata data OFX → Excel
function formatDate(raw) {
  const clean = raw.replace(/000000.*$/, "").trim();
  const year = clean.substring(0, 4);
  const month = clean.substring(4, 6);
  const day = clean.substring(6, 8);
  return `${day}/${month}/${year}`;
}

// Converte lista de transações em planilha Excel
function toExcel(transactions) {
  const worksheetData = [
    ["Date", "Memo", "Amount"],
    ...transactions.map(t => [t.date, t.memo, parseFloat(t.amount)])
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Transações");

  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new Blob([wbout], { type: "application/octet-stream" });
}

// OFX → Excel (um arquivo de saída por OFX)
document.getElementById("btnConvertOFX").addEventListener("click", () => {
  const files = document.getElementById("ofxInput").files;
  const progress = document.getElementById("progressOFX");
  const bar = document.getElementById("progressBarOFX");
  const linksContainer = document.getElementById("downloadExcelLinks");
  progress.textContent = "";
  bar.style.width = "0%";
  linksContainer.innerHTML = "";

  if (!files.length) {
    alert("Selecione um ou mais arquivos OFX.");
    return;
  }

  let processed = 0;

  [...files].forEach(file => {
    const reader = new FileReader();
    reader.onload = () => {
      const content = reader.result;
      const matches = content.match(/<STMTTRN>[\s\S]*?<\/STMTTRN>/g) || [];
      const transactions = matches.map(trn => {
        const rawDate = trn.match(/<DTPOSTED>([^<]*)/)?.[1] || '';
        const amount = trn.match(/<TRNAMT>([^<]*)/)?.[1] || '';
        const memo = trn.match(/<MEMO>([^<]*)/)?.[1] || '';
        return { date: formatDate(rawDate), amount, memo };
      });

      const blob = toExcel(transactions);
      const url = URL.createObjectURL(blob);

      // cria link para cada arquivo convertido
      const link = document.createElement("a");
      link.href = url;
      link.download = file.name.replace(/\.ofx$/i, "") + ".xlsx";
      link.textContent = `⬇️ Baixar ${link.download}`;
      link.style.display = "block";
      linksContainer.appendChild(link);

      processed++;
      const percent = Math.floor((processed / files.length) * 100);
      progress.textContent = `🔄 ${percent}% concluído`;
      bar.style.width = `${percent}%`;

      if (processed === files.length) {
        progress.textContent = `✅ Todos os arquivos OFX foram convertidos.`;
        bar.style.width = "100%";
      }
    };
    reader.readAsText(file);
  });
});

// Função auxiliar: formata data Excel → OFX
function formatDateToOFX(raw) {
  if (!raw) return "";

  if (!isNaN(raw)) {
    const baseDate = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(baseDate.getTime() + raw * 86400000);
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const day = String(date.getUTCDate()).padStart(2, "0");
    return `${year}${month}${day}000000[-03:BRT]`;
  }

  const str = raw.toString().trim();
  const parts = str.split(/[\/\-]/);
  if (parts.length === 3) {
    const [d, m, y] = parts;
    return `${y}${m.padStart(2, "0")}${d.padStart(2, "0")}000000[-03:BRT]`;
  }

  if (/^\d{8}$/.test(str)) {
    const d = str.substring(0, 2);
    const m = str.substring(2, 4);
    const y = str.substring(4, 8);
    return `${y}${m}${d}000000[-03:BRT]`;
  }

  return "";
}

// Gera conteúdo OFX a partir de transações
function generateOFX(transactions) {
  const header = `
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
  <BANKMSGSRSV1>
    <STMTTRNRS>
      <STMTRS>
        <BANKTRANLIST>
`;

  const footer = `
        </BANKTRANLIST>
      </STMTRS>
    </STMTTRNRS>
  </BANKMSGSRSV1>
</OFX>
`;

  const body = transactions.map((t, i) => `
          <STMTTRN>
            <TRNTYPE>OTHER
            <DTPOSTED>${t.date}
            <TRNAMT>${t.amount}
            <FITID>${i + 1}
            <MEMO>${t.memo}
          </STMTTRN>
  `).join("\n");

  return header + body + footer;
}

// Excel → OFX (um arquivo de saída por planilha)
document.getElementById("btnConvertCSV").addEventListener("click", () => {
  const files = document.getElementById("csvInput").files;
  const progress = document.getElementById("progressCSV");
  const bar = document.getElementById("progressBarCSV");
  const linksContainer = document.getElementById("downloadOFXLinks");
  progress.textContent = "";
  bar.style.width = "0%";
  linksContainer.innerHTML = "";

  if (!files.length) {
    alert("Selecione uma ou mais planilhas Excel (.xlsx).");
    return;
  }

  let processed = 0;

  [...files].forEach(file => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const transactions = [];
      for (let i = 1; i < rows.length; i++) {
        const [dateRaw, memoRaw, amountRaw] = rows[i];
        if (!dateRaw || !memoRaw || !amountRaw) continue;
        const date = formatDateToOFX(dateRaw);
        const memo = String(memoRaw).trim();
        const amount = String(amountRaw).replace(",", ".").trim();
        transactions.push({ date, amount, memo });
      }

      const ofx = generateOFX(transactions);
      const blob = new Blob([ofx], { type: "text/plain;charset=utf-8;" });
      const url = URL.createObjectURL(blob);

      // cria link para cada arquivo convertido
      const link = document.createElement("a");
      link.href = url;
      link.download = file.name.replace(/\.xlsx$/i, "") + ".ofx";
      link.textContent = `⬇️ Baixar ${link.download}`;
      link.style.display = "block";
      linksContainer.appendChild(link);

      processed++;
      const percent = Math.floor((processed / files.length) * 100);
      progress.textContent = `🔄 ${percent}% concluído`;
      bar.style.width = `${percent}%`;

      if (processed === files.length) {
        progress.textContent = `✅ Todas as planilhas foram convertidas para OFX.`;
        bar.style.width = "100%";
      }
    };
    reader.readAsArrayBuffer(file);
  });
});