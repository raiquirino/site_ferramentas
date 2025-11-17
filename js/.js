// OFX → Excel (.xlsx)
function formatDate(raw) {
  const clean = raw.replace(/000000.*$/, "").trim();
  const year = clean.substring(0, 4);
  const month = clean.substring(4, 6);
  const day = clean.substring(6, 8);
  return `${day}/${month}/${year}`;
}

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

document.getElementById("btnConvertOFX").addEventListener("click", () => {
  const file = document.getElementById("ofxInput").files[0];
  const progress = document.getElementById("progressOFX");
  const bar = document.getElementById("progressBarOFX");
  const link = document.getElementById("downloadExcel");
  progress.textContent = "";
  bar.style.width = "0%";
  link.style.display = "none";

  if (!file) {
    alert("Selecione um arquivo OFX.");
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    const content = reader.result;
    const matches = content.match(/<STMTTRN>[\s\S]*?<\/STMTTRN>/g);
    const total = matches ? matches.length : 0;

    if (total === 0) {
      progress.textContent = "Nenhuma transação encontrada.";
      return;
    }

    const transactions = [];
    let index = 0;

    const processNext = () => {
      if (index >= total) {
        const blob = toExcel(transactions);
        const url = URL.createObjectURL(blob);

        link.href = url;
        link.download = "transacoes.xlsx";
        link.style.display = "inline";
        link.click();

        progress.textContent = `✅ ${total} transações processadas e salvas.`;
        bar.style.width = "100%";
        return;
      }

      const trn = matches[index];
      const rawDate = trn.match(/<DTPOSTED>([^<]*)/)?.[1] || '';
      const amount = trn.match(/<TRNAMT>([^<]*)/)?.[1] || '';
      const memo = trn.match(/<MEMO>([^<]*)/)?.[1] || '';
      const date = formatDate(rawDate);
      transactions.push({ date, amount, memo });

      index++;
      const percent = Math.floor((index / total) * 100);
      progress.textContent = `🔄 ${percent}% concluído`;
      bar.style.width = `${percent}%`;

      setTimeout(processNext, 10);
    };

    processNext();
  };

  reader.readAsText(file);
});

// Excel (.xlsx) → OFX
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

document.getElementById("btnConvertCSV").addEventListener("click", () => {
  const file = document.getElementById("csvInput").files[0];
  const progress = document.getElementById("progressCSV");
  const bar = document.getElementById("progressBarCSV");
  const link = document.getElementById("downloadOFX");
  progress.textContent = "";
  bar.style.width = "0%";
  link.style.display = "none";

  if (!file) {
    alert("Selecione uma planilha Excel (.xlsx).");
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const total = rows.length - 1;
    const transactions = [];

    for (let i = 1; i < rows.length; i++) {
      const [dateRaw, memoRaw, amountRaw] = rows[i];
      if (!dateRaw || !memoRaw || !amountRaw) continue;

      const date = formatDateToOFX(dateRaw);
      const memo = String(memoRaw).trim();
      const amount = String(amountRaw).replace(",", ".").trim();

      transactions.push({ date, amount, memo });

      const percent = Math.floor((i / total) * 100);
      progress.textContent = `🔄 ${percent}% concluído`;
      bar.style.width = `${percent}%`;
    }

    const ofx = generateOFX(transactions);
    const blob = new Blob([ofx], { type: "text/plain;charset=utf-8;" });
    const url = URL.createObjectURL(blob);

    link.href = url;
    link.download = "arquivo.ofx";
    link.style.display = "inline";
    link.click();

    progress.textContent = `✅ ${transactions.length} transações convertidas para OFX.`;
    bar.style.width = "100%";
  };

  reader.readAsArrayBuffer(file);
});