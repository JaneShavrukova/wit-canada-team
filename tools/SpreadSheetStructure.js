function showFullJson() {
  const json = exportSpreadsheetStructure();

  const html = HtmlService.createHtmlOutput(`
    <div style="padding:12px;font-family:monospace;">
      <button onclick="copyText()" style="margin-bottom:10px;">Copy</button>
      <pre id="content" style="white-space:pre-wrap;word-break:break-word;">
${escapeHtml(json)}
      </pre>
    </div>

    <script>
      function copyText() {
        const text = document.getElementById('content').innerText;
        navigator.clipboard.writeText(text);
        alert('Copied');
      }
    </script>
  `)
  .setWidth(900)
  .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, "Sheet Structure");
}


function exportSpreadsheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const result = {
    spreadsheetName: ss.getName(),
    sheets: sheets.map(sheet => {
      const data = sheet.getDataRange().getValues();

      if (data.length === 0) {
        return {
          sheetName: sheet.getName(),
          rowCount: 0,
          columnCount: 0,
          columns: []
        };
      }

      const headers = data[0];
      const rows = data.slice(1);

      const columns = headers
        .map((header, colIndex) => {
          if (!header) return null;

          const values = rows
            .map(row => row[colIndex])
            .filter(v => v !== "" && v !== null);

          return {
            name: header,
            normalizedName: header.toString().toLowerCase().replace(/\s+/g, "_"),
            type: detectType(values),
            sample: values.slice(0, 3)
          };
        })
        .filter(Boolean);

      return {
        sheetName: sheet.getName(),
        rowCount: rows.length,
        columnCount: headers.length,
        columns: columns
      };
    })
  };

  return JSON.stringify(result, null, 2);
}


function detectType(values) {
  if (values.length === 0) return "empty";

  if (values.every(v => typeof v === "number")) return "number";

  if (values.every(v => typeof v === "boolean")) return "boolean";

  if (values.some(v => v instanceof Date)) return "date";

  const unique = new Set(values);
  if (unique.size <= 10) return "enum";

  return "string";
}


function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}