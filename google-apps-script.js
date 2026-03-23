// ============================================================
// APPS SCRIPT — Cole isso no Google Sheets
// Menu: Extensões > Apps Script
// ============================================================
//
// SETUP:
// 1. Abra sua planilha no Google Sheets
// 2. Vá em Extensões > Apps Script
// 3. Cole todo este código no editor
// 4. Clique em "Implantar" > "Nova implantação"
// 5. Tipo: "App da Web"
// 6. Executar como: "Eu" (sua conta)
// 7. Quem tem acesso: "Qualquer pessoa"
// 8. Copie a URL gerada e cole na ferramenta HTML
//
// Cada aba deve seguir o padrão: "SPO - São Paulo", "STS - Santos", etc.
// O script busca a aba que começa com o código da unidade.
// Colunas: A = Data | B = CNPJ | C = Nome (sem coluna Unidade)
// ============================================================

var ABA_MAP = {
  "SPO": "SPO - São Paulo",
  "STS": "STS - Santos",
  "SBC": "SBC - São Bernardo",
  "CPS": "CPS - Campinas"
};

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    var rows = data.rows;
    var now = new Date();
    var timestamp = Utilities.formatDate(now, "America/Sao_Paulo", "dd/MM/yyyy HH:mm");
    var inserted = 0;

    for (var i = 0; i < rows.length; i++) {
      var unidade = (rows[i].unidade || "").toUpperCase().trim();
      var sheetName = ABA_MAP[unidade];

      if (!sheetName) continue;

      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;

      var lastRow = sheet.getLastRow();
      var nextRow = lastRow < 2 ? 3 : lastRow + 1;

      sheet.getRange(nextRow, 1).setValue(timestamp);
      sheet.getRange(nextRow, 2).setValue(rows[i].cnpj);
      sheet.getRange(nextRow, 3).setValue(rows[i].nome);
      inserted++;
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", count: inserted }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Lançamento Rápido API ativa" }))
    .setMimeType(ContentService.MimeType.JSON);
}
