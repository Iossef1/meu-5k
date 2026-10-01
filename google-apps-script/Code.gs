const SHEET_NAME = "Treinos";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, service:"Meu 5K"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const ss = SpreadsheetApp.openById(body.sheetId);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "id","data","semana","sessao","distancia_km","tempo_min",
        "continuo_min","panturrilha","respiracao","recuperacao",
        "observacoes","recebido_em"
      ]);
    }

    (body.runs || []).forEach(function(r) {
      sheet.appendRow([
        r.id, r.date, r.week, r.ses, r.dist, r.time, r.cont,
        r.calf, r.breath, r.rec, r.notes, new Date()
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ok:true, count:(body.runs || []).length}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}