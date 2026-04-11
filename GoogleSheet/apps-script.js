// ══════════════════════════════════════════════════════
// HELBREDLOG — Google Apps Script Backend
// ══════════════════════════════════════════════════════
//
// SETUP:
// 1. Opret et nyt Google Sheet
// 2. Gå til Udvidelser → Apps Script
// 3. Indsæt HELE denne fil
// 4. Skift SECRET_TOKEN til dit eget hemmelige ord
// 5. Klik Deploy → Ny implementering → Webapp
// 6. Kør som: Dig selv
//    Adgang: Alle (du er den eneste der kender URL'en)
// 7. Kopiér URL og indsæt i appen under Indstillinger
// ══════════════════════════════════════════════════════

const SECRET_TOKEN = 'HealtTracker6500!'; // ← SKIFT DETTE!
const SHEET_NAME_PAIN = 'Smerter';
const SHEET_NAME_SICK = 'Sygdom';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Sikkerhedstjek
    if (SECRET_TOKEN !== '' && data.secret !== SECRET_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Ugyldig nøgle' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date(data.timestamp);
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm');

    if (data.type === 'pain') {
      // ── SMERTE-SHEET ──
      let sheet = ss.getSheetByName(SHEET_NAME_PAIN);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_PAIN);
        sheet.appendRow(['Dato', 'Tidspunkt', 'Område', 'Intensitet (1-5)', 'Note', 'ISO Timestamp']);
        // Formatér header
        sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a2a38').setFontColor('#3ecfcf');
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 100);
        sheet.setColumnWidth(2, 80);
        sheet.setColumnWidth(3, 150);
        sheet.setColumnWidth(4, 130);
        sheet.setColumnWidth(5, 300);
        sheet.setColumnWidth(6, 180);
      }

      const row = [
        formattedDate,
        formattedTime,
        data.location || '',
        data.intensity,
        data.note || '',
        data.timestamp
      ];
      sheet.appendRow(row);

      // Farvemarkér intensitet
      const lastRow = sheet.getLastRow();
      const intensityCell = sheet.getRange(lastRow, 4);
      const colors = { 1: '#d4edda', 2: '#e8f5c8', 3: '#fff3cd', 4: '#ffe0b2', 5: '#f8d7da' };
      intensityCell.setBackground(colors[data.intensity] || '#ffffff');

    } else if (data.type === 'sick') {
      // ── SYGDOMS-SHEET ──
      let sheet = ss.getSheetByName(SHEET_NAME_SICK);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME_SICK);
        sheet.appendRow(['Dato', 'Tidspunkt', 'Symptomer', 'Alvorlighed (1-5)', 'Note', 'ISO Timestamp']);
        sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#2a1a1a').setFontColor('#e05c5c');
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 100);
        sheet.setColumnWidth(2, 80);
        sheet.setColumnWidth(3, 250);
        sheet.setColumnWidth(4, 140);
        sheet.setColumnWidth(5, 300);
        sheet.setColumnWidth(6, 180);
      }

      const row = [
        formattedDate,
        formattedTime,
        data.symptoms || '',
        data.intensity,
        data.note || '',
        data.timestamp
      ];
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Tillad GET for at teste at scriptet virker
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'HelbredLog API kører ✓' }))
    .setMimeType(ContentService.MimeType.JSON);
}
