// © Dev-MB | dev-mb.dev
// Google Apps Script für Termine-API
// DIESES SCRIPT IN DEIN GOOGLE SHEET EINFÜGEN

function doGet(e) {
  try {
    // Zugriff auf das aktuelle Google Sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Alle Daten aus dem Sheet lesen (ab Zeile 2, da Zeile 1 Header ist)
    // Header: Datum | Start | Thema/Was | Wo | Wer | Beschreibung (optional)
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Termine als Array von Objekten erstellen
    const termine = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Leere Zeilen überspringen
      if (!row[0]) continue;
      
      // Datum formatieren
      let datum = '';
      if (row[0] instanceof Date) {
        const d = row[0];
        datum = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
      } else {
        datum = String(row[0]);
      }
      
      // Uhrzeit formatieren - ODER Emoji+Text
      let start = '';
      if (row[1] instanceof Date) {
        // Google Sheets speichert Uhrzeiten als Date mit Basis 30.12.1899
        // Wir wollen nur HH:MM
        const t = row[1];
        start = `${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`;
      } else if (row[1]) {
        // Falls es ein String ist, direkt übernehmen (für Zeit ODER Emoji+Text)
        const timeStr = String(row[1]).trim();
        if (timeStr) {
          start = timeStr;
        }
      }
      
      // Titel - wenn mehrere Zeilen, nur erste Zeile als Titel
      let titel = row[2] || 'Kein Titel';
      let beschreibung = '';
      
      if (titel.includes('\n')) {
        const lines = titel.split('\n');
        titel = lines[0];
        beschreibung = lines.slice(1).join('\n');
      }
      
      // Optional: Spalte 6 für zusätzliche Beschreibung
      if (row[5]) {
        if (beschreibung) {
          beschreibung += '\n' + String(row[5]);
        } else {
          beschreibung = String(row[5]);
        }
      }
      
      termine.push({
        datum: datum,
        start: start,
        titel: titel,
        wo: row[3] || '',
        wer: row[4] || '',
        beschreibung: beschreibung
      });
    }
    
    // JSON zurückgeben
    const output = {
      success: true,
      count: termine.length,
      lastUpdate: new Date().toISOString(),
      termine: termine
    };
    
    return ContentService.createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
