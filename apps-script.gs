/**
 * ============================================================
 *  COOPERADORA ESCOLAR - BACKEND GOOGLE SHEETS
 * ============================================================
 * Copiar y pegar esto en: Extensiones > Apps Script
 * Luego hacer Deploy > New deployment > Web app
 * 
 * Acceso: Anyone, even anonymous
 * Ejecutar como: Me (tu cuenta de Google)
 * ============================================================
 */

const API_KEY = 'coop2026'; // <-- Cambiar por una clave propia

const SHEETS = {
  socios: 'Socios',
  bonos: 'Bonos',
  config: 'Config'
};

function doGet(e) {
  try {
    // Validar que existan los parámetros (cuando se abre la URL sin "?" puede no haber e.parameter)
    const params = (e && e.parameter) ? e.parameter : {};
    const key = params.apiKey || '';
    
    if (key !== API_KEY) {
      return jsonResponse({ error: 'Unauthorized', hint: 'Falta o es incorrecta la apiKey. Usa: ?action=read&apiKey=TU_CLAVE' });
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return jsonResponse({
      socios: readSheet(ss, SHEETS.socios),
      bonos: readSheet(ss, SHEETS.bonos),
      config: readConfig(ss)
    });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ error: 'No data received', hint: 'El body del POST está vacío o mal formado.' });
    }
    const data = JSON.parse(e.postData.contents);
    if (data.apiKey !== API_KEY) {
      return jsonResponse({ error: 'Unauthorized' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === 'write') {
      if (data.socios !== undefined) writeSheet(ss, SHEETS.socios, data.socios);
      if (data.bonos !== undefined) writeSheet(ss, SHEETS.bonos, data.bonos);
      if (data.config !== undefined) writeConfig(ss, data.config);
      return jsonResponse({ success: true, synced: true });
    }

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function readSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).trim());
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = rows[i][j] !== undefined ? rows[i][j] : '';
    }
    data.push(row);
  }
  return data;
}

function writeSheet(ss, sheetName, data) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();
  if (!data || data.length === 0) return;
  const headers = Object.keys(data[0]);
  sheet.appendRow(headers);
  const values = data.map(item => headers.map(h => item[h] !== undefined ? item[h] : ''));
  const range = sheet.getRange(2, 1, values.length, headers.length);
  range.setValues(values);
}

function readConfig(ss) {
  const sheet = ss.getSheetByName(SHEETS.config);
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) config[String(rows[i][0])] = rows[i][1];
  }
  return config;
}

function writeConfig(ss, config) {
  let sheet = ss.getSheetByName(SHEETS.config);
  if (!sheet) sheet = ss.insertSheet(SHEETS.config);
  sheet.clear();
  sheet.appendRow(['key', 'value']);
  Object.keys(config).forEach(k => {
    sheet.appendRow([k, config[k]]);
  });
}
