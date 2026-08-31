const SPREADSHEET_ID = '1_A8FyTO3u1L1Pn9l6YGCGS0MIPHIrjCO2LiKZgBtj3U';
const SHEET_NAME = 'Registros';
const EMPLOYEE_SHEET = 'Funcionarios';
const LEGACY_CSV_URL = 'https://raw.githubusercontent.com/escuelaoriente/registro-diario-funcionarios/main/historico-excel.csv';
const LEGACY_RECORDS_CSV_URL = 'https://raw.githubusercontent.com/willfret/registro-diario-funcionarios/main/legacy-records.csv';

function doGet(request) {
  if (request && request.parameter && request.parameter.action) {
    return apiResponse_(request);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Control de salidas diarias')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiResponse_(request) {
  const p = request.parameter || {};
  const action = p.action;
  let data;
  if (action === 'getEmployees') data = getEmployees();
  else if (action === 'addEmployee') data = addEmployee(p.name);
  else if (action === 'savePairedRecord') data = savePairedRecord({
    funcionario: p.funcionario,
    fecha: p.fecha,
    salidaHora: p.salidaHora,
    regresoHora: p.regresoHora,
    motivo: p.motivo
  });
  else if (action === 'getEmployeeDetail') data = getEmployeeDetail(p.funcionario, p.mode, p.value);
  else if (action === 'getAllRecords') data = getAllRecords();
  else throw new Error('Acción no reconocida.');

  const callback = String(p.callback || 'callback').replace(/[^a-zA-Z0-9_$]/g, '');
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  let employeeSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!employeeSheet) {
    employeeSheet = ss.insertSheet(EMPLOYEE_SHEET);
  }
  const headers = ['Timestamp', 'Fecha', 'Hora', 'Funcionario', 'Tipo', 'Motivo', 'Usuario'];
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  if (employeeSheet.getLastRow() === 0) {
    employeeSheet.getRange(1, 1).setValue('Nombre');
  }
  return { ok: true, message: 'Hoja Registros creada o reiniciada.' };
}

function importLegacyCsv() {
  const response = UrlFetchApp.fetch(LEGACY_CSV_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('No se pudo leer el histórico publicado en GitHub.');
  const rows = Utilities.parseCsv(response.getContentText());
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Historico Excel');
  if (!sheet) sheet = ss.insertSheet('Historico Excel');
  sheet.clearContents();
  if (rows.length) sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
  return { ok: true, rows: Math.max(0, rows.length - 1), message: 'Histórico del Excel importado correctamente.' };
}

function syncLegacyRecords() {
  const response = UrlFetchApp.fetch(LEGACY_RECORDS_CSV_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('No se pudo leer el histórico normalizado.');
  const rows = Utilities.parseCsv(response.getContentText()).slice(1);
  const { sheet } = ensureDatabase_();
  const existing = new Set();
  if (sheet.getLastRow() >= 2) {
    sheet.getDataRange().getValues().slice(1).forEach(row => existing.add([row[1], row[2], row[3], row[4]].join('|')));
  }
  const newRows = rows.filter(row => row.length >= 6).filter(row => {
    const key = [row[0], row[1], row[2], row[3]].join('|');
    if (existing.has(key)) return false;
    existing.add(key);
    return true;
  }).map(row => [new Date(), row[0], row[1], row[2], row[3], row[4], row[5]]);
  if (newRows.length) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  return { ok: true, rows: newRows.length, message: `${newRows.length} registros históricos sincronizados.` };
}

function replaceImportedRecords() {
  const response = UrlFetchApp.fetch(LEGACY_RECORDS_CSV_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('No se pudo leer el Excel sincronizado.');
  const rows = Utilities.parseCsv(response.getContentText()).slice(1).filter(row => row.length >= 6);
  const { sheet } = ensureDatabase_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || ['Timestamp', 'Fecha', 'Hora', 'Funcionario', 'Tipo', 'Motivo', 'Usuario'];
  const currentRecords = values.slice(1).filter(row => String(row[5] || '').trim() !== 'Registro importado del Excel');
  const importedRecords = rows.map(row => [new Date(), row[0], row[1], row[2], row[3], row[4], row[5]]);
  const output = [headers].concat(currentRecords, importedRecords);
  sheet.clearContents();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.setFrozenRows(1);
  return { ok: true, rows: importedRecords.length, preserved: currentRecords.length, message: `${importedRecords.length} registros del Excel reemplazados; ${currentRecords.length} registros nuevos conservados.` };
}

function replaceMarchMayRecords() {
  const response = UrlFetchApp.fetch(LEGACY_RECORDS_CSV_URL, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('No se pudo leer el Excel sincronizado.');
  const months = ['2026-03', '2026-04', '2026-05'];
  const rows = Utilities.parseCsv(response.getContentText()).slice(1)
    .filter(row => row.length >= 6 && months.some(month => String(row[0]).startsWith(month)));
  const { sheet } = ensureDatabase_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || ['Timestamp', 'Fecha', 'Hora', 'Funcionario', 'Tipo', 'Motivo', 'Usuario'];
  const preserved = values.slice(1).filter(row => !months.includes(dateText_(row[1]).slice(0, 7)));
  const imported = rows.map(row => [new Date(), row[0], row[1], row[2], row[3], row[4], row[5]]);
  const output = [headers].concat(preserved, imported);
  sheet.clearContents();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.setFrozenRows(1);
  return { ok: true, rows: imported.length, preserved: preserved.length, message: `Marzo, abril y mayo reemplazados con ${imported.length} registros del Excel.` };
}

function deduplicateRecords() {
  const { sheet } = ensureDatabase_();
  if (sheet.getLastRow() < 2) return { ok: true, removed: 0, message: 'No hay registros para limpiar.' };
  const values = sheet.getDataRange().getValues();
  const seen = new Set();
  const kept = [values[0]];
  values.slice(1).forEach(row => {
    const key = [dateText_(row[1]), timeText_(row[2]), String(row[3] || '').trim().toLowerCase(), String(row[4] || '').trim().toLowerCase(), String(row[5] || '').trim().toLowerCase()].join('|');
    if (!seen.has(key)) {
      seen.add(key);
      kept.push(row);
    }
  });
  sheet.clearContents();
  sheet.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
  sheet.setFrozenRows(1);
  return { ok: true, removed: values.length - kept.length, message: `${values.length - kept.length} duplicados eliminados.` };
}

function ensureDatabase_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7).setValues([['Timestamp', 'Fecha', 'Hora', 'Funcionario', 'Tipo', 'Motivo', 'Usuario']]);
    sheet.setFrozenRows(1);
  }
  let employeeSheet = ss.getSheetByName(EMPLOYEE_SHEET);
  if (!employeeSheet) employeeSheet = ss.insertSheet(EMPLOYEE_SHEET);
  if (employeeSheet.getLastRow() === 0) employeeSheet.getRange(1, 1).setValue('Nombre');
  return { ss, sheet, employeeSheet };
}

function getEmployees() {
  const { employeeSheet: sheet } = ensureDatabase_();
  if (sheet && sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().filter(Boolean);
    if (values.length) return values.sort();
  }

  return [
    'Ruama Acevedo', 'Nuvia Aguero', 'Pedro Almonacid', 'Norma Alvarado',
    'Ivan Antianco', 'Bernardino Antipan', 'Angela Orozco', 'Valeria Avendano',
    'Ivan Barra', 'Carla Burgos', 'Alonso Cadin', 'Mirta Calbuyahue',
    'Evelyn Carcamo', 'Liliana Carcamo', 'Gissela Carocca', 'Orfelina Casanova',
    'Marlene Castillo', 'Magaly Colivoro', 'Antonia Colivoro', 'Sara Farfan',
    'Humberto Gallardo', 'Hanuxa Geoffroy', 'Emilia Gonzalez', 'Ricardo Gonzalez',
    'Grenny Gutierrez', 'Marcela Haro', 'Luis Haro', 'Katheryn Jorquera',
    'Cristina Manquemilla', 'Cristian Mansilla', 'Oscar Mardones', 'Marco Marin',
    'Gerardo Marquez', 'Paola Martinez', 'Carla Mayorga', 'Richard Millaldeo',
    'Viviana Millaldeo', 'Lorena Moraga', 'Yacqueline Munoz', 'Clara Ojeda',
    'Josua Opazo', 'Priscila Ormeno', 'Rosa Oyarzo', 'Katrina Oyarzo',
    'Nataly Pardo', 'Angela Pereira', 'Fabian Pino', 'Rosa Rain',
    'Claudio Reyes', 'Jacqueline Riquelme', 'Alejandra Rodriguez', 'Jessica Rogel',
    'Carolina Rojas', 'Pamela Sanchez', 'Viviana Sanchez', 'Bianca Sanchez',
    'Lorena Sandoval', 'Paola Silva', 'Cinthia Soto', 'Carolina Tecay',
    'Jonathan Valderas', 'Loreto Valle', 'Angel Velasquez', 'Yanett Vasquez',
    'Barbara Vera', 'Cecilia Vera', 'Sandra Vera', 'Viviana Vera',
    'Diego Vera', 'Alicia Vergara', 'Paulina Vidal', 'Cesar Vidal',
    'Arlen Villagran', 'Bernardita Villarroel', 'Camila Yanez', 'Ma Jose Zamorano'
  ].sort();
}

function addEmployee(name) {
  const clean = String(name || '').trim();
  if (!clean) throw new Error('Debes escribir un nombre.');

  const { employeeSheet: sheet } = ensureDatabase_();

  const existing = getEmployees().map(n => String(n).trim().toLowerCase());
  if (existing.includes(clean.toLowerCase())) {
    return { ok: true, message: 'El funcionario ya existe.' };
  }

  sheet.appendRow([clean]);
  return { ok: true, message: 'Funcionario agregado correctamente.' };
}

function saveRecord(payload) {
  if (!payload || !payload.funcionario || !payload.tipo) {
    throw new Error('Faltan datos obligatorios.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    setup();
  }

  const now = new Date();
  const fecha = payload.fecha || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const hora = payload.hora || Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || 'anonimo';

  sheet.appendRow([
    now,
    fecha,
    hora,
    payload.funcionario,
    payload.tipo,
    payload.motivo || '',
    user
  ]);

  return { ok: true, message: 'Registro guardado correctamente.' };
}

function savePairedRecord(payload) {
  if (!payload || !payload.funcionario || !payload.salidaHora || !payload.regresoHora) {
    throw new Error('Faltan datos para registrar la salida y el regreso.');
  }

  const { sheet } = ensureDatabase_();

  const now = new Date();
  const fecha = payload.fecha || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const user = Session.getActiveUser().getEmail() || 'anonimo';
  const motivo = payload.motivo || 'Salida corta';
  sheet.appendRow([now, fecha, payload.salidaHora, payload.funcionario, 'Salida', motivo, user]);
  sheet.appendRow([now, fecha, payload.regresoHora, payload.funcionario, 'Regreso', motivo, user]);

  return { ok: true, message: 'Salida y regreso guardados correctamente.' };
}

function getTodayRecords(fecha) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const tz = Session.getScriptTimeZone();
  const target = fecha || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(row => dateText_(row[1]) === target);

  return rows.map(row => ({
    timestamp: row[0],
    fecha: dateText_(row[1]),
    hora: row[2],
    funcionario: row[3],
    tipo: row[4],
    motivo: row[5],
    usuario: row[6]
  }));
}

function getMonthlyRecords(funcionario, month) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues().slice(1);
  const targetMonth = month || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const nameFilter = (funcionario || '').trim().toLowerCase();

  return values
    .filter(row => {
      const rowDate = dateText_(row[1]);
      const matchesMonth = rowDate.startsWith(targetMonth);
      const matchesName = !nameFilter || String(row[3] || '').toLowerCase().includes(nameFilter);
      return matchesMonth && matchesName;
    })
    .map(row => ({
      timestamp: row[0],
      fecha: dateText_(row[1]),
      hora: timeText_(row[2]),
      funcionario: row[3],
      tipo: row[4],
      motivo: row[5],
      usuario: row[6]
    }));
}

function getDailyRecordsByEmployee(funcionario, date) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const targetDate = date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nameFilter = (funcionario || '').trim().toLowerCase();
  return sheet.getDataRange().getValues().slice(1)
    .filter(row => dateText_(row[1]) === targetDate && String(row[3] || '').toLowerCase().includes(nameFilter))
    .map(row => ({
      timestamp: row[0],
      fecha: dateText_(row[1]),
      hora: row[2],
      funcionario: row[3],
      tipo: row[4],
      motivo: row[5],
      usuario: row[6]
    }));
}

function getSummary(funcionario, date) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { totalMinutes: 0, totalHours: '0h 0m' };
  }

  const targetDate = date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const nameFilter = (funcionario || '').trim().toLowerCase();
  const rows = sheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[3] || '').toLowerCase().includes(nameFilter) && dateText_(row[1]) <= targetDate)
    .sort((a, b) => String(a[1]).localeCompare(String(b[1])) || timeToMinutes(a[2]) - timeToMinutes(b[2]));

  let totalMinutes = 0;
  let lastExit = null;
  rows.forEach(row => {
    const type = String(row[4] || '').toLowerCase();
    const minutes = timeToMinutes(row[2]);
    if (type === 'salida' || type === 'salida final') {
      lastExit = minutes;
    } else if ((type === 'regreso' || type === 'entrada') && lastExit !== null) {
      totalMinutes += Math.max(0, minutes - lastExit);
      lastExit = null;
    }
  });

  return { totalMinutes: totalMinutes, totalHours: formatMinutes(totalMinutes) };
}

function getMonthlySummary(funcionario, month) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { totalMinutes: 0, totalHours: '0h 0m' };
  }

  const targetMonth = month || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const nameFilter = (funcionario || '').trim().toLowerCase();
  const rows = sheet.getDataRange().getValues().slice(1)
    .filter(row => dateText_(row[1]).startsWith(targetMonth) && String(row[3] || '').toLowerCase().includes(nameFilter))
    .sort((a, b) => String(a[1]).localeCompare(String(b[1])) || timeToMinutes(a[2]) - timeToMinutes(b[2]));

  let totalMinutes = 0;
  let lastExit = null;
  rows.forEach(row => {
    const type = String(row[4] || '').toLowerCase();
    const minutes = timeToMinutes(row[2]);
    if (type === 'salida' || type === 'salida final') {
      lastExit = minutes;
    } else if ((type === 'regreso' || type === 'entrada') && lastExit !== null) {
      totalMinutes += Math.max(0, minutes - lastExit);
      lastExit = null;
    }
  });

  return { totalMinutes: totalMinutes, totalHours: formatMinutes(totalMinutes) };
}

function getEmployeeDetail(funcionario, mode, value) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { items: [], summary: { entries: 0, exits: 0, returns: 0, totalMinutes: 0, totalHours: '0h 0m' } };
  }

  const targetName = String(funcionario || '').trim().toLowerCase();
  const targetMode = String(mode || 'month');
  const targetValue = String(value || '').trim();
  const rows = sheet.getDataRange().getValues().slice(1).filter(row => String(row[3] || '').toLowerCase().includes(targetName));

  const filtered = rows.filter(row => {
    const rowDate = dateText_(row[1]);
    return targetMode === 'day' ? rowDate === targetValue : rowDate.startsWith(targetValue);
  }).sort((a, b) => String(a[1]).localeCompare(String(b[1])) || timeToMinutes(a[2]) - timeToMinutes(b[2]));

  let entries = 0;
  let exits = 0;
  let returns = 0;
  let totalMinutes = 0;
  let lastExit = null;

  filtered.forEach(row => {
    const type = String(row[4] || '').toLowerCase();
    const minutes = timeToMinutes(row[2]);
    if (type === 'entrada') entries += 1;
    if (type === 'salida' || type === 'salida final') {
      exits += 1;
      lastExit = minutes;
    } else if (type === 'regreso') {
      returns += 1;
      if (lastExit !== null) {
        totalMinutes += Math.max(0, minutes - lastExit);
        lastExit = null;
      }
    }
  });

  return {
    items: filtered.map(row => ({
      fecha: dateText_(row[1]),
      hora: row[2],
      funcionario: row[3],
      tipo: row[4],
      motivo: row[5]
    })),
    summary: {
      entries,
      exits,
      returns,
      totalMinutes,
      totalHours: formatMinutes(totalMinutes)
    }
  };
}

function getAllRecords() {
  const { sheet } = ensureDatabase_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    fecha: dateText_(row[1]),
    hora: timeText_(row[2]),
    funcionario: row[3],
    tipo: row[4],
    motivo: row[5],
    usuario: row[6]
  }));
}

function timeToMinutes(value) {
  if (value instanceof Date) {
    return value.getHours() * 60 + value.getMinutes();
  }
  if (typeof value === 'string') {
    const parts = value.trim().split(':');
    if (parts.length >= 2) {
      return Number(parts[0]) * 60 + Number(parts[1]);
    }
  }
  return 0;
}

function timeText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  const text = String(value || '').trim();
  const match = text.match(/(\d{1,2}):(\d{2})/);
  return match ? `${String(match[1]).padStart(2, '0')}:${match[2]}` : text;
}

function dateText_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value || '').slice(0, 10);
}

function formatMinutes(totalMinutes) {
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${h}h ${m}m`;
}
