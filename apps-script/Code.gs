/**
 * Reporte de fallos – Pruebas del Agente de IA CMV20J
 * ---------------------------------------------------
 * Script vinculado a una Google Sheet (Extensiones > Apps Script).
 * - doPost(): recibe los reportes que envía la landing publicada en Vercel.
 * - setup(): crea la hoja "Fallos", encabezados, validaciones y la carpeta de evidencias en Drive.
 * - submitReport(): recibe el formulario, guarda las 2 fotos en Drive y agrega la fila en la hoja.
 */

const CONFIG = {
  SHEET_NAME: 'Fallos',
  FOLDER_NAME: 'Evidencias - Pruebas Agente CMV20J',
  ID_PREFIX: 'F-',
  MAX_IMAGE_BYTES: 8 * 1024 * 1024
};

const HEADERS = [
  'ID', 'Fecha de registro', 'Fecha de la prueba', 'Tester', 'Área / servicio',
  'Tipo de fallo', 'Severidad', 'Paso del flujo', 'Lo que escribió el tester',
  'Lo que respondió el agente', 'Lo que debió pasar', 'Descripción del fallo',
  'Evidencia 1', 'Evidencia 2', 'Carpeta de evidencias',
  'Estado', 'Responsable', 'Notas de corrección'
];

const ESTADOS = ['Pendiente', 'En corrección', 'Corregido', 'Re-probado OK', 'Descartado'];

/* ---------- API para la landing en Vercel ---------- */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = submitReport(data);
    return json_({ ok: true, id: result.id });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'Reporte de fallos – Agente CMV20J' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Configuración inicial (ejecutar una vez) ---------- */

function setup() {
  const sheet = getSheet_();
  getFolder_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Hoja y carpeta de evidencias listas.', 'Configuración', 5);
  return sheet.getName();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight('bold')
      .setBackground('#10302C')
      .setFontColor('#FFFFFF')
      .setWrap(true)
      .setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    sheet.setRowHeight(1, 42);

    const widths = [80, 140, 110, 140, 170, 230, 90, 160, 260, 260, 260, 320, 110, 110, 130, 130, 140, 260];
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

    // Lista desplegable de Estado para el equipo que corrige
    const estadoCol = HEADERS.indexOf('Estado') + 1;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build();
    sheet.getRange(2, estadoCol, sheet.getMaxRows() - 1, 1).setDataValidation(rule);

    // Colores por severidad
    const sevCol = HEADERS.indexOf('Severidad') + 1;
    const sevRange = sheet.getRange(2, sevCol, sheet.getMaxRows() - 1, 1);
    const colors = { 'Crítica': '#F8D7D3', 'Alta': '#FBE3CF', 'Media': '#FBF0C9', 'Baja': '#E3ECE9' };
    const rules = Object.keys(colors).map(k =>
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(k).setBackground(colors[k]).setRanges([sevRange]).build()
    );
    sheet.setConditionalFormatRules(rules);
  }
  return sheet;
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* se recrea abajo */ }
  }
  const folder = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

/* ---------- Recepción del formulario ---------- */

function submitReport(data) {
  const required = ['tester', 'fechaPrueba', 'area', 'tipoFallo', 'severidad', 'esperado', 'descripcion'];
  required.forEach(k => {
    if (!data || !String(data[k] || '').trim()) throw new Error('Falta el campo obligatorio: ' + k);
  });
  if (!Array.isArray(data.imagenes) || data.imagenes.length !== 2 || data.imagenes.some(i => !i || !i.base64)) {
    throw new Error('Debes adjuntar las 2 fotos de evidencia.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    const nextNumber = Math.max(sheet.getLastRow(), 1); // fila 1 = encabezados
    const reportId = CONFIG.ID_PREFIX + String(nextNumber).padStart(4, '0');

    const parent = getFolder_();
    const folder = parent.createFolder(reportId + ' - ' + clean_(data.area).slice(0, 40));

    const links = data.imagenes.map((img, i) => {
      const bytes = Utilities.base64Decode(img.base64);
      if (bytes.length > CONFIG.MAX_IMAGE_BYTES) throw new Error('La foto ' + (i + 1) + ' es demasiado pesada.');
      const ext = (img.mimeType || 'image/jpeg').split('/')[1].replace('jpeg', 'jpg');
      const blob = Utilities.newBlob(bytes, img.mimeType || 'image/jpeg', reportId + '_evidencia' + (i + 1) + '.' + ext);
      const file = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* dominio puede bloquearlo */ }
      return file.getUrl();
    });

    const row = [
      reportId,
      new Date(),
      clean_(data.fechaPrueba),
      clean_(data.tester),
      clean_(data.area),
      clean_(data.tipoFallo),
      clean_(data.severidad),
      clean_(data.paso),
      clean_(data.mensajeTester),
      clean_(data.respuestaAgente),
      clean_(data.esperado),
      clean_(data.descripcion),
      '=HYPERLINK("' + links[0] + '","Ver foto 1")',
      '=HYPERLINK("' + links[1] + '","Ver foto 2")',
      '=HYPERLINK("' + folder.getUrl() + '","Abrir carpeta")',
      'Pendiente',
      '',
      ''
    ];
    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), 1, 1, HEADERS.length).setWrap(true).setVerticalAlignment('top');

    return { id: reportId };
  } finally {
    lock.releaseLock();
  }
}

/** Evita que un texto que empiece con =, +, - o @ se interprete como fórmula. */
function clean_(value) {
  const s = String(value == null ? '' : value).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
