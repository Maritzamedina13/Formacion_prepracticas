/**
 * OVA Preprácticas ITM — Web App para Google Sheets
 *
 * Instrucciones de despliegue:
 *  1. En el Google Sheet: Extensiones > Apps Script
 *  2. Pega todo este código (reemplaza el contenido existente)
 *  3. Guardar (Ctrl+S)
 *  4. Implementar > Nueva implementación
 *     - Tipo: Aplicación web
 *     - Ejecutar como: Yo (tu cuenta)
 *     - Quién tiene acceso: Cualquier usuario
 *  5. Autorizar el acceso cuando se solicite
 *  6. Copia la URL generada
 *  7. En index.html, reemplaza REEMPLAZAR_CON_URL_APPS_SCRIPT con esa URL
 */

var SPREADSHEET_ID = '1etkSENFncJgRmQdnoSkJMihk4B_i1OnF80L7-FKDZU0';
var SHEET_NAME     = 'Participantes';

var MODULES = [
  {id: 1, title: 'Compromiso y Reglamentos',     xp: 100, quiz: 3},
  {id: 2, title: 'Seguridad de la Información',  xp: 100, quiz: 3},
  {id: 3, title: 'Habilidades Blandas',           xp: 150, quiz: 3},
  {id: 4, title: 'El Mundo Organizacional',       xp: 100, quiz: 3},
  {id: 5, title: 'Metodologías Ágiles',           xp: 150, quiz: 3},
  {id: 6, title: 'Manejo del Tiempo',             xp: 100, quiz: 3},
  {id: 7, title: 'IA y Uso Responsable',          xp: 150, quiz: 3},
  {id: 8, title: 'Herramientas y Actualización',  xp: 150, quiz: 3}
];

// ── Punto de entrada POST ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    guardarParticipante(data);
    return _resp({ok: true});
  } catch (err) {
    return _resp({ok: false, error: err.message});
  }
}

// ── Punto de entrada GET (prueba de conectividad) ─────────────────────────────
function doGet(e) {
  return _resp({ok: true, mensaje: 'Web App OVA Preprácticas ITM activa'});
}

// ── Guardar / actualizar participante ─────────────────────────────────────────
function guardarParticipante(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Hoja "' + SHEET_NAME + '" no encontrada. Ejecuta primero setup_template() en Python.');
  }

  // Buscar fila existente por cédula (columna B = 2)
  var lastRow    = Math.max(sheet.getLastRow(), 2);
  var cedulaVals = sheet.getRange(3, 2, lastRow - 2 + 1, 1).getValues();
  var existingRow = -1;
  for (var i = 0; i < cedulaVals.length; i++) {
    if (String(cedulaVals[i][0]).trim() === String(data.cc).trim()) {
      existingRow = i + 3;   // 1-indexed, saltamos 2 filas de encabezado
      break;
    }
  }

  var fecha = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm');

  // Construir fila
  var row = [
    data.nombre                 || '',
    data.cc                     || '',
    data.email                  || '',
    data.modulos_completados    || 0,
    (data.progreso_pct || 0) + '%',
    data.xp_total               || 0,
    fecha
  ];

  var modulos = data.modulos || [];
  for (var j = 0; j < MODULES.length; j++) {
    var m = j < modulos.length ? modulos[j] : {};
    row.push(
      m.contenido     ? '✓' : '—',
      m.quiz_aprobado ? '✓' : '—',
      (m.puntaje || 0) + '/' + (m.total || MODULES[j].quiz),
      m.xp || 0
    );
  }

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

// ── Helper respuesta JSON ─────────────────────────────────────────────────────
function _resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
