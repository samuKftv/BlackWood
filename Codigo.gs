/**
 * Blackwood, 1926 · Reparto de personajes
 * Script de Google Apps Script que guarda las reservas en esta hoja de cálculo.
 * Cambia el código de docente antes de publicar.
 */
const PIN_DOCENTE = 'cambia-este-codigo';
const HOJA = 'Reparto';
const COLUMNAS = ['id', 'fam', 'role', 'custom', 'roleName', 'alumna', 'personaje', 'estilo', 'at'];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(COLUMNAS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function leer_() {
  const sh = hoja_();
  const datos = sh.getDataRange().getValues();
  const cab = datos.shift();
  return datos.map(fila => {
    const o = {};
    cab.forEach((c, i) => (o[c] = fila[i]));
    return o;
  });
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function limpia_(v, max) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'").slice(0, max || 80);
}

function doGet() {
  return salida_({ ok: true, claims: leer_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const sh = hoja_();
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]));

    if (d.action === 'reserve') {
      if (!d.id || !d.alumna) return salida_({ ok: false, error: 'datos' });
      if (ids.indexOf(String(d.id)) !== -1) return salida_({ ok: false, error: 'ocupado' });
      sh.appendRow([
        limpia_(d.id, 60), limpia_(d.fam, 20), limpia_(d.role, 60), d.custom === true,
        limpia_(d.roleName, 60), limpia_(d.alumna, 60), limpia_(d.personaje, 60),
        limpia_(d.estilo, 20), new Date()
      ]);
      return salida_({ ok: true });
    }

    if (d.action === 'release') {
      if (d.pin !== PIN_DOCENTE) return salida_({ ok: false, error: 'pin' });
      const fila = ids.indexOf(String(d.id));
      if (fila > 0) sh.deleteRow(fila + 1);
      return salida_({ ok: true });
    }

    return salida_({ ok: false, error: 'accion' });
  } finally {
    lock.releaseLock();
  }
}
