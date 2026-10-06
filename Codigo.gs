/**
 * Blackwood, 1926 · Reparto de personajes
 * Script de Google Apps Script que guarda las reservas en esta hoja de cálculo.
 * Cambia el código de docente antes de publicar.
 */
const PIN_DOCENTE = 'cambia-este-codigo';
const HOJA = 'Reparto';
const COLUMNAS = ['id', 'fam', 'role', 'custom', 'roleName', 'alumna', 'personaje', 'estilo', 'at'];
const HOJA_HISTORIAS = 'Historias';
const CAMPOS_HISTORIA = [
  ['nombre', 'Nombre del personaje'], ['edad_ap', 'Edad aparente'], ['edad_real', 'Edad real'], ['papel', 'Papel en su familia'],
  ['nacimiento', 'Nacimiento e infancia'], ['marca', 'Qué le marcó'], ['familia_rel', 'Relación con su familia'], ['acuerdo', 'Opinión sobre el Acuerdo'],
  ['rasgos', 'Rasgos de carácter'], ['desea', 'Qué quiere y qué teme'], ['gesto', 'Gesto o manía'],
  ['porque_gala', 'Por qué vino a la gala'], ['severina', 'Relación con Severina'], ['donde', 'Dónde estaba'], ['culpable', 'Por qué podría ser culpable'],
  ['ll_rostro', 'Llegada: rostro'], ['ll_cuerpo', 'Llegada: cuerpo y manos'], ['ll_cabello', 'Llegada: cabello'], ['ll_vestuario', 'Llegada: vestuario'],
  ['ec_rostro', 'Eclipse: rostro'], ['ec_cuerpo', 'Eclipse: cuerpo y manos'], ['ec_cabello', 'Eclipse: cabello'], ['ec_vestuario', 'Eclipse: vestuario'],
  ['just1', 'Rasgo justificado 1'], ['just2', 'Rasgo justificado 2'], ['just3', 'Rasgo justificado 3']
];

function hojaHistorias_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_HISTORIAS);
  if (!sh) sh = ss.insertSheet(HOJA_HISTORIAS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Fecha de envío', 'Rol (id)', 'Familia', 'Rol', 'Alumna'].concat(CAMPOS_HISTORIA.map(c => c[1])));
    sh.setFrozenRows(1);
  }
  return sh;
}

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  if (sh.getLastRow() === 0) {
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
    const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));

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

    if (d.action === 'historia') {
      if (!d.roleId || !d.alumna) return salida_({ ok: false, error: 'datos' });
      const datos = d.data || {};
      hojaHistorias_().appendRow(
        [new Date(), limpia_(d.roleId, 60), limpia_(d.familia, 30), limpia_(d.rol, 60), limpia_(d.alumna, 60)]
          .concat(CAMPOS_HISTORIA.map(c => limpia_(datos[c[0]], 3000)))
      );
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
