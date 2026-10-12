/**
 * Blackwood, 1926 · Reparto, historias y entregas
 * Script de Google Apps Script que guarda los datos del proyecto en esta hoja de cálculo
 * y las imágenes de las entregas en tu Google Drive.
 *
 * ANTES DE PUBLICAR, cambia estos dos códigos:
 *  - PIN_DOCENTE: solo lo sabe el profesor (liberar roles y ocultar entregas).
 *  - CODIGO_CLASE: lo sabe todo el grupo (para entrar en la web).
 */
const PIN_DOCENTE = 'cambia-este-codigo';
const CODIGO_CLASE = 'cambia-codigo-clase';
const CARPETA_DRIVE = 'Blackwood, 1926 · Entregas';

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
const HOJA_ENTREGAS = 'Entregas';
const COLUMNAS_ENTREGAS = ['id', 'Fecha', 'Rol (id)', 'Familia', 'Rol', 'Alumna', 'Unidad', 'Tipo', 'Nota', 'Colores', 'Enlace', 'Archivo (id)', 'Visible'];

/* ---------- Hojas ---------- */

function pestana_(nombre, cabecera) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cabecera);
    sh.setFrozenRows(1);
  }
  return sh;
}
function hoja_() { return pestana_(HOJA, COLUMNAS); }
function hojaHistorias_() {
  return pestana_(HOJA_HISTORIAS, ['Fecha de envío', 'Rol (id)', 'Familia', 'Rol', 'Alumna'].concat(CAMPOS_HISTORIA.map(c => c[1])));
}
function hojaEntregas_() { return pestana_(HOJA_ENTREGAS, COLUMNAS_ENTREGAS); }

function filas_(sh) {
  const datos = sh.getDataRange().getValues();
  datos.shift();
  return datos.filter(f => f.some(v => v !== ''));
}

function leer_() {
  const sh = hoja_();
  const datos = sh.getDataRange().getValues();
  const cab = datos.shift();
  return datos.filter(f => f.some(v => v !== '')).map(fila => {
    const o = {};
    cab.forEach((c, i) => (o[c] = fila[i]));
    return o;
  });
}

/* La última historia enviada de cada personaje. */
function ultimasHistorias_() {
  const out = {};
  filas_(hojaHistorias_()).forEach(f => {
    const data = {};
    CAMPOS_HISTORIA.forEach((c, i) => (data[c[0]] = String(f[5 + i] == null ? '' : f[5 + i])));
    out[String(f[1])] = { fecha: f[0], alumna: String(f[4]), data: data };
  });
  return out;
}

/* Entregas visibles. */
function leerEntregas_() {
  return filas_(hojaEntregas_())
    .filter(f => f[12] !== false && String(f[12]).toUpperCase() !== 'FALSE')
    .map(f => ({
      id: String(f[0]), fecha: f[1], roleId: String(f[2]), alumna: String(f[5]),
      ut: String(f[6]), tipo: String(f[7]), nota: String(f[8]),
      colores: String(f[9]), enlace: String(f[10]), fileId: String(f[11])
    }));
}

/* ---------- Drive ---------- */

function carpeta_(roleId) {
  const raiz = DriveApp.getFoldersByName(CARPETA_DRIVE);
  const base = raiz.hasNext() ? raiz.next() : DriveApp.createFolder(CARPETA_DRIVE);
  const sub = base.getFoldersByName(roleId);
  return sub.hasNext() ? sub.next() : base.createFolder(roleId);
}

function guardarImagen_(roleId, dataUrl, nombre) {
  const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) throw new Error('imagen');
  if (m[2].length > 12000000) throw new Error('grande');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombre);
  const file = carpeta_(roleId).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

/* ---------- Utilidades ---------- */

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function limpia_(v, max) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'").slice(0, max || 80);
}
function codigoOk_(c) { return String(c || '') === CODIGO_CLASE; }

/* ---------- Web ---------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!codigoOk_(p.code)) return salida_({ ok: false, error: 'codigo' });
  return salida_({ ok: true, claims: leer_(), historias: ultimasHistorias_(), entregas: leerEntregas_() });
}

function doPost(e) {
  const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  if (!codigoOk_(d.code)) return salida_({ ok: false, error: 'codigo' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
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

    if (d.action === 'entrega') {
      if (!d.roleId || !d.alumna || !d.ut || !d.tipo) return salida_({ ok: false, error: 'datos' });
      if (!d.imagen && !d.enlace && !d.colores) return salida_({ ok: false, error: 'vacia' });
      let fileId = '';
      if (d.imagen) {
        try { fileId = guardarImagen_(limpia_(d.roleId, 60), d.imagen, limpia_(d.nombreArchivo || 'entrega.jpg', 120)); }
        catch (err) { return salida_({ ok: false, error: String(err.message) === 'grande' ? 'grande' : 'drive', detalle: String(err.message) }); }
      }
      const enlace = /^https?:\/\//.test(String(d.enlace || '')) ? limpia_(d.enlace, 500) : '';
      const id = Utilities.getUuid();
      hojaEntregas_().appendRow([
        id, new Date(), limpia_(d.roleId, 60), limpia_(d.familia, 30), limpia_(d.rol, 60), limpia_(d.alumna, 60),
        limpia_(d.ut, 10), limpia_(d.tipo, 60), limpia_(d.nota, 1000), limpia_(d.colores, 200), enlace, fileId, true
      ]);
      return salida_({ ok: true, id: id, fileId: fileId });
    }

    if (d.action === 'ocultar') {
      if (d.pin !== PIN_DOCENTE) return salida_({ ok: false, error: 'pin' });
      const he = hojaEntregas_();
      const col = he.getRange(1, 1, Math.max(he.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
      const fila = col.indexOf(String(d.id));
      if (fila > 0) he.getRange(fila + 1, 13).setValue(false);
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

/* Ejecuta esta función una vez desde el editor (botón ▶) para dar permiso de Drive. */
function autorizar() {
  const f = DriveApp.getFoldersByName(CARPETA_DRIVE);
  if (!f.hasNext()) DriveApp.createFolder(CARPETA_DRIVE);
  Logger.log('Listo: la carpeta "' + CARPETA_DRIVE + '" existe en tu Drive.');
}
