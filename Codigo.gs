/**
 * Blackwood, 1926 · Reparto, historias y fichas de personaje
 * Script de Google Apps Script que guarda los datos en esta hoja de cálculo.
 *
 * IMPORTANTE: antes de pegar este código, copia tu código de docente actual
 * y vuelve a escribirlo en la línea PIN_DOCENTE.
 */
const PIN_DOCENTE = 'cambia-este-codigo';
const HOJA = 'Reparto';
const COLUMNAS = ['id', 'fam', 'role', 'custom', 'roleName', 'alumna', 'personaje', 'estilo', 'at', 'carpeta'];
const HOJA_HISTORIAS = 'Historias';
const HOJA_CLAVES = 'Contraseñas';
const CARPETA_RAIZ = 'Blackwood · Imágenes';
const CAMPOS_HISTORIA = [
  ['nombre', 'Nombre del personaje'], ['edad_ap', 'Edad aparente'], ['edad_real', 'Edad real'], ['papel', 'Papel en su familia'],
  ['nacimiento', 'Nacimiento e infancia'], ['marca', 'Qué le marcó'], ['familia_rel', 'Relación con su familia'], ['acuerdo', 'Opinión sobre el Acuerdo'],
  ['rasgos', 'Rasgos de carácter'], ['desea', 'Qué quiere y qué teme'], ['gesto', 'Gesto o manía'],
  ['porque_gala', 'Por qué vino a la gala'], ['severina', 'Relación con Severina'], ['donde', 'Dónde estaba'], ['culpable', 'Por qué podría ser culpable'],
  ['ll_rostro', 'Llegada: rostro'], ['ll_cuerpo', 'Llegada: cuerpo y manos'], ['ll_cabello', 'Llegada: cabello'], ['ll_vestuario', 'Llegada: vestuario'],
  ['ec_rostro', 'Eclipse: rostro'], ['ec_cuerpo', 'Eclipse: cuerpo y manos'], ['ec_cabello', 'Eclipse: cabello'], ['ec_vestuario', 'Eclipse: vestuario'],
  ['just1', 'Rasgo justificado 1'], ['just2', 'Rasgo justificado 2'], ['just3', 'Rasgo justificado 3']
];

/* ---------- menú de la hoja ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Blackwood')
    .addItem('Crear carpetas de imágenes', 'crearCarpetas')
    .addToUi();
}

/**
 * Crea una carpeta de Drive por cada rol reservado que aún no tenga carpeta,
 * la deja visible para «cualquier persona con el enlace» y pega su enlace en la columna «carpeta».
 * Se puede ejecutar todas las veces que haga falta: solo crea las que faltan.
 */
function crearCarpetas() {
  const sh = hoja_();
  const datos = sh.getDataRange().getValues();
  const cab = datos[0];
  const cId = cab.indexOf('id'), cAl = cab.indexOf('alumna'), cPj = cab.indexOf('personaje'), cCa = cab.indexOf('carpeta');
  const raices = DriveApp.getFoldersByName(CARPETA_RAIZ);
  const raiz = raices.hasNext() ? raices.next() : DriveApp.createFolder(CARPETA_RAIZ);
  let creadas = 0, avisos = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (!f[cId] || !f[cAl] || f[cCa]) continue;
    const nombre = (f[cPj] || f[cId]) + ' (' + f[cAl] + ')';
    const carpeta = raiz.createFolder(nombre);
    try {
      carpeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      avisos.push(nombre);
    }
    sh.getRange(i + 1, cCa + 1).setValue(carpeta.getUrl());
    creadas++;
  }
  let msg = 'Carpetas creadas: ' + creadas + '. Están dentro de «' + CARPETA_RAIZ + '» en tu Drive.';
  if (avisos.length) msg += '\n\nNo se pudo activar «Cualquier persona con el enlace» en: ' + avisos.join(', ') + '. Actívalo a mano en cada una.';
  msg += '\n\nAhora comparte cada carpeta con su alumna como Editora.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* ---------- hojas ---------- */

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNAS);
    sh.setFrozenRows(1);
  } else {
    const cab = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    COLUMNAS.forEach((c, i) => { if (cab[i] !== c && cab.indexOf(c) === -1) sh.getRange(1, i + 1).setValue(c); });
  }
  return sh;
}

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

function hojaClaves_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA_CLAVES);
  if (!sh) sh = ss.insertSheet(HOJA_CLAVES);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Rol (id)', 'Personaje', 'Alumna', 'Contraseña', 'Creada']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function leer_() {
  const datos = hoja_().getDataRange().getValues();
  const cab = datos.shift();
  return datos.map(fila => {
    const o = {};
    cab.forEach((c, i) => (o[c] = fila[i]));
    return o;
  });
}

function claves_() {
  const datos = hojaClaves_().getDataRange().getValues();
  datos.shift();
  const m = {};
  datos.forEach(f => { if (f[0]) m[String(f[0])] = String(f[3]); });
  return m;
}

function galeriaAbierta_() {
  return PropertiesService.getScriptProperties().getProperty('galeria') === 'si';
}

/* ---------- utilidades ---------- */

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function limpia_(v, max) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'").slice(0, max || 80);
}

function idCarpeta_(url) {
  const m = String(url || '').match(/folders\/([-\w]{10,})/) || String(url || '').match(/[?&]id=([-\w]{10,})/);
  return m ? m[1] : '';
}

/** Lista las imágenes y PDF de una carpeta y las clasifica por prefijo. */
function imagenes_(urlCarpeta) {
  const id = idCarpeta_(urlCarpeta);
  const out = { cs: [], bocetos: [], refs: [] };
  if (!id) return out;
  const cache = CacheService.getScriptCache();
  const enCache = cache.get('img_' + id);
  if (enCache) return JSON.parse(enCache);
  let it;
  try { it = DriveApp.getFolderById(id).getFiles(); } catch (e) { return out; }
  while (it.hasNext()) {
    const f = it.next();
    const tipo = f.getMimeType();
    if (tipo.indexOf('image/') !== 0 && tipo !== 'application/pdf') continue;
    const nombre = f.getName();
    const o = { id: f.getId(), n: nombre, t: f.getLastUpdated().getTime() };
    const up = nombre.toUpperCase();
    if (up.indexOf('CS_') === 0) out.cs.push(o);
    else if (up.indexOf('BOCETO_') === 0) out.bocetos.push(o);
    else out.refs.push(o);
  }
  ['cs', 'bocetos', 'refs'].forEach(k => out[k].sort((a, b) => b.t - a.t));
  try { cache.put('img_' + id, JSON.stringify(out), 60); } catch (e) {}
  return out;
}

function ficha_(fila, conCarpeta) {
  const o = {
    id: fila.id, fam: fila.fam, alumna: fila.alumna, personaje: fila.personaje,
    tieneCarpeta: !!idCarpeta_(fila.carpeta), img: imagenes_(fila.carpeta)
  };
  if (conCarpeta) o.carpeta = fila.carpeta;
  return o;
}


/** Miniaturas servidas por el script con los permisos del profesor, para no depender de cómo estén compartidos los archivos. */
function miniatura_(id, lado) {
  const cache = CacheService.getScriptCache();
  const clave = 'th_' + lado + '_' + id;
  const enCache = cache.get(clave);
  if (enCache) return JSON.parse(enCache);
  const token = ScriptApp.getOAuthToken();
  const meta = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?fields=thumbnailLink,mimeType&supportsAllDrives=true',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (meta.getResponseCode() !== 200) return null;
  const link = JSON.parse(meta.getContentText()).thumbnailLink;
  if (!link) return null;
  const url = link.replace(/=s\d+(-[a-z0-9-]*)?$/i, '') + '=s' + lado;
  const img = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (img.getResponseCode() !== 200) return null;
  const blob = img.getBlob();
  const out = { mime: blob.getContentType() || 'image/jpeg', data: Utilities.base64Encode(blob.getBytes()) };
  const json = JSON.stringify(out);
  if (json.length < 95000) { try { cache.put(clave, json, 21600); } catch (e) {} }
  return out;
}

/* ---------- web ---------- */

function datosPublicos_() {
  const cache = CacheService.getScriptCache();
  const enCache = cache.get('publico');
  if (enCache) return enCache;
  const claves = claves_();
  const claims = leer_().map(c => {
    const o = Object.assign({}, c);
    delete o.carpeta;          // el enlace de la carpeta nunca se publica
    return o;
  });
  const json = JSON.stringify({ ok: true, claims: claims, conClave: Object.keys(claves), galeria: galeriaAbierta_() });
  try { cache.put('publico', json, 20); } catch (e) {}
  return json;
}

function limpiaCache_() {
  try { CacheService.getScriptCache().remove('publico'); } catch (e) {}
}

function doGet() {
  return ContentService.createTextOutput(datosPublicos_()).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const d = JSON.parse((e.postData && e.postData.contents) || '{}');

  /* --- lecturas: no necesitan bloquear la hoja --- */
  if (d.action === 'ficha') {
    const filas = leer_();
    const fila = filas.filter(f => String(f.id) === String(d.id))[0];
    if (!fila) return salida_({ ok: false, error: 'rol' });
    const esProfe = d.pin && d.pin === PIN_DOCENTE;
    if (!esProfe) {
      const clave = claves_()[String(d.id)];
      if (!clave || clave !== String(d.clave || '')) return salida_({ ok: false, error: 'clave' });
    }
    return salida_({ ok: true, ficha: ficha_(fila, true), galeria: galeriaAbierta_() });
  }

  if (d.action === 'thumbs') {
    const esProfe = d.pin && d.pin === PIN_DOCENTE;
    if (!esProfe) {
      const clave = claves_()[String(d.id)];
      if (!clave || clave !== String(d.clave || '')) return salida_({ ok: false, error: 'clave' });
    }
    const ids = (Array.isArray(d.ids) ? d.ids : []).slice(0, 8).map(String);
    const lado = Math.min(Math.max(parseInt(d.lado, 10) || 640, 200), 1600);
    const imgs = {};
    ids.forEach(id => { try { imgs[id] = miniatura_(id, lado); } catch (e) { imgs[id] = null; } });
    return salida_({ ok: true, imgs: imgs });
  }

  if (d.action === 'galeria') {
    const esProfe = d.pin && d.pin === PIN_DOCENTE;
    if (!esProfe) {
      if (!galeriaAbierta_()) return salida_({ ok: false, error: 'cerrada' });
      const clave = claves_()[String(d.id)];
      if (!clave || clave !== String(d.clave || '')) return salida_({ ok: false, error: 'clave' });
    }
    const fichas = leer_().filter(f => f.id && f.alumna).map(f => ficha_(f, esProfe));
    return salida_({ ok: true, fichas: fichas, galeria: galeriaAbierta_() });
  }

  /* --- escrituras --- */
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  limpiaCache_();
  try {
    const sh = hoja_();
    const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]));

    if (d.action === 'reserve') {
      if (!d.id || !d.alumna) return salida_({ ok: false, error: 'datos' });
      if (ids.indexOf(String(d.id)) !== -1) return salida_({ ok: false, error: 'ocupado' });
      sh.appendRow([
        limpia_(d.id, 60), limpia_(d.fam, 20), limpia_(d.role, 60), d.custom === true,
        limpia_(d.roleName, 60), limpia_(d.alumna, 60), limpia_(d.personaje, 60),
        limpia_(d.estilo, 20), new Date(), ''
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

    if (d.action === 'clave') {
      const clave = String(d.clave || '');
      if (!d.id || clave.length < 4) return salida_({ ok: false, error: 'corta' });
      if (ids.indexOf(String(d.id)) === -1) return salida_({ ok: false, error: 'rol' });
      if (claves_()[String(d.id)]) return salida_({ ok: false, error: 'existe' });
      const fila = leer_().filter(f => String(f.id) === String(d.id))[0];
      hojaClaves_().appendRow([limpia_(d.id, 60), limpia_(fila.personaje, 60), limpia_(fila.alumna, 60), limpia_(clave, 60), new Date()]);
      return salida_({ ok: true });
    }

    if (d.action === 'abrirGaleria') {
      if (d.pin !== PIN_DOCENTE) return salida_({ ok: false, error: 'pin' });
      PropertiesService.getScriptProperties().setProperty('galeria', d.abierta ? 'si' : 'no');
      return salida_({ ok: true, galeria: !!d.abierta });
    }

    if (d.action === 'release') {
      if (d.pin !== PIN_DOCENTE) return salida_({ ok: false, error: 'pin' });
      const fila = ids.indexOf(String(d.id));
      if (fila > 0) sh.deleteRow(fila + 1);
      return salida_({ ok: true });
    }

    return salida_({ ok: false, error: 'accion' });
  } finally {
    limpiaCache_();
    lock.releaseLock();
  }
}
