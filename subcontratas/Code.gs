/**
 * Recam Subcontratas · puente entre el Google Sheet «Master subcontratas» y la página web.
 *
 * INSTALACIÓN (con la cuenta exteriorsrecam@gmail.com, que es la dueña del Sheet):
 *  1. Entra en https://script.google.com → «Nuevo proyecto» (o, desde el Sheet, Extensiones → Apps Script).
 *  2. Borra lo que haya, pega este archivo entero y guarda.
 *  3. (Opcional) Escribe una clave en TOKEN y la misma en la pestaña Datos de la página.
 *  4. Implementar → Nueva implementación → tipo «Aplicación web».
 *       Ejecutar como: Yo   ·   Quién tiene acceso: Cualquier usuario
 *  5. Autoriza los permisos, copia la URL que acaba en /exec y pégala en Datos → «Guardar y actualizar».
 *
 * Si cambias este código después: Implementar → Gestionar implementaciones → editar → Nueva versión.
 *
 * Uso diario: pega el volcado del RP en la primera hoja del Sheet (cabeceras en la fila 1).
 * La página guarda sus propios datos en dos hojas que crea sola, para que no se pierdan al sustituir el volcado:
 *   «Seguimiento»           → contactos, reclamaciones y promesas de fecha.
 *   «Contactos proveedores» → email y teléfono de cada proveedor.
 */
const SHEET_ID = '1VMv1lN1HI5AEgTQkOdnlMte75ixw-Znx8FUOxPfXLBw';   // Master subcontratas
const MASTER_TAB = '';            // nombre de la pestaña del volcado; vacío = la primera hoja que no sea de la página
const TOKEN = '';                 // déjalo vacío si no quieres clave
const SEG = 'Seguimiento';
const CON = 'Contactos proveedores';
const SEG_COLS = ['id', 'ts', 'key', 'prov', 'oc', 'linea', 'tipo', 'nota', 'promesa', 'user'];
const CON_COLS = ['prov', 'email', 'tel'];

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fmt_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;
}
function okToken_(t) { return !TOKEN || t === TOKEN; }
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function masterSheet_(ss) {
  if (MASTER_TAB) return ss.getSheetByName(MASTER_TAB);
  const own = [SEG, CON];
  return ss.getSheets().filter(function (s) { return own.indexOf(s.getName()) < 0; })[0];
}

function doGet(e) {
  try {
    if (!okToken_(e && e.parameter && e.parameter.token)) return out_({ error: 'Clave incorrecta' });
    const ss = ss_();
    const sh = masterSheet_(ss);
    if (!sh) return out_({ error: 'No se encuentra la hoja del volcado' });
    const values = sh.getDataRange().getValues().map(function (r) { return r.map(fmt_); });
    return out_({ values: values, seg: readTable_(ss, SEG, SEG_COLS), provs: readTable_(ss, CON, CON_COLS), updated: new Date().toISOString() });
  } catch (err) {
    return out_({ error: String(err) });
  }
}

function readTable_(ss, name, cols) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues().map(function (r) {
    const o = {};
    cols.forEach(function (c, i) { o[c] = fmt_(r[i]); });
    if (o.ts !== undefined) o.ts = String(o.ts);
    return o;
  });
}
function ensure_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    if (!okToken_(body.token)) return out_({ error: 'Clave incorrecta' });
    if (body.action !== 'push') return out_({ error: 'Acción desconocida' });
    const ss = ss_();
    let added = 0, upd = 0;

    const entries = body.entries || [];
    if (entries.length) {
      const sh = ensure_(ss, SEG, SEG_COLS);
      const have = {};
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[r[0]] = true; });
      entries.forEach(function (en) {
        if (have[en.id]) return;   // evita duplicados si se reintenta
        sh.appendRow(SEG_COLS.map(function (c) { return en[c] == null ? '' : en[c]; }));
        added++;
      });
    }

    const provs = body.provs || [];
    if (provs.length) {
      const sh = ensure_(ss, CON, CON_COLS);
      const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
      provs.forEach(function (p) {
        let at = -1;
        for (let i = 0; i < rows.length; i++) if (rows[i][0] === p.prov) { at = i; break; }
        if (at >= 0) { sh.getRange(at + 2, 1, 1, 3).setValues([[p.prov, p.email || '', p.tel || '']]); }
        else { sh.appendRow([p.prov, p.email || '', p.tel || '']); rows.push([p.prov]); }
        upd++;
      });
    }
    return out_({ ok: true, added: added, contacts: upd });
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
