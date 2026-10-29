// Backend de leads de Super Motos (Google Apps Script sobre una Google Sheet).
// Cambiá esta clave: es la que Carlos escribe para entrar al panel.
const CLAVE_PANEL = 'cambiar-esta-clave';

const COLS = ['id','fecha','nombre','nivel','puntaje','tipo','uso','pago','ingresos','clearing','entrega','cuando','zona','sucursal','whatsapp','estado','nota'];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Leads');
  if (!sh) { sh = ss.insertSheet('Leads'); sh.appendRow(COLS); sh.setFrozenRows(1); }
  return sh;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function filaDe_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let r = 0; r < ids.length; r++) if (String(ids[r][0]) === String(id)) return r + 2;
  return 0;
}

// La landing manda: {accion:'lead', ...datos} y {accion:'click', id}
// El panel manda:   {accion:'estado', clave, id, estado, nota}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const sh = hoja_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (d.accion === 'lead') {
      if (filaDe_(sh, d.id)) return json_({ok: true});
      sh.appendRow(COLS.map(c => c === 'fecha' ? new Date() : c === 'whatsapp' ? 'no' : c === 'estado' ? 'nuevo' : (d[c] ?? '')));
      return json_({ok: true});
    }
    const fila = filaDe_(sh, d.id);
    if (!fila) return json_({ok: false, error: 'no existe'});
    if (d.accion === 'click') {
      sh.getRange(fila, COLS.indexOf('whatsapp') + 1).setValue('sí');
      return json_({ok: true});
    }
    if (d.accion === 'estado') {
      if (d.clave !== CLAVE_PANEL) return json_({ok: false, error: 'clave'});
      sh.getRange(fila, COLS.indexOf('estado') + 1).setValue(d.estado);
      sh.getRange(fila, COLS.indexOf('nota') + 1).setValue(d.nota || '');
      return json_({ok: true});
    }
    return json_({ok: false, error: 'accion'});
  } finally { lock.releaseLock(); }
}

// El panel lee: GET ?clave=...
function doGet(e) {
  if (e.parameter.clave !== CLAVE_PANEL) return json_({ok: false, error: 'clave'});
  const v = hoja_().getDataRange().getValues();
  const head = v.shift();
  return json_({ok: true, leads: v.map(r => Object.fromEntries(head.map((h, i) => [h, r[i]])))});
}
