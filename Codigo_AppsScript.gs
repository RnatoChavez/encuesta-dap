/**
 * CODIGO PARA GOOGLE APPS SCRIPT
 * Hoja de destino: "Encuesta DAP Entrega de Medicamentos - Medi Ruta-360"
 *
 * INSTALACION:
 * 1. Abre la Google Sheet -> Extensiones -> Apps Script.
 * 2. Borra el contenido de Code.gs y pega todo este archivo.
 * 3. Ejecuta la funcion "setupSheet" una vez (menu Ejecutar).
 *    - La primera vez pedira autorizacion: acepta los permisos.
 *    - Renombra la pestana a "Respuestas" y deja los encabezados en negrita.
 * 4. Implementar -> Nueva implementacion -> Tipo: Aplicacion web.
 *    - Ejecutar como: Yo (tu cuenta).
 *    - Quien tiene acceso: Cualquier usuario.
 *    - Clic en "Implementar" y copia la URL del Web App (termina en /exec).
 * 5. Pega esa URL en index.html, en la constante WEB_APP_URL
 *    (o enviasela a Claude para que la actualice en GitHub).
 *
 * Asignacion de versiones: doGet devuelve cuantas respuestas tiene cada
 * version (V1-V5) y cual es la siguiente, para que las 5 queden balanceadas.
 */

const SHEET_NAME = "Respuestas";
const META = 150;                       // meta total: 30 por version
const BIDS = [5, 10, 15, 20, 25];       // precio inicial por version (S/)

const HEADERS = ["FECHA_HORA","ID","VERSION","BID","SI_DAP","BID2","SI_DAP2","SEGURO",
  "MOTIVO_NO","PROTESTA","EDAD","MUJER","ESSALUD","SIS","POLIMED","DEPEND","TRABAJA",
  "PERMISO","VECES","MOVILIDAD","HORAS","SINSTOCK","WHATSAPP","DELIVERY","INGRESO",
  "DISTRITO","ENCUESTADOR","AVISAR","CELULAR"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function setupSheet() {
  const sh = getSheet_();
  sh.setName(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, HEADERS.length);
}

function countByVersion_(sh) {
  const counts = [0, 0, 0, 0, 0];
  const last = sh.getLastRow();
  if (last < 2) return counts;
  const col = HEADERS.indexOf("VERSION") + 1;
  sh.getRange(2, col, last - 1, 1).getValues().forEach(function (r) {
    const v = parseInt(r[0], 10);
    if (v >= 1 && v <= 5) counts[v - 1]++;
  });
  return counts;
}

function nextVersion_(counts) {
  let best = 0;
  for (let i = 1; i < counts.length; i++) if (counts[i] < counts[best]) best = i;
  return best + 1;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const counts = countByVersion_(getSheet_());
  const total = counts.reduce(function (a, b) { return a + b; }, 0);
  const next = nextVersion_(counts);
  return json_({ ok: true, total: total, meta: META, porVersion: counts,
                 siguiente: next, bid: BIDS[next - 1] });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    const id = Math.max(sh.getLastRow() - 1, 0) + 1;
    const row = HEADERS.map(function (h) {
      if (h === "FECHA_HORA") return new Date();
      if (h === "ID") return id;
      const v = d[h];
      return (v === undefined || v === null) ? "" : v;
    });
    sh.appendRow(row);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
