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
 * Escalera de precios: S/ 10 -> 20 -> 30 -> 40 (si dice No en 10, se pregunta 5).
 * DAP_MIN = ultimo precio aceptado; DAP_MAX = primer precio rechazado (vacio si acepta 40).
 */

const SHEET_NAME = "Respuestas";
const META = 100;                       // meta de encuestas

const HEADERS = ["FECHA_HORA","ID","DISTRITO","EDAD","MUJER","POLIMED","DEPEND","TRABAJA",
  "PERMISO","VECES","MOVILIDAD","HORAS","SINSTOCK","WHATSAPP","DELIVERY","INGRESO",
  "SI_10","SI_5","SI_20","SI_30","SI_40","DAP_MIN","DAP_MAX","DAP_ABIERTA","SEGURO",
  "MOTIVO_NO","PROTESTA","AVISAR","CELULAR"];

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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const total = Math.max(getSheet_().getLastRow() - 1, 0);
  return json_({ ok: true, total: total, meta: META });
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
