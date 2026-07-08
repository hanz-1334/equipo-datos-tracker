/**
 * Backend de "Notas de proyecto" para tracker_v2.html
 * ---------------------------------------------------
 * Guarda notas por proyecto en una Google Sheet, con columnas:
 *   id | fecha | autor | proyecto | texto | resuelto
 *
 * Cómo instalarlo:
 * 1. Crea una Google Sheet nueva (o usa una existente).
 * 2. En la primera fila (fila 1) escribe estos encabezados exactos:
 *    id | fecha | autor | proyecto | texto | resuelto
 * 3. Abre Extensiones → Apps Script en esa Sheet.
 * 4. Borra el contenido de Code.gs y pega todo este archivo.
 * 5. Guarda el proyecto (nómbralo como quieras, ej. "Notas Tracker").
 * 6. Click en "Implementar" → "Nueva implementación".
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario (o "Cualquiera con el enlace")
 * 7. Autoriza los permisos que pida Google (es tu propio script, es seguro).
 * 8. Copia la URL que te da ("URL de la aplicación web") — termina en /exec.
 * 9. Pega esa URL en la constante APPS_SCRIPT_URL dentro de tracker_v2.html.
 *
 * Nota de diseño: el POST se envía como text/plain (no application/json)
 * desde el HTML para evitar el preflight CORS que Apps Script no maneja bien.
 * Este script parsea el body manualmente con JSON.parse.
 */

const SHEET_NAME = "Sheet1"; // cambia esto si tu hoja tiene otro nombre

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readAllRows_() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * GET ?proyecto=ID_DEL_PROYECTO
 * Devuelve todas las notas de ese proyecto, ordenadas cronológicamente (más nueva al final).
 * Si no se pasa "proyecto", devuelve todas las notas (uso interno/debug).
 */
function doGet(e) {
  try {
    const proyecto = e.parameter.proyecto;
    let rows = readAllRows_();
    if (proyecto) {
      rows = rows.filter(r => String(r.proyecto) === String(proyecto));
    }
    rows.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
    return jsonResponse_({ ok: true, notas: rows });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

/**
 * POST body (text/plain, contenido JSON):
 * { "action": "create", "proyecto": "control-pagos", "autor": "Miguel", "texto": "..." }
 * { "action": "resolve", "id": "abc123" }
 * { "action": "delete",  "id": "abc123" }
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (payload.action === "create") {
      const id = Utilities.getUuid();
      const fecha = new Date().toISOString();
      sheet.appendRow([
        id,
        fecha,
        payload.autor || "Desconocido",
        payload.proyecto || "",
        payload.texto || "",
        false
      ]);
      return jsonResponse_({ ok: true, id: id, fecha: fecha });
    }

    if (payload.action === "resolve" || payload.action === "delete") {
      const data = sheet.getDataRange().getValues();
      const idColIndex = data[0].indexOf("id");
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idColIndex]) === String(payload.id)) {
          if (payload.action === "resolve") {
            const resueltoColIndex = data[0].indexOf("resuelto");
            sheet.getRange(i + 1, resueltoColIndex + 1).setValue(true);
          } else {
            sheet.deleteRow(i + 1);
          }
          return jsonResponse_({ ok: true });
        }
      }
      return jsonResponse_({ ok: false, error: "id no encontrado" });
    }

    return jsonResponse_({ ok: false, error: "acción no reconocida" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}
