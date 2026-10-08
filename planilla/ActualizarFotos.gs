/* ============================================================
   FOTOS DEL CATÁLOGO RIOMAR — botón "Actualizar fotos"
   ============================================================
   Se pega en la planilla de precios: Extensiones → Apps Script.

   Cómo se usa:
   1) Subí la foto a la carpeta de Drive "Fotos catálogo" con el
      código del producto como nombre (ej.: 2000.jpg, 2000.png).
   2) En la planilla: menú "Fotos catálogo" → "Actualizar fotos".
   3) El script completa la columna "Foto" con el link de cada foto
      cuyo nombre coincida con el código de la columna "Cod.".

   - Si un producto no tiene foto en la carpeta, su celda "Foto" no
     se toca (así se puede pegar un link a mano si hace falta).
   - Si hay dos archivos con el mismo código, se usa el más nuevo
     (para cambiar una foto alcanza con subir la nueva).

   Actualización automática (recomendado): en el editor de Apps
   Script elegí la función "activarActualizacionAutomatica" y tocá
   Ejecutar UNA sola vez. Desde ahí la columna "Foto" se actualiza
   sola cada hora, sin usar el menú. (El menú a veces no aparece si
   hay varias cuentas de Google abiertas en el mismo navegador.)
   ============================================================ */

const CARPETA_FOTOS_ID = "1jXAQuDt56-rFb1hLNyg0xvfSXm9uydTP";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Fotos catálogo")
    .addItem("Actualizar fotos", "actualizarFotos")
    .addToUi();
}

/* Ejecutar una sola vez: programa actualizarFotos() para que corra
   sola cada hora. Si se ejecuta de nuevo, no duplica el programa. */
function activarActualizacionAutomatica() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "actualizarFotos") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("actualizarFotos").timeBased().everyHours(1).create();
  actualizarFotos();
}

/* Avisos en pantalla: cuando corre sola (cada hora) no hay pantalla,
   así que si fallan los ignoramos. */
function avisar_(ss, texto, titulo) {
  try { ss.toast(texto, titulo, 6); } catch (e) {}
}

function actualizarFotos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheets()[0]; // la pestaña de precios (la primera)
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(function (h) { return String(h).trim().toLowerCase(); });

  const colCod = encabezados.findIndex(function (h) { return h.indexOf("cod") === 0; });
  if (colCod < 0) {
    avisar_(ss, 'No encontré la columna "Cod." en la fila 1.', "Fotos");
    return;
  }
  let colFoto = encabezados.indexOf("foto");
  if (colFoto < 0) {
    // Creamos la columna "Foto" a la derecha de la última columna usada
    colFoto = hoja.getLastColumn();
    hoja.getRange(1, colFoto + 1).setValue("Foto");
  }

  // Archivos de la carpeta: nombre sin extensión → link (el más nuevo gana)
  const fotos = {};
  const fechas = {};
  const archivos = DriveApp.getFolderById(CARPETA_FOTOS_ID).getFiles();
  while (archivos.hasNext()) {
    const a = archivos.next();
    if (a.getMimeType().indexOf("image/") !== 0) continue;
    const cod = a.getName().replace(/\.[^.]+$/, "").trim();
    const fecha = a.getLastUpdated().getTime();
    if (!fotos[cod] || fecha > fechas[cod]) {
      fotos[cod] = "https://drive.google.com/file/d/" + a.getId() + "/view";
      fechas[cod] = fecha;
    }
  }

  // Columna "Foto" completa de una sola vez (más rápido que celda por celda)
  const filas = Math.max(datos.length - 1, 0);
  if (!filas) return;
  const rango = hoja.getRange(2, colFoto + 1, filas, 1);
  const actuales = rango.getValues();
  let conFoto = 0;
  for (let i = 0; i < filas; i++) {
    const cod = String(datos[i + 1][colCod]).trim();
    if (cod && fotos[cod]) {
      actuales[i][0] = fotos[cod];
      conFoto++;
    }
  }
  rango.setValues(actuales);

  avisar_(ss, conFoto + " productos con foto (" + Object.keys(fotos).length + " fotos en la carpeta).",
          "Fotos actualizadas");
}
