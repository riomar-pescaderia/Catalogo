/* ============================================================
   MAPA DE PUNTOS DE VENTA — sincronización desde Ruta Fría
   ============================================================
   Se pega en la planilla PÚBLICA "Puntos de venta":
   Extensiones → Apps Script.

   Qué hace: cada hora le pide a Ruta Fría la lista de comercios
   marcados con "Mostrar este comercio en el mapa público del
   catálogo" y la copia en la primera pestaña de esta planilla. El
   catálogo web lee esta planilla (nunca Ruta Fría).

   Seguridad:
   - Ruta Fría solo entrega esa lista con una clave secreta. La clave
     NO va escrita en este código: se guarda en las propiedades del
     script (ver paso 1), que solo ven los editores de la planilla.
   - Ruta Fría solo manda datos pensados para publicar (nombre
     comercial, dirección, ubicación, teléfono/WhatsApp del comercio,
     logo, horario y nombres de productos). Nunca CUIT, notas, saldos,
     contactos personales, cantidades ni precios.
   - Si Ruta Fría no responde, la planilla queda como estaba (no se
     borra nada).

   Puesta en marcha (una sola vez):
   1. En el editor de Apps Script: ⚙ Configuración del proyecto →
      Propiedades del script → Agregar propiedad:
        Propiedad: MAPA_TOKEN   Valor: (la clave que se cargó en Render)
   2. Elegir la función "activarSincronizacion" y tocar Ejecutar.
      Aceptar los permisos que pide Google.
   ============================================================ */

const URL_RUTA_FRIA = "https://ruta-fria.onrender.com/api/mapa/puntos-venta";
const ENCABEZADOS = ["ID", "Nombre", "Dirección", "Latitud", "Longitud", "Teléfono",
                     "WhatsApp", "Logo", "Horario", "Productos", "Actualizado"];

/* Ejecutar una sola vez: programa la sincronización cada hora (sin
   duplicarla si se ejecuta de nuevo) y sincroniza ya mismo. */
function activarSincronizacion() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sincronizarPuntosDeVenta") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sincronizarPuntosDeVenta").timeBased().everyHours(1).create();
  sincronizarPuntosDeVenta();
}

function sincronizarPuntosDeVenta() {
  const clave = PropertiesService.getScriptProperties().getProperty("MAPA_TOKEN");
  if (!clave) throw new Error("Falta la propiedad del script MAPA_TOKEN (ver paso 1 arriba).");

  const datos = pedirARutaFria_(clave);
  if (!datos) return; // Ruta Fría no respondió: se deja la planilla como estaba

  const filas = (datos.puntos || []).map(function (p) {
    return [
      p.id, p.nombre, p.direccion, p.lat, p.lng, p.telefono,
      p.whatsapp, p.logo, p.horario, (p.productos || []).join(", "), datos.actualizado
    ].map(limpiarCelda_);
  });

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  hoja.clearContents();
  hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);
  if (filas.length) hoja.getRange(2, 1, filas.length, ENCABEZADOS.length).setValues(filas);
}

/* Ruta Fría está en el plan gratuito de Render y "se duerme": el
   primer pedido puede tardar en despertarlo, así que se reintenta. */
function pedirARutaFria_(clave) {
  for (let intento = 1; intento <= 3; intento++) {
    try {
      const r = UrlFetchApp.fetch(URL_RUTA_FRIA, {
        headers: { Authorization: "Bearer " + clave },
        muteHttpExceptions: true,
      });
      const codigo = r.getResponseCode();
      if (codigo === 200) return JSON.parse(r.getContentText());
      if (codigo === 401 || codigo === 503) {
        throw new Error("Ruta Fría rechazó el pedido (" + codigo + "): revisar que MAPA_TOKEN sea igual en Render y en este script.");
      }
    } catch (e) {
      if (String(e.message).indexOf("Ruta Fría rechazó") === 0) throw e;
    }
    Utilities.sleep(20000);
  }
  console.warn("Ruta Fría no respondió; la planilla queda como estaba.");
  return null;
}

/* Un texto que empieza con =, +, - o @ se tomaría como fórmula en la
   planilla: se le antepone un apóstrofo para que quede como texto. */
function limpiarCelda_(v) {
  if (typeof v !== "string") return v;
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
