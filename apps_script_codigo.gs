/**
 * CÓDIGO DE GOOGLE APPS SCRIPT — Captura de Mantenimiento Flexigrip
 * ===================================================================
 *
 * QUÉ HACE:
 *   Recibe los registros del formulario HTML y los guarda en la pestaña
 *   "Registros" de esta Hoja de Google. También se los regresa cuando
 *   el formulario los pide (al abrir, al dar "Actualizar", etc.).
 *
 * CÓMO INSTALARLO (una sola vez):
 *   1. Crea una Hoja de Google nueva (sheets.google.com → Hoja en blanco).
 *   2. Ve a Extensiones → Apps Script.
 *   3. Borra el código de ejemplo que aparece y pega TODO este archivo.
 *   4. Guarda (ícono de disquete o Ctrl+S). Ponle un nombre al proyecto,
 *      por ejemplo "Backend Mantenimiento".
 *   5. Arriba a la derecha: botón "Implementar" → "Nueva implementación".
 *   6. Junto a "Selecciona el tipo": ícono de engrane → elige "Aplicación web".
 *   7. Configuración:
 *        - Descripción: lo que quieras (ej. "v1")
 *        - Ejecutar como: "Yo" (tu cuenta)
 *        - Quién tiene acceso: "Cualquier usuario"
 *   8. Clic en "Implementar". Google te va a pedir autorizar permisos
 *      la primera vez (clic en "Autorizar acceso", elige tu cuenta,
 *      si sale una advertencia de "app no verificada" clic en
 *      "Configuración avanzada" → "Ir a [nombre del proyecto] (no seguro)"
 *      — es normal, es tu propio script, no de un tercero).
 *   9. Te va a mostrar una URL que termina en /exec. Cópiala completa.
 *  10. Pega esa URL en generar_formulario.py, en la línea:
 *        SHEET_WEBAPP_URL = "PEGA_AQUI_TU_URL_DE_APPS_SCRIPT"
 *      y vuelve a correr el script para regenerar el formulario.
 *
 * SI DESPUÉS EDITAS ESTE CÓDIGO:
 *   Tienes que volver a hacer "Implementar → Nueva implementación" cada
 *   vez que cambies algo aquí — si solo guardas, los cambios no se
 *   publican solos. (O usa "Administrar implementaciones" → lápiz de
 *   editar → sube la versión, para conservar la misma URL).
 */

const SHEET_NAME = 'Registros';

/**
 * Se ejecuta cuando el formulario pide los datos (fetch GET).
 * Regresa un JSON con el arreglo completo de registros.
 */
function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return respond_([]);
  }

  const jsonColumna = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  const registros = jsonColumna
    .map(function (fila) {
      try {
        return JSON.parse(fila[0]);
      } catch (err) {
        return null;
      }
    })
    .filter(function (r) { return r !== null; });

  return respond_(registros);
}

/**
 * Se ejecuta cuando el formulario guarda cambios (fetch POST).
 * Recibe el arreglo COMPLETO de registros y reemplaza el contenido
 * de la hoja con ese arreglo (así se mantienen sincronizados).
 */
function doPost(e) {
  try {
    const registros = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    if (registros.length > 0) {
      const filas = registros.map(function (r) {
        return [
          r.id,
          r.estado,
          r.maquina || '',
          JSON.stringify(r)
        ];
      });
      sheet.getRange(2, 1, filas.length, 4).setValues(filas);
    }

    return respond_({ ok: true, guardados: registros.length });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  }
}

/**
 * Devuelve la pestaña "Registros", creándola con encabezados si no existe.
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['ID', 'ESTADO', 'MAQUINA', 'DATA_JSON']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Arma la respuesta HTTP en formato JSON.
 */
function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
