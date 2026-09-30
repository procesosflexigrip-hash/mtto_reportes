/**
 * CÓDIGO DE GOOGLE APPS SCRIPT — Captura de Mantenimiento Flexigrip
 * ===================================================================
 *
 * QUÉ HACE:
 *   Recibe los registros del formulario HTML y los guarda en la pestaña
 *   "Registros" de esta Hoja de Google. También se los regresa cuando
 *   el formulario los pide (al abrir, al dar "Actualizar", etc.).
 *   Si hay un bot de Telegram configurado, avisa al grupo cuando:
 *     - se genera un reporte nuevo (estado Pendiente)
 *     - se asignan o suman técnicos a un servicio
 *   Si existe TELEGRAM_CODIGOS_CHAT_ID, un segundo grupo (privado)
 *   recibe solo el número de servicio, la máquina y el código de
 *   confirmación al crear el reporte.
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
 *
 * ===================================================================
 * TELEGRAM — GRUPO DE MANTENIMIENTO (Modo A)
 * ===================================================================
 *
 * El token del bot NUNCA va en index.html ni en GitHub. Solo aquí,
 * en Propiedades del script.
 *
 * 1) Crear el bot
 *    - Abre Telegram y busca @BotFather
 *    - Envía /newbot, elige nombre y username (debe terminar en bot)
 *    - Copia el token (parece: 123456:ABC-xxxxx)
 *
 * 2) Crear el grupo
 *    - Nuevo grupo, por ejemplo "Mantenimiento Flexigrip"
 *    - Invita a supervisores y técnicos
 *    - Añade el bot al grupo (Buscar → el username del bot)
 *    - Envíale un mensaje cualquiera al grupo (para que Telegram
 *      registre el chat)
 *
 * 3) Guardar secretos en Apps Script
 *    - En el editor: ⚙ Configuración del proyecto (o Archivo en
 *      versiones antiguas) → Propiedades del script → Añadir
 *    - TELEGRAM_BOT_TOKEN  = el token de BotFather
 *    - TELEGRAM_CHAT_ID    = el id del grupo (número negativo,
 *      ej. -1001234567890)
 *
 * 4) Cómo sacar el chat_id del grupo
 *    - En el editor, selecciona la función mostrarChatsTelegram
 *      y pulsa Ejecutar (la primera vez autoriza UrlFetchApp).
 *    - Ve a Ver → Registros. Aparecen los chats recientes.
 *    - Copia el id del grupo y pégalo en TELEGRAM_CHAT_ID.
 *    - Si no sale nada: escribe un mensaje en el grupo y vuelve
 *      a ejecutar mostrarChatsTelegram.
 *
 * 5) Probar
 *    - Ejecuta la función probarTelegram. Debe llegar un mensaje
 *      de prueba al grupo.
 *
 * 6) Publicar este código
 *    - Implementar → Administrar implementaciones → lápiz →
 *      Versión: Nueva → Implementar
 *    - Conserva la misma URL /exec si editas la implementación
 *      existente.
 *
 * 7) Probar los dos eventos reales
 *    - Genera un reporte en el formulario → debe llegar "Nuevo reporte"
 *    - Como supervisor, asigna un técnico → debe llegar "Servicio asignado"
 *
 * 8) Grupo privado de códigos (solo operaciones y gerentes)
 *    - Crea OTRO grupo, por ejemplo "Códigos mtto — interno"
 *    - Invita SOLO a quien debe ver el código. No técnicos,
 *      no operadores, no supervisores de piso.
 *    - Añade el MISMO bot y escribe un mensaje en ese grupo
 *    - Ejecuta mostrarChatsTelegram y copia el chat_id (negativo)
 *    - Propiedad: TELEGRAM_CODIGOS_CHAT_ID = ese id
 *    - Ejecuta probarTelegramCodigos — debe llegar al grupo chico,
 *      NO al de mantenimiento
 *    - Al generar un reporte, ese grupo recibe: servicio, máquina
 *      y el código de 4 dígitos. El grupo grande no lo ve.
 *
 * Alta de gente nueva al grupo de mantenimiento: solo invitarlos.
 * No hay que tocar código ni Excel.
 *
 * ===================================================================
 * TELEGRAM — MENSAJES PRIVADOS (Modo B, no activo)
 * ===================================================================
 *
 * Este archivo crea la pestaña "Contactos" (NOMBRE, ROL, CHAT_ID)
 * para cuando se quiera avisar solo al técnico asignado.
 * Hoy NO se usa: todos los avisos van al grupo.
 *
 * El día que se active:
 *   - Cada persona abre el bot y pulsa Inicio (/start)
 *   - Ejecuta mostrarChatsTelegram y anota el chat_id privado
 *   - Pega NOMBRE (igual que en validacion.xlsx), ROL (tecnico o
 *     supervisor) y CHAT_ID en la pestaña Contactos
 *   - Telegram no deja al bot escribir primero a quien no le
 *     haya hablado
 */

const SHEET_NAME = 'Registros';
const CONTACTOS_SHEET = 'Contactos';
const RESPALDOS_SHEET = 'Respaldos';
const PROTECCION_DESC = 'Hoja escrita por el script — no editar a mano';

/**
 * Se ejecuta cuando el formulario pide los datos (fetch GET).
 * Regresa un JSON con el arreglo completo de registros.
 */
function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    asegurarContactos_();
    return respond_(leerRegistros_());
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Se ejecuta cuando el formulario guarda cambios (fetch POST).
 * Recibe el arreglo COMPLETO de registros y reemplaza el contenido
 * de la hoja con ese arreglo (así se mantienen sincronizados).
 * Después intenta avisar a Telegram; si el aviso falla, el guardado
 * igual queda hecho.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const registros = JSON.parse(e.postData.contents);
    if (!Array.isArray(registros)) {
      return respond_({ ok: false, error: 'El cuerpo debe ser un arreglo de registros' });
    }

    asegurarContactos_();
    const anteriores = leerRegistros_();
    const validacion = validarReemplazo_(anteriores, registros);
    if (!validacion.ok) {
      return respond_({
        ok: false,
        error: validacion.error,
        idsFaltantes: validacion.idsFaltantes || []
      });
    }
    const maquinas = validarMaquinasDisponibles_(anteriores, registros);
    if (!maquinas.ok) {
      return respond_({
        ok: false,
        error: maquinas.error,
        servicioExistente: maquinas.servicioExistente || null
      });
    }
    const tecnicos = validarTecnicosDisponibles_(anteriores, registros);
    if (!tecnicos.ok) {
      return respond_({
        ok: false,
        error: tecnicos.error,
        servicioExistente: tecnicos.servicioExistente || null
      });
    }
    const eventos = detectarEventosTelegram_(anteriores, registros);

    respaldarRegistros_(anteriores);
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

    var telegram = { enviados: 0, error: null };
    try {
      telegram.enviados = notificarTelegram_(eventos);
    } catch (errTg) {
      telegram.error = String(errTg);
    }

    return respond_({
      ok: true,
      guardados: registros.length,
      telegram: telegram
    });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
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
  asegurarBlindaje_(sheet);
  return sheet;
}

/**
 * Pestaña lista para Modo B (mensajes privados). No se usa en el aviso
 * al grupo. Columnas: NOMBRE, ROL, CHAT_ID.
 */
function asegurarContactos_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONTACTOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONTACTOS_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['NOMBRE', 'ROL', 'CHAT_ID']]);
    sheet.setFrozenRows(1);
    sheet.getRange('A2').setNote(
      'Modo B (aún no activo). NOMBRE igual que en el formulario. ' +
      'ROL: supervisor o tecnico. CHAT_ID: el id privado tras /start al bot.'
    );
  }
  return sheet;
}

function leerRegistros_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const jsonColumna = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  return jsonColumna
    .map(function (fila) {
      try {
        return JSON.parse(fila[0]);
      } catch (err) {
        return null;
      }
    })
    .filter(function (r) { return r !== null; });
}

/**
 * Evita que un navegador que no pudo cargar la hoja la reemplace con una
 * lista vacía o incompleta. Este sistema no tiene una acción de eliminar,
 * por lo que ningún guardado válido debe perder IDs existentes.
 */
function validarReemplazo_(anteriores, nuevos) {
  const idsNuevos = {};
  const nuevosPorId = {};
  const duplicados = [];
  (nuevos || []).forEach(function (r) {
    if (!r || r.id == null) return;
    const id = String(r.id);
    if (idsNuevos[id]) duplicados.push(r.id);
    idsNuevos[id] = true;
    nuevosPorId[id] = r;
  });

  if (duplicados.length > 0) {
    return {
      ok: false,
      error: 'Guardado rechazado: hay IDs duplicados (' + duplicados.join(', ') + ').'
    };
  }

  const faltantes = (anteriores || [])
    .filter(function (r) { return r && r.id != null && !idsNuevos[String(r.id)]; })
    .map(function (r) { return r.id; });

  if (faltantes.length > 0) {
    return {
      ok: false,
      error: 'Guardado rechazado para proteger los datos: faltan registros que ya existen en la hoja.',
      idsFaltantes: faltantes
    };
  }

  const colisiones = (anteriores || [])
    .filter(function (anterior) {
      if (!anterior || anterior.id == null) return false;
      const nuevo = nuevosPorId[String(anterior.id)];
      if (!nuevo) return false;
      if (anterior.creadoEn && nuevo.creadoEn) {
        return String(anterior.creadoEn) !== String(nuevo.creadoEn);
      }
      return huellaRegistro_(anterior) !== huellaRegistro_(nuevo);
    })
    .map(function (r) { return r.id; });

  if (colisiones.length > 0) {
    return {
      ok: false,
      error: 'Guardado rechazado: otro navegador ya utilizó uno de estos IDs. Actualiza antes de continuar.',
      idsFaltantes: colisiones
    };
  }

  return { ok: true };
}

/**
 * Identidad de origen de un registro. Los registros viejos no tienen
 * creadoEn, así que se comparan por lo que se capturó al solicitarlos:
 * si eso cambia bajo el mismo id, son dos servicios distintos.
 */
function huellaRegistro_(r) {
  if (!r) return '';
  return [
    String(r.fecha || ''),
    String(r.horaSolicita || ''),
    String(r.maquina || '').trim().toUpperCase()
  ].join('|');
}

/**
 * Trabajos generales (lámparas, pintura, etc.) no se amarran a una
 * máquina: pueden existir varios servicios abiertos a la vez.
 */
function esMaquinaSinBloqueo_(maquina) {
  const clave = String(maquina || '').trim().toUpperCase();
  return clave === 'SERVICIOS' || clave === 'NO APLICA';
}

/**
 * Segunda barrera para la regla "una máquina, un servicio abierto".
 * El navegador también lo valida, pero aquí se evita que dos equipos que
 * capturan al mismo tiempo abran dos tickets para la misma máquina.
 * SERVICIOS (y el nombre viejo NO APLICA) queda fuera: no ocupa ni se rechaza.
 */
function validarMaquinasDisponibles_(anteriores, nuevos) {
  const anterioresPorId = mapaPorId_(anteriores);
  const ocupadas = {};

  (anteriores || []).forEach(function (r) {
    if (!r || !r.maquina || r.estado === 'Liberado') return;
    if (esMaquinaSinBloqueo_(r.maquina)) return;
    ocupadas[String(r.maquina).trim().toUpperCase()] = r;
  });

  for (var i = 0; i < (nuevos || []).length; i++) {
    const r = nuevos[i];
    if (!r || r.id == null || anterioresPorId[String(r.id)] || !r.maquina) continue;
    if (esMaquinaSinBloqueo_(r.maquina)) continue;
    const clave = String(r.maquina).trim().toUpperCase();
    if (ocupadas[clave]) {
      const existente = ocupadas[clave];
      return {
        ok: false,
        error: 'La máquina ya tiene el servicio #' + existente.id +
          ' en estado "' + existente.estado + '". Debe liberarse antes de crear otro.',
        servicioExistente: existente.id
      };
    }
    if (r.estado !== 'Liberado') ocupadas[clave] = r;
  }

  return { ok: true };
}

function tecnicoOcupaServicio_(r) {
  return !!r && (r.estado === 'Asignado' || r.estado === 'En reparación');
}

/**
 * Un técnico en Asignado o En reparación no puede figurar en otro
 * servicio a la vez. Pausado, Cerrado y Liberado no lo ocupan.
 */
function validarTecnicosDisponibles_(anteriores, nuevos) {
  const vistos = {};
  for (var i = 0; i < (nuevos || []).length; i++) {
    const r = nuevos[i];
    if (!tecnicoOcupaServicio_(r)) continue;
    const lista = r.tecnicosActivos || [];
    for (var j = 0; j < lista.length; j++) {
      const nombre = String(lista[j] || '').trim();
      if (!nombre) continue;
      const clave = nombre.toUpperCase();
      if (vistos[clave] && String(vistos[clave].id) !== String(r.id)) {
        return {
          ok: false,
          error: 'El técnico ' + nombre + ' ya está en el servicio #' + vistos[clave].id +
            ' (' + vistos[clave].estado + '). Debe quedar libre antes de asignarlo a otro.',
          servicioExistente: vistos[clave].id
        };
      }
      vistos[clave] = r;
    }
  }
  return { ok: true };
}

/**
 * Conserva una instantánea antes de cada reemplazo aceptado. La pestaña
 * Respaldos permite recuperar el JSON aunque el historial de Sheets no esté
 * disponible. No se respalda cuando la hoja todavía está vacía.
 */
function respaldarRegistros_(registros) {
  if (!registros || registros.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESPALDOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RESPALDOS_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([
      ['RESPALDO_EN', 'ID', 'ESTADO', 'MAQUINA', 'DATA_JSON']
    ]);
    sheet.setFrozenRows(1);
  }

  const momento = new Date();
  const filas = registros.map(function (r) {
    return [
      momento,
      r.id,
      r.estado || '',
      r.maquina || '',
      JSON.stringify(r)
    ];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, filas.length, 5).setValues(filas);
}

function mapaPorId_(registros) {
  const mapa = {};
  (registros || []).forEach(function (r) {
    if (r && r.id != null) {
      mapa[String(r.id)] = r;
    }
  });
  return mapa;
}

function listaTecnicos_(arr) {
  return (arr || []).map(function (t) { return String(t); });
}

function tecnicosNuevos_(antes, ahora) {
  const vistos = {};
  listaTecnicos_(antes).forEach(function (t) { vistos[t] = true; });
  return listaTecnicos_(ahora).filter(function (t) { return !vistos[t]; });
}

function detectarEventosTelegram_(anteriores, actuales) {
  const eventos = [];
  const mapaAntes = mapaPorId_(anteriores);

  (actuales || []).forEach(function (r) {
    if (!r || r.id == null) return;
    const prev = mapaAntes[String(r.id)];

    if (!prev) {
      eventos.push({ tipo: 'nuevo', rec: r });
      return;
    }

    const estadoAhora = r.estado || '';
    const estadoAntes = prev.estado || '';
    const nuevos = tecnicosNuevos_(prev.tecnicosActivos, r.tecnicosActivos);

    if (estadoAhora === 'Asignado' && estadoAntes !== 'Asignado') {
      eventos.push({ tipo: 'asignado', rec: r, tecnicos: listaTecnicos_(r.tecnicosActivos) });
      return;
    }

    if (nuevos.length > 0 && (estadoAhora === 'Asignado' || estadoAhora === 'En reparación' || estadoAhora === 'Pausado')) {
      eventos.push({ tipo: 'asignado', rec: r, tecnicos: nuevos });
    }
  });

  return eventos;
}

function textoEvento_(ev) {
  const r = ev.rec || {};
  const id = r.id != null ? r.id : '?';
  const maquina = r.maquina || 'sin máquina';
  const area = r.area || 'sin área';
  const prioridad = r.prioridad || '—';
  const problema = r.problema || 'sin descripción';
  const solicitante = r.solicitante || '—';
  const hora = r.horaSolicita || r.horaAsignado || '';

  if (ev.tipo === 'nuevo') {
    return [
      'Nuevo reporte #' + id,
      'Máquina: ' + maquina + ' (' + area + ')',
      'Prioridad: ' + prioridad,
      'Solicitante: ' + solicitante,
      hora ? ('Hora: ' + hora) : '',
      'Problema: ' + problema
    ].filter(Boolean).join('\n');
  }

  const tecnicos = (ev.tecnicos && ev.tecnicos.length)
    ? ev.tecnicos.join(', ')
    : (listaTecnicos_(r.tecnicosActivos).join(', ') || '—');

  return [
    'Servicio asignado #' + id,
    'Máquina: ' + maquina + ' (' + area + ')',
    'Técnico(s): ' + tecnicos,
    'Prioridad: ' + prioridad,
    'Problema: ' + problema
  ].join('\n');
}

function propsTelegram_() {
  const props = PropertiesService.getScriptProperties();
  return {
    token: (props.getProperty('TELEGRAM_BOT_TOKEN') || '').trim(),
    chatId: (props.getProperty('TELEGRAM_CHAT_ID') || '').trim(),
    chatIdCodigos: (props.getProperty('TELEGRAM_CODIGOS_CHAT_ID') || '').trim()
  };
}

function textoCodigoConfirmacion_(r) {
  const rec = r || {};
  const id = rec.id != null ? rec.id : '?';
  const maquina = rec.maquina || 'sin máquina';
  const codigo = rec.codigoConfirmacion || 'sin código';
  return [
    'Servicio #' + id,
    'Máquina: ' + maquina,
    'Código: ' + codigo
  ].join('\n');
}

function enviarTelegram_(texto, chatId) {
  const cfg = propsTelegram_();
  const destino = (chatId == null || chatId === '') ? cfg.chatId : String(chatId).trim();
  if (!cfg.token || !destino) {
    return { ok: false, skip: true, error: 'Faltan TELEGRAM_BOT_TOKEN o el chat_id de destino en Propiedades del script' };
  }
  const url = 'https://api.telegram.org/bot' + cfg.token + '/sendMessage';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: destino,
      text: texto,
      disable_web_page_preview: true
    }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    return { ok: false, error: 'HTTP ' + code + ' ' + body };
  }
  return { ok: true };
}

function notificarTelegram_(eventos) {
  if (!eventos || eventos.length === 0) return 0;
  const cfg = propsTelegram_();
  if (!cfg.token) return 0;

  var enviados = 0;
  eventos.forEach(function (ev) {
    if (cfg.chatId) {
      const resultado = enviarTelegram_(textoEvento_(ev), cfg.chatId);
      if (resultado.ok) enviados += 1;
    }
    if (ev && ev.tipo === 'nuevo' && cfg.chatIdCodigos) {
      const codigo = enviarTelegram_(textoCodigoConfirmacion_(ev.rec), cfg.chatIdCodigos);
      if (codigo.ok) enviados += 1;
    }
  });
  return enviados;
}

/**
 * Envía un mensaje de prueba al grupo. Ejecutar a mano desde el editor
 * (selecciona esta función → Ejecutar). Sirve para autorizar UrlFetchApp
 * y comprobar token + chat_id.
 */
function probarTelegram() {
  const r = enviarTelegram_('Prueba de avisos — Mantenimiento Flexigrip.\nSi lees esto, el bot y el grupo están bien configurados.');
  if (r.skip) {
    throw new Error(r.error);
  }
  if (!r.ok) {
    throw new Error(r.error);
  }
  Logger.log('Mensaje de prueba enviado al grupo de mantenimiento.');
}

/**
 * Prueba el grupo PRIVADO de códigos. No debe llegar al grupo de
 * técnicos. Ejecutar a mano después de guardar TELEGRAM_CODIGOS_CHAT_ID.
 */
function probarTelegramCodigos() {
  const cfg = propsTelegram_();
  if (!cfg.chatIdCodigos) {
    throw new Error('Falta TELEGRAM_CODIGOS_CHAT_ID en Propiedades del script.');
  }
  const r = enviarTelegram_(
    'Prueba de códigos — grupo interno.\nSi lees esto, los códigos de confirmación llegarán aquí (servicio + código).',
    cfg.chatIdCodigos
  );
  if (r.skip) {
    throw new Error(r.error);
  }
  if (!r.ok) {
    throw new Error(r.error);
  }
  Logger.log('Mensaje de prueba enviado al grupo de códigos.');
}

/**
 * Lista chats recientes del bot para copiar el TELEGRAM_CHAT_ID.
 * Escribe un mensaje en el grupo justo antes de ejecutarla.
 */
function mostrarChatsTelegram() {
  const cfg = propsTelegram_();
  if (!cfg.token) {
    throw new Error('Falta TELEGRAM_BOT_TOKEN en Propiedades del script.');
  }
  const url = 'https://api.telegram.org/bot' + cfg.token + '/getUpdates';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());
  if (!data.ok) {
    throw new Error('Telegram: ' + res.getContentText());
  }
  const vistos = {};
  (data.result || []).forEach(function (upd) {
    const c = (upd.message && upd.message.chat)
      || (upd.my_chat_member && upd.my_chat_member.chat)
      || (upd.channel_post && upd.channel_post.chat);
    if (!c) return;
    const key = String(c.id);
    if (vistos[key]) return;
    vistos[key] = true;
    Logger.log(
      'chat_id=' + c.id +
      ' tipo=' + (c.type || '') +
      ' titulo=' + (c.title || c.username || c.first_name || '')
    );
  });
  if (Object.keys(vistos).length === 0) {
    Logger.log('No hay chats. Escribe en el grupo (con el bot dentro) y vuelve a ejecutar.');
  }
}

/**
 * Lectura de Contactos para un futuro Modo B. Hoy no se llama desde doPost.
 */
function leerContactos_() {
  const sheet = asegurarContactos_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const valores = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return valores
    .map(function (fila) {
      return {
        nombre: String(fila[0] || '').trim(),
        rol: String(fila[1] || '').trim().toLowerCase(),
        chatId: String(fila[2] || '').trim()
      };
    })
    .filter(function (c) { return c.nombre && c.chatId; });
}

/**
 * Arma la respuesta HTTP en formato JSON.
 */
function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ===================================================================
 * MANTENIMIENTO DE LA HOJA (ejecutar a mano desde el editor)
 * ===================================================================
 *
 * El sistema SOLO lee la columna DATA_JSON. Las columnas ID, ESTADO y
 * MAQUINA son un espejo para poder leer la hoja a simple vista; el
 * script las reescribe en cada guardado.
 *
 * Si alguien edita la hoja a mano (arrastrando el cuadrito de relleno,
 * pegando celdas u ordenando una sola columna), un DATA_JSON puede
 * quedar duplicado. Cuando eso pasa, validarReemplazo_ rechaza TODOS
 * los guardados y el formulario se bloquea. Estas funciones sirven
 * para detectarlo y repararlo:
 *
 *   1. auditarRegistros()              → dice qué está mal
 *   2. recuperarDesdeRespaldo(id)      → muestra el JSON perdido
 *   3. recuperarDesdeRespaldo(id, true)→ lo vuelve a escribir
 *   4. quitarFilaDuplicada(fila)       → borra una fila repetida
 *   5. repararColumnasEspejo()         → realinea ID/ESTADO/MAQUINA
 *   6. blindarHoja()                   → protección y nota de aviso
 */

/**
 * Revisa la pestaña Registros y reporta en el log:
 *   - ids repetidos dentro de DATA_JSON (esto es lo que bloquea el guardado)
 *   - filas donde la columna ID no coincide con el id del DATA_JSON
 *   - celdas DATA_JSON vacías o ilegibles
 * No modifica nada.
 */
function auditarRegistros() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const reporte = { filas: 0, maxId: 0, duplicados: [], desajustes: [], invalidas: [] };

  if (lastRow < 2) {
    Logger.log('La hoja ' + SHEET_NAME + ' está vacía.');
    return reporte;
  }

  const valores = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const vistos = {};
  reporte.filas = valores.length;

  valores.forEach(function (fila, i) {
    const numeroFila = i + 2;
    const columnaId = fila[0];
    const crudo = fila[3];

    if (crudo === '' || crudo == null) {
      if (columnaId !== '' && columnaId != null) {
        reporte.invalidas.push({ fila: numeroFila, motivo: 'DATA_JSON vacío', columnaId: columnaId });
      }
      return;
    }

    var rec = null;
    try {
      rec = JSON.parse(crudo);
    } catch (err) {
      reporte.invalidas.push({ fila: numeroFila, motivo: 'JSON ilegible', columnaId: columnaId });
      return;
    }
    if (!rec || rec.id == null) {
      reporte.invalidas.push({ fila: numeroFila, motivo: 'JSON sin id', columnaId: columnaId });
      return;
    }

    const id = String(rec.id);
    if (vistos[id]) {
      reporte.duplicados.push({ fila: numeroFila, id: rec.id, primeraFila: vistos[id] });
    } else {
      vistos[id] = numeroFila;
    }
    if (String(columnaId) !== id) {
      reporte.desajustes.push({ fila: numeroFila, columnaId: columnaId, idJson: rec.id });
    }
    if (Number(rec.id) > reporte.maxId) reporte.maxId = Number(rec.id);
  });

  Logger.log('Filas con datos: ' + reporte.filas + '  —  id más alto: ' + reporte.maxId);

  if (reporte.duplicados.length === 0) {
    Logger.log('Ids duplicados: ninguno.');
  } else {
    Logger.log('IDS DUPLICADOS (el guardado está bloqueado por esto):');
    reporte.duplicados.forEach(function (d) {
      Logger.log('  fila ' + d.fila + ' repite el id ' + d.id + ' que ya estaba en la fila ' + d.primeraFila);
    });
  }

  if (reporte.desajustes.length === 0) {
    Logger.log('Columna ID contra DATA_JSON: todo coincide.');
  } else {
    Logger.log('COLUMNA ID QUE NO COINCIDE CON EL DATA_JSON:');
    reporte.desajustes.forEach(function (d) {
      Logger.log('  fila ' + d.fila + ': la columna dice ' + d.columnaId + ' y el DATA_JSON dice ' + d.idJson);
    });
  }

  if (reporte.invalidas.length > 0) {
    Logger.log('FILAS CON DATA_JSON QUE NO SE PUEDE LEER:');
    reporte.invalidas.forEach(function (d) {
      Logger.log('  fila ' + d.fila + ': ' + d.motivo + ' (columna ID = ' + d.columnaId + ')');
    });
  }

  if (reporte.duplicados.length > 0) {
    Logger.log('');
    Logger.log('QUÉ SIGUE: el id que aparece en la columna ID de esa fila es el servicio que se perdió.');
    Logger.log('Ejecuta recuperarDesdeRespaldo(ESE_ID) para verlo y recuperarDesdeRespaldo(ESE_ID, true) para restaurarlo.');
  }

  return reporte;
}

/**
 * Busca en la pestaña Respaldos la última copia guardada de un servicio.
 * Con aplicar = true la vuelve a escribir en Registros.
 *
 *   recuperarDesdeRespaldo(75)        → solo muestra el JSON en el log
 *   recuperarDesdeRespaldo(75, true)  → lo restaura en la hoja
 */
function recuperarDesdeRespaldo(id, aplicar) {
  if (id == null || id === '') {
    throw new Error('Indica el id a recuperar. Ejemplo: recuperarDesdeRespaldo(75)');
  }
  const buscado = String(id);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPALDOS_SHEET);
  if (!sheet) {
    throw new Error('Todavía no existe la pestaña ' + RESPALDOS_SHEET + '.');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('La pestaña ' + RESPALDOS_SHEET + ' está vacía.');
  }

  const valores = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (var i = valores.length - 1; i >= 0; i--) {
    const crudo = valores[i][4];
    if (!crudo) continue;
    var rec = null;
    try {
      rec = JSON.parse(crudo);
    } catch (err) {
      continue;
    }
    if (!rec || String(rec.id) !== buscado) continue;

    Logger.log('Encontrado el id ' + buscado + ' en el respaldo del ' + valores[i][0] +
      ' (fila ' + (i + 2) + ' de ' + RESPALDOS_SHEET + ').');
    Logger.log(crudo);

    if (aplicar !== true) {
      Logger.log('');
      Logger.log('Revisa el JSON de arriba. Si es el servicio correcto, ejecuta:');
      Logger.log('  recuperarDesdeRespaldo(' + buscado + ', true)');
      return rec;
    }

    const filaDestino = escribirRegistroRecuperado_(rec, String(crudo));
    Logger.log('Restaurado en la fila ' + filaDestino + ' de ' + SHEET_NAME +
      '. Corre auditarRegistros() para confirmar que ya no hay duplicados.');
    return rec;
  }

  throw new Error('Ningún respaldo tiene el id ' + buscado +
    '. Revisa Archivo → Historial de versiones de la hoja, o vuelve a capturar el servicio.');
}

/**
 * Escribe un registro recuperado. Reutiliza la fila cuya columna ID ya
 * apunta a ese servicio (la que quedó con el DATA_JSON equivocado); si no
 * existe, lo agrega al final. Se niega a escribir si el id ya está vivo.
 */
function escribirRegistroRecuperado_(rec, crudo) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const buscado = String(rec.id);
  var destino = 0;

  if (lastRow >= 2) {
    const valores = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    for (var i = 0; i < valores.length; i++) {
      var actual = null;
      try {
        actual = JSON.parse(valores[i][3]);
      } catch (err) {
        actual = null;
      }
      if (actual && String(actual.id) === buscado) {
        throw new Error('La fila ' + (i + 2) + ' ya tiene el id ' + buscado +
          ' en su DATA_JSON. No hay nada que recuperar.');
      }
      if (!destino && String(valores[i][0]) === buscado) destino = i + 2;
    }
  }

  if (!destino) destino = Math.max(lastRow, 1) + 1;
  sheet.getRange(destino, 1, 1, 4).setValues([
    [rec.id, rec.estado || '', rec.maquina || '', crudo]
  ]);
  return destino;
}

/**
 * Borra una fila repetida de Registros. Solo la borra si su id sigue
 * existiendo en otra fila, para no perder información.
 * El número de fila lo da auditarRegistros().
 */
function quitarFilaDuplicada(fila) {
  const numeroFila = Number(fila);
  if (!numeroFila || numeroFila < 2) {
    throw new Error('Indica el número de fila que reportó auditarRegistros(). Ejemplo: quitarFilaDuplicada(9)');
  }
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (numeroFila > lastRow) {
    throw new Error('La fila ' + numeroFila + ' no existe en ' + SHEET_NAME + '.');
  }

  var objetivo = null;
  try {
    objetivo = JSON.parse(sheet.getRange(numeroFila, 4).getValue());
  } catch (err) {
    objetivo = null;
  }
  if (!objetivo || objetivo.id == null) {
    throw new Error('La fila ' + numeroFila + ' no tiene un DATA_JSON legible. Revísala a mano antes de borrarla.');
  }

  const valores = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  var apariciones = 0;
  valores.forEach(function (f) {
    var rec = null;
    try {
      rec = JSON.parse(f[0]);
    } catch (err) {
      rec = null;
    }
    if (rec && String(rec.id) === String(objetivo.id)) apariciones += 1;
  });

  if (apariciones < 2) {
    throw new Error('El id ' + objetivo.id + ' solo aparece una vez. Esta fila no es un duplicado y no se borra.');
  }

  sheet.deleteRow(numeroFila);
  Logger.log('Fila ' + numeroFila + ' borrada. El id ' + objetivo.id + ' queda ' + (apariciones - 1) + ' vez/veces.');
  return objetivo.id;
}

/**
 * Reescribe las columnas ID, ESTADO y MAQUINA a partir del DATA_JSON,
 * para que lo que se ve en la hoja coincida con lo que lee el sistema.
 */
function repararColumnasEspejo() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('La hoja ' + SHEET_NAME + ' está vacía, no hay nada que reparar.');
    return 0;
  }

  const rango = sheet.getRange(2, 1, lastRow - 1, 4);
  const valores = rango.getValues();
  var cambios = 0;

  valores.forEach(function (fila) {
    var rec = null;
    try {
      rec = JSON.parse(fila[3]);
    } catch (err) {
      rec = null;
    }
    if (!rec || rec.id == null) return;
    const estado = rec.estado || '';
    const maquina = rec.maquina || '';
    if (String(fila[0]) === String(rec.id) && String(fila[1]) === estado && String(fila[2]) === maquina) return;
    fila[0] = rec.id;
    fila[1] = estado;
    fila[2] = maquina;
    cambios += 1;
  });

  if (cambios > 0) rango.setValues(valores);
  Logger.log('Columnas espejo corregidas en ' + cambios + ' fila(s).');
  return cambios;
}

/**
 * Pone una protección de solo advertencia sobre Registros y una nota en
 * A1. No impide editar (el script necesita escribir), pero Sheets avisa
 * antes de que alguien modifique la hoja por accidente.
 */
function blindarHoja() {
  aplicarBlindaje_(getSheet_());
  PropertiesService.getScriptProperties().setProperty('HOJA_BLINDADA', '1');
  Logger.log('Listo: protección de advertencia y nota aplicadas en ' + SHEET_NAME + '.');
}

function aplicarBlindaje_(sheet) {
  const aviso =
    'Esta hoja la escribe el formulario automáticamente.\n' +
    'El sistema solo lee la columna DATA_JSON; ID, ESTADO y MAQUINA son una copia para leerla a simple vista.\n' +
    'NO edites, arrastres, pegues ni ordenes celdas aquí: un DATA_JSON repetido bloquea el guardado de todos.\n' +
    'Si algo se ve mal, ejecuta auditarRegistros() en Apps Script.';

  sheet.getRange('A1').setNote(aviso);

  const existentes = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < existentes.length; i++) {
    if (existentes[i].getDescription() === PROTECCION_DESC) return;
  }
  sheet.protect().setDescription(PROTECCION_DESC).setWarningOnly(true);
}

/**
 * Aplica el blindaje una sola vez, sin castigar cada petición con
 * llamadas extra a la hoja.
 */
function asegurarBlindaje_(sheet) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('HOJA_BLINDADA') === '1') return;
    aplicarBlindaje_(sheet);
    props.setProperty('HOJA_BLINDADA', '1');
  } catch (err) {
    // El blindaje es un extra: si la cuenta no puede protegerla, el resto sigue igual.
  }
}
