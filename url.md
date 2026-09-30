https://script.google.com/macros/s/AKfycbzqFLePhr_RQo1aZeHO3BLJZDx6RCcmGekwBZSmhpPW0K-7FGX293qtUt1pKKDhGQjV/exec

Avisos Telegram (después de pegar el nuevo apps_script_codigo.gs):

1. En Apps Script: Propiedades del script
   TELEGRAM_BOT_TOKEN        = token de @BotFather
   TELEGRAM_CHAT_ID          = id del grupo de mantenimiento (número negativo)
   TELEGRAM_CODIGOS_CHAT_ID  = id del grupo PRIVADO de códigos (otro número negativo)

2. Ejecutar mostrarChatsTelegram si aún no tienes algún chat_id (escribe antes en ese grupo).
3. Ejecutar probarTelegram — debe llegar al grupo de mantenimiento.
   Ejecutar probarTelegramCodigos — debe llegar SOLO al grupo interno de códigos.
4. Implementar → Administrar implementaciones → lápiz → Versión: Nueva → Implementar
   (así se conserva esta misma URL /exec).
5. Probar en el formulario: generar un reporte
   - grupo de mantenimiento: "Nuevo reporte" (sin código)
   - grupo de códigos: servicio, máquina y el código de 4 dígitos
   Asignar un técnico → solo el grupo de mantenimiento recibe "Servicio asignado".

Alta de técnicos/supervisores en Telegram: invitarlos al grupo de mantenimiento.
Al grupo de códigos: SOLO operaciones y gerentes. No técnicos ni piso.
La pestaña Contactos de la Hoja queda lista para avisos privados (aún no se usan).


Si el formulario dice "IDs duplicados" o un registro se ve raro en la hoja:

El sistema solo lee la columna DATA_JSON. Las columnas ID, ESTADO y MAQUINA son
una copia para leer la hoja a simple vista. Nadie debe editar la hoja a mano:
si un DATA_JSON queda repetido, se bloquea el guardado de todos.

En Apps Script (selecciona la función arriba y pulsa Ejecutar, luego Ver → Registros):

1. auditarRegistros()
   Dice qué filas tienen id repetido, cuáles no coinciden con su DATA_JSON
   y cuáles no se pueden leer. No modifica nada.

2. recuperarDesdeRespaldo(75)
   Muestra la última copia guardada de ese servicio en la pestaña Respaldos.
   Revisa que sea el correcto.

3. recuperarDesdeRespaldo(75, true)
   Lo vuelve a escribir en Registros, en la fila cuya columna ID ya decía 75.

4. quitarFilaDuplicada(9)
   Borra una fila repetida (el número de fila lo da auditarRegistros).
   Solo borra si ese id sigue existiendo en otra fila.

5. repararColumnasEspejo()
   Realinea ID, ESTADO y MAQUINA con el DATA_JSON.

6. blindarHoja()
   Pone la nota de aviso en A1 y la protección de advertencia. Se aplica sola
   la primera vez, pero puedes volver a ejecutarla si alguien la quitó.

Al final: auditarRegistros() debe salir sin duplicados, y en el formulario
pulsar "Actualizar" y comprobar que un cambio se guarda sin rechazo.
