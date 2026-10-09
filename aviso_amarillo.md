El aviso amarillo es la franja de arriba, siempre empieza con ⚠. Estos son los textos que puede mostrar.

Al abrir, si no llegan los servicios

"No se pudo cargar desde Google Sheets. El guardado está bloqueado para proteger los registros. Detalle: …" Es el de anoche si la lista se quedó vacía. En el detalle puede decir "Failed to fetch" (la red no llegó a Google) o "El servidor no devolvió una lista válida de registros." (Google contestó, pero no con la lista).
"Este formulario aún no está conectado a Google Sheets…" Solo saldría si la página se publicara sin la dirección del servidor.
Al intentar guardar y la página no tiene la lista

"Guardado bloqueado para proteger los datos: primero actualiza y confirma que los registros cargan desde Google Sheets."
"No se puede guardar: falta configurar la conexión con Google Sheets."
Al guardar y se corta la red

"No se guardó ningún cambio. Se bloquearon las acciones para proteger los registros. Actualiza la página. Detalle: …"
Cuando el servidor rechaza el cambio. El aviso muestra el texto de Google tal cual:

"Otro usuario actualizó el servicio #…. Se recargó la lista; vuelve a intentar." Alguien más modificó ese servicio al mismo tiempo.
"La máquina ya tiene el servicio #… en estado …. Debe liberarse antes de crear otro."
"El técnico … ya está en el servicio #… (…). Debe quedar libre antes de asignarlo a otro."
"Actualiza la página: esta versión del formulario ya no es compatible." La página y el script de Google no son la misma versión.
"La hoja de Google tiene un registro repetido y por eso no se puede guardar nada. Avisa al administrador para que ejecute la auditoría de la hoja (auditarRegistros). Detalle: …" Hay un servicio duplicado en la hoja y se bloquea todo.
Al liberar una máquina

"Para liberar, escribe y selecciona un nombre válido de la lista."
"Ese servicio ya no está pendiente de liberación. Actualiza la página."