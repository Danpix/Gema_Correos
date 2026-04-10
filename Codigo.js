var API_KEY = API_KEY_ENV; // Reemplaza con tu API Key de Google Cloud

var MODELO = "gemini-2.5-flash";
var MODELO_FALLBACK = "gemini-2.0-flash-lite";
var SEARCH_QUERY = "in:inbox category:primary";
var SEARCH_LIMIT = 4;
var GEMINI_MAX_ATTEMPTS = 2;
var GEMINI_RETRY_DELAY_MS = 10000;
var MAX_CONTENIDO_CHARS = 4000;
var LABEL_NO_INCIDENCIA = "NO INCIDENCIA";
var TEAM_LABELS = [
  "PLATFORM",
  "CIB ONLINE",
  "OGB",
  "SYBASE / BBDD",
  "HPC GRID",
  "CAU",
  "EQUIPO DESCONOCIDO"
];

function obtenerFirma() {
  try {
    var aliases = Gmail.Users.Settings.SendAs.list("me").sendAs;
    for (var i = 0; i < aliases.length; i++) {
      if (aliases[i].isDefault) {
        return aliases[i].signature || "";
      }
    }
    return aliases[0].signature || "";
  } catch (error) {
    Logger.log("No se pudo obtener la firma desde Gmail avanzado: " + error.message);
    return "";
  }
}

function leerCorreos() {
  var threads = obtenerThreadsObjetivo();
  var miCorreo = obtenerCorreoActual();

  for (var i = 0; i < threads.length; i++) {
    procesarThread(threads[i], i + 1, miCorreo);
  }
}

function obtenerThreadsObjetivo() {
  return GmailApp.search(SEARCH_QUERY, 0, SEARCH_LIMIT);
}

function obtenerCorreoActual() {
  return Session.getActiveUser().getEmail().toLowerCase();
}

function procesarThread(thread, indiceHilo, miCorreo) {
  var asunto = "";
  var equipoAsignado = "";
  var threadId = thread.getId();
  var operationContext = crearContextoOperacion();

  try {
    asunto = thread.getFirstMessageSubject();
    var mensajes = thread.getMessages();
    var mensajeReferencia = obtenerMensajeReferencia(mensajes, miCorreo);
    var contenidoReferencia = mensajeReferencia.getPlainBody();

    if (debeOmitirsePorEtiquetaNoIncidencia(thread, contenidoReferencia, miCorreo)) {
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Etiqueta NO INCIDENCIA activa sin mención @, se omite");
      return;
    }

    if (threadTieneRespuestaPropia(mensajes, miCorreo)) {
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Ya respondido por el usuario, se omite");
      return;
    }

    if (threadTieneBorradorAbierto(threadId)) {
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Ya existe un borrador abierto, se omite");
      return;
    }

    var contenidoHilo = construirContenidoHilo(mensajes);
    var equipoFijo = obtenerEquipoFijoDesdeEtiquetas(thread);

    if (equipoFijo) {
      equipoAsignado = equipoFijo;
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Equipo fijo detectado por etiqueta: " + equipoAsignado);
    } else {
      equipoAsignado = clasificarEquipo(asunto, contenidoHilo, operationContext);
    }

    if (esNoIncidenciaPorCc(mensajeReferencia, miCorreo, contenidoReferencia)) {
      aplicarEtiquetasNoIncidencia(thread, equipoAsignado);
      saveEmailAnalysis(threadId, asunto, equipoAsignado, "no incidencia", operationContext.tokensUsed);
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Marcado como No incidencia (usuario en CC sin mención @)");
      return;
    }

    var estadoBorrador = revisarEstadoBorrador(threadId, asunto, equipoAsignado);
    var incidenciaAbierta = evaluarIncidenciaAbierta(equipoAsignado, threadId, contenidoHilo, operationContext);

    if (!incidenciaAbierta) {
      aplicarEtiquetaExistente(thread, "NO INCIDENCIA");
      aplicarEtiquetaExistente(thread, equipoAsignado);
      saveEmailAnalysis(threadId, asunto, equipoAsignado, "resuelta", operationContext.tokensUsed);
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Incidencia resuelta, no se crea borrador");
      return;
    }

    aplicarEtiquetaExistente(thread, equipoAsignado);

    if (estadoBorrador === "exists") {
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Ya existe borrador, no se modifica");
      return;
    }

    if (estadoBorrador === "deleted") {
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Borrador eliminado detectado, no se recrea en esta corrida");
      return;
    }

    var messageId = mensajeReferencia.getHeader("Message-ID");
    var destinatarios = construirDestinatarios(mensajeReferencia, miCorreo, indiceHilo);

    if (!destinatarios.paraFinal) {
      saveEmailAnalysis(threadId, asunto, equipoAsignado, "incorrecto", operationContext.tokensUsed);
      Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Sin destinatarios, se omite");
      return;
    }

    var cuerpo = construirCuerpoBorrador(equipoAsignado);
    var draftId = crearBorradorRespuesta(
      thread,
      miCorreo,
      destinatarios.paraFinal,
      destinatarios.ccFinal,
      asunto,
      messageId,
      cuerpo.texto,
      cuerpo.html
    );

    registrarBorradorThread(threadId, draftId);
    saveEmailAnalysis(threadId, asunto, equipoAsignado, "creado", operationContext.tokensUsed);
    Logger.log("Hilo #" + indiceHilo + " | Asunto: " + asunto + " | Equipo: " + equipoAsignado + " | Borrador creado");
  } catch (error) {
    saveEmailAnalysis(threadId, asunto, equipoAsignado, "incorrecto", operationContext.tokensUsed);
    Logger.log("ERROR en Hilo #" + indiceHilo + " | Asunto: " + asunto + " | " + error.message);
  }
}

function debeOmitirsePorEtiquetaNoIncidencia(thread, contenidoReferencia, miCorreo) {
  if (!threadTieneEtiqueta(thread, LABEL_NO_INCIDENCIA)) {
    return false;
  }

  return !contenidoTieneMencionArroba(contenidoReferencia, miCorreo);
}

function threadTieneEtiqueta(thread, labelName) {
  return labelExisteEnLista(thread.getLabels(), labelName);
}

function threadTieneRespuestaPropia(mensajes, miCorreo) {
  for (var i = 0; i < mensajes.length; i++) {
    if (esMensajePropio(mensajes[i], miCorreo)) {
      return true;
    }
  }

  return false;
}

function threadTieneBorradorAbierto(threadId) {
  if (!threadId) {
    return false;
  }

  var drafts = GmailApp.getDrafts();
  for (var i = 0; i < drafts.length; i++) {
    var draftMessage = drafts[i].getMessage();
    if (!draftMessage) {
      continue;
    }

    var draftThread = draftMessage.getThread();
    if (draftThread && String(draftThread.getId()) === String(threadId)) {
      return true;
    }
  }

  return false;
}

function crearContextoOperacion() {
  return {
    tokensUsed: 0,
    modelError: false
  };
}

function obtenerMensajeReferencia(mensajes, miCorreo) {
  for (var i = mensajes.length - 1; i >= 0; i--) {
    if (!esMensajePropio(mensajes[i], miCorreo)) {
      return mensajes[i];
    }
  }

  return mensajes[mensajes.length - 1];
}

function esMensajePropio(mensaje, miCorreo) {
  var fromEmail = extraerEmails(mensaje.getFrom())[0] || "";
  return fromEmail.toLowerCase() === String(miCorreo || "").toLowerCase();
}

function construirContenidoHilo(mensajes) {
  var contenido = "";
  for (var j = 0; j < mensajes.length; j++) {
    contenido += mensajes[j].getPlainBody() + "\n---\n";
  }
  if (contenido.length > MAX_CONTENIDO_CHARS) {
    contenido = contenido.substring(contenido.length - MAX_CONTENIDO_CHARS);
  }
  return contenido;
}

function evaluarIncidenciaAbierta(equipoAsignado, threadId, contenidoHilo, operationContext) {
  var resultadoIncidencia = extraerIncidencia(equipoAsignado, threadId, contenidoHilo, operationContext);
  return esIncidenciaAbierta(resultadoIncidencia);
}

function aplicarEtiquetasNoIncidencia(thread, equipoAsignado) {
  aplicarEtiquetaExistente(thread, "NO INCIDENCIA");
  aplicarEtiquetaExistente(thread, equipoAsignado);
}

function obtenerEquipoFijoDesdeEtiquetas(thread) {
  var labels = thread.getLabels();

  for (var i = 0; i < TEAM_LABELS.length; i++) {
    if (labelExisteEnLista(labels, TEAM_LABELS[i])) {
      return TEAM_LABELS[i];
    }
  }

  return "";
}

function labelExisteEnLista(labels, labelName) {
  var expected = String(labelName || "").trim().toLowerCase();

  for (var i = 0; i < labels.length; i++) {
    if (String(labels[i].getName() || "").trim().toLowerCase() === expected) {
      return true;
    }
  }

  return false;
}

function construirCuerpoBorrador(equipoAsignado) {
  var firma = obtenerFirma();
  var texto = "Saludos, este tipo de incidencia no la resolvemos nosotros, sin embargo agregamos al equipo encargado: " + equipoAsignado;
  var html = "<p>" + texto + "</p><br>" + firma;
  return {
    texto: texto,
    html: html
  };
}

function construirDestinatarios(ultimoMensaje, miCorreo, indiceHilo) {
  var fromEmail = extraerEmails(ultimoMensaje.getFrom())[0] || "";
  var toEmails = extraerEmails(ultimoMensaje.getTo());
  var ccEmails = extraerEmails(ultimoMensaje.getCc());
  var yoEnvie = fromEmail.toLowerCase() === miCorreo;

  Logger.log("DEBUG Hilo #" + indiceHilo + " | From: " + fromEmail + " | To: " + toEmails.join(", ") + " | CC: " + ccEmails.join(", ") + " | yoEnvie: " + yoEnvie);

  var destinatariosPara = yoEnvie ? toEmails : [fromEmail].concat(toEmails).filter(function(email) {
    return email.toLowerCase() !== miCorreo;
  });

  var paraDeduplicado = quitarDuplicados(destinatariosPara);
  var vistos = {};
  for (var i = 0; i < paraDeduplicado.length; i++) {
    vistos[paraDeduplicado[i].toLowerCase()] = true;
  }

  var ccFinal = ccEmails.filter(function(email) {
    var key = email.toLowerCase();
    return key !== miCorreo && !vistos[key];
  });

  var paraFinal = paraDeduplicado.join(", ");
  var ccTexto = ccFinal.join(", ");
  Logger.log("ENVIAR Para: [" + paraFinal + "] CC: [" + ccTexto + "]");

  return {
    paraFinal: paraFinal,
    ccFinal: ccTexto
  };
}

function esNoIncidenciaPorCc(ultimoMensaje, miCorreo, contenidoHilo) {
  var toEmails = extraerEmails(ultimoMensaje.getTo());
  var ccEmails = extraerEmails(ultimoMensaje.getCc());
  var miCorreoLower = String(miCorreo || "").toLowerCase();

  var estoyEnCc = contieneEmail(ccEmails, miCorreoLower);
  var estoyEnTo = contieneEmail(toEmails, miCorreoLower);

  if (!estoyEnCc || estoyEnTo) {
    return false;
  }

  return !contenidoTieneMencionArroba(contenidoHilo, miCorreoLower);
}

function contieneEmail(listaEmails, emailBuscado) {
  for (var i = 0; i < listaEmails.length; i++) {
    if (String(listaEmails[i] || "").toLowerCase() === emailBuscado) {
      return true;
    }
  }

  return false;
}

function contenidoTieneMencionArroba(contenidoHilo, miCorreo) {
  var texto = String(contenidoHilo || "").toLowerCase();
  var tokens = obtenerTokensMencion(miCorreo);

  for (var i = 0; i < tokens.length; i++) {
    if (texto.indexOf("@" + tokens[i]) !== -1) {
      return true;
    }
  }

  return false;
}

function obtenerTokensMencion(miCorreo) {
  var localPart = String(miCorreo || "").split("@")[0].toLowerCase();
  if (!localPart) {
    return [];
  }

  var tokens = [localPart];
  var porPunto = localPart.split(".");
  for (var i = 0; i < porPunto.length; i++) {
    if (porPunto[i] && porPunto[i].length >= 3) {
      tokens.push(porPunto[i]);
    }
  }

  var vistos = {};
  var salida = [];
  for (var j = 0; j < tokens.length; j++) {
    var t = tokens[j];
    if (!vistos[t]) {
      vistos[t] = true;
      salida.push(t);
    }
  }

  return salida;
}

function extraerEmails(cadena) {
  if (!cadena) {
    return [];
  }

  var partes = cadena.split(",");
  var emails = [];

  for (var i = 0; i < partes.length; i++) {
    var parte = partes[i].trim();
    if (!parte) {
      continue;
    }

    var match = parte.match(/<([^>]+)>/);
    var email = match ? match[1].trim() : parte.trim();
    if (email.indexOf("@") !== -1) {
      emails.push(email);
    }
  }

  return emails;
}

function quitarDuplicados(listaEmails) {
  var vistos = {};
  var salida = [];

  for (var i = 0; i < listaEmails.length; i++) {
    var email = String(listaEmails[i] || "").trim();
    var key = email.toLowerCase();
    if (!email || vistos[key]) {
      continue;
    }

    vistos[key] = true;
    salida.push(email);
  }

  return salida;
}

function crearBorradorRespuesta(thread, miCorreo, paraFinal, cc, asunto, messageId, textoBorrador, htmlBorrador) {
  try {
    var rawMessage = "MIME-Version: 1.0\r\n" +
      "From: " + miCorreo + "\r\n" +
      "To: " + paraFinal + "\r\n" +
      (cc ? "Cc: " + cc + "\r\n" : "") +
      "Subject: =?UTF-8?B?" + Utilities.base64Encode(Utilities.newBlob("Re: " + asunto).getBytes()) + "?=\r\n" +
      "In-Reply-To: " + messageId + "\r\n" +
      "References: " + messageId + "\r\n" +
      "Content-Type: text/html; charset=UTF-8\r\n" +
      "Content-Transfer-Encoding: base64\r\n" +
      "\r\n" +
      Utilities.base64Encode(Utilities.newBlob(htmlBorrador).getBytes());

    var encodedMessage = Utilities.base64EncodeWebSafe(Utilities.newBlob(rawMessage).getBytes());

    var draft = Gmail.Users.Drafts.create(
      { message: { raw: encodedMessage, threadId: thread.getId() } },
      "me"
    );
    return draft && draft.id ? String(draft.id) : "";
  } catch (error) {
    Logger.log("No se pudo crear el borrador con Gmail avanzado: " + error.message);
  }

  var draftOptions = {
    htmlBody: htmlBorrador
  };

  if (cc) {
    draftOptions.cc = cc;
  }

  var fallbackDraft = thread.createDraftReply(textoBorrador, draftOptions);
  if (fallbackDraft && typeof fallbackDraft.getId === "function") {
    return String(fallbackDraft.getId());
  }
  return "";
}

function registrarBorradorThread(threadId, draftId) {
  if (!threadId || !draftId) {
    return;
  }

  var key = "draftThread:" + String(threadId);
  PropertiesService.getScriptProperties().setProperty(key, String(draftId));
}

function revisarEstadoBorrador(threadId, asunto, equipoAsignado) {
  if (!threadId) {
    return "none";
  }

  var key = "draftThread:" + String(threadId);
  var properties = PropertiesService.getScriptProperties();
  var draftId = properties.getProperty(key);

  if (!draftId) {
    return "none";
  }

  if (existeBorrador(draftId)) {
    return "exists";
  }

  saveEmailAnalysis(threadId, asunto, equipoAsignado, "borrado", 0);
  properties.deleteProperty(key);
  Logger.log("Hilo con borrador eliminado detectado | Asunto: " + asunto + " | DraftId: " + draftId);
  return "deleted";
}

function existeBorrador(draftId) {
  try {
    Gmail.Users.Drafts.get("me", draftId);
    return true;
  } catch (error) {
    // Si no hay Gmail avanzado, se usa el fallback con GmailApp.
  }

  var drafts = GmailApp.getDrafts();
  for (var i = 0; i < drafts.length; i++) {
    if (String(drafts[i].getId()) === String(draftId)) {
      return true;
    }
  }

  return false;
}


function clasificarEquipo(asunto, contenido, operationContext) {
  var prompt = "Instrucciones: Actúa como un motor de enrutamiento de tickets de soporte técnico para Murex 3. Analiza el correo proporcionado y responde ÚNICAMENTE con el nombre del equipo responsable de la lista siguiente.\n" +
    "Si no puedes determinarlo, responde \"EQUIPO DESCONOCIDO\".\n" +
    "Prioriza el contenido técnico sobre los nombres de departamentos en las firmas (ej. si el hilo menciona \"CIB\" en la firma pero el problema es de base de datos, asigna SYBASE / BBDD).\n\n" +
    "PLATFORM: Si el problema implica validar discrepancias en cálculos financieros (Yield, Precios, Accruals), lógica de negocio en Murex, o validación de datos funcionales tras un cambio de configuración (STP/Templates).\n" +
    "CIB ONLINE: Si el error es estrictamente técnico de conectividad, fallos en el enrutador de mensajes (ESB/Argos), o errores de mapeo XML/MxML que impiden que el mensaje llegue a su destino.\n" +
    "OGB: Relanzamiento de jobs, falta de espacio en batch o ejecución de aperiodicos SQL.\n" +
    "SYBASE / BBDD: Bloqueos de tablas, errores de logs de transacciones o problemas con la base de datos temporal.\n" +
    "HPC GRID: Timeouts de conexión al cluster de valoración o hilos de servicios colgados.\n" +
    "CAU: Problemas de acceso a la máquina (Putty/SSH), desbloqueo de usuarios UNIX, configuración de permisos en Murex, cuentas bloqueadas, reseteo o cambio de contraseña, o cualquier problema de autenticación/login.\n\n" +
    "Correo a analizar:\n" +
    "Asunto: " + asunto + "\n" + contenido + "\n\n" +
    "Solo responde con el nombre del equipo";

  var respuesta = enviarAGemini(prompt, operationContext);
  return normalizarEquipoClasificado(respuesta);
}

function normalizarEquipoClasificado(respuestaModelo) {
  var texto = String(respuestaModelo || "").trim();
  var upper = texto.toUpperCase();

  var equiposValidos = [
    "PLATFORM",
    "CIB ONLINE",
    "OGB",
    "SYBASE / BBDD",
    "HPC GRID",
    "CAU",
    "EQUIPO DESCONOCIDO"
  ];

  for (var i = 0; i < equiposValidos.length; i++) {
    if (upper.indexOf(equiposValidos[i]) !== -1) {
      return equiposValidos[i];
    }
  }

  return "EQUIPO DESCONOCIDO";
}

function extraerIncidencia(equipoAsignado, threadId, contenido, operationContext) {
  var prompt = "Actúa como un analista de soporte técnico. Lee el siguiente hilo de correos (puede estar en español o inglés) y determina si la incidencia sigue ABIERTA o ya fue RESUELTA/CERRADA.\n\n" +
    "REGLAS PARA DETERMINAR EL ESTADO:\n\n" +
    "La incidencia está CERRADA (false) si:\n" +
    "- El equipo asignado es \"MX3 ANS\"\n" +
    "- El último mensaje del hilo indica que el problema fue resuelto, actualizado, completado o cerrado\n" +
    "- Frases como: \"cerramos la incidencia\", \"no existe incidencia\", \"queda resuelta\", \"updated\", \"done\", \"completed\", \"fixed\", \"thank you\" seguido de confirmación de acción realizada, etc.\n" +
    "- La intención final del hilo es de cierre, agradecimiento por resolución o confirmación de que se realizó la acción solicitada\n\n" +
    "La incidencia está ABIERTA (true) si:\n" +
    "- El último mensaje reporta un problema nuevo o pendiente\n" +
    "- Se solicita una acción que aún no ha sido confirmada\n" +
    "- No hay señal de resolución en los mensajes recientes\n\n" +
    "IMPORTANTE: Enfócate en la INTENCIÓN FINAL del hilo, especialmente el último mensaje. Analiza en el idioma original del correo.\n\n" +
    "DATOS:\n" +
    "Equipo asignado: " + equipoAsignado + "\n" +
    "Contenido del correo:\n" + contenido + "\n\n" +
    "Responde SOLO con el JSON (sin bloques de código Markdown):\n" +
    "{\n\"equipo_final\": \"" + equipoAsignado + "\",\n\"incidencia\": \"true o false\"\n}";

  return enviarAGemini(prompt, operationContext);
}

function esIncidenciaAbierta(respuestaIncidencia) {
  if (!respuestaIncidencia) {
    return true;
  }

  var texto = String(respuestaIncidencia).trim();

  try {
    var json = JSON.parse(texto);
    var valor = String(json.incidencia || "").toLowerCase().trim();
    if (valor === "false") {
      return false;
    }
    if (valor === "true") {
      return true;
    }
  } catch (error) {
    Logger.log("No se pudo parsear JSON de incidencia, se toma como abierta: " + error.message);
  }

  var normalizado = texto.toLowerCase();
  if (normalizado.indexOf('"incidencia":"false"') !== -1 || normalizado.indexOf('"incidencia": "false"') !== -1) {
    return false;
  }

  return true;
}

function aplicarEtiquetaExistente(thread, equipoAsignado) {
  if (!equipoAsignado) {
    return;
  }

  var nombreEtiqueta = obtenerEtiquetaPorEquipo(equipoAsignado);
  if (!nombreEtiqueta) {
    Logger.log("No hay etiqueta configurada para el equipo: " + equipoAsignado);
    return;
  }

  var etiqueta = GmailApp.getUserLabelByName(nombreEtiqueta);
  if (!etiqueta) {
    Logger.log("Etiqueta no encontrada, se omite: " + nombreEtiqueta);
    return;
  }

  etiqueta.addToThread(thread);
}

function obtenerEtiquetaPorEquipo(equipoAsignado) {
  var equipo = String(equipoAsignado || "").toUpperCase().trim();

  switch (equipo) {
    case "NO INCIDENCIA":
      return "NO INCIDENCIA";
    case "PLATFORM":
      return "PLATFORM"; 
    case "CIB ONLINE":
      return "CIB ONLINE";
    case "OGB":
      return "OGB";
    case "SYBASE / BBDD":
      return "SYBASE / BBDD";
    case "HPC GRID":
      return "HPC GRID";
    case "CAU":
      return "CAU";
    case "EQUIPO DESCONOCIDO":
      return "EQUIPO DESCONOCIDO";
    default:
      return "";
  }
}

function enviarAGemini(prompt, operationContext) {
  var modelos = [MODELO, MODELO_FALLBACK];

  for (var m = 0; m < modelos.length; m++) {
    var modeloActual = modelos[m];
    var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modeloActual + ":generateContent?key=" + API_KEY;

    var payload = {
      contents: [
        {
          parts: [
            {
              text: prompt
            }
          ]
        }
      ]
    };

    var opciones = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    for (var attempt = 1; attempt <= GEMINI_MAX_ATTEMPTS; attempt++) {
      try {
        var respuesta = UrlFetchApp.fetch(url, opciones);
        var rawText = respuesta.getContentText();
        var httpCode = respuesta.getResponseCode();
        var json;

        try {
          json = JSON.parse(rawText);
        } catch (parseError) {
          json = null;
        }

        acumularTokensOperacion(operationContext, json);

        if (json && json.candidates && json.candidates[0]) {
          if (m > 0) {
            Logger.log("Respuesta obtenida con modelo fallback: " + modeloActual);
          }
          return json.candidates[0].content.parts[0].text;
        }

        var unavailableError = json && json.error && (json.error.code === 503 || json.error.status === "UNAVAILABLE");
        var shouldRetry = (httpCode === 503 || unavailableError) && attempt < GEMINI_MAX_ATTEMPTS;

        if (shouldRetry) {
          Logger.log("Gemini saturado (503) con " + modeloActual + ". Reintentando en " + (GEMINI_RETRY_DELAY_MS / 1000).toFixed(1) + " segundos...");
          Utilities.sleep(GEMINI_RETRY_DELAY_MS);
          continue;
        }

        // Si 503 y ya no hay reintentos, saltar al siguiente modelo
        if (httpCode === 503 || unavailableError) {
          Logger.log("Modelo " + modeloActual + " no disponible (503). Probando siguiente modelo...");
          break;
        }

        // Cualquier otro error: loguear y probar siguiente modelo
        var errorMsg = json && json.error ? (json.error.message || String(json.error.code)) : rawText.substring(0, 200);
        Logger.log("Modelo " + modeloActual + " error (" + httpCode + "): " + errorMsg + ". Probando siguiente modelo...");
        break;
      } catch (error) {
        if (attempt < GEMINI_MAX_ATTEMPTS) {
          Logger.log("Error llamando a Gemini (" + modeloActual + "). Reintentando en " + (GEMINI_RETRY_DELAY_MS / 1000).toFixed(1) + " segundos... " + error.message);
          Utilities.sleep(GEMINI_RETRY_DELAY_MS);
          continue;
        }

        Logger.log("Modelo " + modeloActual + " falló. Probando siguiente modelo... " + error.message);
        break;
      }
    }
  }

  if (operationContext) {
    operationContext.modelError = true;
  }

  return "Error: no fue posible obtener respuesta de ningún modelo";
}

function calcularEsperaGemini(attempt) {
  return GEMINI_RETRY_DELAY_MS;
}

function acumularTokensOperacion(operationContext, responseJson) {
  if (!operationContext || !responseJson || !responseJson.usageMetadata) {
    return;
  }

  var usage = responseJson.usageMetadata;
  var tokenCount = Number(usage.totalTokenCount || 0);
  if (!tokenCount && (usage.promptTokenCount || usage.candidatesTokenCount)) {
    tokenCount = Number(usage.promptTokenCount || 0) + Number(usage.candidatesTokenCount || 0);
  }

  operationContext.tokensUsed += tokenCount;
}

function main() {
  leerCorreos();
}

function limpiarPropiedadesBorradores() {
  var properties = PropertiesService.getScriptProperties();
  var todas = properties.getProperties();
  var eliminadas = 0;
  for (var key in todas) {
    if (key.indexOf("draftThread:") === 0) {
      properties.deleteProperty(key);
      eliminadas++;
    }
  }
  Logger.log("Propiedades de borradores eliminadas: " + eliminadas);
}
