/**
 * Módulo de Operaciones V1.2
 * Funciones de escritura con protección anti-sobreescritura.
 */

/**
 * Bloquea un caso para el usuario actual. Optimizada para alta concurrencia.
 * Realiza la lectura de la base de datos antes de pedir el candado para 
 * minimizar el tiempo de bloqueo a menos de 1 segundo.
 * * @param {number} numeroFila Fila real en la hoja de cálculo.
 * @param {boolean} esRescate Si es true, permite sobreescribir al usuario anterior (usado en casos en Limbo).
 * @returns {Object} Un objeto indicando el resultado: { exito: boolean, mensaje?: string }
 */
function tomarCasoDinamico(numeroFila, esRescate) {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  
  // 1. TRABAJO PESADO AFUERA DEL CANDADO (Ahorra tiempo de bloqueo)
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colUSRIdx = headers.indexOf("USR") + 1;
  var colInicioAtnIdx = headers.indexOf("InicioAtencion") + 1;
  var colFolioIdx = headers.indexOf("FOLIO") + 1;
  
  var folio = "Desconocido";
  if (colFolioIdx > 0) {
    folio = sheet.getRange(numeroFila, colFolioIdx).getValue().toString().trim();
  }

  // LOG DE INTENTO
  console.log(`[INTENTO] Usuario: ${idUsuario} | Folio: ${folio} (Fila: ${numeroFila}) | Rescate: ${esRescate}`);

  var lock = LockService.getScriptLock();

  try {
    // 2. PACIENCIA AUMENTADA A 25 SEGUNDOS (25000 ms)
    lock.waitLock(25000); 
    
    // --- ZONA CRÍTICA: SOLO LECTURA RÁPIDA Y ESCRITURA ---
    var usuarioActualEnCelda = sheet.getRange(numeroFila, colUSRIdx).getValue().toString().trim();
    
    // Si la celda ya tiene a alguien y no es un rescate permitido, abortamos
    if (!esRescate && usuarioActualEnCelda !== "") {
      lock.releaseLock();
      console.warn(`[COLISIÓN] ${idUsuario} intentó tomar el Folio ${folio}, ya bloqueado por ${usuarioActualEnCelda}`);
      return { exito: false, mensaje: "⚠️ El caso ya fue tomado por: " + usuarioActualEnCelda };
    }

    // Escribimos al nuevo usuario
    sheet.getRange(numeroFila, colUSRIdx).setValue(idUsuario);
    if (colInicioAtnIdx > 0) {
      sheet.getRange(numeroFila, colInicioAtnIdx).setValue(new Date());
    }
    
    // Forzamos guardado y liberamos inmediatamente
    SpreadsheetApp.flush();
    lock.releaseLock(); 
    // -----------------------------------------------------
    
    console.log(`[ÉXITO] Folio ${folio} asignado a ${idUsuario}`);
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock();
    console.error(`[ERROR FATAL] Fallo al asignar Folio ${folio}. Detalle: ${e.toString()}`);
    return { exito: false, mensaje: "Error de servidor (Tráfico Alto): " + e.toString() }; 
  }
}

/**
 * Finaliza la atención de un caso registrando la clasificación, notas y la marca de tiempo de cierre.
 * Optimizada para alta concurrencia extrayendo la fase de lectura fuera del bloqueo (LockService).
 * * @param {number} numeroFila Fila real en la hoja de cálculo que se va a actualizar.
 * @param {string} clasificacion El tipo de fraude o resolución seleccionada por el usuario.
 * @param {string} indicaciones Los comentarios finales del especialista.
 * @returns {Object} Un objeto indicando el resultado: { exito: boolean, mensaje?: string }
 */
function finalizarCasoDinamico(numeroFila, clasificacion, indicaciones) {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  
  // 1. TRABAJO PESADO AFUERA DEL CANDADO (Lectura y preparación)
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange("1:1").getValues()[0];
  
  // Pre-calculamos las columnas para no perder tiempo dentro del bloqueo
  var colFolioIdx = headers.indexOf("FOLIO") + 1;
  var colClasifIdx = headers.indexOf("Clasificación") + 1;
  var colIndicIdx = headers.indexOf("Indicaciones") + 1;
  var colAtnIdx = headers.indexOf("Atención") + 1;
  
  var folio = "Desconocido";
  if (colFolioIdx > 0) {
    folio = sheet.getRange(numeroFila, colFolioIdx).getValue().toString().trim();
  }

  // LOG DE INTENTO
  console.log(`[FINALIZAR - INTENTO] ${idUsuario} va a cerrar el Folio: ${folio} con clasificación: ${clasificacion}`);

  var lock = LockService.getScriptLock();
  
  try {
    // 2. PACIENCIA AUMENTADA A 25 SEGUNDOS
    lock.waitLock(25000);
    
    // --- ZONA CRÍTICA: SOLO ESCRITURA RÁPIDA ---
    if (colClasifIdx > 0) sheet.getRange(numeroFila, colClasifIdx).setValue(clasificacion);
    if (colIndicIdx > 0) sheet.getRange(numeroFila, colIndicIdx).setValue(indicaciones);
    if (colAtnIdx > 0) sheet.getRange(numeroFila, colAtnIdx).setValue(new Date()); 
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    // ------------------------------------------
    
    console.log(`[FINALIZAR - ÉXITO] Folio ${folio} cerrado correctamente por ${idUsuario}`);
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    console.error(`[FINALIZAR - ERROR] Fallo al cerrar Folio ${folio} (Fila: ${numeroFila}) por ${idUsuario}. Detalle: ${e.toString()}`);
    return { exito: false, mensaje: "Error de servidor (Tráfico Alto): " + e.toString() }; 
  }
}

/**
 * Finaliza la atención de un caso registrando la clasificación, notas y la marca de tiempo de cierre.
 * @param {number} numeroFila Fila real en la hoja de cálculo que se va a actualizar.
 * @param {string} clasificacion El tipo de fraude o resolución seleccionada por el usuario.
 * @param {string} indicaciones Los comentarios finales del especialista.
 */
function finalizarCasoDinamico(numeroFila, clasificacion, indicaciones) {
  // Identificamos quién hace la acción para la auditoría
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  
  var lock = LockService.getScriptLock();
  var folio = "Desconocido"; // Declaramos arriba para que el catch siempre lo vea
  
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var headers = sheet.getRange("1:1").getValues()[0];
    
    // Ubicamos la columna del FOLIO para los logs
    var colFolioIdx = headers.indexOf("FOLIO") + 1;
    if (colFolioIdx > 0) {
      folio = sheet.getRange(numeroFila, colFolioIdx).getValue().toString().trim();
    }

    // 1. LOG DE INTENTO: Registramos el inicio de la operación de guardado
    console.log(`[FINALIZAR - INTENTO] ${idUsuario} va a cerrar el Folio: ${folio} con clasificación: ${clasificacion}`);

    // Escribimos los datos de cierre
    sheet.getRange(numeroFila, headers.indexOf("Clasificación") + 1).setValue(clasificacion);
    sheet.getRange(numeroFila, headers.indexOf("Indicaciones") + 1).setValue(indicaciones);
    sheet.getRange(numeroFila, headers.indexOf("Atención") + 1).setValue(new Date()); 
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    // 2. LOG DE ÉXITO: Confirmamos que se cerró bien
    console.log(`[FINALIZAR - ÉXITO] Folio ${folio} cerrado correctamente por ${idUsuario}`);
    
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    
    // 3. LOG DE ERROR CRÍTICO: Atrapamos el detalle técnico exacto
    console.error(`[FINALIZAR - ERROR] Fallo al cerrar Folio ${folio} (Fila: ${numeroFila}) por ${idUsuario}. Detalle: ${e.toString()}`);
    
    return { exito: false, mensaje: e.toString() }; 
  }
}
