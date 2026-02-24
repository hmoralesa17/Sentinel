/**
 * Módulo de Operaciones V1.2
 * Funciones de escritura con protección anti-sobreescritura.
 */

/**
 * Bloquea el caso para el usuario actual.
 * @param {number} numeroFila Fila real en la hoja.
 * @param {boolean} esRescate Si es true, permite pisar al usuario anterior (usado en casos Limbo).
 */
function tomarCasoDinamico(numeroFila, esRescate) {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  var lock = LockService.getScriptLock();
  
  try {
    // Esperamos turno en la fila (máximo 10 segundos para obtener el candado)
    lock.waitLock(10000);
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    var colUSRIdx = headers.indexOf("USR") + 1;
    var colInicioAtnIdx = headers.indexOf("InicioAtencion") + 1;

    // --- PROTECCIÓN DE CONCURRENCIA ---
    // Volvemos a leer la celda directamente para ver el valor REAL en este milisegundo
    var usuarioActualEnCelda = sheet.getRange(numeroFila, colUSRIdx).getValue().toString().trim();

    // Si NO es una acción de rescate y la celda ya tiene un usuario asignado...
    if (!esRescate && usuarioActualEnCelda !== "") {
      lock.releaseLock(); // Soltamos el candado
      return { 
        exito: false, 
        mensaje: "⚠️ El caso ya fue tomado hace unos instantes por: " + usuarioActualEnCelda 
      };
    }

    // Si llegamos aquí es porque: o la celda estaba vacía, o es un RESCATE permitido.
    sheet.getRange(numeroFila, colUSRIdx).setValue(idUsuario);
    if (colInicioAtnIdx > 0) {
      sheet.getRange(numeroFila, colInicioAtnIdx).setValue(new Date());
    }
    
    // Forzamos la escritura inmediata en la base de datos
    SpreadsheetApp.flush();
    
    // Liberamos para que el siguiente en la fila pueda entrar (y ver que ya ocupamos el lugar)
    lock.releaseLock();
    
    return { exito: true };
  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    return { exito: false, mensaje: "Error de servidor: " + e.toString() }; 
  }
}

function liberarCasoDinamico(numeroFila) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange("1:1").getValues()[0];
  
  var colUSRIdx = headers.indexOf("USR") + 1;
  var colInicioAtnIdx = headers.indexOf("InicioAtencion") + 1;

  if (colUSRIdx > 0) sheet.getRange(numeroFila, colUSRIdx).clearContent();
  if (colInicioAtnIdx > 0) sheet.getRange(numeroFila, colInicioAtnIdx).clearContent();
  
  return { exito: true };
}

function finalizarCasoDinamico(numeroFila, clasificacion, indicaciones) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var headers = sheet.getRange("1:1").getValues()[0];
    
    sheet.getRange(numeroFila, headers.indexOf("Clasificación") + 1).setValue(clasificacion);
    sheet.getRange(numeroFila, headers.indexOf("Indicaciones") + 1).setValue(indicaciones);
    sheet.getRange(numeroFila, headers.indexOf("Atención") + 1).setValue(new Date()); 
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    return { exito: true };
  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    return { exito: false, mensaje: e.toString() }; 
  }
}
