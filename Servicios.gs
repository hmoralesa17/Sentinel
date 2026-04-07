/**
 * Módulo de Operaciones V2.0
 * Funciones de escritura, lectura protegida y Motor de Búsqueda
 * (Totalmente optimizado con Variables Globales)
 */

// =======================================================
// ⚙️ CONFIGURACIÓN INTERNA DEL MÓDULO (Solo para Devs)
// =======================================================
var CFG_SERVICIOS = {
  TIEMPO_ESPERA_CANDADO: 25000, 
  
  // Diccionario de Mensajes de Error
  MSG_COLISION: "⚠️ El caso acaba de ser tomado por: ",
  MSG_TRAFICO_ALTO: "Tráfico alto en el servidor. Cierra esta alerta e intenta guardar o atender de nuevo.",
  MSG_ERROR_COLUMNA: "Error interno: Faltan columnas clave (USR, FOLIO) en el Diccionario.",
  MSG_ERROR_LIBERAR: "No se pudo liberar el folio por saturación. Por favor, recarga tu página e intenta de nuevo.",
  MSG_ERROR_BUSQUEDA: "Error al realizar la búsqueda. Por favor, revisa tus filtros e intenta de nuevo."
};

// =======================================================
// 🚀 FUNCIONES OPERATIVAS (Tomar, Liberar, Finalizar)
// =======================================================

function tomarCasoDinamico(numeroFila, esRescate) {
  var idUsuario = DICCIONARIO_USUARIOS[GLOBAL_EMAIL] || GLOBAL_EMAIL.split('@')[0];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  // Usamos nuestras variables globales (Sumamos 1 porque getRange es Base 1)
  var folio = IDX_FOLIO > -1 ? sheet.getRange(numeroFila, IDX_FOLIO + 1).getValue().toString().trim() : "S/N";

  console.log(`[TOMA CASO] Usr: ${idUsuario} | Folio: ${folio} | Fila: ${numeroFila} | Rescate: ${esRescate}`);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(CFG_SERVICIOS.TIEMPO_ESPERA_CANDADO); 
    
    if (IDX_USR === -1) throw new Error(CFG_SERVICIOS.MSG_ERROR_COLUMNA);

    var celdaUSR = sheet.getRange(numeroFila, IDX_USR + 1);
    var usuarioActualEnCelda = celdaUSR.getValue().toString().trim();
    
    if (!esRescate && usuarioActualEnCelda !== "") {
      lock.releaseLock();
      console.warn(`[COLISIÓN] ${idUsuario} rebotado. Folio ${folio} ocupado por ${usuarioActualEnCelda}`);
      return { exito: false, mensaje: CFG_SERVICIOS.MSG_COLISION + usuarioActualEnCelda };
    }

    celdaUSR.setValue(idUsuario);
    if (IDX_INICIO_ATN > -1) sheet.getRange(numeroFila, IDX_INICIO_ATN + 1).setValue(new Date());
    
    SpreadsheetApp.flush();
    lock.releaseLock(); 
    
    console.log(`[ÉXITO] Folio ${folio} bloqueado por ${idUsuario}`);
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    console.error(`[ERROR CRÍTICO] Fallo en tomarCasoDinamico. Detalle: ${e.toString()}`);
    return { exito: false, mensaje: CFG_SERVICIOS.MSG_TRAFICO_ALTO }; 
  }
}

function liberarCasoDinamico(numeroFila) {
  var idUsuario = DICCIONARIO_USUARIOS[GLOBAL_EMAIL] || GLOBAL_EMAIL.split('@')[0];
  console.log(`[LIBERAR - INTENTO] ${idUsuario} está cancelando la atención de la fila: ${numeroFila}`);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(CFG_SERVICIOS.TIEMPO_ESPERA_CANDADO); 
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (IDX_USR > -1) sheet.getRange(numeroFila, IDX_USR + 1).clearContent();
    if (IDX_INICIO_ATN > -1) sheet.getRange(numeroFila, IDX_INICIO_ATN + 1).clearContent();
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    console.log(`[LIBERAR - ÉXITO] Caso en la fila ${numeroFila} liberado correctamente.`);
    return { exito: true };
    
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    console.error(`[LIBERAR - ERROR] Fallo al liberar la fila ${numeroFila}: ${e.toString()}`);
    return { exito: false, mensaje: CFG_SERVICIOS.MSG_ERROR_LIBERAR };
  }
}

function finalizarCasoDinamico(numeroFila, clasificacion, indicaciones) {
  var idUsuario = DICCIONARIO_USUARIOS[GLOBAL_EMAIL] || GLOBAL_EMAIL.split('@')[0];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  var folio = IDX_FOLIO > -1 ? sheet.getRange(numeroFila, IDX_FOLIO + 1).getValue().toString().trim() : "Desconocido";

  console.log(`[FINALIZAR - INTENTO] ${idUsuario} va a cerrar el Folio: ${folio}`);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(CFG_SERVICIOS.TIEMPO_ESPERA_CANDADO);
    
    if (IDX_CLASIF > -1) sheet.getRange(numeroFila, IDX_CLASIF + 1).setValue(clasificacion);
    if (IDX_INDICACIONES > -1) sheet.getRange(numeroFila, IDX_INDICACIONES + 1).setValue(indicaciones);
    if (IDX_ATENCION > -1) sheet.getRange(numeroFila, IDX_ATENCION + 1).setValue(new Date()); 
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    console.log(`[FINALIZAR - ÉXITO] Folio ${folio} cerrado por ${idUsuario}`);
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    console.error(`[FINALIZAR - ERROR] Fallo al cerrar Folio ${folio} (Fila: ${numeroFila}). Detalle: ${e.toString()}`);
    return { exito: false, mensaje: CFG_SERVICIOS.MSG_TRAFICO_ALTO }; 
  }
}


// =======================================================
// 🔎 MOTOR DE BÚSQUEDA AVANZADA (Explorador)
// =======================================================
function buscarFoliosExplorador(filtros) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // 🛡️ EL ESCUDO DEL ROBOT: Usamos la fila global directamente
    var ultimaFilaReal = IDX_FILA_ULTIMOFOLIO;
    
    // Fallback de seguridad por si falla la lectura del Bot
    if (ultimaFilaReal < 2) {
      ultimaFilaReal = sheet.getLastRow(); 
    }
    
    if (ultimaFilaReal < 2) return { exito: true, data: [] };
    
    // Descargamos SOLO el rango exacto
    var data = sheet.getRange(2, 1, ultimaFilaReal - 1, TODAS_LAS_COLUMNAS.length).getValues();
    var resultados = [];

    // --- SANITIZACIÓN PREVIA DE FILTROS ---
    var fFolio = filtros.folio ? filtros.folio.toString().toUpperCase().replace(/\s+/g, '') : "";
    var fCurp = filtros.curp ? filtros.curp.toString().toUpperCase().replace(/\s+/g, '') : "";
    var fCuenta = filtros.cuenta ? filtros.cuenta.toString().toUpperCase().replace(/\s+/g, '') : "";
    var fUsr = filtros.usr ? filtros.usr.toString().toUpperCase().replace(/\s+/g, '') : "";
    var fPalabra = filtros.palabraClave ? filtros.palabraClave.toString().toLowerCase() : "";

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // 1. REGLA DE FOLIO Y USR (Sanitizados y Parciales)
      var rFolio = (IDX_FOLIO > -1 && row[IDX_FOLIO]) ? row[IDX_FOLIO].toString().toUpperCase().replace(/\s+/g, '') : "";
      var rUsr = (IDX_USR > -1 && row[IDX_USR]) ? row[IDX_USR].toString().toUpperCase().replace(/\s+/g, '') : "";
      
      if (fFolio && rFolio.indexOf(fFolio) === -1) continue;
      if (fUsr && rUsr.indexOf(fUsr) === -1) continue;

      // 2. LÓGICA OR PARA CURP / CUENTA (Búsqueda Parcial)
      if (fCurp !== "" || fCuenta !== "") {
        var rCurp = (IDX_CURP > -1 && row[IDX_CURP]) ? row[IDX_CURP].toString().toUpperCase().replace(/\s+/g, '') : "";
        var rCuenta = (IDX_CUENTA > -1 && row[IDX_CUENTA]) ? row[IDX_CUENTA].toString().toUpperCase().replace(/\s+/g, '') : "";
        
        var matchO = false;
        if (fCurp !== "" && rCurp.indexOf(fCurp) > -1) matchO = true;
        if (fCuenta !== "" && rCuenta.indexOf(fCuenta) > -1) matchO = true;
        
        if (!matchO) continue; 
      }

      // 3. REGLA FLEXIBLE (Tienda)
      if (filtros.tienda) {
        var rTienda = (IDX_TIENDA > -1 && row[IDX_TIENDA]) ? row[IDX_TIENDA].toString().toLowerCase() : "";
        if (rTienda.indexOf(filtros.tienda.toLowerCase()) === -1) continue;
      }

      // 4. REGLA DE CORREO (Alias)
      if (filtros.correo) {
        var aliasCelda = (IDX_CORREO > -1 && row[IDX_CORREO]) ? row[IDX_CORREO].toString().toLowerCase().split('@')[0] : "";
        if (aliasCelda.indexOf(filtros.correo.toLowerCase()) === -1) continue;
      }

      // 5. DESPLEGABLES MÚLTIPLES (Arrays)
      if (filtros.tipoId && filtros.tipoId.length > 0) {
        var rTipoId = (IDX_TIPO_ID > -1 && row[IDX_TIPO_ID]) ? row[IDX_TIPO_ID].toString().trim() : "";
        if (filtros.tipoId.indexOf(rTipoId) === -1) continue;
      }
      if (filtros.tipoCaso && filtros.tipoCaso.length > 0) {
        var rTipoCaso = (IDX_TIPO_CASO > -1 && row[IDX_TIPO_CASO]) ? row[IDX_TIPO_CASO].toString().trim() : "";
        if (filtros.tipoCaso.indexOf(rTipoCaso) === -1) continue;
      }
      if (filtros.clasificacion && filtros.clasificacion.length > 0) {
        var rClasif = (IDX_CLASIF > -1 && row[IDX_CLASIF]) ? row[IDX_CLASIF].toString().trim() : "";
        if (filtros.clasificacion.indexOf(rClasif) === -1) continue;
      }

      // 6. WILDCARD (Palabra Clave en Textos Libres)
      if (fPalabra !== "") {
        var rComentarios = (IDX_COMENTARIOS > -1 && row[IDX_COMENTARIOS]) ? row[IDX_COMENTARIOS].toString().toLowerCase() : "";
        var rIndicaciones = (IDX_INDICACIONES > -1 && row[IDX_INDICACIONES]) ? row[IDX_INDICACIONES].toString().toLowerCase() : "";
        
        if (rComentarios.indexOf(fPalabra) === -1 && rIndicaciones.indexOf(fPalabra) === -1) continue; 
      }

      // 7. RANGO DE FECHAS
      var valMarca = (IDX_MARCA > -1 && row[IDX_MARCA]) ? new Date(row[IDX_MARCA]) : null;
      if (valMarca && !isNaN(valMarca.getTime())) {
        if (filtros.fechaInicio && valMarca.getTime() < new Date(filtros.fechaInicio + "T00:00:00").getTime()) continue;
        if (filtros.fechaFin && valMarca.getTime() > new Date(filtros.fechaFin + "T23:59:59").getTime()) continue;
      } else if (filtros.fechaInicio || filtros.fechaFin) {
         continue; 
      }

      // 8. TIEMPO DE RESOLUCIÓN Y SLA
      var valInicioAtn = (IDX_INICIO_ATN > -1 && row[IDX_INICIO_ATN]) ? new Date(row[IDX_INICIO_ATN]) : null;
      var valAtencion = (IDX_ATENCION > -1 && row[IDX_ATENCION]) ? new Date(row[IDX_ATENCION]) : null;
      var minsResolucion = 0;

      if (valInicioAtn && valAtencion && !isNaN(valInicioAtn.getTime()) && !isNaN(valAtencion.getTime())) {
        minsResolucion = (valAtencion.getTime() - valInicioAtn.getTime()) / 60000;
      }

      if (filtros.resCondicion && filtros.resMinutos !== "") {
        var minFiltro = parseFloat(filtros.resMinutos);
        if (!valInicioAtn || !valAtencion) continue; 
        if (filtros.resCondicion === 'mayor' && minsResolucion <= minFiltro) continue;
        if (filtros.resCondicion === 'menor' && minsResolucion >= minFiltro) continue;
      }

      // Formateo limpio
      var txtMarca = (valMarca && !isNaN(valMarca.getTime())) ? Utilities.formatDate(valMarca, GLOBAL_TIMEZONE, "dd/MM/yyyy HH:mm:ss") : "-";
      var txtInicio = (valInicioAtn && !isNaN(valInicioAtn.getTime())) ? Utilities.formatDate(valInicioAtn, GLOBAL_TIMEZONE, "dd/MM/yyyy HH:mm:ss") : "-";
      var txtAtencion = (valAtencion && !isNaN(valAtencion.getTime())) ? Utilities.formatDate(valAtencion, GLOBAL_TIMEZONE, "dd/MM/yyyy HH:mm:ss") : "-";

      // Empuje al array de resultados
      resultados.push({
        fila: i + 2,
        folio: (IDX_FOLIO > -1 && row[IDX_FOLIO]) ? row[IDX_FOLIO].toString() : "S/N",
        curp: (IDX_CURP > -1 && row[IDX_CURP]) ? row[IDX_CURP].toString() : "-",
        cuenta: (IDX_CUENTA > -1 && row[IDX_CUENTA]) ? row[IDX_CUENTA].toString() : "-",
        tienda: (IDX_TIENDA > -1 && row[IDX_TIENDA]) ? row[IDX_TIENDA].toString() : "-",
        correo: (IDX_CORREO > -1 && row[IDX_CORREO]) ? row[IDX_CORREO].toString() : "-",
        usr: (IDX_USR > -1 && row[IDX_USR]) ? row[IDX_USR].toString() : "-",
        tipoId: (IDX_TIPO_ID > -1 && row[IDX_TIPO_ID]) ? row[IDX_TIPO_ID].toString() : "-",
        tipoCaso: (IDX_TIPO_CASO > -1 && row[IDX_TIPO_CASO]) ? row[IDX_TIPO_CASO].toString() : "-",
        clasificacion: (IDX_CLASIF > -1 && row[IDX_CLASIF]) ? row[IDX_CLASIF].toString() : "-",
        marca: txtMarca,          
        inicioAtn: txtInicio,     
        atencion: txtAtencion,    
        resolucion: minsResolucion > 0 ? minsResolucion.toFixed(2) : "0", 
        repetido: (IDX_REPETIDO > -1 && row[IDX_REPETIDO]) ? row[IDX_REPETIDO].toString() : "-",
        tipoCli: (IDX_TIPO_CLI > -1 && row[IDX_TIPO_CLI]) ? row[IDX_TIPO_CLI].toString() : "-",
        servicio: (IDX_SERVICIO > -1 && row[IDX_SERVICIO]) ? row[IDX_SERVICIO].toString() : "-",
        comentarios: (IDX_COMENTARIOS > -1 && row[IDX_COMENTARIOS]) ? row[IDX_COMENTARIOS].toString() : "-",
        indicaciones: (IDX_INDICACIONES > -1 && row[IDX_INDICACIONES]) ? row[IDX_INDICACIONES].toString() : "-",
        _tsMarca: valMarca && !isNaN(valMarca.getTime()) ? valMarca.getTime() : 0  
      });
    }

    resultados.sort(function(a, b) { return a._tsMarca - b._tsMarca; });
    return { exito: true, data: resultados };

  } catch (e) {
    console.error(`[EXPLORADOR - ERROR] Fallo en la búsqueda: ${e.toString()}`);
    return { exito: false, mensaje: CFG_SERVICIOS.MSG_ERROR_BUSQUEDA };
  }
}


/**
 * Busca los casos en la hoja para llenar la tabla.
 * (Versión Optimizada para la Fila Fantasma - Cero indexOf)
 */
function obtenerCasosDinamicos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  // Usamos el límite del Bot o la última fila
  var lastRow = IDX_FILA_ULTIMOFOLIO > 1 ? IDX_FILA_ULTIMOFOLIO : sheet.getLastRow();
  if (lastRow < 2) return [];

  // Leemos según la configuración de la Mesa de Control
  var limite = parseInt(CEREBRO.filas) || 1000;
  var numRows = Math.min(lastRow - 1, limite);
  var startRow = lastRow - numRows + 1;
  
  // getValues puro para velocidad
  var data = sheet.getRange(startRow, 1, numRows, TODAS_LAS_COLUMNAS.length).getValues();
  
  var ahora = new Date();
  var limiteLimboMs = (parseFloat(CEREBRO.limbo) || 15) * 60000;
  var casos = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var tieneUSR = (IDX_USR > -1 && row[IDX_USR]) ? row[IDX_USR].toString().trim() !== "" : false;
    var tieneAtencion = (IDX_ATENCION > -1 && row[IDX_ATENCION]) ? row[IDX_ATENCION].toString().trim() !== "" : false;
    var tieneClasif = (IDX_CLASIF > -1 && row[IDX_CLASIF]) ? row[IDX_CLASIF].toString().trim() !== "" : false;
    
    // Filtro del Bot: Asumimos que el Bot debe haber dejado huella para ser operable
    var tieneBot = (IDX_COMENTARIOS > -1 && row[IDX_COMENTARIOS]) ? row[IDX_COMENTARIOS].toString().trim() !== "" : false;

    // ESTADO: PENDIENTE NORMAL
    var esCasoPendienteNormal = (!tieneAtencion && !tieneUSR && tieneBot);
    
    // ESTADO: LIMBO (Secuestrado)
    var esLimbo = false;
    if ((!tieneAtencion || !tieneClasif) && tieneUSR && tieneBot) {
      var tInic = (IDX_INICIO_ATN > -1 && row[IDX_INICIO_ATN]) ? new Date(row[IDX_INICIO_ATN]) : null;
      if (tInic && !isNaN(tInic.getTime())) {
        if ((ahora.getTime() - tInic.getTime()) > limiteLimboMs) {
          esLimbo = true;
        }
      }
    }

    // SI ES OPERABLE, LO EMPAQUETAMOS
    if (esCasoPendienteNormal || esLimbo) {
      var obj = {};
      // Llenamos el objeto tal como lo espera el render de HTML
      for (var c = 0; c < TODAS_LAS_COLUMNAS.length; c++) {
        var nombreColumna = TODAS_LAS_COLUMNAS[c];
        var val = row[c];
        
        // Si es fecha, la formateamos para que se vea bonita en la tabla
        if (val instanceof Date) {
          val = Utilities.formatDate(val, GLOBAL_TIMEZONE, "dd/MM/yyyy HH:mm:ss");
        }
        obj[nombreColumna] = val !== undefined && val !== null ? val : "";
      }

      casos.push({ 
        numeroFila: startRow + i, 
        datos: obj, 
        esLimbo: esLimbo, 
        usuarioOriginal: tieneUSR ? row[IDX_USR].toString().trim() : "" 
      });
    }
  }
  return casos;
}
