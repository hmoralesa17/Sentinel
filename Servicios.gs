/**
 * Módulo de Operaciones V1.3
 * Funciones de escritura, lectura protegida y liberación de casos.
 */

/**
 * Bloquea un caso para el usuario actual. Optimizada para alta concurrencia.
 */
function tomarCasoDinamico(numeroFila, esRescate) {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  
  // 1. TRABAJO PESADO AFUERA DEL CANDADO
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

  console.log(`[INTENTO] Usuario: ${idUsuario} | Folio: ${folio} (Fila: ${numeroFila}) | Rescate: ${esRescate}`);

  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(25000); 
    
    // ZONA CRÍTICA: SOLO LECTURA RÁPIDA Y ESCRITURA
    var usuarioActualEnCelda = sheet.getRange(numeroFila, colUSRIdx).getValue().toString().trim();
    
    // Si ya está tomado y no es rescate, abortar
    if (!esRescate && usuarioActualEnCelda !== "") {
      lock.releaseLock();
      console.warn(`[COLISIÓN] ${idUsuario} intentó tomar el Folio ${folio}, ya bloqueado por ${usuarioActualEnCelda}`);
      return { exito: false, mensaje: "⚠️ El caso ya fue tomado por: " + usuarioActualEnCelda };
    }

    // Escribir usuario
    sheet.getRange(numeroFila, colUSRIdx).setValue(idUsuario);
    if (colInicioAtnIdx > 0) {
      sheet.getRange(numeroFila, colInicioAtnIdx).setValue(new Date());
    }
    
    SpreadsheetApp.flush();
    lock.releaseLock(); 
    
    console.log(`[ÉXITO] Folio ${folio} asignado a ${idUsuario}`);
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock();
    console.error(`[ERROR FATAL] Fallo al asignar Folio ${folio}. Detalle: ${e.toString()}`);
    return { exito: false, mensaje: "Error de servidor (Tráfico Alto). Intenta de nuevo." }; 
  }
}

/**
 * ¡LA NUEVA FUNCIÓN QUE RECUPERAMOS!
 * Libera un caso si el usuario decide cancelar, borrando su nombre y la hora de inicio.
 * Se le agregó LockService para que no choque con otras operaciones.
 */
function liberarCasoDinamico(numeroFila) {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  console.log(`[LIBERAR - INTENTO] ${idUsuario} está cancelando la atención de la fila: ${numeroFila}`);

  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(15000); // 15 segundos de paciencia
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var headers = sheet.getRange("1:1").getValues()[0];
    
    var colUSRIdx = headers.indexOf("USR") + 1;
    var colInicioAtnIdx = headers.indexOf("InicioAtencion") + 1;

    // Borramos los datos
    if (colUSRIdx > 0) sheet.getRange(numeroFila, colUSRIdx).clearContent();
    if (colInicioAtnIdx > 0) sheet.getRange(numeroFila, colInicioAtnIdx).clearContent();
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    console.log(`[LIBERAR - ÉXITO] Caso en la fila ${numeroFila} liberado correctamente.`);
    return { exito: true };
    
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    console.error(`[LIBERAR - ERROR] Fallo al liberar la fila ${numeroFila}: ${e.toString()}`);
    return { exito: false, mensaje: e.toString() };
  }
}

/**
 * Finaliza la atención del caso. 
 * (Mantenemos nuestra versión superoptimizada que lee antes de bloquear)
 */
function finalizarCasoDinamico(numeroFila, clasificacion, indicaciones) {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  var idUsuario = DICCIONARIO_USUARIOS[email] || email;
  
  // TRABAJO PESADO AFUERA DEL CANDADO
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var headers = sheet.getRange("1:1").getValues()[0];
  
  var colFolioIdx = headers.indexOf("FOLIO") + 1;
  var colClasifIdx = headers.indexOf("Clasificación") + 1;
  var colIndicIdx = headers.indexOf("Indicaciones") + 1;
  var colAtnIdx = headers.indexOf("Atención") + 1;
  
  var folio = "Desconocido";
  if (colFolioIdx > 0) {
    folio = sheet.getRange(numeroFila, colFolioIdx).getValue().toString().trim();
  }

  console.log(`[FINALIZAR - INTENTO] ${idUsuario} va a cerrar el Folio: ${folio} con clasificación: ${clasificacion}`);

  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(25000);
    
    // ZONA CRÍTICA: SOLO ESCRITURA RÁPIDA
    if (colClasifIdx > 0) sheet.getRange(numeroFila, colClasifIdx).setValue(clasificacion);
    if (colIndicIdx > 0) sheet.getRange(numeroFila, colIndicIdx).setValue(indicaciones);
    if (colAtnIdx > 0) sheet.getRange(numeroFila, colAtnIdx).setValue(new Date()); 
    
    SpreadsheetApp.flush();
    lock.releaseLock();
    
    console.log(`[FINALIZAR - ÉXITO] Folio ${folio} cerrado correctamente por ${idUsuario}`);
    return { exito: true };

  } catch (e) { 
    if (lock.hasLock()) lock.releaseLock(); 
    console.error(`[FINALIZAR - ERROR] Fallo al cerrar Folio ${folio} (Fila: ${numeroFila}) por ${idUsuario}. Detalle: ${e.toString()}`);
    return { exito: false, mensaje: "Error de servidor (Tráfico Alto). Intenta guardar de nuevo." }; 
  }
}


/**
 * =======================================================
 * MÓDULO EXPLORADOR: MOTOR DE BÚSQUEDA AVANZADA
 * =======================================================
 */
function buscarFoliosExplorador(filtros) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // Función rápida para encontrar el número de columna
    var getIdx = function(name) { return headers.indexOf(name); };
    
    var idxFolio = getIdx("FOLIO");
    var idxCurp = getIdx("CURP");
    var idxCuenta = getIdx("Cuenta");
    var idxTienda = getIdx("Tienda");
    var idxCorreo = getIdx("Dirección de correo electrónico");
    var idxUsr = getIdx("USR");
    var idxTipoId = getIdx("Tipo de Identificación");
    var idxTipoCaso = getIdx("Tipo Caso");
    var idxClasif = getIdx("Clasificación");
    var idxMarca = getIdx("Marca temporal");
    var idxInicioAtn = getIdx("InicioAtencion");
    var idxAtencion = getIdx("Atención");
    var idxRepetido = getIdx("Repetido");
    var idxTipoCli = getIdx("Tipo Cliente");
    var idxServicio = getIdx("Servicio");
    var idxComentarios = getIdx("Comentarios Bot");
    var idxIndicaciones = getIdx("Indicaciones");

    var resultados = [];
    // Obtenemos la zona horaria de tu hoja para que las fechas coincidan
    var zonaHoraria = Session.getScriptTimeZone();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // 1. REGLA ESTRICTA (100% IDÉNTICOS)
      if (filtros.folio && row[idxFolio].toString().trim() !== filtros.folio.trim()) continue;
      if (filtros.curp && row[idxCurp].toString().trim() !== filtros.curp.trim()) continue;
      if (filtros.cuenta && row[idxCuenta].toString().trim() !== filtros.cuenta.trim()) continue;

      // 2. REGLA FLEXIBLE
      if (filtros.tienda && row[idxTienda].toString().toLowerCase().indexOf(filtros.tienda.toLowerCase()) === -1) continue;
      if (filtros.usr && row[idxUsr].toString().toLowerCase().indexOf(filtros.usr.toLowerCase()) === -1) continue;

      // 3. REGLA DE CORREO
      if (filtros.correo) {
        var correoCelda = row[idxCorreo] ? row[idxCorreo].toString().toLowerCase() : "";
        var aliasCelda = correoCelda.split('@')[0];
        if (aliasCelda.indexOf(filtros.correo.toLowerCase()) === -1) continue;
      }

      // 4. DESPLEGABLES
      if (filtros.tipoId && row[idxTipoId].toString().trim() !== filtros.tipoId) continue;
      if (filtros.tipoCaso && row[idxTipoCaso].toString().trim() !== filtros.tipoCaso) continue;
      if (filtros.clasificacion && row[idxClasif].toString().trim() !== filtros.clasificacion) continue;

      // 5. RANGO DE FECHAS
      var valMarca = row[idxMarca] ? new Date(row[idxMarca]) : null;
      if (valMarca && !isNaN(valMarca.getTime())) {
        if (filtros.fechaInicio) {
          var fIni = new Date(filtros.fechaInicio + "T00:00:00").getTime();
          if (valMarca.getTime() < fIni) continue;
        }
        if (filtros.fechaFin) {
          var fFin = new Date(filtros.fechaFin + "T23:59:59").getTime();
          if (valMarca.getTime() > fFin) continue;
        }
      } else if (filtros.fechaInicio || filtros.fechaFin) {
         continue; 
      }

      // 6. TIEMPO DE RESOLUCIÓN
      var valInicioAtn = row[idxInicioAtn] ? new Date(row[idxInicioAtn]) : null;
      var valAtencion = row[idxAtencion] ? new Date(row[idxAtencion]) : null;
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

      // ✨ CORRECCIÓN: Dar formato limpio a las fechas
      var txtMarca = (valMarca && !isNaN(valMarca.getTime())) ? Utilities.formatDate(valMarca, zonaHoraria, "dd/MM/yyyy HH:mm:ss") : "-";
      var txtInicio = (valInicioAtn && !isNaN(valInicioAtn.getTime())) ? Utilities.formatDate(valInicioAtn, zonaHoraria, "dd/MM/yyyy HH:mm:ss") : "-";
      var txtAtencion = (valAtencion && !isNaN(valAtencion.getTime())) ? Utilities.formatDate(valAtencion, zonaHoraria, "dd/MM/yyyy HH:mm:ss") : "-";

      // SI PASÓ LOS FILTROS, LO GUARDAMOS
      resultados.push({
        fila: i + 1,
        folio: row[idxFolio] ? row[idxFolio].toString() : "S/N",
        curp: row[idxCurp] ? row[idxCurp].toString() : "-",
        cuenta: row[idxCuenta] ? row[idxCuenta].toString() : "-",
        tienda: row[idxTienda] ? row[idxTienda].toString() : "-",
        correo: row[idxCorreo] ? row[idxCorreo].toString() : "-",
        usr: row[idxUsr] ? row[idxUsr].toString() : "-",
        tipoId: row[idxTipoId] ? row[idxTipoId].toString() : "-",
        tipoCaso: row[idxTipoCaso] ? row[idxTipoCaso].toString() : "-",
        clasificacion: row[idxClasif] ? row[idxClasif].toString() : "-",
        marca: txtMarca,          // Fecha limpia
        inicioAtn: txtInicio,     // Fecha limpia
        atencion: txtAtencion,    // Fecha limpia
        resolucion: minsResolucion > 0 ? minsResolucion.toFixed(2) : "0", 
        repetido: row[idxRepetido] ? row[idxRepetido].toString() : "-",
        tipoCli: row[idxTipoCli] ? row[idxTipoCli].toString() : "-",
        servicio: row[idxServicio] ? row[idxServicio].toString() : "-",
        comentarios: row[idxComentarios] ? row[idxComentarios].toString() : "-",
        indicaciones: row[idxIndicaciones] ? row[idxIndicaciones].toString() : "-",
        _tsMarca: valMarca && !isNaN(valMarca.getTime()) ? valMarca.getTime() : 0  // ✨ Variable oculta para ordenar
      });
    }

    // ✨ CORRECCIÓN: Ordenar por Marca Temporal (de menor a mayor = más viejo a más reciente)
    resultados.sort(function(a, b) {
      return a._tsMarca - b._tsMarca;
    });

    // Se eliminó el límite de seguridad, mandamos todo.
    return { exito: true, data: resultados };

  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}
