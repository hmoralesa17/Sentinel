/**
 * Módulo de Historial V1.0
 * Búsqueda rápida en toda la base de datos con ordenamiento seguro.
 */

function obtenerHistorialCaso(curp, cuenta) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // getValues es 10x más rápido que getDisplayValues para escanear toda la base
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var idxCurp = headers.indexOf("CURP"), 
      idxCuenta = headers.indexOf("Cuenta"), 
      idxFecha = headers.indexOf("Marca temporal"), 
      idxAtencion = headers.indexOf("Atención"), 
      idxFolio = headers.indexOf("FOLIO"), 
      idxUsr = headers.indexOf("USR"), 
      idxClasif = headers.indexOf("Clasificación"), 
      idxIndic = headers.indexOf("Indicaciones");

  var historial = [];
  var sCurp = curp ? curp.toString().trim() : "";
  var sCuenta = cuenta ? cuenta.toString().trim() : "";

  // Optimización: Empezamos en 1 para saltar los encabezados
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rCurp = row[idxCurp] ? row[idxCurp].toString().trim() : "";
    var rCuenta = row[idxCuenta] ? row[idxCuenta].toString().trim() : "";

    // Si coincide el CURP o la Cuenta
    if ((sCurp !== "" && rCurp === sCurp) || (sCuenta !== "" && rCuenta === sCuenta)) {
      
      // 1. Manejo seguro de la Fecha de Creación
      var rawFecha = row[idxFecha];
      var strFecha = rawFecha;
      var tsOrdenamiento = 0; // Timestamp para ordenar
      
      if (rawFecha instanceof Date) {
        strFecha = Utilities.formatDate(rawFecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        tsOrdenamiento = rawFecha.getTime();
      } else if (rawFecha !== "") {
        // Si no es un objeto fecha, intentamos convertirlo
        tsOrdenamiento = new Date(rawFecha).getTime() || 0;
      }

      // 2. Manejo seguro de la Fecha de Atención
      var rawAtn = row[idxAtencion];
      var strAtencion = "En Proceso";
      
      if (rawAtn instanceof Date) {
        strAtencion = Utilities.formatDate(rawAtn, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      } else if (rawAtn !== "") {
        strAtencion = rawAtn.toString();
      }

      historial.push({ 
        fecha: strFecha, 
        atencion: strAtencion, 
        folio: row[idxFolio] ? row[idxFolio].toString() : "S/N", 
        curp: rCurp,      // <--- NUEVO
        cuenta: rCuenta,  // <--- NUEVO
        usr: row[idxUsr] ? row[idxUsr].toString() : "", 
        clasificacion: row[idxClasif] ? row[idxClasif].toString() : "", 
        indicaciones: row[idxIndic] ? row[idxIndic].toString() : "",
        _timestamp: tsOrdenamiento 
      });
    }
  }
  
  // 3. Ordenar matemáticamente usando el timestamp oculto (Descendente: más reciente primero)
  historial.sort((a, b) => b._timestamp - a._timestamp);
  
  return historial;
}
