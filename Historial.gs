/**
 * Módulo de Historial V1.0
 * Búsqueda rápida en toda la base de datos.
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

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rCurp = row[idxCurp] ? row[idxCurp].toString().trim() : "";
    var rCuenta = row[idxCuenta] ? row[idxCuenta].toString().trim() : "";

    if ((sCurp !== "" && rCurp === sCurp) || (sCuenta !== "" && rCuenta === sCuenta)) {
      historial.push({ 
        fecha: Utilities.formatDate(new Date(row[idxFecha]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"), 
        atencion: row[idxAtencion] ? Utilities.formatDate(new Date(row[idxAtencion]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") : "En Proceso", 
        folio: row[idxFolio], 
        usr: row[idxUsr], 
        clasificacion: row[idxClasif], 
        indicaciones: row[idxIndic] 
      });
    }
  }
  
  // Ordenar por folio de forma descendente (el más reciente primero)
  historial.sort((a, b) => b.folio - a.folio);
  return historial;
}
