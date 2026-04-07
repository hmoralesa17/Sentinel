/**
 * Módulo de Historial V2.0 (Ultra-Optimizado)
 * Búsqueda rápida conectada a la Mesa de Control (Variables Globales).
 */
function obtenerHistorialCaso(curp, cuenta) {
  // 1. VALIDACIÓN TEMPRANA (Early Exit)
  var sCurp = curp ? curp.toString().trim().toUpperCase() : "";
  var sCuenta = cuenta ? cuenta.toString().trim().toUpperCase() : "";
  
  // Si no hay nada que buscar, abortamos sin tocar Google Sheets
  if (sCurp === "" && sCuenta === "") return [];

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  // 2. EL ESCUDO DEL ROBOT (Lectura estricta)
  var endRow = IDX_FILA_ULTIMOFOLIO > 1 ? IDX_FILA_ULTIMOFOLIO : sheet.getLastRow();
  if (endRow < 2) return [];

  // Descargamos SOLO los datos reales, saltando los encabezados (Fila 2 en adelante)
  var data = sheet.getRange(2, 1, endRow - 1, TODAS_LAS_COLUMNAS.length).getValues();
  var historial = [];

  // 3. BÚSQUEDA VECTORIAL
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    // Extracción directa usando los Índices Globales de tu Mesa de Control
    var rCurp = (IDX_CURP > -1 && row[IDX_CURP]) ? row[IDX_CURP].toString().trim().toUpperCase() : "";
    var rCuenta = (IDX_CUENTA > -1 && row[IDX_CUENTA]) ? row[IDX_CUENTA].toString().trim().toUpperCase() : "";

    // ¿Hay coincidencia?
    if ((sCurp !== "" && rCurp === sCurp) || (sCuenta !== "" && rCuenta === sCuenta)) {
      
      // Manejo de la Fecha de Creación
      var rawFecha = (IDX_MARCA > -1) ? row[IDX_MARCA] : "";
      var tsOrdenamiento = 0; 
      var strFecha = rawFecha;
      
      if (rawFecha instanceof Date) {
        // Usamos la constante GLOBAL_TIMEZONE para cuadrar con el resto del sistema
        strFecha = Utilities.formatDate(rawFecha, GLOBAL_TIMEZONE, "dd/MM/yyyy HH:mm:ss");
        tsOrdenamiento = rawFecha.getTime();
      } else if (rawFecha !== "") {
        tsOrdenamiento = new Date(rawFecha).getTime() || 0;
      }

      // Manejo de la Fecha de Atención
      var rawAtn = (IDX_ATENCION > -1) ? row[IDX_ATENCION] : "";
      var strAtencion = "En Proceso";
      
      if (rawAtn instanceof Date) {
        strAtencion = Utilities.formatDate(rawAtn, GLOBAL_TIMEZONE, "dd/MM/yyyy HH:mm:ss");
      } else if (rawAtn !== "") {
        strAtencion = rawAtn.toString();
      }

      // Inyección al Historial
      historial.push({ 
        fecha: strFecha, 
        atencion: strAtencion, 
        folio: (IDX_FOLIO > -1 && row[IDX_FOLIO]) ? row[IDX_FOLIO].toString() : "S/N", 
        curp: rCurp,      
        cuenta: rCuenta,  
        usr: (IDX_USR > -1 && row[IDX_USR]) ? row[IDX_USR].toString() : "", 
        clasificacion: (IDX_CLASIF > -1 && row[IDX_CLASIF]) ? row[IDX_CLASIF].toString() : "", 
        indicaciones: (IDX_INDICACIONES > -1 && row[IDX_INDICACIONES]) ? row[IDX_INDICACIONES].toString() : "",
        _timestamp: tsOrdenamiento 
      });
    }
  }
  
  // 4. Ordenamiento matemático (Descendente: más reciente primero)
  historial.sort(function(a, b) { return b._timestamp - a._timestamp; });
  
  return historial;
}
