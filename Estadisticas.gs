/**
 * Módulo de Análisis V4.0 (Versión Final Control Total)
 * 1. Cuadre absoluto con el Sheet (Folios Hoy).
 * 2. Detección de rezago de días anteriores (Pendientes/Gestión).
 * 3. Independencia del BOT (Si el bot no reclama, es Pendiente).
 */
function obtenerEstadisticasHoy() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  // --- 1. LÍMITES DE LECTURA (BASADO EN CONFIG BOT) ---
  var endRow = IDX_FILA_ULTIMOFOLIO; 
  var startRow = IDX_FILA_PRIMER_FOLIO_XDIAS;

  if (endRow < 2 || startRow < 2) return {};

  var numRows = endRow - startRow + 1;
  if (numRows < 1) return {};

  var data = sheet.getRange(startRow, 1, numRows, TODAS_LAS_COLUMNAS.length).getValues();

  // --- 2. INICIALIZACIÓN DE VARIABLES ---
  var META_SLA_MINUTOS = parseFloat(CEREBRO.sla); 
  var usrActual = DICCIONARIO_USUARIOS[GLOBAL_EMAIL] || "";
  var hoy = new Date();
  var dHoy = hoy.getDate(), mHoy = hoy.getMonth(), yHoy = hoy.getFullYear();
  
  var stats = {
    foliosHoy: 0, atendidosHoy: 0, pendientesHoy: 0, enGestionHoy: 0,
    promedioAtencion: "0:00", promedioTotal: "0:00", ultimoCaso: "---", primerCaso: "---",
    maxAtencion: "---", viejoPendiente: "---",
    atendidosBot: 0, promedioBot: "0:00", ultimoBot: "---", primerBot: "---",
    miAtendidos: 0, miPromedio: "0:00", miMin: "---", miMax: "---",
    miPrimerAtn: "---", miUltimoAtn: "---",
    cumplimientoSLA: "0%", miSLA: "0%"
  };
  
  var sumaResHumano = 0, sumaResMi = 0, sumaResBot = 0;
  var maxResTotal = 0, minResMi = Infinity, maxResMi = 0;
  var atendidosHumanos = 0, bajoSLA_General = 0, bajoSLA_Mi = 0;
  var tsGlobal = [], tsBot = [], tsMiAtencion = [], tsPendientes = [];

  // --- 3. MOTOR DE PROCESAMIENTO ---
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var valMarca = (IDX_MARCA > -1) ? row[IDX_MARCA] : null;
    if (!valMarca || valMarca === "") continue;

    var tObj = (valMarca instanceof Date) ? valMarca : new Date(valMarca);
    if (isNaN(tObj.getTime())) continue; 

    var esDeHoy = (tObj.getDate() === dHoy && tObj.getMonth() === mHoy && tObj.getFullYear() === yHoy);

    var usrCode = (IDX_USR > -1 && row[IDX_USR]) ? row[IDX_USR].toString().trim() : "";
    var tieneUSR = usrCode !== "";
    var tieneAtencion = (IDX_ATENCION > -1 && row[IDX_ATENCION]) ? row[IDX_ATENCION].toString().trim() !== "" : false;
    var tieneClasificacion = (IDX_CLASIF > -1 && row[IDX_CLASIF]) ? row[IDX_CLASIF].toString().trim() !== "" : false;
    var esBot = usrCode.toUpperCase() === "BOT";
    
    var valResolucion = (IDX_SLA > -1) ? row[IDX_SLA] : 0;
    var resMinutos = normalizarResolucionAMinutos(valResolucion);

    // A. Conteo Total (Solo Hoy)
    if (esDeHoy) {
      stats.foliosHoy++; 
      tsGlobal.push(tObj);
    }

    // B. Lógica de Estados (Pendientes y Gestión - Incluye Rezago)
    var estaTerminado = (tieneUSR && tieneAtencion && tieneClasificacion);

    if (!estaTerminado) {
      if (tieneUSR && !tieneAtencion) {
        stats.enGestionHoy++;
      } else if (!tieneUSR) {
        stats.pendientesHoy++; // Si no tiene USR, es pendiente (Bot caído o folio nuevo)
      }
      tsPendientes.push(tObj); 
    }

    // C. Productividad (Solo lo finalizado hoy)
    if (estaTerminado && esDeHoy) {
      stats.atendidosHoy++;

      if (esBot) {
        stats.atendidosBot++;
        sumaResBot += resMinutos;
        tsBot.push(tObj);
      } else {
        atendidosHumanos++;
        sumaResHumano += resMinutos;
        if (resMinutos > maxResTotal) maxResTotal = resMinutos;
        if (resMinutos <= META_SLA_MINUTOS) bajoSLA_General++; 

        if (usrCode === usrActual) {
          stats.miAtendidos++;
          sumaResMi += resMinutos;
          if (resMinutos < minResMi) minResMi = resMinutos;
          if (resMinutos > maxResMi) maxResMi = resMinutos;
          if (resMinutos <= META_SLA_MINUTOS) bajoSLA_Mi++; 
          tsMiAtencion.push(tObj);
        }
      }
    }
  } 

  // --- 4. CÁLCULOS FINALES ---
  if (atendidosHumanos > 0) {
    stats.promedioAtencion = formatearTiempo(sumaResHumano / atendidosHumanos);
    stats.maxAtencion = formatearTiempo(maxResTotal);
    stats.cumplimientoSLA = ((bajoSLA_General / atendidosHumanos) * 100).toFixed(0) + "%";
  }
  
  var totalAtendidos = atendidosHumanos + stats.atendidosBot;
  if (totalAtendidos > 0) {
    stats.promedioTotal = formatearTiempo((sumaResHumano + sumaResBot) / totalAtendidos);
  }
  
  if (stats.miAtendidos > 0) {
    stats.miPromedio = formatearTiempo(sumaResMi / stats.miAtendidos);
    stats.miMin = formatearTiempo(minResMi);
    stats.miMax = formatearTiempo(maxResMi);
    stats.miSLA = ((bajoSLA_Mi / stats.miAtendidos) * 100).toFixed(0) + "%";
  }
  
  if (stats.atendidosBot > 0) {
    stats.promedioBot = formatearTiempo(sumaResBot / stats.atendidosBot);
  }
  
  const fmtH = function(d) { return Utilities.formatDate(d, GLOBAL_TIMEZONE, "HH:mm:ss"); };
  
  if (tsGlobal.length > 0) {
    stats.primerCaso = fmtH(new Date(Math.min.apply(null, tsGlobal)));
    stats.ultimoCaso = fmtH(new Date(Math.max.apply(null, tsGlobal)));
  }
  if (tsBot.length > 0) {
    stats.primerBot = fmtH(new Date(Math.min.apply(null, tsBot)));
    stats.ultimoBot = fmtH(new Date(Math.max.apply(null, tsBot)));
  }
  if (tsMiAtencion.length > 0) {
    stats.miPrimerAtn = fmtH(new Date(Math.min.apply(null, tsMiAtencion)));
    stats.miUltimoAtn = fmtH(new Date(Math.max.apply(null, tsMiAtencion)));
  }

  if (tsPendientes.length > 0) {
    var masViejo = new Date(Math.min.apply(null, tsPendientes));
    var textoViejo = (masViejo.getDate() !== hoy.getDate()) 
                     ? Utilities.formatDate(masViejo, GLOBAL_TIMEZONE, "dd/MM HH:mm")
                     : fmtH(masViejo);
    stats.viejoPendiente = textoViejo;
  }

  return stats;
}

/**
 * Busca los casos en la hoja para llenar la tabla.
 */
function obtenerCasosDinamicos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numRows = Math.min(lastRow - 1, FILAS_A_LEER);
  var startRow = lastRow - numRows + 1;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getDisplayValues();
  
  var colAtencionIdx = headers.indexOf("Atención"), colUSRIdx = headers.indexOf("USR"),
      colBotIdx = headers.indexOf("Comentarios Bot"), colFolioIdx = headers.indexOf("FOLIO"),
      colInicioAtnIdx = headers.indexOf("InicioAtencion"), colClasifIdx = headers.indexOf("Clasificación"),
      colIndicIdx = headers.indexOf("Indicaciones");

  var ahora = new Date(), casos = [];

  for (var i = 0; i < data.length; i++) {
    var fila = data[i];
    
    // Validaciones seguras (evita errores de nulos)
    var tieneAtencion = fila[colAtencionIdx] ? fila[colAtencionIdx].toString().trim() !== "" : false;
    var tieneUSR = fila[colUSRIdx] ? fila[colUSRIdx].toString().trim() !== "" : false;
    var tieneBot = fila[colBotIdx] ? fila[colBotIdx].toString() !== "" && fila[colBotIdx].toString() !== "0" : false;
    var tieneClasif = colClasifIdx > -1 && fila[colClasifIdx] ? fila[colClasifIdx].toString().trim() !== "" : true;
    var tieneIndic = colIndicIdx > -1 && fila[colIndicIdx] ? fila[colIndicIdx].toString().trim() !== "" : true;
    
    var esCasoPendienteNormal = (!tieneAtencion && !tieneUSR && tieneBot);
    var esLimbo = false;

    // Lógica del Limbo
    if ((!tieneAtencion || !tieneClasif || !tieneIndic) && tieneUSR && tieneBot) {
      if (colInicioAtnIdx > -1 && fila[colInicioAtnIdx] !== "") {
        try {
          var inicioAtnStr = fila[colInicioAtnIdx].toString();
          var fechaInicio;
          
          if (inicioAtnStr.includes("/") && inicioAtnStr.includes(":")) {
             var partes = inicioAtnStr.split(" ");
             var fP = partes[0].split("/");
             var hP = partes[1].split(":");
             fechaInicio = new Date(fP[2], fP[1]-1, fP[0], hP[0], hP[1], hP[2] || 0);
          } else {
             fechaInicio = new Date(inicioAtnStr);
          }
          
          if (!isNaN(fechaInicio.getTime())) {
            var diffMinutos = (ahora - fechaInicio) / (1000 * 60);
            if (diffMinutos > TOLERANCIA_LIMBO_MINUTOS) esLimbo = true; // Parametrizado
          }
        } catch(e) {}
      }
    }

    if (esCasoPendienteNormal || esLimbo) {
      var obj = {};
      TODAS_LAS_COLUMNAS.forEach(function(col) { 
        var idx = headers.indexOf(col); 
        obj[col] = idx > -1 ? fila[idx] : ""; 
      });
      casos.push({ 
        numeroFila: startRow + i, 
        datos: obj, 
        esLimbo: esLimbo, 
        usuarioOriginal: tieneUSR ? fila[colUSRIdx] : "" 
      });
    }
  }
  return casos;
}
