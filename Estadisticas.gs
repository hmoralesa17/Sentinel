/**
 * Módulo de Análisis V1.2
 * Lectura acotada (3000 filas) con detección de rezago (ayer y anteriores).
 */

function obtenerEstadisticasHoy() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  // Ventana de 3000 filas para cubrir hoy y rezagos de días anteriores
  var numRows = Math.min(lastRow - 1, 1500);
  var startRow = lastRow - numRows + 1;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getDisplayValues();
  
  var idxFecha = headers.indexOf("Marca temporal");
  var idxUSR = headers.indexOf("USR");
  var idxAtencion = headers.indexOf("Atención");
  var idxRes = headers.indexOf("Resolución");
  var idxBot = headers.indexOf("Comentarios Bot");
  var idxClasif = headers.indexOf("Clasificación");

  var emailActual = Session.getActiveUser().getEmail().toLowerCase();
  var usrActual = DICCIONARIO_USUARIOS[emailActual] || "";
  var hoy = new Date();
  var dHoy = hoy.getDate(), mHoy = hoy.getMonth() + 1, yHoy = hoy.getFullYear();
  
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

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var fechaRowStr = row[idxFecha];
    if (!fechaRowStr) continue;
    
    // Parseo de fecha de la fila
    var fPartes = fechaRowStr.split(" ")[0].split("/");
    var dRow = parseInt(fPartes[0]), mRow = parseInt(fPartes[1]), yRow = parseInt(fPartes[2]);
    var esDeHoy = (dRow === dHoy && mRow === mHoy && yRow === yHoy);

    // Objeto de tiempo para comparaciones y chips
    var tObj = new Date(yRow, mRow - 1, dRow);
    if (fechaRowStr.split(" ")[1]) {
      var hP = fechaRowStr.split(" ")[1].split(":");
      tObj.setHours(parseInt(hP[0]), parseInt(hP[1]), parseInt(hP[2] || 0));
    }

    var usrCode = row[idxUSR].toString().trim();
    var tieneUSR = usrCode !== "";
    var tieneAtencion = row[idxAtencion].toString().trim() !== "";
    var tieneClasificacion = idxClasif > -1 ? row[idxClasif].toString().trim() !== "" : false;
    var esBot = usrCode.toUpperCase() === "BOT";
    var resMinutos = normalizarResolucionAMinutos(row[idxRes]);

    // --- 1. LÓGICA DE PRODUCTIVIDAD (SÓLO HOY) ---
    if (esDeHoy) {
      stats.foliosHoy++;
      tsGlobal.push(tObj);

      if (esBot) {
        stats.atendidosBot++;
        sumaResBot += resMinutos;
        tsBot.push(tObj);
      }

      if (tieneUSR && tieneAtencion && tieneClasificacion) {
        stats.atendidosHoy++;
        if (!esBot) {
          atendidosHumanos++;
          sumaResHumano += resMinutos;
          if (resMinutos > maxResTotal) maxResTotal = resMinutos;
          if (resMinutos <= 30) bajoSLA_General++;

          if (usrCode === usrActual) {
            stats.miAtendidos++;
            sumaResMi += resMinutos;
            if (resMinutos < minResMi) minResMi = resMinutos;
            if (resMinutos > maxResMi) maxResMi = resMinutos;
            if (resMinutos <= 30) bajoSLA_Mi++;
            tsMiAtencion.push(tObj);
          }
        }
      }
    }

    // --- 2. LÓGICA DE PENDIENTES Y REZAGO (TODA LA VENTANA) ---
    // Si es un caso válido (tiene comentarios bot)
    if (row[idxBot] !== "" && row[idxBot] !== "0") {
      // Si falta CUALQUIERA de los datos de cierre (Limbo o Pendiente puro)
      if (!tieneAtencion || !tieneUSR || !tieneClasificacion) {
        
        // Conteo para chip "En Gestión" (Si tiene usuario pero no cierre)
        if (tieneUSR && !tieneAtencion) {
          stats.enGestionHoy++;
        }
        
        // Conteo para chip "Pendientes" (Si nadie lo ha tomado)
        if (!tieneUSR && !tieneAtencion) {
          stats.pendientesHoy++;
        }
        
        // Lista para encontrar el folio más viejo (incluyendo días anteriores)
        tsPendientes.push(tObj);
      }
    }
  }
  
  // --- PROCESAMIENTO FINAL ---
  if (atendidosHumanos > 0) {
    stats.promedioAtencion = formatearTiempo(sumaResHumano / atendidosHumanos);
    stats.maxAtencion = formatearTiempo(maxResTotal);
    stats.cumplimientoSLA = ((bajoSLA_General / atendidosHumanos) * 100).toFixed(0) + "%";
  }
  var totalAtendidos = atendidosHumanos + stats.atendidosBot;
  if (totalAtendidos > 0) stats.promedioTotal = formatearTiempo((sumaResHumano + sumaResBot) / totalAtendidos);
  if (stats.miAtendidos > 0) {
    stats.miPromedio = formatearTiempo(sumaResMi / stats.miAtendidos);
    stats.miMin = formatearTiempo(minResMi);
    stats.miMax = formatearTiempo(maxResMi);
    stats.miSLA = ((bajoSLA_Mi / stats.miAtendidos) * 100).toFixed(0) + "%";
  }
  if (stats.atendidosBot > 0) stats.promedioBot = formatearTiempo(sumaResBot / stats.atendidosBot);
  
  const fmtH = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm:ss");
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

  // Chip "+ Viejo": Inteligente para mostrar fecha si es de otro día
  if (tsPendientes.length > 0) {
    var masViejo = new Date(Math.min.apply(null, tsPendientes));
    if (masViejo.getDate() !== hoy.getDate()) {
       // Si es rezago de ayer, mostramos el día y mes
       stats.viejoPendiente = Utilities.formatDate(masViejo, Session.getScriptTimeZone(), "dd/MM HH:mm");
    } else {
       stats.viejoPendiente = fmtH(masViejo);
    }
  }

  return stats;
}

function obtenerCasosDinamicos() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var numRows = Math.min(lastRow - 1, 1500);
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
    var tieneAtencion = fila[colAtencionIdx].trim() !== "", tieneUSR = fila[colUSRIdx].trim() !== "",
        tieneBot = fila[colBotIdx] !== "" && fila[colBotIdx] !== "0", tieneFolio = fila[colFolioIdx] !== "",
        tieneClasif = colClasifIdx > -1 ? fila[colClasifIdx].trim() !== "" : true,
        tieneIndic = colIndicIdx > -1 ? fila[colIndicIdx].trim() !== "" : true;
    
    var esCasoPendienteNormal = (!tieneAtencion && !tieneUSR && tieneBot );//&& tieneFolio); te arreglo despues dependes del bot
    var esLimbo = false;

    if ((!tieneAtencion || !tieneClasif || !tieneIndic) && tieneUSR && tieneBot ) {//&& tieneFolio
      if (colInicioAtnIdx > -1) {
        var inicioAtnStr = fila[colInicioAtnIdx];
        if (inicioAtnStr !== "") {
          try {
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
              if (diffMinutos > 10) esLimbo = true;
            }
          } catch(e) {}
        }
      }
    }

    if (esCasoPendienteNormal || esLimbo) {
      var obj = {};
      TODAS_LAS_COLUMNAS.forEach(function(col) { 
        var idx = headers.indexOf(col); 
        obj[col] = idx > -1 ? fila[idx] : ""; 
      });
      casos.push({ numeroFila: startRow + i, datos: obj, esLimbo: esLimbo, usuarioOriginal: tieneUSR ? fila[colUSRIdx] : "" });
    }
  }
  return casos;
}
