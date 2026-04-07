/**
 * Módulo de Análisis V5.1 (Control de Extremos y 3 Tiempos)
 * Archivo de referencia: Estadisticas_260406.gs
 */
function obtenerEstadisticasHoy() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  var endRow = IDX_FILA_ULTIMOFOLIO; 
  var startRow = IDX_FILA_PRIMER_FOLIO_XDIAS;

  if (endRow < 2 || startRow < 2) return {};
  var numRows = endRow - startRow + 1;
  if (numRows < 1) return {};

  var data = sheet.getRange(startRow, 1, numRows, TODAS_LAS_COLUMNAS.length).getValues();

  // --- CONFIGURACIÓN ---
  var META_SLA_MINUTOS = parseFloat(CEREBRO.sla) || 30; 
  var LIMBO_MINUTOS = parseFloat(CEREBRO.limbo) || 15;
  var usrActual = DICCIONARIO_USUARIOS[GLOBAL_EMAIL] || "";
  var ahora = new Date();
  var dHoy = ahora.getDate(), mHoy = ahora.getMonth(), yHoy = ahora.getFullYear();
  
  // CAJA DE DATOS AMPLIADA
  var stats = {
    foliosHoy: 0, atendidosHoy: 0, pendientesHoy: 0, enGestionHoy: 0, limboDetectado: 0,
    // ESPERA (General)
    promedioEspera: "0:00", maxEspera: "0:00", minEspera: "---",
    // GESTIÓN (General Humano)
    promedioGestion: "0:00", maxGestion: "0:00", minGestion: "---",
    // GESTIÓN (Bot)
    atendidosBot: 0, promedioBot: "0:00", maxBot: "---", minBot: "---",
    // MI GESTIÓN (Personal)
    miAtendidos: 0, miPromedioGestion: "0:00", miMaxGestion: "---", miMinGestion: "---",
    // TOTALES Y SLA
    promedioTotal: "0:00", cumplimientoSLA: "0%", miSLA: "0%",
    ultimoCaso: "---", primerCaso: "---", viejoPendiente: "---",
    miPrimerAtn: "---", miUltimoAtn: "---"
  };
  
  // Acumuladores Matemáticos
  var sumaEsp = 0, maxEsp = 0, minEsp = Infinity;
  var sumaGesHum = 0, maxGesHum = 0, minGesHum = Infinity;
  var sumaGesBot = 0, maxGesBot = 0, minGesBot = Infinity;
  var sumaGesMi = 0, maxGesMi = 0, minGesMi = Infinity;
  var sumaTotal = 0, bajoSLA = 0, bajoSLAMi = 0;
  
  var tsGlobal = [], tsBot = [], tsMiAtencion = [], tsPendientes = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var vMarca = (IDX_MARCA > -1) ? row[IDX_MARCA] : null;
    if (!vMarca || vMarca === "") continue;
    
    var tMarca = (vMarca instanceof Date) ? vMarca : new Date(vMarca);
    if (isNaN(tMarca.getTime())) continue; 

    var tInic = (IDX_INICIO_ATN > -1 && row[IDX_INICIO_ATN]) ? new Date(row[IDX_INICIO_ATN]) : null;
    var tFin = (IDX_ATENCION > -1 && row[IDX_ATENCION]) ? new Date(row[IDX_ATENCION]) : null;
    var esDeHoy = (tMarca.getDate() === dHoy && tMarca.getMonth() === mHoy && tMarca.getFullYear() === yHoy);

    var usrCode = (IDX_USR > -1 && row[IDX_USR]) ? row[IDX_USR].toString().trim() : "";
    var tieneUSR = usrCode !== "";
    var tieneInic = (tInic && !isNaN(tInic.getTime()));
    var tieneFin = (tFin && !isNaN(tFin.getTime()));
    var tieneClasif = (IDX_CLASIF > -1 && row[IDX_CLASIF]) ? row[IDX_CLASIF].toString().trim() !== "" : false;
    var esBot = usrCode.toUpperCase() === "BOT";

    if (esDeHoy) { stats.foliosHoy++; tsGlobal.push(tMarca); }

    // --- LÓGICA DE ESTADOS ---
    var terminado = tieneUSR && tieneFin && tieneClasif;
    if (!terminado) {
      if (tieneUSR && tieneInic && !tieneFin) {
        stats.enGestionHoy++;
        if ((ahora - tInic) / 60000 > LIMBO_MINUTOS) stats.limboDetectado++;
      } else if (!tieneUSR) {
        stats.pendientesHoy++;
      }
      tsPendientes.push(tMarca); 
    }

    // --- CÁLCULO DE MÉTRICAS (TERMINADOS HOY) ---
    if (terminado && esDeHoy) {
      stats.atendidosHoy++;
      var minEsp_F = tieneInic ? Math.max(0, (tInic - tMarca) / 60000) : 0;
      var minGes_F = tieneInic ? Math.max(0, (tFin - tInic) / 60000) : 0;
      var minTot_F = Math.max(0, (tFin - tMarca) / 60000);

      // ESPERA (Solo General)
      sumaEsp += minEsp_F;
      if (minEsp_F > maxEsp) maxEsp = minEsp_F;
      if (minEsp_F < minEsp) minEsp = minEsp_F;

      if (esBot) {
        stats.atendidosBot++;
        sumaGesBot += minGes_F;
        if (minGes_F > maxGesBot) maxGesBot = minGes_F;
        if (minGes_F < minGesBot) minGesBot = minGes_F;
        tsBot.push(tFin);
      } else {
        // EQUIPO HUMANO
        var atnH = (stats.atendidosHoy - stats.atendidosBot); 
        sumaGesHum += minGes_F;
        sumaTotal += minTot_F;
        if (minGes_F > maxGesHum) maxGesHum = minGes_F;
        if (minGes_F < minGesHum) minGesHum = minGes_F;
        if (minTot_F <= META_SLA_MINUTOS) bajoSLA++;

        // MI GESTIÓN
        if (usrCode === usrActual) {
          stats.miAtendidos++;
          sumaGesMi += minGes_F;
          if (minGes_F > maxGesMi) maxGesMi = minGes_F;
          if (minGes_F < minGesMi) minGesMi = minGes_F;
          if (minTot_F <= META_SLA_MINUTOS) bajoSLAMi++;
          tsMiAtencion.push(tFin);
        }
      }
    }
  } 

  // --- FORMATEO FINAL ---
  const fT = (m) => (m === Infinity || m === -Infinity) ? "---" : formatearTiempo(m);
  const fmtH = (d) => Utilities.formatDate(d, GLOBAL_TIMEZONE, "HH:mm:ss");

  // Espera General
  if (stats.atendidosHoy > 0) {
    stats.promedioEspera = fT(sumaEsp / stats.atendidosHoy);
    stats.maxEspera = fT(maxEsp);
    stats.minEspera = fT(minEsp);
  }

  // Gestión Equipo (Humano)
  var cantHum = stats.atendidosHoy - stats.atendidosBot;
  if (cantHum > 0) {
    stats.promedioGestion = fT(sumaGesHum / cantHum);
    stats.maxGestion = fT(maxGesHum);
    stats.minGestion = fT(minGesHum);
    stats.promedioTotal = fT(sumaTotal / cantHum);
    stats.cumplimientoSLA = ((bajoSLA / cantHum) * 100).toFixed(0) + "%";
  }

  // Gestión Bot
  if (stats.atendidosBot > 0) {
    stats.promedioBot = fT(sumaGesBot / stats.atendidosBot);
    stats.maxBot = fT(maxGesBot);
    stats.minBot = fT(minGesBot);
  }

  // Mi Gestión
  if (stats.miAtendidos > 0) {
    stats.miPromedioGestion = fT(sumaGesMi / stats.miAtendidos);
    stats.miMaxGestion = fT(maxGesMi);
    stats.miMinGestion = fT(minGesMi);
    stats.miSLA = ((bajoSLAMi / stats.miAtendidos) * 100).toFixed(0) + "%";
  }

  // Tiempos de Reloj
  if (tsGlobal.length > 0) {
    stats.primerCaso = fmtH(new Date(Math.min.apply(null, tsGlobal)));
    stats.ultimoCaso = fmtH(new Date(Math.max.apply(null, tsGlobal)));
  }
  if (tsMiAtencion.length > 0) {
    stats.miPrimerAtn = fmtH(new Date(Math.min.apply(null, tsMiAtencion)));
    stats.miUltimoAtn = fmtH(new Date(Math.max.apply(null, tsMiAtencion)));
  }
  if (tsPendientes.length > 0) {
    var vP = new Date(Math.min.apply(null, tsPendientes));
    stats.viejoPendiente = (vP.getDate() !== dHoy) ? Utilities.formatDate(vP, GLOBAL_TIMEZONE, "dd/MM HH:mm") : fmtH(vP);
  }

  return stats;
}
