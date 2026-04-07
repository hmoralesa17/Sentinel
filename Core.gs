/**
 * Módulo Core y Configuración V2.1 (Conectado a la Mesa y con Variables Globales)
 * Archivo Principal de Configuración y Control de Acceso
 */

/**
 * Verifica si el usuario activo tiene privilegios de administrador.
 */
function verificarSiEsAdmin() {
  var prefijo = GLOBAL_EMAIL.split('@')[0];
  return ADMINS.indexOf(prefijo) > -1;
}

/**
 * Función de arranque (Punto de entrada de la aplicación)
 */
function doGet() {
  try {
    if (verificarAcceso(GLOBAL_EMAIL)) {
      var template = HtmlService.createTemplateFromFile('Index');
      return template.evaluate()
        .setTitle(CEREBRO.appTitle) // Lee el Título de tu Mesa de Control
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } else {
      return crearPantallaBloqueo(GLOBAL_EMAIL);
    }
  } catch (e) {
    console.error("Fallo Crítico doGet: " + e.toString());
    return crearPantallaError(e, GLOBAL_EMAIL);
  }
}

/** * Función verificación de acceso a la plataforma 
 */
function verificarAcceso(email) {
  if (!email || typeof email !== 'string') {
    console.warn("Intento de acceso con email inválido: " + email);
    return false;
  }
  return USUARIOS_AUTORIZADOS.indexOf(email) > -1; // "email" ya viene en minúsculas desde GLOBAL_EMAIL
}

/**
 * Función auxiliar para incluir archivos HTML (CSS y JS)
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    console.warn("⚠️ Archivo no encontrado: " + filename);
    return "<script>console.error('Fallo al cargar componente: " + filename + "');</script>";
  }
}

/**
 * Obtiene la versión de la app directamente del Cerebro
 */
function obtenerVersionFormateada() {
  return VERSION_APP; 
}


/**
 * Lee la hoja 'Config BOT' (o como se llame en la Mesa) y crea un diccionario global.
 * Columna B = Llave (Nombre de la variable)
 * Columna C = Valor
 */
function leerConfiguracionBot() {
  var diccionarioBot = {};
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    
    // 👇 USAMOS LA VARIABLE GLOBAL DINÁMICA
    var sheetBot = ss.getSheetByName(BOT_SHEET_NAME); 
    
    if (!sheetBot) return diccionarioBot;

    var ultimaFilaBot = sheetBot.getLastRow();
    if (ultimaFilaBot < 2) return diccionarioBot;

    // Leemos de la fila 2, columna 2 (B) hasta la columna 3 (C)
    var dataBot = sheetBot.getRange(2, 2, ultimaFilaBot - 1, 2).getValues();

    for (var i = 0; i < dataBot.length; i++) {
      var nombreVariable = dataBot[i][0] ? dataBot[i][0].toString().trim() : "";
      var valorVariable = dataBot[i][1];
      
      if (nombreVariable !== "") {
        diccionarioBot[nombreVariable] = valorVariable;
      }
    }
  } catch (e) {
    console.error("Fallo al leer hoja del Bot (" + BOT_SHEET_NAME + "): " + e.toString());
  }
  return diccionarioBot;
}
