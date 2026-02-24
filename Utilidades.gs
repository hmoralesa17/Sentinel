/**
 * Módulo de Utilidades V1.0
 * Funciones de apoyo y soporte para el sistema.
 */

/**
 * Obtiene la información detallada del usuario desde la hoja de equipo.
 */
function obtenerInfoUsuario() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  try {
    var ss = SpreadsheetApp.openById(SHEET_EQUIPO_ID);
    var sheet = ss.getSheetByName(SHEET_EQUIPO_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxCorreo = headers.indexOf("Correo"), 
        idxNombre = headers.indexOf("Nombre Completo"), 
        idxIDVPL = headers.indexOf("ID_VPL"), 
        idxUSR = headers.indexOf("USR"), 
        idxNoEmpleado = headers.indexOf("No_Empleado");
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][idxCorreo].toString().toLowerCase().trim() === email) {
        return {
          nombreCompleto: data[i][idxNombre],
          vpl_usr: data[i][idxIDVPL] + " - " + data[i][idxUSR],
          noEmpleado: data[i][idxNoEmpleado],
          usrCode: data[i][idxUSR]
        };
      }
    }
  } catch (e) {
    console.error("Error en obtenerInfoUsuario: " + e.toString());
  }
  return { nombreCompleto: email, vpl_usr: "N/A", noEmpleado: "N/A", usrCode: "USR" };
}

/**
 * Obtiene el catálogo de clasificaciones e indicaciones predefinidas.
 */
function obtenerCatalogoNotas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(CATALOG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.map(row => ({ clasificacion: row[0], indicacion: row[1] }));
}

/**
 * Convierte minutos decimales a formato de tiempo legible (H:MM:SS o M:SS).
 */
function formatearTiempo(minutos) {
  if (!minutos || minutos < 0 || isNaN(minutos)) return "0:00";
  var totalSeg = Math.round(minutos * 60);
  var h = Math.floor(totalSeg / 3600), 
      m = Math.floor((totalSeg % 3600) / 60), 
      s = totalSeg % 60;
  var sStr = s < 10 ? "0" + s : s;
  
  if (h > 0) {
    var mStr = m < 10 ? "0" + m : m;
    return h + ":" + mStr + ":" + sStr;
  }
  return m + ":" + sStr;
}

/**
 * Normaliza valores de resolución (texto o decimal) a minutos para cálculos.
 */
function normalizarResolucionAMinutos(valor) {
  if (!valor || valor === "" || valor === "0") return 0;
  var str = valor.toString().trim();
  
  // Si ya viene en formato HH:MM:SS
  if (str.indexOf(':') > -1) {
    var partes = str.split(':');
    var h = parseInt(partes[0]) || 0;
    var m = parseInt(partes[1]) || 0;
    var s = parseInt(partes[2]) || 0;
    return (h * 60) + m + (s / 60);
  }
  
  // Si viene como decimal de Excel (días)
  var num = parseFloat(str.replace(',', '.'));
  return !isNaN(num) ? num * 1440 : 0;
}

/**
 * Genera la interfaz de acceso denegado con el estilo visual de la marca.
 */
function crearPantallaBloqueo(email) {
  var nombreUser = email.split('@')[0].toUpperCase();
  
  var html = `
  <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #fdf4fa; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background-color: #ffffff; border-top: 5px solid #D40099; padding: 40px; border-radius: 12px; max-width: 480px; text-align: center; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
        h1 { color: #78256F; margin-top: 0; font-size: 24px; font-weight: 800; }
        p { color: #555; font-size: 15px; line-height: 1.6; }
        .user-name { color: #D40099; font-weight: 800; }
        .footer { margin-top: 30px; font-size: 11px; color: #aaa; border-top: 1px solid #f1f1f1; padding-top: 20px; }
      </style>
    </head>
    <body>
      <div class="card">
        <h1>Acceso Restringido</h1>
        <p>Hola <span class="user-name">${nombreUser}</span>,</p>
        <p>Actualmente tu usuario no cuenta con los permisos necesarios para acceder a esta aplicación.</p>
        <p>Por favor, contacta a <b>hmoralesa@liverpool.com.mx</b> para gestionar tu autorización correspondiente.</p>
        <div class="footer">ID de cuenta detectado: ${email}</div>
      </div>
    </body>
  </html>`;
  
  return HtmlService.createHtmlOutput(html).setTitle("Acceso Denegado");
}

/**
 * Genera una pantalla de error con la identidad visual de Sentinel.
 * Basado en el diseño de crearPantallaBloqueo [1] y paleta de colores [3].
 */
function crearPantallaError(error, email) {
  var nombreUser = email ? email.split('@')[0].toUpperCase() : "USUARIO";
  
  // Estilos en línea para asegurar que se vean incluso si falla la carga de CSS externo
  var html = `
    <div style="font-family: 'Segoe UI', sans-serif; text-align: center; padding: 40px; color: #78256F;">
      <h1 style="color: #e74c3c; margin-bottom: 10px;">⚠️ Error de Carga</h1>
      
      <h3 style="color: #D40099;">Hola ${nombreUser},</h3>
      
      <p style="font-size: 16px;">
        Sentinel encontró un problema inesperado y no puede iniciar.
      </p>

      <div style="background-color: #f4f6f9; border-left: 5px solid #e74c3c; padding: 15px; margin: 20px auto; max-width: 600px; text-align: left; color: #333; font-family: monospace;">
        <strong>Detalle Técnico:</strong><br>
        ${error.toString()}
      </div>

      <p>Por favor, envía una captura de esta pantalla a <strong>hmoralesa@liverpool.com.mx</strong>.</p>
      
      <hr style="border: 0; border-top: 1px solid #E4D3E2; margin: 30px 0;">
      <small style="color: #BF97BA;">Cuenta detectada: ${email}</small>
    </div>
  `;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Error - Sentinel Fraudes")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
