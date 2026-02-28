/**
 * Módulo de Utilidades V1.0
 * Funciones de apoyo y soporte para el sistema.
 */

/**
 * Obtiene la información detallada del usuario desde la hoja de equipo.
 * Si no encuentra al usuario o falla la conexión, devuelve un objeto genérico (Fallback).
 * @returns {Object} Objeto con nombreCompleto, vpl_usr, noEmpleado y usrCode.
 */
function obtenerInfoUsuario() {
  var email = Session.getActiveUser().getEmail().toLowerCase();
  
  // 1. LOG DE INTENTO: Registramos quién está intentando cargar su perfil
  console.log(`[INFO USUARIO - INTENTO] Buscando datos en catálogo para el correo: ${email}`);

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
        
        // 2. LOG DE ÉXITO: Confirmamos que encontramos sus credenciales
        console.log(`[INFO USUARIO - ÉXITO] Datos encontrados para: ${email} (Identificador USR: ${data[i][idxUSR]})`);
        
        return {
          nombreCompleto: data[i][idxNombre],
          vpl_usr: data[i][idxIDVPL] + " - " + data[i][idxUSR],
          noEmpleado: data[i][idxNoEmpleado],
          usrCode: data[i][idxUSR]
        };
      }
    }
    
    // 3. LOG DE ADVERTENCIA: El ciclo terminó y el correo no estaba en el archivo
    // Esto es vital para saber si se te olvidó registrar a un analista nuevo
    console.warn(`[INFO USUARIO - NO ENCONTRADO] El correo ${email} no está registrado en la hoja de equipo. Se enviarán datos "N/A".`);
    
  } catch (e) {
    // 4. LOG DE ERROR CRÍTICO: Atrapamos fallos de Google Sheets o permisos
    console.error(`[INFO USUARIO - ERROR] Fallo técnico al buscar la información de ${email}. Detalle: ${e.toString()}`);
  }
  
  // Fallback (Valores por defecto si no lo encuentra o hay error)
  return { nombreCompleto: email, vpl_usr: "N/A", noEmpleado: "N/A", usrCode: "USR" };
}

/**
 * Obtiene el catálogo de clasificaciones e indicaciones predefinidas.
 * Lee directamente de la pestaña de configuración para alimentar los desplegables de atención.
 * @returns {Array<Object>} Lista de objetos con el formato { clasificacion: string, indicacion: string }.
 */
function obtenerCatalogoNotas() {
  // 1. LOG DE INTENTO (Útil para monitorear si se consulta muchas veces)
  console.log(`[CATÁLOGO - INTENTO] Solicitando lista de clasificaciones predefinidas...`);

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(CATALOG_SHEET_NAME);
    
    // Validación: ¿Existe la pestaña y tiene datos más allá del encabezado?
    if (!sheet || sheet.getLastRow() < 2) {
      // 2. LOG DE ADVERTENCIA: Te avisa si tu pestaña quedó vacía
      console.warn(`[CATÁLOGO - VACÍO] La pestaña '${CATALOG_SHEET_NAME}' no existe o está vacía. El desplegable estará en blanco.`);
      return [];
    }

    // Extraemos solo los datos reales (omitimos fila 1, leemos 2 columnas)
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    
    // 3. LOG DE ÉXITO: Confirmamos cuántas opciones se cargaron
    console.log(`[CATÁLOGO - ÉXITO] Se cargaron ${data.length} opciones de clasificación exitosamente.`);
    
    return data.map(row => ({ clasificacion: row[0], indicacion: row[1] }));

  } catch (e) {
    // 4. LOG DE ERROR CRÍTICO: Atrapamos fallos técnicos de conexión
    console.error(`[CATÁLOGO - ERROR] Fallo al leer la hoja de configuraciones '${CATALOG_SHEET_NAME}'. Detalle: ${e.toString()}`);
    
    // Devolvemos un arreglo vacío como "salvavidas" para que la app no colapse
    return []; 
  }
}

/**
 * Convierte minutos decimales a formato de tiempo legible (H:MM:SS o M:SS).
 * Se utiliza para mostrar los promedios y SLA en el Dashboard.
 * @param {number} minutos Cantidad de minutos en formato decimal o entero.
 * @returns {string} Cadena de texto con el tiempo formateado.
 */
function formatearTiempo(minutos) {
  // 1. DEFENSA Y AUDITORÍA: Detectamos anomalías matemáticas
  if (!minutos || minutos < 0 || isNaN(minutos)) {
    
    // Solo lanzamos el Warning si el dato está genuinamente roto (negativo o no-numérico)
    // Evitamos loguear si simplemente viene en 0
    if (minutos < 0 || isNaN(minutos)) {
      console.warn(`[TIEMPO - ANOMALÍA] Se recibió un valor de tiempo inválido o negativo: ${minutos}. Se regresará '0:00'.`);
    }
    
    return "0:00";
  }
  
  // 2. Ejecución matemática normal
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
 * Normaliza valores de resolución (texto de reloj o decimal de día) a minutos para cálculos de SLA.
 * @param {string|number} valor El tiempo de resolución extraído directamente de la celda.
 * @returns {number} El equivalente en minutos totales decimales.
 */
function normalizarResolucionAMinutos(valor) {
  if (!valor || valor === "" || valor === "0") return 0;
  
  try {
    var str = valor.toString().trim();
    
    // --- RUTA A: Si ya viene en formato HH:MM:SS ---
    if (str.indexOf(':') > -1) {
      var partes = str.split(':');
      var h = parseInt(partes[0]) || 0;
      var m = parseInt(partes[1]) || 0;
      var s = parseInt(partes[2]) || 0;
      
      if (h < 0 || m < 0 || s < 0) {
        console.warn(`[NORMALIZAR - ANOMALÍA RELOJ] Se detectó tiempo negativo: ${str}. Se asumirá 0.`);
        return 0;
      }
      
      return (h * 60) + m + (s / 60);
    }
    
    // --- RUTA B: Si viene como decimal de Excel (fracción de día) ---
    var num = parseFloat(str.replace(',', '.'));
    
    if (isNaN(num)) {
      console.warn(`[NORMALIZAR - DATO INVÁLIDO] No se pudo interpretar el valor como tiempo: '${valor}'. Se devolverá 0.`);
      return 0;
    }
    
    // ¡Aquí está tu ajuste a 2 días!
    if (num > 2) {
      console.warn(`[NORMALIZAR - ALERTA SLA] Atención inusualmente larga detectada (${(num).toFixed(2)} días) en el valor original: ${valor}.`);
    }
    
    return num * 1440;
    
  } catch (e) {
    console.error(`[NORMALIZAR - ERROR CRÍTICO] Fallo al procesar el valor: '${valor}'. Detalle: ${e.toString()}`);
    return 0;
  }
}

/**
 * Genera la interfaz de acceso denegado con el estilo visual de la marca.
 * Además, registra el intento de acceso no autorizado en la bitácora de seguridad.
 * @param {string} email El correo del usuario que intentó acceder.
 * @returns {HtmlOutput} La página HTML renderizada con el bloqueo.
 */
function crearPantallaBloqueo(email) {
  // 1. LOG DE SEGURIDAD: Auditoría de intrusiones (Nivel Warning)
  console.warn(`[SEGURIDAD - ACCESO DENEGADO] El usuario ${email || "DESCONOCIDO"} intentó acceder a Sentinel sin autorización.`);

  try {
    // 2. DEFENSA DE VARIABLE: Prevenimos que el split() rompa el código si el email viene corrupto
    var emailSeguro = (typeof email === 'string' && email.includes('@')) ? email : "usuario@desconocido.com";
    var nombreUser = emailSeguro.split('@')[0].toUpperCase();
    
    var html = `
    <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background-color: #fdf4fa; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;
          }
          .card { background-color: #ffffff; border-top: 5px solid #D40099; padding: 40px; border-radius: 12px;
          max-width: 480px; text-align: center; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
          h1 { color: #78256F;
          margin-top: 0; font-size: 24px; font-weight: 800; }
          p { color: #555; font-size: 15px;
          line-height: 1.6; }
          .user-name { color: #D40099; font-weight: 800;
          }
          .footer { margin-top: 30px; font-size: 11px; color: #aaa; border-top: 1px solid #f1f1f1;
          padding-top: 20px; }
        </style>
      </head>
      <body>
        <div class="card">
          <h1>Acceso Restringido</h1>
          <p>Hola <span class="user-name">${nombreUser}</span>,</p>
          <p>Actualmente tu usuario no cuenta con los permisos necesarios para acceder a esta aplicación.</p>
          <p>Por favor, contacta a <b>hmoralesa@liverpool.com.mx</b> para gestionar tu autorización correspondiente.</p>
          <div class="footer">ID de cuenta detectado: ${emailSeguro}</div>
        </div>
      </body>
    </html>`;
    
    return HtmlService.createHtmlOutput(html).setTitle("Acceso Denegado");

  } catch (e) {
    // 3. LOG DE ERROR CRÍTICO: Por si falla el motor de plantillas HTML
    console.error(`[SEGURIDAD - ERROR VISUAL] Fallo al generar la pantalla de bloqueo para ${email}. Detalle: ${e.toString()}`);
    
    // Fallback de emergencia ultrabásico
    return HtmlService.createHtmlOutput("<h1>Acceso Denegado</h1><p>Contacte a hmoralesa@liverpool.com.mx</p>").setTitle("Bloqueado");
  }
}

/**
 * Genera una pantalla de error con la identidad visual de Sentinel.
 * Actúa como el último recurso visual si el sistema falla de manera crítica.
 * @param {Error|string} error El objeto o mensaje de error capturado.
 * @param {string} email El correo del usuario afectado.
 * @returns {HtmlOutput} La página HTML renderizada con el mensaje de fallo.
 */
function crearPantallaError(error, email) {
  // 1. DEFENSA DE VARIABLES: Prevenimos que la pantalla de error colapse
  var emailSeguro = (typeof email === 'string' && email.includes('@')) ? email : "usuario@desconocido.com";
  var nombreUser = emailSeguro.split('@')[0].toUpperCase();
  
  // Si el error viene nulo, le ponemos un texto genérico para que no explote el .toString()
  var mensajeError = (error && error.toString) ? error.toString() : "Error desconocido del sistema.";

  // 2. LOG CRÍTICO CENTRALIZADO: Registramos que un usuario vio la "Pantalla de la Muerte"
  console.error(`[SISTEMA - PANTALLA ERROR] Mostrando pantalla de fallo crítico a ${emailSeguro}. Causa: ${mensajeError}`);

  try {
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
          ${mensajeError}
        </div>

        <p>Por favor, envía una captura de esta pantalla a <strong>hmoralesa@liverpool.com.mx</strong>.</p>
        
        <hr style="border: 0; border-top: 1px solid #E4D3E2; margin: 30px 0;">
        <small style="color: #BF97BA;">Cuenta detectada: ${emailSeguro}</small>
      </div>
    `;

    return HtmlService.createHtmlOutput(html)
      .setTitle("Error - Sentinel Fraudes")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (e) {
    // 3. FALLBACK DE EMERGENCIA: Si incluso la generación de HTML falla
    console.error(`[SISTEMA - ERROR FATAL] No se pudo ni siquiera generar la pantalla de error. Detalle: ${e.toString()}`);
    return HtmlService.createHtmlOutput("<h1>Error Fatal</h1><p>El sistema colapsó por completo. Contacte a hmoralesa@liverpool.com.mx de inmediato.</p>");
  }
}
