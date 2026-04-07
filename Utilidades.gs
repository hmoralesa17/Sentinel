/**
 * Módulo de Utilidades V2.0
 * Funciones de apoyo y soporte para el sistema (Optimizadas con Globales)
 */


/**
 * Obtiene la información detallada del usuario desde la hoja de equipo.
 * Si no encuentra al usuario o falla la conexión, devuelve un objeto genérico (Fallback).
 * @returns {Object} Objeto con nombreCompleto, vpl_usr, noEmpleado y usrCode.
 */
function obtenerInfoUsuario() {
  // 1. MAGIA GLOBAL: Usamos la variable que se cargó al inicio del sistema
  var email = GLOBAL_EMAIL; 
  
  console.log(`[INFO USUARIO - INTENTO] Buscando datos en catálogo para el correo: ${email}`);

  try {
    var ss = SpreadsheetApp.openById(SHEET_EQUIPO_ID);
    var sheet = ss.getSheetByName(SHEET_EQUIPO_NAME);
    
    // Como esta es una tabla minúscula (solo el equipo), getValues e indexOf son instantáneos
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var idxCorreo = headers.indexOf("Correo"), 
        idxNombre = headers.indexOf("Nombre Completo"), 
        idxIDVPL = headers.indexOf("ID_VPL"), 
        idxUSR = headers.indexOf("USR"), 
        idxNoEmpleado = headers.indexOf("No_Empleado");
        
    for (var i = 1; i < data.length; i++) {
      if (data[i][idxCorreo] && data[i][idxCorreo].toString().toLowerCase().trim() === email) {
        
        console.log(`[INFO USUARIO - ÉXITO] Datos encontrados para: ${email} (Identificador USR: ${data[i][idxUSR]})`);
        
        return {
          nombreCompleto: data[i][idxNombre],
          vpl_usr: data[i][idxIDVPL] + " - " + data[i][idxUSR],
          noEmpleado: data[i][idxNoEmpleado],
          usrCode: data[i][idxUSR]
        };
      }
    }
    
    console.warn(`[INFO USUARIO - NO ENCONTRADO] El correo ${email} no está en la hoja de equipo. Se enviarán datos "N/A".`);
    
  } catch (e) {
    console.error(`[INFO USUARIO - ERROR] Fallo al buscar información de ${email}. Detalle: ${e.toString()}`);
  }
  
  // Fallback (Valores por defecto si no lo encuentra o hay error)
  return { nombreCompleto: email, vpl_usr: "N/A", noEmpleado: "N/A", usrCode: "USR" };
}

/**
 * Obtiene el catálogo de clasificaciones e indicaciones predefinidas.
 * Lee directamente de la pestaña de configuración para alimentar los desplegables.
 * @returns {Array<Object>} Lista de objetos con el formato { clasificacion: string, indicacion: string }.
 */
function obtenerCatalogoNotas() {
  console.log(`[CATÁLOGO - INTENTO] Solicitando lista de clasificaciones predefinidas...`);

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(CATALOG_SHEET_NAME);
    
    if (!sheet) {
      console.warn(`[CATÁLOGO - VACÍO] La pestaña '${CATALOG_SHEET_NAME}' no existe. El desplegable estará en blanco.`);
      return [];
    }

    // ⚡ OPTIMIZACIÓN: Le preguntamos a Google la última fila UNA sola vez
    var ultimaFila = sheet.getLastRow();

    if (ultimaFila < 2) {
      console.warn(`[CATÁLOGO - VACÍO] La pestaña '${CATALOG_SHEET_NAME}' está vacía. El desplegable estará en blanco.`);
      return [];
    }

    // Usamos nuestra variable en memoria
    var data = sheet.getRange(2, 1, ultimaFila - 1, 2).getValues();
    
    console.log(`[CATÁLOGO - ÉXITO] Se cargaron ${data.length} opciones de clasificación exitosamente.`);
    
    return data.map(function(row) { 
      return { clasificacion: row[0], indicacion: row[1] }; 
    });

  } catch (e) {
    console.error(`[CATÁLOGO - ERROR] Fallo al leer la hoja '${CATALOG_SHEET_NAME}'. Detalle: ${e.toString()}`);
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
          <p>Por favor, contacta a <b>${CEREBRO.soporte}</b> para gestionar tu autorización correspondiente.</p>
          <div class="footer">ID de cuenta detectado: ${emailSeguro}</div>
        </div>
      </body>
    </html>`;
    
    return HtmlService.createHtmlOutput(html).setTitle("Acceso Denegado");

  } catch (e) {
    // 3. LOG DE ERROR CRÍTICO: Por si falla el motor de plantillas HTML
    console.error(`[SEGURIDAD - ERROR VISUAL] Fallo al generar la pantalla de bloqueo para ${email}. Detalle: ${e.toString()}`);
    
    // Fallback de emergencia ultrabásico (También conectamos la variable aquí sumando el texto)
    return HtmlService.createHtmlOutput("<h1>Acceso Denegado</h1><p>Contacte a " + CEREBRO.soporte + "</p>").setTitle("Bloqueado");
  }
}


/**
 * Genera una pantalla de error con la identidad visual de la app.
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
          El sistema encontró un problema inesperado y no puede iniciar.
        </p>

        <div style="background-color: #f4f6f9; border-left: 5px solid #e74c3c; padding: 15px; margin: 20px auto; max-width: 600px; text-align: left; color: #333; font-family: monospace;">
          <strong>Detalle Técnico:</strong><br>
          ${mensajeError}
        </div>

        <p>Por favor, envía una captura de esta pantalla a <strong>${CEREBRO.soporte}</strong>.</p>
        
        <hr style="border: 0; border-top: 1px solid #E4D3E2; margin: 30px 0;">
        <small style="color: #BF97BA;">Cuenta detectada: ${emailSeguro}</small>
      </div>
    `;

    return HtmlService.createHtmlOutput(html)
      .setTitle("Error - " + CEREBRO.appTitle) // 👇 MAGIA: Título dinámico desde tu Mesa de Control
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (e) {
    // 3. FALLBACK DE EMERGENCIA: Si incluso la generación de HTML falla
    console.error(`[SISTEMA - ERROR FATAL] No se pudo ni siquiera generar la pantalla de error. Detalle: ${e.toString()}`);
    
    // 👇 MAGIA: Correo dinámico en el fallback de emergencia
    return HtmlService.createHtmlOutput("<h1>Error Fatal</h1><p>El sistema colapsó por completo. Contacte a " + CEREBRO.soporte + " de inmediato.</p>");
  }
}


/**
 * =======================================================
 * MÓDULO EXPLORADOR: EXTRAER CATÁLOGOS DINÁMICOS
 * (Versión Optimizada con Globales y Escudo de Bot)
 * =======================================================
 */
function obtenerCatalogosExplorador() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    // 🛡️ EL ESCUDO DEL ROBOT
    var ultimaFilaReal = IDX_FILA_ULTIMOFOLIO;
    if (ultimaFilaReal < 2) {
      ultimaFilaReal = sheet.getLastRow(); // Fallback por seguridad
    }
    
    // Si no hay datos, devolvemos listas vacías para que no truene el Front
    if (ultimaFilaReal < 2) {
      return { exito: true, tipoId: [], tipoCaso: [], clasificacion: [] };
    }
    
    // Descargamos SOLO el bloque de datos real (desde la fila 2)
    var data = sheet.getRange(2, 1, ultimaFilaReal - 1, TODAS_LAS_COLUMNAS.length).getValues();
    
    // Usamos objetos como "Sets" para eliminar duplicados fácilmente
    var setTipoId = {};
    var setTipoCaso = {};
    var setClasif = {};
    
    // Recorremos la base de datos volando con nuestras globales
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // Validamos que la columna exista (> -1) y que la celda no esté vacía
      if (IDX_TIPO_ID > -1 && row[IDX_TIPO_ID]) {
        setTipoId[row[IDX_TIPO_ID].toString().trim()] = true;
      }
      
      if (IDX_TIPO_CASO > -1 && row[IDX_TIPO_CASO]) {
        setTipoCaso[row[IDX_TIPO_CASO].toString().trim()] = true;
      }
      
      if (IDX_CLASIF > -1 && row[IDX_CLASIF]) {
        setClasif[row[IDX_CLASIF].toString().trim()] = true;
      }
    }
    
    return { 
      exito: true, 
      tipoId: Object.keys(setTipoId).sort(),
      tipoCaso: Object.keys(setTipoCaso).sort(),
      clasificacion: Object.keys(setClasif).sort()
    };
    
  } catch (e) {
    console.error(`[CATÁLOGOS EXPLORADOR - ERROR] Fallo al extraer opciones: ${e.toString()}`);
    return { exito: false, mensaje: e.toString() };
  }
}
