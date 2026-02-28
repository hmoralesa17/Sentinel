Session.getActiveUser().getEmail();
/**
 * Módulo Core y Configuración V1.0
 * Archivo Principal de Configuración y Control de Acceso
 */

var SHEET_ID = '1h9zkrvgmH0r-K1wokCVgo3c_aGEKbqBMH03VHL6flDo';
var SHEET_NAME = 'Respuestas de formulario 2';
var CATALOG_SHEET_NAME = 'Notas Bot'; 
var SHEET_EQUIPO_ID = '1JhuWplwWV8k2yj2qI9g0j0niCJlkmnnzH50II1V8aQ0';
var SHEET_EQUIPO_NAME = 'Datos_Equipo';

var DICCIONARIO_USUARIOS = {
  "ajmejian@liverpool.com.mx": "AMN",
  "eurodriguezs@liverpool.com.mx": "GRZ",
  "hmoralesa@liverpool.com.mx": "HOA",
  "jmtenorior@liverpool.com.mx": "UTR",
  "mparedesm@liverpool.com.mx": "MPI",
  "mhuertasa@liverpool.com.mx": "YHA",
  "mabalderasv@liverpool.com.mx": "LBR",
  "nchilpac@liverpool.com.mx": "NPC"
};

var USUARIOS_AUTORIZADOS = Object.keys(DICCIONARIO_USUARIOS);

var TODAS_LAS_COLUMNAS = [
  "FOLIO", "Marca temporal", "CURP", "Cuenta", "Tipo de Identificación", 
  "Tipo Cliente", "Tipo Caso", "Servicio", "Tienda", 
  "Dirección de correo electrónico", "Repetido", "Comentarios Bot"
];

/**
 * Función de arranque: Carga la interfaz usando plantillas para permitir 
 * la inclusión de archivos CSS y JS separados.
 */
function doGet() {
  var emailUsuario = Session.getActiveUser().getEmail();

  try {
    if (verificarAcceso(emailUsuario)) {
      var template = HtmlService.createTemplateFromFile('Index');
      return template.evaluate()
        .setTitle('Sentinel Fraudes RI-FI v1.14')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } else {
      return crearPantallaBloqueo(emailUsuario); // Tu función original [1]
    }
  } catch (e) {
    console.error("Fallo Crítico doGet: " + e.toString());
    // Aquí llamamos a la nueva función visual
    return crearPantallaError(e, emailUsuario);
  }
}

/** 
 * Función verificación de acceso a la plataforma 
 */
function verificarAcceso(email) {
  // 1. Defensa: Si no hay email o no es texto, denegamos acceso sin romper el código
  if (!email || typeof email !== 'string') {
    console.warn("Intento de acceso con email inválido: " + email);
    return false;
  }
  
  // 2. Ejecución normal
  return USUARIOS_AUTORIZADOS.indexOf(email.toLowerCase()) > -1;
}

/**
 * Función auxiliar para incluir archivos HTML (CSS y JS) en la plantilla principal.
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    // Si falla, registramos el error en la consola interna
    console.warn("⚠️ Archivo no encontrado: " + filename);
    // Devolvemos un comentario HTML inofensivo para que la página siga cargando
    return "<script>console.error('Fallo al cargar componente: " + filename + "');</script>";
  }
}
