/**
 * Módulo Variables Globales V1.0
 * Archivo de configración de variables que se ocupan en la pagina, tanto las que vienen de mesa como las directas del código.
 */


// =======================================================
// 1. EL CEREBRO (Conexión a la Bóveda de Configuración)
// =======================================================
var CEREBRO = leerConfiguracionGeneral();

// =======================================================
// 2. VARIABLES DINÁMICAS (Controladas desde la Interfaz)
// =======================================================
var CATALOG_SHEET_NAME = CEREBRO.catalogName;
var VERSION_APP = CEREBRO.appVersion;
var SHEET_ID = CEREBRO.sheetId;
var SHEET_NAME = CEREBRO.sheetName;
var BOT_SHEET_NAME = CEREBRO.botSheetName || 'Config BOT';
var CORREO_SOPORTE = CEREBRO.soporte;

// Extraemos los admins de la lista separada por comas y los limpiamos
var ADMINS = CEREBRO.admins.split(',').map(function(a) { return a.trim().toLowerCase(); });

// ¡MAGIA!: Generamos la lista de columnas leyendo directamente tu Diccionario JSON
var DICCIONARIO_JSON = [];
try { DICCIONARIO_JSON = JSON.parse(CEREBRO.diccionario); } catch(e) {}
var TODAS_LAS_COLUMNAS = DICCIONARIO_JSON.map(function(col) { return col.bd; });

// =======================================================
// 3. VARIABLES FIJAS (Usuarios y Equipo)
// =======================================================
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

// =======================================================
// 4. CONTEXTO DE SESIÓN (Variables Globales de Entorno)
// =======================================================
// Se calculan una sola vez al despertar el servidor, ahorrando tiempo en cada función.
var GLOBAL_EMAIL = Session.getActiveUser().getEmail().toLowerCase();
var GLOBAL_TIMEZONE = Session.getScriptTimeZone();

// =======================================================
// 5. VARIABLES DEL ROBOT (Diccionario Global)
// =======================================================
var VARIABLES_BOT = leerConfiguracionBot();

// =======================================================
// 6. MAPA DE COLUMNAS (Índices Globales Listos para Usar)
// =======================================================

// Mini-función para evitar el "bug del cero" en JavaScript
function getIdx(nombreBot) {
  var numero = parseInt(VARIABLES_BOT[nombreBot]);
  return isNaN(numero) ? -1 : numero; // Si no es un número, regresa -1. Si es 0 o más, lo respeta.
}

var IDX_USR = getIdx("USR"); 
var IDX_CLASIF = getIdx("Clasificación");
var IDX_ATENCION = getIdx("Atención");
var IDX_SLA = getIdx("Resolución");
var IDX_CURP = getIdx("CURP");
var IDX_CUENTA = getIdx("CTA"); 
var IDX_REPETIDO = getIdx("Repetido");
var IDX_MARCA = getIdx("MarcaRecibido");
var IDX_CORREO = getIdx("Correo");
var IDX_TIPO_CLI = getIdx("TipoCliente");
var IDX_TIPO_CASO = getIdx("TipoCaso");
var IDX_SERVICIO = getIdx("Servicio");
var IDX_COMENTARIOS = getIdx("NotasBot");
var IDX_TIENDA = getIdx("TDA"); 
var IDX_TIPO_ID = getIdx("TipoId");
var IDX_INDICACIONES = getIdx("Indicaciones");
var IDX_FOLIO = getIdx("Folio");
var IDX_INICIO_ATN = getIdx("InicioAtencion");
var IDX_FILA_FOLIO_PENDIENTE = getIdx("FilaFolioPendiente");
var IDX_FILA_ULTIMOFOLIO = getIdx("FilaUltimoFolio");
var IDX_FILA_PRIMER_FOLIO_HOY = getIdx("FilaPrimerFolioHoy");
var IDX_FILA_PRIMER_FOLIO_XDIAS = getIdx("FilaPrimerFolioxDias");
