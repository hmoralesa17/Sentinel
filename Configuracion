/**
 * Módulo de Configuración Segura V2.0 (Con Diccionario y Backup Drive)
 */

function leerConfiguracionGeneral() {
  function getSafe(key, defaultVal) {
    try { return SentSec.getSecret(key); } 
    catch(e) { return defaultVal; }
  }

  return {
    filas:          getSafe('CFG_FILAS', "3000"),
    limbo:          getSafe('CFG_LIMBO', "10"),
    sla:            getSafe('CFG_SLA', "30"),
    sheetId:        getSafe('CFG_SHEET_ID', "1h9zkrvgmH0r-K1wokCVgo3c_aGEKbqBMH03VHL6flDo"),
    sheetName:      getSafe('CFG_SHEET_NAME', "Respuestas de formulario 2"),
    catalogName:    getSafe('CFG_CATALOG_NAME', "Notas"),
    backupFolderId: getSafe('CFG_BACKUP_FOLDER_ID', "1Xx0jHr3-uJxCj9boUXSkrVC_ZHe7H58q"), // <--- Tu carpeta
    admins:         getSafe('CFG_ADMINS', "eurodriguezs, hmoralesa, mhuertasa"),
    soporte:        getSafe('CFG_SOPORTE', "hmoralesa@liverpool.com.mx"),
    appTitle:       getSafe('CFG_APP_TITLE', "Hub Sentinel Fraudes RI-FI"),
    appVersion:     getSafe('CFG_APP_VERSION', "v1.5.3 - Estable"),
    mod1Title:      getSafe('CFG_MOD1_TITLE', "Módulo de Atención Folios"),
    mod1Desc:       getSafe('CFG_MOD1_DESC', "Atención de folios biométricos en tienda."),
    mod2Title:      getSafe('CFG_MOD2_TITLE', "Módulo Buscador de Folios"),
    mod2Desc:       getSafe('CFG_MOD2_DESC', "Búsqueda avanzada y auditoría de casos históricos."),
    diccionario:    getSafe('CFG_DICCIONARIO', '[]') 
  };
}

function guardarConfiguracionGeneral(datos) {
  try {
    const clean = (val) => val ? val.toString().trim() : "";

    // 1. Guardar en memoria encriptada (Propiedades)
    SentSec.setSecret('CFG_FILAS',            clean(datos.filas));
    SentSec.setSecret('CFG_LIMBO',            clean(datos.limbo));
    SentSec.setSecret('CFG_SLA',              clean(datos.sla));
    SentSec.setSecret('CFG_SHEET_ID',         clean(datos.sheetId));
    SentSec.setSecret('CFG_SHEET_NAME',       clean(datos.sheetName));
    SentSec.setSecret('CFG_CATALOG_NAME',     clean(datos.catalogName));
    SentSec.setSecret('CFG_BACKUP_FOLDER_ID', clean(datos.backupFolderId));
    SentSec.setSecret('CFG_ADMINS',           clean(datos.admins));
    SentSec.setSecret('CFG_SOPORTE',          clean(datos.soporte));
    SentSec.setSecret('CFG_APP_TITLE',        clean(datos.appTitle));
    SentSec.setSecret('CFG_APP_VERSION',      clean(datos.appVersion));
    SentSec.setSecret('CFG_MOD1_TITLE',       clean(datos.mod1Title));
    SentSec.setSecret('CFG_MOD1_DESC',        clean(datos.mod1Desc));
    SentSec.setSecret('CFG_MOD2_TITLE',       clean(datos.mod2Title));
    SentSec.setSecret('CFG_MOD2_DESC',        clean(datos.mod2Desc));
    SentSec.setSecret('CFG_DICCIONARIO',      datos.diccionario ? datos.diccionario : "[]"); 

    // 2. BACKUP MÁGICO EN GOOGLE DRIVE
    const folderId = clean(datos.backupFolderId);
    if(folderId) {
      const folder = DriveApp.getFolderById(folderId);
      const fileName = "Respaldo_Master_HubSentinel.json";
      const fileContent = JSON.stringify(datos, null, 2); // JSON formateado bonito

      // Busca si el archivo ya existe
      const files = folder.searchFiles(`title = '${fileName}' and trashed = false`);
      if (files.hasNext()) {
        const file = files.next();
        file.setContent(fileContent); // Sobrescribe y crea nueva versión en el historial de Drive
      } else {
        folder.createFile(fileName, fileContent, MimeType.PLAIN_TEXT); // Lo crea si es la primera vez
      }
    }

    return { exito: true };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}
