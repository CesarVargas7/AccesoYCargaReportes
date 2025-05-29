function crearArchivo(email, url, nombreArchivo) {
  let idDocCreado;

  // Valida si se recibe un Excel
  if(url.includes("1drv.ms")) {
    // Hace la conversión de Excel a Google Sheets
    idDocCreado = importarExcelDesdeOneDrive(url, nombreArchivo);
  }
  // Si no, es un Drive
  else {
    idDocCreado = copiarArchivoDeDrive(url)
  }
  // Valida que si regresó un error (false).
  if(!idDocCreado) { 
    return "Error al importar de Excel a Google Sheets."
  }

  registrarImportacionEnSheets(url, idDocCreado)
  permisoEdicion(email, idDocCreado)
}

// Además de volver a copiar el archivo original al creado,
// sobreescribe el URL del archivo fuente
function cargarArchivo(id) {
  const fila = buscarIDCreado(id);
  const sheet = SpreadsheetApp.openById("1NOQyzaGSHXWFcfHd7WjVnTk_b_zwrt2UKW1_xo8v_58").getSheets()[0];
  const url = sheet.getRange(fila, obtenerColumna("Archivo Original")).getValue();

  if (/onedrive|1drv\.ms/.test(url)) {
    const blob = obtenerBlobDesdeOneDrive(url);
    copiarContenidoDesdeBlob(blob, id); // la puedes crear si quieres una función específica para esto
  } else {
    const idFuente = extraerIDDrive(url);
    copiarDesdeDrive(idFuente, id);
  }
}


/* function actualizarArchivo(url, idDocCreado) {
  const fila = buscarIDCreado(idDocCreado);
  if (fila === -1) return Logger.log('❌ No se encontró el ID en la hoja.');

  const sheet = getSheet();
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf('Archivo Original') + 1;

  // Agrega columna si no existe
  if (!col) {
    headers.push('Archivo Original');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    col = headers.length;
  }

  // Escribe el nuevo URL y recarga el archivo
  sheet.getRange(fila, col).setValue(url);
  cargarArchivo(idDocCreado);
} */

// Solamente vuelve a copiar el archivo original al creado
function cargarArchivo(idCreado) {
  const fila = buscarIDCreado(idCreado);
  if (fila === -1) return Logger.log('❌ No se encontró el ID en la hoja.');

  const sheet = getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colOriginal = headers.indexOf('Archivo Original') + 1;

  if (!colOriginal) return Logger.log('❌ No existe la columna "Archivo Original".');

  const url = sheet.getRange(fila, colOriginal).getValue().toString().trim();

  if (!url) return Logger.log('⚠️ No hay URL en "Archivo Original"');

  // Detecta si es un enlace de OneDrive
  if (/1drv\.ms|onedrive\.live/.test(url)) {
    copiarDesdeOneDrive(url, idCreado);
  } else {
    const match = url.match(/[-\w]{25,}/); // Extraer ID de Drive
    if (match) {
      copiarDesdeDrive(match[0], idCreado);
    } else {
      Logger.log('❌ No se pudo extraer ID válido de Drive');
    }
  }
}

// Asignación y remoción de permiso de lectura al archivo
function permisoLectura(email, archivoId) {
  try {
    // Obtener el archivo por ID
    const file = DriveApp.getFileById(archivoId);
    
    let viewers = file.getViewers();
    // Si encuentra el lector, entonces la llamada es de remoción
    if(viewers.includes(email)) {
      file.removeViewer(email);
      return `Permiso de lectura removido a ${email}`;
    }

    // Si no lo encuentra, la llamada es de asignación
    // Asigna el permiso de solo lectura sin enviar notificación
    file.addViewer(email);
    
    return `Permiso de lectura asignado a ${email}`;
    
  } catch (error) {
    return `Error al asignar permiso: ${error.toString()}`;
  }
}

// Asignación y remoción de permiso de edición al archivo
function permisoEdicion(email, archivoId) {
  try {
    // Obtener el archivo por ID
    const file = DriveApp.getFileById(archivoId);
    
    let editors = file.getEditors();
    // Si encuentra el lector, entonces la llamada es de remoción
    if(editors.includes(email)) {
      file.removeEditor(email);
      return `Permiso de edición removido a ${email}`;
    }

    // Si no lo encuentra, la llamada es de asignación
    // Asignar permiso de solo lectura sin enviar notificación
    file.addEditor(email);
    // Asegura que a quienes se les da acceso el archivo como editor, no puedan compartir
    file.setShareableByEditors(false); 
    
    return `Permiso de edición asignado a ${email}`;
    
  } catch (error) {
    return `Error al asignar permiso: ${error.toString()}`;
  }
}