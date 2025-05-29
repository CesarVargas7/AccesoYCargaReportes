function importarExcelDesdeOneDrive(url, nombreArchivo) {
  const nombreCarpeta = 'Sheets Reportes';

  try {
    const carpeta = asegurarCarpeta(nombreCarpeta);
    const blob = obtenerBlobDesdeOneDrive(url).setName(nombreArchivo);
    carpeta.createFile(blob); // opcional: guardar copia original

    const idConvertido = convertirBlobAGoogleSheet(blob, nombreArchivo, carpeta);

    Logger.log('✅ Google Sheet creado: https://docs.google.com/spreadsheets/d/' + idConvertido);
    return idConvertido;

  } catch (error) {
    Logger.log('❌ Error OneDrive: ' + error.toString());
    return false;
  }
}
/* function importarExcelDesdeOneDrive(url, nombreArchivo) {
  try {
    const carpeta = asegurarCarpeta('Sheets Reportes');
    const nuevoId = crearSpreadsheet(nombreArchivo, carpeta);
    copiarDesdeOneDrive(url, nuevoId);
    Logger.log('✅ Spreadsheet creado desde OneDrive: ' + nuevoId);
    return nuevoId;
  } catch (e) {
    Logger.log('❌ Error al importar desde OneDrive: ' + e);
    return false;
  }
} */

function copiarArchivoDeDrive(idOrigen) {
  try {
    const carpeta = asegurarCarpeta('Sheets Reportes');
    const nombre = DriveApp.getFileById(idOrigen).getName();
    const nuevoId = crearSpreadsheet(nombre, carpeta);
    copiarDesdeDrive(idOrigen, nuevoId);
    Logger.log('✅ Copiado desde Drive: ' + nuevoId);
    return nuevoId;
  } catch (e) {
    Logger.log('❌ Error al copiar desde Drive: ' + e);
    return false;
  }
}

function copiarDesdeOneDrive(id) {
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
/* function copiarDesdeOneDrive(url, idDestino) {
  const respuesta = UrlFetchApp.fetch(url, { followRedirects: true });
  const blob = respuesta.getBlob().setName('temp.xlsx');
  const archivoTemp = Drive.Files.insert({ title: 'temp', mimeType: MimeType.GOOGLE_SHEETS }, blob);
  const libroTemporal = SpreadsheetApp.openById(archivoTemp.id);
  const libroDestino = SpreadsheetApp.openById(idDestino);

  // Copiar cada hoja
  libroTemporal.getSheets().forEach(hoja => {
    const nuevaHoja = hoja.copyTo(libroDestino);
    nuevaHoja.setName(hoja.getName());
  });

  // Borrar hoja por defecto y archivo temporal
  libroDestino.deleteSheet(libroDestino.getSheets()[0]);
  DriveApp.getFileById(archivoTemp.id).setTrashed(true);
} */

function copiarDesdeDrive(idOrigen, idDestino) {
  const origen = SpreadsheetApp.openById(idOrigen);
  const destino = SpreadsheetApp.openById(idDestino);

  origen.getSheets().forEach(hoja => {
    const copia = hoja.copyTo(destino);
    copia.setName(hoja.getName());
  });

  // Borrar hoja por defecto si existe solo una al inicio
  if (destino.getSheets().length > origen.getSheets().length) {
    destino.deleteSheet(destino.getSheets()[0]);
  }
}

function obtenerBlobDesdeOneDrive(url) {
  // Detectar redirección
  const redir = UrlFetchApp.fetch(url, {
    followRedirects: false,
    muteHttpExceptions: true
  });

  const status = redir.getResponseCode();
  let finalUrl = url;

  if (status === 302 || status === 301) {
    const location = redir.getAllHeaders()['Location'];
    if (!location) throw new Error('No se pudo seguir la redirección.');
    finalUrl = location;
  }

  // Descargar archivo
  const respuesta = UrlFetchApp.fetch(finalUrl, {
    followRedirects: true,
    muteHttpExceptions: false
  });

  const contentType = respuesta.getHeaders()['Content-Type'];
  if (!/excel|spreadsheet/.test(contentType)) {
    throw new Error('El archivo descargado no es un Excel válido.');
  }

  return respuesta.getBlob();
}


function convertirBlobAGoogleSheet(blob, nombreArchivo, carpetaDestino) {
  return Drive.Files.insert(
    {
      title: nombreArchivo,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: carpetaDestino.getId() }]
    },
    blob
  ).id;
}
