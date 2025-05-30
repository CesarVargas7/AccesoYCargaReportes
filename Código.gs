function doPost(e) {
  try { 
    const data = JSON.parse(e.postData.contents);
    const { email, accion, url, id, nombreArchivo } = data;

    let result;
    switch(accion) {
      case "Crear": 
        result = crearArchivo(email, url, nombreArchivo);
        break;

      case "Actualizar":
        result = actualizarArchivo(url, id);
        break;

      case "Cargar":
        result = cargarArchivo(id);
        break;

      case "Editar":
        result = permisoEdicion(email, id);
        break;

      case "Lectura":
        result = permisoLectura(email, id);
        break;

      default:
        result = { success: false, error: "Acción no válida" };
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      result: result
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function test() {
  const payload = {
    email: "ses.tono@gmail.com",
    accion: "Crear",
    url: "https://1drv.ms/x/c/70363ac35e9e3256/EYcGFvMjfaZPjEpTXjnG080BVfBHPGAnsL8WhQ8jeZABgA?e=RlD1PB",
    id:"",
    nombreArchivo: "Presupuesto mensual"
  };

  const e = {
    postData: {
      contents: JSON.stringify(payload)
    }
  };

  const respuesta = doPost(e);
  Logger.log(respuesta);
}

///////// Funciones Complementarias /////////

function getSheet() {
  return SpreadsheetApp.openById('1NOQyzaGSHXWFcfHd7WjVnTk_b_zwrt2UKW1_xo8v_58').getSheetByName('RegistroDocs');
}

function registrarImportacionEnSheets(url, idDocCreado) {
  const colDocOriginal = 'Archivo Original';
  const colDocCreado = 'Archivo Creado';
  const sheet = getSheet();

  // Verificar si la hoja tiene columnas
  const lastColumn = sheet.getLastColumn();
  let headers = [];
  
  if (lastColumn > 0) {
    headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }
  
  const columnas = [colDocOriginal, colDocCreado];
  
  columnas.forEach(col => {
    if (!headers.includes(col)) headers.push(col);
  });
  
  // Escribir headers solo si hay contenido
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sheet.getLastRow() + 1;
  const col1 = headers.indexOf(colDocOriginal) + 1;
  const col2 = headers.indexOf(colDocCreado) + 1;

  sheet.getRange(lastRow, col1).setValue(url);
  sheet.getRange(lastRow, col2).setValue(idDocCreado);
}

function buscarIDCreado(idDocCreado) {
  const sheet = getSheet();
  const lastColumn = sheet.getLastColumn();
  
  if (lastColumn === 0) return -1;
  
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const col = headers.indexOf('Archivo Creado') + 1;
  if (!col) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1; // Solo headers
  
  const datos = sheet.getRange(2, col, lastRow - 1).getValues();
  return datos.findIndex(r => r[0]?.toString().includes(idDocCreado)) + 2 || -1;
}

function obtenerColumna(nombreColumna) {
  const sheet = getSheet();
  const lastColumn = sheet.getLastColumn();
  
  if (lastColumn === 0) return 0;
  
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  return headers.indexOf(nombreColumna) + 1;
}

function asegurarCarpeta(nombreCarpeta) {
  const carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  return carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(nombreCarpeta);
}

function crearSpreadsheet(nombre, carpeta) {
  const archivo = SpreadsheetApp.create(nombre);
  const archivoId = archivo.getId();
  const archivoFile = DriveApp.getFileById(archivoId);
  carpeta.addFile(archivoFile);
  DriveApp.getRootFolder().removeFile(archivoFile);
  return archivoId;
}

function extraerIDDrive(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}