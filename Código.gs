function doPost(e) {
  try {
    // Recibimos la llamada dede PrivilegiosReporteController
    const data = JSON.parse(e.postData.contents);
    const { email, accion, url, id, nombreArchivo } = data;

    switch(accion) {
      case "Crear": 
        return crearArchivo(email, url, nombreArchivo);

      case "Actualizar":
        return actualizarArchivo(url, id);

      case "Cargar":
        return cargarArchivo(id);

      case "Editar":
        return permisoEdicion(email, id);

      case "Lectura":
        return permisoLectura(email, id);
    }
    
    // Asignación de permiso al archivo
    const result = assignReadOnlyPermission(email, archivoId);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
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

// Se lleva un registro con el URL enviado y la copia creada
function registrarImportacionEnSheets(url, idDocCreado) {
  const colDocOriginal = 'Archivo Original';
  const colDocCreado = 'Archivo Creado';

  const sheet = getSheet();

  // Obtiene las columnas
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Las columnas que deben existir
  const columnas = [colDocOriginal, colDocCreado];
  columnas.forEach(col => {
    // Asegurar que las columnas existan
    if (!headers.includes(col)) headers.push(col);
  });
  // Las ingresa en el sheet
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Actualizar encabezados si se agregaron columnas
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Buscar la siguiente fila vacía
  const lastRow = sheet.getLastRow() + 1;
  // Obtiene el índice de las columnas
  const col1 = headers.indexOf(colDocOriginal) + 1;
  const col2 = headers.indexOf(colDocCreado) + 1;

  // Escribir el enlace original en la hoja
  sheet.getRange(lastRow, col1).setValue(url);

  // Escribir el id de la copia en la hoja
  sheet.getRange(lastRow, col2).setValue(idDocCreado);
}

function buscarIDCreado(idDocCreado) {
  const sheet = getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = headers.indexOf('Archivo Creado') + 1;
  if (!col) return -1;

  const datos = sheet.getRange(2, col, sheet.getLastRow() - 1).getValues();
  return datos.findIndex(r => r[0]?.includes(idDocCreado)) + 2 || -1;
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
  DriveApp.getRootFolder().removeFile(archivoFile); // Opcional: mover solo a carpeta destino
  return archivoId;
}

