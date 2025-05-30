function crearArchivo(email, url, nombreArchivo) {
  let idDocCreado;

  try {
    if(url.includes("1drv.ms") || url.includes("onedrive")) {
      idDocCreado = importarExcelDesdeOneDrive(url, nombreArchivo);
    } else {
      const idOrigen = extraerIDDrive(url);
      if (!idOrigen) {
        Logger.log('❌ No se pudo extraer ID de Drive del URL: ' + url);
        return "Error: URL de Drive no válido";
      }
      idDocCreado = copiarArchivoDeDrive(idOrigen);
    }

    if(!idDocCreado) { 
      return "Error al importar archivo";
    }

    registrarImportacionEnSheets(url, idDocCreado);
    const permisoResult = permisoEdicion(email, idDocCreado);
    
    return {
      idCreado: idDocCreado,
      permiso: permisoResult,
      url: `https://docs.google.com/spreadsheets/d/${idDocCreado}`
    };

  } catch (error) {
    Logger.log('❌ Error en crearArchivo: ' + error.toString());
    return "Error al crear archivo: " + error.toString();
  }
}

function cargarArchivo(idCreado) {
  try {
    const fila = buscarIDCreado(idCreado);
    if (fila === -1) {
      Logger.log('❌ No se encontró el ID en la hoja.');
      return "Error: ID no encontrado";
    }

    const sheet = getSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colOriginal = headers.indexOf('Archivo Original') + 1;

    if (!colOriginal) {
      Logger.log('❌ No existe la columna "Archivo Original".');
      return "Error: Columna 'Archivo Original' no encontrada";
    }

    const url = sheet.getRange(fila, colOriginal).getValue().toString().trim();

    if (!url) {
      Logger.log('⚠️ No hay URL en "Archivo Original"');
      return "Error: No hay URL en 'Archivo Original'";
    }

    if (/1drv\.ms|onedrive\.live/.test(url)) {
      copiarDesdeOneDrive(url, idCreado);
    } else {
      const idOrigen = extraerIDDrive(url);
      if (idOrigen) {
        copiarDesdeDrive(idOrigen, idCreado);
      } else {
        Logger.log('❌ No se pudo extraer ID válido de Drive');
        return "Error: ID de Drive no válido";
      }
    }

    return "Archivo cargado exitosamente";

  } catch (error) {
    Logger.log('❌ Error en cargarArchivo: ' + error.toString());
    return "Error al cargar archivo: " + error.toString();
  }
}

function actualizarArchivo(url, idDocCreado) {
  try {
    const fila = buscarIDCreado(idDocCreado);
    if (fila === -1) {
      Logger.log('❌ No se encontró el ID en la hoja.');
      return "Error: ID no encontrado";
    }

    const sheet = getSheet();
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let col = headers.indexOf('Archivo Original') + 1;

    if (!col) {
      headers.push('Archivo Original');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      col = headers.length;
    }

    sheet.getRange(fila, col).setValue(url);
    return cargarArchivo(idDocCreado);

  } catch (error) {
    Logger.log('❌ Error en actualizarArchivo: ' + error.toString());
    return "Error al actualizar archivo: " + error.toString();
  }
}

function permisoLectura(email, archivoId) {
  try {
    const file = DriveApp.getFileById(archivoId);
    const viewers = file.getViewers();
    
    const tienePermiso = viewers.some(viewer => viewer.getEmail() === email);
    
    if(tienePermiso) {
      file.removeViewer(email);
      return `Permiso de lectura removido a ${email}`;
    } else {
      file.addViewer(email);
      return `Permiso de lectura asignado a ${email}`;
    }
    
  } catch (error) {
    Logger.log('❌ Error en permisoLectura: ' + error.toString());
    return `Error al asignar permiso: ${error.toString()}`;
  }
}

function permisoEdicion(email, archivoId) {
  try {
    const file = DriveApp.getFileById(archivoId);
    const editors = file.getEditors();
    
    const tienePermiso = editors.some(editor => editor.getEmail() === email);
    
    if(tienePermiso) {
      file.removeEditor(email);
      return `Permiso de edición removido a ${email}`;
    } else {
      file.addEditor(email);
      file.setShareableByEditors(false); 
      return `Permiso de edición asignado a ${email}`;
    }
    
  } catch (error) {
    Logger.log('❌ Error en permisoEdicion: ' + error.toString());
    return `Error al asignar permiso: ${error.toString()}`;
  }
}
