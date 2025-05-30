function importarExcelDesdeOneDrive(url, nombreArchivo) {
  const nombreCarpeta = 'Sheets Reportes';

  try {
    const carpeta = asegurarCarpeta(nombreCarpeta);
    const blob = obtenerBlobDesdeOneDrive(url).setName(nombreArchivo);
    
    // Crear archivo Excel como respaldo (opcional)
    // carpeta.createFile(blob);

    const idConvertido = convertirBlobAGoogleSheet(blob, nombreArchivo, carpeta);

    Logger.log('✅ Google Sheet creado: https://docs.google.com/spreadsheets/d/' + idConvertido);
    return idConvertido;

  } catch (error) {
    Logger.log('❌ Error OneDrive: ' + error.toString());
    return false;
  }
}

function copiarArchivoDeDrive(idOrigen) {
  try {
    const carpeta = asegurarCarpeta('Sheets Reportes');
    const archivoOrigen = DriveApp.getFileById(idOrigen);
    const nombre = archivoOrigen.getName();
    const nuevoId = crearSpreadsheet(nombre, carpeta);
    copiarDesdeDrive(idOrigen, nuevoId);
    Logger.log('✅ Copiado desde Drive: ' + nuevoId);
    return nuevoId;
  } catch (e) {
    Logger.log('❌ Error al copiar desde Drive: ' + e);
    return false;
  }
}

function copiarDesdeOneDrive(url, idDestino) {
  try {
    const blob = obtenerBlobDesdeOneDrive(url);
    copiarContenidoDesdeBlob(blob, idDestino);
    Logger.log('✅ Contenido copiado desde OneDrive');
  } catch (error) {
    Logger.log('❌ Error al copiar desde OneDrive: ' + error.toString());
    throw error;
  }
}

function copiarContenidoDesdeBlob(blob, idDestino) {
  try {
    // Crear archivo temporal
    const archivoTemp = Drive.Files.insert({
      title: 'temp_' + Date.now(),
      mimeType: MimeType.GOOGLE_SHEETS
    }, blob);
    
    const libroTemporal = SpreadsheetApp.openById(archivoTemp.id);
    const libroDestino = SpreadsheetApp.openById(idDestino);

    // Limpiar hojas existentes en destino (excepto una)
    const hojasDestino = libroDestino.getSheets();
    if (hojasDestino.length > 1) {
      for (let i = 1; i < hojasDestino.length; i++) {
        libroDestino.deleteSheet(hojasDestino[i]);
      }
    }

    // Copiar cada hoja del temporal al destino
    const hojasTemporal = libroTemporal.getSheets();
    hojasTemporal.forEach((hoja, index) => {
      if (index === 0) {
        // Primera hoja: reemplazar la hoja existente
        const hojaDestino = libroDestino.getSheets()[0];
        const datos = hoja.getDataRange().getValues();
        if (datos.length > 0) {
          hojaDestino.clear();
          hojaDestino.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
        }
        hojaDestino.setName(hoja.getName());
      } else {
        // Hojas adicionales: copiar como nuevas hojas
        const nuevaHoja = hoja.copyTo(libroDestino);
        nuevaHoja.setName(hoja.getName());
      }
    });

    // Eliminar archivo temporal
    DriveApp.getFileById(archivoTemp.id).setTrashed(true);
    
  } catch (error) {
    Logger.log('❌ Error en copiarContenidoDesdeBlob: ' + error.toString());
    throw error;
  }
}

function copiarDesdeDrive(idOrigen, idDestino) {
  try {
    const origen = SpreadsheetApp.openById(idOrigen);
    const destino = SpreadsheetApp.openById(idDestino);

    // Limpiar hojas existentes en destino (excepto una)
    const hojasDestino = destino.getSheets();
    if (hojasDestino.length > 1) {
      for (let i = 1; i < hojasDestino.length; i++) {
        destino.deleteSheet(hojasDestino[i]);
      }
    }

    // Copiar hojas del origen
    const hojasOrigen = origen.getSheets();
    hojasOrigen.forEach((hoja, index) => {
      if (index === 0) {
        // Primera hoja: copiar con formato completo
        const hojaDestino = destino.getSheets()[0];
        copiarHojaCompleta(hoja, hojaDestino);
        hojaDestino.setName(hoja.getName());
      } else {
        // Hojas adicionales: usar copyTo que preserva formato
        const nuevaHoja = hoja.copyTo(destino);
        nuevaHoja.setName(hoja.getName());
      }
    });

    Logger.log('✅ Contenido copiado desde Drive');
    
  } catch (error) {
    Logger.log('❌ Error en copiarDesdeDrive: ' + error.toString());
    throw error;
  }
}

function copiarHojaCompleta(hojaOrigen, hojaDestino) {
  try {
    // Limpiar destino
    hojaDestino.clear();
    
    // Copiar datos
    const rango = hojaOrigen.getDataRange();
    if (rango.getNumRows() > 0) {
      const datos = rango.getValues();
      const rangoDestino = hojaDestino.getRange(1, 1, datos.length, datos[0].length);
      rangoDestino.setValues(datos);
      
      // Copiar formatos
      const formatos = rango.getTextStyles();
      rangoDestino.setTextStyles(formatos);
      
      const fondos = rango.getBackgrounds();
      rangoDestino.setBackgrounds(fondos);
      
      const bordes = rango.getBorder();
      if (bordes) {
        rangoDestino.setBorder(true, true, true, true, true, true);
      }
      
      // Copiar anchos de columna
      for (let col = 1; col <= datos[0].length; col++) {
        const ancho = hojaOrigen.getColumnWidth(col);
        hojaDestino.setColumnWidth(col, ancho);
      }
    }
  } catch (error) {
    Logger.log('❌ Error en copiarHojaCompleta: ' + error.toString());
    // Fallback: solo copiar datos
    const datos = hojaOrigen.getDataRange().getValues();
    if (datos.length > 0) {
      hojaDestino.clear();
      hojaDestino.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
    }
  }
}

function obtenerBlobDesdeOneDrive(url) {
  try {
    // Convertir URL de OneDrive a formato de descarga directa
    let downloadUrl = url;
    
    // Si es un enlace compartido de OneDrive, convertirlo
    if (url.includes('1drv.ms') || url.includes('onedrive.live.com')) {
      // Para enlaces 1drv.ms, primero seguir la redirección
      if (url.includes('1drv.ms')) {
        const response = UrlFetchApp.fetch(url, {
          followRedirects: false,
          muteHttpExceptions: true
        });
        
        const location = response.getAllHeaders()['Location'] || response.getAllHeaders()['location'];
        if (location) {
          downloadUrl = location;
        }
      }
      
      // Convertir a URL de descarga directa
      downloadUrl = downloadUrl.replace('?e=', '&e=').replace('view.aspx', 'download.aspx');
      if (!downloadUrl.includes('download.aspx')) {
        downloadUrl = downloadUrl.replace(/(\?|&)e=.*/, '') + '&download=1';
      }
    }

    Logger.log('📥 Descargando desde: ' + downloadUrl);

    // Descargar el archivo
    const response = UrlFetchApp.fetch(downloadUrl, {
      followRedirects: true,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Error HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
    }

    const blob = response.getBlob();
    const contentType = response.getHeaders()['Content-Type'] || '';
    
    Logger.log('📄 Content-Type: ' + contentType);
    Logger.log('📊 Tamaño del archivo: ' + blob.getBytes().length + ' bytes');

    // Validar que sea un archivo Excel
    if (blob.getBytes().length === 0) {
      throw new Error('El archivo descargado está vacío');
    }

    return blob;

  } catch (error) {
    Logger.log('❌ Error en obtenerBlobDesdeOneDrive: ' + error.toString());
    throw new Error('No se pudo descargar el archivo de OneDrive: ' + error.toString());
  }
}

function convertirBlobAGoogleSheet(blob, nombreArchivo, carpetaDestino) {
  try {
    Logger.log('🔄 Iniciando conversión de blob a Google Sheets...');
    
    // Método 1: Conversión directa
    try {
      const file = Drive.Files.create({
        name: nombreArchivo,
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [carpetaDestino.getId()]
      }, blob);
      
      Logger.log('✅ Conversión directa exitosa: ' + file.id);
      return file.id;
    } catch (directError) {
      Logger.log('⚠️ Conversión directa falló, intentando método alternativo...');
    }
    
    // Método 2: Crear archivo Excel primero, luego convertir
    const tempName = 'temp_' + Date.now() + '.xlsx';
    const excelFile = carpetaDestino.createFile(blob.setName(tempName));
    
    Logger.log('📄 Archivo Excel temporal creado: ' + excelFile.getId());
    
    // Convertir usando Drive API
    const convertedFile = Drive.Files.copy(excelFile.getId(), {
      name: nombreArchivo,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [carpetaDestino.getId()]
    });
    
    // Eliminar archivo temporal
    excelFile.setTrashed(true);
    
    Logger.log('✅ Conversión por método alternativo exitosa: ' + convertedFile.id);
    return convertedFile.id;
    
  } catch (error) {
    Logger.log('❌ Error en convertirBlobAGoogleSheet: ' + error.toString());
    throw new Error('No se pudo convertir el archivo: ' + error.toString());
  }
}