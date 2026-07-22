/**
 * CÓDIGO DE GOOGLE APPS SCRIPT PARA CONTROL DE LICENCIAS DE ACCESOQR (MULTIDISPOSITIVO)
 * 
 * INSTRUCCIONES:
 * 1. Crea una hoja de cálculo en Google Sheets.
 * 2. Coloca en la primera fila (encabezados): 
 *    Celda A1: Clave Licencia
 *    Celda B1: Cliente
 *    Celda C1: Estado
 *    Celda D1: Dispositivos
 *    Celda E1: Limite Dispositivos
 * 3. En el menú superior, ve a "Extensiones" -> "Apps Script".
 * 4. Borra el código por defecto, pega este código de abajo y guarda.
 * 5. Haz clic en "Implementar" -> "Nueva implementación".
 * 6. Tipo: "Aplicación web".
 * 7. Ejecutar como: "Tú (tu correo)".
 * 8. Quién tiene acceso: "Cualquiera".
 * 9. Copia la URL de la aplicación web obtenida y pégala en el archivo `app.js` de la aplicación.
 */

function doGet(e) {
  return handleLicenseRequest(e);
}

function doPost(e) {
  return handleLicenseRequest(e);
}

function handleLicenseRequest(e) {
  // Manejo de CORS
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    var params = e.parameter;
    
    // Si la petición viene como JSON en el cuerpo (POST)
    if (e.postData && e.postData.contents) {
      try {
        var postParams = JSON.parse(e.postData.contents);
        params = Object.assign({}, params, postParams);
      } catch(err) {
        // No es JSON válido, continuamos con parámetros URL
      }
    }

    var licenseKey = params.key;
    var deviceId = params.device;
    
    if (!licenseKey || !deviceId) {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "Parámetros insuficientes. Se requiere 'key' y 'device'." 
      }));
    }

    // Abrir hoja de cálculo activa
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    var licenseRowIndex = -1;
    
    // Buscar la clave en la columna A (Clave Licencia)
    // Empezamos desde i = 1 para saltar la fila de títulos
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(licenseKey).trim()) {
        licenseRowIndex = i;
        break;
      }
    }
    
    // Si no se encuentra la licencia
    if (licenseRowIndex === -1) {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "La clave de licencia ingresada no existe." 
      }));
    }
    
    // Leer valores de la fila encontrada
    var currentStatus = String(data[licenseRowIndex][2]).trim().toLowerCase(); // Columna C (Estado)
    var devicesString = String(data[licenseRowIndex][3]).trim();               // Columna D (Dispositivos)
    var deviceLimitRaw = data[licenseRowIndex][4];                            // Columna E (Límite Dispositivos)
    
    // Determinar límite numérico
    var deviceLimit = parseInt(deviceLimitRaw, 10);
    if (isNaN(deviceLimit) || deviceLimit <= 0) {
      deviceLimit = 1; // Por defecto mínimo 1 dispositivo si está vacío o no es un número
    }
    
    // Si la licencia está desactivada o suspendida
    if (currentStatus !== "activa") {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "Esta licencia se encuentra suspendida o inactiva." 
      }));
    }
    
    // Analizar la lista de dispositivos actuales
    var activeDevices = [];
    if (devicesString !== "") {
      activeDevices = devicesString.split(",").map(function(item) {
        return item.trim();
      }).filter(function(item) {
        return item !== "";
      });
    }
    
    var guestLimitRaw = data[licenseRowIndex][5];                            // Columna F (Límite Invitados)
    var guestLimit = parseInt(guestLimitRaw, 10);
    if (isNaN(guestLimit) || guestLimit <= 0) {
      guestLimit = 9999; // Por defecto 9999 (sin límite) si no está definido
    }

    // Si el dispositivo actual ya está registrado en la lista
    if (activeDevices.indexOf(deviceId) !== -1) {
      return output.setContent(JSON.stringify({ 
        success: true, 
        message: "Licencia verificada (Dispositivo ya registrado).",
        client: data[licenseRowIndex][1],
        guestLimit: guestLimit
      }));
    }
    
    // Si el dispositivo es nuevo pero ya se alcanzó el límite permitido
    if (activeDevices.length >= deviceLimit) {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "Límite de dispositivos alcanzado (" + activeDevices.length + "/" + deviceLimit + "). Contacta a soporte para ampliar tu plan." 
      }));
    }
    
    // Registrar el nuevo dispositivo
    activeDevices.push(deviceId);
    var newDevicesString = activeDevices.join(",");
    
    // Escribir en la Columna D (Fila es 1-indexed en Google Sheets, así que es licenseRowIndex + 1)
    sheet.getRange(licenseRowIndex + 1, 4).setValue(newDevicesString);
    
    // Retorno exitoso
    return output.setContent(JSON.stringify({ 
      success: true, 
      message: "Licencia verificada con éxito (" + activeDevices.length + "/" + deviceLimit + " dispositivos registrados).",
      client: data[licenseRowIndex][1], // Devolver el nombre del cliente
      guestLimit: guestLimit
    }));

  } catch(error) {
    return output.setContent(JSON.stringify({ 
      success: false, 
      message: "Error interno del servidor: " + error.toString() 
    }));
  }
}
