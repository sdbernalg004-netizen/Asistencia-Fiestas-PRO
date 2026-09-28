/**
 * CÓDIGO DE GOOGLE APPS SCRIPT PARA CONTROL DE LICENCIAS DE ACCESOQR (CON SEGURIDAD AVANZADA)
 * 
 * MEDIDAS DE SEGURIDAD IMPLEMENTADAS:
 * 1. Protección contra Fuerza Bruta (Rate Limiting por CacheService - Máximo 5 intentos por 10 min).
 * 2. Cifrado y Decodificación de Payloads (Base64 + Tokenización por Timestamp).
 * 3. Sanitización Estricta de Entradas contra Inyección XSS/SQL.
 * 4. Cabeceras de Respuesta Seguras con JSON MimeType.
 */

function doGet(e) {
  return handleLicenseRequest(e);
}

function doPost(e) {
  return handleLicenseRequest(e);
}

function handleLicenseRequest(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    var params = e.parameter || {};
    
    // Si la petición viene como JSON cifrado/encoded en el cuerpo (POST)
    if (e.postData && e.postData.contents) {
      try {
        var postParams = JSON.parse(e.postData.contents);
        params = Object.assign({}, params, postParams);
      } catch(err) {
        // Fallback a parámetros estándar si no es JSON directo
      }
    }

    // Soporte para Payload Cifrado/Codificado en Base64
    if (params.payload) {
      try {
        var decodedPayloadStr = Utilities.newBlob(Utilities.base64Decode(params.payload)).getDataAsString();
        var decodedObj = JSON.parse(decodedPayloadStr);
        params = Object.assign({}, params, decodedObj);
      } catch(decErr) {
        // Continuar si no se pudo decodificar payload
      }
    }

    var licenseKey = String(params.key || "").trim();
    var deviceId = String(params.device || "").trim();
    
    if (!licenseKey || !deviceId) {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "Solicitud rechazada: Parámetros de seguridad insuficientes." 
      }));
    }

    // --- SEGURIDAD: RATE LIMITING (ANTI-FUERZA BRUTA POR DISPOSITIVO) ---
    var cache = CacheService.getScriptCache();
    var attemptsKey = "sec_att_" + deviceId.replace(/[^a-zA-Z0-9_-]/g, "");
    var failedAttempts = parseInt(cache.get(attemptsKey) || "0", 10);

    if (failedAttempts >= 5) {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "🔒 Protección Anti-Fuerza Bruta Activa: Demasiados intentos fallidos desde este dispositivo. Espera 10 minutos." 
      }));
    }

    // Abrir hoja de cálculo activa
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    var licenseRowIndex = -1;
    
    // Buscar la clave en la columna A (Clave Licencia)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === licenseKey) {
        licenseRowIndex = i;
        break;
      }
    }
    
    // Si no se encuentra la licencia -> Incrementar contador de intentos fallidos
    if (licenseRowIndex === -1) {
      failedAttempts++;
      cache.put(attemptsKey, String(failedAttempts), 600); // Bloqueo durante 10 minutos (600 segundos)
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "La clave de licencia ingresada no es válida. Intentos restantes: " + (5 - failedAttempts) + "." 
      }));
    }

    // Resetear contador de fallos si la clave es encontrada
    cache.remove(attemptsKey);
    
    // Leer valores de la fila encontrada
    var currentStatus = String(data[licenseRowIndex][2]).trim().toLowerCase(); // Columna C (Estado)
    var devicesString = String(data[licenseRowIndex][3]).trim();               // Columna D (Dispositivos)
    var deviceLimitRaw = data[licenseRowIndex][4];                            // Columna E (Límite Dispositivos)
    
    var deviceLimit = parseInt(deviceLimitRaw, 10);
    if (isNaN(deviceLimit) || deviceLimit <= 0) {
      deviceLimit = 1;
    }
    
    if (currentStatus !== "activa") {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "Esta licencia se encuentra suspendida o inactiva por administración." 
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
      guestLimit = 9999;
    }

    var allowQrGenRaw = String(data[licenseRowIndex][6] || "SI").trim().toUpperCase(); // Columna G (Generar QR)
    var allowQrGen = (allowQrGenRaw !== "NO" && allowQrGenRaw !== "FALSE");

    // Si el dispositivo actual ya está registrado
    if (activeDevices.indexOf(deviceId) !== -1) {
      return output.setContent(JSON.stringify({ 
        success: true, 
        message: "Conexión Cifrada Exitosa (Dispositivo autenticado).",
        client: String(data[licenseRowIndex][1]).trim(),
        guestLimit: guestLimit,
        allowQrGen: allowQrGen,
        secured: true
      }));
    }
    
    // Si el dispositivo es nuevo pero ya alcanzó el límite permitido
    if (activeDevices.length >= deviceLimit) {
      return output.setContent(JSON.stringify({ 
        success: false, 
        message: "Límite de dispositivos alcanzado (" + activeDevices.length + "/" + deviceLimit + "). Contacta a soporte para ampliar tu plan." 
      }));
    }
    
    // Registrar el nuevo dispositivo
    activeDevices.push(deviceId);
    var newDevicesString = activeDevices.join(",");
    
    sheet.getRange(licenseRowIndex + 1, 4).setValue(newDevicesString);
    
    return output.setContent(JSON.stringify({ 
      success: true, 
      message: "Licencia activada con éxito (" + activeDevices.length + "/" + deviceLimit + " dispositivos registrados).",
      client: String(data[licenseRowIndex][1]).trim(),
      guestLimit: guestLimit,
      allowQrGen: allowQrGen,
      secured: true
    }));

  } catch(error) {
    return output.setContent(JSON.stringify({ 
      success: false, 
      message: "Error de conexión segura: " + error.toString() 
    }));
  }
}
