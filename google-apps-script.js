// ============================================================
// GOOGLE APPS SCRIPT — Nuevo Sol Inversiones
// Registro de Propiedades → Google Sheets
// ============================================================
//
// INSTRUCCIONES DE INSTALACIÓN:
//
// 1. Ve a Google Sheets → Crea un nuevo spreadsheet
//    Nómbralo: "Nuevo Sol — Registro de Propiedades"
//
// 2. Crea 3 hojas (tabs en la parte inferior):
//    - "Propiedades" (hoja principal)
//    - "Dashboard" (resumen automático)  
//    - "Config" (configuración)
//
// 3. En la hoja "Propiedades", pega estos headers en la fila 1:
//    A1: Timestamp
//    B1: Referencia ID
//    C1: Editar
//    D1: Título Propiedad
//    E1: Moneda
//    F1: Precio Venta
//    G1: Sector
//    H1: Ciudad
//    I1: Estado Propiedad
//    J1: M² Construcción
//    K1: M² Terreno
//    L1: Nivel/Piso
//    M1: Habitaciones
//    N1: Baños
//    O1: Medio Baño
//    P1: Parqueos Cantidad
//    Q1: Tipo Parqueos
//    R1: Espacios Incluidos
//    S1: Terminaciones
//    T1: Áreas Sociales
//    U1: Servicios Edificio
//    V1: Costo Mantenimiento
//    W1: Reserva
//    X1: Separación
//    Y1: Cuotas Obra
//    Z1: Fecha Entrega
//    AA1: Descripción Web
//    AB1: Documentos Disponibles
//    AC1: Notas Internas
//    AD1: Cuarto de Servicio (Cant.)
//    AE1: Jacuzzi (Cant.)
//    AF1: Antigüedad
//    AG1: Bauleras
//    AH1: Orientación
//    AI1: Amueblado
//
// 4. En la hoja "Config", pon:
//    A1: Versión      B1: 1.0
//    A2: Último Registro  B2: (se llenará automáticamente)
//
// 5. Ve a Extensiones (Extensions) → Apps Script
//
// 6. Borra todo el contenido de Code.gs y pega TODO este código
//
// 7. Guarda (Ctrl+S)
//
// 8. Click en "Implementar" (Deploy) → "Nueva implementación" (New deployment)
//    - Tipo: "Aplicación web" (Web app)
//    - Ejecutar como: "Yo" (Me)
//    - Quién tiene acceso: "Cualquiera" (Anyone)
//
// 9. Click en "Implementar" (Deploy)
//    - La primera vez te pedirá autorizar permisos → Autoriza
//
// 10. Copia la URL que te da (algo como https://script.google.com/macros/s/xxx/exec)
//
// 11. Pega esa URL en el formulario web (en el banner amarillo de configuración)
//
// ¡LISTO! Cada vez que envíes un formulario, los datos llegarán a tu Google Sheet.
//
// ============================================================

/**
 * Maneja las solicitudes POST del formulario web
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    
    // Configura aquí la URL de tu formulario publicado en GitHub Pages
    var FORM_URL = 'https://nuevosolinversionesrd-cloud.github.io/fichadepropiedades/index.html'; 
    
    // Determinar a qué hoja enviar los datos
    var sheetName = 'Propiedades';
    var isExternal = false;
    
    if (data.form_type === 'terceros' || data.isExternal) {
      sheetName = 'Publicaciones Externas';
      isExternal = true;
    }
    
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (isExternal) {
        addExternalHeaders(sheet);
      } else {
        addHeaders(sheet);
      }
    }
    
    // Formatear timestamp en zona horaria de RD
    var timestamp = Utilities.formatDate(
      new Date(), 
      'America/Santo_Domingo', 
      'dd/MM/yyyy HH:mm:ss'
    );
    
    // Manejar carga de firma a Google Drive si existe
    var firmaUrl = 'No adjunta';
    if (data.firma_base64) {
      try {
        var folderName = 'Firmas — Nuevo Sol';
        var folders = DriveApp.getFoldersByName(folderName);
        var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
        
        var contentType = data.firma_mimetype || 'image/png';
        var decoded = Utilities.base64Decode(data.firma_base64);
        var blob = Utilities.newBlob(decoded, contentType, 'Firma_' + data.nombre.replace(/ /g, '_') + '_' + data.cedula);
        var file = folder.createFile(blob);
        firmaUrl = file.getUrl();
      } catch (fError) {
        firmaUrl = 'Error al subir: ' + fError.toString();
      }
    }
    
    if (isExternal) {
      // Guardar en Publicaciones Externas
      sheet.appendRow([
        timestamp,                  // A: Timestamp
        data.nombre || '',          // B: Nombre
        data.cedula || '',          // C: Cédula / ID
        data.telefono || '',        // D: Teléfono
        data.email || '',           // E: Email
        data.titulo || '',          // F: Título Propiedad
        data.tipo || '',            // G: Tipo
        data.operacion || '',       // H: Operación
        data.precio || '',          // I: Precio
        data.ubicacion || '',       // J: Ubicación
        data.descripcion || '',     // K: Descripción
        data.fotos_link || '',      // L: Link Fotos
        firmaUrl,                   // M: Firma Digital (URL Drive)
        data.terminos ? 'Aceptado' : 'No Aceptado', // N: Términos
        data.comision_acepta ? 'Aceptado (2%)' : 'No Aceptado' // O: Comisión
      ]);
    } else {
      // Guardar en Propiedades (Original)
      var refId = data.referencia_id || '';
      var rowToUpdate = -1;
      
      // Buscar si la referencia ya existe
      if (refId) {
        var values = sheet.getRange(2, 2, sheet.getLastRow(), 1).getValues();
        for (var i = 0; i < values.length; i++) {
          if (values[i][0] === refId) {
            rowToUpdate = i + 2;
            break;
          }
        }
      }

      var editLink = '=HYPERLINK("' + FORM_URL + '?ref=' + refId + '"; "✏️ Editar")';
      
      var rowData = [
        timestamp,                        // A: Timestamp
        refId,                            // B: Referencia ID
        editLink,                         // C: Editar (Link)
        data.titulo_propiedad || '',      // D: Título Propiedad
        data.moneda || '',                // E: Moneda
        data.precio_venta || '',          // F: Precio Venta
        data.sector || '',                // G: Sector
        data.ciudad || '',                // H: Ciudad
        data.estado_propiedad || '',      // I: Estado Propiedad
        data.metros_construccion || '',   // J: M² Construcción
        data.metros_terreno || '',        // K: M² Terreno
        data.nivel_piso || '',            // L: Nivel/Piso
        data.habitaciones || '',          // M: Habitaciones
        data.banos || '',                 // N: Baños
        data.medios_banos || data.medio_bano || '', // O: Medio Baño
        data.parqueos_cantidad || '',     // P: Parqueos Cantidad
        data.tipo_parqueos || '',         // Q: Tipo Parqueos
        data.espacios_incluidos || '',    // R: Espacios Incluidos
        data.terminaciones || '',         // S: Terminaciones
        data.areas_sociales || '',        // T: Áreas Sociales
        data.servicios_edificio || '',    // U: Servicios Edificio
        data.costo_mantenimiento || '',   // V: Costo Mantenimiento
        data.reserva || '',               // W: Reserva
        data.separacion || '',            // X: Separación
        data.cuotas_obra || '',           // Y: Cuotas Obra
        data.fecha_entrega || '',         // Z: Fecha Entrega
        data.descripcion_web || '',       // AA: Descripción Web
        data.documentos || '',            // AB: Documentos Disponibles
        data.notas_internas || '',        // AC: Notas Internas
        data.cuarto_servicio || '',       // AD: Cuarto de Servicio (Cant.)
        data.cantidad_jacuzzi || '',      // AE: Jacuzzi (Cant.)
        data.antiguedad || '',            // AF: Antigüedad
        data.bauleras || '',              // AG: Bauleras
        data.orientacion || '',           // AH: Orientación
        data.amueblado || ''              // AI: Amueblado
      ];

      if (rowToUpdate > -1) {
        // Actualizar fila existente
        sheet.getRange(rowToUpdate, 1, 1, rowData.length).setValues([rowData]);
      } else {
        // Nueva fila
        sheet.appendRow(rowData);
      }
    }
    
    // Actualizar Config con último registro
    updateConfig(ss, timestamp, data.referencia_id || data.nombre);
    
    // Actualizar Dashboard (solo para propiedades internas por ahora)
    if (!isExternal) {
      updateDashboard(ss);
    }
    
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', message: 'Registro completado exitosamente' })
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Agrega los headers a la hoja de Publicaciones Externas
 */
function addExternalHeaders(sheet) {
  var headers = [
    'Timestamp', 'Nombre', 'Cédula / ID', 'Teléfono', 'Email', 
    'Título Propiedad', 'Tipo', 'Operación', 'Precio (USD)', 
    'Ubicación', 'Descripción', 'Link Fotos', 'Firma Digital', 
    'Aceptación Términos', 'Aceptación Comisión'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatear headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#E79E24');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontSize(10);
  
  // Congelar fila de headers
  sheet.setFrozenRows(1);
  
  // Ajustar anchos de columnas
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 160);
  }
}

/**
 * Maneja solicitudes GET (para verificar que el script funciona)
 */
function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'getProperty') {
    var ref = e.parameter.ref;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Propiedades');
    var data = sheet.getDataRange().getValues();
    
    // Headers para mapeo
    var headers = data[0];
    var property = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === ref) { // Columna B: Referencia ID
        property = {};
        for (var j = 0; j < headers.length; j++) {
          property[headers[j]] = data[i][j];
        }
        break;
      }
    }
    
    return ContentService.createTextOutput(
      JSON.stringify(property)
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ 
      status: 'active', 
      message: 'Nuevo Sol Inversiones — API de Registro activa',
      version: '1.1'
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Agrega los headers a la hoja de Propiedades
 */
function addHeaders(sheet) {
  var headers = [
    'Timestamp', 'Referencia ID', 'Editar', 'Título Propiedad', 'Moneda', 'Precio Venta',
    'Sector', 'Ciudad', 'Estado Propiedad', 'M² Construcción', 'M² Terreno',
    'Nivel/Piso', 'Habitaciones', 'Baños', 'Medio Baño', 'Parqueos Cantidad',
    'Tipo Parqueos', 'Espacios Incluidos', 'Terminaciones', 'Áreas Sociales',
    'Servicios Edificio', 'Costo Mantenimiento', 'Reserva', 'Separación',
    'Cuotas Obra', 'Fecha Entrega', 'Descripción Web', 'Documentos Disponibles',
    'Notas Internas', 'Cuarto de Servicio (Cant.)', 'Jacuzzi (Cant.)',
    'Antigüedad', 'Bauleras', 'Orientación', 'Amueblado'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatear headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a3c5e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontSize(10);
  
  // Congelar fila de headers
  sheet.setFrozenRows(1);
  
  // Ajustar anchos de columnas
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 140);
  }
}

/**
 * Actualiza la hoja Config con la última actividad
 */
function updateConfig(ss, timestamp, refId) {
  var configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.getRange('A1').setValue('Versión');
    configSheet.getRange('B1').setValue('1.0');
    configSheet.getRange('A2').setValue('Último Registro');
    configSheet.getRange('A3').setValue('Total Propiedades');
    
    // Formatear
    configSheet.getRange('A1:A3').setFontWeight('bold');
    configSheet.setColumnWidth(1, 180);
    configSheet.setColumnWidth(2, 250);
  }
  
  configSheet.getRange('B2').setValue(timestamp + ' — ' + (refId || 'Sin Ref'));
  
  var propSheet = ss.getSheetByName('Propiedades');
  if (propSheet) {
    var totalRows = Math.max(0, propSheet.getLastRow() - 1);
    configSheet.getRange('B3').setValue(totalRows);
  }
}

/**
 * Actualiza el Dashboard con estadísticas resumidas
 */
function updateDashboard(ss) {
  var dashSheet = ss.getSheetByName('Dashboard');
  
  if (!dashSheet) {
    dashSheet = ss.insertSheet('Dashboard');
  }
  
  var propSheet = ss.getSheetByName('Propiedades');
  if (!propSheet) return;
  
  var lastRow = propSheet.getLastRow();
  if (lastRow < 2) return; // No hay datos aún
  
  // Limpiar dashboard
  dashSheet.clear();
  
  // --- TÍTULO ---
  dashSheet.getRange('A1').setValue('📊 DASHBOARD — NUEVO SOL INVERSIONES');
  dashSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#1a3c5e');
  dashSheet.getRange('A2').setValue('Actualizado: ' + Utilities.formatDate(new Date(), 'America/Santo_Domingo', 'dd/MM/yyyy HH:mm'));
  dashSheet.getRange('A2').setFontSize(10).setFontColor('#888');
  
  // --- RESUMEN GENERAL ---
  dashSheet.getRange('A4').setValue('📋 RESUMEN GENERAL').setFontWeight('bold').setFontSize(12).setFontColor('#2e6da4');
  
  var totalProps = lastRow - 1;
  dashSheet.getRange('A5').setValue('Total Propiedades Registradas');
  dashSheet.getRange('B5').setValue(totalProps).setFontWeight('bold').setFontSize(14);
  
  // --- POR ESTADO ---
  dashSheet.getRange('A7').setValue('🏗️ POR ESTADO').setFontWeight('bold').setFontSize(12).setFontColor('#2e6da4');
  
  var estados = {};
  if (lastRow >= 2) {
    var estadoData = propSheet.getRange(2, 9, lastRow - 1, 1).getValues(); // Columna I
    estadoData.forEach(function(row) {
      var val = row[0] ? row[0].toString().trim() : 'Sin especificar';
      estados[val] = (estados[val] || 0) + 1;
    });
  }
  
  var row = 8;
  for (var estado in estados) {
    dashSheet.getRange('A' + row).setValue(estado);
    dashSheet.getRange('B' + row).setValue(estados[estado]).setFontWeight('bold');
    row++;
  }
  
  // --- POR CIUDAD ---
  row += 1;
  dashSheet.getRange('A' + row).setValue('🌆 POR CIUDAD').setFontWeight('bold').setFontSize(12).setFontColor('#2e6da4');
  row++;
  
  var ciudades = {};
  if (lastRow >= 2) {
    var ciudadData = propSheet.getRange(2, 8, lastRow - 1, 1).getValues(); // Columna H
    ciudadData.forEach(function(r) {
      var val = r[0] ? r[0].toString().trim() : 'Sin especificar';
      ciudades[val] = (ciudades[val] || 0) + 1;
    });
  }
  
  for (var ciudad in ciudades) {
    dashSheet.getRange('A' + row).setValue(ciudad);
    dashSheet.getRange('B' + row).setValue(ciudades[ciudad]).setFontWeight('bold');
    row++;
  }
  
  // --- RANGO DE PRECIOS ---
  row += 1;
  dashSheet.getRange('A' + row).setValue('💰 RANGO DE PRECIOS').setFontWeight('bold').setFontSize(12).setFontColor('#2e6da4');
  row++;
  
  if (lastRow >= 2) {
    var precioData = propSheet.getRange(2, 6, lastRow - 1, 1).getValues(); // Columna F
    var precios = [];
    precioData.forEach(function(r) {
      var val = parseFloat(r[0].toString().replace(/[^0-9.]/g, ''));
      if (!isNaN(val) && val > 0) precios.push(val);
    });
    
    if (precios.length > 0) {
      dashSheet.getRange('A' + row).setValue('Precio Mínimo');
      dashSheet.getRange('B' + row).setValue('$' + Math.min.apply(null, precios).toLocaleString());
      row++;
      dashSheet.getRange('A' + row).setValue('Precio Máximo');
      dashSheet.getRange('B' + row).setValue('$' + Math.max.apply(null, precios).toLocaleString());
      row++;
      var avg = precios.reduce(function(a, b) { return a + b; }, 0) / precios.length;
      dashSheet.getRange('A' + row).setValue('Precio Promedio');
      dashSheet.getRange('B' + row).setValue('$' + Math.round(avg).toLocaleString());
    } else {
      dashSheet.getRange('A' + row).setValue('Sin datos de precios aún');
    }
  }
  
  // Ajustar columnas
  dashSheet.setColumnWidth(1, 280);
  dashSheet.setColumnWidth(2, 180);
}

/**
 * Función de inicialización — Ejecutar una vez para crear las hojas
 * Puedes ejecutar esta función manualmente desde el editor de Apps Script
 * haciendo clic en ▶️ Run
 */
function inicializarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Crear hoja Propiedades si no existe
  var propSheet = ss.getSheetByName('Propiedades');
  if (!propSheet) {
    propSheet = ss.insertSheet('Propiedades');
    addHeaders(propSheet);
    Logger.log('✅ Hoja "Propiedades" creada con headers');
  } else {
    Logger.log('ℹ️ Hoja "Propiedades" ya existe');
  }
  
  // Crear hoja Dashboard si no existe
  var dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) {
    dashSheet = ss.insertSheet('Dashboard');
    dashSheet.getRange('A1').setValue('📊 DASHBOARD — NUEVO SOL INVERSIONES');
    dashSheet.getRange('A1').setFontSize(14).setFontWeight('bold');
    dashSheet.getRange('A2').setValue('Sin datos aún. Envía tu primer formulario.');
    Logger.log('✅ Hoja "Dashboard" creada');
  } else {
    Logger.log('ℹ️ Hoja "Dashboard" ya existe');
  }
  
  // Crear hoja Config si no existe
  var configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.getRange('A1').setValue('Versión');
    configSheet.getRange('B1').setValue('1.0');
    configSheet.getRange('A2').setValue('Último Registro');
    configSheet.getRange('B2').setValue('Ninguno aún');
    configSheet.getRange('A3').setValue('Total Propiedades');
    configSheet.getRange('B3').setValue(0);
    configSheet.getRange('A1:A3').setFontWeight('bold');
    configSheet.setColumnWidth(1, 180);
    configSheet.setColumnWidth(2, 250);
    Logger.log('✅ Hoja "Config" creada');
  } else {
    Logger.log('ℹ️ Hoja "Config" ya existe');
  }
  
  Logger.log('🎉 Inicialización completada');
}

/**
 * Función para actualizar los headers manualmente en una hoja existente.
 * Ejecuta esta función desde el editor de Apps Script si has añadido campos nuevos.
 */
function actualizarHeadersManual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Propiedades');
  
  if (!sheet) {
    Logger.log('❌ No se encontró la hoja "Propiedades"');
    return;
  }

  // 1. Detectar estado actual
  var headerC = sheet.getRange(1, 3).getValue();
  var headerD = sheet.getRange(1, 4).getValue();
  var lastRow = sheet.getLastRow();
  
  // CASO: La columna "Editar" ya existe pero los datos están desplazados (como en tu captura)
  if (headerC === 'Editar' && lastRow >= 2) {
    // Revisamos si la columna D tiene errores o está vacía mientras la E tiene el título
    var testCellD = sheet.getRange(2, 4).getValue(); // D2
    var testCellE = sheet.getRange(2, 5).getValue(); // E2
    
    // Si D está mal y E parece tener datos que deberían estar en D
    if ((testCellD === '' || testCellD.toString().includes('#')) && testCellE !== '') {
      Logger.log('⚠️ Detectada desalineación. Corrigiendo...');
      // Mover datos de E2:LastCol a D2
      var lastCol = sheet.getLastColumn();
      var dataRange = sheet.getRange(2, 5, lastRow - 1, lastCol - 4);
      dataRange.moveTo(sheet.getRange(2, 4));
      Logger.log('✅ Datos desplazados a la izquierda correctamente');
    }
  } 
  // CASO: Aún no existe la columna "Editar"
  else if (headerC !== 'Editar') {
    sheet.insertColumnBefore(3);
    Logger.log('✅ Columna C insertada para "Editar"');
  }

  // 2. Lista completa de headers actualizada
  var headers = [
    'Timestamp', 'Referencia ID', 'Editar', 'Título Propiedad', 'Moneda', 'Precio Venta',
    'Sector', 'Ciudad', 'Estado Propiedad', 'M² Construcción', 'M² Terreno',
    'Nivel/Piso', 'Habitaciones', 'Baños', 'Medio Baño', 'Parqueos Cantidad',
    'Tipo Parqueos', 'Espacios Incluidos', 'Terminaciones', 'Áreas Sociales',
    'Servicios Edificio', 'Costo Mantenimiento', 'Reserva', 'Separación',
    'Cuotas Obra', 'Fecha Entrega', 'Descripción Web', 'Documentos Disponibles',
    'Notas Internas', 'Cuarto de Servicio (Cant.)', 'Jacuzzi (Cant.)',
    'Antigüedad', 'Bauleras', 'Orientación', 'Amueblado'
  ];
  
  // 3. Aplicar headers a la primera fila (esto asegura que los nombres coincidan)
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 4. Formatear headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a3c5e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontSize(10);
  
  // 5. Generar/Actualizar los enlaces de edición
  var FORM_URL = 'https://nuevosolinversionesrd-cloud.github.io/fichadepropiedades/index.html';
  if (lastRow >= 2) {
    var refIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var editLinks = [];
    for (var i = 0; i < refIds.length; i++) {
      var refId = refIds[i][0];
      if (refId) {
        editLinks.push(['=HYPERLINK("' + FORM_URL + '?ref=' + refId + '"; "✏️ Editar")']);
      } else {
        editLinks.push(['']);
      }
    }
    sheet.getRange(2, 3, editLinks.length, 1).setValues(editLinks);
  }
  
  // 6. Ajustar anchos
  for (var i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, 140);
  }
  
  Logger.log('🎉 Proceso completado. Hoja sincronizada.');
}

