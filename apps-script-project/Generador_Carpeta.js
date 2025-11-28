/**
 * Transfiere el contenido de la hoja Generador a un nuevo archivo en la carpeta del cliente
 * y luego limpia la hoja Generador eliminando las filas usadas
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Acciones')
      .addItem('📁 Enviar a Carpeta Cliente', 'transferirGeneradorACliente')
      .addSeparator()
      .addItem('🔐 Autorizar Permisos Drive', 'autorizarPermisosDrive')
      .addItem('🧪 Probar Acceso a Carpeta', 'probarAccesoCarpeta')
      .addToUi();
}


function transferirGeneradorACliente() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    var ui = SpreadsheetApp.getUi();

    Logger.log("🚀 INICIANDO transferirGeneradorACliente...");

    // ✅ PASO 1: Obtener información del cliente y carpeta
    Logger.log("📋 PASO 1: Obteniendo información del cliente...");
    var infoCliente = obtenerInformacionCliente(hojaGenerador);
    if (!infoCliente.exito) {
      Logger.log("❌ FALLÓ PASO 1: " + infoCliente.mensaje);
      ui.alert('Error en Paso 1', infoCliente.mensaje, ui.ButtonSet.OK);
      return infoCliente.mensaje;
    }
    Logger.log("✅ PASO 1 COMPLETADO");
    
    // ✅ PASO 2: Determinar el rango de datos a copiar
    Logger.log("📋 PASO 2: Determinando rango de datos...");
    var ultimaFila = hojaGenerador.getRange("K1").getValue();
    if (!ultimaFila || ultimaFila < 4) {
      Logger.log("❌ FALLÓ PASO 2: No hay datos para transferir (K1 = " + ultimaFila + ")");
      ui.alert('Error en Paso 2', "No hay datos para transferir (K1 = " + ultimaFila + ")", ui.ButtonSet.OK);
      return "Error: No hay datos para transferir (K1 = " + ultimaFila + ")";
    }

    var rangoDatos = "A1:M" + ultimaFila; // Desde encabezados hasta última fila
    Logger.log("Rango a copiar: " + rangoDatos);
    Logger.log("✅ PASO 2 COMPLETADO");

    // ✅ PASO 3: Crear nuevo archivo en carpeta del cliente
    Logger.log("📋 PASO 3: Creando archivo en carpeta del cliente...");
    var nuevoArchivo = crearArchivoEnCarpetaCliente(infoCliente.carpetaId, infoCliente.nombreCliente);
    if (!nuevoArchivo.exito) {
      Logger.log("❌ FALLÓ PASO 3: " + nuevoArchivo.mensaje);
      ui.alert('Error en Paso 3 - Creación de archivo', nuevoArchivo.mensaje, ui.ButtonSet.OK);
      return nuevoArchivo.mensaje;
    }
    Logger.log("✅ PASO 3 COMPLETADO - Archivo creado: " + nuevoArchivo.url);

    // ✅ PASO 4: Copiar datos al nuevo archivo
    Logger.log("📋 PASO 4: Copiando datos al nuevo archivo...");
    var resultadoCopia = copiarDatosAlNuevoArchivo(hojaGenerador, nuevoArchivo.archivo, rangoDatos);
    if (!resultadoCopia.exito) {
      Logger.log("❌ FALLÓ PASO 4: " + resultadoCopia.mensaje);
      ui.alert('Error en Paso 4 - Copia de datos', resultadoCopia.mensaje, ui.ButtonSet.OK);
      return resultadoCopia.mensaje;
    }
    Logger.log("✅ PASO 4 COMPLETADO");

    // ✅ PASO 5: Limpiar hoja Generador (eliminar filas y desplazar hacia arriba)
    Logger.log("📋 PASO 5: Limpiando hoja Generador...");
    var resultadoLimpieza = limpiarGeneradorConDesplazamiento(hojaGenerador, ultimaFila);
    if (!resultadoLimpieza.exito) {
      Logger.log("❌ FALLÓ PASO 5: " + resultadoLimpieza.mensaje);
      ui.alert('Error en Paso 5 - Limpieza', resultadoLimpieza.mensaje, ui.ButtonSet.OK);
      return resultadoLimpieza.mensaje;
    }
    Logger.log("✅ PASO 5 COMPLETADO");

    Logger.log("🎉 TODO COMPLETADO EXITOSAMENTE");
    ui.alert('¡Éxito!', 'Archivo creado en carpeta del cliente y Generador limpiado.\n\nURL: ' + nuevoArchivo.url, ui.ButtonSet.OK);
    return "✅ Éxito: Archivo creado en carpeta del cliente y Generador limpiado. URL: " + nuevoArchivo.url;
    
  } catch (e) {
    Logger.log("💥 ERROR GENERAL en transferirGeneradorACliente: " + e.toString());
    Logger.log("📍 Stack trace: " + e.stack);
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error General', 'Ocurrió un error inesperado:\n\n' + e.message + '\n\nRevisa los logs para más detalles (Extensiones > Apps Script > Ver logs)', ui.ButtonSet.OK);
    return "Error: " + e.message;
  }
}

/**
 * Obtiene la información del cliente y el ID de su carpeta desde F2
 */
function obtenerInformacionCliente(hoja) {
  try {
    // Obtener el valor y el hipervínculo de F2
    var celdaF2 = hoja.getRange("F2");
    var valorF2 = celdaF2.getValue().toString().trim();
    var hipervinculos = celdaF2.getRichTextValue().getLinkUrl();
    
    if (!hipervinculos) {
      // Si no hay hipervínculo en el RichText, intentar con getFormula
      var formula = celdaF2.getFormula();
      var matchHyperlink = formula.match(/HYPERLINK\("([^"]+)"/);
      if (matchHyperlink) {
        hipervinculos = matchHyperlink[1];
      }
    }
    
    if (!hipervinculos) {
      return { exito: false, mensaje: "Error: No se encontró hipervínculo en F2" };
    }
    
    // Extraer ID de carpeta del URL de Google Drive
    var carpetaId = extraerIdCarpeta(hipervinculos);
    if (!carpetaId) {
      return { exito: false, mensaje: "Error: No se pudo extraer ID de carpeta de: " + hipervinculos };
    }
    
    // Usar valorF2 como nombre del cliente (sin el "Ver Carpeta")
    var nombreCliente = valorF2.replace("Ver Carpeta", "").trim();
    if (!nombreCliente) {
      nombreCliente = "Cliente_" + Utilities.formatDate(new Date(), "GMT-6", "yyyyMMdd");
    }
    
    Logger.log("Cliente: " + nombreCliente + " | Carpeta ID: " + carpetaId);
    
    return {
      exito: true,
      nombreCliente: nombreCliente,
      carpetaId: carpetaId
    };
    
  } catch (e) {
    return { exito: false, mensaje: "Error obteniendo info cliente: " + e.message };
  }
}

/**
 * Extrae el ID de carpeta de un URL de Google Drive
 */
function extraerIdCarpeta(url) {
  try {
    // Diferentes formatos de URLs de Google Drive
    var patterns = [
      /\/folders\/([a-zA-Z0-9-_]+)/,  // /folders/ID
      /id=([a-zA-Z0-9-_]+)/,          // ?id=ID
      /^([a-zA-Z0-9-_]+)$/            // Solo el ID
    ];
    
    for (var i = 0; i < patterns.length; i++) {
      var match = url.match(patterns[i]);
      if (match) {
        return match[1];
      }
    }
    
    return null;
  } catch (e) {
    Logger.log("Error extrayendo ID carpeta: " + e.message);
    return null;
  }
}

/**
 * Crea un nuevo archivo de Google Sheets en la carpeta del cliente
 */
function crearArchivoEnCarpetaCliente(carpetaId, nombreCliente) {
  try {
    Logger.log("🔍 Intentando acceder a carpeta ID: " + carpetaId);
    
    // Verificar que la carpeta existe y tenemos acceso
    var carpeta;
    try {
      carpeta = DriveApp.getFolderById(carpetaId);
      Logger.log("✅ Carpeta encontrada: " + carpeta.getName());
    } catch (errorCarpeta) {
      Logger.log("❌ Error accediendo a carpeta: " + errorCarpeta.message);
      return { exito: false, mensaje: "No se puede acceder a la carpeta. Verificar permisos o ID: " + errorCarpeta.message };
    }
    
    // Crear nombre único para el archivo
    var timestamp = Utilities.formatDate(new Date(), "GMT-6", "yyyyMMdd_HHmm");
    var nombreArchivo = "Cotización_" + nombreCliente + "_" + timestamp;
    Logger.log("📄 Creando archivo: " + nombreArchivo);
    
    // Crear nuevo Google Sheets
    var nuevoArchivo;
    try {
      nuevoArchivo = SpreadsheetApp.create(nombreArchivo);
      Logger.log("✅ Archivo creado en Mi Unidad: " + nuevoArchivo.getId());
    } catch (errorCreacion) {
      Logger.log("❌ Error creando archivo: " + errorCreacion.message);
      return { exito: false, mensaje: "Error creando archivo: " + errorCreacion.message };
    }
    
    // Mover el archivo a la carpeta del cliente
    try {
      var archivo = DriveApp.getFileById(nuevoArchivo.getId());
      Logger.log("📁 Moviendo archivo a carpeta del cliente...");
      
      carpeta.addFile(archivo);
      Logger.log("✅ Archivo agregado a carpeta del cliente");
      
      DriveApp.getRootFolder().removeFile(archivo); // Remover de "Mi unidad"
      Logger.log("✅ Archivo removido de Mi Unidad");
      
    } catch (errorMovimiento) {
      Logger.log("❌ Error moviendo archivo: " + errorMovimiento.message);
      // El archivo se creó pero no se pudo mover, mejor dejarlo en Mi Unidad
      Logger.log("⚠️ Archivo quedó en Mi Unidad debido al error de movimiento");
    }
    
    Logger.log("✅ Archivo final URL: " + nuevoArchivo.getUrl());
    
    return {
      exito: true,
      archivo: nuevoArchivo,
      url: nuevoArchivo.getUrl()
    };
    
  } catch (e) {
    Logger.log("💥 Error general en crearArchivoEnCarpetaCliente: " + e.message);
    Logger.log("📍 Stack trace: " + e.stack);
    return { exito: false, mensaje: "Error general creando archivo: " + e.message };
  }
}

/**
 * Copia los datos de la hoja Generador al nuevo archivo
 */
function copiarDatosAlNuevoArchivo(hojaOrigen, archivoDestino, rangoDatos) {
  try {
    Logger.log("📋 Iniciando copia de datos...");
    
    // Obtener la hoja activa del nuevo archivo (por defecto "Hoja 1")
    var hojaDestino = archivoDestino.getActiveSheet();
    hojaDestino.setName("Cotización");
    
    // Obtener rango origen
    var rangoOrigen = hojaOrigen.getRange(rangoDatos);
    Logger.log("📏 Rango origen obtenido: " + rangoDatos);
    
    // ✅ COPIAR VALORES
    var valores = rangoOrigen.getValues();
    var rangoDestino = hojaDestino.getRange(1, 1, valores.length, valores[0].length);
    rangoDestino.setValues(valores);
    Logger.log("✅ Valores copiados");
    
    // ✅ COPIAR FORMATOS (uno por uno para evitar conflictos entre archivos)
    try {
      // Colores de fondo
      var fondos = rangoOrigen.getBackgrounds();
      rangoDestino.setBackgrounds(fondos);
      Logger.log("✅ Fondos copiados");
      
      // Colores de fuente
      var coloresFuente = rangoOrigen.getFontColors();
      rangoDestino.setFontColors(coloresFuente);
      Logger.log("✅ Colores de fuente copiados");
      
      // Tamaños de fuente
      var tamanosFuente = rangoOrigen.getFontSizes();
      rangoDestino.setFontSizes(tamanosFuente);
      Logger.log("✅ Tamaños de fuente copiados");
      
      // Estilos de fuente (negrita, cursiva)
      var fontWeights = rangoOrigen.getFontWeights();
      rangoDestino.setFontWeights(fontWeights);
      
      var fontStyles = rangoOrigen.getFontStyles();
      rangoDestino.setFontStyles(fontStyles);
      Logger.log("✅ Estilos de fuente copiados");
      
      // Alineación horizontal
      var alineacionH = rangoOrigen.getHorizontalAlignments();
      rangoDestino.setHorizontalAlignments(alineacionH);
      
      // Alineación vertical  
      var alineacionV = rangoOrigen.getVerticalAlignments();
      rangoDestino.setVerticalAlignments(alineacionV);
      Logger.log("✅ Alineaciones copiadas");
      
      // Formatos de número
      var formatosNumero = rangoOrigen.getNumberFormats();
      rangoDestino.setNumberFormats(formatosNumero);
      Logger.log("✅ Formatos de número copiados");
      
    } catch (formatError) {
      Logger.log("⚠️ Error copiando algunos formatos: " + formatError.message);
      // Continuar aunque algunos formatos fallen
    }
    
    // ✅ APLICAR BORDES BÁSICOS
    try {
      // Bordes exteriores
      rangoDestino.setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      Logger.log("✅ Bordes aplicados");
    } catch (borderError) {
      Logger.log("⚠️ Error aplicando bordes: " + borderError.message);
    }
    
    // ✅ AJUSTAR COLUMNAS
    try {
      hojaDestino.autoResizeColumns(1, valores[0].length);
      Logger.log("✅ Columnas ajustadas automáticamente");
    } catch (resizeError) {
      Logger.log("⚠️ Error ajustando columnas: " + resizeError.message);
    }
    
    Logger.log("✅ Datos copiados exitosamente al nuevo archivo");
    
    return { exito: true };
    
  } catch (e) {
    Logger.log("❌ Error general copiando datos: " + e.message);
    Logger.log("📍 Stack trace: " + e.stack);
    return { exito: false, mensaje: "Error copiando datos: " + e.message };
  }
}

/**
 * Limpia la hoja Generador eliminando las filas desde A4 hacia abajo y desplazando hacia arriba
 */
function limpiarGeneradorConDesplazamiento(hoja, ultimaFila) {
  try {
    if (ultimaFila <= 3) {
      Logger.log("No hay filas para eliminar (última fila: " + ultimaFila + ")");
      return { exito: true };
    }

    // Calcular cuántas filas eliminar (desde fila 4 hasta ultimaFila)
    var filasAEliminar = ultimaFila - 3; // Restar 3 porque empezamos desde fila 4

    Logger.log("Eliminando " + filasAEliminar + " filas desde la fila 4");

    // Eliminar las filas (esto automáticamente desplaza hacia arriba)
    hoja.deleteRows(4, filasAEliminar);

    // ✅ IMPORTANTE: Restaurar la fórmula en K1 después de eliminar filas
    hoja.getRange("K1").setFormula('=MAX(FILTER(ROW(D:D), D:D<>""))');

    Logger.log("Filas eliminadas exitosamente y fórmula K1 restaurada");

    // ✅ Limpiar rango A4:I completo con formato de bordes blancos
    Logger.log("Limpiando rango A4:I con bordes blancos...");
    var rangoLimpieza = hoja.getRange("A4:I");

    // Limpiar contenido
    rangoLimpieza.clearContent();

    // Aplicar bordes blancos a todo el rango
    rangoLimpieza.setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

    Logger.log("✅ Rango A4:I limpiado con bordes blancos");

    return { exito: true };

  } catch (e) {
    return { exito: false, mensaje: "Error limpiando Generador: " + e.message };
  }
}

/**
 * Versión alternativa: Crea archivo en Mi Unidad y lo comparte con el cliente
 * NO requiere acceso a carpetas externas
 */
function transferirGeneradorAlternativo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    
    Logger.log("🚀 INICIANDO transferirGeneradorAlternativo (sin acceso a carpetas)...");
    
    // ✅ PASO 1: Obtener información del cliente
    Logger.log("📋 PASO 1: Obteniendo información del cliente...");
    var infoCliente = obtenerInformacionClienteAlternativo(hojaGenerador);
    if (!infoCliente.exito) {
      Logger.log("❌ FALLÓ PASO 1: " + infoCliente.mensaje);
      return infoCliente.mensaje;
    }
    Logger.log("✅ PASO 1 COMPLETADO");
    
    // ✅ PASO 2: Determinar rango de datos
    Logger.log("📋 PASO 2: Determinando rango de datos...");
    var ultimaFila = hojaGenerador.getRange("K1").getValue();
    if (!ultimaFila || ultimaFila < 4) {
      Logger.log("❌ FALLÓ PASO 2: No hay datos (K1 = " + ultimaFila + ")");
      return "Error: No hay datos para transferir (K1 = " + ultimaFila + ")";
    }
    
    var rangoDatos = "A1:M" + ultimaFila;
    Logger.log("Rango a copiar: " + rangoDatos);
    Logger.log("✅ PASO 2 COMPLETADO");
    
    // ✅ PASO 3: Crear archivo en Mi Unidad
    Logger.log("📋 PASO 3: Creando archivo en Mi Unidad...");
    var nuevoArchivo = crearArchivoEnMiUnidad(infoCliente.nombreCliente);
    if (!nuevoArchivo.exito) {
      Logger.log("❌ FALLÓ PASO 3: " + nuevoArchivo.mensaje);
      return nuevoArchivo.mensaje;
    }
    Logger.log("✅ PASO 3 COMPLETADO - Archivo creado: " + nuevoArchivo.url);
    
    // ✅ PASO 4: Copiar datos
    Logger.log("📋 PASO 4: Copiando datos...");
    var resultadoCopia = copiarDatosAlNuevoArchivo(hojaGenerador, nuevoArchivo.archivo, rangoDatos);
    if (!resultadoCopia.exito) {
      Logger.log("❌ FALLÓ PASO 4: " + resultadoCopia.mensaje);
      return resultadoCopia.mensaje;
    }
    Logger.log("✅ PASO 4 COMPLETADO");
    
    // ✅ PASO 5: Compartir con cliente (si hay email)
    if (infoCliente.emailCliente) {
      Logger.log("📋 PASO 5: Compartiendo con cliente...");
      var resultadoCompartir = compartirConCliente(nuevoArchivo.archivo, infoCliente.emailCliente);
      Logger.log("✅ PASO 5 COMPLETADO: " + resultadoCompartir);
    } else {
      Logger.log("⚠️ PASO 5 OMITIDO: No se encontró email del cliente");
    }
    
    // ✅ PASO 6: Limpiar Generador
    Logger.log("📋 PASO 6: Limpiando hoja Generador...");
    var resultadoLimpieza = limpiarGeneradorConDesplazamiento(hojaGenerador, ultimaFila);
    if (!resultadoLimpieza.exito) {
      Logger.log("❌ FALLÓ PASO 6: " + resultadoLimpieza.mensaje);
      return resultadoLimpieza.mensaje;
    }
    Logger.log("✅ PASO 6 COMPLETADO");
    
    Logger.log("🎉 TODO COMPLETADO EXITOSAMENTE");
    return "✅ Éxito: Archivo creado y compartido. URL: " + nuevoArchivo.url;
    
  } catch (e) {
    Logger.log("💥 ERROR GENERAL: " + e.toString());
    return "Error: " + e.message;
  }
}

/**
 * Obtiene información del cliente (versión simplificada)
 */
function obtenerInformacionClienteAlternativo(hoja) {
  try {
    // Obtener nombre del cliente desde F2 (sin el hipervínculo)
    var valorF2 = hoja.getRange("F2").getValue().toString().trim();
    var nombreCliente = valorF2.replace("Ver Carpeta", "").trim();
    
    if (!nombreCliente) {
      nombreCliente = "Cliente_" + Utilities.formatDate(new Date(), "GMT-6", "yyyyMMdd");
    }
    
    // Intentar obtener email desde otra celda (puedes modificar esto)
    var emailCliente = null;
    try {
      // Buscar email en F1, G2, o donde esté configurado
      var posiblesEmails = [
        hoja.getRange("F1").getValue().toString().trim(),
        hoja.getRange("G2").getValue().toString().trim(),
        hoja.getRange("E2").getValue().toString().trim()
      ];
      
      for (var i = 0; i < posiblesEmails.length; i++) {
        if (posiblesEmails[i].includes("@") && posiblesEmails[i].includes(".")) {
          emailCliente = posiblesEmails[i];
          break;
        }
      }
    } catch (e) {
      Logger.log("⚠️ No se pudo obtener email del cliente");
    }
    
    Logger.log("Cliente: " + nombreCliente + " | Email: " + (emailCliente || "No encontrado"));
    
    return {
      exito: true,
      nombreCliente: nombreCliente,
      emailCliente: emailCliente
    };
    
  } catch (e) {
    return { exito: false, mensaje: "Error obteniendo info cliente: " + e.message };
  }
}

/**
 * Crea archivo en Mi Unidad (sin mover a carpetas externas)
 */
function crearArchivoEnMiUnidad(nombreCliente) {
  try {
    var timestamp = Utilities.formatDate(new Date(), "GMT-6", "yyyyMMdd_HHmm");
    var nombreArchivo = "Cotización_" + nombreCliente + "_" + timestamp;
    
    Logger.log("📄 Creando archivo: " + nombreArchivo);
    
    var nuevoArchivo = SpreadsheetApp.create(nombreArchivo);
    Logger.log("✅ Archivo creado en Mi Unidad: " + nuevoArchivo.getId());
    
    return {
      exito: true,
      archivo: nuevoArchivo,
      url: nuevoArchivo.getUrl()
    };
    
  } catch (e) {
    Logger.log("❌ Error creando archivo: " + e.message);
    return { exito: false, mensaje: "Error creando archivo: " + e.message };
  }
}

/**
 * Comparte el archivo con el cliente
 */
function compartirConCliente(archivo, emailCliente) {
  try {
    var archivoId = archivo.getId();
    var archivoDrive = DriveApp.getFileById(archivoId);
    
    // Compartir con permisos de editor
    archivoDrive.addEditor(emailCliente);
    
    Logger.log("✅ Archivo compartido con: " + emailCliente);
    return "Compartido con " + emailCliente;
    
  } catch (e) {
    Logger.log("⚠️ No se pudo compartir: " + e.message);
    return "No se pudo compartir: " + e.message;
  }
}

/**
 * Función para solicitar permisos de Google Drive
 * EJECUTAR ESTA FUNCIÓN PRIMERO para autorizar permisos
 */
function autorizarPermisosDrive() {
  try {
    Logger.log("🔐 Solicitando permisos de Google Drive...");
    
    // Esta línea forzará la solicitud de permisos
    var carpetas = DriveApp.getFolders();
    
    // Crear un archivo de prueba para confirmar que funcionan los permisos
    var archivoTest = SpreadsheetApp.create("TEST_Permisos_" + new Date().getTime());
    var archivo = DriveApp.getFileById(archivoTest.getId());
    
    Logger.log("✅ Permisos otorgados correctamente");
    Logger.log("📄 Archivo de prueba creado: " + archivoTest.getUrl());
    
    // Borrar el archivo de prueba
    DriveApp.getFileById(archivoTest.getId()).setTrashed(true);
    Logger.log("🗑️ Archivo de prueba eliminado");
    
    return "✅ Permisos de Drive autorizados correctamente. Ahora puedes usar las funciones de transferencia.";
    
  } catch (e) {
    Logger.log("❌ Error autorizando permisos: " + e.message);
    return "❌ Error: " + e.message + "\n\nDebes autorizar manualmente los permisos en el editor de Apps Script.";
  }
}
 
function soloLimpiarGenerador() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    var ultimaFila = hojaGenerador.getRange("K1").getValue();
    
    var resultado = limpiarGeneradorConDesplazamiento(hojaGenerador, ultimaFila);
    return resultado.exito ? "✅ Generador limpiado" : resultado.mensaje;
    
  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Función de prueba para verificar acceso a la carpeta del cliente
 */
function probarAccesoCarpeta() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    
    Logger.log("🧪 PROBANDO ACCESO A CARPETA...");
    
    // Obtener información del cliente
    var infoCliente = obtenerInformacionCliente(hojaGenerador);
    if (!infoCliente.exito) {
      Logger.log("❌ No se pudo obtener info del cliente: " + infoCliente.mensaje);
      return infoCliente.mensaje;
    }
    
    Logger.log("✅ Info cliente obtenida: " + infoCliente.nombreCliente);
    Logger.log("📁 ID Carpeta: " + infoCliente.carpetaId);
    
    // Intentar acceder a la carpeta
    try {
      var carpeta = DriveApp.getFolderById(infoCliente.carpetaId);
      Logger.log("✅ Carpeta accesible: " + carpeta.getName());
      
      // Verificar permisos
      var archivos = carpeta.getFiles();
      var contador = 0;
      while (archivos.hasNext() && contador < 3) {
        var archivo = archivos.next();
        Logger.log("📄 Archivo en carpeta: " + archivo.getName());
        contador++;
      }
      
      return "✅ Carpeta accesible: " + carpeta.getName() + " (ID: " + infoCliente.carpetaId + ")";
      
    } catch (errorCarpeta) {
      Logger.log("❌ Error accediendo a carpeta: " + errorCarpeta.message);
      return "❌ No se puede acceder a la carpeta: " + errorCarpeta.message;
    }
    
  } catch (e) {
    Logger.log("💥 Error general: " + e.message);
    return "Error: " + e.message;
  }
}

/**
 * Función para crear archivo de prueba en Mi Unidad (sin mover a carpeta)
 */
function crearArchivoPrueba() {
  try {
    Logger.log("🧪 CREANDO ARCHIVO DE PRUEBA...");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    
    // Crear archivo de prueba
    var timestamp = Utilities.formatDate(new Date(), "GMT-6", "yyyyMMdd_HHmm");
    var nombreArchivo = "PRUEBA_Cotización_" + timestamp;
    
    var nuevoArchivo = SpreadsheetApp.create(nombreArchivo);
    Logger.log("✅ Archivo de prueba creado: " + nuevoArchivo.getUrl());
    
    // Copiar algunos datos
    var ultimaFila = hojaGenerador.getRange("K1").getValue();
    var rangoDatos = "A1:M" + Math.min(ultimaFila, 10); // Solo las primeras 10 filas como prueba
    
    var hojaDestino = nuevoArchivo.getActiveSheet();
    var rangoOrigen = hojaGenerador.getRange(rangoDatos);
    var valores = rangoOrigen.getValues();
    
    hojaDestino.getRange(1, 1, valores.length, valores[0].length).setValues(valores);
    
    Logger.log("✅ Datos copiados al archivo de prueba");
    
    return "✅ Archivo de prueba creado exitosamente: " + nuevoArchivo.getUrl();
    
  } catch (e) {
    Logger.log("❌ Error en prueba: " + e.message);
    return "Error en prueba: " + e.message;
  }
}