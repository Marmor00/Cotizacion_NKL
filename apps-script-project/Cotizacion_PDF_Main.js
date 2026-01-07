/**
 * Script principal para la generación de cotizaciones formales en PDF
 * Integra: Parser, Folios, WebApp y Generador de PDF
 */

// ID del spreadsheet que contiene datos compartidos (Personal, ListaProyectos, CONTACTOS)
var SPREADSHEET_DATOS_ID = "1noiFvtA5BXIQMVtY9amQbMGGrPu4DiXXZHxqRmpvi5U";

// ID del spreadsheet de contactos
var SPREADSHEET_CONTACTOS_ID = "1lI1brWvWN24cBjjoXs7qUWJIlUpN6VMQx9W-MSlT-P8";

/**
 * NOTA: La función onOpen() está ahora en Menu_Principal.js
 * para evitar duplicados. Este archivo ya no necesita onOpen().
 */
// function onOpen() {
//   SpreadsheetApp.getUi()
//       .createMenu('📄 Cotización PDF')
//       .addItem('✨ Generar Cotización Formal', 'abrirWebappCotizacion')
//       .addSeparator()
//       .addItem('🧪 Probar Parser', 'probarParser')
//       .addItem('🔢 Ver Último Folio', 'mostrarUltimoFolio')
//       .addToUi();
// }

/**
 * Función para forzar la autorización de todos los permisos necesarios
 * EJECUTAR ESTA FUNCIÓN PRIMERO desde el editor de Apps Script
 */
function autorizarPermisos() {
  // Forzar permisos de Spreadsheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("✅ Permiso de Spreadsheets autorizado");

  // Forzar permisos de Drive
  var folders = DriveApp.getFolders();
  Logger.log("✅ Permiso de Drive autorizado");

  // Forzar permisos de Documents (IMPORTANTE)
  var testDoc = DocumentApp.create("TEST_Permisos_" + new Date().getTime());
  var docId = testDoc.getId();
  DriveApp.getFileById(docId).setTrashed(true);
  Logger.log("✅ Permiso de Documents autorizado");

  Logger.log("🎉 TODOS LOS PERMISOS AUTORIZADOS CORRECTAMENTE");

  return "✅ Todos los permisos han sido autorizados. Ahora puedes usar la webapp sin problemas.";
}

/**
 * Función doGet() - Permite acceder a la webapp mediante URL
 * Esta es la función que se ejecuta cuando abres la webapp con su URL desplegada
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Cotizacion_PDF_Webapp')
      .setTitle('Generar Cotización Formal - NKL')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Abre la webapp para generar cotización (desde el menú de Sheets)
 */
function abrirWebappCotizacion() {
  var html = HtmlService.createHtmlOutputFromFile('Cotizacion_PDF_Webapp')
      .setWidth(950)
      .setHeight(700)
      .setTitle('Generar Cotización Formal');

  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Cotización Formal - NKL');
}

/**
 * Obtiene la lista de vendedores desde la hoja Personal
 * @returns {Array} - Array de objetos con datos de vendedores
 */
function obtenerVendedores() {
  try {
    Logger.log("🔍 Buscando vendedores en hoja Personal...");
    Logger.log("📂 Abriendo spreadsheet de datos: " + SPREADSHEET_DATOS_ID);

    var ss = SpreadsheetApp.openById(SPREADSHEET_DATOS_ID);

    var hojaPersonal = ss.getSheetByName("Personal");

    if (!hojaPersonal) {
      Logger.log("⚠️ No existe la hoja Personal en el spreadsheet");
      return [];
    }

    Logger.log("✅ Hoja Personal encontrada");

    // Leer datos desde el rango L:O (columnas 12-15)
    var ultimaFila = hojaPersonal.getLastRow();
    Logger.log("📊 Última fila en Personal: " + ultimaFila);

    if (ultimaFila < 2) {
      Logger.log("⚠️ No hay datos de personal (última fila < 2)");
      return [];
    }

    // Leer desde fila 2 (asumiendo que fila 1 tiene encabezados)
    var numFilas = ultimaFila - 1;
    Logger.log("📋 Leyendo " + numFilas + " filas desde L2:O" + ultimaFila);

    var datos = hojaPersonal.getRange(2, 12, numFilas, 4).getValues();
    Logger.log("✅ Datos leídos: " + datos.length + " filas");

    var vendedores = [];

    for (var i = 0; i < datos.length; i++) {
      Logger.log("  Fila " + (i+2) + ": [" + datos[i].join(", ") + "]");

      // Solo agregar si tiene nombre
      if (datos[i][0] && String(datos[i][0]).trim() !== "") {
        var vendedor = {
          nombre: String(datos[i][0] || "").trim(),
          puesto: String(datos[i][1] || "").trim(),
          correo: String(datos[i][2] || "").trim(),
          celular: String(datos[i][3] || "").trim()
        };
        vendedores.push(vendedor);
        Logger.log("    ✅ Vendedor agregado: " + vendedor.nombre);
      } else {
        Logger.log("    ⏭️ Fila saltada (sin nombre)");
      }
    }

    Logger.log("✅ TOTAL: " + vendedores.length + " vendedores cargados");
    return vendedores;

  } catch (e) {
    Logger.log("❌ Error en obtenerVendedores: " + e.message);
    Logger.log("📍 Stack: " + e.stack);
    return [];
  }
}

/**
 * Guardar contacto en la hoja CONTACTOS
 * @param {Object} datosCliente - Datos del cliente a guardar
 * @return {Object} - Resultado de la operación
 */
function guardarContacto(datosCliente) {
  try {
    Logger.log("💾 Guardando contacto: " + datosCliente.nombre);

    var ss = SpreadsheetApp.openById(SPREADSHEET_CONTACTOS_ID);
    var hojaContactos = ss.getSheetByName("CONTACTOS");

    if (!hojaContactos) {
      Logger.log("⚠️ No existe la hoja CONTACTOS");
      return {
        exito: false,
        mensaje: "No se encontró la hoja CONTACTOS"
      };
    }

    // Verificar si el contacto ya existe (por nombre o email)
    var ultimaFila = hojaContactos.getLastRow();
    if (ultimaFila > 1) {
      var datosExistentes = hojaContactos.getRange(2, 1, ultimaFila - 1, 3).getValues();

      for (var i = 0; i < datosExistentes.length; i++) {
        var nombreExistente = String(datosExistentes[i][0] || "").trim().toLowerCase();
        var emailExistente = String(datosExistentes[i][2] || "").trim().toLowerCase();
        var nombreNuevo = String(datosCliente.nombre || "").trim().toLowerCase();
        var emailNuevo = String(datosCliente.email || "").trim().toLowerCase();

        if (nombreExistente === nombreNuevo || (emailNuevo && emailExistente === emailNuevo)) {
          Logger.log("⚠️ Contacto ya existe");
          return {
            exito: false,
            mensaje: "Este contacto ya existe en la base de datos"
          };
        }
      }
    }

    // Agregar nuevo contacto
    var nuevaFila = ultimaFila + 1;
    hojaContactos.getRange(nuevaFila, 1).setValue(datosCliente.nombre || "");
    hojaContactos.getRange(nuevaFila, 2).setValue(datosCliente.rfc || "");
    hojaContactos.getRange(nuevaFila, 3).setValue(datosCliente.email || "");
    hojaContactos.getRange(nuevaFila, 4).setValue(datosCliente.telefono || "");
    hojaContactos.getRange(nuevaFila, 5).setValue(datosCliente.domicilioFiscal || "");
    hojaContactos.getRange(nuevaFila, 6).setValue(datosCliente.domicilioEntrega || "");
    hojaContactos.getRange(nuevaFila, 7).setValue(datosCliente.razonSocial || "");

    Logger.log("✅ Contacto guardado exitosamente en fila " + nuevaFila);
    return {
      exito: true,
      mensaje: "Contacto guardado exitosamente"
    };

  } catch (e) {
    Logger.log("❌ Error al guardar contacto: " + e.message);
    return {
      exito: false,
      mensaje: "Error al guardar: " + e.message
    };
  }
}

/**
 * Buscar proyecto por número de orden en ListaProyectos
 * @param {string} numeroOrden - Número de orden (ej: "O-3000")
 * @return {Object|null} - Datos del proyecto o null si no se encuentra
 */
function buscarProyectoPorOrden(numeroOrden) {
  try {
    Logger.log("🔍 Buscando proyecto con orden: " + numeroOrden);

    var ss = SpreadsheetApp.openById(SPREADSHEET_DATOS_ID);
    var hojaProyectos = ss.getSheetByName("ListaProyectos");

    if (!hojaProyectos) {
      Logger.log("⚠️ No existe la hoja ListaProyectos");
      return null;
    }

    var ultimaFila = hojaProyectos.getLastRow();
    if (ultimaFila < 2) {
      Logger.log("⚠️ La hoja ListaProyectos está vacía");
      return null;
    }

    var numFilas = ultimaFila - 1;
    // Columnas S:T = columnas 19:20 (orden, nombre)
    var datos = hojaProyectos.getRange(2, 19, numFilas, 2).getValues();

    Logger.log("📋 Buscando en " + numFilas + " proyectos...");

    for (var i = 0; i < datos.length; i++) {
      var ordenActual = String(datos[i][0] || "").trim();

      if (ordenActual === numeroOrden) {
        Logger.log("✅ Proyecto encontrado: " + datos[i][1]);
        return {
          orden: ordenActual,
          nombre: String(datos[i][1] || "").trim()
        };
      }
    }

    Logger.log("⚠️ No se encontró proyecto con orden: " + numeroOrden);
    return null;

  } catch (e) {
    Logger.log("❌ Error en buscarProyectoPorOrden: " + e.message);
    return null;
  }
}

/**
 * FUNCIÓN DE PRUEBA: Ver datos de ListaProyectos
 */
function probarListaProyectos() {
  Logger.log("=".repeat(50));
  Logger.log("🧪 PRUEBA DE LISTA PROYECTOS");
  Logger.log("=".repeat(50));

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_DATOS_ID);
    var hojaProyectos = ss.getSheetByName("ListaProyectos");

    if (!hojaProyectos) {
      Logger.log("⚠️ No existe la hoja ListaProyectos");
      return;
    }

    var ultimaFila = hojaProyectos.getLastRow();
    Logger.log("📊 Última fila: " + ultimaFila);

    // Leer columnas S:T (orden, nombre)
    var numFilas = Math.min(ultimaFila - 1, 20); // Mostrar primeras 20 filas
    var datos = hojaProyectos.getRange(2, 19, numFilas, 2).getValues();

    Logger.log("📋 Primeras " + numFilas + " filas (columnas S:T - Orden, Nombre):");
    for (var i = 0; i < datos.length; i++) {
      Logger.log("  Fila " + (i+2) + ": Orden=[" + datos[i][0] + "] Nombre=[" + datos[i][1] + "]");
    }

    // Buscar específicamente O-3373 y O-3733
    Logger.log("\n🔍 Buscando O-3373 y O-3733 específicamente...");
    var datosCompletos = hojaProyectos.getRange(2, 19, ultimaFila - 1, 2).getValues();
    for (var i = 0; i < datosCompletos.length; i++) {
      var orden = String(datosCompletos[i][0] || "").trim();
      if (orden.includes("3373") || orden.includes("3733")) {
        Logger.log("✅ ENCONTRADO en fila " + (i+2) + ": Orden=[" + orden + "] Nombre=[" + datosCompletos[i][1] + "]");
        Logger.log("   Longitud del string: " + orden.length);
        Logger.log("   Códigos de caracteres: " + Array.from(orden).map(function(c) { return c.charCodeAt(0); }).join(", "));
      }
    }

  } catch (e) {
    Logger.log("❌ Error: " + e.message);
  }

  Logger.log("=".repeat(50));
}

/**
 * FUNCIÓN DE PRUEBA: Verificar carga de vendedores
 */
function probarVendedores() {
  Logger.log("=".repeat(50));
  Logger.log("🧪 PRUEBA DE VENDEDORES");
  Logger.log("=".repeat(50));

  var vendedores = obtenerVendedores();

  Logger.log("=".repeat(50));
  Logger.log("📊 RESULTADO:");
  Logger.log("  Total vendedores: " + vendedores.length);
  Logger.log(JSON.stringify(vendedores, null, 2));
  Logger.log("=".repeat(50));

  SpreadsheetApp.getUi().alert(
    'Vendedores encontrados',
    'Se encontraron ' + vendedores.length + ' vendedores.\n\nRevisa el log (Ctrl+Enter) para ver los detalles.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Función principal que genera el PDF de cotización
 * Llamada desde la webapp
 */
function generarPDFCotizacion(datosCliente, datosCotizacion, productosSeleccionados, modoManual) {
  try {
    Logger.log("🚀 Iniciando generación de PDF de cotización...");
    Logger.log("📋 Modo manual: " + (modoManual ? "SÍ" : "NO"));

    // PASO 1: Generar folio único
    Logger.log("📋 PASO 1: Generando folio único...");
    var folio = generarFolioUnico();
    Logger.log("✅ Folio generado: " + folio);

    // PASO 2 y 3: Obtener productos (desde Generador o desde webapp si es manual)
    var productosParaPDF = [];

    if (modoManual) {
      // MODO MANUAL: Usar productos enviados directamente desde la webapp
      Logger.log("📋 PASO 2-3: Usando productos manuales desde webapp...");

      for (var i = 0; i < productosSeleccionados.length; i++) {
        var sel = productosSeleccionados[i];

        // En modo manual, el producto completo viene en sel.producto
        if (sel.producto) {
          var producto = sel.producto;

          // Aplicar código editado y descuentos
          producto.codigoEditado = sel.codigo;
          producto.descuentoPorcentaje = sel.descuentoPorcentaje || 0;
          producto.descuentoPesos = sel.descuentoPesos || 0;
          producto.importeFinal = sel.importeFinal || producto.importe;
          productosParaPDF.push(producto);
        }
      }

      Logger.log("✅ Productos manuales procesados: " + productosParaPDF.length);

    } else {
      // MODO NORMAL: Parsear productos de Generador
      Logger.log("📋 PASO 2: Parseando productos desde Generador...");
      var resultadoParser = parsearHojaGenerador();

      if (!resultadoParser.exito) {
        return {
          exito: false,
          mensaje: "Error al parsear productos: " + resultadoParser.mensaje
        };
      }

      Logger.log("✅ Productos parseados: " + resultadoParser.productos.length);

      // PASO 3: Filtrar solo los productos seleccionados y aplicar códigos editados y descuentos
      Logger.log("📋 PASO 3: Filtrando productos seleccionados...");

      for (var i = 0; i < productosSeleccionados.length; i++) {
        var sel = productosSeleccionados[i];
        var producto = resultadoParser.productos[sel.index];

        if (producto) {
          // Aplicar código editado y descuentos
          producto.codigoEditado = sel.codigo;
          producto.descuentoPorcentaje = sel.descuentoPorcentaje || 0;
          producto.descuentoPesos = sel.descuentoPesos || 0;
          producto.importeFinal = sel.importeFinal || producto.importe;
          productosParaPDF.push(producto);
        }
      }

      Logger.log("✅ Productos filtrados: " + productosParaPDF.length);
    }

    // PASO 4: Verificar si es Modo B (Precio Cerrado)
    var modoPrecioCerrado = datosCotizacion.modoPrecioCerrado || false;
    var urlPDFCliente = null;
    var urlPDFInterno = null;
    var totales = null;
    var productosConDescuento = null;

    if (modoPrecioCerrado) {
      Logger.log("💰 MODO B - PRECIO CERRADO ACTIVADO");
      Logger.log("📋 Se generarán 2 PDFs: Cliente (sin desc.) e Interno (con desc. 13.79%)");

      // PASO 4A: Generar PDF para el CLIENTE (sin descuentos)
      Logger.log("📋 PASO 4A: Calculando totales SIN descuentos (PDF Cliente)...");
      var totalesSinDescuento = calcularTotales(productosParaPDF, false); // false = sin descuentos
      Logger.log("✅ Subtotal Cliente: $" + totalesSinDescuento.subtotal.toFixed(2));
      Logger.log("✅ Total Cliente: $" + totalesSinDescuento.total.toFixed(2));

      Logger.log("📋 Generando PDF Cliente (sin descuentos)...");
      urlPDFCliente = generarDocumentoDesdeTemplate(
        folio,
        datosCliente,
        datosCotizacion,
        productosParaPDF,
        totalesSinDescuento,
        "_Cliente" // Sufijo para el nombre
      );

      if (!urlPDFCliente) {
        return {
          exito: false,
          mensaje: "Error al generar el PDF Cliente"
        };
      }
      Logger.log("✅ PDF Cliente generado: " + urlPDFCliente);

      // PASO 4B: Hacer copia profunda de productos para aplicar descuentos
      Logger.log("📋 PASO 4B: Copiando productos para aplicar descuentos...");
      productosConDescuento = JSON.parse(JSON.stringify(productosParaPDF));

      // PASO 4C: Calcular totales CON descuentos (PDF Interno)
      Logger.log("📋 PASO 4C: Calculando totales CON descuentos 13.79% (PDF Interno)...");
      totales = calcularTotales(productosConDescuento, true); // true = con descuentos
      Logger.log("✅ Subtotal Interno: $" + totales.subtotal.toFixed(2));
      Logger.log("✅ Total Interno: $" + totales.total.toFixed(2));

      Logger.log("📋 Generando PDF Interno (con descuentos)...");
      urlPDFInterno = generarDocumentoDesdeTemplate(
        folio,
        datosCliente,
        datosCotizacion,
        productosConDescuento,
        totales,
        "_Interna" // Sufijo para el nombre
      );

      if (!urlPDFInterno) {
        return {
          exito: false,
          mensaje: "Error al generar el PDF Interno"
        };
      }
      Logger.log("✅ PDF Interno generado: " + urlPDFInterno);

    } else {
      // MODO NORMAL: Solo 1 PDF
      Logger.log("📋 PASO 4: Calculando totales (modo normal)...");
      totales = calcularTotales(productosParaPDF, false);
      Logger.log("✅ Subtotal: $" + totales.subtotal.toFixed(2));

      Logger.log("📋 PASO 5: Generando documento desde template...");
      urlPDFInterno = generarDocumentoDesdeTemplate(
        folio,
        datosCliente,
        datosCotizacion,
        productosParaPDF,
        totales,
        "" // Sin sufijo
      );

      if (!urlPDFInterno) {
        return {
          exito: false,
          mensaje: "Error al generar el documento PDF"
        };
      }
      Logger.log("✅ PDF generado: " + urlPDFInterno);
    }

    // PASO 6: Registrar cotización en la hoja Cotizaciones
    Logger.log("📋 PASO 6: Registrando cotización...");
    var productosInterno = modoPrecioCerrado ? productosConDescuento : productosParaPDF;
    var productosCliente = modoPrecioCerrado ? productosParaPDF : null; // Productos SIN descuentos

    registrarCotizacion(
      folio,
      datosCliente,
      datosCotizacion,
      productosInterno,  // Productos con descuentos (o normales si no es Modo B)
      totales,
      urlPDFInterno,
      urlPDFCliente,     // URL del PDF cliente (si existe)
      productosCliente   // Productos sin descuentos (solo en Modo B)
    );
    Logger.log("✅ Cotización registrada");

    return {
      exito: true,
      folio: folio,
      url: urlPDFInterno,  // URL principal (la interna)
      urlCliente: urlPDFCliente,  // URL del cliente (si existe)
      mensaje: modoPrecioCerrado
        ? "PDFs generados exitosamente (Cliente e Interna)"
        : "PDF generado exitosamente",
      datosCliente: datosCliente
    };

  } catch (e) {
    Logger.log("❌ ERROR en generarPDFCotizacion: " + e.message);
    Logger.log("📍 Stack: " + e.stack);

    return {
      exito: false,
      mensaje: "Error general: " + e.message
    };
  }
}

/**
 * Genera el documento PDF desde un template de Google Docs
 * @param {String} sufijo - Sufijo opcional para el nombre del archivo (ej: "_Cliente", "_Interna")
 * @returns {String} - URL del PDF generado
 */
function generarDocumentoDesdeTemplate(folio, datosCliente, datosCotizacion, productos, totales, sufijo) {
  try {
    sufijo = sufijo || ""; // Si no se proporciona sufijo, usar cadena vacía
    var nombreArchivo = "Cotización_" + folio + sufijo + "_" + datosCliente.nombre;

    // Crear nuevo Google Doc
    var doc = DocumentApp.create(nombreArchivo);
    var body = doc.getBody();

    // Configurar márgenes y formato más compactos
    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(50);
    body.setMarginRight(50);

    // ENCABEZADO DE LA EMPRESA
    var headerTable = body.appendTable();
    var headerRow = headerTable.appendTableRow();

    var cellEmpresa = headerRow.appendTableCell();
    var parrafoEmpresa = cellEmpresa.appendParagraph("Nueva Krystalum Lomelí SA de CV");
    parrafoEmpresa.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    parrafoEmpresa.editAsText().setFontSize(12).setBold(true).setForegroundColor("#595959");

    cellEmpresa.appendParagraph("NKL040223FP3").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("Avenida Fray Antonio Alcalde 2418").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("Col.Jardines Alcalde CP 44298").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("Guadalajara, JALISCO").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("Tel.(333) 853-7583").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("contacto@nkrystalum.mx").editAsText().setFontSize(8);

    // Agregar logo desde Google Drive
    var cellLogo = headerRow.appendTableCell();
    cellLogo.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    try {
      var logoId = "1u47sWEv4LYocFqGXNXFdS9M1Vhn_RnqP";
      var logoFile = DriveApp.getFileById(logoId);
      var logoBlob = logoFile.getBlob();
      var logoImage = cellLogo.appendParagraph("").appendInlineImage(logoBlob);
      // Logo horizontal (1200x351) - proporción 3.42:1
      // Establecer tanto ancho como altura manteniendo proporción
      logoImage.setWidth(171);  // 171px de ancho
      logoImage.setHeight(50);  // 50px de alto = proporción 3.42:1
      cellLogo.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    } catch(e) {
      Logger.log("⚠️ No se pudo cargar el logo: " + e.message);
      cellLogo.appendParagraph("").setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }

    headerTable.setBorderWidth(0);
    headerRow.getCell(0).setWidth(330);
    headerRow.getCell(1).setWidth(150);

    body.appendHorizontalRule();

    // TÍTULO
    var fecha = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");
    var tituloParrafo = body.appendParagraph("COTIZACIÓN " + folio);
    tituloParrafo.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    tituloParrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    tituloParrafo.editAsText().setFontSize(14).setBold(true).setForegroundColor("#595959");

    body.appendParagraph(""); // Espacio

    // DATOS DEL CLIENTE Y COTIZACIÓN (formato compacto)
    var infoTable = body.appendTable();

    // Helper function para aplicar estilos
    function aplicarEstiloLabel(cell) {
      cell.getChild(0).asParagraph().editAsText().setFontSize(8).setBold(true).setForegroundColor("#424242");
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setPaddingLeft(4);
      cell.setPaddingRight(4);
      cell.setBackgroundColor("#f5f5f5");
    }

    function aplicarEstiloValor(cell) {
      cell.getChild(0).asParagraph().editAsText().setFontSize(8);
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setPaddingLeft(4);
      cell.setPaddingRight(4);
    }

    // Fila 1: Cliente y Fecha
    var row1 = infoTable.appendTableRow();
    var c1_1 = row1.appendTableCell("CLIENTE:");
    aplicarEstiloLabel(c1_1);
    c1_1.setWidth(70);
    var c1_2 = row1.appendTableCell(datosCliente.nombre || "");
    aplicarEstiloValor(c1_2);
    c1_2.setWidth(190);
    var c1_3 = row1.appendTableCell("FECHA:");
    aplicarEstiloLabel(c1_3);
    c1_3.setWidth(70);
    var c1_4 = row1.appendTableCell(fecha);
    aplicarEstiloValor(c1_4);
    c1_4.setWidth(150);

    // Fila 2: RFC y ID Proyecto
    var row2 = infoTable.appendTableRow();
    var c2_1 = row2.appendTableCell("RFC:");
    aplicarEstiloLabel(c2_1);
    var c2_2 = row2.appendTableCell(datosCliente.rfc || "");
    aplicarEstiloValor(c2_2);
    var c2_3 = row2.appendTableCell("ID PROYECTO:");
    aplicarEstiloLabel(c2_3);
    var c2_4 = row2.appendTableCell(datosCotizacion.idObra || "NINGUNO");
    aplicarEstiloValor(c2_4);

    // Fila 3: Razón Social y Vendedor
    var row3 = infoTable.appendTableRow();
    var c3_1 = row3.appendTableCell("RAZÓN SOCIAL:");
    aplicarEstiloLabel(c3_1);
    var c3_2 = row3.appendTableCell(datosCliente.razonSocial || "");
    aplicarEstiloValor(c3_2);
    var c3_3 = row3.appendTableCell("VENDEDOR:");
    aplicarEstiloLabel(c3_3);
    var c3_4 = row3.appendTableCell(datosCotizacion.vendedor || "");
    aplicarEstiloValor(c3_4);

    // Fila 4: Teléfono y Correo Vendedor
    var row4 = infoTable.appendTableRow();
    var c4_1 = row4.appendTableCell("TELÉFONO:");
    aplicarEstiloLabel(c4_1);
    var c4_2 = row4.appendTableCell(datosCliente.telefono || "");
    aplicarEstiloValor(c4_2);
    var c4_3 = row4.appendTableCell("CORREO VEND.:");
    aplicarEstiloLabel(c4_3);
    var c4_4 = row4.appendTableCell(datosCotizacion.vendedorCorreo || "");
    aplicarEstiloValor(c4_4);

    // Fila 5: Email y Celular Vendedor
    var row5 = infoTable.appendTableRow();
    var c5_1 = row5.appendTableCell("EMAIL:");
    aplicarEstiloLabel(c5_1);
    var c5_2 = row5.appendTableCell(datosCliente.email || "");
    aplicarEstiloValor(c5_2);
    var c5_3 = row5.appendTableCell("CELULAR VEND.:");
    aplicarEstiloLabel(c5_3);
    var c5_4 = row5.appendTableCell(datosCotizacion.vendedorCelular || "");
    aplicarEstiloValor(c5_4);

    // Fila 6: Domicilio fiscal y Vigencia
    var row6 = infoTable.appendTableRow();
    var c6_1 = row6.appendTableCell("DOM. FISCAL:");
    aplicarEstiloLabel(c6_1);
    var c6_2 = row6.appendTableCell(datosCliente.domicilioFiscal || "SIN DEFINIR");
    aplicarEstiloValor(c6_2);
    var c6_3 = row6.appendTableCell("VIGENCIA:");
    aplicarEstiloLabel(c6_3);
    var c6_4 = row6.appendTableCell(datosCotizacion.vigencia + " días");
    aplicarEstiloValor(c6_4);

    // Fila 7: Domicilio entrega y Tiempo de entrega
    var row7 = infoTable.appendTableRow();
    var c7_1 = row7.appendTableCell("DOM. ENTREGA:");
    aplicarEstiloLabel(c7_1);
    var c7_2 = row7.appendTableCell(datosCliente.domicilioEntrega || "SIN DEFINIR");
    aplicarEstiloValor(c7_2);
    var c7_3 = row7.appendTableCell("T. ENTREGA:");
    aplicarEstiloLabel(c7_3);
    var c7_4 = row7.appendTableCell(datosCotizacion.tiempoEntrega + " días");
    aplicarEstiloValor(c7_4);

    // Fila 8: Moneda (solo 2 columnas)
    var row8 = infoTable.appendTableRow();
    var c8_1 = row8.appendTableCell("MONEDA:");
    aplicarEstiloLabel(c8_1);
    var c8_2 = row8.appendTableCell("MXN");
    aplicarEstiloValor(c8_2);
    var c8_3 = row8.appendTableCell("");
    aplicarEstiloLabel(c8_3);
    var c8_4 = row8.appendTableCell("");
    aplicarEstiloValor(c8_4);

    infoTable.setBorderWidth(0.5);
    infoTable.setBorderColor("#cccccc");

    body.appendParagraph(""); // Espacio

    // TABLA DE PRODUCTOS (formato compacto y elegante)
    var productosTable = body.appendTable();

    // Encabezado de tabla
    var headerProductos = productosTable.appendTableRow();
    headerProductos.appendTableCell("No.").setWidth(25);
    headerProductos.appendTableCell("Código").setWidth(55);
    headerProductos.appendTableCell("Descripción").setWidth(150);
    headerProductos.appendTableCell("Cant.").setWidth(35);
    headerProductos.appendTableCell("P.U.").setWidth(55);
    headerProductos.appendTableCell("% Desc.").setWidth(35);
    headerProductos.appendTableCell("$ Desc.").setWidth(55);
    headerProductos.appendTableCell("Importe").setWidth(55);

    // Hacer encabezado elegante
    for (var i = 0; i < headerProductos.getNumCells(); i++) {
      var cell = headerProductos.getCell(i);
      cell.getChild(0).asParagraph().editAsText().setBold(true).setFontSize(8).setForegroundColor("#ffffff");
      cell.setBackgroundColor("#595959");
      cell.setPaddingTop(3);
      cell.setPaddingBottom(3);
      cell.setPaddingLeft(2);
      cell.setPaddingRight(2);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    // Filas de productos
    for (var i = 0; i < productos.length; i++) {
      var prod = productos[i];

      var rowProd = productosTable.appendTableRow();

      // Alternar color de fondo
      var bgColor = i % 2 === 0 ? "#ffffff" : "#f9f9f9";

      var cellNo = rowProd.appendTableCell(String(i + 1));
      var cellCodigo = rowProd.appendTableCell(prod.codigoEditado || "SUM E INS");
      var cellDesc = rowProd.appendTableCell(prod.descripcion || "");

      var cantidad = prod.piezas || 1;
      var cellCant = rowProd.appendTableCell(String(cantidad));

      var precioUnitario = prod.precioUnitario || prod.precioVenta || 0;
      var precioFormateado = typeof precioUnitario === 'number'
        ? precioUnitario.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
        : String(precioUnitario);
      var cellPU = rowProd.appendTableCell(precioFormateado);

      // Descuento porcentaje
      var descPorcentaje = prod.descuentoPorcentaje || 0;
      var cellDescPorc = rowProd.appendTableCell(descPorcentaje.toFixed(2) + "%");

      // Descuento en pesos (puede venir como descuentoPesos o descuentoMonto)
      var descPesos = prod.descuentoPesos || prod.descuentoMonto || 0;
      var descPesosFormateado = typeof descPesos === 'number'
        ? descPesos.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
        : String(descPesos);
      var cellDescPesos = rowProd.appendTableCell(descPesosFormateado);

      // Importe final
      var importeFinal = prod.importeFinal || prod.importe || 0;
      var importeFormateado = typeof importeFinal === 'number'
        ? importeFinal.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
        : String(importeFinal);
      var cellImporte = rowProd.appendTableCell(importeFormateado);

      // Aplicar estilos a todas las celdas
      var cells = [cellNo, cellCodigo, cellDesc, cellCant, cellPU, cellDescPorc, cellDescPesos, cellImporte];
      for (var j = 0; j < cells.length; j++) {
        cells[j].getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor("#000000");
        cells[j].setPaddingTop(1.5);
        cells[j].setPaddingBottom(1.5);
        cells[j].setPaddingLeft(2);
        cells[j].setPaddingRight(2);
        cells[j].setBackgroundColor(bgColor);
      }

      // Alinear columnas numéricas a la derecha
      cellNo.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cellCant.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cellPU.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      cellDescPorc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cellDescPesos.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      cellImporte.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }

    productosTable.setBorderWidth(0.5);
    productosTable.setBorderColor("#cccccc");

    body.appendParagraph(""); // Espacio

    // TOTALES (formato elegante)
    var totalesTable = body.appendTable();
    totalesTable.setBorderWidth(0);

    // Fila Subtotal
    var rowSubtotal = totalesTable.appendTableRow();
    var cellLabelSub = rowSubtotal.appendTableCell("Subtotal:");
    cellLabelSub.setWidth(400);
    cellLabelSub.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellLabelSub.getChild(0).asParagraph().editAsText().setFontSize(10).setBold(true);
    var cellValorSub = rowSubtotal.appendTableCell("$" + totales.subtotal.toLocaleString('es-MX', {minimumFractionDigits: 2}));
    cellValorSub.setWidth(100);
    cellValorSub.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellValorSub.getChild(0).asParagraph().editAsText().setFontSize(10);

    // Fila IVA
    var rowIVA = totalesTable.appendTableRow();
    var cellLabelIVA = rowIVA.appendTableCell("IVA 16%:");
    cellLabelIVA.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellLabelIVA.getChild(0).asParagraph().editAsText().setFontSize(10).setBold(true);
    var cellValorIVA = rowIVA.appendTableCell("$" + totales.iva.toLocaleString('es-MX', {minimumFractionDigits: 2}));
    cellValorIVA.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellValorIVA.getChild(0).asParagraph().editAsText().setFontSize(10);

    // Fila Total
    var rowTotal = totalesTable.appendTableRow();
    var cellLabelTotal = rowTotal.appendTableCell("TOTAL:");
    cellLabelTotal.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellLabelTotal.getChild(0).asParagraph().editAsText().setFontSize(12).setBold(true).setForegroundColor("#595959");
    cellLabelTotal.setBackgroundColor("#f0f0f0");
    var cellValorTotal = rowTotal.appendTableCell("$" + totales.total.toLocaleString('es-MX', {minimumFractionDigits: 2}));
    cellValorTotal.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellValorTotal.getChild(0).asParagraph().editAsText().setFontSize(12).setBold(true).setForegroundColor("#595959");
    cellValorTotal.setBackgroundColor("#f0f0f0");

    body.appendParagraph(""); // Espacio

    // Total en letras
    var totalEnLetras = numeroALetras(totales.total);
    var parrafoLetras = body.appendParagraph("SON: " + totalEnLetras);
    parrafoLetras.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    parrafoLetras.editAsText().setItalic(true).setFontSize(9).setBold(true);

    body.appendParagraph(""); // Espacio
    body.appendParagraph(""); // Espacio
    body.appendParagraph(""); // Espacio

    // SECCIÓN DE FIRMAS
    var tablaFirmas = body.appendTable();
    var filaFirmas = tablaFirmas.appendTableRow();

    // Firma del vendedor
    var celdaVendedor = filaFirmas.appendTableCell();
    celdaVendedor.appendParagraph("").appendHorizontalRule();
    var nombreVendedor = celdaVendedor.appendParagraph("Firma del Vendedor");
    nombreVendedor.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    nombreVendedor.editAsText().setFontSize(9).setBold(true);
    var datoVendedor = celdaVendedor.appendParagraph(datosCotizacion.vendedor || "");
    datoVendedor.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    datoVendedor.editAsText().setFontSize(8);
    celdaVendedor.setPaddingTop(10);
    celdaVendedor.setPaddingBottom(10);

    // Firma del cliente
    var celdaCliente = filaFirmas.appendTableCell();
    celdaCliente.appendParagraph("").appendHorizontalRule();
    var labelCliente = celdaCliente.appendParagraph("Firma del Cliente");
    labelCliente.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    labelCliente.editAsText().setFontSize(9).setBold(true);
    var autorizoLabel = celdaCliente.appendParagraph("Autorizo: " + (datosCliente.nombre || ""));
    autorizoLabel.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    autorizoLabel.editAsText().setFontSize(8);
    celdaCliente.setPaddingTop(10);
    celdaCliente.setPaddingBottom(10);

    tablaFirmas.setBorderWidth(0);
    filaFirmas.getCell(0).setWidth(240);
    filaFirmas.getCell(1).setWidth(240);

    // TÉRMINOS Y CONDICIONES
    body.appendParagraph(""); // Espacio
    body.appendParagraph(""); // Espacio

    var tituloTerminos = body.appendParagraph("TÉRMINOS Y CONDICIONES");
    tituloTerminos.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    tituloTerminos.editAsText().setBold(true).setFontSize(12).setForegroundColor("#595959");

    body.appendParagraph(""); // Espacio

    // Leer términos y condiciones del archivo
    var terminos = [
      "1. Cotizaciones y Contratación",
      "- Las cotizaciones tienen validez en tanto no exista incremento en materiales.",
      "- Se requiere firma, nombre del cliente y fecha en la cotización o contrato.",
      "",
      "2. Precios, Pagos y Devoluciones",
      "- Precios: Los precios no incluyen IVA (se desglosa al final, en la cotización).",
      "- Anticipo y pagos: Se requiere un anticipo del 60% para iniciar trabajos. El saldo se cubrirá mediante pagos parciales por avance de obra y/o contra entrega.",
      "- Estimaciones: Se realizarán con base en el avance real de obra, y su frecuencia será semanal, salvo que el cliente y el vendedor acuerden un periodo distinto. El pago de cada estimación deberá efectuarse en un plazo máximo de dos días hábiles después de su emisión.",
      "  El objetivo principal de una estimación es que el cliente pague únicamente por el avance real de lo instalado en obra.",
      "  A modo de ejemplo, el procedimiento podrá desarrollarse de la siguiente manera:",
      "  • El anticipo del 60% se prorratea proporcionalmente entre todas las partidas; este anticipo se destina a la compra de materiales y garantiza el precio contratado, por lo que se considera aplicado desde el inicio.",
      "  • Conforme avanza la instalación, se cobra el porcentaje correspondiente al progreso físico de cada partida. Si una partida se instala al 100%, se factura el porcentaje restante de esa partida. Si una partida queda parcialmente instalada, se cobra solo el porcentaje equivalente al avance ejecutado, dejando el saldo pendiente para la entrega final.",
      "  Este esquema busca mantener una relación justa y transparente entre el avance físico y los pagos, permitiendo ajustes flexibles conforme a los acuerdos específicos con el vendedor.",
      "- Devoluciones: Las devoluciones deberán ser autorizadas por la Dirección Comercial. En caso de proceder, se aplicará un cargo mínimo del 20% sobre el valor del producto o servicio, además de los costos de materiales y mano de obra utilizados. Los productos fabricados a la medida no tienen devolución.",
      "",
      "3. Exclusiones",
      "- Si la boquilla no está correctamente nivelada, plomeada o presenta descuadres, se generará un costo adicional por corrección o ajuste de la pieza.",
      "- No incluye trabajos adicionales como albañilería, pintura, pisos, electricidad, plomería, etc.",
      "- Dichos servicios podrán cotizarse de manera independiente, si el cliente lo solicita.",
      "",
      "4. Condiciones de Obra",
      "- Para iniciar fabricación e instalación, la obra debe contar con boquilla fondeada y una primera mano de pintura aplicada. La segunda mano de pintura deberá realizarse una vez instalada la ventana, por cuenta del cliente o de su contratista. Adicionalmente se requiere que pisos y azulejo estén instalados si aplica.",
      "- El cliente deberá proporcionar acceso, reglamento de seguridad y notificar si se requiere documentación especial al personal (ejemplo: certificaciones, exámenes, credenciales, etc).",
      "",
      "5. Materiales",
      "- Si el cliente aporta materiales propios, Grupo Lomelí no se hace responsable por daños o maltrato de los mismos.",
      "- Todas las medidas están expresadas en metros.",
      "- Resguardo: Una vez instaladas las ventanas, la custodia, cuidado y protección de las mismas será responsabilidad del cliente, quien deberá cubrirlas y resguardarlas adecuadamente durante la continuación de los trabajos de obra, pintura, limpieza o cualquier otra actividad posterior.",
      "- Limpieza: La limpieza realizada por nuestros técnicos corresponde únicamente a limpieza gruesa, limitada al retiro de residuos directamente generados por los trabajos ejecutados. No incluye limpieza fina, detallada ni de mantenimiento general del área.",
      "",
      "6. Garantías",
      "- Trabajos nuevos cuentan con 1 año de garantía a partir de la entrega.",
      "- Reparaciones y mantenimientos no tienen garantía.",
      "- La garantía es válida sólo si el proyecto está liquidado al 100%.",
      "- Modificaciones no autorizadas invalidan la garantía.",
      "- En caso de que las boquillas presenten imperfecciones, no se garantiza la estética o funcionamiento de la cancelería.",
      "",
      "7. Tiempos y Entregas",
      "- Los tiempos inician con anticipo confirmado y boquillas terminadas al 100%.",
      "- Se cuentan solo días hábiles (lunes a viernes, 8:00 a 18:00 hrs).",
      "- Retrasos por áreas no aptas o por falta de pagos reprogramarán la instalación.",
      "- En caso de retraso atribuible a Grupo Lomelí, se atenderá con urgencia. Si el retraso es atribuible al cliente, se reprogramará según disponibilidad.",
      "",
      "8. Comunicación",
      "- Toda comunicación oficial entre el cliente y Grupo Lomelí deberá canalizarse a través del representante asignado al proyecto (vendedor), quien será el enlace autorizado para brindar información sobre avances, fechas y acuerdos.",
      "- El personal técnico podrá mantener contacto directo únicamente para fines operativos, como coordinar accesos, horarios o detalles de instalación.",
      "- El personal técnico no está autorizado para negociar precios, realizar cambios al proyecto o comprometer fechas de entrega.",
      "  Esta medida busca mantener una comunicación clara, ordenada y profesional, evitando malentendidos o acuerdos fuera del alcance autorizado.",
      "",
      "9. Lineamientos",
      "Se reconoce una tolerancia de error de ±3 mm en cualquier producto. De igual manera, se establece que toda variación que no sea perceptible a una distancia mínima de tres metros no será considerada defecto ni motivo de reclamación.",
      "",
      "10. Disposiciones Generales",
      "- La persona contratante es responsable del pago total.",
      "- Cualquier situación no prevista en este documento será atendida conforme a la Política de Ventas interna de Grupo Lomelí, privilegiando siempre la transparencia y la comunicación con el cliente."
    ];

    for (var i = 0; i < terminos.length; i++) {
      var linea = terminos[i];
      var parrafo = body.appendParagraph(linea);

      if (linea.match(/^\d+\./)) {
        // Títulos de sección (ej: "1. Cotizaciones...")
        parrafo.editAsText().setBold(true).setFontSize(9).setForegroundColor("#595959");
        parrafo.setSpacingBefore(6);
      } else if (linea === "") {
        // Líneas vacías para separación
        parrafo.editAsText().setFontSize(6);
      } else {
        // Contenido normal
        parrafo.editAsText().setFontSize(7).setForegroundColor("#333333");
        parrafo.setLineSpacing(1.15);
      }
    }

    // Guardar y cerrar
    doc.saveAndClose();

    // Convertir a PDF
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(nombreArchivo + ".pdf");

    // Guardar PDF en carpeta del cliente
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    var infoCliente = obtenerInformacionCliente(hojaGenerador);

    if (infoCliente.exito && infoCliente.carpetaId) {
      try {
        var carpeta = DriveApp.getFolderById(infoCliente.carpetaId);
        var pdfFile = carpeta.createFile(pdfBlob);

        // Eliminar el Doc original de Mi Unidad
        DriveApp.getFileById(doc.getId()).setTrashed(true);

        Logger.log("✅ PDF guardado en carpeta del cliente");
        return pdfFile.getUrl();

      } catch (errorCarpeta) {
        Logger.log("⚠️ No se pudo guardar en carpeta del cliente, guardando en Mi Unidad");
        // Guardar en Mi Unidad si falla
        var pdfFile = DriveApp.createFile(pdfBlob);
        DriveApp.getFileById(doc.getId()).setTrashed(true);
        return pdfFile.getUrl();
      }
    } else {
      // Guardar en Mi Unidad
      var pdfFile = DriveApp.createFile(pdfBlob);
      DriveApp.getFileById(doc.getId()).setTrashed(true);
      return pdfFile.getUrl();
    }

  } catch (e) {
    Logger.log("❌ Error en generarDocumentoDesdeTemplate: " + e.message);
    Logger.log("📍 Stack: " + e.stack);
    return null;
  }
}

/**
 * Muestra el último folio generado
 */
function mostrarUltimoFolio() {
  var ultimoFolio = obtenerUltimoFolio();
  var ui = SpreadsheetApp.getUi();
  ui.alert('Último Folio Generado', ultimoFolio, ui.ButtonSet.OK);
}
