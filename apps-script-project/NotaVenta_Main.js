/**
 * Sistema de Notas de Venta
 * Genera notas de venta basadas en cotizaciones existentes
 */

/**
 * Obtiene la lista de cotizaciones disponibles
 * @returns {Array} - Array de objetos con información de cotizaciones
 */
function obtenerCotizacionesDisponibles() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaCotizaciones = ss.getSheetByName("Cotizaciones");

    if (!hojaCotizaciones) {
      Logger.log("⚠️ No existe la hoja Cotizaciones");
      return [];
    }

    var ultimaFila = hojaCotizaciones.getLastRow();

    if (ultimaFila < 2) {
      Logger.log("⚠️ No hay cotizaciones registradas");
      return [];
    }

    // Leer todas las cotizaciones (desde la fila 2)
    // Ahora tenemos 16 columnas con Productos Cliente (JSON)
    var datos = hojaCotizaciones.getRange(2, 1, ultimaFila - 1, 16).getValues();
    var cotizaciones = [];

    for (var i = 0; i < datos.length; i++) {
      try {
        var modoPrecioCerrado = String(datos[i][15] || "") === "SÍ";  // Modo B (columna 16)

        // Si es Modo B, usar productos CLIENTE (sin descuentos) para mostrar descuento visual
        // Si NO es Modo B, usar productos INTERNO (normales)
        var productosInterno = JSON.parse(datos[i][10] || "[]");  // Productos Interno JSON (columna 11)
        var productosCliente = JSON.parse(datos[i][11] || "[]");  // Productos Cliente JSON (columna 12)
        var productos = modoPrecioCerrado && productosCliente.length > 0 ? productosCliente : productosInterno;

        var datosCliente = JSON.parse(datos[i][12] || "{}");  // Datos Cliente JSON (columna 13)
        var datosCotizacion = JSON.parse(datos[i][13] || "{}");  // Datos Cotización JSON (columna 14)

        // Convertir fecha a string para evitar problemas de serialización
        var fechaStr = "";
        if (datos[i][1]) {
          try {
            fechaStr = Utilities.formatDate(new Date(datos[i][1]), "GMT-6", "dd/MM/yyyy");
          } catch (e) {
            fechaStr = String(datos[i][1]);
          }
        }

        cotizaciones.push({
          folio: Number(datos[i][0]) || 0,
          fecha: fechaStr,
          cliente: String(datos[i][2] || ""),
          proyecto: String(datos[i][3] || ""),
          vendedor: String(datos[i][4] || ""),
          subtotal: Number(datos[i][5]) || 0,
          iva: Number(datos[i][6]) || 0,
          total: Number(datos[i][7]) || 0,
          linkPDFInterno: String(datos[i][8] || ""),  // Link PDF Interno (columna 9)
          linkPDFCliente: String(datos[i][9] || ""),  // Link PDF Cliente (columna 10)
          productos: productos,  // Productos CLIENTE si es Modo B, sino INTERNO
          datosCliente: datosCliente,
          datosCotizacion: datosCotizacion,
          modoPrecioCerrado: modoPrecioCerrado
        });

      } catch (e) {
        Logger.log("⚠️ Error parseando cotización fila " + (i + 2) + ": " + e.message);
      }
    }

    Logger.log("✅ " + cotizaciones.length + " cotizaciones cargadas");
    return cotizaciones;

  } catch (e) {
    Logger.log("❌ Error en obtenerCotizacionesDisponibles: " + e.message);
    return [];
  }
}

/**
 * Genera un PDF de Nota de Venta basado en una cotización
 * @param {Object} datosNotaVenta - Datos completos de la nota de venta
 * @returns {Object} - Resultado con éxito, folio y URL del PDF
 */
function generarPDFNotaVenta(datosNotaVenta) {
  try {
    Logger.log("🚀 Iniciando generación de Nota de Venta...");

    // PASO 1: Generar folio único para la nota de venta
    Logger.log("📋 PASO 1: Generando folio de Nota de Venta...");
    var folioNV = generarFolioNotaVenta();
    Logger.log("✅ Folio NV generado: " + folioNV);

    // PASO 1.5: Si es Modo B, aplicar descuento 13.79% a los productos
    var productos = datosNotaVenta.productos;
    var modoPrecioCerrado = datosNotaVenta.modoPrecioCerrado || false;

    if (modoPrecioCerrado) {
      Logger.log("💰 MODO B DETECTADO - Aplicando descuento 13.79% a productos...");
      var DESCUENTO_IVA = 13.79;

      // Aplicar descuento a cada producto
      productos = productos.map(function(p) {
        var importe = p.importe || 0;
        var descuento = importe * (DESCUENTO_IVA / 100);
        var importeConDescuento = importe - descuento;

        return {
          codigo: p.codigo,
          descripcion: p.descripcion,
          cantidad: p.cantidad,
          precioUnitario: p.precioUnitario,
          descuentoPorcentaje: DESCUENTO_IVA,
          descuentoPesos: descuento,
          importe: importeConDescuento
        };
      });

      // Recalcular totales con descuentos aplicados
      var subtotal = 0;
      for (var i = 0; i < productos.length; i++) {
        subtotal += productos[i].importe;
      }
      var iva = subtotal * 0.16;
      var total = subtotal + iva;

      datosNotaVenta.totales = {
        subtotal: subtotal,
        iva: iva,
        total: total
      };

      Logger.log("✅ Descuentos aplicados - Nuevo total: $" + total.toFixed(2));
    }

    // PASO 2: Generar el documento PDF
    Logger.log("📋 PASO 2: Generando documento PDF...");
    var urlPDF = generarDocumentoNotaVenta(
      folioNV,
      datosNotaVenta.datosCliente,
      datosNotaVenta.datosCotizacion,
      productos,
      datosNotaVenta.totales,
      datosNotaVenta.modoProductos,
      {
        metodoPago: datosNotaVenta.metodoPago,
        condicionesPago: datosNotaVenta.condicionesPago,
        fechaEntrega: datosNotaVenta.fechaEntrega,
        lugarEntrega: datosNotaVenta.lugarEntrega,
        observaciones: datosNotaVenta.observaciones
      }
    );

    if (!urlPDF) {
      return {
        exito: false,
        mensaje: "Error al generar el documento PDF"
      };
    }

    Logger.log("✅ PDF generado: " + urlPDF);

    // PASO 3: Registrar nota de venta
    Logger.log("📋 PASO 3: Registrando Nota de Venta...");
    registrarNotaVenta(
      folioNV,
      datosNotaVenta.folioCotizacion,
      datosNotaVenta.datosCliente,
      datosNotaVenta.datosCotizacion,
      productos,  // Usar productos con descuentos si es Modo B
      datosNotaVenta.totales,
      datosNotaVenta.modoProductos,
      datosNotaVenta.metodoPago,
      datosNotaVenta.condicionesPago,
      urlPDF
    );
    Logger.log("✅ Nota de Venta registrada");

    return {
      exito: true,
      folio: folioNV,
      url: urlPDF,
      mensaje: "Nota de Venta generada exitosamente"
    };

  } catch (e) {
    Logger.log("❌ ERROR en generarPDFNotaVenta: " + e.message);
    Logger.log("📍 Stack: " + e.stack);
    return {
      exito: false,
      mensaje: "Error: " + e.message
    };
  }
}

/**
 * Genera el folio único para Nota de Venta
 * Formato: NV-0001, NV-0002, etc.
 * @returns {String} - Folio de Nota de Venta
 */
function generarFolioNotaVenta() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaNotasVenta = ss.getSheetByName("NotasVenta");
    var ultimoNumero = 0;

    if (hojaNotasVenta) {
      var ultimaFila = hojaNotasVenta.getLastRow();
      if (ultimaFila >= 2) {
        // Leer el último folio y extraer el número
        var ultimoFolio = hojaNotasVenta.getRange(ultimaFila, 1).getValue();
        if (ultimoFolio && typeof ultimoFolio === 'string') {
          var match = ultimoFolio.match(/NV-(\d+)/);
          if (match) {
            ultimoNumero = parseInt(match[1]);
          }
        }
      }
    }

    // Incrementar y formatear
    var nuevoNumero = ultimoNumero + 1;
    var nuevoFolio = "NV-" + String(nuevoNumero).padStart(4, '0');

    Logger.log("✅ Folio NV generado: " + nuevoFolio + " (último número: " + ultimoNumero + ")");

    return nuevoFolio;

  } catch (e) {
    Logger.log("❌ Error generando folio NV: " + e.message);
    return "NV-0001";
  }
}

/**
 * Registra una Nota de Venta en la hoja NotasVenta
 */
function registrarNotaVenta(folioNV, folioCotizacion, datosCliente, datosCotizacion, productos, totales, modoProductos, metodoPago, condicionesPago, urlPDF) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaNotasVenta = ss.getSheetByName("NotasVenta");

    // Crear hoja si no existe
    if (!hojaNotasVenta) {
      Logger.log("📝 Creando hoja NotasVenta...");
      hojaNotasVenta = ss.insertSheet("NotasVenta");

      var encabezados = [
        "Folio NV",
        "Fecha",
        "Folio Cotización",
        "Cliente",
        "Proyecto",
        "Vendedor",
        "Modo Productos",
        "Método Pago",
        "Condiciones Pago",
        "Subtotal",
        "IVA",
        "Total",
        "Link PDF",
        "Productos (JSON)"
      ];

      hojaNotasVenta.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
      hojaNotasVenta.getRange(1, 1, 1, encabezados.length).setFontWeight("bold").setBackground("#1b5e20").setFontColor("#ffffff");

      // Ajustar anchos
      hojaNotasVenta.setColumnWidth(1, 100);  // Folio NV
      hojaNotasVenta.setColumnWidth(2, 120);  // Fecha
      hojaNotasVenta.setColumnWidth(3, 100);  // Folio Cotización
      hojaNotasVenta.setColumnWidth(4, 200);  // Cliente
      hojaNotasVenta.setColumnWidth(5, 200);  // Proyecto
      hojaNotasVenta.setColumnWidth(6, 150);  // Vendedor
      hojaNotasVenta.setColumnWidth(7, 100);  // Modo
      hojaNotasVenta.setColumnWidth(8, 120);  // Método Pago
      hojaNotasVenta.setColumnWidth(9, 150);  // Condiciones
      hojaNotasVenta.setColumnWidth(10, 100); // Subtotal
      hojaNotasVenta.setColumnWidth(11, 100); // IVA
      hojaNotasVenta.setColumnWidth(12, 100); // Total
      hojaNotasVenta.setColumnWidth(13, 400); // Link PDF
      hojaNotasVenta.setColumnWidth(14, 300); // Productos JSON

      Logger.log("✅ Hoja NotasVenta creada");
    }

    // Agregar registro
    var ultimaFila = hojaNotasVenta.getLastRow() + 1;
    var fecha = new Date();

    var fila = [
      folioNV,
      fecha,
      folioCotizacion,
      datosCliente.nombre || "",
      datosCotizacion.idObra || datosCotizacion.proyecto || "",
      datosCotizacion.vendedor || Session.getActiveUser().getEmail(),
      modoProductos,
      metodoPago,
      condicionesPago,
      totales.subtotal,
      totales.iva,
      totales.total,
      urlPDF,
      JSON.stringify(productos)
    ];

    hojaNotasVenta.getRange(ultimaFila, 1, 1, fila.length).setValues([fila]);

    // Formatear
    hojaNotasVenta.getRange(ultimaFila, 10, 1, 3).setNumberFormat("$#,##0.00");
    hojaNotasVenta.getRange(ultimaFila, 2).setNumberFormat("dd/mm/yyyy hh:mm:ss");

    Logger.log("✅ Nota de Venta registrada: " + folioNV);

  } catch (e) {
    Logger.log("❌ Error al registrar Nota de Venta: " + e.message);
  }
}

/**
 * Muestra el formulario de Nota de Venta
 */
function mostrarFormularioNotaVenta() {
  var html = HtmlService.createHtmlOutputFromFile('NotaVenta_Webapp')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📄 Generar Nota de Venta');
}

/**
 * Genera el documento PDF de Nota de Venta
 * @param {String} folio - Folio de la nota de venta
 * @param {Object} datosCliente - Información del cliente
 * @param {Object} datosCotizacion - Información de la cotización
 * @param {Array} productos - Array de productos
 * @param {Object} totales - Totales calculados
 * @param {String} modoProductos - "resumido" o "desglosado"
 * @param {Object} datosAdicionales - Datos adicionales de la nota de venta
 * @returns {String} - URL del PDF generado
 */
function generarDocumentoNotaVenta(folio, datosCliente, datosCotizacion, productos, totales, modoProductos, datosAdicionales) {
  try {
    Logger.log("📄 Creando documento de Nota de Venta...");

    var nombreArchivo = "NotaVenta_" + folio + "_" + (datosCliente.nombre || "").replace(/[^a-zA-Z0-9]/g, '_');
    var doc = DocumentApp.create(nombreArchivo);
    var body = doc.getBody();

    // Configurar márgenes
    body.setMarginTop(40);
    body.setMarginBottom(40);
    body.setMarginLeft(50);
    body.setMarginRight(50);

    // ENCABEZADO CON LOGO Y DATOS DE LA EMPRESA
    var headerTable = body.appendTable();
    var headerRow = headerTable.appendTableRow();

    // Datos de la empresa
    var cellEmpresa = headerRow.appendTableCell();
    cellEmpresa.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    var nombreEmpresa = cellEmpresa.appendParagraph("NUEVA KRYSTALUM LOMELÍ");
    nombreEmpresa.editAsText().setBold(true).setFontSize(12).setForegroundColor("#1b5e20");
    cellEmpresa.appendParagraph("Rafael Camacho 1396, Ladrón de Guevara").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("Guadalajara, JALISCO").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("Tel.(333) 853-7583").editAsText().setFontSize(8);
    cellEmpresa.appendParagraph("contacto@nkrystalum.mx").editAsText().setFontSize(8);

    // Logo
    var cellLogo = headerRow.appendTableCell();
    cellLogo.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    try {
      var logoId = "1u47sWEv4LYocFqGXNXFdS9M1Vhn_RnqP";
      var logoFile = DriveApp.getFileById(logoId);
      var logoBlob = logoFile.getBlob();
      var logoImage = cellLogo.appendParagraph("").appendInlineImage(logoBlob);
      logoImage.setWidth(171);
      logoImage.setHeight(50);
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
    var titulo = body.appendParagraph("NOTA DE VENTA " + folio);
    titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titulo.editAsText().setBold(true).setFontSize(14).setForegroundColor("#1b5e20");

    body.appendParagraph(""); // Espacio

    // DATOS DEL CLIENTE (tabla estilo verde)
    var clienteTable = body.appendTable();

    function aplicarEstiloLabelVerde(cell) {
      cell.getChild(0).asParagraph().editAsText().setFontSize(8).setBold(true).setForegroundColor("#1b5e20");
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setPaddingLeft(4);
      cell.setPaddingRight(4);
      cell.setBackgroundColor("#e8f5e9");
    }

    function aplicarEstiloValorVerde(cell) {
      cell.getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor("#000000");
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setPaddingLeft(4);
      cell.setPaddingRight(4);
    }

    // Fila 1: Cliente y Proyecto
    var row1 = clienteTable.appendTableRow();
    var cellLabel1 = row1.appendTableCell("CLIENTE:");
    aplicarEstiloLabelVerde(cellLabel1);
    var cellValue1 = row1.appendTableCell(datosCliente.nombre || "");
    aplicarEstiloValorVerde(cellValue1);
    var cellLabel2 = row1.appendTableCell("PROYECTO:");
    aplicarEstiloLabelVerde(cellLabel2);
    var cellValue2 = row1.appendTableCell(datosCotizacion.idObra || datosCotizacion.proyecto || "");
    aplicarEstiloValorVerde(cellValue2);

    // Fila 2: Contacto y Teléfono
    var row2 = clienteTable.appendTableRow();
    var cellLabel3 = row2.appendTableCell("CONTACTO:");
    aplicarEstiloLabelVerde(cellLabel3);
    var cellValue3 = row2.appendTableCell(datosCliente.contacto || "");
    aplicarEstiloValorVerde(cellValue3);
    var cellLabel4 = row2.appendTableCell("TELÉFONO:");
    aplicarEstiloLabelVerde(cellLabel4);
    var cellValue4 = row2.appendTableCell(datosCliente.telefono || "");
    aplicarEstiloValorVerde(cellValue4);

    // Fila 3: Email y Fecha
    var row3 = clienteTable.appendTableRow();
    var cellLabel5 = row3.appendTableCell("EMAIL:");
    aplicarEstiloLabelVerde(cellLabel5);
    var cellValue5 = row3.appendTableCell(datosCliente.email || "");
    aplicarEstiloValorVerde(cellValue5);
    var cellLabel6 = row3.appendTableCell("FECHA:");
    aplicarEstiloLabelVerde(cellLabel6);
    var fechaActual = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy");
    var cellValue6 = row3.appendTableCell(fechaActual);
    aplicarEstiloValorVerde(cellValue6);

    clienteTable.setBorderWidth(0.5);
    clienteTable.setBorderColor("#cccccc");

    body.appendParagraph(""); // Espacio

    // DATOS ADICIONALES DE NOTA DE VENTA
    var datosNVTable = body.appendTable();

    // Fila 1: Método de Pago y Condiciones
    var rowNV1 = datosNVTable.appendTableRow();
    var cellLabelNV1 = rowNV1.appendTableCell("MÉTODO DE PAGO:");
    aplicarEstiloLabelVerde(cellLabelNV1);
    var cellValueNV1 = rowNV1.appendTableCell(datosAdicionales.metodoPago || "");
    aplicarEstiloValorVerde(cellValueNV1);
    var cellLabelNV2 = rowNV1.appendTableCell("CONDICIONES:");
    aplicarEstiloLabelVerde(cellLabelNV2);
    var cellValueNV2 = rowNV1.appendTableCell(datosAdicionales.condicionesPago || "");
    aplicarEstiloValorVerde(cellValueNV2);

    // Fila 2: Fecha Entrega y Lugar
    if (datosAdicionales.fechaEntrega || datosAdicionales.lugarEntrega) {
      var rowNV2 = datosNVTable.appendTableRow();
      var cellLabelNV3 = rowNV2.appendTableCell("ENTREGA:");
      aplicarEstiloLabelVerde(cellLabelNV3);
      var cellValueNV3 = rowNV2.appendTableCell(datosAdicionales.fechaEntrega || "");
      aplicarEstiloValorVerde(cellValueNV3);
      var cellLabelNV4 = rowNV2.appendTableCell("LUGAR:");
      aplicarEstiloLabelVerde(cellLabelNV4);
      var cellValueNV4 = rowNV2.appendTableCell(datosAdicionales.lugarEntrega || "");
      aplicarEstiloValorVerde(cellValueNV4);
    }

    datosNVTable.setBorderWidth(0.5);
    datosNVTable.setBorderColor("#cccccc");

    body.appendParagraph(""); // Espacio

    // TABLA DE PRODUCTOS (modo resumido o desglosado)
    // Ambos modos usan la misma tabla completa
    var productosTable = body.appendTable();

    var headerProductos = productosTable.appendTableRow();
    headerProductos.appendTableCell("No.").setWidth(25);
    headerProductos.appendTableCell("Código").setWidth(55);
    headerProductos.appendTableCell("Descripción").setWidth(150);
    headerProductos.appendTableCell("Cant.").setWidth(35);
    headerProductos.appendTableCell("P.U.").setWidth(55);
    headerProductos.appendTableCell("% Desc.").setWidth(35);
    headerProductos.appendTableCell("$ Desc.").setWidth(55);
    headerProductos.appendTableCell("Importe").setWidth(55);

    for (var i = 0; i < headerProductos.getNumCells(); i++) {
      var cell = headerProductos.getCell(i);
      cell.getChild(0).asParagraph().editAsText().setBold(true).setFontSize(8).setForegroundColor("#ffffff");
      cell.setBackgroundColor("#1b5e20");
      cell.setPaddingTop(3);
      cell.setPaddingBottom(3);
      cell.setPaddingLeft(2);
      cell.setPaddingRight(2);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    // Filas de productos (una sola en resumido, todas en desglosado)
    if (modoProductos === "resumido") {
      // MODO RESUMIDO: Una sola fila con descripción editada
      var prod = productos[0]; // En resumido solo hay un producto
      var rowProd = productosTable.appendTableRow();
      var bgColor = "#ffffff";

      var cellNo = rowProd.appendTableCell("1");
      var cellCodigo = rowProd.appendTableCell(prod.codigo || "");
      var cellDesc = rowProd.appendTableCell(prod.descripcion || "");

      var cantidad = prod.cantidad || 0;
      var cellCant = rowProd.appendTableCell(String(cantidad));

      var precioUnitario = prod.precioUnitario || 0;
      var puFormateado = typeof precioUnitario === 'number'
        ? precioUnitario.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
        : String(precioUnitario);
      var cellPU = rowProd.appendTableCell(puFormateado);

      var descPorc = prod.descuentoPorcentaje || 0;
      var cellDescPorc = rowProd.appendTableCell(descPorc + "%");

      var descPesos = prod.descuentoPesos || 0;
      var descPesosFormateado = typeof descPesos === 'number'
        ? descPesos.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
        : String(descPesos);
      var cellDescPesos = rowProd.appendTableCell(descPesosFormateado);

      var importe = prod.importe || 0;
      var importeFormateado = typeof importe === 'number'
        ? importe.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
        : String(importe);
      var cellImporte = rowProd.appendTableCell(importeFormateado);

      var cells = [cellNo, cellCodigo, cellDesc, cellCant, cellPU, cellDescPorc, cellDescPesos, cellImporte];
      for (var j = 0; j < cells.length; j++) {
        cells[j].getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor("#000000");
        cells[j].setPaddingTop(1.5);
        cells[j].setPaddingBottom(1.5);
        cells[j].setPaddingLeft(2);
        cells[j].setPaddingRight(2);
        cells[j].setBackgroundColor(bgColor);
      }

      cellNo.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cellCant.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cellPU.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      cellDescPorc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      cellDescPesos.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      cellImporte.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

    } else {
      // MODO DESGLOSADO: Todas las filas de productos
      for (var i = 0; i < productos.length; i++) {
        var prod = productos[i];
        var rowProd = productosTable.appendTableRow();
        var bgColor = (i % 2 === 0) ? "#ffffff" : "#f1f8e9";

        var cellNo = rowProd.appendTableCell(String(i + 1));
        var cellCodigo = rowProd.appendTableCell(prod.codigo || "");
        var cellDesc = rowProd.appendTableCell(prod.descripcion || "");

        var cantidad = prod.cantidad || 0;
        var cellCant = rowProd.appendTableCell(String(cantidad));

        var precioUnitario = prod.precioUnitario || 0;
        var puFormateado = typeof precioUnitario === 'number'
          ? precioUnitario.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
          : String(precioUnitario);
        var cellPU = rowProd.appendTableCell(puFormateado);

        var descPorc = prod.descuentoPorcentaje || 0;
        var cellDescPorc = rowProd.appendTableCell(descPorc + "%");

        var descPesos = prod.descuentoPesos || 0;
        var descPesosFormateado = typeof descPesos === 'number'
          ? descPesos.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
          : String(descPesos);
        var cellDescPesos = rowProd.appendTableCell(descPesosFormateado);

        var importe = prod.importe || 0;
        var importeFormateado = typeof importe === 'number'
          ? importe.toLocaleString('es-MX', {minimumFractionDigits: 2, maximumFractionDigits: 2})
          : String(importe);
        var cellImporte = rowProd.appendTableCell(importeFormateado);

        var cells = [cellNo, cellCodigo, cellDesc, cellCant, cellPU, cellDescPorc, cellDescPesos, cellImporte];
        for (var j = 0; j < cells.length; j++) {
          cells[j].getChild(0).asParagraph().editAsText().setFontSize(8).setForegroundColor("#000000");
          cells[j].setPaddingTop(1.5);
          cells[j].setPaddingBottom(1.5);
          cells[j].setPaddingLeft(2);
          cells[j].setPaddingRight(2);
          cells[j].setBackgroundColor(bgColor);
        }

        cellNo.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        cellCant.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        cellPU.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        cellDescPorc.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        cellDescPesos.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
        cellImporte.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      }

      productosTable.setBorderWidth(0.5);
      productosTable.setBorderColor("#cccccc");
    }

    body.appendParagraph(""); // Espacio

    // TOTALES
    var totalesTable = body.appendTable();

    var rowSubtotal = totalesTable.appendTableRow();
    var cellLabelSub = rowSubtotal.appendTableCell("Subtotal:");
    cellLabelSub.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellLabelSub.getChild(0).asParagraph().editAsText().setFontSize(10).setBold(true);
    var cellValorSub = rowSubtotal.appendTableCell("$" + totales.subtotal.toLocaleString('es-MX', {minimumFractionDigits: 2}));
    cellValorSub.getChild(0).asParagraph().editAsText().setFontSize(10);

    var rowIVA = totalesTable.appendTableRow();
    var cellLabelIVA = rowIVA.appendTableCell("IVA (16%):");
    cellLabelIVA.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellLabelIVA.getChild(0).asParagraph().editAsText().setFontSize(10).setBold(true);
    var cellValorIVA = rowIVA.appendTableCell("$" + totales.iva.toLocaleString('es-MX', {minimumFractionDigits: 2}));
    cellValorIVA.getChild(0).asParagraph().editAsText().setFontSize(10);

    var rowTotal = totalesTable.appendTableRow();
    var cellLabelTotal = rowTotal.appendTableCell("TOTAL:");
    cellLabelTotal.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    cellLabelTotal.getChild(0).asParagraph().editAsText().setFontSize(12).setBold(true).setForegroundColor("#1b5e20");
    cellLabelTotal.setBackgroundColor("#e8f5e9");
    var cellValorTotal = rowTotal.appendTableCell("$" + totales.total.toLocaleString('es-MX', {minimumFractionDigits: 2}));
    cellValorTotal.getChild(0).asParagraph().editAsText().setFontSize(12).setBold(true).setForegroundColor("#1b5e20");
    cellValorTotal.setBackgroundColor("#e8f5e9");

    totalesTable.setBorderWidth(0);
    rowSubtotal.getCell(0).setWidth(360);
    rowSubtotal.getCell(1).setWidth(120);

    body.appendParagraph(""); // Espacio

    // Total en letras
    var totalEnLetras = numeroALetras(totales.total);
    var parrafoLetras = body.appendParagraph("SON: " + totalEnLetras);
    parrafoLetras.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    parrafoLetras.editAsText().setItalic(true).setFontSize(9).setBold(true);

    // Observaciones
    if (datosAdicionales.observaciones) {
      body.appendParagraph(""); // Espacio
      var tituloObs = body.appendParagraph("OBSERVACIONES:");
      tituloObs.editAsText().setBold(true).setFontSize(9).setForegroundColor("#1b5e20");
      var textoObs = body.appendParagraph(datosAdicionales.observaciones);
      textoObs.editAsText().setFontSize(8).setItalic(true);
    }

    body.appendParagraph(""); // Espacio

    // ========================================
    // TÉRMINOS Y CONDICIONES (NUEVA PÁGINA)
    // ========================================
    body.appendPageBreak();

    var tituloTerminos = body.appendParagraph("TÉRMINOS Y CONDICIONES");
    tituloTerminos.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    tituloTerminos.editAsText().setBold(true).setFontSize(12).setForegroundColor("#595959");

    body.appendParagraph(""); // Espacio

    // Array con los términos y condiciones
    var terminosYCondiciones = [
      { titulo: "1. Cotizaciones y Contratación", items: [
        "- Las cotizaciones tienen validez en tanto no exista incremento en materiales.",
        "- Se requiere firma, nombre del cliente y fecha en la cotización o contrato."
      ]},
      { titulo: "2. Precios, Pagos y Devoluciones", items: [
        "- Precios: Los precios no incluyen IVA (se desglosa al final, en la cotización).",
        "- Anticipo y pagos: Se requiere un anticipo del 60% para iniciar trabajos. El saldo se cubrirá mediante pagos parciales por avance de obra y/o contra entrega.",
        "- Estimaciones: Se realizarán con base en el avance real de obra, y su frecuencia será semanal, salvo que el cliente y el vendedor acuerden un periodo distinto.",
        "El pago de cada estimación deberá efectuarse en un plazo máximo de dos días hábiles después de su emisión.",
        "El objetivo principal de una estimación es que el cliente pague únicamente por el avance real de lo instalado en obra.",
        "A modo de ejemplo, el procedimiento podrá desarrollarse de la siguiente manera:",
        "- El anticipo del 60 % se prorratea proporcionalmente entre todas las partidas; este anticipo se destina a la compra de materiales y garantiza el precio contratado, por lo que se considera aplicado desde el inicio.",
        "- Conforme avanza la instalación, se cobra el porcentaje correspondiente al progreso físico de cada partida. Si una partida se instala al 100 %, se factura el porcentaje restante de esa partida. Si una partida queda parcialmente instalada, se cobra solo el porcentaje equivalente al avance ejecutado, dejando el saldo pendiente para la entrega final.",
        "Este esquema busca mantener una relación justa y transparente entre el avance físico y los pagos, permitiendo ajustes flexibles conforme a los acuerdos específicos con el vendedor.",
        "-Devoluciones: Las devoluciones deberán ser autorizadas por la Dirección Comercial. En caso de proceder, se aplicará un cargo mínimo del 20% sobre el valor del producto o servicio, además de los costos de materiales y mano de obra utilizados. Los productos fabricados a la medida no tienen devolución."
      ]},
      { titulo: "3. Exclusiones", items: [
        "- Si la boquilla no está correctamente nivelada, plomeada o presenta descuadres, se generará un costo adicional por corrección o ajuste de la pieza.",
        "- No incluye trabajos adicionales como albañilería, pintura, pisos, electricidad, plomería, etc.",
        "- Dichos servicios podrán cotizarse de manera independiente, si el cliente lo solicita."
      ]},
      { titulo: "4. Condiciones de Obra", items: [
        "- Para iniciar fabricación e instalación, la obra debe contar con boquilla fondeada y una primera mano de pintura aplicada. La segunda mano de pintura deberá realizarse una vez instalada la ventana, por cuenta del cliente o de su contratista. Adicionalmente se requiere que pisos y azulejo estén instalados si aplica.",
        "- El cliente deberá proporcionar acceso, reglamento de seguridad y notificar si se requiere documentación especial al personal (ejemplo: certificaciones, exámenes, credenciales, etc)."
      ]},
      { titulo: "5. Materiales", items: [
        "- Si el cliente aporta materiales propios, Grupo Lomelí no se hace responsable por daños o maltrato de los mismos.",
        "- Todas las medidas están expresadas en metros.",
        "- Resguardo: Una vez instaladas las ventanas, la custodia, cuidado y protección de las mismas será responsabilidad del cliente, quien deberá cubrirlas y resguardarlas adecuadamente durante la continuación de los trabajos de obra, pintura, limpieza o cualquier otra actividad posterior.",
        "- Limpieza: La limpieza realizada por nuestros técnicos corresponde únicamente a limpieza gruesa, limitada al retiro de residuos directamente generados por los trabajos ejecutados. No incluye limpieza fina, detallada ni de mantenimiento general del área."
      ]},
      { titulo: "6. Garantías", items: [
        "- Trabajos nuevos cuentan con 1 año de garantía a partir de la entrega.",
        "- Reparaciones y mantenimientos no tienen garantía.",
        "- La garantía es válida sólo si el proyecto está liquidado al 100%.",
        "- Modificaciones no autorizadas invalidan la garantía.",
        "- En caso de que las boquillas presenten imperfecciones, no se garantiza la estética o funcionamiento de la cancelería."
      ]},
      { titulo: "7. Tiempos y Entregas", items: [
        "- Los tiempos inician con anticipo confirmado y boquillas terminadas al 100%.",
        "- Se cuentan solo días hábiles (lunes a viernes, 8:00 a 18:00 hrs).",
        "- Retrasos por áreas no aptas o por falta de pagos reprogramarán la instalación.",
        "- En caso de retraso atribuible a Grupo Lomelí, se atenderá con urgencia. Si el retraso es atribuible al cliente, se reprogramará según disponibilidad."
      ]},
      { titulo: "8. Comunicación", items: [
        "- Toda comunicación oficial entre el cliente y Grupo Lomelí deberá canalizarse a través del representante asignado al proyecto (vendedor), quien será el enlace autorizado para brindar información sobre avances, fechas y acuerdos.",
        "- El personal técnico podrá mantener contacto directo únicamente para fines operativos, como coordinar accesos, horarios o detalles de instalación.",
        "- El personal técnico no está autorizado para negociar precios, realizar cambios al proyecto o comprometer fechas de entrega.",
        "Esta medida busca mantener una comunicación clara, ordenada y profesional, evitando malentendidos o acuerdos fuera del alcance autorizado."
      ]},
      { titulo: "9. Lineamientos", items: [
        "Se reconoce una tolerancia de error de ±3 mm en cualquier producto. De igual manera, se establece que toda variación que no sea perceptible a una distancia mínima de tres metros no será considerada defecto ni motivo de reclamación."
      ]},
      { titulo: "10. Disposiciones Generales", items: [
        "- La persona contratante es responsable del pago total.",
        "- Cualquier situación no prevista en este documento será atendida conforme a la Política de Ventas interna de Grupo Lomelí, privilegiando siempre la transparencia y la comunicación con el cliente."
      ]}
    ];

    // Insertar cada sección de términos
    for (var i = 0; i < terminosYCondiciones.length; i++) {
      var seccion = terminosYCondiciones[i];

      // Título de la sección
      var parrafo = body.appendParagraph(seccion.titulo);
      parrafo.editAsText().setBold(true).setFontSize(9).setForegroundColor("#595959");

      // Items de la sección
      for (var j = 0; j < seccion.items.length; j++) {
        var parrafo = body.appendParagraph(seccion.items[j]);
        parrafo.editAsText().setFontSize(7).setForegroundColor("#333333");
      }

      body.appendParagraph(""); // Espacio entre secciones
    }

    body.appendParagraph(""); // Espacio

    // FIRMAS
    var tablaFirmas = body.appendTable();
    var filaFirmas = tablaFirmas.appendTableRow();

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

    var celdaCliente = filaFirmas.appendTableCell();
    celdaCliente.appendParagraph("").appendHorizontalRule();
    var nombreCliente = celdaCliente.appendParagraph("Firma del Cliente");
    nombreCliente.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    nombreCliente.editAsText().setFontSize(9).setBold(true);
    var datoCliente = celdaCliente.appendParagraph(datosCliente.nombre || "");
    datoCliente.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    datoCliente.editAsText().setFontSize(8);
    celdaCliente.setPaddingTop(10);
    celdaCliente.setPaddingBottom(10);

    tablaFirmas.setBorderWidth(0);
    filaFirmas.getCell(0).setWidth(240);
    filaFirmas.getCell(1).setWidth(240);

    // Guardar y cerrar
    doc.saveAndClose();

    // Convertir a PDF
    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs('application/pdf');
    pdfBlob.setName(nombreArchivo + ".pdf");

    // Guardar en carpeta del cliente
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");
    var infoCliente = obtenerInformacionCliente(hojaGenerador);

    if (infoCliente.exito && infoCliente.carpetaId) {
      try {
        var carpeta = DriveApp.getFolderById(infoCliente.carpetaId);
        var pdfFile = carpeta.createFile(pdfBlob);
        DriveApp.getFileById(doc.getId()).setTrashed(true);
        Logger.log("✅ PDF guardado en carpeta del cliente");
        return pdfFile.getUrl();
      } catch (errorCarpeta) {
        Logger.log("⚠️ No se pudo guardar en carpeta del cliente, guardando en Mi Unidad");
        var pdfFile = DriveApp.createFile(pdfBlob);
        DriveApp.getFileById(doc.getId()).setTrashed(true);
        return pdfFile.getUrl();
      }
    } else {
      var pdfFile = DriveApp.createFile(pdfBlob);
      DriveApp.getFileById(doc.getId()).setTrashed(true);
      Logger.log("✅ PDF guardado en Mi Unidad");
      return pdfFile.getUrl();
    }

  } catch (e) {
    Logger.log("❌ Error en generarDocumentoNotaVenta: " + e.message);
    Logger.log("📍 Stack: " + e.stack);
    return null;
  }
}
