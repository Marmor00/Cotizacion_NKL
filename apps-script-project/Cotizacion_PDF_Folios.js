/**
 * Sistema de generación de folios únicos para cotizaciones
 * Mantiene un contador en una hoja oculta para generar números secuenciales
 */

/**
 * Genera un nuevo folio único para la cotización
 * Formato: Número consecutivo empezando desde 7744
 * @returns {Number} - Folio único generado
 */
function generarFolioUnico() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Intentar leer el último folio de la hoja Cotizaciones
    var hojaCotizaciones = ss.getSheetByName("Cotizaciones");
    var ultimoNumero = 7115; // Folio inicial menos 1 (primer folio será 7116)

    if (hojaCotizaciones) {
      var ultimaFila = hojaCotizaciones.getLastRow();
      if (ultimaFila >= 2) { // Si hay datos además del encabezado
        var ultimoFolioRegistrado = hojaCotizaciones.getRange(ultimaFila, 1).getValue();
        if (ultimoFolioRegistrado && !isNaN(ultimoFolioRegistrado)) {
          ultimoNumero = parseInt(ultimoFolioRegistrado);
        }
      }
    }

    // Incrementar para el nuevo folio
    // Si es la primera vez (hoja vacía o no existe), será 7117
    var nuevoFolio = ultimoNumero + 1;

    Logger.log("✅ Folio generado: " + nuevoFolio + " (último: " + ultimoNumero + ")");

    return nuevoFolio;

  } catch (e) {
    Logger.log("❌ Error generando folio: " + e.message);
    // Generar folio de emergencia - primer folio
    return 7116;
  }
}

/**
 * Obtiene o crea la hoja de control de folios
 * @param {Spreadsheet} ss - Spreadsheet activo
 * @returns {Sheet} - Hoja de folios
 */
function obtenerHojaFolios(ss) {
  var nombreHoja = "Config_Folios";
  var hojaFolios = ss.getSheetByName(nombreHoja);

  if (!hojaFolios) {
    Logger.log("📝 Creando hoja de configuración de folios...");

    // Crear hoja nueva
    hojaFolios = ss.insertSheet(nombreHoja);

    // Configurar encabezados
    hojaFolios.getRange("A1").setValue("CONFIGURACIÓN DE FOLIOS").setFontWeight("bold").setFontSize(12);
    hojaFolios.getRange("A2").setValue("Último Número:").setFontWeight("bold");
    hojaFolios.getRange("B2").setValue(0);

    hojaFolios.getRange("D1").setValue("HISTORIAL DE FOLIOS").setFontWeight("bold").setFontSize(12);
    hojaFolios.getRange("D2").setValue("Fecha/Hora").setFontWeight("bold");
    hojaFolios.getRange("E2").setValue("Folio Generado").setFontWeight("bold");

    // Ajustar anchos de columna
    hojaFolios.setColumnWidth(1, 150);
    hojaFolios.setColumnWidth(2, 120);
    hojaFolios.setColumnWidth(4, 150);
    hojaFolios.setColumnWidth(5, 200);

    // Ocultar la hoja (opcional - puedes comentar esta línea si quieres verla)
    hojaFolios.hideSheet();

    Logger.log("✅ Hoja de folios creada y configurada");
  }

  return hojaFolios;
}

/**
 * Resetea el contador de folios (USAR CON CUIDADO)
 * Solo para testing o cuando cambie el año
 */
function resetearContadorFolios() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.alert(
    'Confirmar Reset',
    '⚠️ ¿Estás seguro de resetear el contador de folios a 0?\n\nEsto puede causar duplicados si ya existen cotizaciones.',
    ui.ButtonSet.YES_NO
  );

  if (respuesta === ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaFolios = obtenerHojaFolios(ss);
    hojaFolios.getRange("B2").setValue(0);

    ui.alert("✅ Contador reseteado a 0");
    Logger.log("⚠️ Contador de folios reseteado a 0");
  }
}

/**
 * Función de prueba para generar varios folios
 */
function probarGeneracionFolios() {
  Logger.log("🧪 Probando generación de folios...");

  for (var i = 1; i <= 5; i++) {
    var folio = generarFolioUnico();
    Logger.log("Folio " + i + ": " + folio);
  }

  Logger.log("✅ Prueba completada");
}

/**
 * Obtiene el último folio generado (sin crear uno nuevo)
 * @returns {String} - Último folio o mensaje si no hay
 */
function obtenerUltimoFolio() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaCotizaciones = ss.getSheetByName("Cotizaciones");

    if (!hojaCotizaciones) {
      return "No se han generado folios aún";
    }

    var ultimaFila = hojaCotizaciones.getLastRow();

    if (ultimaFila < 2) {
      return "No se han generado folios aún";
    }

    var ultimoFolio = hojaCotizaciones.getRange(ultimaFila, 1).getValue();
    return ultimoFolio;

  } catch (e) {
    return "Error: " + e.message;
  }
}

/**
 * Registra una nueva cotización en la hoja Cotizaciones
 * @param {Number} folio - Número de folio
 * @param {Object} datosCliente - Información del cliente
 * @param {Object} datosCotizacion - Información de la cotización
 * @param {Array} productos - Array de productos incluidos en la cotización
 * @param {Object} totales - Totales calculados
 * @param {String} urlPDF - URL del PDF generado
 */
function registrarCotizacion(folio, datosCliente, datosCotizacion, productos, totales, urlPDFInterno, urlPDFCliente) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaCotizaciones = ss.getSheetByName("Cotizaciones");

    // Si no existe la hoja, crearla
    if (!hojaCotizaciones) {
      Logger.log("📝 Creando hoja Cotizaciones...");
      hojaCotizaciones = ss.insertSheet("Cotizaciones");

      // Crear encabezados
      var encabezados = [
        "Folio",
        "Fecha",
        "Cliente",
        "Proyecto",
        "Vendedor",
        "Subtotal",
        "IVA",
        "Total",
        "Link PDF Interno",
        "Link PDF Cliente",
        "Productos (JSON)",
        "Datos Cliente (JSON)",
        "Datos Cotización (JSON)",
        "Descripción Resumida",
        "Modo Precio Cerrado"
      ];

      hojaCotizaciones.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
      hojaCotizaciones.getRange(1, 1, 1, encabezados.length).setFontWeight("bold").setBackground("#8E8E8E").setFontColor("#ffffff");

      // Ajustar anchos de columna
      hojaCotizaciones.setColumnWidth(1, 80);   // Folio
      hojaCotizaciones.setColumnWidth(2, 120);  // Fecha
      hojaCotizaciones.setColumnWidth(3, 200);  // Cliente
      hojaCotizaciones.setColumnWidth(4, 200);  // Proyecto
      hojaCotizaciones.setColumnWidth(5, 150);  // Vendedor
      hojaCotizaciones.setColumnWidth(6, 100);  // Subtotal
      hojaCotizaciones.setColumnWidth(7, 100);  // IVA
      hojaCotizaciones.setColumnWidth(8, 100);  // Total
      hojaCotizaciones.setColumnWidth(9, 400);  // Link PDF Interno
      hojaCotizaciones.setColumnWidth(10, 400); // Link PDF Cliente
      hojaCotizaciones.setColumnWidth(11, 300); // Productos JSON
      hojaCotizaciones.setColumnWidth(12, 250); // Datos Cliente JSON
      hojaCotizaciones.setColumnWidth(13, 250); // Datos Cotización JSON
      hojaCotizaciones.setColumnWidth(14, 400); // Descripción Resumida
      hojaCotizaciones.setColumnWidth(15, 120); // Modo Precio Cerrado

      Logger.log("✅ Hoja Cotizaciones creada");
    }

    // Agregar nueva fila con los datos
    var ultimaFila = hojaCotizaciones.getLastRow() + 1;
    var fecha = new Date();

    // Convertir productos a JSON (solo campos relevantes)
    var productosSimplificados = productos.map(function(p) {
      return {
        codigo: p.codigoEditado || p.codigo,
        descripcion: p.descripcion,
        cantidad: p.piezas || p.cantidad,
        precioUnitario: p.precioUnitario || p.precioVenta,
        descuentoPorcentaje: p.descuentoPorcentaje || 0,
        descuentoPesos: p.descuentoPesos || p.descuentoMonto || 0,
        importe: p.importeFinal || p.importe
      };
    });

    // Generar descripción resumida (similar a NotasVenta pero sin "Atendido por")
    var descripcionResumida = "Cotización correspondiente";
    var idProyecto = datosCotizacion.idObra || datosCotizacion.proyecto || "";
    var domicilio = datosCliente.domicilioEntrega || "";

    if (idProyecto) {
      descripcionResumida += " al proyecto \"" + idProyecto + "\"";
    }
    if (domicilio) {
      descripcionResumida += ", instalado en " + domicilio;
    }
    descripcionResumida += ".";

    var fila = [
      folio,
      fecha,
      datosCliente.nombre || "",
      datosCotizacion.idObra || datosCotizacion.proyecto || "",
      datosCotizacion.vendedor || Session.getActiveUser().getEmail(),
      totales.subtotal,
      totales.iva,
      totales.total,
      urlPDFInterno,
      urlPDFCliente || "",  // Puede ser null si no es Modo B
      JSON.stringify(productosSimplificados),
      JSON.stringify(datosCliente),
      JSON.stringify(datosCotizacion),
      descripcionResumida,
      datosCotizacion.modoPrecioCerrado ? "SÍ" : "NO"
    ];

    hojaCotizaciones.getRange(ultimaFila, 1, 1, fila.length).setValues([fila]);

    // Formatear montos como moneda
    hojaCotizaciones.getRange(ultimaFila, 6, 1, 3).setNumberFormat("$#,##0.00");

    // Formatear fecha
    hojaCotizaciones.getRange(ultimaFila, 2).setNumberFormat("dd/mm/yyyy hh:mm:ss");

    Logger.log("✅ Cotización registrada en hoja Cotizaciones: Folio " + folio);

  } catch (e) {
    Logger.log("❌ Error al registrar cotización: " + e.message);
  }
}
