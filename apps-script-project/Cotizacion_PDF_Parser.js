/**
 * Parser para extraer productos de la hoja "Generador" y estructurarlos para la cotización formal
 * Lee bloques de productos y extrae solo la información necesaria para el PDF
 */

/**
 * Función principal que lee toda la hoja Generador y devuelve productos estructurados
 * @returns {Object} - { productos: [...], totales: {...}, proyecto: "..." }
 */
function parsearHojaGenerador() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaGenerador = ss.getSheetByName("Generador");

    if (!hojaGenerador) {
      throw new Error("No se encontró la hoja 'Generador'");
    }

    Logger.log("📖 Iniciando parseo de hoja Generador...");

    // Obtener última fila con datos (desde K1)
    var ultimaFila = hojaGenerador.getRange("K1").getValue();
    if (!ultimaFila || ultimaFila < 4) {
      throw new Error("No hay datos en la hoja Generador (K1 = " + ultimaFila + ")");
    }

    Logger.log("Última fila detectada: " + ultimaFila);

    // Leer todo el rango de datos
    var rangoDatos = hojaGenerador.getRange(1, 1, ultimaFila, 12).getValues(); // A1:L{ultimaFila}

    // Extraer información del proyecto (fila 1, columna E - índice 4)
    var nombreProyecto = rangoDatos[1] ? rangoDatos[1][4] : ""; // Fila 2, columna E
    var totalMaterial = rangoDatos[1] ? rangoDatos[1][11] : ""; // Fila 2, columna L

    Logger.log("Proyecto: " + nombreProyecto);
    Logger.log("Total Material: " + totalMaterial);

    // Detectar y extraer cada bloque de producto
    var productos = [];
    var filaActual = 3; // Empezar desde fila 4 (índice 3)

    while (filaActual < rangoDatos.length) {
      var producto = detectarYExtraerProducto(rangoDatos, filaActual);

      if (producto) {
        productos.push(producto);
        filaActual = producto.filaFin + 1; // Saltar al siguiente bloque
      } else {
        filaActual++; // Avanzar una fila si no se detectó producto
      }
    }

    Logger.log("✅ Total de productos detectados: " + productos.length);

    // Calcular totales generales
    var totales = calcularTotales(productos);

    return {
      proyecto: nombreProyecto,
      totalMaterial: totalMaterial,
      productos: productos,
      totales: totales,
      exito: true
    };

  } catch (e) {
    Logger.log("❌ Error parseando Generador: " + e.message);
    return {
      exito: false,
      mensaje: "Error parseando hoja Generador: " + e.message
    };
  }
}

/**
 * Detecta un bloque de producto y extrae su información
 * @param {Array} datos - Matriz de datos de la hoja
 * @param {Number} filaInicio - Índice de fila donde empezar a buscar
 * @returns {Object|null} - Producto estructurado o null si no se encuentra
 */
function detectarYExtraerProducto(datos, filaInicio) {
  try {
    // Buscar el inicio de un bloque de producto
    // Identificadores: "Datos de la Pieza" en A, o "Medidas" en A, o "Categoría" en A
    var fila = filaInicio;
    var inicioBloqueEncontrado = false;
    var tipoProducto = null;

    // Buscar inicio del bloque
    while (fila < datos.length && !inicioBloqueEncontrado) {
      var celdaA = String(datos[fila][0]).trim();

      if (celdaA === "Datos de la Pieza") {
        inicioBloqueEncontrado = true;
        tipoProducto = "PIEZA";
        break;
      } else if (celdaA === "Medidas" && String(datos[fila][3]).trim().toLowerCase().includes("descripción del modelo")) {
        inicioBloqueEncontrado = true;
        tipoProducto = "MEDIDAS";
        break;
      }

      fila++;
    }

    if (!inicioBloqueEncontrado) {
      return null;
    }

    Logger.log("🔍 Producto detectado en fila " + (fila + 1) + " - Tipo: " + tipoProducto);

    // Extraer información según el tipo
    if (tipoProducto === "PIEZA") {
      return extraerProductoTipoPieza(datos, fila);
    } else if (tipoProducto === "MEDIDAS") {
      return extraerProductoTipoMedidas(datos, fila);
    }

    return null;

  } catch (e) {
    Logger.log("⚠️ Error detectando producto en fila " + (filaInicio + 1) + ": " + e.message);
    return null;
  }
}

/**
 * Extrae un producto tipo "Datos de la Pieza" (puertas, ventanas, etc.)
 */
function extraerProductoTipoPieza(datos, filaInicio) {
  var producto = {
    tipo: "PIEZA",
    filaInicio: filaInicio,
    filaFin: filaInicio
  };

  var fila = filaInicio;

  // Extraer campos principales
  // Categoría (siguiente fila, columna A y B)
  if (datos[fila + 1]) {
    producto.categoria = String(datos[fila + 1][1]).trim(); // Columna B
  }

  // ✅ CORREGIDO: Descripción está DEBAJO de "Descripción del Modelo"
  // Buscar la fila que contiene "Descripción del Modelo" y tomar la siguiente
  var descripcionEncontrada = false;
  for (var i = fila; i < Math.min(fila + 10, datos.length); i++) {
    var celdaD = String(datos[i][3]).trim();
    if (celdaD.toLowerCase().includes("descripción del modelo")) {
      // La descripción REAL está en la fila siguiente
      if (datos[i + 1]) {
        producto.descripcion = String(datos[i + 1][3]).trim().replace(/\s+/g, ' ');
        descripcionEncontrada = true;
      }
      break;
    }
  }

  // Fallback si no se encontró el patrón
  if (!descripcionEncontrada && datos[fila]) {
    producto.descripcion = String(datos[fila][3]).trim().replace(/\s+/g, ' ');
  }

  // Modelo (fila +2, columna B)
  if (datos[fila + 2]) {
    producto.modelo = String(datos[fila + 2][1]).trim();
  }

  // Clave (fila +3, columna B)
  if (datos[fila + 3]) {
    producto.clave = String(datos[fila + 3][1]).trim();
  }

  // ✅ NUEVO: Buscar "Piezas" para la cantidad
  producto.piezas = 1; // Default
  for (var i = fila; i < Math.min(fila + 25, datos.length); i++) {
    var celdaA = String(datos[i][0]).trim();
    if (celdaA.toLowerCase() === "piezas") {
      var valorPiezas = datos[i][1]; // Columna B
      if (valorPiezas && !isNaN(parseFloat(valorPiezas))) {
        producto.piezas = parseFloat(valorPiezas);
      }
      break;
    }
  }

  // Precio de Venta (columna K, buscar "Pecio de Venta" o "Precio de Venta")
  for (var i = fila; i < Math.min(fila + 15, datos.length); i++) {
    var celdaK = String(datos[i][10]).trim();
    if (celdaK.toLowerCase().includes("pecio de venta") || celdaK.toLowerCase().includes("precio de venta")) {
      producto.precioVenta = datos[i][11]; // Columna L
      break;
    }
  }

  // Importe (columna K, buscar "Importe")
  for (var i = fila; i < Math.min(fila + 15, datos.length); i++) {
    var celdaK = String(datos[i][10]).trim();
    if (celdaK === "Importe") {
      producto.importe = datos[i][11]; // Columna L
      // ✅ Calcular precio unitario
      if (producto.importe && !isNaN(parseFloat(String(producto.importe).replace(/[$,]/g, '')))) {
        var importeNum = parseFloat(String(producto.importe).replace(/[$,]/g, ''));
        producto.precioUnitario = importeNum / producto.piezas;
      }
      break;
    }
  }

  // Detectar fin del bloque (siguiente fila vacía o siguiente producto)
  fila = filaInicio + 4;
  while (fila < datos.length) {
    var celdaA = String(datos[fila][0]).trim();
    var celdaD = String(datos[fila][3]).trim();

    // Si encontramos otra sección de producto, terminamos
    if (celdaA === "Datos de la Pieza" || celdaA === "Medidas") {
      break;
    }

    // Si encontramos fila completamente vacía
    if (celdaA === "" && datos[fila][1] === "" && celdaD === "") {
      break;
    }

    fila++;
  }

  producto.filaFin = fila - 1;

  Logger.log("  ✓ Categoría: " + producto.categoria);
  Logger.log("  ✓ Descripción: " + (producto.descripcion ? producto.descripcion.substring(0, 50) + "..." : "N/A"));
  Logger.log("  ✓ Precio: " + producto.precioVenta);

  return producto;
}

/**
 * Extrae un producto tipo "Medidas" (cristales, etc.)
 */
function extraerProductoTipoMedidas(datos, filaInicio) {
  var producto = {
    tipo: "MEDIDAS",
    filaInicio: filaInicio,
    filaFin: filaInicio
  };

  var fila = filaInicio;

  // ✅ CORREGIDO: Descripción está DEBAJO de "Descripción del Modelo"
  var descripcionEncontrada = false;
  for (var i = fila; i < Math.min(fila + 10, datos.length); i++) {
    var celdaD = String(datos[i][3]).trim();
    if (celdaD.toLowerCase().includes("descripción del modelo")) {
      // La descripción REAL está en la fila siguiente
      if (datos[i + 1]) {
        producto.descripcion = String(datos[i + 1][3]).trim().replace(/\s+/g, ' ');
        descripcionEncontrada = true;
      }
      break;
    }
  }

  // Fallback
  if (!descripcionEncontrada && datos[fila]) {
    producto.descripcion = String(datos[fila][3]).trim().replace(/\s+/g, ' ');
  }

  // Categoría o clave (si existe en siguientes filas)
  // Para cristales, a veces no hay categoría explícita
  producto.categoria = "CRISTAL"; // Por defecto
  producto.clave = ""; // Puede no tener

  // ✅ NUEVO: Buscar "Piezas" para la cantidad (cristales también pueden tener piezas)
  producto.piezas = 1; // Default
  for (var i = fila; i < Math.min(fila + 25, datos.length); i++) {
    var celdaA = String(datos[i][0]).trim();
    if (celdaA.toLowerCase() === "piezas") {
      var valorPiezas = datos[i][1]; // Columna B
      if (valorPiezas && !isNaN(parseFloat(valorPiezas))) {
        producto.piezas = parseFloat(valorPiezas);
      }
      break;
    }
  }

  // Precio de Venta (columna K)
  for (var i = fila; i < Math.min(fila + 15, datos.length); i++) {
    var celdaK = String(datos[i][10]).trim();
    if (celdaK.toLowerCase().includes("pecio de venta") || celdaK.toLowerCase().includes("precio de venta")) {
      producto.precioVenta = datos[i][11]; // Columna L
      break;
    }
  }

  // Importe
  for (var i = fila; i < Math.min(fila + 15, datos.length); i++) {
    var celdaK = String(datos[i][10]).trim();
    if (celdaK === "Importe") {
      producto.importe = datos[i][11]; // Columna L
      // ✅ Calcular precio unitario
      if (producto.importe && !isNaN(parseFloat(String(producto.importe).replace(/[$,]/g, '')))) {
        var importeNum = parseFloat(String(producto.importe).replace(/[$,]/g, ''));
        producto.precioUnitario = importeNum / producto.piezas;
      }
      break;
    }
  }

  // Detectar fin del bloque
  fila = filaInicio + 4;
  while (fila < datos.length) {
    var celdaA = String(datos[fila][0]).trim();
    var celdaD = String(datos[fila][3]).trim();

    if (celdaA === "Datos de la Pieza" || celdaA === "Medidas") {
      break;
    }

    if (celdaA === "" && datos[fila][1] === "" && celdaD === "") {
      break;
    }

    fila++;
  }

  producto.filaFin = fila - 1;

  Logger.log("  ✓ Tipo: CRISTAL");
  Logger.log("  ✓ Descripción: " + (producto.descripcion ? producto.descripcion.substring(0, 50) + "..." : "N/A"));
  Logger.log("  ✓ Precio: " + producto.precioVenta);

  return producto;
}

/**
 * Calcula totales a partir de los productos
 * @param {Array} productos - Array de productos
 * @param {Boolean} modoPrecioCerrado - Si es true, aplica 13.79% de descuento por partida (Modo B)
 */
function calcularTotales(productos, modoPrecioCerrado) {
  var subtotal = 0;
  var DESCUENTO_IVA = 13.79; // Descuento aplicado en Modo B

  // Si es Modo B, aplicar descuento del 13.79% a cada producto
  if (modoPrecioCerrado) {
    for (var i = 0; i < productos.length; i++) {
      var importe = productos[i].importeFinal !== undefined ? productos[i].importeFinal : productos[i].importe;

      if (typeof importe === "string") {
        importe = parseFloat(importe.replace(/[$,]/g, ""));
      }

      if (!isNaN(importe)) {
        // Aplicar descuento del 13.79% al importe de cada producto
        var descuento = importe * (DESCUENTO_IVA / 100);
        var importeConDescuento = importe - descuento;

        // Guardar el descuento en el producto para mostrarlo en el PDF
        productos[i].descuentoPorcentaje = DESCUENTO_IVA;
        productos[i].descuentoMonto = descuento;
        productos[i].importeFinal = importeConDescuento;

        subtotal += importeConDescuento;
      }
    }
  } else {
    // Modo normal - sin descuentos adicionales
    for (var i = 0; i < productos.length; i++) {
      var importe = productos[i].importeFinal !== undefined ? productos[i].importeFinal : productos[i].importe;

      if (typeof importe === "string") {
        importe = parseFloat(importe.replace(/[$,]/g, ""));
      }

      if (!isNaN(importe)) {
        subtotal += importe;
      }
    }
  }

  var iva = subtotal * 0.16;
  var total = subtotal + iva;

  return {
    subtotal: subtotal,
    iva: iva,
    total: total,
    modoPrecioCerrado: modoPrecioCerrado || false
  };
}

/**
 * Convierte un número a letras en español (para importes)
 * @param {Number} numero - Número a convertir
 * @returns {String} - Número en letras
 */
function numeroALetras(numero) {
  var unidades = ['', 'UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE'];
  var decenas = ['', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
  var especiales = ['DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE'];
  var centenas = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];

  function convertirGrupo(n) {
    if (n === 0) return '';
    if (n < 10) return unidades[n];
    if (n < 20) return especiales[n - 10];
    if (n < 100) {
      var dec = Math.floor(n / 10);
      var uni = n % 10;
      if (uni === 0) return decenas[dec];
      if (dec === 2) return 'VEINTI' + unidades[uni];
      return decenas[dec] + ' Y ' + unidades[uni];
    }
    if (n < 1000) {
      var cen = Math.floor(n / 100);
      var resto = n % 100;
      if (n === 100) return 'CIEN';
      if (resto === 0) return centenas[cen];
      return centenas[cen] + ' ' + convertirGrupo(resto);
    }
    return '';
  }

  if (numero === 0) return 'CERO PESOS';

  var entero = Math.floor(numero);
  var decimales = Math.round((numero - entero) * 100);

  var resultado = '';

  // Millones
  if (entero >= 1000000) {
    var millones = Math.floor(entero / 1000000);
    resultado += (millones === 1 ? 'UN MILLÓN' : convertirGrupo(millones) + ' MILLONES');
    entero = entero % 1000000;
    if (entero > 0) resultado += ' ';
  }

  // Miles
  if (entero >= 1000) {
    var miles = Math.floor(entero / 1000);
    if (miles === 1) {
      resultado += 'MIL';
    } else {
      resultado += convertirGrupo(miles) + ' MIL';
    }
    entero = entero % 1000;
    if (entero > 0) resultado += ' ';
  }

  // Centenas, decenas y unidades
  if (entero > 0) {
    resultado += convertirGrupo(entero);
  }

  resultado += ' PESOS';

  if (decimales > 0) {
    resultado += ' ' + decimales.toString().padStart(2, '0') + '/100 MXN';
  }

  return resultado.trim();
}

/**
 * Función de prueba para verificar el parseo
 */
function probarParser() {
  var resultado = parsearHojaGenerador();

  if (resultado.exito) {
    Logger.log("✅ Parseo exitoso!");
    Logger.log("Proyecto: " + resultado.proyecto);
    Logger.log("Productos encontrados: " + resultado.productos.length);
    Logger.log("Subtotal: $" + resultado.totales.subtotal.toFixed(2));
    Logger.log("IVA: $" + resultado.totales.iva.toFixed(2));
    Logger.log("Total: $" + resultado.totales.total.toFixed(2));
    Logger.log("Total en letras: " + numeroALetras(resultado.totales.total));

    // Mostrar primer producto como ejemplo
    if (resultado.productos.length > 0) {
      Logger.log("\n--- Producto 1 ---");
      Logger.log(JSON.stringify(resultado.productos[0], null, 2));
    }
  } else {
    Logger.log("❌ Error: " + resultado.mensaje);
  }
}
