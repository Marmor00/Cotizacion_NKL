/**
 * Extrae datos de productos desde varias hojas de cálculo basado en un ID de producto.
 * Utiliza una lógica de "Clave Universal" para buscar precios de materiales.
 *
 * VERSIÓN 5.2 (CORREGIDA):
 * - La hoja "Generadores" ("sheetLimpio") ahora se lee con la estructura A:E
 * (Código, Fórmula, Unidad, Material).
 * - La hoja "Productos" se mantiene sin cambios (para buscar descripción especial).
 * - Corregidos errores de sintaxis por caracteres extraños.
 */
function ExtraerDatosCotizador() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCotizador = ss.getSheetByName("Cotizador");
  var sheetLimpio = ss.getSheetByName("Generadores"); // <-- Esta es tu hoja A:E
  var sheetAluminio = ss.getSheetByName("Aluminio");
  var sheetVidrio = ss.getSheetByName("Cristales");
  var sheetHerrajes = ss.getSheetByName("Herrajes");
  var sheetProductos = ss.getSheetByName("Productos"); // <-- Esta se queda igual (para Desc. Especial)
  var sheetOtros = ss.getSheetByName("Otros");

  // Verificar que todas las hojas existan (se mantiene como V5.0)
  if (!sheetCotizador || !sheetLimpio || !sheetAluminio || !sheetVidrio || !sheetHerrajes || !sheetProductos || !sheetOtros) {
    Logger.log("❌ Error: Una o más hojas no existen (Cotizador, Generadores, Aluminio, Cristales, Herrajes, Productos, Otros)");
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var valorBuscado = sheetCotizador.getRange("B3").getValue().toString().trim(); // ID producto
  var valorAcabado = sheetCotizador.getRange("B5").getValue().toString().trim(); // Acabado
  var vidrioPersonalizado = sheetCotizador.getRange("B6").getValue().toString().trim(); // Vidrio específico

  if (valorBuscado === "") {
    sheetCotizador.getRange("D8:H58").clearContent();
    return;
  }
  if (valorAcabado === "") {
    ui.alert('Atención', 'Por favor, selecciona un acabado antes de continuar.', ui.ButtonSet.OK);
    return;
  }

  Logger.log("🚀 Iniciando cotización para ID: '" + valorBuscado + "' | Acabado: '" + valorAcabado + "' | Vidrio: '" + vidrioPersonalizado + "'");

  // =========================================================================
  // Bloque para buscar la descripción de medidas especiales (SIN CAMBIOS)
  // Sigue buscando en la hoja "Productos"
  // =========================================================================
  Logger.log("🔍 Buscando descripción especial en la hoja 'Productos'...");
  var datosProductos = sheetProductos.getRange("B2:L" + sheetProductos.getLastRow()).getValues();
  var descripcionEspecial = "Medidas regulares";

  for (var k = 0; k < datosProductos.length; k++) {
    var codigoProducto = datosProductos[k][0] ? datosProductos[k][0].toString().trim() : "";
    if (codigoProducto === valorBuscado) {
      var valorColumnaL = datosProductos[k][10];
      if (valorColumnaL && valorColumnaL.toString().trim() !== "") {
        descripcionEspecial = valorColumnaL;
        ui.alert('¡Atención!', 'Estás por cotizar un modelo especial. Revisa con cuidado el recuadro naranja.', ui.ButtonSet.OK);
      }
      Logger.log("✅ Coincidencia encontrada. Descripción especial: '" + descripcionEspecial + "'");
      break;
    }
  }
  // =========================================================================

  // --- Carga de Datos y Creación de Mapa Optimizado ---

  // --- CAMBIO ---
  // "datosLimpio" (de "Generadores") ahora se lee asumiendo 5 columnas (A:E).
  // Se asume que la Fila 1 es encabezado, por eso empieza en A2.
  var datosLimpio = sheetLimpio.getRange("A2:E" + sheetLimpio.getLastRow()).getValues();

  // El resto de las cargas de datos no cambia
  var datosAluminio = sheetAluminio.getDataRange().getValues();
  var datosVidrio = sheetVidrio.getRange("A1:F" + sheetVidrio.getLastRow()).getValues();
  var datosHerrajes = sheetHerrajes.getRange("A2:F" + sheetHerrajes.getLastRow()).getValues();
  var datosOtros = sheetOtros.getDataRange().getValues();
  var mapaPreciosOtros = crearMapaDePrecios(datosOtros);

  // Buscar columna de acabado en Aluminio (Con trim para evitar espacios extra del QUERY)
  Logger.log("🔍 Buscando acabado: '" + valorAcabado + "' (longitud: " + valorAcabado.length + ")");

  var encabezadosAluminio = datosAluminio[0].map(function(header) {
    return header ? header.toString().trim() : "";
  });

  Logger.log("🔍 Encabezados después de trim: [" + encabezadosAluminio.map(function(h) {
    return "'" + h + "'(" + h.length + ")";
  }).join(", ") + "]");

  var idxAcabadoAluminio = encabezadosAluminio.indexOf(valorAcabado);
  if (idxAcabadoAluminio === -1) {
    Logger.log("❌ No se encontró la columna de acabado '" + valorAcabado + "' en la hoja 'Aluminio'");
    ui.alert('Error', 'El acabado "' + valorAcabado + '" no se encontró en la base de datos de Aluminios.\n\nEncabezados disponibles:\n' + encabezadosAluminio.join(", "), ui.ButtonSet.OK);
    return;
  }
  Logger.log("✅ Acabado encontrado en columna índice: " + idxAcabadoAluminio);

  // --- Búsqueda de materiales y precios ---
  Logger.log("🔍 Buscando materiales del producto...");
  var resultados = [];
  var seNecesitaVidrio = false;

  // --- CAMBIO ---
  // El bucle ahora itera sobre "datosLimpio" (la hoja "Generadores" A:E)
  // El bucle empieza en 0 porque getRange("A2:E...") ya omitió el encabezado
  for (var i = 0; i < datosLimpio.length; i++) {
    // Compara Columna A (índice 0)
    if (datosLimpio[i][0].toString().trim() === valorBuscado) {

      // --- CAMBIO EN ÍNDICES ---
      // (Originalmente eran 7, 8, 9)
      var formulaOriginal = datosLimpio[i][1] || ""; // Col B (índice 1)
      var unidad = datosLimpio[i][2] ? datosLimpio[i][2].toString().trim().toLowerCase() : ""; // Col C (índice 2)
      var material = datosLimpio[i][3] ? datosLimpio[i][3].toString().trim() : ""; // Col D (índice 3)
      // La Col E (P.U.) se ignora, ya que el script calcula el costo dinámicamente.

      var costo = 0;
      var materialFinal = material;

      // --- LÓGICA DE BÚSQUEDA UNIFICADA (Sin cambios) ---
      if (unidad === "perfil por ml") {
        // Caso 1: Aluminio
        costo = buscarPrecioAluminio(material, datosAluminio, idxAcabadoAluminio);

      } else if (unidad === "m2") {
        // Caso 2: Cristal (Lógica Nueva)
        seNecesitaVidrio = true;
        costo = buscarPrecioVidrioPersonalizado(vidrioPersonalizado, datosVidrio);
        materialFinal = vidrioPersonalizado;

      } else {
        // Caso 3: Herrajes, Otros, etc. (Búsqueda Universal)
        var claveBuscada = material.split(' ')[0];
        costo = buscarPrecioUniversal(claveBuscada, datosHerrajes, datosVidrio, mapaPreciosOtros);
      }

      var cantidadCalculada = calcularFormula(formulaOriginal);
      resultados.push([cantidadCalculada, unidad, materialFinal, costo, "'" + formulaOriginal]);
    }
  }

  // Si el producto requiere vidrio pero no se especificó uno, se detiene la ejecución.
  if (seNecesitaVidrio && vidrioPersonalizado === "") {
      ui.alert('Acción Requerida', 'Este producto requiere un tipo de vidrio. Por favor, selecciónalo en la celda B6.', ui.ButtonSet.OK);
      Logger.log("⚠️ Operación detenida: El producto necesita vidrio pero la celda B6 está vacía.");
      return;
  }


  // --- Actualizar resultados en cotizador (sin cambios) ---
  if (resultados.length > 0) {
    var numFilas = resultados.length;
    sheetCotizador.getRange("D8:H" + (7 + numFilas + 20)).clearContent(); // Limpieza más segura
    sheetCotizador.getRange(8, 4, numFilas, 5).setValues(resultados);
    Logger.log("✅ Cotización actualizada con " + numFilas + " materiales.");
  } else {
    Logger.log("⚠️ No se encontraron materiales para el producto: " + valorBuscado);
    sheetCotizador.getRange("D8:H58").clearContent();
  }
}

// ==================================================================================
// FUNCIONES AUXILIARES (DE AYUDA)
// (Estas funciones no necesitan cambios)
// ==================================================================================

/**
 * Busca un precio de vidrio por su descripción exacta.
 * @param {string} descripcion La descripción completa del vidrio a buscar (de la celda B6).
 * @param {Array<Array<any>>} datosVidrio Los datos de la hoja "Cristales".
 * @returns {any} El precio encontrado o una cadena de error.
*/
function buscarPrecioVidrioPersonalizado(descripcion, datosVidrio) {
  if (!descripcion) return "Elegir Vidrio"; // Si la celda B6 está vacía

  Logger.log("🔎 Búsqueda Específica de Vidrio: '" + descripcion + "'");
  // Recorremos buscando en la columna D (índice 3) la descripción exacta
  for (var i = 1; i < datosVidrio.length; i++) { // Empezar en 1 para saltar encabezado
    var descripcionEnHoja = datosVidrio[i][3] ? datosVidrio[i][3].toString().trim() : "";
    if (descripcionEnHoja === descripcion) {
      Logger.log("✅ Vidrio encontrado. Precio: " + datosVidrio[i][5]);
      return datosVidrio[i][5]; // Retorna el precio de la columna F (índice 5)
    }
  }
  Logger.log("❌ No se encontró el vidrio '" + descripcion + "' en la hoja 'Cristales'.");
  return "NE - Vidrio"; // No Encontrado
}


/**
* Busca un precio usando una clave, siguiendo un orden de prioridad:
 * 1. Hoja de Herrajes, 2. Hoja de Cristales (para casos residuales), 3. Mapa de "Otros"
*/
function buscarPrecioUniversal(claveBuscada, datosHerrajes, datosVidrio, mapaOtros) {
  Logger.log("🔎 Búsqueda Universal para clave: '" + claveBuscada + "'");

  // 1. Buscar en Herrajes (Columna C para la clave, F para el precio)
  for (var i = 0; i < datosHerrajes.length; i++) {
    if (datosHerrajes[i][2] && datosHerrajes[i][2].toString().trim() === claveBuscada) {
      Logger.log("✅ Encontrado en 'Herrajes'. Precio: " + datosHerrajes[i][5]);
      return datosHerrajes[i][5];
    }
  }

  // 2. Buscar en Cristales (Columna C para la clave, F para el precio)
  for (var j = 1; j < datosVidrio.length; j++) {
    if (datosVidrio[j][2] && datosVidrio[j][2].toString().trim() === claveBuscada) {
      Logger.log("✅ Encontrado en 'Cristales'. Precio: " + datosVidrio[j][5]);
      return datosVidrio[j][5];
    }
  }

  // 3. Buscar en el mapa de "Otros" (Búsqueda casi instantánea)
  if (mapaOtros.hasOwnProperty(claveBuscada)) {
    Logger.log("✅ Encontrado en 'Otros'. Precio: " + mapaOtros[claveBuscada]);
    return mapaOtros[claveBuscada];
  }

  Logger.log("❌ No se encontró '" + claveBuscada + "' en ninguna hoja.");
  return "NE"; // No Encontrado
}


/**
 * Crea un mapa de búsqueda rápida (objeto) a partir de los datos de "Otros".
 */
function crearMapaDePrecios(datos) {
  var mapa = {};
  for (var i = 1; i < datos.length; i++) {
    var clave = datos[i][2] ? datos[i][2].toString().trim() : "";
    var precio = datos[i][5];
    if (clave !== "") {
      mapa[clave] = precio;
    }
  }
  Logger.log("📊 Mapa de precios 'Otros' creado con " + Object.keys(mapa).length + " entradas.");
  return mapa;
}


/**
 * Busca el precio del aluminio basado en un acabado específico.
 */
function buscarPrecioAluminio(material, datosAluminio, idxAcabado) {
  var claveMaterialMatch = material.match(/\b\d{4,5}\b/);
  var claveMaterial = claveMaterialMatch ? claveMaterialMatch[0] : "";
  if (claveMaterial === "") return "N/A - Sin clave";

  for (var j = 1; j < datosAluminio.length; j++) {
    var claveAluminio = datosAluminio[j][3] ? datosAluminio[j][3].toString().trim() : "";
    if (claveAluminio === claveMaterial) {
      return datosAluminio[j][idxAcabado];
    }
  }
  return "NE";
}


/**
 * Evalúa una fórmula matemática reemplazando variables por valores de celdas.
 */
function calcularFormula(formula) {
  if (typeof formula !== "string" || formula.trim() === "") return formula;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Cotizador");

    var valorLargo1 = hoja.getRange("A9").getValue() || 0;
    var valorAlto1 = hoja.getRange("B9").getValue() || 0;
    var valorLargo2 = hoja.getRange("A10").getValue() || valorLargo1;
    var valorAlto2 = hoja.getRange("B10").getValue() || valorAlto1;
    var valorDesp = hoja.getRange("B15").getValue() || 1;

    let formulaProcesada = formula.toString().replace(/^=/, "");
    formulaProcesada = formulaProcesada
      .replace(/\bLargo1\b/g, valorLargo1)
      .replace(/\bAlto1\b/g, valorAlto1)
      .replace(/\bLargo2\b/g, valorLargo2)
      .replace(/\bAlto2\b/g, valorAlto2)
      .replace(/\bDesp\b/g, valorDesp);

    var resultado = Function('"use strict"; return (' + formulaProcesada + ')')();
    return Math.round(resultado * 100) / 100;
  } catch (e) {
    Logger.log(`Error al calcular fórmula: ${formula} - ${e.message}`);
    return formula;
  }
}
