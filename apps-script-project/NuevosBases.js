/**
 * Script de DEMOSTRACIÓN.
 * Guarda un nuevo producto desde la hoja "Nuevos" en rangos
 * temporales de "Productos" (R2) y "Generadores" (G2).
 *
 * VERSIÓN 1.2:
 * - Lee las fórmulas de la Col H como TEXTO y les añade un apóstrofe (').
 * - Lee todos los materiales de la lista, no solo uno.
 */
function NuevosBases() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var sheetNuevos = ss.getSheetByName("Nuevos");
  var sheetProductos = ss.getSheetByName("Productos");
  var sheetGeneradores = ss.getSheetByName("Generadores");

  if (!sheetNuevos || !sheetProductos || !sheetGeneradores) {
    ui.alert("Error", "Faltan hojas 'Nuevos', 'Productos' o 'Generadores'.", ui.ButtonSet.OK);
    return;
  }

  // --- 1. Leer Datos Generales (de "Nuevos") ---
  
  // ▼▼▼ ¡IMPORTANTE! ▼▼▼
  // AÚN FALTAN ESTAS CELDAS. REEMPLAZA "???"
  var marca = sheetNuevos.getRange("L10").getValue();      // ¿Celda para Marca?
  var familia = sheetNuevos.getRange("L11").getValue();    // ¿Celda para Familia?
  var especiales = sheetNuevos.getRange("A14").getValue(); // ¿Celda para Especiales?
  
  var descripcion = sheetNuevos.getRange("D2").getValue(); 
  
  // Leemos las 5 claves de L10:L14
  var datosClave = sheetNuevos.getRange("L12:L16").getValues();
  var serie = datosClave[0][0];       // L10
  var funcion = datosClave[1][0];     // L11
  var d1 = datosClave[2][0];          // L12
  var d2 = datosClave[3][0];          // L13
  var anotacion = datosClave[4][0];   // L14
  
  if (serie === "" || funcion === "" || d1 === "" || d2 === "") {
    ui.alert("Datos Incompletos", "Faltan datos clave en el rango L10:L14.", ui.ButtonSet.OK);
    return;
  }

  // --- 2. Preparar datos para la hoja "Productos" (Destino R2:AA2) ---
  
  var codigo = [serie, funcion, d1, d2, anotacion].filter(Boolean).join('-');

  var filaProducto = [
    codigo,       // R
    marca,        // S
    familia,      // T
    serie,        // U
    funcion,      // V
    d1,           // W
    d2,           // X
    anotacion,    // Y
    descripcion,  // Z (Leído de D2)
    especiales    // AA
  ];

  // --- 3. Preparar datos para la hoja "Generadores" (Destino G2:K...) ---
  
  var rangoReceta = sheetNuevos.getRange("E10:I" + sheetNuevos.getLastRow());
  
  // ▼▼▼ CAMBIO: Leemos FÓRMULAS y VALORES por separado ▼▼▼
  var datosRecetaValores = rangoReceta.getValues();
  var datosRecetaFormulas = rangoReceta.getFormulas();
  
  var filasGeneradores = [];

  for (var i = 0; i < datosRecetaValores.length; i++) {
    
    // Leemos los valores
    var unidad = datosRecetaValores[i][0];  // Col E (índice 0)
    var descMat = datosRecetaValores[i][1]; // Col F (índice 1)
    var pu = datosRecetaValores[i][4];      // Col I (índice 4)

    // ▼▼▼ CAMBIO: Lógica para fórmulas como texto ▼▼▼
    var formulaStr = datosRecetaFormulas[i][3]; // Texto de la fórmula (Col H)
    var formulaVal = datosRecetaValores[i][3];  // Valor (Col H)

    // Si la fórmula y el valor están vacíos, es el fin de la lista
    if (formulaStr === "" && (formulaVal === "" || formulaVal === 0)) {
      break; 
    }
    
    var formulaParaGuardar = "";
    if (formulaStr.startsWith("=")) {
      // Si es una fórmula (empieza con =), le añadimos un apóstrofe
      formulaParaGuardar = "'" + formulaStr;
    } else {
      // Si no, es un valor simple (ej. 4), lo guardamos tal cual
      formulaParaGuardar = formulaVal;
    }
    // ▲▲▲ FIN DEL CAMBIO ▲▲▲
    
    
    // Armar la fila de 5 columnas para "Generadores" (G:K)
    filasGeneradores.push([
      codigo, // La clave-modelo concatenada
      formulaParaGuardar, // <--- Ya viene como texto
      unidad,
      descMat,
      pu
    ]);
  }

  if (filasGeneradores.length === 0) {
    ui.alert("Receta Vacía", "No se encontraron materiales en la lista (E10:I...). Asegúrate de agregar materiales y sus FÓRMULAS (Col H).", ui.ButtonSet.OK);
    return;
  }

  // --- 4. Confirmación y Guardado ---
  var confirm = ui.alert(
    "Confirmar Guardado (DEMO)",
    "Se va a guardar el producto:\n\nCÓDIGO: " + codigo + "\nMATERIALES: " + filasGeneradores.length + "\n\nEsto se guardará en los rangos temporales (G2 y R2) para demostración.\n\n¿Deseas continuar?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  try {
    // Limpiar rangos de destino antes de pegar
    sheetProductos.getRange("R2:AA2").clearContent();
    var lastRowGen = sheetGeneradores.getLastRow();
    if (lastRowGen >= 2) {
      sheetGeneradores.getRange("G2:K" + lastRowGen).clearContent();
    }

    // Guardar en "Productos" (en la fila 2, R:AA)
    sheetProductos.getRange("R2:AA2").setValues([filaProducto]);

    // Guardar en "Generadores" (a partir de G2)
    sheetGeneradores.getRange(
      2, // Fila 2
      7, // Columna G
      filasGeneradores.length,
      5  // 5 columnas (G:K)
    ).setValues(filasGeneradores);

    ui.alert("¡Éxito!", "El producto '" + codigo + "' ha sido guardado en los rangos de demostración (G2 y R2).", ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ui.alert("Error al guardar: " + e.message);
  }
}