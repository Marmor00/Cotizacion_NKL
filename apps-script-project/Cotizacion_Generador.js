/**
 * Transfiere y formatea los datos de la hoja "Cotizador" a la hoja "Generador".
 * VERSIÓN 4.2: Se extiende la tabla de materiales a la columna J y se aplica ajuste de texto a las filas 12-13.
 */
/**
 * Transfiere y formatea los datos de la hoja "Cotizador" a la hoja "Generador".
 * VERSIÓN 4.3: El rango de materiales ahora es D:I y la tabla lateral se mueve a K:L.
 */
function transferirDatosCotizador() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaOrigen = ss.getSheetByName("Cotizador");
    var hojaDestino = ss.getSheetByName("Generador");

    Logger.log("Iniciando transferencia optimizada...");

    // PASO 1: Leer la última fila del destino para saber dónde pegar
    var ultimaFilaColumnaD = hojaDestino.getRange("A1").getValue();
    var filaPegado = 4; // Por defecto fila 4

    if (ultimaFilaColumnaD && ultimaFilaColumnaD > 0) {
      filaPegado = ultimaFilaColumnaD + 2; // Saltar 2 filas libres
    }
    Logger.log("Pegando en fila: " + filaPegado);

    // PASO 2: Copiar A1:B17, aplicar bordes y formato de texto (SIN CAMBIOS)
    Logger.log("Copiando información general...");
    hojaOrigen.getRange("A1:B17").copyTo(hojaDestino.getRange(filaPegado, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
    
    var rangoDestinoGeneral = hojaDestino.getRange(filaPegado, 1, 17, 2); 
    rangoDestinoGeneral.clearDataValidations();
    
    var rangoOrigenGeneral = hojaOrigen.getRange("A1:B17");
    var valoresGeneral = rangoOrigenGeneral.getValues();
    rangoDestinoGeneral.setValues(valoresGeneral);

    rangoDestinoGeneral.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    Logger.log("Formato de texto 'Clip' aplicado ✓");
    
    var rangoFila12y13 = hojaDestino.getRange(filaPegado + 11, 1, 2, 2);
    rangoFila12y13.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    Logger.log("Ajuste de texto 'Wrap' aplicado a las filas 12-13 ✓");

    rangoDestinoGeneral.setBorder(true, true, true, true, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID);
    Logger.log("Borde exterior aplicado a la tabla general ✓");

    var alturaFila13 = hojaOrigen.getRowHeight(13); 
    var filaDestino13 = filaPegado + 12; 
    hojaDestino.setRowHeight(filaDestino13, alturaFila13); 
    Logger.log("Alto de la fila 13 (" + alturaFila13 + "px) aplicado en la fila " + filaDestino13 + " ✓");

    // PASO 3: Copiar K1:L7 y corregir bordes
    Logger.log("Copiando tabla adicional K1:L7..."); // <-- CAMBIO: Mensaje actualizado
    // <-- CAMBIO: El rango de origen ahora es K1:L7 y el destino empieza en la columna 11 (K)
    hojaOrigen.getRange("K1:L7").copyTo(hojaDestino.getRange(filaPegado, 11), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
    
    var rangoDestinoAdicional = hojaDestino.getRange(filaPegado, 11, 7, 2); // <-- CAMBIO: Columna 11
    rangoDestinoAdicional.clearDataValidations();
    
    var rangoOrigenAdicional = hojaOrigen.getRange("K1:L7"); // <-- CAMBIO: Rango K1:L7
    var valoresAdicional = rangoOrigenAdicional.getValues();
    rangoDestinoAdicional.setValues(valoresAdicional);
    
    var rangoColumnaL = hojaDestino.getRange(filaPegado, 12, 7, 1); // <-- CAMBIO: Columna 12 (L) para el formato de moneda
    rangoColumnaL.setNumberFormat("$#,##0.00");
    Logger.log("Tabla adicional K:L copiada ✓");

    // Eliminar bordes en la fila del medio
    var filaMedia = filaPegado + 3; 
    var rangoBordeMedio = hojaDestino.getRange(filaMedia, 11, 1, 2); // <-- CAMBIO: Columna 11
    rangoBordeMedio.setBorder(false, false, false, false, false, false);
    Logger.log("Bordes de la fila intermedia eliminados ✓");

    // PASO 4: Copiar la tabla de materiales
    var ultimaFilaRealOrigen = hojaOrigen.getRange("J1").getValue();
    
    if (ultimaFilaRealOrigen && ultimaFilaRealOrigen >= 1) { 
      // <-- CAMBIO: El rango ahora es D1:I...
      Logger.log("Copiando materiales desde D1:I" + ultimaFilaRealOrigen);
      var rangoMaterialesOrigen = hojaOrigen.getRange("D1:I" + ultimaFilaRealOrigen); // <-- CAMBIO: Rango D1:I
      rangoMaterialesOrigen.copyTo(hojaDestino.getRange(filaPegado, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);

      var valoresMateriales = rangoMaterialesOrigen.getValues();
      var rangoDestinoMateriales = hojaDestino.getRange(filaPegado, 4, valoresMateriales.length, valoresMateriales[0].length);
      rangoDestinoMateriales.clearDataValidations();

      // Este bucle para formatear la columna de la fórmula sigue funcionando
      // porque el índice [4] corresponde a la columna H dentro del rango D:I
      for (var i = 0; i < valoresMateriales.length; i++) {
        var celdaFormula = valoresMateriales[i][4]; // Columna H (Fórmula)
        if (typeof celdaFormula === 'string' && celdaFormula !== "") {
          valoresMateriales[i][4] = "'" + celdaFormula.toString().replace(/'/g, "");
        }
      }

      rangoDestinoMateriales.setValues(valoresMateriales);

      var ultimaFilaTabla = filaPegado + valoresMateriales.length - 1;
      var rangoUltimaFila = hojaDestino.getRange(ultimaFilaTabla, 4, 1, valoresMateriales[0].length);
      rangoUltimaFila.setBorder(null, null, true, null, null, null, "#666666", SpreadsheetApp.BorderStyle.SOLID);
      Logger.log("Borde inferior gris aplicado a la tabla de materiales ✓");
    }

    Logger.log("Transferencia completada");
    return "Éxito - Pegado en fila " + filaPegado;

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    return "Error: " + e.message;
  }
}

// --- Esta función no necesita cambios ---
function limpiarHojaGenerador() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaDestino = ss.getSheetByName("Generador");
    hojaDestino.clear();
    hojaDestino.getRange("J1").setFormula('=MAX(FILTER(ROW(D:D), D:D<>""))');
    return "Hoja limpiada";
  } catch (e) {
    return "Error: " + e.message;
  }
}