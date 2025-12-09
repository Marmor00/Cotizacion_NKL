/**
 * Transfiere y formatea los datos de la hoja "Nuevos" a la hoja "Generador".
 * Similar a Cotizacion_Generador pero adaptado para productos nuevos/especiales.
 */
function transferirDatosNuevos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaOrigen = ss.getSheetByName("Nuevos");
    var hojaDestino = ss.getSheetByName("Generador");

    Logger.log("Iniciando transferencia desde Nuevos...");

    // PASO 1: Leer la última fila del destino para saber dónde pegar
    // K1 contiene la fórmula que calcula el máximo entre columna A y columna D
    var ultimaFilaDestino = hojaDestino.getRange("K1").getValue();
    var filaPegado = 4; // Por defecto fila 4

    if (ultimaFilaDestino && ultimaFilaDestino > 0) {
      filaPegado = ultimaFilaDestino + 2; // Saltar 2 filas libres
    }
    Logger.log("Última fila calculada: " + ultimaFilaDestino);
    Logger.log("Pegando en fila: " + filaPegado);

    // PASO 2: Copiar A9:B19 (tabla de información general de Nuevos)
    // A diferencia de Cotizador que usa A1:B17, Nuevos tiene la tabla en A9:B19
    Logger.log("Copiando información general desde A9:B19...");
    hojaOrigen.getRange("A9:B19").copyTo(hojaDestino.getRange(filaPegado, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);

    var rangoDestinoGeneral = hojaDestino.getRange(filaPegado, 1, 11, 2); // 11 filas (A9:B19)
    rangoDestinoGeneral.clearDataValidations();

    var rangoOrigenGeneral = hojaOrigen.getRange("A9:B19");
    var valoresGeneral = rangoOrigenGeneral.getValues();
    rangoDestinoGeneral.setValues(valoresGeneral);

    rangoDestinoGeneral.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    Logger.log("Formato de texto 'Clip' aplicado ✓");

    // Aplicar ajuste de texto 'Wrap' a las filas equivalentes a 12-13 del origen
    // (si es que existe ese patrón en Nuevos)
    var rangoFila12y13 = hojaDestino.getRange(filaPegado + 3, 1, 2, 2);
    rangoFila12y13.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    Logger.log("Ajuste de texto 'Wrap' aplicado a filas específicas ✓");

    rangoDestinoGeneral.setBorder(true, true, true, true, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID);
    Logger.log("Borde exterior aplicado a la tabla general ✓");

    // Copiar altura de fila si es necesario (ajustar según la fila específica de Nuevos)
    var alturaFila13Origen = hojaOrigen.getRowHeight(13);
    var filaDestino13 = filaPegado + 4; // Ajustar según posición relativa
    hojaDestino.setRowHeight(filaDestino13, alturaFila13Origen);
    Logger.log("Alto de fila aplicado ✓");

    // PASO 2.5: Agregar etiqueta "Producto Nuevo" debajo de la tabla izquierda
    var filaEtiqueta = filaPegado + 11; // Debajo de las 11 filas de A9:B19
    var rangoEtiqueta = hojaDestino.getRange(filaEtiqueta, 1, 1, 2); // A y B
    rangoEtiqueta.merge();
    rangoEtiqueta.setValue("PRODUCTO NUEVO");
    rangoEtiqueta.setFontWeight("bold");
    rangoEtiqueta.setFontSize(10);
    rangoEtiqueta.setHorizontalAlignment("center");
    rangoEtiqueta.setBackground("#FFD966"); // Amarillo claro para distinguir
    rangoEtiqueta.setFontColor("#000000");
    rangoEtiqueta.setBorder(true, true, true, true, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID);
    Logger.log("Etiqueta 'PRODUCTO NUEVO' agregada ✓");

    // PASO 3: Copiar K1:L7 (tabla verde y roja, igual que en Cotizador)
    Logger.log("Copiando tabla adicional K1:L7...");
    hojaOrigen.getRange("K1:L7").copyTo(hojaDestino.getRange(filaPegado, 11), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);

    var rangoDestinoAdicional = hojaDestino.getRange(filaPegado, 11, 7, 2);
    rangoDestinoAdicional.clearDataValidations();

    var rangoOrigenAdicional = hojaOrigen.getRange("K1:L7");
    var valoresAdicional = rangoOrigenAdicional.getValues();
    rangoDestinoAdicional.setValues(valoresAdicional);

    var rangoColumnaL = hojaDestino.getRange(filaPegado, 12, 7, 1); // Columna L para formato de moneda
    rangoColumnaL.setNumberFormat("$#,##0.00");
    Logger.log("Tabla adicional K:L copiada ✓");

    // Eliminar bordes en la fila del medio (fila 4 de la tabla K:L)
    var filaMedia = filaPegado + 3;
    var rangoBordeMedio = hojaDestino.getRange(filaMedia, 11, 1, 2);
    rangoBordeMedio.setBorder(false, false, false, false, false, false);
    Logger.log("Bordes de la fila intermedia eliminados ✓");

    // PASO 4: Copiar la tabla de materiales D1:I (igual que en Cotizador)
    var ultimaFilaRealOrigen = hojaOrigen.getRange("J1").getValue();

    if (ultimaFilaRealOrigen && ultimaFilaRealOrigen >= 1) {
      Logger.log("Copiando materiales desde D1:I" + ultimaFilaRealOrigen);
      var rangoMaterialesOrigen = hojaOrigen.getRange("D1:I" + ultimaFilaRealOrigen);
      rangoMaterialesOrigen.copyTo(hojaDestino.getRange(filaPegado, 4), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);

      var valoresMateriales = rangoMaterialesOrigen.getValues();
      var rangoDestinoMateriales = hojaDestino.getRange(filaPegado, 4, valoresMateriales.length, valoresMateriales[0].length);
      rangoDestinoMateriales.clearDataValidations();

      // Formatear la columna de fórmula (columna H)
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

    Logger.log("Transferencia desde Nuevos completada");
    return "Éxito - Pegado en fila " + filaPegado;

  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    return "Error: " + e.message;
  }
}

/**
 * Limpia la hoja Generador (misma función que en Cotizacion_Generador)
 * Nota: Esta función ya existe en Cotizacion_Generador, no es necesario duplicarla
 */
function limpiarHojaGeneradorNuevos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaDestino = ss.getSheetByName("Generador");
    hojaDestino.clear();
    // Nueva fórmula: calcula el máximo entre última fila de columna A y columna D
    hojaDestino.getRange("K1").setFormula('=MAX(FILTER(ROW(A:A), A:A<>""), FILTER(ROW(D:D), D:D<>""))');
    return "Hoja limpiada";
  } catch (e) {
    return "Error: " + e.message;
  }
}
