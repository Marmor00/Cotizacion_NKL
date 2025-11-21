function borrarDatosCotizador() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCotizador = ss.getSheetByName("Cotizador"); // Asegúrate de que el nombre coincide exactamente

  if (!sheetCotizador) {
    Logger.log("No se encontró la hoja Cotizador");
    return;
  }

  // Borrar solo el rango específico A7:I28
  sheetCotizador.getRange("D8:H80").clearContent();
  sheetCotizador.getRange("B2:B5").clearContent();
  sheetCotizador.getRange("A9:B11").clearContent();
  sheetCotizador.getRange("D2:I5").clearContent();
  Logger.log("Datos borrados correctamente en el rango A7:I28 de cotizador_PBI.");
}