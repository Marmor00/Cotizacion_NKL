/**
 * Traduce las fórmulas de la columna D a un formato de texto con variables
 * personalizadas (Largo, Alto) en la columna H.
 * VERSIÓN 1.1: Agrega una comilla simple (') al inicio para forzar
 * el resultado como TEXTO y evitar el error #NAME?
 */
function traducirFormulasLargoAlto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet(); 
  var ui = SpreadsheetApp.getUi();

  var ultimaFila = sheet.getRange("J1").getValue();

  if (typeof ultimaFila !== 'number' || ultimaFila < 10) {
    ui.alert("Aviso", "Asegúrate de que la celda J1 contenga un número válido y que sea 10 o mayor.", ui.ButtonSet.OK);
    return;
  }

  var rangoOrigen = sheet.getRange("D10:D" + ultimaFila);
  var rangoDestino = sheet.getRange("H10:H" + ultimaFila);
  var formulasOriginales = rangoOrigen.getFormulas();
  var formulasTraducidas = [];

  for (var i = 0; i < formulasOriginales.length; i++) {
    var formula = formulasOriginales[i][0];
    
    if (formula === "") {
      formulasTraducidas.push([""]);
      continue;
    }

    var formulaTraducida = formula
      .replace(/A11/gi, "Largo1")
      .replace(/A12/gi, "Largo2")
      .replace(/A13/gi, "Largo3")
      .replace(/B11/gi, "Alto1")
      .replace(/B12/gi, "Alto2")
      .replace(/B13/gi, "Alto3");

    // --- ¡ESTE ES EL CAMBIO! ---
    // Le añadimos una comilla simple al principio para que se guarde como texto.
    formulasTraducidas.push(["'" + formulaTraducida]);
  }

  rangoDestino.setValues(formulasTraducidas);
  ui.alert("Proceso completado", "Las fórmulas han sido traducidas en la columna H.", ui.ButtonSet.OK);
}