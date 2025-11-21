/**
 * Busca un material basado en los criterios de la hoja "Nuevos",
 * determina su hoja de origen usando "Aux" como índice, extrae los
 * datos correspondientes (unidad, descripción y costo) y los agrega
 * a la lista a partir de la celda E10.
 *
 * VERSIÓN 1.2:
 * - B4 (Clave Combinada) se usa para buscar en Col F de "Aux".
 * - De la fila encontrada en "Aux", se extrae la Clave Real de Col D.
 * - Esta Clave Real se usa para buscar en las hojas de materiales.
 */
function buscarYAgregarMaterial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNuevos = ss.getSheetByName("Nuevos");
  var sheetAux = ss.getSheetByName("Aux");

  // Verificar que las hojas principales existan
  if (!sheetNuevos || !sheetAux) {
    SpreadsheetApp.getUi().alert("Error", "No se encontraron las hojas 'Nuevos' o 'Aux'. Verifica sus nombres.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // --- 1. Leer los datos de búsqueda de la hoja "Nuevos" ---
  var familia = sheetNuevos.getRange("B2").getValue().toString().trim();
  var serie = sheetNuevos.getRange("B3").getValue().toString().trim();
  // Esta es la clave "combinada" que se usa para buscar en Aux
  var claveCombinada = sheetNuevos.getRange("B4").getValue().toString().trim(); 
  var acabadoAluminio = sheetNuevos.getRange("B5").getValue().toString().trim();

  if (claveCombinada === "") {
    SpreadsheetApp.getUi().alert("Información", "Por favor, ingresa una clave en la celda B4 para buscar.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  Logger.log("🔎 Iniciando búsqueda en 'Aux' para Clave Combinada: '" + claveCombinada + "'");

  // --- 2. Determinar la hoja de origen y la Clave Real ---
  var datosAux = sheetAux.getRange("A4:F" + sheetAux.getLastRow()).getValues(); 
  var hojaOrigen = "";
  var claveDeBusqueda = ""; // <-- Variable para la Clave Real (de Col D)

  // Convertimos los valores de búsqueda a minúsculas
  var familiaLower = familia.toLowerCase();
  var serieLower = serie.toLowerCase();
  var claveCombinadaLower = claveCombinada.toLowerCase();

  for (var i = 0; i < datosAux.length; i++) {
    var auxFamilia = datosAux[i][1] ? datosAux[i][1].toString().trim().toLowerCase() : ""; // Col B (índice 1)
    var auxSerie = datosAux[i][2] ? datosAux[i][2].toString().trim().toLowerCase() : "";   // Col C (índice 2)
    var auxClaveCombinada = datosAux[i][5] ? datosAux[i][5].toString().trim().toLowerCase() : "";   // Col F (índice 5)

    // Comparamos los tres criterios (B, C y F)
    if (auxFamilia === familiaLower && auxSerie === serieLower && auxClaveCombinada === claveCombinadaLower) {
      hojaOrigen = datosAux[i][0].toString().trim(); // Col A: Nombre de la hoja
      
      // ▼▼▼ CAMBIO IMPORTANTE ▼▼▼
      // Obtenemos la Clave Real de la Columna D (índice 3)
      claveDeBusqueda = datosAux[i][3] ? datosAux[i][3].toString().trim() : ""; 
      break;
    }
  }

  if (hojaOrigen === "") {
    SpreadsheetApp.getUi().alert("Error de Búsqueda", "No se pudo determinar la hoja de origen para la combinación de Familia, Serie y Clave proporcionada. Revisa la hoja 'Aux'.", SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("❌ No se encontró el origen en la hoja 'Aux'.");
    return;
  }
  if (claveDeBusqueda === "") {
    SpreadsheetApp.getUi().alert("Error", "Se encontró la fila en 'Aux', pero la clave de búsqueda (Columna D de 'Aux') está vacía.", SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("❌ Fila encontrada en 'Aux' pero la Col D (Clave Real) estaba vacía.");
    return;
  }

  Logger.log("✅ Hoja de origen: '" + hojaOrigen + "'. Usando Clave Real: '" + claveDeBusqueda + "'");

  // --- 3. Buscar el material y extraer sus datos ---
  var sheetMaterial = ss.getSheetByName(hojaOrigen);
  if (!sheetMaterial) {
    SpreadsheetApp.getUi().alert("Error", "La hoja '" + hojaOrigen + "' no existe en este archivo.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var datosMaterial;
  var unidad = "";
  var descripcion = "";
  var costo = "NE"; 

  var datosHojaMaterial = sheetMaterial.getDataRange().getValues();
  var filaEncontrada = -1;
  var claveDeBusquedaLower = claveDeBusqueda.toLowerCase(); // <-- Clave Real en minúsculas

  // El bucle busca la Clave Real en la columna correspondiente
  for (var j = 1; j < datosHojaMaterial.length; j++) {
    var claveEnHoja = "";
    if (hojaOrigen === "Aluminio") {
      claveEnHoja = datosHojaMaterial[j][3] ? datosHojaMaterial[j][3].toString().trim() : ""; // Columna D para Aluminio
    } else {
      claveEnHoja = datosHojaMaterial[j][2] ? datosHojaMaterial[j][2].toString().trim() : ""; // Columna C para el resto
    }

    // ▼▼▼ CAMBIO IMPORTANTE ▼▼▼
    // Comparamos usando la Clave Real (claveDeBusquedaLower)
    if (claveEnHoja.toLowerCase() === claveDeBusquedaLower) {
      filaEncontrada = j;
      break;
    }
  }

  if (filaEncontrada === -1) {
    // Error actualizado para mostrar la Clave Real
    SpreadsheetApp.getUi().alert("Error de Búsqueda", "No se encontró la clave '" + claveDeBusqueda + "' en la hoja '" + hojaOrigen + "'.", SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("❌ No se encontró la Clave Real '" + claveDeBusqueda + "' en la hoja de origen.");
    return;
  }

  // --- 4. "Traducir" los datos según la hoja de origen ---
  // (Esta sección no necesita cambios, ya que se basa en la filaEncontrada)
  var filaDatos = datosHojaMaterial[filaEncontrada];

  if (hojaOrigen === "Aluminio") {
    unidad = "perfil por ml";
    var descAluminio = filaDatos[4] ? filaDatos[4].toString().trim() : ""; // Columna E
    var claveAluminio = filaDatos[3] ? filaDatos[3].toString().trim() : ""; // Columna D
    descripcion = claveAluminio + " - " + descAluminio;

    // Buscar el costo según el acabado
    var encabezadosAluminio = datosHojaMaterial[0];
    var idxAcabado = encabezadosAluminio.indexOf(acabadoAluminio);
    if (idxAcabado !== -1) {
      costo = filaDatos[idxAcabado];
    } else {
      costo = "Acabado NE";
      Logger.log("⚠️ No se encontró el acabado '" + acabadoAluminio + "' en los encabezados de 'Aluminio'.");
    }

    } else { // Para "Herrajes", "Cristales", "Otros"
      // Extraemos la clave y la descripción
      var claveOtros = filaDatos[2] ? filaDatos[2].toString().trim() : "";    // Columna C (Clave)
      var descOtros = filaDatos[3] ? filaDatos[3].toString().trim() : "";      // Columna D (Descripción)
      
      // Asignamos los valores finales
      unidad = filaDatos[4]; // Columna E
      descripcion = claveOtros + " - " + descOtros; // Combinamos Clave y Descripción
      costo = filaDatos[5]; // Columna F
    }

  Logger.log("📦 Datos extraídos: Unidad: " + unidad + " | Descripción: " + descripcion + " | Costo: " + costo);

  // --- 5. Pegar los resultados en la primera fila vacía a partir de E10 ---
  var rangoDestino = sheetNuevos.getRange("E10:E" + sheetNuevos.getMaxRows());
  var valoresDestino = rangoDestino.getValues();
  var filaVacia = 9; // El índice base es 9 (para que la primera escritura sea en la fila 10)

  for (var k = 0; k < valoresDestino.length; k++) {
    if (valoresDestino[k][0] === "") {
      filaVacia += (k + 1); // Suma 1 (por el índice base 0) + k (filas recorridas)
      break;
    }
  }
  
  // Si no encontró fila vacía en el rango chequeado, escribirá al final.
  if (filaVacia === 9 && valoresDestino[0][0] !== "") {
     filaVacia = sheetNuevos.getLastRow() + 1;
     // Asegurarnos de que no sea menor a 10
     if (filaVacia < 10) filaVacia = 10;
  }

  sheetNuevos.getRange(filaVacia, 5).setValue(unidad);       // Columna E
  sheetNuevos.getRange(filaVacia, 6).setValue(descripcion);  // Columna F
  sheetNuevos.getRange(filaVacia, 7).setValue(costo);        // Columna G

  Logger.log("✅ Material agregado en la fila " + filaVacia + " de la hoja 'Nuevos'.");
  
  // Opcional: Limpiar los campos de búsqueda después de agregar el material
  // sheetNuevos.getRange("B2:B5").clearContent();
}