/**
 * MINI-CÓDIGO
 * Busca descripciones adicionales para productos marcados con "S" en el catálogo
 * y las coloca en la columna L.
 */
function actualizarDescripcionesAdicionales() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalogoSheet = ss.getSheetByName('CATALOGO DE PRODUCTOS');
  const generadoresSheet = ss.getSheetByName('GENERADORES');

  if (!catalogoSheet || !generadoresSheet) {
    ui.alert('No se encontraron las hojas "CATALOGO DE PRODUCTOS" o "GENERADORES".');
    return;
  }

  // Obtenemos los datos a revisar: Columna A (marcador "S") y C (fila en Generadores)
  const rangoParaRevisar = catalogoSheet.getRange('A2:C189').getValues();
  let productosActualizados = 0;

  // 1. Recorremos el rango especificado en el catálogo
  for (let i = 0; i < rangoParaRevisar.length; i++) {
    const marcador = rangoParaRevisar[i][0]; // Columna A
    const filaInicioGeneradores = rangoParaRevisar[i][2]; // Columna C

    // 2. Si la columna A tiene una "S" y hay un número de fila válido...
    if (marcador.toString().toUpperCase() === 'S' && isFinite(filaInicioGeneradores) && filaInicioGeneradores > 0) {
      
      // 3. Vamos a buscar solo las descripciones a la hoja GENERADORES
      const descripciones = buscarSoloDescripcionesEnGenerador(filaInicioGeneradores, generadoresSheet);

      // 4. Si encontramos alguna descripción...
      if (descripciones.length > 0) {
        const textoParaColumnaL = descripciones.join('\n');
        const filaActualCatalogo = i + 2; // i es el índice del array (empieza en 0), +2 para la fila real de la hoja
        
        // 5. La escribimos en la columna L (12)
        catalogoSheet.getRange(filaActualCatalogo, 12).setValue(textoParaColumnaL);
        productosActualizados++;
      }
    }
  }

  ui.alert(`¡Proceso completado! Se actualizaron ${productosActualizados} productos.`);
}

/**
 * Función auxiliar que busca el rango de un producto y extrae
 * únicamente las descripciones de la columna C junto a las medidas.
 * @param {number} filaInicio La fila donde empieza el producto en "GENERADORES".
 * @param {Sheet} sheet La hoja "GENERADORES".
 * @returns {string[]} Un array con las descripciones encontradas.
 */
function buscarSoloDescripcionesEnGenerador(filaInicio, sheet) {
  // Usamos la misma función de antes para saber dónde termina el producto
  const rangoProducto = determinarRangoDelProducto(filaInicio, sheet);
  if (!rangoProducto) return [];

  const valores = sheet.getRange(rangoProducto.inicio, 1, rangoProducto.fin - rangoProducto.inicio + 1, 3).getValues();
  const descripcionesEncontradas = [];

  for (let i = 0; i < valores.length; i++) {
    const celdaA = String(valores[i][0]).trim().toLowerCase();
    const celdaB = String(valores[i][1]).trim().toLowerCase();

    // Si encontramos el encabezado de las medidas...
    if (celdaA === 'largo' && celdaB === 'alto') {
      // Revisamos hasta 3 filas por debajo por si hay múltiples medidas
      for (let j = 1; j <= 3; j++) {
        const filaDeMedida = i + j;
        if (valores[filaDeMedida] && !isNaN(parseFloat(valores[filaDeMedida][0]))) {
          const descripcion = valores[filaDeMedida][2]; // Columna C
          if (descripcion && String(descripcion).trim() !== '') {
            descripcionesEncontradas.push(String(descripcion).trim());
          }
        }
      }
      break; // Encontramos las medidas, ya no necesitamos seguir buscando
    }
  }
  return descripcionesEncontradas;
}

/* NOTA: Esta función de abajo (`determinarRangoDelProducto`) es una copia
  de la que ya tienes en tu script principal. Se necesita aquí para que
  el "mini-código" pueda funcionar de forma independiente.
*/
function determinarRangoDelProducto(filaInicio, sheet) {
  const columnaMarcadores = sheet.getRange(filaInicio + 1, 8, sheet.getLastRow() - filaInicio, 1).getValues();
  let filaFin = sheet.getLastRow();
  for (let i = 0; i < columnaMarcadores.length; i++) {
    if (columnaMarcadores[i][0] !== '') {
      filaFin = (filaInicio + 1) + i - 1;
      break;
    }
  }
  return { inicio: filaInicio, fin: filaFin };
}
