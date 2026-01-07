/**
 * @OnlyCurrentDoc
 * Este script busca uno o varios productos, extrae sus componentes, transforma fórmulas,
 * añade características repetitivas y actualiza el catálogo de origen con la descripción.
 * VERSIÓN 6.2: Maneja dinámicamente hasta 6 medidas y marca los productos procesados con "Registrado".
 */

/**
 * Crea un menú personalizado en la UI de Google Sheets cuando el archivo se abre.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🚀 Herramientas de Productos')
      .addItem('Extraer y Actualizar Producto(s)', 'iniciarProcesoDeExtraccion')
      .addToUi();
}

/**
 * Función principal que busca automáticamente los productos a procesar y los ejecuta en lote.
 * Guarda el progreso para poder continuar después de un tiempo de espera agotado.
 */
function iniciarProcesoDeExtraccion() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const configSheet = obtenerHojaConfig(ss);
  const celdaEstado = configSheet.getRange('A2');
  
  const catalogoSheet = ss.getSheetByName('CATALOGO DE PRODUCTOS');
  if (!catalogoSheet) {
    ui.alert('No se pudo encontrar la hoja "CATALOGO DE PRODUCTOS".');
    return;
  }

  let ultimaFilaProcesada = celdaEstado.getValue();
  if (!isFinite(ultimaFilaProcesada) || ultimaFilaProcesada < 1) {
    ultimaFilaProcesada = 1;
  }

  const filaDeInicio = ultimaFilaProcesada + 1;
  const ultimaFilaCatalogo = catalogoSheet.getLastRow();

  if (filaDeInicio > ultimaFilaCatalogo) {
    ui.alert('¡Proceso completado! Todos los productos del catálogo ya han sido procesados. El marcador de progreso se ha reiniciado.');
    celdaEstado.clearContent();
    return;
  }

  const rangoDeBusqueda = catalogoSheet.getRange(filaDeInicio, 2, ultimaFilaCatalogo - filaDeInicio + 1, 12);
  const valoresDelCatalogo = rangoDeBusqueda.getValues();

  let productosParaProcesar = [];
  for (let i = 0; i < valoresDelCatalogo.length; i++) {
    const fila = valoresDelCatalogo[i];
    const codigoProducto = fila[0];
    const valorFilaGenerador = fila[1];
    const estadoRegistro = fila[11]; // Columna M (13) está en el índice 11 del array

    if (codigoProducto && isFinite(valorFilaGenerador) && estadoRegistro !== "Registrado") {
      productosParaProcesar.push({
        codigo: codigoProducto,
        fila: filaDeInicio + i
      });
    }
  }

  if (productosParaProcesar.length === 0) {
    ui.alert('No se encontraron más productos para procesar desde la fila ' + filaDeInicio + '.');
    celdaEstado.clearContent();
    return;
  }

  const mensajeConfirmacion = `Continuando desde la fila ${filaDeInicio}. Se encontraron ${productosParaProcesar.length} productos restantes. ¿Deseas continuar?`;
  const respuesta = ui.alert('Confirmar Procesamiento', mensajeConfirmacion, ui.ButtonSet.YES_NO);

  if (respuesta !== ui.Button.YES) {
    ui.alert('Operación cancelada por el usuario.');
    return;
  }

  let exitosos = 0;
  let fallidos = [];
  const generadoresSheet = ss.getSheetByName('GENERADORES');
  let productosSheet = ss.getSheetByName('Productos');
  if (!productosSheet) { productosSheet = ss.insertSheet('Productos'); }

  try {
    for (const item of productosParaProcesar) {
      const resultado = procesarProductoUnico(item.codigo, ss, catalogoSheet, generadoresSheet, productosSheet);
      if (resultado) {
        exitosos++;
        celdaEstado.setValue(item.fila); 
      } else {
        fallidos.push(item.codigo);
      }
    }

    let mensajeFinal = `¡Proceso completado con éxito!\n\n- Productos procesados en esta ejecución: ${exitosos}`;
    if (fallidos.length > 0) {
      mensajeFinal += `\n- Productos fallidos: ${fallidos.length} (${fallidos.join(', ')})`;
    }
    mensajeFinal += '\n\nEl marcador de progreso ha sido reiniciado.';
    ui.alert(mensajeFinal);
    celdaEstado.clearContent();

  } catch (e) {
    if (e.message.includes("exceeded maximum execution time")) {
      ui.alert(`Se superó el tiempo de ejecución. Se procesaron ${exitosos} productos.\n\nLa próxima vez que ejecutes el script, continuará automáticamente desde donde se quedó.`);
    } else {
      ui.alert(`Ocurrió un error inesperado: ${e.message}`);
    }
  }
}

/**
 * Procesa un único producto. Esta función contiene la lógica de extracción y actualización.
 */
function procesarProductoUnico(productoCodigo, ss, catalogoSheet, generadoresSheet, productosSheet) {
  const infoBusqueda = buscarFilaDeInicio(productoCodigo, catalogoSheet);
  if (infoBusqueda === null) {
    Logger.log(`No se pudo encontrar el código "${productoCodigo}" en el catálogo.`);
    return false;
  }
  const { filaGeneradores: filaInicio, filaCatalogo: filaDestinoCatalogo } = infoBusqueda;

  const datosRepetitivos = generadoresSheet.getRange(filaInicio, 1, 1, 6).getValues()[0];
  const rangoFilas = determinarRangoDelProducto(filaInicio, generadoresSheet);
  if (!rangoFilas) { return false; }
  
  const extraccion = extraerYTransformarDatos(rangoFilas.inicio, rangoFilas.fin, generadoresSheet);
  const datosMateriales = extraccion.datosProcesados;
  
  if (datosMateriales.length === 0) { return false; }

  const descripcionesAdicionales = extraccion.descripcionesMedidas;
  if (descripcionesAdicionales && descripcionesAdicionales.length > 0) {
    const textoParaColumnaL = descripcionesAdicionales.join('\n');
    catalogoSheet.getRange(filaDestinoCatalogo, 12).setValue(textoParaColumnaL);
  }
  
  const datosFinales = datosMateriales.map(filaMaterial => [productoCodigo, ...datosRepetitivos, ...filaMaterial]);

  const ultimaFila = productosSheet.getLastRow();
  if (ultimaFila === 0) {
    const encabezados = [
      'Código Buscado', 'Característica 1', 'Característica 2', 'Característica 3', 'Característica 4', 'Característica 5', 'Característica 6',
      'Cantidad (Fórmula General)', 'Unidad', 'Descripción', 'P.U.'
    ];
    productosSheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
  }
  productosSheet.getRange(ultimaFila + 1, 1, datosFinales.length, datosFinales[0].length).setValues(datosFinales);
  productosSheet.autoResizeColumns(1, productosSheet.getLastColumn());
  
  const descripcionProducto = buscarDescripcionProducto(generadoresSheet, rangoFilas);
  if (descripcionProducto) {
    catalogoSheet.getRange(filaDestinoCatalogo, 11).setValue(descripcionProducto);
  }
  
  catalogoSheet.getRange(filaDestinoCatalogo, 13).setValue("Registrado");
  
  return true;
}

/**
 * EXTRAE LOS DATOS DE MATERIALES Y TRANSFORMA LAS FÓRMULAS.
 * Usa un bucle para buscar hasta 6 juegos de medidas.
 */
function extraerYTransformarDatos(inicio, fin, sheet) {
  const rangoDatos = sheet.getRange(inicio, 1, fin - inicio + 1, 4);
  const formulas = rangoDatos.getFormulas();
  const valores = rangoDatos.getValues();
  const datosProcesados = [];
  const descripcionesMedidas = [];
  const mapaDeMedidas = {};
  let indiceFilaDeInicioDatos = 0;

  for (let i = 0; i < valores.length; i++) {
    const celdaA = String(valores[i][0]).trim().toLowerCase();
    const celdaB = String(valores[i][1]).trim().toLowerCase();

    if (celdaA === 'largo' && celdaB === 'alto') {
      
      for (let j = 1; j <= 6; j++) {
        const filaDeMedida = i + j;
        
        if (valores[filaDeMedida] && !isNaN(parseFloat(valores[filaDeMedida][0])) && !isNaN(parseFloat(valores[filaDeMedida][1]))) {
          mapaDeMedidas['A' + (inicio + filaDeMedida)] = 'Largo' + j;
          mapaDeMedidas['B' + (inicio + filaDeMedida)] = 'Alto' + j;
          
          const desc = valores[filaDeMedida][2];
          if (desc && String(desc).trim() !== '') {
            descripcionesMedidas.push(String(desc).trim());
          }
        } else {
          break; 
        }
      }
    }

    if (celdaA === 'cantidad' && celdaB === 'unidad') {
      indiceFilaDeInicioDatos = i + 1;
      break;
    }
  }
  
  if (indiceFilaDeInicioDatos === 0) {
      return { datosProcesados: [], descripcionesMedidas: descripcionesMedidas };
  }

  for (let i = indiceFilaDeInicioDatos; i < valores.length; i++) {
    if (valores[i][1] === '') { continue; }

    const formulaOriginal = formulas[i][0];
    let cantidadTransformada = (formulaOriginal) ? transformarFormula(formulaOriginal, mapaDeMedidas) : valores[i][0];

    const [unidad, descripcion, precioUnitario] = [valores[i][1], valores[i][2], valores[i][3]];
    datosProcesados.push([cantidadTransformada, unidad, descripcion, precioUnitario]);
  }

  return { datosProcesados, descripcionesMedidas };
}

/**
 * Busca la descripción principal de un producto dentro de su rango de datos.
 */
function buscarDescripcionProducto(sheet, rango) {
  const rangoBusqueda = sheet.getRange(rango.inicio, 1, rango.fin - rango.inicio + 1, 4).getValues();
  for (let i = rangoBusqueda.length - 1; i >= 0; i--) {
    const [colA, colB, colC, colD] = rangoBusqueda[i];
    if (colA !== '' && colB === '' && colC === '' && colD === '') { return colA; }
  }
  return null;
}

/**
 * Busca el código de un producto en la hoja de Catálogo y devuelve su fila de inicio en Generadores.
 */
function buscarFilaDeInicio(codigo, sheet) {
  const data = sheet.getRange('B2:C' + sheet.getLastRow()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(codigo).trim().toLowerCase()) {
      return { filaGeneradores: data[i][1], filaCatalogo: i + 2 };
    }
  }
  return null;
}

/**
 * Determina el rango de filas que ocupa un producto en la hoja de Generadores.
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

/**
 * Transforma una fórmula de Google Sheets a un texto generalizado.
 */
function transformarFormula(formula, mapaDeMedidas) {
  if (!formula) return '';

  let transformada = formula;
  transformada = transformada.replace(/\$B\$2/g, 'Desp');

  for (const celda in mapaDeMedidas) {
    const regex = new RegExp(celda, 'gi');
    transformada = transformada.replace(regex, mapaDeMedidas[celda]);
  }

  return transformada.startsWith('=') ? "'" + transformada : transformada;
}

/**
 * Obtiene la hoja de configuración, creándola si no existe.
 */
function obtenerHojaConfig(ss) {
  const nombreHoja = 'Configuración';
  let configSheet = ss.getSheetByName(nombreHoja);
  if (!configSheet) {
    configSheet = ss.insertSheet(nombreHoja);
    configSheet.getRange('A1').setValue('Última Fila Procesada (CATALOGO):').setFontWeight('bold');
    configSheet.hideSheet();
  }
  return configSheet;
}