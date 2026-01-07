/**
 * Menú principal unificado para todas las funciones del sistema
 * Combina: Cotización PDF, Generador de Carpeta, y otras herramientas
 */

/**
 * Crea el menú personalizado cuando se abre el spreadsheet
 * IMPORTANTE: Solo puede haber UN onOpen() en todo el proyecto
 * Este archivo reemplaza a los onOpen() individuales
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Menú principal - Generar Documentos
  ui.createMenu('📄 Generar Documentos')
      .addItem('🚀 Iniciar - Seleccionar Tipo de Documento', 'mostrarMenuSelector')
      .addSeparator()
      .addItem('📝 Cotización (Directo)', 'abrirWebappCotizacion')
      .addItem('📄 Nota de Venta (Directo)', 'mostrarFormularioNotaVenta')
      .addSeparator()
      .addItem('🔐 Autorizar Permisos (EJECUTAR PRIMERO)', 'autorizarPermisos')
      .addSeparator()
      .addItem('🧪 Probar Parser', 'probarParser')
      .addItem('🔢 Ver Último Folio', 'mostrarUltimoFolio')
      .addItem('👥 Probar Vendedores', 'probarVendedores')
      .addItem('📋 Probar Lista Proyectos', 'probarListaProyectos')
      .addToUi();

  // Menú de Acciones (Generador Carpeta)
  ui.createMenu('Acciones')
      .addItem('📁 Enviar a Carpeta Cliente', 'transferirGeneradorACliente')
      .addSeparator()
      .addItem('🔐 Autorizar Permisos Drive', 'autorizarPermisosDrive')
      .addItem('🧪 Probar Acceso a Carpeta', 'probarAccesoCarpeta')
      .addToUi();
}

/**
 * Muestra el menú selector de tipo de documento
 */
function mostrarMenuSelector() {
  var html = HtmlService.createHtmlOutputFromFile('Menu_Selector')
    .setWidth(850)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📄 Sistema de Documentos');
}
