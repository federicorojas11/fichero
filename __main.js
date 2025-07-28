// Inicializar formulario
function onOpen() {
  cargarFormulario();
  agregarToolbar();
}

// función para importar archivos js.html dentro de otro html (modularización)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Función para cargar formulario en Google Sheets
function cargarFormulario() {
  const template = HtmlService.createTemplateFromFile("formulario.html");
  const html = template.evaluate().setTitle('Dirección Registro General de Alojados'); 
  SpreadsheetApp.getUi().showSidebar(html);
}

// Función para agregar 'Acciones->Abrir formulario' desde la barra de herramientas
function agregarToolbar() {
    SpreadsheetApp.getUi()
    .createMenu('Acciones')
    .addItem('Abrir formulario', 'cargarFormulario')
    .addToUi();
}
