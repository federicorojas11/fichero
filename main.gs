// Inicializar formulario
function onOpen() {
  cargarFormulario();
  agregarToolbar();
}

// Función para cargar formulario en Google Sheets
function cargarFormulario() {
  const template = HtmlService.createTemplateFromFile("formulario.html");
  const html = template.evaluate().setTitle('Fichero'); 
  SpreadsheetApp.getUi().showSidebar(html);
}

// Función para agregar 'Acciones->Abrir formulario' desde la barra de herramientas
function agregarToolbar() {
    SpreadsheetApp.getUi()
    .createMenu('Acciones')
    .addItem('Abrir formulario', 'cargarFormulario')
    .addToUi();
}

// Función principal de protección
function protegerEdicion(e) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = spreadsheet.getSheetByName('Configuraciones');
  let estado = configSheet.getRange('C4').getValue();
  
  if(estado){
    const range = e.range;
    const libro = range.getSheet().getParent();
    
    // Limitar protección a columnas B-F (2-6) y filas 3 en adelante
    const col = range.getColumn();
    const row = range.getRow();
    if (col < 2 || col > 6 || row < 3) {
      return; // Salir si está fuera del rango protegido
    }
    
    // Deshacer el cambio mostrando el valor anterior
    range.setValue(e.oldValue || '');
      
    // Mostrar alert
    try {
      SpreadsheetApp.getUi().alert(
        '❌ Edición no permitida',
        'Todas las modificaciones deben realizarse mediante el formulario ➡️\n' +
        'Utilice la opción de Entrada/Salida para editar el fichero.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      cargarFormulario();
    } catch (e) {
      // Fallback para cuando no está disponible la UI modal
    }
    }
}

// función para importar archivos js.html dentro de otro html (modularización)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
