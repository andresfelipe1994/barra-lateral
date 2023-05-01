

var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Technology')
  .addItem('Mostra Barra Lateral','mostrarBarraLateral')
  .addToUi();
}

function mostrarBarraLateral()
{
  var barra = HtmlService.createHtmlOutputFromFile('BarraLateral').setTitle('Barra Lateral');
  SpreadsheetApp.getUi().showSidebar(barra);
}

function aplicarEstilo1(){

  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas     = hojaActual.getActiveRange();

  celdas.setBackground("blue")
        .setFontColor("white")
        .setHorizontalAlignment("center")
        .setValue("Estilo 1")

}

function aplicarEstilo2(){

  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas     = hojaActual.getActiveRange();

  celdas.setBackground("green")
        .setFontColor("white")
        .setFontWeight("bold")
        .setValue("Estilo 2")

}

function aplicarEstilo(){

  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas     = hojaActual.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty('color'))
        .setBackground(estilos_sheet.getProperty('ColorFondo'))
        .setFontSize(estilos_sheet.getProperty('size'));
}

function guardarEstilos(){

  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  estilos_sheet.setProperty('color', celda.getFontColor())
               .setProperty('colorFondo', celda.getBackground())
               .setProperty('size', celda.getFontSize()+'');

  return { colorFondo: estilos_sheet.getProperty('colorFondo'),
           colorLetra: estilos_sheet.getProperty('color')};

}



function borrarTodo (){
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}

function borrarEstilos (){
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly:true});
}
