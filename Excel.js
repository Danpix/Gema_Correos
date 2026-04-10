//Codigo para appscript donde guardo datos de el codigo principal en sheets

function guardarEnSheetPruebaFecha() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PruebaFecha");
    var fechaActual = new Date();
    sheet.appendRow([fechaActual]);
  }
