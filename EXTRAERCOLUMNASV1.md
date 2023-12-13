function importDataFromGoogleSheetsInFolder() {
  try {
    // ID de la carpeta desde una celda específica
    var folderIDCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGA").getRange("N5");
    var folderID = folderIDCell.getValue();

    // Nombre de la hoja de cálculo
    var sheetNameCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGA").getRange("L5");
    var sheetName = sheetNameCell.getValue();

    // Nombre de la hoja destino
    var destinationSheetNameCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGA").getRange("M5");
    var destinationSheetName = destinationSheetNameCell.getValue();
    
    // Rango destino
    var destinationRangeCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGA").getRange("P5");
    var destinationRange = destinationRangeCell.getValue();

    // Lista de columnas
    var columnListRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGA").getRange("O5:O"); // Ajusta el rango según tus necesidades
    var columnListValues = columnListRange.getValues().flat();

    // Filtra columnas vacías y convierte los valores a números
    var columnsToExtract = columnListValues.filter(function (value) {
      return value !== "" && !isNaN(value);
    });

    // Busca la hoja de cálculo en la carpeta
    var folder = DriveApp.getFolderById(folderID);
    var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

    // Verifica si se encontró la hoja de cálculo
    if (files.hasNext()) {
      var file = files.next();

      // Abre la hoja de cálculo
      var spreadsheet = SpreadsheetApp.open(file);

      // Obtiene la hoja específica
      var sheet = spreadsheet.getSheetByName(sheetName);

      // Obtiene los datos de la hoja
      var sourceData = sheet.getDataRange().getValues();

      // Filtra solo las columnas seleccionadas
      var data = sourceData.map(function (row) {
        return columnsToExtract.map(function (col) {
          return row[col - 1]; // Resta 1 porque los números de columna comienzan en 1
        });
      });

      // Obtiene la hoja destino dinámica
      var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(destinationSheetName);
      
      // Define la celda de inicio para pegar los datos
      var startRow = parseInt(destinationRange.match(/\d+/)[0], 10);
      var startColumn = destinationRange.match(/[A-Z]+/)[0];
      var columnNumber = startColumn.charCodeAt(0) - 65 + 1;
      
      // Pega los datos en el rango destino
      destinationSheet.getRange(startRow, columnNumber, data.length, data[0].length).setValues(data);

      Logger.log("Datos importados con éxito en el rango '" + destinationRange + "' de la hoja '" + destinationSheetName + "'.");
    } else {
      Logger.log("No se encontraron hojas de cálculo en la carpeta.");
    }
  } catch (error) {
    Logger.log("Error: " + error.toString());
  }
}
