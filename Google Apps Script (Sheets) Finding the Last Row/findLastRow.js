function findLastRow(sheetName, columnName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) { // If sheet doesn't exist
    //Logger.log("Sheet '" + sheetName + "' not found.");
    return 0; // Return zero if the sheet is not found
  }

  if (!columnName) {
    //Logger.log("Column '" + columnName + "' not found.");
    return 0; // Return zero if the column name is not found
  }
  
  var lastRow = sheet.getLastRow(); // This is the original method (https://developers.google.com/apps-script/reference/spreadsheet/sheet#getLastRow())
  var columnValues = sheet.getRange(columnName + "1:" + columnName + lastRow).getValues(); // To retrieve the values on the specified column
  
  for (var i = lastRow; i > 0; i--) {
    if (columnValues[i - 1][0] !== "") {
      return i;
    }
  }
  // Return zero if the column is empty
  return 0;
  //Logger.log("Column '" + columnName + "' is empty in sheet '" + sheetName + "'.");
}




// To call the function and assign the result into a variable:
var lastRow = findLastRow("Sample Sheet Name", "A"); 
