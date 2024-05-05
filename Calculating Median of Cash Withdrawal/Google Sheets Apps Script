function calculatingMedian() {
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");

  // Activate the source sheet
  SpreadsheetApp.setActiveSheet(mySheet);

  // Clear the result column
  mySheet.getRange("D:D").clearContent();
  mySheet.getRange("D1").setValue("Alert Status");

  // Starting row
  startRow = 2;

  //Default value of counter
  var counter = 0;

  // Finding the last non-empty row in Column A (custom function)
  var untilRow = findLastRow("Data", "A");

  for (var i = startRow; i <= untilRow; i++) {
    //Logger.log(mySheet.getRange(i,2).getValue());

    if (mySheet.getRange(i,1).getValue() != mySheet.getRange(i-1, 1).getValue()) {
      // Reset the counter then customer changes
      counter = 1;
    }
    else {
      counter++;
    }

    // Comparing the median value to double of the transaction and writing the result into "Alert Status" column.
    if (counter < 3) {
      mySheet.getRange(i,4).setValue("DO NOT SEND ALERT"); // We are calculating last 3 transaction of customer
    }
    else {
        
        var medianValues = mySheet.getRange(i-2,2,3,1).getValues();
        var medianResult = calculateMedian(medianValues); 
        
        var doubleOfTransaction = mySheet.getRange(i,2).getValue() * 2;
        
        if (medianResult >= doubleOfTransaction) {
          mySheet.getRange(i,4).setValue("SEND ALERT");
        }
        else {
          mySheet.getRange(i,4).setValue("DO NOT SEND ALERT");
        }
        //Logger.log("medianValues >> " + medianValues);
        //Logger.log("medianResult >> " + medianResult);
    }
  }
}




// For this function to work (above) there are 2 other functions must also be defined.
function calculateMedian(values) {
  // Since there is no built-in function to calculate median in Apps Script, I asked chatGPT to write the code. Thank you chatGPT :)
  // Sort the values in ascending order
  values.sort(function(a, b) {
    return a - b;
  });
  
  var length = values.length;
  
  if (length % 2 === 0) {
    // If the number of values is even, calculate the average of the two middle values
    var mid1 = values[length / 2 - 1];
    var mid2 = values[length / 2];
    return (mid1 + mid2) / 2;
  } else {
    // If the number of values is odd, return the middle value
    return values[Math.floor(length / 2)];
  }
}




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


