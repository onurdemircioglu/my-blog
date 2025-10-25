function retrieveWordsFromWebPage() {
  // URL of the web page
  var url = "https://en.wikipedia.org/wiki/Main_Page";
  
  // Fetch the content of the web page
  var response;
  try {
    response = UrlFetchApp.fetch(url);
  } catch (error) {
    Logger.log("Error fetching webpage: " + error);
    return;
  }
  
  // Check if the request was successful (HTTP status code 200)
  if (response.getResponseCode() !== 200) {
    Logger.log("Failed to fetch webpage. HTTP status code: " + response.getResponseCode());
    return;
  }
  
  // Extract the words from the HTML content
  var content = response.getContentText();
  var words = extractWordsFromHTML(content);
  
  // Eliminate duplicate words
  var uniqueWords = eliminateDuplicates(words);

  // Remove words containing numbers
  var wordsWithoutNumbers = removeWordsWithNumbers(uniqueWords);

  // Remove all uppercase words
  var wordsWithoutUppercase = removeUppercaseWords(wordsWithoutNumbers);

 // Remove two-character words
  var wordsWithoutTwoCharacters = removeTwoCharacterWords(wordsWithoutUppercase);

  // Remove words containing underscores, commas, or dots
  var filteredWords = removeSpecialCharacters(wordsWithoutTwoCharacters);

  // Remove specific words
  var finalWords = eliminateWords(filteredWords)
  
  // Write the unique words onto the "ListofWords" sheet
  writeWordsToSheet(finalWords);
}




function extractWordsFromHTML(html) {
  // Ensure that html is not undefined
  if (!html) {
    return [];
  }
  
  // Remove HTML tags
  // var text = html.replace(/<[^>]*>/g, '');
  var text = html.replace(/<[^>]*>/g, '').replace(/&[a-z]+;/gi, ''); // 2025-08-31 ChatGPT >> extractWordsFromHTML using regex is okay, but it might still leave behind junk like encoded HTML entities (&nbsp;, &amp;). You could clean them:

  
  // Split the text into words
  var words = text.match(/\b\w+\b/g);
  
  return words || [];
}


// Eliminating duplicates from the HTML parse
function eliminateDuplicates(words) {
  // Create a Set to store unique words
  var uniqueWordsSet = new Set(words);
  
  // Convert the Set back to an array
  var uniqueWords = Array.from(uniqueWordsSet);
  
  return uniqueWords;
}


// Eliminating the words with numbers
function removeWordsWithNumbers(words) {
  // Filter out words containing numbers
  var filteredWords = words.filter(function(word) {
    return !/\d/.test(word); // Test if word contains a digit
  });
  
  return filteredWords;
}


// Eliminating upper case words
function removeUppercaseWords(words) {
  // Filter out words with uppercase letters other than the first character
  var filteredWords = words.filter(function(word) {
    // Check if the word has uppercase letters other than the first character
    for (var i = 0; i < word.length; i++) {
      if (word[i] === word[i].toUpperCase()) {
        return false; // Word contains uppercase letter(s) other than the first character
      }
    }
    return true; // Word is valid (no uppercase letter(s) other than the first character)
  });
  
  return filteredWords;
}


// Eliminating length of the word < 3
function removeTwoCharacterWords(words) {
  // Filter out two-character words
  var filteredWords = words.filter(function(word) {
    return word.length > 2; // Test if word has more than two characters
  });
  
  return filteredWords;
}


// Eliminating special characters
function removeSpecialCharacters(words) {
  // Filter out words containing underscores, commas, or dots
  var filteredWords = words.filter(function(word) {
    return !(/[_,.]/.test(word)); // Test if word contains underscores, commas, or dots
  });
  
  return filteredWords;
}


// Read excluded words from EXCLUDED_WORDS!A column
function getExcludedWords() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EXCLUDED_WORDS");
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  // Read all values from column A (ignore empty cells)
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  return values
    .map(function(row) { return row[0]; })
    .filter(function(word) { return word && word.toString().trim() !== ""; });
}


// Filter words against the sheet values
function eliminateWords(words) {
  var excluded = getExcludedWords();
  var excludedSet = new Set(excluded.map(function(w){ return w.toString().toLowerCase(); })); // case-insensitive
  return words.filter(function(word) {
    return !excludedSet.has(word.toString().toLowerCase());
  });
}


// Write the word to the result sheet.
function writeWordsToSheet(words) {
  // Get the active spreadsheet and the "ListofWords" sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("ListofWords");
  
  var lastRow = findLastRow("ListofWords", "A");
  
  // Write the words onto the sheet
  var numRows = words.length;
  var range = sheet.getRange(lastRow + 1, 1, numRows, 1);
  
  range.setValues(words.map(function(word) { return [word]; }));
}


// Remove Duplications on Column A
function removeColumnDuplicates() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("ListofWords");
  //var sheet = SpreadsheetApp.getActive();
  sheet.getRange('A:A').activate();
  sheet.getActiveRange().offset(1, 0, sheet.getActiveRange().getNumRows() - 1).activate();
  sheet.getActiveRange().removeDuplicates().activate();
  sheet.getRange('C2').activate();
};



////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

// writeWordsToSheet function uses findLastRow function:
function findLastRow(sheetName, columnName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    //Logger.log("Sheet '" + sheetName + "' not found.");
    return 0; // Return zero if the sheet is not found
  }
  
  var lastRow = sheet.getLastRow();
  var columnValues = sheet.getRange(columnName + "1:" + columnName + lastRow).getValues();
  
  for (var i = lastRow; i > 0; i--) {
    if (columnValues[i - 1][0] !== "") {
      return i;
    }
  }
  // Return zero if the column is empty
  return 0;
  //Logger.log("Column '" + columnName + "' is empty in sheet '" + sheetName + "'.");
}
