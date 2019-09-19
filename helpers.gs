/* Helper functions */

function replaceWithFormula(colHeader, formula, rowCount, sheet) {
  var coords = getCoords(colHeader, sheet);
  var lastRow = rowCount + coords.row;
  
  for (i = coords.row + 1; i <= lastRow; i++) {
    var cell = sheet.getRange(i, coords.col);
    if (!cell.isBlank()) {
      cell.setFormulaR1C1(formula);
    }
  }
}

function setNumberFormat(searchString, type, length, sheet) {
  var coords = getCoords(searchString, sheet);
  var targetRange = sheet.getRange(coords.row, coords.col, length);
  targetRange.setNumberFormat(type);
}


function searchReplace(searchString, newValue, sheet) {
  var coords = getCoords(searchString, sheet);
  sheet.getRange(coords.row, coords.col).setValue(newValue);
}


function getRange(searchString, sheet) {
  var coords = getCoords(searchString, sheet);
  var range = sheet.getRange(coords.row, coords.col);
  return range;
}


function getCoords(searchString, sheet) {
  var searchValues = sheet.getDataRange().getValues();
  
  for (j = 0; j < searchValues.length; j++) {
    for (k = 0; k < searchValues[j].length; k++) {
      if (searchValues[j][k] === searchString) {
        return {row: j+1, col: k+1};
      }
    }
  }
}


function addDates(arr, sheet) {
  var searchValues = sheet.getDataRange().getValues();
  
  // find location of {dates}
  for (j = 0; j < searchValues.length; j++) {
    for (k = 0; k < searchValues[j].length; k++) {
      if (searchValues[j][k] === "{dates}") {
        var range = sheet.getRange((j+1), (k+1));
      }
    }
  }
  
  for (j = 0; j < arr.length; j++) {
    range.offset(j, 0).setValue(arr[j]);
  }
  
  // add border around "Datum(s)"
  var row = range.getRow();
  var col = range.getColumn();
  sheet.getRange(row, col-1, arr.length, 2).setBorder(true, true, true, true, false, false, "#45818e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  
  // change "Datums" to "Datum" if there's only one date
  if (arr.length === 1) {
    range.offset(0, -1).setValue("Datum:");
  }  
}


function getInvoiceNo() {
  // count files in folder "Facturen"
  var query = 'trashed = false and ' +
        "'your folder ID' in parents"; // REDACT before upload
  
  var filesInFolder = Drive.Files.list({q: query});
  
  // count Sheets starting with "Factuur..."
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var counter = 0;
  
  for (j = 0; j < allSheets.length; j++) {
    if (allSheets[j].getName().indexOf("Factuur") >= 0) {
      counter++;
    }
  }
    
  var invoiceNo = filesInFolder.items.length + counter + 1;
  var date = new Date();
  var year = JSON.stringify(date.getFullYear()).slice(-2);  
  var padding;
  
  if (invoiceNo < 10) {
    var padding = "00";
  } else if (invoiceNo < 100) {
    var padding = "0";
  }
  invoiceNo = year + padding + invoiceNo + "";
  
  return invoiceNo;
}


function findIndex(arr, string) {
  for (j = 0; j < arr.length; j++) {
    if (arr[j].student === string) {
      return j;
    }
  }
  //Browser.msgBox('Student called ' + string + ' not found');
  return -1;
}


function today() {
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; //January is 0!
  var yyyy = today.getFullYear();
  
  if(dd<10) {
      dd = '0'+dd
  } 
  
  if(mm<10) {
      mm = '0'+mm
  } 
  
  today = dd + '/' + mm + '/' + yyyy;
  return today;
}