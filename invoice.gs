/* Invoice functions */

function createInvoiceServer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Import calendar");
  var calendar = sheet.getRange(1, 1).getValue();
  
  // extract month from first date
  var month = sheet.getRange(2, 3).getValue();
  
  if (calendar === "Private lessons") {
    var counter = 0;
    var selection = sheet.getSelection().getActiveRange();
    var firstRow = selection.getRow();
    var lastRow = selection.getLastRow();
    
    for (i = firstRow; i <= lastRow; i++) {
  
      var data = {
        name : sheet.getRange(i, colName).getValue(),
        length : sheet.getRange(i, colLength).getValue(),
        dates : sheet.getRange(i, colDates).getValue(),
        price : sheet.getRange(i, colPrice).getValue(),
        amount : sheet.getRange(i, colAmount).getValue(),
        btw : sheet.getRange(i, colBtw).getValue()
      }
      
      // check if a row with data has been selected, skip rows without data
      if (!data.length || typeof data.length != "number") {
        if (firstRow === lastRow) {
          Browser.msgBox("Please select a row containing a student.");
          return;
        }
        continue;
      }
          
      // add student information to data
      var studentsSheet = ss.getSheetByName("Students");
      var studentsValues = studentsSheet.getDataRange().getValues();
      
      for (j = 0; j < studentsValues.length; j++) {
        if (data.name === studentsValues[j][1].split(" ")[0]) {
          data.fullName = studentsValues[j][1];
          data.recipient = (studentsValues[j][2]) ? studentsValues[j][2] : studentsValues[j][1];
          data.address1 = studentsValues[j][3];
          data.address2 = studentsValues[j][4];
          data.email = studentsValues[j][9];
          break;
        }
      }
      
      var template = ss.getSheetByName("template-invoice-private");
      var invoiceNo = getInvoiceNo();
      var newSheet = template.copyTo(ss).setName("Factuur " + invoiceNo + " - " + data.name);
      newSheet.showSheet();
      
      //search within newSheet and replace    
      searchReplace("{recipient}", data.recipient, newSheet);
      searchReplace("{address1}", data.address1, newSheet);
      searchReplace("{address2}", data.address2, newSheet);
      searchReplace("{name}", data.name, newSheet);
      searchReplace("{invoice-no}", invoiceNo, newSheet);
      searchReplace("{date}", today(), newSheet),
      searchReplace("{description}", ("Pianolessen in " + month), newSheet),
      searchReplace("{btw}", data.btw, newSheet),
      searchReplace("{price}", data.price, newSheet),
      searchReplace("{amount}", data.amount, newSheet)
      
      //add dates to "Datum(s)"
      var datesArr = data.dates.split(", ");
      addDates(datesArr, newSheet);
      
      counter++
    }
    var s = (counter > 1) ? "s" : "";
    Browser.msgBox(counter + " invoice" + s + " created.");
    
  } else if (calendar === "Music for Life") {
  
    var firstRow = 2;
    var lastRow = sheet.getLastRow();
    var rowCount = lastRow - firstRow + 1;
    
    var firstCol = 1;
    var lastCol = sheet.getLastColumn();
    var colCount = lastCol - firstCol + 1;
    
    // copy all data from import sheet
    var allNames = sheet.getRange(firstRow, 1, rowCount).getValues();
    var restOfData = sheet.getRange(firstRow, 2, rowCount, colCount).getValues();
    
    // copy template sheet
    var template = ss.getSheetByName("template-invoice-m4l");
    var invoiceNo = getInvoiceNo();
    var newSheet = template.copyTo(ss).setName("Factuur " + invoiceNo + " - M4L");
    newSheet.activate();
    
    searchReplace("{invoiceNo}", invoiceNo, newSheet);
    searchReplace("{date}", today(), newSheet);
    
    // insert rows for data
    var startCell = getCoords("{tableStart}", newSheet);
    newSheet.insertRowsAfter(startCell.row, rowCount - 1);
    
    // insert data
    var nameRange = newSheet.getRange(startCell.row, startCell.col, rowCount);
    nameRange.setValues(allNames);
    
    var restOfDataRange = newSheet.getRange(startCell.row, startCell.col + 2, rowCount, colCount);
    restOfDataRange.setValues(restOfData);
    
    setNumberFormat("Prijs p/s", '_([$€-809]* #,##0.00_);_([$€-809]* \(#,##0.00\);_([$€-809]* "-"??_);_(@_)', rowCount + 1, newSheet);
    setNumberFormat("Prijs", '_([$€-809]* #,##0.00_);_([$€-809]* \(#,##0.00\);_([$€-809]* "-"??_);_(@_)', rowCount + 1, newSheet);
    setNumberFormat("BTW %", "#0%", rowCount + 1, newSheet);
    setNumberFormat("BTW €", '_([$€-809]* #,##0.00_);_([$€-809]* \(#,##0.00\);_([$€-809]* "-"??_);_(@_)', rowCount + 1, newSheet);
    
    var startCoords = getCoords("Lengte les", newSheet);
    var table = newSheet.getRange(startCoords.row, startCoords.col, rowCount + 1, 7);
    table.setHorizontalAlignment("center");
    
    var subTotalFormula = getRange("{subTotalFormula}", newSheet).setFormulaR1C1("=SUM(R[-" + (rowCount - 1) + "]C[0]:R[-1]C[0])");
    var btwFormula = getRange("{btwFormula}", newSheet).setFormulaR1C1("=SUM(R[-" + rowCount + "]C[0]:R[-2]C[0])");;
    var totalFormula = getRange("{totalFormula}", newSheet).setFormulaR1C1("=R[-1]C[0]+R[-2]C[-2]");

    // format rows that start with a weekday
    var weekdays = ["Zondag", "Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag"];
    weekdays.forEach(function(day) {
      var coords = getCoords(day, newSheet);
      if (coords) {
        var headerRow = newSheet.getRange(coords.row, coords.col, 1, newSheet.getLastColumn() - 1);
        headerRow.setBorder(true, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID).setFontWeight("bold");
      }
    });
    
    // put this towards the end because it takes long but changes nothing visually
    replaceWithFormula("Prijs", "=R[0]C[-1]*R[0]C[-2]", rowCount, newSheet);
    replaceWithFormula("BTW €", "=R[0]C[-1]*R[0]C[-2]", rowCount, newSheet);
    
  }
}


function printSheet() {
    
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheetId = spreadsheet.getId();
  var sheet = SpreadsheetApp.getActiveSheet();
  sheetId = sheet.getSheetId();
  var pdfName = sheet.getName();
  
  if (pdfName.indexOf("Factuur") < 0) {
    Browser.msgBox("This sheet is not an invoice!")
    return;
  }
  
  folderId = 'your folder ID'; // REDACT before upload
  var folder = DriveApp.getFolderById(folderId)
  var url_base = spreadsheet.getUrl().replace(/edit$/, '');
  var url_ext = 'export?exportFormat=pdf&format=pdf'
  + '&gid=' + sheetId
  + '&horizontal_alignment=CENTER'
  + '&fitw=true'      // fit to width, false for actual size - (&source=labnol)
  + '&size=A4'      // paper size
  + '&portrait=true'    // orientation, false for landscape
  + '&fitw=true'        // fit to width, false for actual size
  + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
  + '&gridlines=false'  // hide gridlines
  + '&fzr=false'       // do not repeat row headers (frozen rows) on each page
  + '&top_margin=0.00'              //All four margins must be set!
  + '&bottom_margin=0.00'           //All four margins must be set!
  + '&left_margin=0.00'             //All four margins must be set!
  + '&right_margin=0.00';            //All four margins must be set!

  var options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
    }
  }
  
  var response = UrlFetchApp.fetch(url_base + url_ext, options);
  var blob = response.getBlob().setName(pdfName + '.pdf');
  var invoiceId = folder.createFile(blob).getId();
  var newInvoice = DriveApp.getFileById(invoiceId);
  
  Browser.msgBox(pdfName + ".pdf is saved in Financiën/Facturen.");

}


function deleteAllInvoices() {
  // delete Sheets starting with "Factuur..."
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  var delArr = [];
  var delArrNames = []
  
  for (j = 0; j < allSheets.length; j++) {
    if (allSheets[j].getName().indexOf("Factuur") >= 0) {
      delArr.push(j);
      delArrNames.push(allSheets[j].getName());
    }
  }
  
  if (delArr.length === 0) {
    Browser.msgBox("There are no \"Factuur\" sheets");
    return;
  }
  
  var delNames = delArrNames.join("\n");
  var s = (delArr.length > 1) ? "s" : "";
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Delete sheet' + s + '?',
                          'Do you want to delete the following sheet' + s + ':'
                          +'\n' + delNames,
                          ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    for (k = 0; k < delArr.length; k++) {
      ss.deleteSheet(allSheets[delArr[k]]);
    }
  }
}