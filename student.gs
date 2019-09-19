/* Add a new student functions */

function addStudentDialog() {
  // Display a modeless dialog box with custom HtmlService content.
  var htmlOutput = HtmlService
    .createTemplateFromFile('addstudent.html')
    .evaluate()
    .setWidth(250)
    .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Add student');
}

function addStudent(form) {
  if (!form.name || !form.email || !form.calendar) { }
  else {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var row = sheet.getLastRow() + 1;
    
    sheet.getRange(row, 1).setValue(form.calendar);
    sheet.getRange(row, 2).setValue(form.name);
    sheet.getRange(row, 3).setValue(form.recipient);
    sheet.getRange(row, 4).setValue(form.address1);
    sheet.getRange(row, 5).setValue(form.address2);
    sheet.getRange(row, 6).setValue(form.phone);
    sheet.getRange(row, 7).setValue(form.price);
    sheet.getRange(row, 8).setValue(form.btw);
    sheet.getRange(row, 9).setValue(form.birthday);
    sheet.getRange(row, 10).setValue(form.email);
    sheet.getRange(row, 11).setValue(today());
    
    Browser.msgBox("Student \"" + form.name + "\" added");
  
  }
}
