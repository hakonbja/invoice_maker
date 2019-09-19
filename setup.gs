
/* Global variables */

var calendarGlobal;

var colName = 1,
    colLength = 2,
    colDates = 3,
    colPrice = 4,
    colAmount = 5,
    colBtw = 7;
    

/* Setup functions */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Invoice Maker')
      .addItem('Show sidebar', 'showSidebar')
      .addSeparator()
      .addItem('Add student', 'addStudentDialog')
      .addToUi();
}


function showSidebar() {
  var html = HtmlService
    .createTemplateFromFile('sidebar.html')
    .evaluate();
    
  html.setTitle(' ');
  SpreadsheetApp.getUi().showSidebar(html);
}


function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}
