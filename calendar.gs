/* Calendar and Events functions */


function importCalendarServer(dates, calendar) {

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  
  if (activeSheet !== "Import calendar") {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Wrong sheet', 'Do you want to go to the "Import calendar" sheet and continue?', ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.NO) {
      return;
    } else if (response == ui.Button.YES) {
      clearEvents();
      importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Import calendar");
      SpreadsheetApp.setActiveSheet(importSheet);
    }
  } else {
    clearEvents();
  }
  
  calendarGlobal = calendar;
  
  if (calendar === "Music for Life") {
   var calendarId = 'your calendar ID'; // REDACT before upload
  } else if (calendar === "Private lessons") {
    var calendarId = 'your calendar ID'; // REDACT before upload
  }
  
  var optionalArgs = {
    timeMin: dates[0],
    timeMax: dates[1],
    showDeleted: false,
    singleEvents: true,
    //maxResults: 2,
    orderBy: 'startTime'
  };

  var response = Calendar.Events.list(calendarId, optionalArgs);
  var events = response.items;
  
  if (events.length > 0) {
    var eventsArr = [];
    
    for (i = 0; i < events.length; i++) {
      var event = events[i];
      var student = events[i].summary;
      var startTime = new Date(event.start.dateTime).getHours() + "." + new Date(event.start.dateTime).getMinutes();
      var length = new Date(event.end.dateTime) - new Date(event.start.dateTime);
      var date = new Date(event.start.dateTime);
      var dayOfMonth = date.getDate();
      var month = date.getMonth() + 1;
      var year = date.getYear();
      var fullDate = dayOfMonth + "." + month + "." + year;
      var weekday = date.getDay();
      
      if (calendar === "Music for Life") {
        var dateFormatted = [dayOfMonth];
      } else {
        var dateFormatted = [fullDate];
      }
      
      var eventObj = {
        student: student,
        length: length/60000,
        quantity: 1,
        date: [dateFormatted],
        weekday: weekday,
        startTime: startTime,
        monthNr: month
      }

      var studentIndex = findIndex(eventsArr, student);

      if (studentIndex >= 0) {
        eventsArr[studentIndex].quantity++;
        eventsArr[studentIndex].date.push(dateFormatted);
      } else if (studentIndex === -1) {      
        eventsArr.push(eventObj);
      }
    }
  }
  
  if (!eventsArr) {
    Browser.msgBox("No events found.");
    return;
  }
  
  //sort array of events
  eventsArr.sort(function(a, b) {
    return a.startTime - b.startTime;
  }).sort(function(a, b) {
    return a.weekday - b.weekday;
  });

  //show events on sheet
  renderEvents(eventsArr);
}

function renderEvents(arr) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Import calendar");
  
  var pricesSheet = ss.getSheetByName("Prices");
  var pricesLastCol = pricesSheet.getLastColumn();
  var pricesLastRow = pricesSheet.getLastRow();
  var pricesValues = pricesSheet.getRange(2, 1, pricesLastRow, pricesLastCol).getValues();
  
  var studentsSheet = ss.getSheetByName("Students");
  var studentsValues = studentsSheet.getDataRange().getValues();
  
  var weekdays = ["Zondag", "Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag"];
  var startRow = 2;
  var startCol = 1;
  
  var months = ['Januari', 'Februari', 'Maart', 'April', 'Mei', 'Juni', 'Juli', 'Augustus', 'September', 'Oktober', 'November', 'December'];
  var month = months[arr[0].monthNr-1];
  
  // find price of classes and add to each object
  if (calendarGlobal === "Music for Life") {
    arr.forEach(function(obj, index) {
      for (i = 0; i < pricesValues.length; i++) {
        if (obj.student.toLowerCase().indexOf("proef") >= 0) {
          if (calendarGlobal === pricesValues[i][0] && obj.length === pricesValues[i][3] && pricesValues[i][2] === 0 && pricesValues[i][1] === "proefles") { //assumes every lesson is a single lesson and without BTW
            obj.price = pricesValues[i][4];
            break;
          }
        } else if (calendarGlobal === pricesValues[i][0] && obj.length === pricesValues[i][3] && pricesValues[i][2] === 0) { //assumes every lesson is a single lesson and without BTW
          obj.price = pricesValues[i][4];
          break;
        }
      }     
    });
    
    
  } else if (calendarGlobal === "Private lessons") {
    arr.forEach(function(obj, i) {
      //find row in pricesValues where obj.length === column(D).value && column(a).value === calendarGlobal
      for (i = 0; i < pricesValues.length; i++) {
        if (obj.student.toLowerCase().indexOf("proef") >= 0) {
          obj.price = 0;
          break;
        }
        if (calendarGlobal === pricesValues[i][0] && obj.length === pricesValues[i][3]) {
          obj.price = pricesValues[i][4];
          break;
        }
      }
      
    });
  }
  
  //if obj.student appears on students list change obj.price
    arr.forEach(function(obj, index) {
      for (i = 0; i < studentsValues.length; i++) {
        if (studentsValues[i][1].split(" ")[0] === obj.student.split(" ")[0]) {
          obj.price = studentsValues[i][6];
          obj.btw = studentsValues[i][7];
        }
      }
    });
  
  // paste values into correct fields
  sheet.getRange(1, 1).setValue(calendarGlobal);

  
  arr.forEach(function(obj, index) {
    if (index === 0 || arr[index].weekday !== arr[index-1].weekday) {
      sheet.getRange(startRow, startCol).setValue(weekdays[obj.weekday]).setFontWeight("bold");
      sheet.getRange(startRow, startCol + 2).setValue(month).setFontWeight("bold").setHorizontalAlignment("center");
      sheet.getRange(startRow, startCol, 1, 8).setBackground("#d9d9d9");
      startRow++;
    }
    sheet.getRange(startRow, startCol).setValue(obj.student);
    sheet.getRange(startRow, startCol+1).setValue(obj.length);
    var dates = obj.date.join(", ");
    sheet.getRange(startRow, startCol+2).setValue(dates).setHorizontalAlignment("right");
    
    if (obj.price) {
      sheet.getRange(startRow, startCol+3).setValue(obj.price).setNumberFormat('_([$€-809]* #,##0.00_);_([$€-809]* \(#,##0.00\);_([$€-809]* "-"??_);_(@_)');
    }
    
    sheet.getRange(startRow, startCol+4).setValue(obj.quantity);
    
    sheet.getRange(startRow, startCol+5).setFormula("=E" + startRow + "*D" + startRow).setNumberFormat('_([$€-809]* #,##0.00_);_([$€-809]* \(#,##0.00\);_([$€-809]* "-"??_);_(@_)');
    
    var btw = obj.btw ? obj.btw : 0
    sheet.getRange(startRow, startCol+6).setValue(btw).setNumberFormat("0%");
    
    sheet.getRange(startRow, startCol+7).setFormula("=G" + startRow + "*F" + startRow).setNumberFormat('_([$€-809]* #,##0.00_);_([$€-809]* \(#,##0.00\);_([$€-809]* "-"??_);_(@_)');
    
    
    startRow++;
  });
}


function clearEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Import calendar");
  var startRow = 2;
  var startCol = 1;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  sheet.getRange(1, 1).setValue("[calendar]");
  
  if (startRow < lastRow) {
    var clearRange = sheet.getRange(startRow, startCol, lastRow-startRow+1, lastCol-startCol+1);
    clearRange.setValue("").setFontWeight(null).setBackground(null);
  }
}