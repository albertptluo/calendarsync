/*
Author: Albert Luo
For copying Birthdays from google spreadsheet to google calendar for NCS team
*/

function copyBirthdays() { 
  function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Birthdays"); // Put the tab name here
  var myCalendar = CalendarApp.getCalendarsByName("NCS Team")[0]; // Put the Calendar name here
  if( sheet == null || myCalendar == null) {return;}

  var columnRange = "A" + ":" + "D";
  var allCellsInWUD = sheet.getRange(columnRange).getValues();
  var allCellValuesInWUD = sheet.getRange(columnRange).getDisplayValues();
  
  var index = 1;
  var toCreate = [];
  
  
  var currentDate = new Date(Utilities.formatDate(new Date(), "EST", "MM/dd/yyy"));

  // assume less than 50 rows
  while(index < 50)
  {
    // check if already added to calendar
    var alreadyAddedToCalendar = allCellValuesInWUD[index][2] == 'TRUE';
    if (alreadyAddedToCalendar) {
      index++;
      continue;
    }
    
    
    // Process the student name name for the row
    var name = allCellsInWUD[index][0];      
    if(name == "") {
      index++;
      continue;
    }    
    
    // Get Birthday
    var birthday = allCellsInWUD[index][1];
    if (birthday == "" || !isValidDate(birthday)) {
      index++;
      continue;
    }
    
    var description = "";

    toCreate.push({name: "Birthday - " + name, 
                   birthday: birthday, 
                   options: {description: description}
                  });
    
    var addedColumn = "C" + (index+ 1);
    sheet.getRange(addedColumn).setValue('TRUE');
    
    index++;
  }

  
  
  for (var e in toCreate) {
    var recurrence = CalendarApp.newRecurrence().addYearlyRule();
    myCalendar.createAllDayEventSeries(toCreate[e]['name'], getDate(toCreate[e]['birthday']), recurrence);
  }
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function getDate(d) {
  return new Date(Utilities.formatDate(new Date(d), "EST", "MM/dd/yyyy"));
}
