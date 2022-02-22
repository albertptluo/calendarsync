function NCS_sync2Weeks() {
  function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2 Wk At A Glance");
  var myCalendar = CalendarApp.getCalendarsByName("NCS Team")[0];
  if( sheet == null || myCalendar == null) {return;}

  //SpreadsheetApp.getActiveSpreadsheet().getRange('E270').setValue('a');
  var columnRange = "A" + ":" + "J";
  var allCellsInWUD = sheet.getRange(columnRange).getValues();
  var allCellValuesInWUD = sheet.getRange(columnRange).getDisplayValues();
  
  var numRows = sheet.getLastRow();
  var index = 9;
  var eventName = null;
  var eventStart = null;
  var eventEnd = null;
  var eventLoc = null;
  var events = null;
  var toCreate = [];
  var toCreateNames = [];
  var toDelete = {};
  var descriptionText; 
  
  var allDates = [];
  
  var currentDate = new Date(Utilities.formatDate(new Date(), "EST", "MM/dd/yyy"));
  allDates[currentDate.getTime()] = currentDate;
  
  // get date for the current block of rows
  var dateOnCell = allCellsInWUD[index][0];
  
  // if it's before today, just skip to the next index
  var notToday = true;
  var today = new Date(Utilities.formatDate(new Date(), "EST", "MM/dd/yyy"));
  while (notToday) {
    dateOnCell = allCellsInWUD[index][0];
    if (dateOnCell != '' && getDate(dateOnCell).getTime() >= today.getTime()) {
      notToday = false;
      break;
    }
    index++;
  }
  
  
  while(index < numRows)
  {
    dateOnCell = allCellsInWUD[index][0];
    if (dateOnCell != "" && dateOnCell >= currentDate) {
      currentDate = dateOnCell;
      allDates[currentDate.getTime()] = currentDate;
    }

    // // check if we want to add to calendar
    // var addToCalendar = allCellValuesInWUD[index][9] == 'TRUE';
    // if (!addToCalendar) {
    //   index++;
    //   continue;
    // }

    // check ministry category
    var ministry = allCellValuesInWUD[index][1];
    if (ministry != "NC State") {
      index++;
      continue;
    }

    // Get event start/end times
    var eventDate = new Date(currentDate);    
    eventStart = getDateFromTimeString(allCellValuesInWUD[index][2]);
    if (eventStart == null) {
      index++;
      continue;
    }
    eventStart.setMonth(eventDate.getMonth());
    eventStart.setDate(eventDate.getDate());
    
    eventEnd = getDateFromTimeString(allCellValuesInWUD[index][3]);
    if (eventEnd == null) {
      eventEnd = new Date(eventStart);
      eventEnd.setHours(eventStart.getHours() + 1);
    }
    eventEnd.setMonth(eventDate.getMonth());
    eventEnd.setDate(eventDate.getDate());
    
    // Process the event name for the row
    eventName = allCellsInWUD[index][4];      
    if(eventName == "") {
      index++;
      continue;
    }
    
    eventLoc = allCellsInWUD[index][5];
    
    // Process extra columns
    descriptionText = "In Charge: " + allCellsInWUD[index][6] 
                    + "\n"
                    + "Who Else: " + allCellsInWUD[index][7]
                    + "\n"
                    + "Notes: " + allCellsInWUD[index][8] 
                    + "\n"
                    + "Childcare: " + allCellsInWUD[index][9]
                    + "\n\n\n"
                    + "Updated via script at: " + Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy HH:mm:ss")
    
    toCreate.push({name: eventName, 
                   start: eventStart, 
                   end: eventEnd, 
                   options: {location: eventLoc, 
                             description: descriptionText}
                  });
    toCreateNames.push(eventName);
    index++;
  }
 
  for (var d in allDates) {
    
    events = myCalendar.getEventsForDay(allDates[d]);
    for(var e in events){
      var name = events[e].getTitle();
      // delete all calendar events that were created by the script
      // except for all day events
      if (events[e].getDescription().indexOf("Updated via script at:") > 0 ) {
        toDelete[events[e].getId()] = events[e];
      }
    }
    
  }
  for(var key in toDelete){
    var title = toDelete[key].getTitle();
    toDelete[key].deleteEvent();
    Utilities.sleep(5);
  }
  for (var e in toCreate) {
    var start = toCreate[e]['start'].getHours();
    var end = toCreate[e]['end'].getHours();
    myCalendar.createEvent(toCreate[e]['name'], toCreate[e]['start'], toCreate[e]['end'], toCreate[e]['options']);
    Utilities.sleep(5);
  }
}
    
// for debugging purposes. Delete all events in the last 8 days and 20 days from now
function clearAllEvents() {
  var myCalendar = CalendarApp.getCalendarsByName("INSERT CALENDAR NAME")[0];
  var now = new Date();
  events = myCalendar.getEvents(new Date(now.getTime() - 8*24*60*60*1000), new Date(now.getTime() + 20*24*60*60*1000));
  for(var e in events){
    events[e].deleteEvent();
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

function splitTime(time) {
  var regExp = new RegExp('^(\\d{1,2}):?(\\d{2})?\\s?([a,p]?[m]?)?-?(\\d{1,2}):?(\\d{2})?\\s?([a,p]?[m]?)?', "gi");
  return regExp.exec(time);
}

function getDateFromTimeString(timeString) {
  var timeArray = splitTime(timeString);
  if (timeArray == null) {
    return null;
  }
  var startHour = parseInt(timeArray[1]);
  var startMinute= parseInt(timeArray[4]);
  if (timeArray[6] == "PM" && startHour < 12) startHour += 12;
  var time = new Date();
  time.setHours(startHour);
  time.setMinutes(startMinute);
  return time;
}
