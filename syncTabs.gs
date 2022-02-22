function Sync_Tabs() {
  
  var currentDate = new Date(Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy"));
  var tabName = formatDate(currentDate);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  while (sheet != null) {
    syncTabToCalendars(sheet, currentDate);
    currentDate.setDate(currentDate.getDate()+1);
    tabName = formatDate(currentDate);
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  }
}

function syncTabToCalendars(sheet, currentDate) {
  var ncs_calendar = CalendarApp.getCalendarsByName("NCS Team")[0];
  var churchwide_calendar = CalendarApp.getCalendarsByName("nc_churchwide")[0]; 
  var unc_calendar = CalendarApp.getCalendarsByName("unc_a2f_team")[0];
  if( sheet == null) {return;}
  if (ncs_calendar == null || unc_calendar == null || churchwide_calendar == null) { return; }

  //SpreadsheetApp.getActiveSpreadsheet().getRange('A2').setValue('a');
  var columnRange = "A" + ":" + "J";
  var allCellsInTab = sheet.getRange(columnRange).getValues();
  var allCellValuesInTab = sheet.getRange(columnRange).getDisplayValues();
  
  var numRows = sheet.getLastRow();
  var index = 2;
  var eventName = null;
  var eventStart = null;
  var eventEnd = null;
  var eventLoc = null;
  var events = null;
  var ncs_events = [];
  var churchwide_events = [];
  var unc_events = [];
  var toDelete = {};
  var descriptionText; 
  
  
  
  while(index < numRows)
  {
    // Get event start/end times
    var eventDate = new Date(currentDate);    
    eventStart = getDateFromTimeString(allCellValuesInTab[index][1]);
    if (eventStart == null) {
      index++;
      continue;
    }
    eventStart.setMonth(eventDate.getMonth());
    eventStart.setDate(eventDate.getDate());
    
    eventEnd = getDateFromTimeString(allCellValuesInTab[index][2]);
    if (eventEnd == null) {
      eventEnd = new Date(eventStart);
      eventEnd.setHours(eventStart.getHours() + 1);
    }
    eventEnd.setMonth(eventDate.getMonth());
    eventEnd.setDate(eventDate.getDate());
    
    // Process the event name for the row
    eventName = allCellsInTab[index][3];      
    if(eventName == "") {
      index++;
      continue;
    }
    
    eventLoc = allCellsInTab[index][4];
    
    // Process extra columns
    descriptionText = "In Charge: " + allCellsInTab[index][5] 
                    + "\n"
                    + "Who Else: " + allCellsInTab[index][6]
                    + "\n"
                    + "Notes: " + allCellsInTab[index][7] 
                    + "\n"
                    + "Tech: " + allCellsInTab[index][8]
                    + "\n"
                    + "Childcare: " + allCellsInTab[index][9]
                    + "\n\n\n"
                    + "Updated via script at: " + Utilities.formatDate(new Date(), "EST", "MM/dd/yyyy HH:mm:ss")
    
                    
                    
  
  
    // check ministry category
    var ministry = allCellValuesInTab[index][0];
    var event = {  name: eventName, 
                   start: eventStart, 
                   end: eventEnd, 
                   options: {location: eventLoc, 
                             description: descriptionText}
                  }
    if (ministry == "NC State") {
      ncs_events.push(event);
    } else if (ministry == "Churchwide" || ministry == "College") {
      churchwide_events.push(event);
    } else if (ministry == "UNC") {
      unc_events.push(event);
    }
    index++;
  }

  events = ncs_calendar.getEventsForDay(currentDate);
  events = events.concat(unc_calendar.getEventsForDay(currentDate));
  events = events.concat(churchwide_calendar.getEventsForDay(currentDate));
  for(var e in events){
    var name = events[e].getTitle();
    // delete all calendar events that were created by the script
    // except for all day events
    if (events[e].getDescription().indexOf("Updated via script at:") > 0 ) {
      toDelete[events[e].getId()] = events[e];
    }
  }
  for(var key in toDelete){
    var title = toDelete[key].getTitle();
    toDelete[key].deleteEvent();
    Utilities.sleep(5);
  }
  for (var e in ncs_events) {
    var start = ncs_events[e]['start'].getHours();
    var end = ncs_events[e]['end'].getHours();
    ncs_calendar.createEvent(ncs_events[e]['name'], ncs_events[e]['start'], ncs_events[e]['end'], ncs_events[e]['options']);
    Utilities.sleep(5);
  }
  for (var e in unc_events) {
    var start = unc_events[e]['start'].getHours();
    var end = unc_events[e]['end'].getHours();
    unc_calendar.createEvent(unc_events[e]['name'], unc_events[e]['start'], unc_events[e]['end'], unc_events[e]['options']);
    Utilities.sleep(5);
  }
  for (var e in churchwide_events) {
    var start = churchwide_events[e]['start'].getHours();
    var end = churchwide_events[e]['end'].getHours();
    churchwide_calendar.createEvent(churchwide_events[e]['name'], churchwide_events[e]['start'], churchwide_events[e]['end'], churchwide_events[e]['options']);
    Utilities.sleep(5);
  }
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

// Helper Function
// Gets current date in MM/dd DAY format (eg 1/1 Fri), with leading zeros removed
function formatDate(date) {
  var day = date.getDate();
  var month = date.getMonth() + 1;
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var dayOfWeek = days[date.getDay()];
  return "" + month + "/" + day + " " + dayOfWeek;
}


