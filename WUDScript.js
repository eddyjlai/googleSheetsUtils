function syncCalendarScript() { 
  function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }
  function isValidDate(d) {
    if ( Object.prototype.toString.call(d) !== "[object Date]" )
      return false;
    return !isNaN(d.getTime());
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INSERT TAB NAME HERE"); // Put the tab name here
  var myCalendar = CalendarApp.getCalendarsByName("INSERT CALENDAR NAME HERE")[0]; // Put the Calendar name here
  if( sheet == null || myCalendar == null) {return;}
  var firstColumn = "A";
  var lastColumn = "H";

  var hasExtraColumn1 = true;
  var descriptionText1 = "In Charge: "; // Modify these as needed for each column
  var hasExtraColumn2 = true;
  var descriptionText2 = "Food: ";
  var hasExtraColumn3 = true;
  var descriptionText3 = "Notes: ";

  var columnRange = firstColumn + ":" + lastColumn;
  var allCellsInWUD = sheet.getRange(columnRange).getValues();
  
  var numRows = sheet.getLastRow();
  var index = 0;
  var dateOnCell = null;
  var curDateValue = null;
  var eventName = null;
  var eventStart = null;
  var eventEnd = null;
  var eventLoc = null;
  var events = null;
  var toCreate = {};
  var matching = null;
  var firstDateValue = null;
  var missingEnd = false;
  var toDelete = {};
  var descriptionText; 
  var today = new Date();
  var threeDaysAgo = today.getTime() - 1000*60*60*24*3;
  var key;
  
  while (index < numRows) {
    dateOnCell = allCellsInWUD[index][0];
    if(isValidDate(dateOnCell)){
      var tempDate = convertDateToUTC(dateOnCell);
      if (tempDate.getTime() >= threeDaysAgo) {
        break;
      }
    }
    index++;
  }
  
  while( index < numRows)
  {
    dateOnCell = allCellsInWUD[index][0];
    if(curDateValue != null || isValidDate(dateOnCell)){
      if(isValidDate(dateOnCell)){
        curDateValue = convertDateToUTC(dateOnCell);
        if (firstDateValue == null) {
          firstDateValue = curDateValue;
          firstDateValue.setHours(0,0,0,0);
        }
      }
      
      eventStart = allCellsInWUD[index][1];
      eventEnd = allCellsInWUD[index][2];
      eventName = allCellsInWUD[index][3];      
      if(eventStart == "") {
        index++;
        continue;
      }
      if (!isValidDate(eventStart)) {
        if (eventStart && (eventStart.toString().toLowerCase() === "all day" || eventStart.toString().toLowerCase().equals === "allday")) {
          eventLoc = allCellsInWUD[index][4];
          descriptionText = "";
          if (hasExtraColumn1) {
            descriptionText += descriptionText1 + allCellsInWUD[index][5];
          }
          if (hasExtraColumn2) {
            descriptionText += "\n" + descriptionText2 + allCellsInWUD[index][6];
          }
          if (hasExtraColumn3) {
            descriptionText += "\n" + descriptionText3 + allCellsInWUD[index][7];
          }
          eventEnd = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate()+1, curDateValue.getHours(), 0);
          key = "allday;" + day + ";" + eventName;
          toCreate[key] = {name: eventName, start: curDateValue, end: eventEnd, 
                         options: {location: eventLoc, 
                                   description: descriptionText}, date: curDateValue, allDay: true};
          
        }
        index++;
        continue;
      }
      if (eventEnd == "" || !isValidDate(eventEnd)) {
        eventEnd = new Date(eventStart);
        missingEnd = true;
      }
      // sometimes these have issues with timezones. You can use getUTCHours() and getUTCMinutes() if you're running into timezone issues
      eventStart = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate(), eventStart.getHours(), eventStart.getMinutes());
      eventEnd = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate(), eventEnd.getHours(), eventEnd.getMinutes());
      if (missingEnd) {
        eventEnd.setTime(eventStart.getTime() + 60*60*1000); // add an hour if no end time is set
        missingEnd = false;
      }
      if (eventEnd.getTime() < eventStart.getTime()) { //hack fix for when the end time is 12AM or something. Add 24 hours to move it to next day.
        eventEnd.setTime(eventEnd.getTime() + 1000*60*60*24);
      }
      eventLoc = allCellsInWUD[index][4];
      descriptionText = "";
      if (hasExtraColumn1) {
        descriptionText += descriptionText1 + allCellsInWUD[index][5];
      }
      if (hasExtraColumn2) {
        descriptionText += "\n" + descriptionText2 + allCellsInWUD[index][6];
      }
      if (hasExtraColumn3) {
        descriptionText += "\n" + descriptionText3 + allCellsInWUD[index][7];
      }
      key = "" + eventStart.getTime() + ";" + eventEnd.getTime() + ";" + eventName;
      toCreate[key] = {name: eventName, start: eventStart, end: eventEnd, 
                     options: {location: eventLoc, 
                               description: descriptionText}, date: curDateValue, allDay: false};
    }
    index++;
  }
//  Logger.log(toCreate);
//  Logger.log(".................");
  while (firstDateValue.getTime() <= curDateValue.getTime()) {
    events = myCalendar.getEventsForDay(firstDateValue);
    for(var e in events){
      matching = null;
      if (events[e].isAllDayEvent())
        key = "allday;" + events[e].getAllDayStartDate().getDate() + ";" + events[e].getTitle();
      else
        key = "" + events[e].getStartTime().getTime() + ";" + events[e].getEndTime().getTime() + ";" + events[e].getTitle();
      if (!(key in toCreate) && (events[e].getStartTime().valueOf() >= firstDateValue.getTime().valueOf() || events[e].isAllDayEvent())) {
        // can update who's in charge, location, and who's involved if needed, then pop the event
        toDelete[events[e].getId()] = events[e];
      } else {
        if (events[e].getLocation() != toCreate[key]['options']['location'] || events[e].getDescription() != toCreate[key]['options']['description']) {
          events[e].setLocation(toCreate[key]['options']['location']);
          events[e].setDescription(toCreate[key]['options']['description']);
        }
        delete toCreate[key];
      }

    }
    firstDateValue.setDate(firstDateValue.getUTCDate() + 1);
  }
//  Logger.log(toDelete);
//  Logger.log(toCreate);
  var counter = 0;
  for (var e in toCreate) {
    if (toCreate[e]['allDay']) {
      myCalendar.createAllDayEvent(toCreate[e]['name'], toCreate[e]['date'], toCreate[e]['options']);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(1000);
    } else {
      myCalendar.createEvent(toCreate[e]['name'], toCreate[e]['start'], toCreate[e]['end'], toCreate[e]['options']);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(1000);
    }
  }
  for(var key in toDelete){
    toDelete[key].deleteEvent();
    counter++;
    if (counter % 20 == 0)
      Utilities.sleep(1000);
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

