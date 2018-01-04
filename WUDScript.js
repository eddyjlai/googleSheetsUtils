
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

function syncCalendarK1() { 
  function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("K1"); // Put the tab name here
  var myCalendar = CalendarApp.getCalendarsByName("DJ HG Calendar")[0]; // Put the Calendar name here
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
  var toCreate = [];
  var matching = null;
  var firstDateValue = null;
  var missingEnd = false;
  var toDelete = {};
  var descriptionText; 
  var today = new Date();
  var threeDaysAgo = today.getTime() - 1000*60*60*24*3;
  
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
        if (eventStart && eventStart.toString().toLowerCase().equals("all day")) {
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
          toCreate.push({name: eventName, start: curDateValue, end: eventEnd, 
                         options: {location: eventLoc, 
                                   description: descriptionText}, date: curDateValue, allDay: true});
          
        }
        index++;
        continue;
      }
      if (eventEnd == "") {
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
      toCreate.push({name: eventName, start: eventStart, end: eventEnd, 
                     options: {location: eventLoc, 
                               description: descriptionText}, date: curDateValue, allDay: false});
    }
    index++;
  }
//  Logger.log(toCreate);
//  Logger.log(".................");
  while (firstDateValue.getTime() <= curDateValue.getTime() && toCreate.length > 0) {
    events = myCalendar.getEventsForDay(firstDateValue);
    for(var e in events){
      matching = null;
      for (var item in toCreate) {
        if (events[e].isAllDayEvent() && toCreate[item]['allDay'] == true) {
          if (toCreate[item]['date'].getTime() == firstDateValue.getTime() 
            && toCreate[item]['name'] == events[e].getTitle() 
            && events[e].getLocation() == toCreate[item]['options']['location']
            && events[e].getDescription() == toCreate[item]['options']['description']) {
              matching = toCreate[item];
              break;
            }
        }
        if (toCreate[item]['date'].getTime() == firstDateValue.getTime() 
          && toCreate[item]['name'] == events[e].getTitle()
          && events[e].getStartTime().valueOf() == toCreate[item]['start'].valueOf()
          && events[e].getEndTime().valueOf() == toCreate[item]['end'].valueOf()
          && events[e].getLocation() == toCreate[item]['options']['location']
          && events[e].getDescription() == toCreate[item]['options']['description']) {
          matching = toCreate[item];
          break;
        }
      }
      if (matching == null && (events[e].getStartTime().valueOf() >= firstDateValue.getTime().valueOf() || events[e].isAllDayEvent())) { //calendar event not found, deleting 
        toDelete[events[e].getId()] = events[e];
      } else { //calendar event found, don't create (i.e. remove from create list)
        toCreate.splice(toCreate.indexOf(matching), 1);
      }
    }
    firstDateValue.setDate(firstDateValue.getUTCDate() + 1);
  }
//  Logger.log(toDelete);
//  Logger.log(toCreate);
  for (var e in toCreate) {
    if (toCreate[e]['allDay']) {
      myCalendar.createAllDayEvent(toCreate[e]['name'], toCreate[e]['date'], toCreate[e]['options']);
    } else {
      myCalendar.createEvent(toCreate[e]['name'], toCreate[e]['start'], toCreate[e]['end'], toCreate[e]['options']);
    }
  }
  for(var key in toDelete){
    toDelete[key].deleteEvent();
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

