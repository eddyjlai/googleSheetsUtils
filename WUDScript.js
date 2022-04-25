// CONFIGURATION SECTION
var isTesting = false; // toggle between real and test calendar
var calendarName = isTesting ? "test_calendar_jerell" : "GPNB - WUD Calendar"; // Put the Calendar name here
var tabName = "SCHEDULE"; // Put the tab name here

// These column labels span two columns on row #2, but are specified in more detail on row #3
var doubleColumnSubgroupMapping = {
  "WHEN": [
    "START", "END"
  ],
  "WHO": [
    "LEAD", "OTHERS"
  ]
}

var firstColumn = "A";
var lastColumn = "K";
var extraColumns = ["GROUP", "LEAD", "OTHERS", "CHILDCARE", "FOOD", "NOTES"]; // Extra info columns to append as a string to the calendar event description
// format:
// GROUP: test group value
// LEAD: test lead value
// etc

// https://developers.google.com/apps-script/reference/calendar/event-color
var groupColorKey = {
  "Churchwide": "1",
  "Hosting": "2",
  "Rutgers": "4",
  "Youth": "5",
  "1st Years": "10",
  "Princeton": "6",
  "Canceled": "8"
}

// Row 2 contains all the column headers
function buildColumnKeyMap(headerLabelsRow) {
  let columnLetterIncrementer = "A";
  return headerLabelsRow.reduce((prev, curr) => {
    const columnName = curr.toUpperCase().trim();
    if (!columnName) return prev;

    const subGroupNames = doubleColumnSubgroupMapping[columnName];

    if (subGroupNames) {
      subGroupNames.forEach((subName) => {
        prev[subName] = columnLetterIncrementer;
        columnLetterIncrementer = nextLetter(columnLetterIncrementer);
      });
    } else {
      prev[columnName] = columnLetterIncrementer;
      columnLetterIncrementer = nextLetter(columnLetterIncrementer);
    }

    return prev;
  }, {});
}
// helper function to get next column letter i.e. A -> B -> C
function nextLetter(s){
    return s.replace(/([a-zA-Z])[^a-zA-Z]*$/, function(a){
        var c= a.charCodeAt(0);
        switch(c){
            case 90: return 'A';
            case 122: return 'a';
            default: return String.fromCharCode(++c);
        }
    });
}

const sheetNamePrefixes = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];

function syncCalendarAllSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetNamePrefixes.some(p => sheetName.toLowerCase().includes(p))) {
      console.log(sheetName);
      syncCalendarScript(sheet);
    }

  }
}

// SCRIPT SECTION
function syncCalendarScript(sheet) {
  function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }
  function isValidDate(d) {
    if ( Object.prototype.toString.call(d) !== "[object Date]" )
      return false;
    return !isNaN(d.getTime());
  }

  var myCalendar = CalendarApp.getCalendarsByName(calendarName)[0];
  if( sheet == null || myCalendar == null) {
    console.log("sheet or calendar not found");
    return;
  }

  var columnRange = firstColumn + ":" + lastColumn;
  var allCellsInWUD = sheet.getRange(columnRange).getValues();
  var headerLabelsRow = allCellsInWUD[1];
  var columnKeys = buildColumnKeyMap(headerLabelsRow)

  var columnIndex = Object.keys(columnKeys).reduce((acc, curr) => {
    var columnName = curr;
    var columnLetter = columnKeys[columnName];
    var columnIndex = columnLetter.toLowerCase().charCodeAt(0) - 97
    acc[columnName] = columnIndex;
    return acc;
  }, {})

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
  var dateOnCell;

  // get date from sheet name
  // test for 2 formats
  let dateFromSheetName;
  // format 1: EX - Tue(2/10)
  if (sheet.getName().match(/\(([^)]+)\)/)?.length === 2) {
    dateFromSheetName = sheet.getName().match(/\(([^)]+)\)/)[1];
  } else {
  // format 2: EX - Tue 2/10
    dateFromSheetName = sheet.getName().split(' ')[0];
  }
  const year = today.getFullYear();
  dateOnCell = new Date(dateFromSheetName + '/' + year);

  if(isValidDate(dateOnCell)){
    var tempDate = convertDateToUTC(dateOnCell);
    if (tempDate.getTime() < threeDaysAgo) { // don't update sheets older than 3 days ago
      return;
    }
  } else {
    return; // invalid date, skip sheet
  }
  index = 3; // account for headers on all tabs now
  while(index < numRows)
  {
    if(curDateValue != null || isValidDate(dateOnCell)){
      if(isValidDate(dateOnCell)){
        curDateValue = convertDateToUTC(dateOnCell);
        if (firstDateValue == null) {
          firstDateValue = curDateValue;
          firstDateValue.setHours(0,0,0,0);
        }
      }

      eventStart = allCellsInWUD[index][columnIndex["START"]];
      eventEnd = allCellsInWUD[index][columnIndex["END"]];
      eventName = allCellsInWUD[index][columnIndex["WHAT"]];
      eventGroup = allCellsInWUD[index][columnIndex["GROUP"]];
      if(allCellsInWUD[index].every(value => value === "")) {
        index++;
        continue;
      }
      if (!isValidDate(eventStart)) {
        if ((eventStart && (eventStart.toString().toLowerCase() === "all day" || eventStart.toString().toLowerCase().equals === "allday")) || (eventStart === "" && eventEnd === "")) {
          eventLoc = allCellsInWUD[index][columnIndex["WHERE"]];
          descriptionText = "";

          extraColumns.forEach(columnName => {
            var currColIndex = columnIndex[columnName];
            var currColValue = allCellsInWUD[index][currColIndex];
            var currDescriptionText = columnName + ": " + currColValue;
            descriptionText += currDescriptionText + "\n";
          });

          var allDayEventStart = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate(), curDateValue.getHours(), 0);
          eventEnd = new Date(curDateValue.getUTCFullYear(), curDateValue.getUTCMonth(), curDateValue.getUTCDate()+1, curDateValue.getHours(), 0);
          key = "allday;" + curDateValue.getDate() + ";" + eventName;
          toCreate[key] = {name: eventName, start: allDayEventStart, end: eventEnd, group: eventGroup,
                         options: {location: eventLoc,
                                   description: descriptionText}, date: allDayEventStart, allDay: true};

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
      eventLoc = allCellsInWUD[index][columnIndex["WHERE"]];
      descriptionText = "";

      extraColumns.forEach(columnName => {
        var currColIndex = columnIndex[columnName];
        var currColValue = allCellsInWUD[index][currColIndex];
        var currDescriptionText = columnName + ": " + currColValue;
        descriptionText += currDescriptionText + "\n";
      });
      // prepends CANCELED to lines with strike through applied
      var currentRange = "A" + (index + 1) + ":" + "K" + (index + 1);
      var eventRange = sheet.getRange(currentRange);
      // If line is striked through, or if canceled is in the name
      var isEventCanceledStrikeThrough = eventRange.getFontLine() === 'line-through';
      var isEventcanceledInName = (new RegExp('canceled|cancelled', 'i')).test(eventName);

      // getting all links in cells
      const getHyperlinks = (eventRange) => {
        const response = eventRange.getRichTextValues();
        return response[0]
          .filter(cell => cell.getLinkUrl())
          .map(cell => cell.getLinkUrl() + ' \n');
      }
      const linkylinks = getHyperlinks(eventRange);
      descriptionText += "LINKS: " + linkylinks.join();
      Logger.log(linkylinks);


      var updatedEventName = (!isEventcanceledInName && isEventCanceledStrikeThrough) ? "[CANCELED] " + eventName : eventName;
      eventGroup = (isEventCanceledStrikeThrough || isEventcanceledInName) ? "Canceled" : eventGroup;

      key = "" + eventStart.getTime() + ";" + eventEnd.getTime() + ";" + updatedEventName;
      toCreate[key] = {name: updatedEventName, start: eventStart, end: eventEnd, group: eventGroup,
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
        if (events[e].getLocation() != toCreate[key]?.options?.location || events[e].getDescription() != toCreate[key]?.options?.description) {
          events[e].setLocation(toCreate[key]?.options?.location);
          events[e].setDescription(toCreate[key]?.options?.description);
        }
        delete toCreate[key];
      }

    }
    firstDateValue.setDate(firstDateValue.getUTCDate() + 1);
  }
 Logger.log(toDelete);
//  Logger.log(toCreate);
  var counter = 0;
  for (var e in toCreate) {
    if (toCreate[e]['allDay']) {
      let createdEvent = myCalendar.createAllDayEvent(toCreate[e]['name'], toCreate[e]['date'], toCreate[e]?.options);
      let calendarEventColor = groupColorKey[toCreate[e]['group']] ?? "7";
      createdEvent.setColor(calendarEventColor);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(2000);
    } else {
      let createdEvent = myCalendar.createEvent(toCreate[e]['name'], toCreate[e]['start'], toCreate[e]['end'], toCreate[e]?.options);
      let calendarEventColor = groupColorKey[toCreate[e]['group']] ?? "7";
      createdEvent.setColor(calendarEventColor);
      counter++;
      if (counter % 20 == 0)
        Utilities.sleep(2000);
    }
  }
  for(var key in toDelete){
    let eventToDelete = toDelete[key];
    let eventExists = !!myCalendar.getEventById(eventToDelete.getId());
    // Logger.log(eventToDelete.getId());
    if (eventExists) {
      eventToDelete.deleteEvent();
    }
    counter++;
    if (counter % 20 == 0)
      Utilities.sleep(2000);
  }
}

// for debugging purposes. Delete all events in the last 8 days and 20 days from now
function clearAllEvents() {
  var myCalendar = CalendarApp.getCalendarsByName(calendarName)[0];
  var now = new Date();
  events = myCalendar.getEvents(new Date(now.getTime() - 8*24*60*60*1000), new Date(now.getTime() + 20*24*60*60*1000));
  for(var e in events){
    events[e].deleteEvent();
  }
}