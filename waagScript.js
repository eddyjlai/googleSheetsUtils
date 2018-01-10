function toCalendar() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var myCalendar = CalendarApp.getCalendarsByName("INSERT CALENDAR NAME HERE")[0]; // Put the Calendar name here
  if( myCalendar == null) {return;}
  var firstColumn = "A";
  var lastColumn = "G";
  var columnRange = firstColumn + ":" + lastColumn;

  var hasExtraColumn1 = true;
  var descriptionText1 = "In Charge: "; // Modify these as needed for each column
  var hasExtraColumn2 = true;
  var descriptionText2 = "Who: ";

  var eventName = null;
  var eventDate = null;
  var eventStart = null;
  var eventEnd = null;
  var eventLoc = null;
  var toCreate = {}; 
  //key is start.getTime();end;name, value is {...}
  //for all day event: key is allday;name
  var toDelete = {};
  var sheetNum = 0;
  var allday = false;
  var sheetName, sheetNameSplit, sheetNameDate, sheetNameDateSplit, sheetNameDay, curSheet, allRowsInSheet, index, numRows;
  var today = new Date();
  var year = today.getYear();
  var allCellsOnSheet = null;
  var index = 2;
  var startHour, startMinute, startMeridian, endHour, endMinute, endMeridian;
  var firstDateValue, lastDateValue;
  var key;

  while (sheetNum < sheets.length) {
    curSheet = sheets[sheetNum];
    sheetName = curSheet.getName();
    if (curSheet.isSheetHidden()){
      sheetNum++;
      continue;
    }
    
    sheetNameSplit = splitDate(sheetName);
    // [12/9 Sat,12,9,Sat]
    if (sheetNameSplit != undefined && sheetNameSplit.length > 3) {
      if (sheetNameSplit[1] == "" || sheetNameSplit[2] == "")
        continue;
      var month = parseInt(sheetNameSplit[1])-1;
      var day = parseInt(sheetNameSplit[2]);
      if (firstDateValue == undefined) {
        firstDateValue = new Date(year, month, day, 0, 0);
      }
      allCellsOnSheet = curSheet.getRange(columnRange).getValues();
      for (index = 2; index < curSheet.getLastRow(); index++) {
        allday = false;
        if (allCellsOnSheet[index][0].toString().toLowerCase() === "all day") {
          startHour = 0;
          startMinute = 0;
          endHour = 12;
          endMinute = 0;
          allday = true;
        } else if ( Object.prototype.toString.call(allCellsOnSheet[index][0]) === "[object Date]" ){
          // case for missing end time
          startHour = allCellsOnSheet[index][0].getHours();
          startMinute = allCellsOnSheet[index][0].getMinutes();
          endHour = startHour + 1;
          endMinute = startMinute;
        } else {
          //[4:30-6:30pm, 4, 30, , 6, 30, pm]
          eventDate = splitTime(allCellsOnSheet[index][0].toString().trim());
          if (eventDate == null)
            continue;
          startHour = parseInt(eventDate[1]);
          startMinute = eventDate[2] != null ? parseInt(eventDate[2]) : 0;
          startMeridian = eventDate[3];
          endHour = parseInt(eventDate[4]);
          endMinute = eventDate[5] != null ? parseInt(eventDate[5]) : 0;
          endMeridian = eventDate[6];
          if (endMeridian == "pm" && endHour < 12) {
            if (startMeridian == "") {
              if (startHour < endHour)
                startHour = startHour + 12; // assume the starting time is also pm if something like 2-7pm, but NOT when 11-12:30pm
              else if (startHour == endHour && (startHour < 7 || startHour == 11)) // bug for 1:15-1:30pm, or 11-11:30pm
                startHour = startHour + 12;
            }
            endHour = endHour + 12;
          }
          if (startMeridian == "pm")
            startHour = startHour + 12;
        }
        eventStart = new Date(year, month, day, startHour, startMinute);
        eventEnd = new Date(year, month, day, endHour, endMinute);
        eventName = allCellsOnSheet[index][1];
        eventLoc = allCellsOnSheet[index][2];
        descriptionText = "";
        if (hasExtraColumn1) {
          descriptionText += descriptionText1 + allCellsOnSheet[index][3];
        }
        if (hasExtraColumn2) {
          descriptionText += "\n" + descriptionText2 + allCellsOnSheet[index][4];
        }
        key = "" + eventStart.getTime() + ";" + eventEnd.getTime() + ";" + eventName;
        toCreate[key] = {name: eventName, start: eventStart, end: eventEnd, 
                       options: {location: eventLoc, 
                                 description: descriptionText}, date: eventStart, allDay: allday};
      }
    }
    sheetNum++;
  }

  // Now update the calendar
  lastDateValue = new Date(eventEnd.getYear(), eventEnd.getMonth(), eventEnd.getDate());
  lastDateValue = new Date(lastDateValue.getTime() + 1000*60*60*24);
  var events = myCalendar.getEvents(firstDateValue, lastDateValue);
  for(var e in events){
    if (events[e].isAllDayEvent())
      continue;
    key = "" + events[e].getStartTime().getTime() + ";" + events[e].getEndTime().getTime() + ";" + events[e].getTitle();
    if (key in toCreate) {
      // can update who's in charge, location, and who's involved if needed, then pop the event
      if (events[e].getLocation() != toCreate[key]['options']['location'] || events[e].getDescription() != toCreate[key]['options']['description']) {
        events[e].setLocation(toCreate[key]['options']['location']);
        events[e].setDescription(toCreate[key]['options']['description']);
      }
      delete toCreate[key];
    } else {
      toDelete[events[e].getId()] = events[e];
    }
  }
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

