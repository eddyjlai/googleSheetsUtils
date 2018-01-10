function splitTime(time) {
  var regExp = new RegExp("^(\\d{1,2}):?(\\d{2})?\\s?([a,p]?[m]?)-(\\d{1,2}):?(\\d{2})?([a,p]?[m]?)?", "gi");
  return regExp.exec(time);
}

function splitDate(date) {
  var regExp = new RegExp("^(\\d{1,2})\\/(\\d{0,2})\\s([a-z]{3})$", "gi");
  return regExp.exec(date);
}

function convertDateToUTC(date) { return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()); }

// for debugging purposes
function clearAllEvents() {
  var myCalendar = CalendarApp.getCalendarsByName("ATRDemo")[0];
  var now = new Date();
  events = myCalendar.getEvents(new Date(now.getTime() - 80*24*60*60*1000), new Date(now.getTime() + 1*24*60*60*1000));
  for(var e in events){
    events[e].deleteEvent();
  }
}