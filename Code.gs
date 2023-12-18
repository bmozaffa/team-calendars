let lastWeek = new Date();
lastWeek.setDate(lastWeek.getDate() - 7);
let nextYear = new Date();
nextYear.setFullYear(nextYear.getFullYear() + 1);

function sync() {
  let calendarMap = getTeamCalendarSetup();
  let userMap = new Map();
  for (let [calendarId, emails] of calendarMap) {
    for (let email of emails) {
      if (!userMap.has(email)) {
        userMap.set(email, getUserPTO(email));
      }
      ptoEvents = userMap.get(email);
      for (let pto of ptoEvents) {
        let title = email.concat(" - OOO");
        let start;
        let end;
        if (pto.start.date) {
          start = parseDate(pto.start.date);
          end = parseDate(pto.end.date);
        } else {
          start = parseDateTime(pto.start.dateTime);
          end = parseDateTime(pto.end.dateTime);
        }
        let calendar = CalendarApp.getCalendarById(calendarId);
        if (calendar === null) {
          console.error("No calendar found under %s", calendarId);
        } else {
          calendar.createAllDayEvent(title, start, end);
        }
      }
    }
  }
}

function getTeamCalendarSetup() {
  let map = new Map();
  let sheet = SpreadsheetApp.openById("1PjNQylEVTOHmoLbxzjtQmufyzvRKYYrINLTnAT3ZzQ0").getSheetByName("main");
  let values = sheet.getRange(2, 1, sheet.getLastRow(), 2).getValues();
  values.forEach(value => {
    if (value[0].length > 0 ) {
      let emails = value[1].split(',');
      map.set(value[0], emails);
    }
  });
  return map;
}

function getUserPTO(email) {
  let params = {
    timeMin: Utilities.formatDate(lastWeek, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
    timeMax: Utilities.formatDate(nextYear, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
    eventTypes: ['outOfOffice'],
    showDeleted: false,
  };
  let response = Calendar.Events.list(email, params);
  let events = [];
  events = events.concat(response.items.filter(function(event) {
    return include(event);
  }));
  return events;
}

function include(event) {
  if (event.start != null && event.start.date != null ) {
    //If a date and not a dateTime, this is a full day event as entered
    return true;
  } else {
    let start = parseDateTime(event.start.dateTime);
    let end = parseDateTime(event.end.dateTime);
    return hoursBetween(start, end) > 23;
  }
}

function parseDateTime(dateTime) {
  return parseDate(dateTime.substring(0, 10));
}

function parseDate(date) {
  return Utilities.parseDate(date, Intl.DateTimeFormat().resolvedOptions().timeZone, 'yyyy-MM-dd');
}

function hoursBetween(startDate, endDate) {
  return (endDate - startDate) / 3600000;
}
