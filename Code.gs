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
        if (pto.start.date) {
          start = Utilities.parseDate(pto.start.date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssX');
          end = Utilities.parseDate(pto.end.date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssX');
        } else {
          start = Utilities.parseDate(pto.start.dateTime, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssX');
          end = Utilities.parseDate(pto.end.dateTime, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssX');
        }
        let calendar = CalendarApp.getCalendarById(calendarId);
        if (calendar === null) {
          console.error("No calendar found under %s", calendarId);
        } else {
          calendar.createEvent(title, start, end);
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
    let start = parseDate(event.start.dateTime);
    let end = parseDate(event.end.dateTime);
    return hoursBetween(start, end) > 4;
  }
}

function parseDate(dateTime) {
  return Utilities.parseDate(dateTime, "UTC", 'yyyy-MM-dd\'T\'HH:mm:ssX');
}

function hoursBetween(startDate, endDate) {
  return (endDate - startDate) / 3600000;
}
