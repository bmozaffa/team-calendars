let lastWeek = new Date();
lastWeek.setDate(lastWeek.getDate() - 7);
let nextYear = new Date();
nextYear.setFullYear(nextYear.getFullYear() + 1);

function print() {
  let calendarMap = getTeamCalendarSetup();
  let userMap = new Map();
  for (let [calendarId, emails] of calendarMap) {
    for (let email of emails) {
      if (!userMap.has(email)) {
        userMap.set(email, getUserPTO(email));
      }
      ptoEvents = userMap.get(email);
      for (let pto of ptoEvents) {
        let title = pto.getSummary();
        // CalendarApp.getCalendarById(calendarId).createEvent()
      }
    }
  }

  // for (let [calendarId, emails] of map) {
  //   let calendar = CalendarApp.getCalendarById(calendarId);
  //   console.log("Calendar name is %s", calendar.getName());
  //   for (let email of emails) {
  //     let params = {
  //       timeMin: Utilities.formatDate(lastWeek, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
  //       timeMax: Utilities.formatDate(nextYear, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
  //       eventTypes: ['outOfOffice'],
  //       showDeleted: false,
  //     };
  //     let response = Calendar.Events.list(email, params);
  //     console.log("Found %d events", response.items.length);
  //     response.items.forEach(entry => {
  //       console.log(entry);
  //     });

  //   }
  // }
  // console.log(map);
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
  for (let event of response.items) {
    if (isFullDay(event)) {
      events.push(event);
    }
  }
  return events;
}

function isFullDay(event) {
  if (event.start != null && event.start.date != null ) {
    //If a date and not a dateTime, this is a full day event as entered
    return true;
  } else {
    let start = Utilities.parseDate(event.start.dateTime, "UTC", 'yyyy-MM-dd\'T\'HH:mm:ssX');
    let end = Utilities.parseDate(event.end.dateTime, "UTC", 'yyyy-MM-dd\'T\'HH:mm:ssX');
    let durationHours = (end - start) / 3600000;
    return durationHours > 4;
  }
}
