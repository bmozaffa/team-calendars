let lastWeek = new Date();
lastWeek.setDate(lastWeek.getDate() - 7);
let nextYear = new Date();
nextYear.setFullYear(nextYear.getFullYear() + 1);

function sync() {
  let syncTime = new Date();
  let calendarSetupArray = getTeamCalendarSetup();
  let userMap = new Map();
  let nameMap = new Map();
  for (let calendarSetup of calendarSetupArray) {
    let optSince = calendarSetup.lastRun;
    let resolvedEmails = [];
    for (let email of calendarSetup.emails) {
      try {
        resolvedEmails = resolvedEmails.concat(getAllMembers(email));
      } catch (e) {
        resolvedEmails.push(email);
      }
    }
    for (let email of resolvedEmails) {
      if (!userMap.has(email) || isAfter(userMap.get(email).optSince, optSince)) {
        let userPTO = getUserPTO(email, lastWeek, nextYear, optSince);
        userMap.set(email, {optSince, userPTO});
        nameMap.set(email, getDisplayName(email));
      }
      for (let pto of userMap.get(email).userPTO) {
        if (pto.status === 'canceled') {
          //TODO find and delete from calendar
        } else {
          //TODO later figure out updates too!
          let title = nameMap.get(email).concat(" - OOO");
          let start;
          let end;
          if (pto.start.date) {
            start = parseDate(pto.start.date);
            end = parseDate(pto.end.date);
          } else {
            start = parseDateTime(pto.start.dateTime);
            end = parseDateTime(pto.end.dateTime);
          }
          let calendar = CalendarApp.getCalendarById(calendarSetup.calendarId);
          if (calendar === null) {
            console.error("No calendar found under %s", calendarSetup.calendarId);
          } else {
            calendar.createAllDayEvent(title, start, end);
          }
        }
        console.log(pto);
      }
    }
    getCalendarSheet().getRange(calendarSetup.row, 3, 1, 1).setValue(syncTime);
  }

  // // Determines the time the the script was last run.
  // let lastRun = PropertiesService.getScriptProperties().getProperty('lastRun');
  // lastRun = lastRun ? new Date(lastRun) : null;

  // for (let [calendarId, emails] of calendarMap) {
  //   let resolvedEmails = [];
  //   for (let email of emails) {
  //     try {
  //       resolvedEmails = resolvedEmails.concat(getAllMembers(email));
  //     } catch (e) {
  //       resolvedEmails.push(email);
  //     }
  //   }
  //   for (let email of resolvedEmails) {
  //     if (!userMap.has(email)) {
  //       userMap.set(email, getUserPTO(email, lastWeek, nextYear, lastRun));
  //       nameMap.set(email, getDisplayName(email));
  //     }
  //     ptoEvents = userMap.get(email);
  //     for (let pto of ptoEvents) {
  //       let title = nameMap.get(email).concat(" - OOO");
  //       let start;
  //       let end;
  //       if (pto.start.date) {
  //         start = parseDate(pto.start.date);
  //         end = parseDate(pto.end.date);
  //       } else {
  //         start = parseDateTime(pto.start.dateTime);
  //         end = parseDateTime(pto.end.dateTime);
  //       }
  //       let calendar = CalendarApp.getCalendarById(calendarId);
  //       if (calendar === null) {
  //         console.error("No calendar found under %s", calendarId);
  //       } else {
  //         calendar.createAllDayEvent(title, start, end);
  //       }
  //     }
  //   }
  // }
  PropertiesService.getScriptProperties().setProperty('lastRun', new Date());
}

function getTeamCalendarSetup() {
  let calendarSetup = [];
  let sheet = getCalendarSheet();
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    let calendarId = sheet.getRange(row, 1, 1, 1).getValue();
    if (calendarId.length > 0 ) {
      let emails = sheet.getRange(row, 2, 1, 1).getValue().split(',');
      let lastRun = sheet.getRange(row, 3, 1, 1).getValue();
      calendarSetup.push({calendarId, emails, row, lastRun});
    }
  }
  return calendarSetup;
}

function getCalendarSheet() {
  return SpreadsheetApp.openById("1PjNQylEVTOHmoLbxzjtQmufyzvRKYYrINLTnAT3ZzQ0").getSheetByName("main");
}

/**
 * In a given user's calendar, looks for OutOfOffice
 * events within the specified date range and returns any such events
 * found.
 * @param {string} email The user email to retrieve events for.
 * @param {Date} start The starting date of the range to examine.
 * @param {Date} end The ending date of the range to examine.
 * @param {Date} optSince A date indicating the last time this script was run.
 * @return {Calendar.Event[]} An array of calendar events.
 */
function getUserPTO(email, start, end, optSince) {
  let params = {
    timeMin: formatDateAsRFC3339(start),
    timeMax: formatDateAsRFC3339(end),
    eventTypes: ['outOfOffice'],
    showDeleted: true,
  };
  if (optSince) {
    // This prevents the script from examining events that have not been
    // modified since the specified date (that is, the last time the
    // script was run).
    params.updatedMin = formatDateAsRFC3339(optSince);
  }
  let pageToken = null;
  let events = [];
  do {
    params.pageToken = pageToken;
    let response;
    try {
      response = Calendar.Events.list(email, params);
    } catch (e) {
      console.error('Error retriving events for %s: %s; skipping',
          email, e.toString());
      continue;
    }
    events = events.concat(response.items.filter(function(event) {
      return shouldImportEvent(event);
    }));
    pageToken = response.nextPageToken;
  } while (pageToken);
  return events;
}

/**
 * Returns an RFC3339 formated date String corresponding to the given
 * Date object.
 * @param {Date} date a Date.
 * @return {string} a formatted date string.
 */
function formatDateAsRFC3339(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ');
}

function shouldImportEvent(event) {
  if (event.start != null && event.start.date != null ) {
    //If a date and not a dateTime, this is a full day event as entered
    return true;
  } else {
    let start = parseDateTime(event.start.dateTime);
    let end = parseDateTime(event.end.dateTime);
    return hoursBetween(start, end) > 23;
  }
}

function isAfter(date1, date2) {
  if (!date1) {
    return false;
  } else if (date2) {
    return true;
  } else {
    return date1 > date2;
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

/**
* Get both direct and indirect members (and delete duplicates).
* @param {string} the e-mail address of the group.
* @return {object} direct and indirect members.
*/
function getAllMembers(groupEmail) {
  var group = GroupsApp.getGroupByEmail(groupEmail);
  var users = group.getUsers();
  var childGroups = group.getGroups();
  for (var i = 0; i < childGroups.length; i++) {
    var childGroup = childGroups[i];
    users = users.concat(getAllMembers(childGroup.getEmail()));
  }
  // Remove duplicate members
  let userEmails = new Set();
  for (var i = 0; i < users.length; i++) {
    userEmails.add(users[i].getEmail());
  }
  return Array.from(userEmails);
}

function getDisplayName(email) {
  let person = People.People.searchDirectoryPeople({
    readMask: 'names',
    query: email,
    sources: [
      'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
    ]
  });
  if (person.totalSize === 1) {
    return  person.people[0].names[0].displayName;
  } else {
    return email;
  }
}
