let lastWeek = new Date();
lastWeek.setDate(lastWeek.getDate() - 7);
const nextYear = new Date();
nextYear.setFullYear(nextYear.getFullYear() + 1);
const knownGroups = new Map();
const knownUsers = new Set();
let errors = false;

function sync() {
  let userMap = new Map();
  const calendars = getTeamCalendarSetup();
  for (let calendarSetup of calendars) {
    let resolvedEmails = [];
    for (let email of calendarSetup.emails) {
      try {
        resolvedEmails = resolvedEmails.concat(getAllMembers(email));
      } catch (e) {
        resolvedEmails.push(email);
      }
    }
    calendarSetup.resolvedEmails = resolvedEmails;

    const optSince = calendarSetup.lastRun;
    for (let email of resolvedEmails) {
      if (!userMap.has(email) || isAfter(userMap.get(email).optSince, optSince)) {
        userMap.set(email, {optSince});
      }
    }
  }
  for (const [email, json] of userMap) {
    const optSince = json.optSince;
    const userPTO = getUserPTO(email, lastWeek, nextYear, optSince);
    json.userPTO = userPTO;
  }

  let syncTime = new Date();
  let nameMap = new Map();
  for (let calendarSetup of calendars) {
    let calendar = CalendarApp.getCalendarById(calendarSetup.calendarId);
    if (calendar === null) {
      Logger.log("No calendar found under " + calendarSetup.calendarId);
      continue;
    }
    for (let email of calendarSetup.resolvedEmails) {
      for (let pto of userMap.get(email).userPTO) {
        let imported = calendar.getEvents(lastWeek, nextYear, {search: pto.htmlLink});
        if (pto.status === 'cancelled') {
          for (let existing of imported) {
            if (imported.length > 1) {
              Logger.log("Duplicate team calendar event found for entry: " + existing);
            }
            existing.deleteEvent();
            Logger.log("Deleted event for " + pto.htmlLink);
          }
        } else {
          if (imported.length === 0) {
            if (!nameMap.has(email)) {
              nameMap.set(email, getDisplayName(email));
            }
            let mappedEvent = mapEvent(nameMap.get(email), pto.start, pto.end);
            try {
              calendar.createAllDayEvent(mappedEvent.title, mappedEvent.start, mappedEvent.end, {description: pto.htmlLink});
              Logger.log("Created all day event for " + pto + " in " + calendarSetup.calendarId);
            } catch( error ) {
              Logger.log('Error creating all day event for ' + pto + " in " + calendarSetup.calendarId);
              errors = true;
            }
          } else {
            for (let existing of imported) {
              if (imported.length > 1) {
                Logger.log("Duplicate team calendar event found for entry: " + existing);
              }
              //Just updating start and end, so no need for person name
              let mappedEvent = mapEvent(email, pto.start, pto.end);
              existing.setAllDayDates(mappedEvent.start, mappedEvent.end);
              Logger.log("Updated event for " + pto.htmlLink);
            }
          }
        }
      }
    }
    if (!errors) {
      getCalendarSheet().getRange(calendarSetup.row, 5, 1, 1).setValue(syncTime);
    }
  }
}

function mapEvent(name, origStart, origEnd) {
  let title = name.concat(" - OOO");
  let start;
  let end;
  if (origStart.date) {
    start = parseDate(origStart.date);
    end = parseDate(origEnd.date);
  } else {
    start = parseDateTime(origStart.dateTime);
    end = parseDateTime(origEnd.dateTime);
  }
  return {title, start, end};
}

function getTeamCalendarSetup() {
  let calendarSetup = [];
  let sheet = getCalendarSheet();
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    let calendarId = sheet.getRange(row, 4, 1, 1).getValue();
    if (calendarId.length > 0 ) {
      let emails = sheet.getRange(row, 2, 1, 1).getValue().split(',').map(entry => entry.trim());
      let lastRun = sheet.getRange(row, 5, 1, 1).getValue();
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
      Logger.log('Error retrieving events for ' + email + ': ' + e.toString() + '; skipping');
      errors = true;
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
 * @param {string} groupEmail the e-mail address of the group.
 * @return {object} direct and indirect members.
 */
function getAllMembers(groupEmail) {
  if (knownGroups.has(groupEmail)) {
    return knownGroups.get(groupEmail);
  } else if (knownUsers.has(groupEmail)) {
    return groupEmail;
  }
  var group = GroupsApp.getGroupByEmail(groupEmail);
  var users = group.getUsers();
  var childGroups = group.getGroups();
  for (let i = 0; i < childGroups.length; i++) {
    var childGroup = childGroups[i];
    users = users.concat(getAllMembers(childGroup.getEmail()));
  }
  // Remove duplicate members
  let userEmails = new Set();
  for (let i = 0; i < users.length; i++) {
    userEmails.add(users[i].getEmail());
    knownUsers.add(users[i].getEmail()); //For efficiency and later use
  }
  const members = Array.from(userEmails);
  knownGroups.set(groupEmail, members)
  return members;
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
