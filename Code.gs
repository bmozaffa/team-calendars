let lastWeek = new Date();
lastWeek.setDate(lastWeek.getDate() - 7);
let nextYear = new Date();
nextYear.setFullYear(nextYear.getFullYear() + 1);

function sync() {
  let calendarMap = getTeamCalendarSetup();
  let userMap = new Map();
  let nameMap = new Map();
  for (let [calendarId, emails] of calendarMap) {
    for (let email of emails) {
      if (!userMap.has(email)) {
        userMap.set(email, getUserPTO(email));
        nameMap.set(email, getDisplayName(email));
      }
      ptoEvents = userMap.get(email);
      for (let pto of ptoEvents) {
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
  var uniqueUsers = [];
  var userEmails = {};
  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    if (!userEmails[user.getEmail()]) {
      uniqueUsers.push(user);
      userEmails[user.getEmail()] = true;
    }
  }
  return uniqueUsers;
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
