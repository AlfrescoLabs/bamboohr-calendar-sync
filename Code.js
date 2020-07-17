const properties = PropertiesService.getScriptProperties();
const bambooApiKey = properties.getProperty('bambooApiKey');
const bambooCompanyDomain = properties.getProperty('bambooCompanyDomain');
const googleCalendarId = properties.getProperty('googleCalendarId');

function syncBambooHRWhosOutToGCal() {
  const calendarUsers = getCalendarUsers(googleCalendarId);
  const timeOff = getBambooTimeOff(calendarUsers);
  const timeNow = new Date();
  const existingEventIds = listEvents(googleCalendarId, {timeMin: timeNow}).map((event) => event.id);
  timeOff.forEach(function(timeOffItem) {
    const customId = generateCalendarId(timeOffItem);
    if (existingEventIds.indexOf(customId) > -1) {
      console.log('Updating existing event with ID ' + customId, timeOffItem);
      try {
        updateEvent(googleCalendarId, timeOffItem);
      } catch(e) {
        throw e;
      }
    } else {
      console.log('Creating new event with ID ' + customId, timeOffItem);
      createEvent(googleCalendarId, timeOffItem);
    }
  });
}

function getCalendarUsers(calendarId) {
  let users = [];
  try {
    const acl = Calendar.Acl.list(calendarId);
    const aclItems = acl.items;
    for (let i=0; i<aclItems.length; i++) {
      if (aclItems[i].id.indexOf('user:') === 0) {
        const userId = aclItems[i].id.replace('user:', '');
        const role = aclItems[i].role;
        //const userProfile = getAccount(userId);
        //console.log(userProfile ? userProfile.names : userId, aclItems[i].role);
        users.push(userId);
      }
    }
  }
  catch (e) {
    console.log(e);
    // no existing acl record for this user - as expected. Carry on.
  }
  return users;
}

function getBambooTimeOff(emailList) {
  const allEmployees = bambooGetEmployeeList(bambooCompanyDomain).employees;
  const selectedEmployees = filterEmployeesByEmail(allEmployees, emailList);
  const selectedEmployeeIds = selectedEmployees.map((employee) => employee.id);
  const whosOut = bambooGetWhosOut(bambooCompanyDomain);
  const teamTimeOff = whosOut.filter((timeOff) => selectedEmployeeIds.indexOf('' + timeOff.employeeId) > -1);
  return teamTimeOff;
}

function bambooApiGet(companyDomain, path) {
  const apiUrl = `https://api.bamboohr.com/api/gateway.php/${companyDomain}${path}`;
  const authToken = Utilities.base64Encode(`${bambooApiKey}:x`);
  return UrlFetchApp.fetch(apiUrl, {
    headers: {
      'Authorization': `Basic ${authToken}`,
      'Accept': 'application/json'
    }
  });
}

function bambooGetEmployeeList(companyDomain) {
  return JSON.parse(bambooApiGet(companyDomain, '/v1/employees/directory').getContentText());
}

function filterEmployeesByEmail(employees, emailList) {
  return employees.filter((employee) => employee.workEmail && emailList.indexOf(employee.workEmail) > -1);
}

function bambooGetTimeOff(companyDomain, employeeId) {
  const daysAhead = 30;
  const now = new Date();
  const start = Utilities.formatDate(now, 'GMT', 'YYYY-MM-dd');
  now.setDate(now.getDate() + daysAhead);
  const end = Utilities.formatDate(now, 'GMT', 'YYYY-MM-dd');
  return JSON.parse(bambooApiGet(companyDomain, `/v1/time_off/requests?employeeId=${employeeId}&start=${start}&end=${end}`).getContentText());
}

function bambooGetWhosOut(companyDomain) {
  const daysAhead = 30;
  const now = new Date();
  const start = Utilities.formatDate(now, 'GMT', 'YYYY-MM-dd');
  now.setDate(now.getDate() + daysAhead);
  const end = Utilities.formatDate(now, 'GMT', 'YYYY-MM-dd');
  return JSON.parse(bambooApiGet(companyDomain, `/v1/time_off/whos_out?start=${start}&end=${end}`).getContentText());
}

function listEvents(calendarId, options={}) {
  const eventList = [];
  let pageToken, events, eventItems;
  do {
    const listOptions = {
      maxResults: 100,
      pageToken: pageToken
    };
    if (options.timeMax) {
      listOptions.timeMax = options.timeMax.toISOString();
    }
    if (options.timeMin) {
      listOptions.timeMin = options.timeMin.toISOString();
    }
    events = Calendar.Events.list(calendarId, listOptions);
    eventItems = events.items;
    for (let i=0; i<eventItems.length; i++) {
      eventList.push(eventItems[i]);
    }
    pageToken = events.nextPageToken;
  } while (pageToken);
  return eventList;
}

function generateCalendarId(timeOut) {
  return 'bamboohr' + timeOut.id;
}

function eventFromTimeOut(timeOut) {
  const eventId = generateCalendarId(timeOut);
  return {
    id: eventId,
    summary: timeOut.name + ' - time off',
    description: 'Imported from BambooHR',
    start: {
      date: timeOut.start
    },
    end: { // In Calendar API end date is considered exclusive
      date: formatDate(addRelativeDays(new Date(timeOut.end), 1))
    }
  };
}

function createEvent(calendarId, timeOut) {
  const event = eventFromTimeOut(timeOut);
  return Calendar.Events.insert(event, calendarId);
}

function updateEvent(calendarId, timeOut) {
  const event = eventFromTimeOut(timeOut);
  return Calendar.Events.update(event, calendarId, event.id);
}


function getAccount(accountId) {
  console.log(accountId);
  try {
    const people = People.People.get('people/' + accountId, {
                                     personFields: 'names,emailAddresses'
                                     });
    //console.log('Public Profile: %s', JSON.stringify(people, null, 2));
    return people;
  } catch (e) {
    console.log(e);
    // no person found, could be the calendar system user?
  }
}

function addRelativeDays(reference, numDays) {
  return new Date(reference.getFullYear(), reference.getMonth(), reference.getDate() + numDays);
}

function pad(inputNumber) {
  return (inputNumber < 10 ? '0' : '') + inputNumber;
}

function formatDate(dateObj) {
  return '' + dateObj.getFullYear() + '-' + pad(dateObj.getMonth() + 1) + '-' + pad(dateObj.getDate());
}
