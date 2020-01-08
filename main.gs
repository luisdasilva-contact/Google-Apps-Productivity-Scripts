/* 
 * Writes information on meetings from a list of given emails across a 
      user-specified number of days.
 */
function writeTimeInMeetings() {
  var calendarIDList = getCalendarIDList();
  var numberOfDays = getNumberOfDays();
  var currentTime = new Date();
  var timeMin = new Date();
  timeMin.setDate(currentTime.getDate() - numberOfDays);
  var calendarArgs = {
    timeMin: timeMin.toISOString(),
    timeMax: currentTime.toISOString(),
    showDeleted: false,
    singleEvents: true,
    orderBy: 'startTime'
  };

  if ((calendarIDList === null) || (numberOfDays === null)){
    return;
  } else {
    calendarIDList.forEach(function(calendarID) {
      var userEvents = getMeetings(calendarID, calendarArgs);
      writeToSheet(calendarID, userEvents);
    });
  };
};

/*
 * Given information on a calendar, returns an array of arrays, with each inner 
      array containing up to 50 of the user's valid events. For example, 
      the summary of the 1st event in the second container can be accessed with 
      outerEventsContainer[2][0].summary. This is to avoid write limits via 
      Google's API. NOTE: Use of the Calendar service necessitates enabling
      the Calendar service in Google App Script by going to Resources --> 
      Advanced Google Services, and activating the latest version of the 
      Calendar API. 
  * @param {Array<String>} calendarID The email for which valid events will be
      retrieved.
  * @param {object} calendarArgs An object containing arguments to Google's 
      CalendarApp object.
  * @return {?Array<Event>} A 2D array of Event objects, with
      each inner array containing up to 50 events.
 */
function getMeetings(calendarID, calendarArgs) {
  try {
    var calendarEventsList = Calendar.Events.list(calendarID, calendarArgs);
  } catch (Error) {
   UiFunctions.displayAlert('You don\'t have access to view events for ' + 
      calendarID + ', there is no existing Google Calendar for them, or there' +
          ' was an input error.');
      return;
  };

  var events = calendarEventsList.items;
  var eventContainers = createContainersForCalendarEvents(calendarID, events);
  return eventContainers;
};

/*
 * Creates a 2D array from an existing array, comprised of arrays containing
      50 items each. 
 * @param {Array} An array of any type of object.
 * @return {Array<Array>} An array of arrays, organized into containers of 50.
 */
function createContainers(array){
  const INNER_CONTAINER_SIZE = 50;
  var innerContainer = [];
  var outerContainer = [];

  array.forEach(function(item){
    if (innerContainer.length < INNER_CONTAINER_SIZE){
      innerContainer.push(item);
    } else {
      outerContainer.push(innerContainer);
      innerContainer.push(event);
    };      
  });

  outerContainer.push(innerContainer);
  return outerContainer;
};

/* 
 * Creates a 2D array of Google Calendar Events, given that there are attendees
        in the event, and the event has not been declined. This is for the 
        purpose of avoiding Google's write limits to the Spreadsheet. 
 * @param {string} calendarID the email for the calendar to look up. 
 * @param {Events} events The events to create a 2D array from.  
 */
function createContainersForCalendarEvents(calendarID, events) {
  var arrayOfEvents = [];

  if (events.length > 0){
    events.forEach(function(event) {
      if (((attendeesInMeeting(event)) === true) && 
          ((eventDeclined(calendarID, event)) === false)) {
            arrayOfEvents.push(event);
      };
    });
  };

  return createContainers(arrayOfEvents);
};

/* 
 * Checks if there are attendees in the meeting. 
 * @param {Event} event The event object to check for attendees. 
 * @return {boolean} True if there are attendees, false if there are none. 
 */
function attendeesInMeeting(event) {
  if (event.attendees != undefined) {
      return true;
  } else {
      return false;
  };
};

/* 
 * Checks if an event has been declined.
 * @param {string} calendarID the email for the calendar to look up. 
 * @param {Event} event The event object to check for a decline notice. 
 * @return {boolean} True if there are attendees, false if there are none. 
 */
function eventDeclined(calendarID, event) {
  if (attendeesInMeeting(event) === true) {
      event.attendees.forEach(function(eventAttendee) {
          if (calendarID === eventAttendee.email) {
              if (eventAttendee.responseStatus === 'declined') {
                  return true;
              };
          };
      });
  };
  return false;
};

/* 
 * Returns an array of attendee emails given an event with at least 2 attendees.
 * @param {Event} event The event object to check for attendees.  
 * @return {Array<string>} Array of emails of attendees in the given event. 
 */
function getAttendeeEmails(event) {
    eventAttendeeEmails = [];
    
    event.attendees.forEach(function(attendee) {
        if (((/@resource.calendar.google.com/.test(attendee.email)) != true) &&
              (attendee.email != event.creator.email)) {
                    eventAttendeeEmails.push(attendee.email);
        };
    });
    return eventAttendeeEmails;
};

/* 
 * Returns a 2D array object with the following information on the given event: 
      calendar ID, the event's creator's email, the emails of attendees, 
      the number of attendees, the meeting date, and its duration.  
 * @param {string} calendarID the email for the calendar to look up. 
 * @param {Event} event The event object to look up information on.
 * @return {?Array<String, number>} A 2D array, containing information on 
      the meeting. Contains the calendar ID, the event's creator's email, 
      the emails of attendees, the number of attendees, the meeting date, and 
      its duration. If the user does not have access to look this up, returns 
      null. 
*/
function getItemsForContainer(calendarID, event) {
  var attendeeEmails = [];
  var meetingDate = [];
  var meetingDuration = [];
  var eventEnd = new Date(event.end.dateTime);
  
  try {
    var eventStart = new Date(event.start.dateTime);
  } catch (Error) {
    return null;
  };

  if (event.visibility === 'private') {
      attendeeEmails.push('private event, emails hidden.');
  } else {
      attendeeEmails.push(getAttendeeEmails(event).toString());
  };

  if (eventStart === 'Invalid Date' || eventEnd === 'Invalid Date') {
      meetingDate.push(event.start.date);
      meetingDuration.push('All-day event.');
  } else {
      meetingDate.push(eventStart.toDateString());
      meetingDuration.push(((eventEnd - eventStart) / 3600000).toFixed(2));
  };

  var eventObject = [[calendarID],[event.creator.email], attendeeEmails, 
    [getAttendeeEmails(event).length + 1], meetingDate, meetingDuration];
  return eventObject;
};

/* 
 * Writes a given calendar ID's events to the Spreadsheet in a pre-determined
      format. NOTE: For this function to work, the user must activate the 
      Sheets API by going to Resources --> Advanced Google Resources, and 
      and turning the latest version of the Sheets API on. 
 * @param {string} calendarID * @param {string} calendarID the email for the 
   calendar to look up. 
 * @param {?Array<String, number>} An array of information on the user's
      events, derived from the getMeetings function. 
 */
function writeToSheet(calendarID, events) {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = activeSheet.getSheetByName('Time in Meetings data');
    
    if (dataSheet === null) {
        dataSheet = activeSheet.insertSheet().setName('Time in Meetings ' +
            'data');
        const DATA_SHEET_CELL_RANGE = 'A1:F1';
        const DATA_SHEET_HEADER = [['User Email', 'Meeting Creator',
            'Attendee Emails', 'Number of Attendees', 'Meeting Date',
            'Duration of Meeting (Hours)']];
        dataSheet.getRange(DATA_SHEET_CELL_RANGE).setValues(DATA_SHEET_HEADER);
        dataSheet.getRange(DATA_SHEET_CELL_RANGE).setFontWeight('bold');    
    };

    var valueRange = Sheets.newValueRange();
    var values = [];

    if (events){
      if (!((events.length === 1) && (events[0].length === 0))){
        events.forEach(function(container){
          container.forEach(function(event){
            values.push(getItemsForContainer(calendarID, event));
          });
        });
      } else {
        values.push([calendarID, "-", "-", "-", "-", "-"])
      };
    };

    valueRange.values = values;
    valueRange.majorDimension = 'COLUMNS';

    if (values.length !== 0) {
        dataSheet.getRange(dataSheet.getLastRow() + 1, 1, values.length, 
          values[0].length).setValues(values);
    };
};
