var ui = SpreadsheetApp.getUi();


/*For automatic functions associated with the sheet on launch.*/
function onOpen() {
    buildMenu();
}


/*Creates the custom menu for this app.*/
function buildMenu() {
    ui.createMenu("Time in Meetings")
        .addItem("Find Time for Selected", "getMeetingInput")
        .addToUi();
}


/*Takes the email IDs a user has selected, then asks the user how many days back they want meeting information on. This info is then passed in to the 
getTimeInMeetings function.*/
function getMeetingInput() {
    var calendarIDList = [];
    var userIDs = SpreadsheetApp.getSelection()
                                .getActiveRange()
                                .getValues();

    if (userIDs == "") {
        ui.alert("You have no text selected!");
        return null;
    }

    userIDs.forEach(function(userIDArray) {
        userIDArray.forEach(function(userID) {
        if (userID !== "") {
            calendarIDList.push(userID);
            }
        });
    });

    var userDays = ui.prompt("Number of Days?", "How many days back would you like to see?", ui.ButtonSet.OK_CANCEL);
    
    if (userDays.getSelectedButton() == ui.Button.CANCEL) {
      return null;
    } else {
    while ((userDays.getResponseText() % 1 !== 0) || (userDays.getResponseText() < 1)) {
        userDays = ui.prompt("Number of days?", "Invalid response! Please choose a positive whole number.", ui.ButtonSet.OK_CANCEL);
        if (userDays.getSelectedButton() == ui.Button.CANCEL) {
        return null;
        }
    }

    if (userDays.getSelectedButton() == ui.Button.OK) {
        getTimeInMeetings(userDays.getResponseText(), calendarIDList);
    }
    }
}


/* Retrieves information on meetings from a list of given emails across a specified number of days.*/
function getTimeInMeetings(numberOfDays, calendarIDList) {
    const currentTime = new Date();
    var timeMin = new Date();
    timeMin.setDate(currentTime.getDate() - numberOfDays);

    var calendarArgs = {
        timeMin: timeMin.toISOString(),
        timeMax: currentTime.toISOString(),
        showDeleted: false,
        singleEvents: true,
        orderBy: 'startTime'
    };

    calendarIDList.forEach(function(calendarID) {
        var userEvents = getMeetings(calendarID, calendarArgs);
        writeToSheet(calendarID, userEvents);
    });
};


/* Given the calendarID and calendarArgs, returns an array of arrays, each inner array containing a number 
of the user's valid events. For example, the summary of the 1st event in the second container
can be accessed with outerEventsContainer[2][0].summary. This is to avoid write limits via Google's API.*/
function getMeetings(calendarID, calendarArgs) {
    try {
        var calendarEventsList = Calendar.Events.list(calendarID, calendarArgs);
    } catch (ERROR) {
        return null;
    }

    var events = calendarEventsList.items;
    var eventContainers = createContainers(calendarID, events);
    return eventContainers;
};


/*Takes a calendarID and list of calendar events, and creates a series of array containers from them. */
function createContainers(calendarID, events) {
    const innerContainerSize = 50;
    var innerEventsContainer = [];
    var outerEventsContainer = [];

    if (events.length > 0)
        events.forEach(function(event) {
            if (((attendeesInMeeting(event)) == true) && ((eventDeclined(calendarID, event)) == false)) {
                if (innerEventsContainer.length < innerContainerSize) {
                    innerEventsContainer.push(event);
                } else {
                    outerEventsContainer.push(innerEventsContainer);
                    innerEventsContainer = [];
                    innerEventsContainer.push(event);
                }
            }
        });

    outerEventsContainer.push(innerEventsContainer);
    return outerEventsContainer;
}


/* Checks if there are attendees in the meeting. */
function attendeesInMeeting(event) {
    if (event.attendees != undefined) {
        return true;
    } else {
        return false;
    }
}


/* Returns a true/false value based upon whether the user declined the event. */
function eventDeclined(calendarID, event) {
    if (attendeesInMeeting(event) == true) {
        event.attendees.forEach(function(eventAttendee) {
            if (calendarID == eventAttendee.email) {
                if (eventAttendee.responseStatus == "declined") {
                    return true;
                }
            }
        });
    }
    return false;
}


/* Returns an array of attendee emails given an event with at least 2 attendees. */
function getAttendeeEmails(event) {
    eventAttendeeEmails = [];
    event.attendees.forEach(function(attendee) {
        if (((/@resource.calendar.google.com/.test(attendee.email)) != true) && (attendee.email != event.creator.email)) {
            eventAttendeeEmails.push(attendee.email);
        }
    });
    return eventAttendeeEmails;
}


/* Receives an ID and event, and returns an array object containing desired outputs as arrays. */
function getItemsForContainer(calendarID, event) {
    
    var attendeeEmails = [];
    var meetingDate = [];
    var meetingDuration = [];
    var eventEnd = new Date(event.end.dateTime);
    
    try {
        var eventStart = new Date(event.start.dateTime);
    } catch (ERROR) {
        return null;
    }

    if (event.visibility == "private") {
        attendeeEmails.push("private event, emails hidden.");
    } else {
        attendeeEmails.push(getAttendeeEmails(event).toString());
    }

    if (eventStart == "Invalid Date" || eventEnd == "Invalid Date") {
        meetingDate.push(event.start.date);
        meetingDuration.push("All-day event.");
    } else {
        meetingDate.push(eventStart.toDateString());
        meetingDuration.push(((eventEnd - eventStart) / 3600000).toFixed(2));
    }

    var eventObject = [[calendarID],[event.creator.email], attendeeEmails, [getAttendeeEmails(event).length + 1], 
    meetingDate, meetingDuration];
    return eventObject;
}


/* Writes data to the data Sheet given an array of events.*/
function writeToSheet(calendarID, eventsArgs) {
    const valueInputOption = "RAW";
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheetName = activeSheet.getSheetByName("Time in Meetings data");
    var outerEventContainer = [];
    
    if (dataSheetName == null) {
    dataSheetName = activeSheet.insertSheet().setName("Time in Meetings data");
    var dataSheetCellRange = "A1:F1"
    var dataSheetValues = [["User Email","Meeting Creator","Attendee Emails","Number of Attendees","Meeting Date","Duration of Meeting"]];
    dataSheetName.getRange(dataSheetCellRange).setValues(dataSheetValues);
    dataSheetName.getRange(dataSheetCellRange).setFontWeight("bold");    
    }

    try {
        eventsArgs.forEach(function(outerEvents) {
            var innerEventContainer = [];
            outerEvents.forEach(function(innerEvent) {
                innerEventContainer.push(getItemsForContainer(calendarID, innerEvent));
            });
            outerEventContainer.push(innerEventContainer);
        });
    } catch (TypeError) {
        outerEventContainer.push([
            [calendarID + " - No events found."]
        ]);
    }

    var lastRow = dataSheetName.getRange(dataSheetName.getLastRow() + 1, 1).getA1Notation();
    var rangeName = dataSheetName.getSheetName() + "!" + lastRow;
    var valueRange = Sheets.newValueRange();
    var values = [];

    outerEventContainer.forEach(function(innerEventContainer) {
        innerEventContainer.forEach(function(event) {
            values.push(event);
        });
    });
    
    valueRange.values = values;
    valueRange.majorDimension = "COLUMNS";

    if (values.length !== 0) {
        dataSheetName.getRange(dataSheetName.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
    }
}
