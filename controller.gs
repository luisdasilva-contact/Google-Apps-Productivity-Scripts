/**
 * Takes the user's input, and parses it as a list of emails.
 * @return {?Array<string>} Returns the user's selection as a list of strings, 
        separated by cell. Null if no text is selected.
*/
function getCalendarIDList() {
  var calendarIDList = [];
  var userIDs = SpreadsheetApp.getSelection()
                              .getActiveRange()
                              .getValues();

  if (userIDs === '') {
    UiFunctions.displayAlert('You have no text selected!');
    return null;
  };

  userIDs.forEach(function(userIDArray) {
    userIDArray.forEach(function(userID) {
    if (userID !== '') {
      calendarIDList.push(userID);
      };
    });
  });
  return calendarIDList;
};

/**
 * Retrieves the number of days back the user states they'd like to look back.
 * @return {?number} The number of days the user would like to look back. Null
        if canceled. 
 */
function getNumberOfDays(){
  var userDays = UiFunctions.displayPrompt('How many days back would you ' + 
      'like to see?');
  
  if (userDays) {
      while ((userDays % 1 !== 0) || (userDays < 1)) {
          userDays = UiFunctions.displayPrompt('Invalid response! Please '  + 
              'choose a positive whole number.');
      };
    return userDays;
    } else {
      return null;
    }
};
