/**
 * Module used to display content via Google Apps Script's built-in UI features. 
      Contains the following functions: 
 * @return {function} displayAlert Displays a window with text to the user.
 * @return {function} displayPrompt Displays a text field to the user, allowing 
      them to submit text, or cancel.
 * @return {function} displayYesNoChoice Displays a window with text to the 
      user, allowing them to select "YES", "NO", or "CANCEL".
 * @return {function} displayHTMLAlert Displays a window with text to the user, 
      taking HTML content.
 */ 
var UiFunctions = (function(){
  var Ui = SpreadsheetApp.getUi();
  
  /**
   * Displays a window with text to the user.
   * @param {string} alertText The text to display to the user.
   */
  function displayAlert(alertText){
      Ui.alert(alertText);
  };
  
  /**
   * Displays a text field to the user, allowing them to submit text, or cancel.
   * @param {string} promptText The text to display to the user above the text 
        field.
   * @return {?string} The text the user has entered if they clicked "OK". Null 
        if "CANCEL" is clicked.
   */
  function displayPrompt(promptText){
      var prompt = Ui.prompt(promptText);
      
      if (prompt.getSelectedButton() === Ui.Button.OK){
          return prompt.getResponseText();
      } else {
          return null;
      };
  };
  
  /**
   * Displays a window with text to the user, allowing them to select "YES", 
        "NO", or "CANCEL".
   * @param {string} promptText The text to display to the user above the "YES", 
        "NO", and "CANCEL" buttons.
   * @return {?string} The user's choice. "YES" if they clicked "YES", "NO" if 
        they clicked "NO", null if they clicked "CANCEL".
    */
  function displayYesNoChoice(promptText){
      var prompt = Ui.alert(promptText, Ui.ButtonSet.YES_NO_CANCEL);
      
      if (prompt === Ui.Button.YES){
        return 'YES';
      } else if (prompt === Ui.Button.NO){
        return 'NO';
      } else {
        return null;
      };
  };
  
  /**
   * Displays a window with text to the user, taking HTML content.
   * @param {string} HTML content to display in the window.
   * @param {string} title The text to display as the window's title.
   */
  function displayHTMLAlert(HtmlOutput, title){
      Ui.showModelessDialog(HtmlOutput, title);            
  };

  return {
      getUi: Ui,
      displayPrompt: displayPrompt,
      displayAlert: displayAlert,
      displayYesNoChoice: displayYesNoChoice,
      displayHTMLAlert: displayHTMLAlert
};
})();

/**
 * HTML text to display to the user, showing them how to navigate the program. 
 * @return {string} HTML to display in the window the user will see upon clicking 
      the Help button.
 */
function helpText(){
return '<html><head><link rel=\'stylesheet\'' +
      'href=\'https://ssl.gstatic.com/docs/script/css/add-ons1.css\'></head>' +
      'For a list of Gmail accounts for which the user has access to their ' +
      'Google calendars, this program will return info on their meetings for ' +
      'a user-entered number of days past. The info will only return events ' +
      'that were not canceled or only had one attendee, and will include ' +
      'such as the meeting\'s creator, emails of attendees, number of ' +
      'attendees, the date of the meeting, and the meeting\'s duration in ' +
      'hours.<br>' +
      'To use the program, enter one email per cell in the spreadsheet, and ' +
      'select each of the emails for which you\'d like to find Calendar info ' +
      'on. With the cells selected, go to Time in Meetings -> Find Time for ' +
      'Selected. A window will pop up asking the number of days back you\'d ' +
      'like to search. Enter a number, and the program will create a new ' +
      'sheet titled "Time in Meetings data", showing all data for the ' +
      'calendars you have access to.'
      '</font></body></html>';
};

/**
 * Function placed in UI to activate method to display HTML content in a window. 
 */
function helpButton(){
  var HtmlOutput = HtmlService.createHtmlOutput(helpText())
                              .setWidth(700)
                              .setHeight(300);
  UiFunctions.displayHTMLAlert(HtmlOutput, "Help");
};

/**
 * Function to automatically build menu upon opening the document.
 * @param {Event} onOpen event containing context regarding the document upon 
      opening.
 */
function onOpen(e){
  UiFunctions.getUi
             .createMenu('Time in Meetings')
             .addItem('Find Time for Selected', 'writeTimeInMeetings')
             .addSeparator()
             .addItem('Help', 'helpButton')
             .addToUi(); 
};
