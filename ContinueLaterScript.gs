var nameOfSheet = 'Form Responses 1';
var columnForEmail = 2; // A=1, B=2, C=3, ...

// Created by Dave Ghidiu
// daveghidiu@gmail.com
// Google+: Dave Ghidiu
// Twitter: FringeEdTech
//
// August 12, 2015

/*
The MIT License (MIT)

Copyright (c) 2015 Dave Ghidiu

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/


/**
 * This will always run when the spreadsheet is opened. It will create a menu
 * for the user. It is important that the user runs 'Initialize'. 
 */
function onOpen() {

  try {
  
    SpreadsheetApp.getUi()
      .createMenu('Continue Later')
        .addItem('Initialize', 'initialize')
        .addSeparator()
        .addItem('About', 'about')
        .addToUi();
  
  } catch (err) {
      Logger.log('There was an error in onOpen() at ' + err.lineNumber + ': ' + err.message);  
  }
}


/**
 * Gives the user some information about the script
 */
function about() {
 
  try {
    
    Browser.msgBox('This script was written by Dave Ghidiu, and is made freely available under the MIT License.\\nNo warranties are given.\\n\\nhttp://opensource.org/licenses/MIT\\n\\nDave Ghidiu\\ndaveghidiu@gmail.com\\nGoogle+: Dave Ghidiu\\nTwitter: FringeEdTech');
    
  } catch (err) {
      Logger.log('There was an error in about() at ' + err.lineNumber + ': ' + err.message);    
  }
}


/**
 * This needs to be run once. It will enable the trigger (which automates emails to
 * submitters once they submit).
 */
function initialize() {
  
  try {
    
    // Ask the user for the URL for the form
    setFormURL();
    
    // Delete all current triggers in spreadsheet
    var allTriggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < allTriggers.length; i++) {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    
    // Create a trigger that fires (and calls 'sendEmailOnSubmit()') whenever a form submission is received
    ScriptApp.newTrigger('sendEmailOnSubmit')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();
        
  } catch (err) {
      Logger.log('There was an error in initialize() at ' + err.lineNumber + ': ' + err.message);  
  }
}


/**
 * Runs when 'initialize()' is called. Prompts the user for the URL of the form
 * and then stores it in the Properties Service moniker 'formURL'.
 */
function setFormURL() {
  
  try {
  
    // Get the URL of the form from the user
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Paste in the URL of the form', ui.ButtonSet.OK_CANCEL);
    
    // Store that URL in the script properties
    PropertiesService.getScriptProperties().setProperty('formURL', response.getResponseText());
  
  } catch (err) {
      Logger.log('There was an error in setFormURL() at ' + err.lineNumber + ': ' + err.message);  
  }
}


/**
 * Returns the ID of the form (assuming the user has initialized it)
 *
 * return {string} The URL of the form (as provided by the user)
 */
function getFormURL() {
  
  try {
    
    // Returns the URL of the form (provided earlier by the user)
    return PropertiesService.getScriptProperties().getProperty('formURL');
  
  } catch (err) {
      Logger.log('There was an error in getFormURL() at ' + err.lineNumber + ': ' + err.message);  
  }
}


/**
 * If the trigger is enabled, this will fire whenever a form submission happens.
 * Calls:
 *   getURLs() - which populates the last column with the unique URL for the submission
 *   sendEmail() - which uses the email address (the first question on the form) to send the URL (the last column in the document)
 *
 * @param {event} e A submission from the form
 */
function sendEmailOnSubmit(e) {
  
  try {
    
    getURLs();
    sendEmail();
    
  } catch (err) {
      Logger.log('There was an error in sendEmailOnSubmit(e) at ' + err.lineNumber + ': ' + err.message);  
  }
}


/**
 * This is a fairly reliable way to get the number of items in the form. We want
 * to know what column the response URL is in - and that is always appended after the
 * last item from the form. So this function will help determine what column to 
 * look for.
 *
 * @return {int} The length of the form.getItems (essentially the number of questions on the form)
 */
function getNumberOfItemsOnForm() {
  
  try {  
  
    // Get the form and find the number of items - this will help determine what column the unique URL 
    // for eah submission will be in
    var form = FormApp.openByUrl(getFormURL());
    return form.getItems().length;
    
  } catch (err) {
      Logger.log('There was an error in getNumberOfItemsOnForm() at ' + err.lineNumber + ': ' + err.message);  
  }
}


/**
 * Walks through all submissions and retrieves the unique URL for each submission and
 * puts it in the last column
 */
function getURLs() {
  
  try {
  
    // Get the form and the responses
    var form = FormApp.openByUrl(getFormURL());
    var responses = form.getResponses();    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameOfSheet);
    
    var lastResponse = responses[responses.length-1];
    var url = lastResponse.getEditResponseUrl();
    var numberOfSubmissions = responses.length;
    
    sheet.getRange(numberOfSubmissions + 1, getNumberOfItemsOnForm() + 2).setValue(url);
    
  } catch(err) {
      Logger.log('There was an error in getURLs() at ' + err.lineNumber + ': ' + err.message);
    }
}


/**
 * The portion of the code that sends the email to the user. It is important to note that the first question
 * on the form must be the user's email.
 */
function sendEmail() {

    try {

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameOfSheet);
      var toAddress = sheet.getRange(sheet.getLastRow(), columnForEmail).getValue();
      var subject = "Link to your assessment";
      var url = sheet.getRange(sheet.getLastRow(),getNumberOfItemsOnForm() + 2).getValue();
      var emailBody = 'Dear ' + toAddress + ', ' +
    
        '\n\nYour assessment can be found at: \n\n' + url + '.' + '\n\nPlease do not lose this email! It is the ' +
               
        'only way to resume progress!';

      MailApp.sendEmail(toAddress, subject, emailBody);
  
    } catch (err) {
      Logger.log('There was an error in sendEmail() at ' + err.lineNumber + ': ' + err.message);
    }
}
