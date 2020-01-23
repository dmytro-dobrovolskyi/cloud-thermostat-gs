//  Setup Instructions (continued):
//
//  For Google Sheets to receive data from the Tilt app
//  deploy script as web app from the Publish menu and set permissions.
//
//  1) Publish > Deploy as web app...
//   
//  2) In the dialog box, set "Who has access to the app:" to "Anyone, even anonymous".
//
//  3) Close Google Scripts tab and return to Google Sheets.
//
var SHEET_NAME = "Sheet1";
var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  return handleResponse(e);
}

function doPost(e){
  return handleResponse(e);
}

function handleResponse(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row = []; 
    // loop through the header columns
    for (i in headers){
      if (headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
        row.push(new Date());
      } else { // else use header name to get data
        row.push(e.parameter[headers[i]]);
      }
    }
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tilt')
    .addItem('View Cloud URL', 'menuItemURL')
    .addItem('Email Cloud URL', 'menuItemEmailURL')
    .addToUi();
  if(SCRIPT_PROP.getProperty("url")==null){
    setup();
  }
  else{
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report").getRange("B5").setValue(SCRIPT_PROP.getProperty("url"));
  }
}

function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", doc.getId());
  
    var html = HtmlService.createHtmlOutputFromFile('setup')
      .setTitle('Cloud Setup Instructions')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

function menuItemURL() {
 
  if(ScriptApp.getService().getUrl()!=null){
    SCRIPT_PROP.setProperty("url", ScriptApp.getService().getUrl());
     SpreadsheetApp.getUi()
      .alert("Copy/Paste the following URL into the Cloud URL field in the Tilt app settings: " + ScriptApp.getService().getUrl());
  }
  else{
    SpreadsheetApp.getUi()
      .alert("Follow setup instructions in sidebar to deploy as web app");
  }
  
}

function menuItemEmailURL(){
  if(ScriptApp.getService().getUrl()!=null){
    SCRIPT_PROP.setProperty("url", ScriptApp.getService().getUrl());  
    MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Tilt Cloud URL', "Copy/Paste the following URL into the Cloud URL field in the Tilt app settings: " + ScriptApp.getService().getUrl());
    SpreadsheetApp.getUi()
      .alert("Email sent to: " + Session.getActiveUser().getEmail());
  }
  else{
    SpreadsheetApp.getUi()
      .alert("Follow setup instructions in sidebar to deploy as web app");
  }
}
