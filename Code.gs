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
var HEADER_ROW = ['Time', 'SG', 'Temp', 'Color', 'Comment'];
var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service
var CONFIG_SHEET = 'Thermostat Configuration';
var DATE_FORMAT = 'HH:mm | MM-dd-yyyy';

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e) {
    return handleResponse(e);
}

function doPost(e) {
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
        if (e.parameter.Beer === undefined) {
            return ContentService
                .createTextOutput(JSON.stringify({
                    'result': 'success',
                    'message': 'All good. Now paste this URL to your Tilt app or Tilt PI interface'
                }))
                .setMimeType(ContentService.MimeType.JSON);
        }
        SCRIPT_PROP.setProperty('beer', e.parameter.Beer);

        // next set where we write the data - you could write to multiple/alternate destinations
        var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key'));
        var readings = doc.getSheetByName(e.parameter.Beer);

        if (readings == null) {
            readings = createNewBrewSheet(doc, e);
        }

        // we'll assume header is in row 1 but you can override with header_row in GET/POST data
        var headers = readings.getRange(1, 1, 1, readings.getLastColumn()).getValues()[0];
        var nextRow = readings.getLastRow() + 1; // get next row
        var row = [];
        // loop through the header columns
        for (i in headers) {
            if (headers[i] === 'Time') { // special case if you include a 'Timestamp' column
                row.push(now());
            } else { // else use header name to get data
                row.push(e.parameter[headers[i]]);
            }
        }
      // more efficient to set values as [][] array than individually
        readings.getRange(nextRow, 1, 1, row.length).setValues([row]);
        // return json success results

        regulateTemperature();

        return ContentService
            .createTextOutput(JSON.stringify({'result': 'success', 'row': nextRow}))
            .setMimeType(ContentService.MimeType.JSON);
    } catch (e) {
        // if error return this
        return ContentService
            .createTextOutput(JSON.stringify({'result': 'error', 'error': e}))
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
    if (SCRIPT_PROP.getProperty('url') == null) {
        setup();
    } else {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Report').getRange('B5').setValue(SCRIPT_PROP.getProperty('url'));
    }
}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty('key', doc.getId());

    var html = HtmlService.createHtmlOutputFromFile('setup')
        .setTitle('Cloud Setup Instructions')
        .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showSidebar(html);
}

function menuItemURL() {

    if (ScriptApp.getService().getUrl() != null) {
        SCRIPT_PROP.setProperty('url', ScriptApp.getService().getUrl());
        SpreadsheetApp.getUi()
            .alert('Copy/Paste the following URL into the Cloud URL field in the Tilt app settings: ' + ScriptApp.getService().getUrl());
    } else {
        SpreadsheetApp.getUi()
            .alert('Follow setup instructions in sidebar to deploy as web app');
    }

}

function menuItemEmailURL() {
    if (ScriptApp.getService().getUrl() != null) {
        SCRIPT_PROP.setProperty('url', ScriptApp.getService().getUrl());
        MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Tilt Cloud URL', 'Copy/Paste the following URL into the Cloud URL field in the Tilt app settings: ' + ScriptApp.getService().getUrl());
        SpreadsheetApp.getUi()
            .alert('Email sent to: ' + Session.getActiveUser().getEmail());
    } else {
        SpreadsheetApp.getUi()
            .alert('Follow setup instructions in sidebar to deploy as web app');
    }
}

function createNewBrewSheet(doc, e) {
    var sheet = doc.insertSheet()
        .setName(e.parameter.Beer)
        .appendRow(HEADER_ROW);

    sheet.getRange(1, 1, 1, HEADER_ROW.length).setFontWeight('bold');

    return sheet;
}

function now() {
    return Utilities.formatDate(new Date(), CalendarApp.getDefaultCalendar().getTimeZone(), DATE_FORMAT);
}

function turnOn() {
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key'));
    var configSheet = doc.getSheetByName(CONFIG_SHEET);
    var config = {
        url: getCellValueByName(configSheet, 'ON URL', 2),
        method: getCellValueByName(configSheet, 'OFF Method', 2),
        body: getCellValueByName(configSheet, 'ON Body', 2),
        headers: getCellValueByName(configSheet, 'Headers', 2)
    };

    doHttpCall(config);
}

function turnOff() {
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key'));
    var configSheet = doc.getSheetByName(CONFIG_SHEET);
    var config = {
        url: getCellValueByName(configSheet, 'OFF URL', 2),
        method: getCellValueByName(configSheet, 'OFF Method', 2),
        body: getCellValueByName(configSheet, 'OFF Body', 2),
        headers: getCellValueByName(configSheet, 'Headers', 2)
    };

    doHttpCall(config);
}

function doHttpCall(config) {
    var options = {
        'method': config.method,
        'contentType': 'application/json'
    };
    if (config.method !== 'GET') {
        options.payload = config.body;
    }
    if (config.headers) {
        options.headers = JSON.stringify(config.headers);
    }
    console.log('HTTP Response: %s', UrlFetchApp.fetch(config.url, options));
}

function regulateTemperature() {
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key'));
    var configSheet = doc.getSheetByName(CONFIG_SHEET);
    var readingsSheet = doc.getSheetByName(SCRIPT_PROP.getProperty('beer'));

    var config = {
        readingsSheetName: getCellValueByName(configSheet, 'Readings sheet name', 2),
        isHeating: function () {
            return getCellValueByName(configSheet, 'Mode', 2) === 'Heating';
        },
        isCooling: function () {
            return !this.isHeating();
        },
        temperature: getCellValueByName(configSheet, 'Desired Temperature', 2),
        turnOffIfDisaster: getCellValueByName(configSheet, 'Turn off if disaster', 2) === 'Yes'
    };
    var currentTemperature = getCellValueByName(readingsSheet, 'Temp', readingsSheet.getLastRow());

    console.log("Current temperature is {}", currentTemperature);

    if (isHeatingRequired(config, currentTemperature)) {
        turnOn();
        console.log("It's too cool but no worries it's gonna be all right. Turning on the heater!");
    } else if (isCoolingRequired(config, currentTemperature)) {
        turnOn();
        console.log("It's too hot but no worries it's gonna be all right. Turning on the cooler!");
    } else if (shouldTurnOff(config, currentTemperature)) {
        turnOff();
        console.log("Looks like we did the job well. Turning off the heater/cooler");
    } else {
        console.log("Looks like you're all set!");
    }
}

function isCoolingRequired(config, currentTemperature) {
    return config.isCooling() && isTooHot(config, currentTemperature);
}

function isHeatingRequired(config, currentTemperature) {
    return config.isHeating() && isTooCold(config, currentTemperature);
}

function shouldTurnOff(config, currentTemperature) {
    return isHeatedEnough(config, currentTemperature) || isCoolEnough(config, currentTemperature);
}

function isCoolEnough(config, currentTemperature) {
    return config.isCooling() && currentTemperature <= config.temperature;
}

function isHeatedEnough(config, currentTemperature) {
    return config.isHeating() && currentTemperature >= config.temperature;
}

function isTooCold(config, currentTemperature) {
    return currentTemperature < config.temperature;
}

function isTooHot(config, currentTemperature) {
    return currentTemperature > config.temperature;
}

function getCellValueByName(sheet, colName, row) {
    var data = sheet.getDataRange().getValues();
    var col = data[0].indexOf(colName);

    if (col !== -1) {
        return data[row - 1][col];
    }
}

function checkResponsiveness() {
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty('key'));
    var configSheet = doc.getSheetByName(CONFIG_SHEET);
    var readingsSheet = doc.getSheetByName(SCRIPT_PROP.getProperty('beer'));
    var lastReading = getCellValueByName(readingsSheet, 'Time', readingsSheet.getLastRow());

    var currentTime = Moment.moment(now(), 'HH:mm');
    var lastReadingTime = Moment.moment(lastReading, 'HH:mm');

    if (lastReadingTime.isAfter(currentTime)) {
        lastReadingTime.subtract(1, 'days');
    }
    if (currentTime.diff(lastReadingTime, 'minutes') > 30) {
        console.error('Disastrous situation!');

        if (getCellValueByName(configSheet, 'Turn off if disaster', 2) === 'Yes') {
            console.log('Turning OFF');
            turnOff();
        } else {
            console.log('Keeping ON according to the configuration given');
        }
    }
}
