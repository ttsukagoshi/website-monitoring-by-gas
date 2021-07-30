// Copyright 2021 Taro Tsukagoshi
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/* exported
deleteTimeBasedTriggers,
extractStatusLogsTriggered,
onOpen,
setupLogExtractionTrigger,
setupStatusCheckTrigger,
websiteMonitoringTriggered
*/

// Sheet Names
const SHEET_NAME_DASHBOARD = '01_Dashboard';
const SHEET_NAME_SPREADSHEETS = '90_Spreadsheets';
const SHEET_NAME_STATUS_LOGS_EXTRACTED = '91_Status Logs Extracted';
const SHEET_NAME_OPTIONS = '99_Options';
// Header Names
const HEADER_NAME_TARGET_URL = 'TARGET URL';
// Range parameters of the list of target websites in SHEET_NAME_DASHBOARD.
const TARGET_WEBSITES_RANGE_POSITION = { row: 5, col: 2 }; // Position of the upper- and left-most cell including the header row
const TARGET_WEBSITES_COL_NUM = 2; // Number of fields names (columns) in the list of target websites
const DASHBOARD_STATUS_COL_NUM = 3; // Number of fields names (columns) in the dashboard status ranges, adjacent to the list of target websites
// Keys in SHEET_NAME_OPTIONS whose value should be converted to arrays
const OPTIONS_CONVERT_TO_ARRAY_KEYS = [
  'ALLOWED_RESPONSE_CODES',
  'ERROR_RESPONSE_CODES',
];
// Document property key(s)
const DP_KEY_SAVED_STATUS = 'savedStatus';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Web Status')
    .addSubMenu(
      ui
        .createMenu('Triggers')
        .addItem('Set Status Check Trigger', 'setupStatusCheckTrigger')
        .addItem('Set Log Extraction Trigger', 'setupLogExtractionTrigger')
        .addSeparator()
        .addItem('Delete Triggers', 'deleteTimeBasedTriggers')
    )
    .addSeparator()
    .addItem('Check Status', 'websiteMonitoring')
    .addItem('Extract Status Logs', 'extractStatusLogs')
    .addToUi();
}

/**
 * Set time-based trigger for website status monitoring.
 */
function setupStatusCheckTrigger() {
  const handlerFunction = 'websiteMonitoringTriggered';
  const frequencyKey = 'TRIGGER_MINUTE_FREQUENCY_STATUS_CHECK';
  const frequencyUnit = 'minute';
  setupTrigger(handlerFunction, frequencyKey, frequencyUnit);
}

/**
 * Set time-based trigger for extracting status logs
 * into the managing spreadsheet.
 */
function setupLogExtractionTrigger() {
  const handlerFunction = 'extractStatusLogsTriggered';
  const frequencyKey = 'TRIGGER_DAYS_FREQUENCY_LOG_EXTRACTION';
  const frequencyUnit = 'day';
  setupTrigger(handlerFunction, frequencyKey, frequencyUnit);
}

/**
 * Set time-based trigger for the input handler function,
 * deleting existing triggers with the same handler function.
 * @param {String} handlerFunction The function name to execute in this trigger.
 * @param {String} frequencyKey Key in the options sheet that refers to the trigger frequency for this handler function.
 * @param {String} frequencyUnit Unit of the value of frequencyKey, i.e., minute, hour, day, or week.
 */
function setupTrigger(handlerFunction, frequencyKey, frequencyUnit) {
  const ui = SpreadsheetApp.getUi();
  const myEmail = Session.getActiveUser().getEmail();
  // Parse options data from spreadsheet
  const optionsArr = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME_OPTIONS)
    .getDataRange()
    .getValues();
  optionsArr.shift();
  const options = optionsArr.reduce((obj, row) => {
    let [key, value] = [row[1], row[2]]; // Assuming that the keys and their options are set in columns B and C, respectively.
    if (key === frequencyKey) {
      obj[key] = value;
    }
    return obj;
  }, {});
  try {
    if (
      !options[frequencyKey] ||
      options[frequencyKey] < 0 ||
      !Number.isInteger(options[frequencyKey])
    ) {
      throw new Error(
        `Invalid ${frequencyKey} value. Check the "${SHEET_NAME_OPTIONS}" worksheet for its value.`
      );
    }
    // Brief descriptions of the handler functions
    const functionDesc = {
      websiteMonitoringTriggered: ' to check website status',
      extractStatusLogsTriggered:
        ' to extract status logs into the managing spreadsheet',
    };
    // Confirm the user if they want to continue with the trigger setup.
    const continueAlert = `Setting up new trigger${
      functionDesc[handlerFunction] ? functionDesc[handlerFunction] : ''
    }. \nThis process will delete existing trigger for this function that was set by ${myEmail}. Are you sure you want to continue?`;
    const continueResponse = ui.alert(
      'Trigger Setup',
      continueAlert,
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (continueResponse !== ui.Button.YES) {
      throw new Error('Trigger setup has been canceled.');
    }
    // Delete existing trigger for this function set by the user.
    ScriptApp.getProjectTriggers().forEach((trigger) => {
      if (trigger.getHandlerFunction() === handlerFunction) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    // Setup a new trigger
    if (frequencyUnit === 'minute') {
      ScriptApp.newTrigger(handlerFunction)
        .timeBased()
        .everyMinutes(options[frequencyKey])
        .create();
    } else if (frequencyUnit === 'hour') {
      ScriptApp.newTrigger(handlerFunction)
        .timeBased()
        .everyHours(options[frequencyKey])
        .create();
    } else if (frequencyUnit === 'day') {
      ScriptApp.newTrigger(handlerFunction)
        .timeBased()
        .everyDays(options[frequencyKey])
        .create();
    } else if (frequencyUnit === 'week') {
      ScriptApp.newTrigger(handlerFunction)
        .timeBased()
        .everyWeeks(options[frequencyKey])
        .create();
    } else {
      throw new Error(`Invalid frequency unit: ${frequencyUnit}`);
    }
    ui.alert(
      `Complete (${handlerFunction})`,
      `Trigger set at ${options[frequencyKey]}-${frequencyUnit} interval.`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(e.stack);
  }
}

/**
 * Delete existing trigger(s).
 */
function deleteTimeBasedTriggers() {
  const ui = SpreadsheetApp.getUi();
  const myEmail = Session.getActiveUser().getEmail();
  try {
    const continueAlert = `Deleting all existing trigger(s) on this spreadsheet/script set by ${myEmail}. Are you sure you want to continue?`;
    const continueResponse = ui.alert(
      'Deleting All Triggers',
      continueAlert,
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (continueResponse !== ui.Button.YES) {
      throw new Error('Trigger deletion has been canceled.');
    }
    // Delete all existing triggers set by the user.
    ScriptApp.getProjectTriggers().forEach((trigger) =>
      ScriptApp.deleteTrigger(trigger)
    );
    ui.alert('Complete', `Trigger(s) deleted.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert(e.stack);
  }
}

/**
 * The core function websiteMonitoring to be executed on timeb-based triggers.
 */
function websiteMonitoringTriggered() {
  const triggered = true;
  websiteMonitoring(triggered);
}

/**
 * The core function for checking the website status.
 * @param {Boolean} triggered Will not show UI.alert popups when true. Defaults to false.
 */
function websiteMonitoring(triggered = false) {
  const myEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  const currentYear = Utilities.formatDate(new Date(), timeZone, 'yyyy');
  const dp = PropertiesService.getDocumentProperties();
  const savedStatus = dp.getProperty(DP_KEY_SAVED_STATUS)
    ? JSON.parse(dp.getProperty(DP_KEY_SAVED_STATUS))
    : {};
  if (!triggered) {
    var ui = SpreadsheetApp.getUi();
  }
  // Get the list of target websites to monitor
  const targetWebsitesSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  const targetWebsitesArr = targetWebsitesSheet
    .getRange(
      TARGET_WEBSITES_RANGE_POSITION.row,
      TARGET_WEBSITES_RANGE_POSITION.col,
      targetWebsitesSheet.getLastRow() - TARGET_WEBSITES_RANGE_POSITION.row + 1,
      TARGET_WEBSITES_COL_NUM
    )
    .getValues();
  const targetWebsitesHeader = targetWebsitesArr.shift();
  const targetWebsites = targetWebsitesArr.map((row) =>
    targetWebsitesHeader.reduce((o, k, i) => {
      if (
        k === HEADER_NAME_TARGET_URL &&
        !savedStatus[Utilities.base64Encode(row[i])]
      ) {
        savedStatus[Utilities.base64Encode(row[i])] = { status: null };
      }
      o[k] = row[i];
      return o;
    }, {})
  );
  // Check and update savedStatus so that it matches with targetWebsites
  const savedStatusUpdated = Object.keys(savedStatus).reduce((obj, key) => {
    if (
      targetWebsites
        .map((site) => Utilities.base64Encode(site[HEADER_NAME_TARGET_URL]))
        .includes(key)
    ) {
      obj[key] = savedStatus[key];
    }
    return obj;
  }, {});
  // Parse options data from spreadsheet
  const optionsArr = ss
    .getSheetByName(SHEET_NAME_OPTIONS)
    .getDataRange()
    .getValues();
  optionsArr.shift();
  const options = optionsArr.reduce((obj, row) => {
    let [key, value] = [row[1], row[2]]; // Assuming that the keys and their options are set in columns B and C, respectively.
    if (key) {
      if (OPTIONS_CONVERT_TO_ARRAY_KEYS.includes(key)) {
        value = String(value).replace(/\s/g, ''); // Remove any whitespaces, should there by any
        obj[key] = value.split(',');
      } else {
        obj[key] = value;
      }
    }
    return obj;
  }, {});
  // Get the list of existing spreadsheets to log the results of status check
  const logSpreadsheetsSheet = ss.getSheetByName(SHEET_NAME_SPREADSHEETS);
  const logSpreadsheetsArr = logSpreadsheetsSheet.getDataRange().getValues();
  const logSpreadsheetsHeader = logSpreadsheetsArr.shift();
  const logSpreadsheets = logSpreadsheetsArr.map((row) =>
    logSpreadsheetsHeader.reduce((o, k, i) => {
      o[k] = row[i];
      return o;
    }, {})
  );
  const logSpreadsheetUrls = logSpreadsheets.filter(
    (row) => row.YEAR == currentYear
  );
  if (!logSpreadsheetUrls.length) {
    // Create a new spreadsheet from template
    // if existing spreadsheet matching currentYear is not available
    const targetFolder = options.DRIVE_FOLDER_ID
      ? DriveApp.getFolderById(options.DRIVE_FOLDER_ID)
      : DriveApp.getRootFolder();
    const templateFile = DriveApp.getFileById(
      SpreadsheetApp.openByUrl(options.TEMPLATE_LOG_SHEET_URL).getId()
      // While the code can be made more simple by asking the user to enter the ID
      // of the template spreadsheet, there is a certain logical usefulness to use
      // the template URL rather than its ID since URLs can be directly referenced
      // from the managing spreadsheet.
    );
    const newFileName = options.LOG_FILE_NAME
      ? options.LOG_FILE_NAME.replace(/{{year}}/g, currentYear)
      : templateFile.getName() + currentYear;
    const newFileUrl = templateFile
      .makeCopy(newFileName, targetFolder)
      .getUrl();
    // Append row to logSpreadsheetsSheet
    logSpreadsheetsSheet.appendRow([currentYear, newFileName, newFileUrl]);
    logSpreadsheetUrls.push({
      YEAR: currentYear,
      NAME: newFileName,
      URL: newFileUrl,
    });
  }
  const logSheet = SpreadsheetApp.openByUrl(
    logSpreadsheetUrls[0].URL
  ).getSheets()[0]; // Assuming that the logs be entered on the left-most worksheet of the log spreadsheet
  try {
    // Replace wildcards in options.ALLOWED_RESPONSE_CODES and options.ERROR_RESPONSE_CODES to actual codes
    options.ALLOWED_RESPONSE_CODES = parseResponseCodes_(
      options.ALLOWED_RESPONSE_CODES
    );
    if (!options.ALLOWED_RESPONSE_CODES.includes('200')) {
      options.ALLOWED_RESPONSE_CODES.push('200');
    }
    options.ERROR_RESPONSE_CODES = parseResponseCodes_(
      options.ERROR_RESPONSE_CODES
    );
    console.log(JSON.stringify(options)); //////////////////
    // Get the actual HTTP response codes
    let dashboardStatus = []; // Array to record on the dashboard worksheet
    let statusChange = targetWebsites.reduce(
      (changes, website) => {
        let responseRecord = {
          websiteName: website['WEBSITE NAME'],
          targetUrl: website[HEADER_NAME_TARGET_URL],
          targetUrlEncoded: Utilities.base64Encode(
            website[HEADER_NAME_TARGET_URL]
          ),
          status:
            savedStatusUpdated[
              Utilities.base64Encode(website[HEADER_NAME_TARGET_URL])
            ].status,
        };
        let checkStart = new Date();
        responseRecord['responseCode'] = String(
          UrlFetchApp.fetch(responseRecord.targetUrl, {
            muteHttpExceptions: true,
          }).getResponseCode()
        );
        let checkEnd = new Date();
        responseRecord['timeStamp'] = standardFormatDate_(checkEnd, timeZone);
        responseRecord['responseTime'] = checkEnd - checkStart; // Time in milliseconds
        // Determine if the returned HTTP response code means UP or DOWN
        if (
          options.ALLOWED_RESPONSE_CODES.includes(responseRecord.responseCode)
        ) {
          if (
            options.ERROR_RESPONSE_CODES.includes(
              responseRecord.responseCode
            ) &&
            (!responseRecord.status || responseRecord.status === 'UP')
          ) {
            responseRecord.status = 'DOWN';
            changes.newErrors.push(responseRecord);
          } else if (responseRecord.status === 'DOWN') {
            responseRecord.status = 'UP';
            changes.resolved.push(responseRecord);
          } else {
            responseRecord.status = 'UP';
          }
        } else if (!responseRecord.status || responseRecord.status === 'UP') {
          responseRecord.status = 'DOWN';
          changes.newErrors.push(responseRecord);
        }
        // Log result to the log spreadsheet
        logSheet.appendRow([
          responseRecord.timeStamp,
          responseRecord.websiteName,
          responseRecord.targetUrl,
          responseRecord.responseCode,
          responseRecord.responseTime,
          responseRecord.status,
        ]);
        // Updates to the dashboard worksheet
        dashboardStatus.push([
          responseRecord.status,
          responseRecord.responseCode,
          responseRecord.timeStamp,
        ]);
        // Update savedStatusUpdated
        savedStatusUpdated[responseRecord.targetUrlEncoded] = responseRecord;
        return changes;
      },
      { newErrors: [], resolved: [] }
    );
    // Update the dashboard status
    targetWebsitesSheet
      .getRange(
        TARGET_WEBSITES_RANGE_POSITION.row + 1,
        TARGET_WEBSITES_RANGE_POSITION.col + TARGET_WEBSITES_COL_NUM,
        dashboardStatus.length,
        DASHBOARD_STATUS_COL_NUM
      )
      .setValues(dashboardStatus);
    // Save the updated savedStatusUpdated in the document properties
    dp.setProperty(DP_KEY_SAVED_STATUS, JSON.stringify(savedStatusUpdated));
    // Update dashboard info
    if (statusChange.newErrors.length > 0) {
      MailApp.sendEmail(
        myEmail,
        '[Website Status] Alert: Site DOWN',
        `The following website(s) are DOWN:\n\n${statusChange.newErrors
          .map(
            (errorResponse) =>
              `Site Name: ${errorResponse.websiteName}\nURL: ${errorResponse.targetUrl}\nResponse Code: ${errorResponse.responseCode}\nResponse Time: ${errorResponse.responseTime}\n`
          )
          .join(
            '\n'
          )}\n\n-----\nThis notice is managed by the following spreadsheet:\n${ss.getUrl()}`
      );
    }
    if (statusChange.resolved.length > 0) {
      MailApp.sendEmail(
        myEmail,
        '[Website Status] Notice: Site UP (Resolved)',
        `The following website(s) that were DOWN are now UP:\n\n${statusChange.resolved
          .map((resolvedResponse) => {
            `Site Name: ${resolvedResponse.websiteName}\nURL: ${resolvedResponse.targetUrl}\nResponse Code: ${resolvedResponse.responseCode}\nResponse Time: ${resolvedResponse.responseTime}\n`;
          })
          .join(
            '\n'
          )}\n\n-----\nThis notice is managed by the following spreadsheet:\n${ss.getUrl()}`
      );
    }
    // Log message
    let completeMessage = 'Website status check is complete.';
    if (statusChange.newErrors.length > 0 || statusChange.resolved.length > 0) {
      completeMessage += `\nChanges to website status have been emailed to ${myEmail}`;
      // Set a one-time trigger to update extracted status logs on the managing spreadsheet
      // that will fire 30 secs later.
      ScriptApp.newTrigger('extractStatusLogsTriggered')
        .timeBased()
        .after(30 * 1000)
        .create();
    }
    logSheet.appendRow([
      standardFormatDate_(new Date(), timeZone),
      '[COMPLETE]',
      completeMessage,
      0,
      0,
      'NA',
    ]);
    if (!triggered) {
      // Show UI message, if triggered = false, i.e., this function is executed manually.
      ui.alert(
        '[Website Status] Complete: Status Check',
        completeMessage,
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    console.error(e.stack);
    logSheet.appendRow([
      standardFormatDate_(new Date(), timeZone),
      '[ERROR]',
      e.stack,
      0,
      0,
      'NA',
    ]);
    MailApp.sendEmail(
      myEmail,
      '[Website Status] Error: Status Check',
      `${
        e.stack
      }\n\n-----\nThis notice is managed by the following spreadsheet:\n${ss.getUrl()}`
    );
    if (!triggered) {
      ui.alert('ERROR', e.stack, ui.ButtonSet.OK);
    }
  }
}

/**
 * extractStatusLogs that will be executed by time-based triggers
 */
function extractStatusLogsTriggered() {
  const triggered = true;
  extractStatusLogs(triggered);
}

/**
 * Extract status logs from the log spreadsheets
 * and copy them into a worksheet in the managing spreadsheet
 * for reporting purposes.
 * @param {Boolean} triggered Will not show UI.alert popups when true. Defaults to false.
 */
function extractStatusLogs(triggered = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  if (!triggered) {
    var ui = SpreadsheetApp.getUi();
  }
  try {
    // Clear the worksheet to enter new status logs
    const extractedLogsSheet = ss.getSheetByName(
      SHEET_NAME_STATUS_LOGS_EXTRACTED
    );
    extractedLogsSheet.getDataRange().clearContent();
    // Get the list of target websites to monitor
    const targetWebsitesSheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
    const targetWebsitesArr = targetWebsitesSheet
      .getRange(
        TARGET_WEBSITES_RANGE_POSITION.row,
        TARGET_WEBSITES_RANGE_POSITION.col,
        targetWebsitesSheet.getLastRow() -
          TARGET_WEBSITES_RANGE_POSITION.row +
          1,
        TARGET_WEBSITES_COL_NUM
      )
      .getValues();
    const targetWebsitesHeader = targetWebsitesArr.shift();
    const targetWebsiteUrls = targetWebsitesArr.map((row) => {
      let urlIndex = targetWebsitesHeader.indexOf(HEADER_NAME_TARGET_URL);
      if (urlIndex < 0) {
        throw new Error(
          `"${HEADER_NAME_TARGET_URL}" is not found. Check the "${SHEET_NAME_DASHBOARD}" worksheet for the header name of the target websites\' URL`
        );
      }
      return row[urlIndex];
    });
    // Parse options data from spreadsheet
    const optionsArr = ss
      .getSheetByName(SHEET_NAME_OPTIONS)
      .getDataRange()
      .getValues();
    optionsArr.shift();
    const options = optionsArr.reduce((obj, row) => {
      let [key, value] = [row[1], row[2]]; // Assuming that the keys and their options are set in columns B and C, respectively.
      if (key === 'EXTRACT_STATUS_LOGS_DAYS') {
        obj[key] = value || 365;
      }
      return obj;
    }, {});
    // Get the start date to obtain status logs
    const today = new Date();
    const startLog = new Date(
      new Date().setDate(today.getDate() - options.EXTRACT_STATUS_LOGS_DAYS)
    );
    // Extract status logs from the list of log spreadsheets
    const logSpreadsheetsArr = ss
      .getSheetByName(SHEET_NAME_SPREADSHEETS)
      .getDataRange()
      .getValues();
    const logSpreadsheetsHeader = logSpreadsheetsArr.shift();
    // Array to note the headers of the yearly logs;
    // if the items in the headers do not match each other,
    // an error will be returned later on.
    let headersArr = [];
    // The actual 2-d array to be extracted and saved on the managing spreadsheet
    let statusLogs = logSpreadsheetsArr
      .reduce((filteredList, row) => {
        let rowObj = logSpreadsheetsHeader.reduce((o, k, i) => {
          o[k] = row[i];
          return o;
        }, {});
        if (
          rowObj.YEAR >=
          parseInt(Utilities.formatDate(startLog, timeZone, 'yyyy'))
        ) {
          filteredList.push(rowObj);
        }
        return filteredList;
      }, [])
      .map((yearLog) => {
        let logsArr = SpreadsheetApp.openByUrl(yearLog.URL)
          .getSheets()[0]
          .getDataRange()
          .getValues();
        let logsHeader = logsArr.shift();
        headersArr.push(logsHeader);
        return logsArr.reduce((logs, log) => {
          let logObj = logsHeader.reduce((o, k, i) => {
            o[k] = log[i];
            return o;
          }, {});
          if (
            logObj.TIMESTAMP >= startLog &&
            targetWebsiteUrls.includes(logObj.URL)
          ) {
            logs.push(log);
          }
          return logs;
        }, []);
      })
      .flat();
    // Check the log headers to see if they match each other
    let controlHeader = headersArr[0];
    headersArr.forEach((headers) => {
      headers.forEach((header, i) => {
        if (header !== controlHeader[i]) {
          throw new Error(
            'There seems to be an inconsistency in the header row between the status log files. Edit the header(s) so that they match and retry.'
          );
        }
      });
    });
    // Copy into the managing spreadsheet
    statusLogs = [controlHeader].concat(statusLogs);
    extractedLogsSheet
      .getRange(1, 1, statusLogs.length, statusLogs[0].length)
      .setValues(statusLogs);
    if (!triggered) {
      ui.alert('Complete', 'Status Log Extraction', ui.ButtonSet.OK);
    }
  } catch (e) {
    console.error(e.stack);
    if (!triggered) {
      ui.alert(`[Website Status] Error: ${e.stack}`);
    }
  }
}

/**
 * Parse a given array of HTTP reponse codes (in strings) into actual codes.
 * For example, ["201", "30x"] will be converted into the following array of codes:
 * ["201", "300", "301", "302", "303", "304", "305", "306", "307", "308", "309"]
 * @param {Array} codes An array of HTTP response codes in strings. Wildcards can be used to replace digits.
 * @param {String} wildcard Placeholder value to denote the digits from 0 to 9. Defaults to "x".
 * @returns {Array} An array of replaced codes.
 */
function parseResponseCodes_(codes, wildcard = 'x') {
  return codes
    .map((code) => {
      if (code.length !== 3 && code.length !== 0) {
        throw new Error(
          `Invalid response code "${code}" at ALLOWED_RESPONSE_CODES`
        );
      }
      let parsedCodes = [];
      let remainingWildcard = 0;
      if (code.includes(wildcard)) {
        for (let i = 0; i < 10; i++) {
          let codeReplaced = code.replace(wildcard, i);
          if (codeReplaced.includes(wildcard)) {
            remainingWildcard += 1;
          }
          parsedCodes.push(codeReplaced);
        }
      } else {
        parsedCodes.push(code);
      }
      if (remainingWildcard > 0) {
        parsedCodes = parseResponseCodes_(parsedCodes, wildcard);
      }
      return parsedCodes.flat();
    })
    .flat();
}

/**
 * Standardized date format for this script.
 * @param {Date} dateObj Date object to format.
 * @param {String} timeZone Time zone. Defaults to the active spreadsheet's time zone.
 * @returns {String} The formatted date.
 */
function standardFormatDate_(
  dateObj,
  timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
) {
  return Utilities.formatDate(dateObj, timeZone, 'yyyy-MM-dd HH:mm:ss');
}
