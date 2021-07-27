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

/* exported onOpen, setupTrigger, websiteMonitoring */

// Sheet Names
const SHEET_NAME_TARGET_WEBSITES = '01_Target Websites';
const SHEET_NAME_SPREADSHEETS = '90_Spreadsheets';
const SHEET_NAME_OPTIONS = '99_Options';
const OPTIONS_CONVERT_TO_ARRAY_KEYS = [
  'ALLOWED_RESPONSE_CODES',
  'ERROR_RESPONSE_CODES',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Web Status')
    .addItem('Manual Status Check', 'websiteMonitoring')
    // .addItem('Setup Trigger', 'setupTrigger')
    .addToUi();
}

function websiteMonitoring() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  // Get the list of target websites to monitor
  const targetWebsitesArr = ss
    .getSheetByName(SHEET_NAME_TARGET_WEBSITES)
    .getDataRange()
    .getValues();
  const targetWebsitesHeader = targetWebsitesArr.shift();
  const targetWebsites = targetWebsitesArr.map((row) =>
    targetWebsitesHeader.reduce((o, k, i) => {
      o[k] = row[i];
      return o;
    }, {})
  );
  console.log(JSON.stringify(targetWebsites)); ////////////
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
        value = value.replace(/\s/g, ''); // Remove any whitespaces, should there by any
        obj[key] = value.split(',');
      } else {
        obj[key] = value;
      }
    }
    return obj;
  }, {});
  console.log(JSON.stringify(options)); ////////////
  try {
    // Parse options.ALLOWED_RESPONSE_CODES and options.ERROR_RESPONSE_CODES
    options.ALLOWED_RESPONSE_CODES = parseResponseCodes_(
      options.ALLOWED_RESPONSE_CODES
    );
    if (!options.ALLOWED_RESPONSE_CODES.includes('200')) {
      options.ALLOWED_RESPONSE_CODES.push('200');
    }
    options.ERROR_RESPONSE_CODES = parseResponseCodes_(
      options.ERROR_RESPONSE_CODES
    );
    console.log(JSON.stringify(options)); ////////////
    // Get the actual HTTP response codes
    let errorResponses = targetWebsites.reduce((errors, website) => {
      let responseRecord = {
        websiteName: website['WEBSITE NAME'],
        targetUrl: website['TARGET URL'],
      };
      let checkStart = new Date();
      responseRecord['responseCode'] = String(
        UrlFetchApp.fetch(responseRecord.targetUrl, {
          muteHttpExceptions: true,
        }).getResponseCode()
      );
      let checkEnd = new Date();
      responseRecord['timeStamp'] = Utilities.formatDate(
        checkEnd,
        timeZone,
        'yyyy-MM-dd HH:mm:ss Z'
      );
      responseRecord['responseTime'] = checkEnd - checkStart;
      console.log(JSON.stringify(responseRecord)); ///////////////
      // appendRow [responseRecord.timeStamp, responseRecord.websiteName, responseRecord.targetUrl, responseRecord.responseCode, responseRecord.responseTime] to the recording spreadsheet here
      if (
        options.ALLOWED_RESPONSE_CODES.includes(responseRecord.responseCode)
      ) {
        if (
          options.ERROR_RESPONSE_CODES.includes(responseRecord.responseCode)
        ) {
          errors.push(responseRecord);
        }
      } else {
        errors.push(responseRecord);
      }
      return errors;
    }, []);

    console.log(errorResponses); ////////
  } catch (e) {
    console.log(e.stack); ///////////
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
          let codeReplaced = code.replace(wildcard, `${i}`);
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
