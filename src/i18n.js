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

/* exported LocalizedMessage */

const MESSAGES = {
  en: {
    menuTitle: 'Web Status',
    menuTriggers: 'Triggers',
    menuSetStatusCheckTrigger: 'Set Status Check Trigger',
    menuSetLogExtractionTrigger: 'Set Log Extraction Trigger',
    menuDeleteTriggers: 'Delete Triggers',
    menuCheckStatus: 'Check Status',
    menuExtractStatusLogs: 'Extract Status Logs',
    errorInvalidFrequencyValue:
      'Invalid {{frequencyKey}} value. Check the "{{SHEET_NAME_OPTIONS}}" worksheet for its value.',
    functionDescWebsiteMonitoringTriggered: ' to check website status',
    functionDescExtractStatusLogsTriggered:
      ' to extract status logs into the managing spreadsheet',
    alertMessageContinueTriggerSetup:
      'Setting up new trigger{{functionDesc}}. \nThis process will delete existing trigger for this function that was set by {{myEmail}}. Are you sure you want to continue?',
    alertTitleContinueTriggerSetup: 'Trigger Setup',
    errorTriggerSetupCanceled: 'Trigger setup has been canceled.',
    errorInvalidFrequencyUnit: 'Invalid frequency unit: {{frequencyUnit}}',
    alertTitleCompleteTriggerSetup: 'Complete ({{handlerFunction}})',
    alertMessageCompleteTriggerSetup:
      'Trigger set at {{frequency}}-{{frequencyUnit}} interval.',
  },
  ja: {
    menuTitle: 'サイト公開ステータス',
    menuTriggers: 'トリガー',
    menuSetStatusCheckTrigger: 'トリガー設定（ステータス確認）',
    menuSetLogExtractionTrigger: 'トリガー設定（ログ抽出）',
    menuDeleteTriggers: 'トリガー削除',
    menuCheckStatus: 'ステータス確認',
    menuExtractStatusLogs: 'ステータスログを抽出',
    errorInvalidFrequencyValue:
      '{{frequencyKey}}の値が無効です。シート「{{SHEET_NAME_OPTIONS}}」で値を確認してください。',
    functionDescWebsiteMonitoringTriggered:
      'ウェブサイトのステータス確認のため、',
    functionDescExtractStatusLogsTriggered:
      'この管理シートにステータスログを抽出してくるため、',
  },
};

class LocalizedMessage {
  constructor(userLocale) {
    this.DEFAULT_LOCALE = 'en';
    this.locale = MESSAGES[userLocale] ? userLocale : this.DEFAULT_LOCALE;
    this.messageList = MESSAGES[this.locale];
    Object.keys(MESSAGES[this.DEFAULT_LOCALE]).forEach((key) => {
      if (!this.messageList[key]) {
        this.messageList[key] = MESSAGES[this.DEFAULT_LOCALE][key];
      }
    });
  }
  /**
   * Replace placeholder values in the designated text. String.prototype.replace() is executed using regular expressions with the 'global' flag on.
   * @param {String} text
   * @param {Array} placeholderValues Array of objects containing a placeholder string expressed in regular expression and its corresponding value.
   * @returns {String} The replaced text.
   */
  replacePlaceholders_(text, placeholderValues = []) {
    let replacedText = placeholderValues.reduce(
      (acc, cur) => acc.replace(new RegExp(cur.regexp, 'g'), cur.value),
      text
    );
    return replacedText;
  }
  /**
   * Replace placeholder string in this.messageList.errorInvalidFrequencyValue
   * @param {String} frequencyKey
   * @param {String} sheetNameOptions Name of options worksheet
   * @returns {String} The replaced text.
   */
  replaceErrorInvalidFrequencyValue(frequencyKey, sheetNameOptions) {
    let text = this.messageList.errorInvalidFrequencyValue;
    let placeholderValues = [
      {
        regexp: '{{frequencyKey}}',
        value: frequencyKey,
      },
      {
        regexp: '{{sheetNameOptions}}',
        value: sheetNameOptions,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertMessageContinueTriggerSetup
   * @param {String} functionDesc
   * @param {String} myEmail
   * @returns {String} The replaced text.
   */
  replaceAlertMessageContinueTriggerSetup(functionDesc, myEmail) {
    let text = this.messageList.alertMessageContinueTriggerSetup;
    let placeholderValues = [
      {
        regexp: '{{functionDesc}}',
        value: functionDesc,
      },
      {
        regexp: '{{myEmail}}',
        value: myEmail,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.errorInvalidFrequencyUnit
   * @param {String} frequencyUnit
   * @returns {String} The replaced text.
   */
  replaceErrorInvalidFrequencyUnit(frequencyUnit) {
    let text = this.messageList.errorInvalidFrequencyUnit;
    let placeholderValues = [
      {
        regexp: '{{frequencyUnit}}',
        value: frequencyUnit,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertTitleCompleteTriggerSetup
   * @param {String} handlerFunction
   * @returns {String} The replaced text.
   */
  replaceAlertTitleCompleteTriggerSetup(handlerFunction) {
    let text = this.messageList.alertTitleCompleteTriggerSetup;
    let placeholderValues = [
      {
        regexp: '{{handlerFunction}}',
        value: handlerFunction,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertMessageCompleteTriggerSetup
   * @param {String} frequency
   * @param {String} frequencyUnit
   * @returns {String} The replaced text.
   */
  replaceAlertMessageCompleteTriggerSetup(frequency, frequencyUnit) {
    let text = this.messageList.alertMessageCompleteTriggerSetup;
    let placeholderValues = [
      {
        regexp: '{{frequency}}',
        value: frequency,
      },
      {
        regexp: '{{frequencyUnit}}',
        value: frequencyUnit,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}
