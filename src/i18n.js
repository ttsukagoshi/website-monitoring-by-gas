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

// See https://www.scriptable-assets.page/gas-solutions/website-monitoring-by-gas/ for latest updates.

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
      'Invalid {{frequencyKey}} value. Check the "{{sheetNameOptions}}" worksheet for its value.',
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
    alertTitleContinueTriggerDelete: 'Deleting All Triggers',
    alertMessageContinueTriggerDelete:
      'Deleting all existing trigger(s) on this spreadsheet/script set by {{myEmail}}. Are you sure you want to continue?',
    errorTriggerDeleteCanceled: 'Trigger deletion has been canceled.',
    alertTitleComplete: 'Complete',
    alertMessageTriggerDelete: 'Trigger(s) deleted.',
    mailSubNewDown: '[Website Status] Alert: Site DOWN',
    mailBodyNewDown:
      'The following website(s) are DOWN:\n\n{{newDownList}}\n\n-----\nThis notice is managed by the following spreadsheet:\n{{spreadsheetUrl}}',
    mailSubResolved: '[Website Status] Notice: Site UP (Resolved)',
    mailBodyResolved:
      'The following website(s) that were DOWN are now UP:\n\n{{resolvedList}}\n\n-----\nThis notice is managed by the following spreadsheet:\n{{spreadsheetUrl}}',
    alertMessageCompleteStatusCheck: 'Website status check is complete.',
    alertMessageCompleteStatusCheckAdd:
      '\nChanges to website status have been emailed to {{myEmail}}',
    alertTitleCompleteStatusCheck: '[Website Status] Complete: Status Check',
    mailSubErrorStatusCheck: '[Website Status] Error: Status Check',
    mailBodyErrorStatusCheck:
      '{{errorStack}}\n\n-----\nThis notice is managed by the following spreadsheet:\n{{spreadsheetUrl}}',
    alertTitleError: 'ERROR',
    errorHeaderNameTargetUrlNotFound:
      '"{{headerNameTargetUrl}}" is not found. Check the "{{sheetNameDashboard}}" worksheet for the header name of the target websites\' URL.',
    errorAddTriggerWillBeDeleted:
      "\nThis function's trigger has been deleted. Fix the error and set up the trigger again.",
    errorInconsistencyInHeader:
      'There seems to be an inconsistency in the header row between the status log files. Edit the header(s) so that they match and retry.',
    alertMessageLogExtractionComplete:
      'Status Log for the last {{extractStatusLogDays}} days have been copied to the "{{sheetNameStatusLogsExtracted}}" worksheet.',
    alertMessageErrorInLogExtraction:
      '[Website Status] Error in Status Log Extraction:\n{{errorStack}}',
    errorInvalidResponseCode:
      'Invalid response code "{{code}}" at ALLOWED_RESPONSE_CODES or ERROR_RESPONSE_CODES.',
  },
  ja_JP: {
    menuTitle: 'サイト公開ステータス',
    menuTriggers: 'トリガー',
    menuSetStatusCheckTrigger: 'トリガー設定（ステータス確認）',
    menuSetLogExtractionTrigger: 'トリガー設定（ログ抽出）',
    menuDeleteTriggers: 'トリガー削除',
    menuCheckStatus: 'ステータス確認',
    menuExtractStatusLogs: '確認ログを抽出',
    errorInvalidFrequencyValue:
      '{{frequencyKey}}の値が無効です。シート「{{sheetNameOptions}}」で値を確認してください。',
    functionDescWebsiteMonitoringTriggered:
      'ウェブサイトのステータス確認のための',
    functionDescExtractStatusLogsTriggered:
      'この管理シートにステータスログを抽出してくるための',
    alertMessageContinueTriggerSetup:
      '{{functionDesc}}新規のトリガーを設定します。\n{{myEmail}} によって設定された既存のトリガーは削除・上書きされます。このまま設定を続けますか？',
    alertTitleContinueTriggerSetup: '確認：トリガー設定',
    errorTriggerSetupCanceled: 'トリガー設定がキャンセルされました。',
    errorInvalidFrequencyUnit: '無効なトリガー間隔単位: {{frequencyUnit}}',
    alertTitleCompleteTriggerSetup: '完了（{{handlerFunction}}）',
    alertMessageCompleteTriggerSetup:
      'トリガー設定完了：{{frequency}}-{{frequencyUnit}}間隔.',
    alertTitleContinueTriggerDelete: '全てのトリガーを削除します',
    alertMessageContinueTriggerDelete:
      'このスプレッドシート/スクリプトで {{myEmail}} によって設定された全てのトリガーを削除します。このまま続けますか？',
    errorTriggerDeleteCanceled: 'トリガー削除がキャンセルされました。',
    alertTitleComplete: '完了',
    alertMessageTriggerDelete: '既存のトリガーが全て削除されました。',
    mailSubNewDown: '[サイト公開ステータス] 警告：サイトが DOWN しました。',
    mailBodyNewDown:
      '次のウェブサイトが DOWN しています：\n\n{{newDownList}}\n\n-----\nこの通知は次のGoogleスプレッドシートによって管理されています：\n{{spreadsheetUrl}}',
    mailSubResolved: '[サイト公開ステータス] 復旧：サイトが復旧（UP）しました',
    mailBodyResolved:
      '次のウェブサイトが復旧しました:\n\n{{resolvedList}}\n\n-----\nこの通知は次のGoogleスプレッドシートによって管理されています：\n{{spreadsheetUrl}}',
    alertMessageCompleteStatusCheck:
      'サイト公開ステータスの確認が完了しました。',
    alertMessageCompleteStatusCheckAdd:
      '\nステータスに変化があったサイトの情報は {{myEmail}} 宛にメールで通知されます。',
    alertTitleCompleteStatusCheck:
      '[サイト公開ステータス] 完了：ステータス確認',
    mailSubErrorStatusCheck: '[サイト公開ステータス] エラー：ステータス確認',
    mailBodyErrorStatusCheck:
      '{{errorStack}}\n\n-----\nこの通知は次のGoogleスプレッドシートによって管理されています：\n{{spreadsheetUrl}}',
    alertTitleError: 'エラー',
    errorHeaderNameTargetUrlNotFound:
      '列「{{headerNameTargetUrl}}」が見つかりません。シート「{{sheetNameDashboard}}」にて、ステータス確認対象のウェブサイトURLが記載された列のヘッダ名を確認してください。',
    errorAddTriggerWillBeDeleted:
      '\nこの処理のトリガーはいったん削除されます。エラーを解決した上で、再度トリガーを設定し直してください。',
    errorInconsistencyInHeader:
      'ステータス確認のログファイルの間で、ヘッダ行に不整合があるようです。ヘッダ行が同一となるよう関係ファイルを編集・整形した上で、再度実行してください。',
    alertMessageLogExtractionComplete:
      '過去{{extractStatusLogDays}}日分のステータス確認ログがシート「{{sheetNameStatusLogsExtracted}}」に転記されました。',
    alertMessageErrorInLogExtraction:
      '[サイト公開ステータス] ログ抽出エラー：\n{{errorStack}}',
    errorInvalidResponseCode:
      '無効なレスポンスコード「{{code}}」が ALLOWED_RESPONSE_CODES または ERROR_RESPONSE_CODES にて指定されています。',
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
   * @param {Number} frequency
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
  /**
   * Replace placeholder string in this.messageList.alertMessageContinueTriggerDelete
   * @param {String} myEmail
   * @returns {String} The replaced text.
   */
  replaceAlertMessageContinueTriggerDelete(myEmail) {
    let text = this.messageList.alertMessageContinueTriggerDelete;
    let placeholderValues = [
      {
        regexp: '{{myEmail}}',
        value: myEmail,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.mailBodyNewDown
   * @param {String} newDownList
   * @param {String} spreadsheetUrl
   * @returns {String} The replaced text.
   */
  replaceMailBodyNewDown(newDownList, spreadsheetUrl) {
    let text = this.messageList.mailBodyNewDown;
    let placeholderValues = [
      {
        regexp: '{{newDownList}}',
        value: newDownList,
      },
      {
        regexp: '{{spreadsheetUrl}}',
        value: spreadsheetUrl,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.mailBodyResolved
   * @param {String} resolvedList
   * @param {String} spreadsheetUrl
   * @returns {String} The replaced text.
   */
  replaceMailBodyResolved(resolvedList, spreadsheetUrl) {
    let text = this.messageList.mailBodyResolved;
    let placeholderValues = [
      {
        regexp: '{{resolvedList}}',
        value: resolvedList,
      },
      {
        regexp: '{{spreadsheetUrl}}',
        value: spreadsheetUrl,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertMessageCompleteStatusCheckAdd
   * @param {String} myEmail
   * @returns {String} The replaced text.
   */
  replaceAlertMessageCompleteStatusCheckAdd(myEmail) {
    let text = this.messageList.alertMessageCompleteStatusCheckAdd;
    let placeholderValues = [
      {
        regexp: '{{myEmail}}',
        value: myEmail,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.mailBodyErrorStatusCheck
   * @param {String} errorStack
   * @param {String} spreadsheetUrl
   * @returns {String} The replaced text.
   */
  replaceMailBodyErrorStatusCheck(errorStack, spreadsheetUrl) {
    let text = this.messageList.mailBodyErrorStatusCheck;
    let placeholderValues = [
      {
        regexp: '{{errorStack}}',
        value: errorStack,
      },
      {
        regexp: '{{spreadsheetUrl}}',
        value: spreadsheetUrl,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.errorHeaderNameTargetUrlNotFound
   * @param {String} headerNameTargetUrl
   * @param {String} sheetNameDashboard
   * @returns {String} The replaced text.
   */
  replaceErrorHeaderNameTargetUrlNotFound(
    headerNameTargetUrl,
    sheetNameDashboard
  ) {
    let text = this.messageList.errorHeaderNameTargetUrlNotFound;
    let placeholderValues = [
      {
        regexp: '{{headerNameTargetUrl}}',
        value: headerNameTargetUrl,
      },
      {
        regexp: '{{sheetNameDashboard}}',
        value: sheetNameDashboard,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertMessageLogExtractionComplete
   * @param {Number} extractStatusLogDays
   * @param {String} sheetNameStatusLogsExtracted
   * @returns {String} The replaced text.
   */
  replaceAlertMessageLogExtractionComplete(
    extractStatusLogDays,
    sheetNameStatusLogsExtracted
  ) {
    let text = this.messageList.alertMessageLogExtractionComplete;
    let placeholderValues = [
      {
        regexp: '{{extractStatusLogDays}}',
        value: extractStatusLogDays,
      },
      {
        regexp: '{{sheetNameStatusLogsExtracted}}',
        value: sheetNameStatusLogsExtracted,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.alertMessageErrorInLogExtraction
   * @param {String} errorStack
   * @returns {String} The replaced text.
   */
  replaceAlertMessageErrorInLogExtraction(errorStack) {
    let text = this.messageList.alertMessageErrorInLogExtraction;
    let placeholderValues = [
      {
        regexp: '{{errorStack}}',
        value: errorStack,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.errorInvalidResponseCode
   * @param {String} code
   * @returns {String} The replaced text.
   */
  replaceErrorInvalidResponseCode(code) {
    let text = this.messageList.errorInvalidResponseCode;
    let placeholderValues = [
      {
        regexp: '{{code}}',
        value: code,
      },
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}
