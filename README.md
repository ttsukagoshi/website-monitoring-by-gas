# Website Status Monitoring

[![GitHub Super-Linter](https://github.com/ttsukagoshi/website-monitoring-by-gas/workflows/Lint%20Code%20Base/badge.svg)](https://github.com/marketplace/actions/super-linter) [![clasp](https://img.shields.io/badge/built%20with-clasp-4285f4.svg?style=flat-square)](https://github.com/google/clasp) [![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg?style=flat-square)](https://github.com/prettier/prettier)

Website status monitoring using Google Sheets and Google Apps Script.

## About

A Spreadsheet-bound apps script solution to conduct automated status monitoring on websites listed by the user in a Google Sheets management file. A separate status log file in Google Sheets will be created so that users can easily integrate data with BI services such as Google Data Studio.

## Install

Copy [this sample spreadsheet](https://docs.google.com/spreadsheets/d/1JvO090VcgvF-WwciNnzRb1_nonKJC5QHN73h_CXS1Cw/edit#gid=0) to your Google Drive.

## How to Use

### Setup

#### `01_Dashboard` Worksheet

Replace the `WEBSITE NAME` and `TARGET URL` columns with those of the website(s) that you want to monitor.

#### `90_Spreadsheets` Worksheet

Delete everything **except the first row**.

#### `99_Options` Worksheet

Go over the parameters that you can set for this status monitoring and edit the `VALUE` items to suit you needs.

#### Set Triggers

From the spreadsheet menu, select `Web Status` > `Triggers` > `Set Status Check Trigger`/`Set Log Extraction Trigger` to set up time-based triggers to conduct automated status checks. The latest results will be shown in the `01_Dashboard` worksheet.

You will be asked to authorized the script the first time you execute it. Users of free Gmail account should expect to see the `Unverified` warning during this authorization process. Note that the owner of the script is yourself, and that this solution will not send or receive any information to and from any other Google accounts or services outside the Google ecosystem (except for checking the HTTP response codes of the websites that you designated because, well, that's what it does for status monitoring) unless you explicitly share the spreadsheet.

## Updates

Updates will be distributed via [@ttsukagoshi/website-monitoring-by-gas (GitHub)](https://github.com/ttsukagoshi/website-monitoring-by-gas).

## Terms and Conditions

You must agree to the [Terms and Conditions](https://www.scriptable-assets.page/terms-and-conditions/) to use this solution.
