/**
 * Original version: https://developers.google.com/adwords/scripts/docs/solutions/account-summary
 *
 * Help: https://developers.google.com/adwords/scripts/docs/reference/adwordsapp/adwordsapp_campaign
 *
 */
var SPREADSHEET_URL = '';
var SLACK_POST_URL = '';

var LOCALES = ['en-US', 'id-ID', 'es-MX', 'hu-HU', 'pl-PL', 'pt-BR', 'ro-RO', 'vi-VN'];
var CAMPAIGN_TYPES = ['search', 'display', 'remarketing'];

function main() {
  /* Opens spreadsheet */
  var spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

  collectDailyData(spreadsheet, getNextDailyDate(spreadsheet).toISOString().slice(0, 10));
  collectWeeklyData(spreadsheet, getNextWeeklyDate(spreadsheet).toISOString().slice(0, 10));
  collectMonthlyData(spreadsheet, getNextMonthlyDate(spreadsheet).toISOString().slice(0, 10));
}

function getNextDailyDate(spreadsheet) {
  var lastDateAsMillisecondsSince1970 = spreadsheet.getRangeByName('lastDailyCampaignsDate').getValue().getTime();
  var nextDay = new Date();
  nextDay.setTime(lastDateAsMillisecondsSince1970 + 24 * 60 * 60 * 1000 - 28 * 24 * 60 * 60 * 1000); /* Calculates next day and creates a 28-days window */
  return nextDay;
}

function getNextWeeklyDate(spreadsheet) {
  var lastDateAsMillisecondsSince1970 = spreadsheet.getRangeByName('lastWeeklyCampaignsDate').getValue().getTime();
  var nextDay = new Date();
  nextDay.setTime(lastDateAsMillisecondsSince1970 + 24 * 60 * 60 * 1000 - 28 * 24 * 60 * 60 * 1000); /* Calculates next day and creates a 28-days window */
  return nextDay;
}

function getNextMonthlyDate(spreadsheet) {
  var lastDateAsMillisecondsSince1970 = spreadsheet.getRangeByName('lastMonthlyCampaignsDate').getValue().getTime();
  var nextDay = new Date();
  nextDay.setTime(lastDateAsMillisecondsSince1970 + 24 * 60 * 60 * 1000); /* Calculates next day and steps back a month */
  nextDay.setMonth(nextDay.getMonth() - 1);
  return nextDay;
}

/**
 * Collects data for several days between the given startDate and the current date/time.
 *
 */
function collectDailyData(spreadsheet, startDate) {
  /* Sets up first day to process */
  startDate = convertDateToTimeZone(new Date(startDate), 0);
  var dayCountPerPeriod = 1;

  Logger.log('collectDailyData has started from: ' + startDate);

  /* Gets data */
  var outputRows = getOutputRowsForDateRange(startDate, new Date(new Date().getTime() - 24 * 3600 * 1000), dayCountPerPeriod);

  /* Writes out output */
  if (outputRows.length > 0) {
    writeToSpreadsheet(spreadsheet, outputRows, 'Daily-Campaigns', startDate);
  }

  Logger.log('collectDailyData has finished!');
}

function collectWeeklyData(spreadsheet, startDate) {
  /* Sets up first day to process */
  startDate = convertDateToTimeZone(new Date(startDate), 0);
  var dayCountPerPeriod = 7;

  Logger.log('collectWeeklyData has started from: ' + startDate);

  /* Gets data */
  var outputRows = getOutputRowsForDateRange(startDate, new Date(new Date().getTime() - 24 * 3600 * 1000), dayCountPerPeriod);

  /* Writes out output */
  if (outputRows.length > 0) {
    writeToSpreadsheet(spreadsheet, outputRows, 'Weekly-Campaigns', startDate);
    if (new Date().getDay() === 5) {
      sendToSlack(outputRows);
    }
  }

  Logger.log('collectWeeklyData has finished!');
}

function collectMonthlyData(spreadsheet, startDate) {
  /* Sets up first day to process */
  startDate = convertDateToTimeZone(new Date(startDate), 0);

  Logger.log('collectMonthlyData has started from: ' + startDate);

  /* Gets data */
  var outputRows = getOutputRowsForMonthlyDateRange(startDate, new Date(new Date().getTime() - 24 * 3600 * 1000));

  /* Writes out output */
  if (outputRows.length > 0) {
    writeMonthlyToSpreadsheet(spreadsheet, outputRows, 'Monthly-Campaigns', startDate);
  }

  Logger.log('collectMonthlyData has finished!');
}

function getOutputRowsForDateRange(startDate, endDate, dayCountPerPeriod) {
  var periodStartDate = new Date(startDate);
  var periodEndDate = new Date(startDate.getTime() + 24 * 3600 * 1000 * (dayCountPerPeriod - 1));
  /* Assembles output from each period's report */
  var outputRows = [];
  while (periodEndDate.getTime() <= endDate.getTime()) {
    var campaignData = getCampaignDataForDateRange(AdWordsApp.campaigns().get(), periodStartDate, periodEndDate);
    var rows = flattenCampaignData(campaignData);

    rows.map(function(row) {
      outputRows.push([new Date(periodStartDate), new Date(periodEndDate), row.locale, row.campaign, row.cost, row.impressions, row.clicks, row.averagePosition, row.signedUpForTrialPlan, row.signedUpForPaidPlan]);
    });
    periodStartDate.setDate(periodStartDate.getDate() + dayCountPerPeriod);
    periodEndDate.setDate(periodEndDate.getDate() + dayCountPerPeriod);
  }

  return outputRows;
}

function getOutputRowsForMonthlyDateRange(startDate, endDate) {
  var periodStartDate = new Date(startDate);
  var periodEndDate = new Date(periodStartDate.getFullYear(), periodStartDate.getMonth() + 1, 0);
  /* Assembles output from each period's report */
  var outputRows = [];
  while (periodEndDate.getTime() <= endDate.getTime()) {
    var campaignData = getCampaignDataForDateRange(AdWordsApp.campaigns().get(), periodStartDate, periodEndDate);
    var rows = flattenCampaignData(campaignData);

    rows.map(function(row) {
      outputRows.push([new Date(periodStartDate), new Date(periodEndDate), row.locale, row.campaign, row.cost, row.impressions, row.clicks, row.averagePosition, row.signedUpForTrialPlan, row.signedUpForPaidPlan]);
    });
    periodStartDate.setDate(periodEndDate.getDate() + 1);
    periodEndDate = new Date(periodStartDate.getFullYear(), periodStartDate.getMonth() + 1, 0);
  }

  return outputRows;
}

function getCampaignDataForDateRange(campaigns, startDate, endDate) {
  var startDateString = Utilities.formatDate(startDate, AdWordsApp.currentAccount().getTimeZone(), 'yyyyMMdd');
  var endDateString = Utilities.formatDate(endDate, AdWordsApp.currentAccount().getTimeZone(), 'yyyyMMdd');

  var campaignData = {};
  var campaignTrialConversions = {};
  var campaignPaidConversions = {};

  /* Get conversions */
  var trialConversionsRows = AdsApp.report("SELECT CampaignName, AllConversions FROM CAMPAIGN_PERFORMANCE_REPORT WHERE ConversionTypeName='[ACQ] CodeBerry - signedUpForTrialPlan' DURING " + startDateString + "," + endDateString).rows();
  var paidConversionsRows = AdsApp.report("SELECT CampaignName, AllConversions FROM CAMPAIGN_PERFORMANCE_REPORT WHERE ConversionTypeName='[ACQ] CodeBerry - signedUpForPaidPlan' DURING " + startDateString + "," + endDateString).rows();

  while (trialConversionsRows.hasNext()) {
    var trialRow = trialConversionsRows.next();
    if (trialRow.AllConversions){
      campaignTrialConversions[trialRow.CampaignName] = trialRow.AllConversions;
    } else {
      campaignTrialConversions[trialRow.CampaignName] = 0;
    }
  }
  while (paidConversionsRows.hasNext()) {
    var paidRow = paidConversionsRows.next();
    if (paidRow.AllConversions){
      campaignPaidConversions[paidRow.CampaignName] = paidRow.AllConversions;
    } else {
      campaignPaidConversions[paidRow.CampaignName] = 0;
    }
  }

  /* Interates through all campaigns and retreives their data */
  while (campaigns.hasNext()) {
    /* Loads campaign stats */
    var campaign = campaigns.next();
    if (!campaign.getName().match(/codeberry/)) { continue; }
    var campaignStats = campaign.getStatsFor(startDateString, endDateString);

    /* Determines campaign locale and name */
    var locale = findArrayItemInString(campaign.getName(), LOCALES);
    var campaignName = campaign.getName();
    if (!locale || !campaignName) { continue; }

    /* Saves results to array */
    if (typeof (campaignData[locale]) === 'undefined') { campaignData[locale] = {}; }
    if (typeof (campaignData[locale][campaignName]) === 'undefined') { campaignData[locale][campaignName] = {count: 0, impressions: 0, clicks: 0, cost: 0, averagePosition: 0, signedUpForTrialPlan: 0, signedUpForPaidPlan: 0}; }
    campaignData[locale][campaignName].count += 1;
    campaignData[locale][campaignName].impressions += campaignStats.getImpressions();
    campaignData[locale][campaignName].clicks += campaignStats.getClicks();
    campaignData[locale][campaignName].cost += campaignStats.getCost();
    campaignData[locale][campaignName].averagePosition = (campaignData[locale][campaignName].count > 1)
        ? (campaignData[locale][campaignName].averagePosition + campaignStats.getAveragePosition()) * (campaignData[locale][campaignName].count - 1) / campaignData[locale][campaignName].count
        : campaignStats.getAveragePosition();
    campaignData[locale][campaignName].signedUpForTrialPlan += (campaignTrialConversions[campaignName]) ? campaignTrialConversions[campaignName] : 0;
    campaignData[locale][campaignName].signedUpForPaidPlan += (campaignPaidConversions[campaignName]) ? campaignPaidConversions[campaignName] : 0;
  }

  return campaignData;
}

function flattenCampaignData(campaignData) {
  /* Flattens output and returns it */
  var outputArray = [];

  for (var localeIndex = 0; localeIndex < LOCALES.length; localeIndex++) {
    var locale = LOCALES[localeIndex];

    for (var campaignName in campaignData[locale]) {
      if (campaignData[locale].hasOwnProperty(campaignName)) {
        if (((campaignData[locale] || [])[campaignName] || []).impressions) {
          outputArray.push({
            locale: locale,
            campaign: campaignName,
            impressions: campaignData[locale][campaignName].impressions,
            clicks: campaignData[locale][campaignName].clicks,
            cost: campaignData[locale][campaignName].cost,
            averagePosition: campaignData[locale][campaignName].averagePosition,
            signedUpForTrialPlan: campaignData[locale][campaignName].signedUpForTrialPlan,
            signedUpForPaidPlan: campaignData[locale][campaignName].signedUpForPaidPlan
          });
        }
      }
    }
  }
  return outputArray;
}

function findArrayItemInString(string, array) {
  for (var i = 0; i < array.length; i++) {
    if (string.match(array[i])) {
      return array[i];
    }
  }
  return false;
}

/**
 * Appends some data rows to the spreadsheet.
 * @param {object} spreadsheet The spreadsheet object returned by SpreadsheetApp.openByUrl(SPREADSHEET_URL).
 * @param {Array<Array<string>>} rows The data rows.
 * @param {string} sheetName The name of the sheet.
 * @param {object} startDate
 */
function writeToSpreadsheet(spreadsheet, rows, sheetName, startDate) {
  var access = new SpreadsheetAccess(spreadsheet, sheetName);
  var sheet = spreadsheet.getSheetByName(sheetName);

  var startDateString = Utilities.formatDate(startDate, AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  var startRowIndex = findInColumn(sheet, 'A', startDateString);
  var endRowIndex = sheet.getLastRow();//Number of last row with content
  //blank rows after last row with content will not be deleted

  sheet.deleteRows(startRowIndex, endRowIndex - startRowIndex + 1);

  var emptyRow = access.findEmptyRow(2, 1);
  if (emptyRow < 0) {
    access.addRows(rows.length);
    emptyRow = access.findEmptyRow(2, 1);
  }
  access.writeRows(rows, emptyRow, 1);
}

function writeMonthlyToSpreadsheet(spreadsheet, rows, sheetName, startDate) {
  var access = new SpreadsheetAccess(spreadsheet, sheetName);
  var sheet = spreadsheet.getSheetByName(sheetName);

  var startDateString = Utilities.formatDate(new Date(startDate.getFullYear(), startDate.getMonth(), 1), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  var startRowIndex = findInColumn(sheet, 'A', startDateString);
  var endRowIndex = sheet.getLastRow();//Number of last row with content
  //blank rows after last row with content will not be deleted

  sheet.deleteRows(startRowIndex, endRowIndex - startRowIndex + 1);

  var emptyRow = access.findEmptyRow(2, 1);
  if (emptyRow < 0) {
    access.addRows(rows.length);
    emptyRow = access.findEmptyRow(2, 1);
  }
  access.writeRows(rows, emptyRow, 1);
}

function SpreadsheetAccess(spreadsheet, sheetName) {
  this.spreadsheet = spreadsheet;
  this.sheet = this.spreadsheet.getSheetByName(sheetName);

  // what column should we be looking at to check whether the row is empty?
  this.findEmptyRow = function(minRow, column) {
    var values = this.sheet.getRange(minRow, column,
        this.sheet.getMaxRows(), 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (!values[i][0]) {
        return i + minRow;
      }
    }
    return -1;
  };
  this.addRows = function(howMany) {
    this.sheet.insertRowsAfter(this.sheet.getMaxRows(), howMany);
  };
  this.writeRows = function(rows, startRow, startColumn) {
    this.sheet.getRange(startRow, startColumn, rows.length, rows[0].length).setValues(rows);
  };
}

function findInColumn(spreadsheet, column, data) {
  var columnRange = spreadsheet.getRange(column + ":" + column);  // like A:A
  var values = columnRange.getValues();
  var row = 0;

  for (i = 1; i < values.length; i++) {
    values[i] = Utilities.formatDate(new Date(values[i][0]), AdWordsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  }

  while (values[row] && values[row] !== data) {
    row++;
  }

  if (values[row] === data) {
    return row + 1;
  } else {
    return -1;
  }
}

function sendToSlack(rows) {
  var last_date = rows[rows.length-1][0];
  var last_date_end = rows[rows.length-1][1];
  var rows_filtered = rows.filter(function(row) {
    return row[0].getDate() === last_date.getDate() && row[0].getMonth() === last_date.getMonth() && row[0].getFullYear() === last_date.getFullYear() &&
        row[1].getDate() === last_date_end.getDate() && row[1].getMonth() === last_date_end.getMonth() && row[1].getFullYear() === last_date_end.getFullYear();
  });

  var campaigns = [];

  rows_filtered.forEach(function(row){
    campaigns.push({
      title: row[3],
      value: '[cost: ' + parseInt(row[4]).toFixed(0) + '] [pos: ' + parseInt(row[7]).toFixed(0) +  '] [trial: ' + parseInt(row[8]).toFixed(0) + '] [paid: ' + parseInt(row[9]).toFixed(0) +  ']',
      short: false
    });
  });

  var month = '' + (last_date.getMonth() + 1);
  var day = '' + last_date.getDate();
  var year = last_date.getFullYear();
  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;
  var dateString = [year, month, day].join('.');

  var month_end = '' + (last_date_end.getMonth() + 1);
  var day_end = '' + last_date_end.getDate();
  var year_end = last_date_end.getFullYear();
  if (month_end.length < 2) month_end = '0' + month_end;
  if (day_end.length < 2) day_end = '0' + day_end;
  var dateStringEnd = [year_end, month_end, day_end].join('.');

  var payload = {
    'channel' : '#ad-stats',
    'username' : 'Google Ads Stats Bot',
    'icon_url' : 'https://upload.wikimedia.org/wikipedia/commons/thumb/c/c7/Google_Ads_logo.svg/820px-Google_Ads_logo.svg.png',
    'attachments': [{
      "fallback": "Required plain-text summary of the attachment.",
      "color": "#2eb886",
      "pretext": "Here’s the weekly stats for Google Ads campaigns: " + dateString + "-" + dateStringEnd,
      "author_name": "Google Ads account summary report",
      "author_link": "",
      "fields": campaigns,
      "footer": "Made with ❤️ by a global team."
    }]
  };

  var options = {
    'method' : 'post',
    'contentType' : 'application/json',
    'payload' : JSON.stringify(payload)
  };

  UrlFetchApp.fetch(SLACK_POST_URL, options);
}

/**
 * @param {object} date The date that will be converted to time zone
 * @param {int} offsetInHours E.g. give it 0 for GMT, 1 for Budapest, -8 for LA.
 */
function convertDateToTimeZone(date, offsetInHours) {
  var utcTimeInMilliseconds = date.getTime() + date.getTimezoneOffset() * 60 * 1000;

  return new Date(utcTimeInMilliseconds + (60 * 60 * 1000 * offsetInHours));
}