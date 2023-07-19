/**
 * @name Link Checker - Manager Accounts
 *
 * @overview The Link Checker script iterates through the ads, keywords, and
 *     sitelinks in your manager account hierarchy and makes sure their URLs
 *     do not produce "Page not found" or other types of error responses.
 *
 * @author Said Tezel [saidtezel@gmail.com]
 *
 * @version 1.0
 *
**/

var CONFIG = {
  // URL of the spreadsheet template.
  // This should be a copy of https://goo.gl/Ysrg9s.
  SPREADSHEET_URL: 'SHEET_URL',

  // Array of addresses to be alerted via email if issues are found.
  RECIPIENT_EMAILS: [
    'said.tezel@foundgroup.co.uk'
  ],

  // Label to use when a link has been checked.
  LABEL: 'linkChecker_complete',

  // A list of ManagedAccountSelector conditions to restrict the population
  // of child accounts that will be processed. Leave blank or comment out
  // to include all child accounts.
  ACCOUNT_CONDITIONS: [],

  // Number of milliseconds to sleep after each URL request. Use this throttle
  // to reduce the load that the script imposes on your web server(s).
  THROTTLE: 0,

  // Number of seconds before timeout that the script should stop checking URLs
  // to make sure it has time to output its findings.
  TIMEOUT_BUFFER: 120,

  // Send the script results to the #ppc channel on Slack.
  SLACK_URL: 'SLACK_WEBHOOK_URL'
};

/**
 * Parameters controlling the script's behavior after hitting a UrlFetchApp
 * QPS quota limit.
 */
var QUOTA_CONFIG = {
  // Initial number of milliseconds to sleep.
  INIT_SLEEP_TIME: 150,

  // Multiplicative factor to increase sleep time by on each retry.
  BACKOFF_FACTOR: 1.5,

  // Maximum number of tries for a single URL.
  MAX_TRIES: 3
};

/**
 * Exceptions that prevent the script from finishing checking all URLs in an
 * account but allow it to resume next time.
 */
var EXCEPTIONS = {
  QPS: 'Reached UrlFetchApp QPS limit',
  LIMIT: 'Reached UrlFetchApp daily quota',
  TIMEOUT: 'Approached script execution time limit'
};

/**
 * Named ranges in the spreadsheet.
 */
var NAMES = {
  CHECK_AD_URLS: 'checkAdUrls',
  CHECK_KEYWORD_URLS: 'checkKeywordUrls',
  CHECK_SITELINK_URLS: 'checkSitelinkUrls',
  CHECK_PAUSED_ADS: 'checkPausedAds',
  CHECK_PAUSED_KEYWORDS: 'checkPausedKeywords',
  CHECK_PAUSED_SITELINKS: 'checkPausedSitelinks',
  VALID_CODES: 'validCodes',
  EMAIL_EACH_RUN: 'emailEachRun',
  EMAIL_NON_ERRORS: 'emailNonErrors',
  EMAIL_ON_COMPLETION: 'emailOnCompletion',
  SAVE_ALL_URLS: 'saveAllUrls',
  FREQUENCY: 'frequency',
  DATE_STARTED: 'dateStarted',
  DATE_COMPLETED: 'dateCompleted',
  DATE_EMAILED: 'dateEmailed',
  NUM_ERRORS: 'numErrors',
  RESULT_HEADERS: 'resultHeaders',
  ARCHIVE_HEADERS: 'archiveHeaders'
};

function main() {
  var spreadsheet = validateAndGetSpreadsheet(CONFIG.SPREADSHEET_URL);
  validateEmailAddresses();
  spreadsheet.setSpreadsheetTimeZone(AdWordsApp.currentAccount().getTimeZone());

  var options = loadOptions(spreadsheet);
  var status = loadStatus(spreadsheet);
  Logger.log(status);

  if (!status.dateStarted) {
    // This is the very first execution of the script.
    startNewAnalysis(spreadsheet);
  } else if (status.dateStarted > status.dateCompleted) {
    Logger.log('Resuming work from a previous execution.');
  } else if (dayDifference(status.dateStarted, new Date()) <
             options.frequency) {
    Logger.log('Waiting until ' + options.frequency +
               ' days have elapsed since the start of the last analysis.');
    return;
  } else {
    // Enough time has passed since the last analysis to start a new one.
    removeLabelsInAccounts();
    removeAccountLabels([CONFIG.LABEL]);
    startNewAnalysis(spreadsheet);
  }

  ensureAccountLabels([CONFIG.LABEL]);

  // Get up to 50 accounts that have not yet had all of their URLs checked.
  var accountSelector = getAccounts(false).withLimit(50);
  if (accountSelector.get().hasNext()) {
    accountSelector.executeInParallel('processAccount', 'processResults',
                                      JSON.stringify(options));
  } else {
    processResults([]);
  }
}

/**
 * Retrieves all accounts that either have had their URLs checked or need to
 * have their URLs checked.
 *
 * @param {boolean} isChecked True to get accounts that have been checked
 *     already, false to get accounts that have not have been checked already.
 *     Ignored if the label does not exist.
 * @return {Object} An account selector.
 */
function getAccounts(isChecked) {
  var accountSelector = MccApp.accounts();

  if (getAccountLabel(CONFIG.LABEL)) {
    accountSelector = accountSelector.forDateRange('LAST_30_DAYS').withCondition('Cost > 0').withCondition('LabelNames ' + (isChecked ? 'CONTAINS' : 'DOES_NOT_CONTAIN') + ' "' + CONFIG.LABEL + '"')
  }

  if (CONFIG.ACCOUNT_CONDITIONS) {
    for (var i = 0; i < CONFIG.ACCOUNT_CONDITIONS.length; i++) {
      accountSelector =
          accountSelector.withCondition(CONFIG.ACCOUNT_CONDITIONS[i]);
    }
  }

  return accountSelector;
}

/**
 * Removes the tracking in each account that was previously analyzed, thereby
 * clearing that account for a new analysis.
 */
function removeLabelsInAccounts() {
  var managerAccount = AdWordsApp.currentAccount();
  var accounts = getAccounts(true).get();

  while (accounts.hasNext()) {
    MccApp.select(accounts.next());
    removeLabels([CONFIG.LABEL]);
  }

  MccApp.select(managerAccount);
}

/**
 * Performs the link checking analysis on the current account.
 *
 * @param {string} options Options from the spreadsheet as JSON.
 * @return {string} JSON stringified results of the analysis.
 */
function processAccount(options) {
  return JSON.stringify(analyzeAccount(JSON.parse(options)));
}

/**
 * Consolidates results from each account and outputs them.
 *
 * @param {Array.<Object>} executionResults A list of ExecutionResult objects.
 */
function processResults(executionResults) {
  var spreadsheet = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  var options = loadOptions(spreadsheet);
  var results = {
    urlChecks: [],
    didComplete: true
  };

  for (var i = 0; i < executionResults.length; i++) {
    if (!executionResults[i].getError()) {
      var accountResult = JSON.parse(executionResults[i].getReturnValue());
      results.urlChecks = results.urlChecks.concat(accountResult.urlChecks);
      results.didComplete = results.didComplete && accountResult.didComplete;

      if (accountResult.didComplete) {
        MccApp.accounts().withIds([executionResults[i].getCustomerId()]).get().
            next().applyLabel(CONFIG.LABEL);
      }
    } else {
      Logger.log('Processing for ' + executionResults[i].getCustomerId() +
          ' failed.');
    }
  }

  // The entire analysis is not complete if there are any accounts that have
  // not been labeled (i.e., the account was not started, or not all URLs in
  // the account have been checked).
  results.didComplete = results.didComplete &&
      !getAccounts(false).get().hasNext();

  outputResults(results, options);
}

/**
 * Checks as many new URLs as possible that have not previously been checked,
 * subject to quota and time limits.
 *
 * @param {Object} options Dictionary of options.
 * @return {Object} An object with fields for the URLs checked and an indication
 *     if the analysis was completed (no remaining URLs to check).
 */
function analyzeAccount(options) {
  // Ensure the label exists before attempting to retrieve already checked URLs.
  ensureLabels([CONFIG.LABEL]);

  var checkedUrls = getAlreadyCheckedUrls(options);
  var urlChecks = [];
  var didComplete = false;

  try {
    // If the script throws an exception, didComplete will remain false.
    didComplete = checkUrls(checkedUrls, urlChecks, options);
  } catch(e) {
    if (e == EXCEPTIONS.QPS ||
        e == EXCEPTIONS.LIMIT ||
        e == EXCEPTIONS.TIMEOUT) {
      Logger.log('Stopped checking URLs early because: ' + e);
      Logger.log('Checked URLs will still be output.');
    } else {
      throw e;
    }
  }

  return {
    urlChecks: urlChecks,
    didComplete: didComplete
  };
}

/**
 * Outputs the results to a spreadsheet and sends emails if appropriate.
 *
 * @param {Object} results An object with fields for the URLs checked and an
 *     indication if the analysis was completed (no remaining URLs to check).
 * @param {Object} options Dictionary of options.
 */
function outputResults(results, options) {
  var spreadsheet = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);

  var numErrors = countErrors(results.urlChecks, options);
  Logger.log('Found ' + numErrors + ' this execution.');

  saveUrlsToSpreadsheet(spreadsheet, results.urlChecks, options);

  // Reload the status to get the total number of errors for the entire
  // analysis, which is calculated by the spreadsheet.
  status = loadStatus(spreadsheet);

  if (results.didComplete) {
    spreadsheet.getRangeByName(NAMES.DATE_COMPLETED).setValue(new Date());
    Logger.log('Found ' + status.numErrors + ' across the entire analysis.');
  }

  if (CONFIG.SLACK_URL) {
    if (results.didComplete && (options.emailEachRun || options.emailOnCompletion) && status.numErrors > 0) {
      sendSlackMessage(spreadsheet, status.numErrors);
    }
  }
}

/**
 * Loads data from a spreadsheet based on named ranges. Strings 'Yes' and 'No'
 * are converted to booleans. One-dimensional ranges are converted to arrays
 * with blank cells omitted. Assumes each named range exists.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @param {Array.<string>} names A list of named ranges that should be loaded.
 * @return {Object} A dictionary with the names as keys and the values
 *     as the cell values from the spreadsheet.
 */
function loadDatabyName(spreadsheet, names) {
  var data = {};

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var range = spreadsheet.getRangeByName(name);

    if (range.getNumRows() > 1 && range.getNumColumns() > 1) {
      // Name refers to a 2d range, so load it as a 2d array.
      data[name] = range.getValues();
    } else if (range.getNumRows() == 1 && range.getNumColumns() == 1) {
      // Name refers to a single cell, so load it as a value and replace
      // Yes/No with boolean true/false.
      data[name] = range.getValue();
      data[name] = data[name] === 'Yes' ? true : data[name];
      data[name] = data[name] === 'No' ? false : data[name];
    } else {
      // Name refers to a 1d range, so load it as an array (regardless of
      // whether the 1d range is oriented horizontally or vertically).
      var isByRow = range.getNumRows() > 1;
      var limit = isByRow ? range.getNumRows() : range.getNumColumns();
      var cellValues = range.getValues();

      data[name] = [];
      for (var j = 0; j < limit; j++) {
        var cellValue = isByRow ? cellValues[j][0] : cellValues[0][j];
        if (cellValue) {
          data[name].push(cellValue);
        }
      }
    }
  }

  return data;
}

/**
 * Loads options from the spreadsheet.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @return {Object} A dictionary of options.
 */
function loadOptions(spreadsheet) {
  return loadDatabyName(spreadsheet,
      [NAMES.CHECK_AD_URLS, NAMES.CHECK_KEYWORD_URLS,
       NAMES.CHECK_SITELINK_URLS, NAMES.CHECK_PAUSED_ADS,
       NAMES.CHECK_PAUSED_KEYWORDS, NAMES.CHECK_PAUSED_SITELINKS,
       NAMES.VALID_CODES, NAMES.EMAIL_EACH_RUN,
       NAMES.EMAIL_NON_ERRORS, NAMES.EMAIL_ON_COMPLETION,
       NAMES.SAVE_ALL_URLS, NAMES.FREQUENCY]);
}

/**
 * Loads state information from the spreadsheet.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @return {Object} A dictionary of status information.
 */
function loadStatus(spreadsheet) {
  return loadDatabyName(spreadsheet,
      [NAMES.DATE_STARTED, NAMES.DATE_COMPLETED,
       NAMES.DATE_EMAILED, NAMES.NUM_ERRORS]);
}

/**
 * Saves the start date to the spreadsheet and archives results of the last
 * analysis to a separate sheet.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 */
function startNewAnalysis(spreadsheet) {
  Logger.log('Starting a new analysis.');

  spreadsheet.getRangeByName(NAMES.DATE_STARTED).setValue(new Date());

  // Helper method to get the output area on the results or archive sheets.
  var getOutputRange = function(rangeName) {
    var headers = spreadsheet.getRangeByName(rangeName);
    return headers.offset(1, 0, headers.getSheet().getDataRange().getLastRow());
  };

  getOutputRange(NAMES.ARCHIVE_HEADERS).clearContent();

  var results = getOutputRange(NAMES.RESULT_HEADERS);
  results.copyTo(getOutputRange(NAMES.ARCHIVE_HEADERS));

  getOutputRange(NAMES.RESULT_HEADERS).clearContent();
}

/**
 * Counts the number of errors in the results.
 *
 * @param {Array.<Object>} urlChecks A list of URL check results.
 * @param {Object} options Dictionary of options.
 * @return {number} The number of errors in the results.
 */
function countErrors(urlChecks, options) {
  var numErrors = 0;

  for (var i = 0; i < urlChecks.length; i++) {
    if (options.validCodes.indexOf(urlChecks[i].responseCode) == -1) {
      numErrors++;
    }
  }

  return numErrors;
}

/**
 * Saves URLs for a particular account to the spreadsheet starting at the first
 * unused row.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @param {Array.<Object>} urlChecks A list of URL check results.
 * @param {Object} options Dictionary of options.
 */
function saveUrlsToSpreadsheet(spreadsheet, urlChecks, options) {
  // Build each row of output values in the order of the columns.
  var outputValues = [];
  for (var i = 0; i < urlChecks.length; i++) {
    var urlCheck = urlChecks[i];

    if (options.saveAllUrls ||
        options.validCodes.indexOf(urlCheck.responseCode) == -1) {
      outputValues.push([
        urlCheck.customerId,
        new Date(urlCheck.timestamp),
        urlCheck.url,
        urlCheck.responseCode,
        urlCheck.entityType,
        urlCheck.campaign,
        urlCheck.adGroup,
        urlCheck.ad,
        urlCheck.keyword,
        urlCheck.sitelink
      ]);
    }
  }

  if (outputValues.length > 0) {
    // Find the first open row on the Results tab below the headers and create a
    // range large enough to hold all of the output, one per row.
    var headers = spreadsheet.getRangeByName(NAMES.RESULT_HEADERS);
    var lastRow = headers.getSheet().getDataRange().getLastRow();
    var outputRange = headers.offset(lastRow - headers.getRow() + 1,
                                     0, outputValues.length);
    outputRange.setValues(outputValues);
  }

  for (var i = 0; i < CONFIG.RECIPIENT_EMAILS.length; i++) {
    spreadsheet.addEditor(CONFIG.RECIPIENT_EMAILS[i]);
  }
}

/**
 * Sends an email to a list of email addresses with a link to the spreadsheet
 * and the results of this execution of the script.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @param {boolean} numErrors The number of errors found in this execution.
 */
function sendIntermediateEmail(spreadsheet, numErrors) {
  spreadsheet.getRangeByName(NAMES.DATE_EMAILED).setValue(new Date());

  MailApp.sendEmail(CONFIG.RECIPIENT_EMAILS.join(','),
      'Link Checker Results',
      'The Link Checker script found ' + numErrors + ' URLs with errors in ' +
      'an execution that just finished. See ' +
      spreadsheet.getUrl() + ' for details.');
}

/**
 * Sends an email to a list of email addresses with a link to the spreadsheet
 * and the results across the entire account.
 *
 * @param {Object} spreadsheet The spreadsheet object.
 * @param {boolean} numErrors The number of errors found in the entire account.
 */
function sendFinalEmail(spreadsheet, numErrors) {
  spreadsheet.getRangeByName(NAMES.DATE_EMAILED).setValue(new Date());

  MailApp.sendEmail(CONFIG.RECIPIENT_EMAILS.join(','),
      'Link Checker Results',
      'The Link Checker script found ' + numErrors + ' URLs with errors ' +
      'across its entire analysis. See ' +
      spreadsheet.getUrl() + ' for details.');
}


/**
* Sends a message to Slack
* @param {Object} spreadsheet The spreadsheet object.
* @param {integer} numErrors The number of errors found in the entire account.
*/
function sendSlackMessage(spreadsheet, numErrors) {
var slackMessage = {
  attachments: [
    {
      fallback: 'The Link Checker script found ' + numErrors + ' URLs with errors across its entire analysis. See ' + spreadsheet.getUrl() + ' for details.',
      color: "#36a64f",
      pretext: "There are some broken links recently detected in the Found Agency Account.",
      author_name: "adSmart",
      author_icon: "https://storage.googleapis.com/gweb-uniblog-publish-prod/images/logo_Google_Ads_192px.max-200x200.png",
      title: "Broken Link Checker",
      title_link: spreadsheet.getUrl(),
      text: "Please check the linked Google Sheet to review the latest broken links detected.",
      fields: [
        {
          title: "URLs with Errors",
          value: numErrors,
          short: false
        }
      ],
    }
  ]
};

var options = {
  method: 'POST',
  contentType: 'application/json',
  payload: JSON.stringify(slackMessage)
};

UrlFetchApp.fetch(CONFIG.SLACK_URL, options);
}

/**
 * Retrieves all final URLs and mobile final URLs in the account across ads,
 * keywords, and sitelinks that were checked in a previous run, as indicated by
 * them having been labeled.
 *
 * @param {Object} options Dictionary of options.
 * @return {Object} A map of previously checked URLs with the URL as the key.
 */
function getAlreadyCheckedUrls(options) {
  var urlMap = {};

  var addToMap = function(items) {
    for (var i = 0; i < items.length; i++) {
      var urls = expandUrlModifiers(items[i]);
      urls.forEach(function(url) {
        urlMap[url] = true;
      });
    }
  };

  if (options.checkAdUrls) {
    addToMap(getUrlsBySelector(AdWordsApp.ads().
                               withCondition(labelCondition(true))));
  }

  if (options.checkKeywordUrls) {
    addToMap(getUrlsBySelector(AdWordsApp.keywords().
                               withCondition(labelCondition(true))));
  }

  if (options.checkSitelinkUrls) {
    addToMap(getAlreadyCheckedSitelinkUrls());
  }

  return urlMap;
}

/**
 * Retrieves all final URLs and mobile final URLs for campaign and ad group
 * sitelinks.
 *
 * @return {Array.<string>} An array of URLs.
 */
function getAlreadyCheckedSitelinkUrls() {
  var urls = [];

  // Helper method to get campaign or ad group sitelink URLs.
  var addSitelinkUrls = function(selector) {
    var iterator = selector.withCondition(labelCondition(true)).get();

    while (iterator.hasNext()) {
      var entity = iterator.next();
      var sitelinks = entity.extensions().sitelinks();
      urls = urls.concat(getUrlsBySelector(sitelinks));
    }
  };

  addSitelinkUrls(AdWordsApp.campaigns());
  addSitelinkUrls(AdWordsApp.adGroups());

  return urls;
}

/**
 * Retrieves all URLs in the entities specified by a selector.
 *
 * @param {Object} selector The selector specifying the entities to use.
 *     The entities should be of a type that has a urls() method.
 * @return {Array.<string>} An array of URLs.
 */
function getUrlsBySelector(selector) {
  var urls = [];
  var entities = selector.get();

  // Helper method to add the url to the list if it exists.
  var addToList = function(url) {
    if (url) {
      urls.push(url);
    }
  };

  while (entities.hasNext()) {
    var entity = entities.next();

    addToList(entity.urls().getFinalUrl());
    addToList(entity.urls().getMobileFinalUrl());
  }

  return urls;
}

/**
 * Retrieves all final URLs and mobile final URLs in the account across ads,
 * keywords, and sitelinks, and checks their response code. Does not check
 * previously checked URLs.
 *
 * @param {Object} checkedUrls A map of previously checked URLs with the URL as
 *     the key.
 * @param {Array.<Object>} urlChecks An array into which the results of each URL
 *     check will be inserted.
 * @param {Object} options Dictionary of options.
 * @return {boolean} True if all URLs were checked.
 */
function checkUrls(checkedUrls, urlChecks, options) {
  var didComplete = true;

  // Helper method to add common conditions to ad group and keyword selectors.
  var addConditions = function(selector, includePaused) {
    var statuses = ['ENABLED'];
    if (includePaused) {
      statuses.push('PAUSED');
    }

    var predicate = ' IN [' + statuses.join(',') + ']';
    return selector.withCondition(labelCondition(false)).
        withCondition('Status' + predicate).
        withCondition('CampaignStatus' + predicate).
        withCondition('AdGroupStatus' + predicate);
  };

  if (options.checkAdUrls) {
    didComplete = didComplete && checkUrlsBySelector(checkedUrls, urlChecks,
        addConditions(AdWordsApp.ads().withCondition('CreativeFinalUrls != ""'),
                      options.checkPausedAds));
  }

  if (options.checkKeywordUrls) {
    didComplete = didComplete && checkUrlsBySelector(checkedUrls, urlChecks,
        addConditions(AdWordsApp.keywords().withCondition('FinalUrls != ""'),
                      options.checkPausedKeywords));
  }

  if (options.checkSitelinkUrls) {
    didComplete = didComplete &&
        checkSitelinkUrls(checkedUrls, urlChecks, options);
  }

  return didComplete;
}

/**
 * Retrieves all final URLs and mobile final URLs in a selector and checks them
 * for a valid response code. Does not check previously checked URLs. Labels the
 * entity that it was checked, if possible.
 *
 * @param {Object} checkedUrls A map of previously checked URLs with the URL as
 *     the key.
 * @param {Array.<Object>} urlChecks An array into which the results of each URL
 *     check will be inserted.
 * @param {Object} selector The selector specifying the entities to use.
 *     The entities should be of a type that has a urls() method.
 * @return {boolean} True if all URLs were checked.
 */
function checkUrlsBySelector(checkedUrls, urlChecks, selector) {
  var customerId = AdWordsApp.currentAccount().getCustomerId();
  var iterator = selector.get();
  var entities = [];

  // Helper method to check a URL.
  var checkUrl = function(entity, url) {
    if (!url) {
      return;
    }

    var urlsToCheck = expandUrlModifiers(url);

    for (var i = 0; i < urlsToCheck.length; i++) {
      var expandedUrl = urlsToCheck[i];
      if (checkedUrls[expandedUrl]) {
        continue;
      }

      var responseCode = requestUrl(expandedUrl);
      var entityType = entity.getEntityType();

      urlChecks.push({
        customerId: customerId,
        timestamp: new Date(),
        url: expandedUrl,
        responseCode: responseCode,
        entityType: entityType,
        campaign: entity.getCampaign ? entity.getCampaign().getName() : '',
        adGroup: entity.getAdGroup ? entity.getAdGroup().getName() : '',
        ad: entityType == 'Ad' ? getAdAsText(entity) : '',
        keyword: entityType == 'Keyword' ? entity.getText() : '',
        sitelink: entityType.indexOf('Sitelink') != -1 ?
            entity.getLinkText() : ''
      });

      checkedUrls[expandedUrl] = true;
    }
  };

  while (iterator.hasNext()) {
    entities.push(iterator.next());
  }

  for (var i = 0; i < entities.length; i++) {
    var entity = entities[i];

    checkUrl(entity, entity.urls().getFinalUrl());
    checkUrl(entity, entity.urls().getMobileFinalUrl());

    // Sitelinks do not have labels.
    if (entity.applyLabel) {
      entity.applyLabel(CONFIG.LABEL);
      checkTimeout();
    }
  }

  // True only if we did not breach an iterator limit.
  return entities.length == iterator.totalNumEntities();
}

/**
 * Retrieves a text representation of an ad, casting the ad to the appropriate
 * type if necessary.
 *
 * @param {Ad} ad The ad object.
 * @return {string} The text representation.
 */
function getAdAsText(ad) {
  // There is no AdTypeSpace method for textAd
  if (ad.getType() === 'TEXT_AD') {
    return ad.getHeadline();
  } else if (ad.isType().expandedTextAd()) {
    var eta = ad.asType().expandedTextAd();
    return eta.getHeadlinePart1() + ' - ' + eta.getHeadlinePart2();
  } else if (ad.isType().gmailImageAd()) {
    return ad.asType().gmailImageAd().getName();
  } else if (ad.isType().gmailMultiProductAd()) {
    return ad.asType().gmailMultiProductAd().getHeadline();
  } else if (ad.isType().gmailSinglePromotionAd()) {
    return ad.asType().gmailSinglePromotionAd().getHeadline();
  } else if (ad.isType().html5Ad()) {
    return ad.asType().html5Ad().getName();
  } else if (ad.isType().imageAd()) {
    return ad.asType().imageAd().getName();
  } else if (ad.isType().responsiveDisplayAd()) {
    return ad.asType().responsiveDisplayAd().getLongHeadline();
  }
  return 'N/A';
}

/**
 * Retrieves all final URLs and mobile final URLs for campaign and ad group
 * sitelinks and checks them for a valid response code. Does not check
 * previously checked URLs. Labels the containing campaign or ad group that it
 * has been checked.
 *
 * @param {Object} checkedUrls A map of previously checked URLs with the URL as
 *     the key.
 * @param {Array.<Object>} urlChecks An array into which the results of each URL
 *     check will be inserted.
 * @param {Object} options Dictionary of options.
 * @return {boolean} True if all URLs were checked.
 */
function checkSitelinkUrls(checkedUrls, urlChecks, options) {
  var didComplete = true;

  // Helper method to check URLs for sitelinks in a campaign or ad group
  // selector.
  var checkSitelinkUrls = function(selector) {
    var iterator = selector.withCondition(labelCondition(false)).get();
    var entities = [];

    while (iterator.hasNext()) {
      entities.push(iterator.next());
    }

    for (var i = 0; i < entities.length; i++) {
      var entity = entities[i];
      var sitelinks = entity.extensions().sitelinks();

      if (sitelinks.get().hasNext()) {
        didComplete = didComplete &&
            checkUrlsBySelector(checkedUrls, urlChecks, sitelinks);
        entity.applyLabel(CONFIG.LABEL);
        checkTimeout();
      }
    }

    // True only if we did not breach an iterator limit.
    didComplete = didComplete &&
        entities.length == iterator.totalNumEntities();
  };

  var statuses = ['ENABLED'];
  if (options.checkPausedSitelinks) {
    statuses.push('PAUSED');
  }

  var predicate = ' IN [' + statuses.join(',') + ']';
  checkSitelinkUrls(AdWordsApp.campaigns().
                    withCondition('Status' + predicate));
  checkSitelinkUrls(AdWordsApp.adGroups().
                    withCondition('Status' + predicate).
                    withCondition('CampaignStatus' + predicate));

  return didComplete;
}

/**
 * Expands a URL that contains ValueTrack parameters such as {ifmobile:mobile}
 * to all the combinations, and returns as an array. The following pairs of
 * ValueTrack parameters are currently expanded:
 *     1. {ifmobile:<...>} and {ifnotmobile:<...>} to produce URLs simulating
 *        clicks from either mobile or non-mobile devices.
 *     2. {ifsearch:<...>} and {ifcontent:<...>} to produce URLs simulating
 *        clicks on either the search or display networks.
 * Any other ValueTrack parameters or customer parameters are stripped out from
 * the URL entirely.
 *
 * @param {string} url The URL which may contain ValueTrack parameters.
 * @return {!Array.<string>} An array of one or more expanded URLs.
 */
function expandUrlModifiers(url) {
  var ifRegex = /({(if\w+):([^}]+)})/gi;
  var modifiers = {};
  var matches;
  while (matches = ifRegex.exec(url)) {
    // Tags are case-insensitive, e.g. IfMobile is valid.
    modifiers[matches[2].toLowerCase()] = {
      substitute: matches[0],
      replacement: matches[3]
    };
  }
  if (Object.keys(modifiers).length) {
    if (modifiers.ifmobile || modifiers.ifnotmobile) {
      var mobileCombinations =
          pairedUrlModifierReplace(modifiers, 'ifmobile', 'ifnotmobile', url);
    } else {
      var mobileCombinations = [url];
    }

    // Store in a map on the offchance that there are duplicates.
    var combinations = {};
    mobileCombinations.forEach(function(url) {
      if (modifiers.ifsearch || modifiers.ifcontent) {
        pairedUrlModifierReplace(modifiers, 'ifsearch', 'ifcontent', url)
            .forEach(function(modifiedUrl) {
              combinations[modifiedUrl] = true;
            });
      } else {
        combinations[url] = true;
      }
    });
    var modifiedUrls = Object.keys(combinations);
  } else {
    var modifiedUrls = [url];
  }
  // Remove any custom parameters
  return modifiedUrls.map(function(url) {
    return url.replace(/{[0-9a-zA-Z\_\+\:]+}/g, '');
  });
}

/**
 * Return a pair of URLs, where each of the two modifiers is mutually exclusive,
 * one for each combination. e.g. Evaluating ifmobile and ifnotmobile for a
 * mobile and a non-mobile scenario.
 *
 * @param {Object} modifiers A map of ValueTrack modifiers.
 * @param {string} modifier1 The modifier to honour in the URL.
 * @param {string} modifier2 The modifier to remove from the URL.
 * @param {string} url The URL potentially containing ValueTrack parameters.
 * @return {Array.<string>} A pair of URLs, as a list.
 */
function pairedUrlModifierReplace(modifiers, modifier1, modifier2, url) {
  return [
    urlModifierReplace(modifiers, modifier1, modifier2, url),
    urlModifierReplace(modifiers, modifier2, modifier1, url)
  ];
}

/**
 * Produces a URL where the first {if...} modifier is set, and the second is
 * deleted.
 *
 * @param {Object} mods A map of ValueTrack modifiers.
 * @param {string} mod1 The modifier to honour in the URL.
 * @param {string} mod2 The modifier to remove from the URL.
 * @param {string} url The URL potentially containing ValueTrack parameters.
 * @return {string} The resulting URL with substitions.
 */
function urlModifierReplace(mods, mod1, mod2, url) {
  var modUrl = mods[mod1] ?
      url.replace(mods[mod1].substitute, mods[mod1].replacement) :
      url;
  return mods[mod2] ? modUrl.replace(mods[mod2].substitute, '') : modUrl;
}

/**
 * Requests a given URL. Retries if the UrlFetchApp QPS limit was reached,
 * exponentially backing off on each retry. Throws an exception if it reaches
 * the maximum number of retries. Throws an exception if the UrlFetchApp daily
 * quota limit was reached.
 *
 * @param {string} url The URL to test.
 * @return {number|string} The response code received when requesting the URL,
 *     or an error message.
 */
function requestUrl(url) {
  var responseCode;
  var sleepTime = QUOTA_CONFIG.INIT_SLEEP_TIME;
  var numTries = 0;

  while (numTries < QUOTA_CONFIG.MAX_TRIES && !responseCode) {
    try {
      // If UrlFetchApp.fetch() throws an exception, responseCode will remain
      // undefined.
      responseCode =
        UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getResponseCode();

      if (CONFIG.THROTTLE > 0) {
        Utilities.sleep(CONFIG.THROTTLE);
      }
    } catch(e) {
      if (e.message.indexOf('Service invoked too many times in a short time:')
          != -1) {
        Utilities.sleep(sleepTime);
        sleepTime *= QUOTA_CONFIG.BACKOFF_FACTOR;
      } else if (e.message.indexOf('Service invoked too many times:') != -1) {
        throw EXCEPTIONS.LIMIT;
      } else {
        return e.message;
      }
    }

    numTries++;
  }

  if (!responseCode) {
    throw EXCEPTIONS.QPS;
  } else {
    return responseCode;
  }
}

/**
 * Throws an exception if the script is close to timing out.
 */
function checkTimeout() {
  if (AdWordsApp.getExecutionInfo().getRemainingTime() <
      CONFIG.TIMEOUT_BUFFER) {
    throw EXCEPTIONS.TIMEOUT;
  }
}

/**
 * Returns the number of days between two dates.
 *
 * @param {Object} from The older Date object.
 * @param {Object} to The newer (more recent) Date object.
 * @return {number} The number of days between the given dates (possibly
 *     fractional).
 */
function dayDifference(from, to) {
  return (to.getTime() - from.getTime()) / (24 * 3600 * 1000);
}

/**
 * Builds a string to be used for withCondition() filtering for whether the
 * label is present or not.
 *
 * @param {boolean} hasLabel True if the label should be present, false if the
 *     label should not be present.
 * @return {string} A condition that can be used in withCondition().
 */
function labelCondition(hasLabel) {
  return 'LabelNames ' + (hasLabel ? 'CONTAINS_ANY' : 'CONTAINS_NONE') +
      ' ["' + CONFIG.LABEL + '"]';
}

/**
 * Retrieves an entity by name.
 *
 * @param {Object} selector A selector for an entity type with a Name field.
 * @param {string} name The name to retrieve the entity by.
 * @return {Object} The entity, if it exists, or null otherwise.
 */
function getEntityByName(selector, name) {
  var entities = selector.withCondition('Name = "' + name + '"').get();

  if (entities.hasNext()) {
    return entities.next();
  } else {
    return null;
  }
}

/**
 * Retrieves a Label object by name.
 *
 * @param {string} labelName The label name to retrieve.
 * @return {Object} The Label object, if it exists, or null otherwise.
 */
function getLabel(labelName) {
  return getEntityByName(AdWordsApp.labels(), labelName);
}

/**
 * Retrieves an AccountLabel object by name.
 *
 * @param {string} labelName The label name to retrieve.
 * @return {Object} The AccountLabel object, if it exists, or null otherwise.
 */
function getAccountLabel(labelName) {
  return getEntityByName(MccApp.accountLabels(), labelName);
}

/**
 * Checks that the account has all provided labels and creates any that are
 * missing. Since labels cannot be created in preview mode, throws an exception
 * if a label is missing.
 *
 * @param {Array.<string>} labelNames An array of label names.
 */
function ensureLabels(labelNames) {
  for (var i = 0; i < labelNames.length; i++) {
    var labelName = labelNames[i];
    var label = getLabel(labelName);

    if (!label) {
      if (!AdWordsApp.getExecutionInfo().isPreview()) {
        AdWordsApp.createLabel(labelName);
      } else {
        throw 'Label ' + labelName + ' is missing and cannot be created in ' +
            'preview mode. Please run the script or create the label manually.';
      }
    }
  }
}

/**
 * Checks that the account has all provided account labels and creates any that
 * are missing. Since labels cannot be created in preview mode, throws an
 * exception if a label is missing.
 *
 * @param {Array.<string>} labelNames An array of label names.
 */
function ensureAccountLabels(labelNames) {
  for (var i = 0; i < labelNames.length; i++) {
    var labelName = labelNames[i];
    var label = getAccountLabel(labelName);

    if (!label) {
      if (!AdWordsApp.getExecutionInfo().isPreview()) {
        MccApp.createAccountLabel(labelName);
      } else {
        throw 'Account label ' + labelName + ' is missing and cannot be ' +
            'created in preview mode. Please run the script or create the ' +
            'label manually.';
      }
    }
  }
}

/**
 * Removes all provided labels from the account. Since labels cannot be removed
 * in preview mode, throws an exception in preview mode.
 *
 * @param {Array.<string>} labelNames An array of label names.
 */
function removeLabels(labelNames) {
  if (AdWordsApp.getExecutionInfo().isPreview()) {
    throw 'Cannot remove labels in preview mode. Please run the script or ' +
        'remove the labels manually.';
  }

  for (var i = 0; i < labelNames.length; i++) {
    var label = getLabel(labelNames[i]);

    if (label) {
      label.remove();
    }
  }
}

/**
 * Removes all provided account labels from the account. Since labels cannot be
 * removed in preview mode, throws an exception in preview mode.
 *
 * @param {Array.<string>} labelNames An array of label names.
 */
function removeAccountLabels(labelNames) {
  if (AdWordsApp.getExecutionInfo().isPreview()) {
    throw 'Cannot remove account labels in preview mode. Please run the ' +
        'script or remove the labels manually.';
  }

  for (var i = 0; i < labelNames.length; i++) {
    var label = getAccountLabel(labelNames[i]);

    if (label) {
      label.remove();
    }
  }
}

/**
 * Validates the provided spreadsheet URL to make sure that it's set up
 * properly. Throws a descriptive error message if validation fails.
 *
 * @param {string} spreadsheeturl The URL of the spreadsheet to open.
 * @return {Spreadsheet} The spreadsheet object itself, fetched from the URL.
 * @throws {Error} If the spreadsheet URL hasn't been set
 */
function validateAndGetSpreadsheet(spreadsheeturl) {
  if (spreadsheeturl == 'YOUR_SPREADSHEET_URL') {
    throw new Error('Please specify a valid Spreadsheet URL. You can find' +
        ' a link to a template in the associated guide for this script.');
  }
  return SpreadsheetApp.openByUrl(spreadsheeturl);
}

/**
 * Validates the provided email address to make sure it's not the default.
 * Throws a descriptive error message if validation fails.
 *
 * @throws {Error} If the list of email addresses is still the default
 */
function validateEmailAddresses() {
  if (CONFIG.RECIPIENT_EMAILS &&
      CONFIG.RECIPIENT_EMAILS[0] == 'YOUR_EMAIL_HERE') {
    throw new Error('Please either specify a valid email address or clear' +
        ' the RECIPIENT_EMAILS field.');
  }
}

