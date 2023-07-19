/**
 * @name Ad Label Scheduler
 *
 * @overview This scripts enables/disables ads with assigned labels based on the specific start/end dates.
 * Updates are reflected on the Google Sheet and changes are notified on the Slack channel.
 *
 * @author Said Tezel [saidtezel@gmail.com]
 *
 * @version 1.0
 *
**/

var CONFIG = {
  // Copy the template here https://bit.ly/2LRuDa3 and change the spreadsheet URL below
  SPREADSHEET_URL: 'SPREADSHEET_URL',

  // Slack webhook endpoint to talk to the #ppc channel
  SLACK_ENDPOINT: 'SLACK_ENDPOINT',

  // Set true if you want to send Slack notifications for recent changes.
  UPDATE_SLACK: true,
};

var REPORTING_OPTIONS = {
  apiVersion: 'v201809'
};


function main() {
  var ss = initialiseSpreadsheet();
  var timezone = ss.getSpreadsheetTimeZone();
  var range = ss.getRangeByName('dataRange');

  var scanQueue = {};
  var recentChanges = [];
  var rangeVal = range.getValues();

  // Grouping the executions for each account ID.
  for (var k = 0; k < rangeVal.length; k++) {
    var rowVal = rangeVal[k];

    if (rowVal[0] == '' || rowVal[1] == '') {
      continue
    };

    var accountId = rowVal[1];

    if (!scanQueue[accountId]) {
      scanQueue[accountId] = [];
    }

    scanQueue[accountId].push({
      label: rowVal[0],
      start: new Date(rowVal[3]),
      end: rowVal[4] == '' ? new Date('2099-01-01') : new Date(rowVal[4]),
      lastStatus: rowVal[5],
      lastUpdated: rowVal[6],
      rowIndex: k + 1
    });
  }

  var accountIds = Object.keys(scanQueue);

  for (var j = 0; j < accountIds.length; j++) {
    var accountId = accountIds[j];
    var account = getAccountWithId(accountId);
    var accountName = account.getName();
    AdsManagerApp.select(account);

    Logger.log('Processing label updates for ' + accountName);
    var labels = scanQueue[accountId];

    for (var m = 0; m < labels.length; m++) {
      var labelData = labels[m];
      Logger.log('***');
      Logger.log('Processing label ' + labelData.label);
      var labelDataIsValid = validateLabelData(labelData);
      var labelNeedsUpdate = shouldLabelUpdate(labelData);

      if (labelDataIsValid && labelNeedsUpdate) {
        Logger.log('Status change detected. Updating ads...');
        var updateStatus = processLabelUpdates(labelData);

        range.getCell(labelData.rowIndex, 6).setValue(updateStatus);
        range.getCell(labelData.rowIndex, 7).setValue(Utilities.formatDate(new Date(), timezone, 'YYYY-MM-dd HH:mm'))
        Logger.log('Successfully changed the status of ads to ' + updateStatus);

        recentChanges.push({
          label: labelData.label,
          accountName: accountName,
          status: updateStatus
        });
      }

    }
  }

  var updateFinishTime = new Date();
  var updateFinishTimeLocal = Utilities.formatDate(updateFinishTime, timezone, 'YYYY-MM-dd HH:mm');
  var updateRange = ss.getRangeByName('updateDate');
  updateRange.setValue(updateFinishTimeLocal);

  if (recentChanges.length > 0 && CONFIG.UPDATE_SLACK) {
    sendSlackUpdate(recentChanges, updateFinishTimeLocal);
  }
}

function processLabelUpdates(labelData) {
  var status;
  var start = labelData.start;
  var end = labelData.end;
  var now = new Date();
  var label = labelData.label

  if (now >= start && now < end) {
    status = 'ACTIVE';
  } else if (now < start) {
    status = 'QUEUED';
  } else {
    status = 'INACTIVE';
  }

  var ads = getAdsWithLabel(label);

  while (ads.hasNext()) {
    var ad = ads.next();

    if (status == 'ACTIVE') {
      ad.enable();
    } else {
      ad.pause()
    }
  };

  return status;
}

function shouldLabelUpdate(labelData) {
  var now = new Date();
  var lastStatus = labelData.lastStatus;
  var currStatus;

  if (now >= labelData.start && now < labelData.end) {
    currStatus = 'ACTIVE'
  } else if (now < labelData.start) {
    currStatus = 'QUEUED'
  } else {
    currStatus = 'INACTIVE'
  }

  Logger.log('Last known status ' + lastStatus);
  Logger.log('Current status ' + currStatus);
  return !(currStatus == lastStatus);
}

function validateLabelData(labelData) {
  var label = labelData.label;

  if (!checkIfLabelExists(label)) {
    Logger.log('Error: Label doesn\'n exist.');
    return false
  }

  if (isNaN(labelData.start) || isNaN(labelData.end) || labelData.end <= labelData.start) {
    Logger.log('Error: Input dates are invalid.');
    return false
  }

  return true;
}



/**
 * Initialises the spreadsheet from the config URL.
 *
 * @return {Object} Spreadsheet instance
**/
function initialiseSpreadsheet() {
  var ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  return ss;
}

/**
 * Checks if the current timestamp is within the date range.
 *
 * @param {string} activation date on row
 * @param {string} deactivation date on row
 * @return {boolean} true if date is in range
 */
function checkIfDateInRange(start, end) {
  var now = new Date()
  var startDate = new Date(start);
  var endDate = new Date(end);

  if (now > startDate && (end === '' || now < endDate)) {
    return true
  } else {
    return false
  }
}

/**
 * Checks if the label assigned exists within the ad accont.
 *
 * @param {string} label name
 * @return {boolean} true if label exists
 */
function checkIfLabelExists(label) {
  var labelSelector = AdsApp.labels()
    .withCondition("Name CONTAINS '" + label + "'")
    .get()

  if (!labelSelector.hasNext()) {
    return false
  }

  return true
}

/**
 * Finds and returns the ad account with ID.
 *
 * @param {string} account ID
 * @return {Object} ad account
 */
function getAccountWithId(id) {
  var accountSelector = AdsManagerApp.accounts()
    .withIds([id])
    .get();

  if (accountSelector.hasNext()) {
    return accountSelector.next();
  }

  return null
}

/**
 * Finds and returns an ad selector with a specific label
 *
 * @param {string} label name
 * @return {Object} ad selector
 */
function getAdsWithLabel(label) {
  var adSelector = AdsApp.ads()
    .withCondition("LabelNames CONTAINS_ANY ['" + label + "']")
    .get()

  if (!adSelector.hasNext()) {
    return null;
  }

  return adSelector;
}

/**
 * Sends a notification to the configured Slack webhook endpoint
 * with the latest changes on account ads
 *
 * @param {Array.<Object>} An object array of latest changes on accounts.
**/
function sendSlackUpdate(results, updateTime) {
  var slackMessage = {
    blocks: [
	    {
		    "type": "section",
		    "text": {
			    "type": "mrkdwn",
			    "text": "I've just mades some changes on your ad accounts based on ad labels."
		    }
	    },
	    {
		    "type": "divider"
	    },
	    {
		    "type": "section",
		    "text": {
			    "type": "mrkdwn",
			    "text": "*<https://docs.google.com/spreadsheets/d/1B4xrSdDQH5Wcblc_UwBHEQ-Aty9niavwPfe7K_4ct4M/edit#gid=0|Time Based Ad Copy Activator - Google Sheets>*\nActivate/deactivate ads on your accounts based on date ranges."
		    },
		    "accessory": {
			    "type": "image",
			    "image_url": "https://api.slack.com/img/blocks/bkb_template_images/approvalsNewDevice.png",
			    "alt_text": "computer icon"
		    }
	    },
	    {
		    "type": "context",
		    "elements": [
			    {
				    "type": "image",
				    "image_url": "https://api.slack.com/img/blocks/bkb_template_images/notificationsWarningIcon.png",
				    "alt_text": "notifications warning icon"
			    },
			    {
				    "type": "mrkdwn",
				    "text": "*Last scan completed on " + updateTime + "*"
			    }
		    ]
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "*Here are the latest changes:*"
        }
	    },
    ]
  };

  for (var j = 0; j < results.length; j++) {
    slackMessage.blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: "*" + results[j].accountName + "*\nStatus of ads with _" + results[j].label + "_ label changed to *" + results[j].status.toLowerCase() + "*."
      }
    });
  };

  slackMessage.blocks.push({
      "type": "section",
      "text": {
        "type": "mrkdwn",
        "text": "*<https://docs.google.com/spreadsheets/d/1B4xrSdDQH5Wcblc_UwBHEQ-Aty9niavwPfe7K_4ct4M/edit#gid=0|See all changes...>*"
      }
  });


  var options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(slackMessage)
  };

  UrlFetchApp.fetch(CONFIG.SLACK_ENDPOINT, options);
}


