/**
 * @name Campaign label based ad scheduling
 *
 * @overview This script is a heavily modified version of the day parting script provided by brainlabsdigital.com
* It provides an MCC version of the script that heavily depeneds on the labels applied to the campaigns.
*
* It will apply ad schedules to campaigns or shopping campaigns and set
* the ad schedule bid modifier and mobile bid modifier at each hour according to
* multiplier timetables in a Google sheet.
*
* This version creates schedules with modifiers for 4 hours, then fills the rest
* of the day and the other days of the week with schedules with no modifier as a
* fail safe.
 *
 * @author Said Tezel [saidtezel@gmail.com]
 *
 * @version 1.0
 *
**/

function main() {
	var spreadsheetUrl = "INSERT GOOGLE SHEET URL";

	// Define the label name to be used for filtering accounts.
	// You'll need to apply a label with this value to all MCC accounts you want to run this script, with a limit of 50.
	// The 50 account limit is because parallel processing only allows up to 50 accounts at once.
	var labelName = "dayPartingEnabled";

	// Shopping or regular campaigns?
	// Use true if you want to run script on shopping campaigns (not regular campaigns).
	// Use false for regular campaigns.
	var shoppingCampaigns = false;

	// Optional parameters for filtering campaign names. The matching is case insensitive.
	var excludeCampaignNameContains = [];

	// Select which campaigns to include e.g ["foo", "bar"] will include only campaigns
	var includeCampaignNameContains = [];

	// When you want to stop running the ad scheduling for good, set the lastRun
	// variable to true to remove all ad schedules.
	var lastRun = false;

	// Initialise for use later.
	var weekDays = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
	var adScheduleCodes = [];

	var scheduleRange = "B2:H25";
	var mobileScheduleRange = "M2:S25";

	// Initiate shared parameters to be used across all accounts in parallel processing
	var sharedParams = {
		"spreadsheetUrl": spreadsheetUrl,
		"shoppingCampaigns": shoppingCampaigns,
		"excludeCampaignNameContains": excludeCampaignNameContains,
		"includeCampaignNameContains": includeCampaignNameContains,
		"lastRun": lastRun,
		"weekDays": weekDays,
		"adScheduleCodes": adScheduleCodes,
		"scheduleRange": scheduleRange,
		"mobileScheduleRange": mobileScheduleRange,
	};

	var accountSelector = AdsManagerApp.accounts()
		.withLimit(50)
		.withCondition("LabelNames CONTAINS \"" + labelName + "\"");

	accountSelector.executeInParallel("processAllAccounts", "allFinished", JSON.stringify(sharedParams));
}

function processAllAccounts(sharedParams) {
	var sharedParams = JSON.parse(sharedParams);

	var spreadsheetUrl = sharedParams.spreadsheetUrl;
	var shoppingCampaigns = sharedParams.shoppingCampaigns;
	var excludeCampaignNameContains = sharedParams.excludeCampaignNameContains;
	var includeCampaignNameContains = sharedParams.includeCampaignNameContains;
	var lastRun = sharedParams.lastRun;
	var weekDays = sharedParams.weekDays;
	var adScheduleCodes = sharedParams.adScheduleCodes;
	var scheduleRange = sharedParams.scheduleRange;
	var mobileScheduleRange = sharedParams.mobileScheduleRange;

	//Retrieving up hourly data
	var timeZone = AdWordsApp.currentAccount().getTimeZone();
	if (timeZone === "Etc/GMT") {
		timeZone = "GMT";
	}
	var date = new Date();
	var dayOfWeek = parseInt(Utilities.formatDate(date, timeZone, "uu"), 10) - 1;
	var hour = parseInt(Utilities.formatDate(date, timeZone, "HH"), 10);

	// Keep track of total number of campaigns processed in this execution.
	var campaignCount = 0;

	var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
	var sheets = spreadsheet.getSheets();

	for (var s = 0; s < sheets.length; s++) {
		var sheet = sheets[s];
		var sheetName = sheets[s].getName();

		if (sheetName == "READ THIS FIRST" || "template") {
			continue;
		}
		var campaignIds = [];

		var data = sheet.getRange(scheduleRange).getValues();

		//This hour's bid multiplier.
		var thisHourMultiplier = data[hour][dayOfWeek];
		var lastHourCell = "I2";
		sheet.getRange(lastHourCell).setValue(thisHourMultiplier);

		//The next few hours' multipliers
		var timesAndModifiers = [];
		var otherDays = weekDays.slice(0);
		for (var h = 0; h < 5; h++) {
			var newHour = (hour + h) % 24;
			if (hour + h > 23) {
				var newDay = (dayOfWeek + 1) % 7;
			} else {
				var newDay = dayOfWeek;
			}
			otherDays[newDay] = "-";

			if (h < 4) {
				// Use the specified bids for the next 4 hours
				var bidModifier = data[newHour][newDay];
				if (isNaN(bidModifier) || (bidModifier < -0.9 && bidModifier > -1) || bidModifier > 9) {
					Logger.log("Bid modifier '" + bidModifier + "' for " + weekDays[newDay] + " " + newHour + " is not valid.");
					timesAndModifiers.push([newHour, newHour + 1, weekDays[newDay], 0]);
				} else if (bidModifier != -1 && bidModifier.length != 0) {
					timesAndModifiers.push([newHour, newHour + 1, weekDays[newDay], bidModifier]);
				}
			} else {
				// Fill in the rest of the day with no adjustment (as a back-up incase the script breaks)
				timesAndModifiers.push([newHour, 24, weekDays[newDay], 0]);
			}
		}

		// Iterate through campaigns with labels
		var labelIterator = AdsApp.labels()
			.withCondition("Name = \"" + sheetName + "\"")
			.get();

		if (labelIterator.hasNext()) {
			var label = labelIterator.next();
			var campaignIterator = label.campaigns()
				.withCondition("Status IN [ENABLED]")
				.get();

			while (campaignIterator.hasNext()) {
				campaignCount += 1;
				var campaign = campaignIterator.next();
				var campaignName = campaign.getName();
				var includeCampaign = false;
				if (includeCampaignNameContains.length === 0) {
					includeCampaign = true;
				}
				for (var i = 0; i < includeCampaignNameContains.length; i++) {
					var index = campaignName.toLowerCase().indexOf(includeCampaignNameContains[i].toLowerCase());
					if (index !== -1) {
						includeCampaign = true;
						break;
					}
				}
				if (includeCampaign) {
					var campaignId = campaign.getId();
					campaignIds.push(campaignId);
				}
			}
		}

		// Continue with next iteration if this one has 0 campaigns.
		if (campaignIds.length == 0) {
			continue;
		}

		//Remove all ad scheduling for the last run.
		if (lastRun) {
			checkAndRemoveAdSchedules(campaignIds, []);
			return;
		}

		// Change the mobile bid adjustment
		var mobileCheck = /_incMobile/;
		if (mobileCheck.test(sheetName)) {
			var data = sheet.getRange(mobileScheduleRange).getValues();

			var thisHourMultiplier_Mobile = data[hour][dayOfWeek];

			if (thisHourMultiplier_Mobile.length === 0) {
				thisHourMultiplier_Mobile = -1;
			}

			if (isNaN(thisHourMultiplier_Mobile) || (thisHourMultiplier_Mobile < -0.9 && thisHourMultiplier_Mobile > -1) || thisHourMultiplier_Mobile > 3) {
				Logger.log("Mobile bid modifier '" + thisHourMultiplier_Mobile + "' for " + weekDays[dayOfWeek] + " " + hour + " is not valid.");
				thisHourMultiplier_Mobile = 0;
			}

			var totalMultiplier = ((1 + thisHourMultiplier_Mobile) * (1 + thisHourMultiplier)) - 1;
			sheet.getRange("T2").setValue(thisHourMultiplier_Mobile);
			sheet.getRange("AE2").setValue(totalMultiplier);
			ModifyMobileBidAdjustment(campaignIds, thisHourMultiplier_Mobile);
		}

		// Check the existing ad schedules, removing those no longer necessary
		var existingSchedules = checkAndRemoveAdSchedules(campaignIds, timesAndModifiers);

		// Add in the new ad schedules
		AddHourlyAdSchedules(campaignIds, timesAndModifiers, existingSchedules, shoppingCampaigns);
	}

	return campaignCount.toFixed(0);
}

/**
 * Post-process the results from processAccount. This method will be called
 * once all the accounts have been processed by the executeInParallel method
 * call.
 *
 * @param {Array.<ExecutionResult>} results An array of ExecutionResult objects,
 * one for each account that was processed by the executeInParallel method.
 */
function allFinished(results) {
	var accountCount = results.length;
	var totalCampaignCount = 0;

	Logger.log("Parallel processing is complete");
	Logger.log("*************");
	for (var w = 0; w < results.length; w++) {
		// Get the ExecutionResult for an account.
		var result = results[w];

		Logger.log("Customer ID: %s; Status: %s.", result.getCustomerId(), result.getStatus());

		// Check the execution status. This can be one of ERROR, OK, or TIMEOUT.
		if (result.getStatus() == "ERROR") {
			Logger.log("-- Failed with error: '%s'.", result.getError());
		} else if (result.getStatus() == "OK") {
			var retval = parseInt(result.getReturnValue());
			totalCampaignCount += retval;
		} else {
			Logger.log("The execution was not able to finish due to timeout");
		}
	}

	Logger.log("*************");
	Logger.log("Total number of accounts processed: " + accountCount);
	Logger.log("Total number of campaigns processed: " + totalCampaignCount);
}

/**
* Function to add ad schedules for the campaigns with the given IDs, unless the schedules are
* referenced in the existingSchedules array. The scheduling will be added as a hour long periods
* as specified in the passed parameter array and will be given the specified bid modifier.
*
* @param array campaignIds array of campaign IDs to add ad schedules to
* @param array timesAndModifiers the array of [hour, day, bid modifier] for which to add ad scheduling
* @param array existingSchedules array of strings identifying already existing schedules.
* @param bool shoppingCampaigns using shopping campaigns?
* @return void
*/
function AddHourlyAdSchedules(campaignIds, timesAndModifiers, existingSchedules, shoppingCampaigns) {
	// times = [[hour,day],[hour,day]]
	var campaignIterator = ConstructIterator(shoppingCampaigns)
		.withIds(campaignIds)
		.get();
	while (campaignIterator.hasNext()) {
		var campaign = campaignIterator.next();
		for (var i = 0; i < timesAndModifiers.length; i++) {
			if (existingSchedules.indexOf(
				timesAndModifiers[i][0] + "|" + (timesAndModifiers[i][1]) + "|" + timesAndModifiers[i][2]
				+ "|" + Utilities.formatString("%.2f", (timesAndModifiers[i][3] + 1)) + "|" + campaign.getId())
				> -1) {
				continue;
			}

			campaign.addAdSchedule({
				dayOfWeek: timesAndModifiers[i][2],
				startHour: timesAndModifiers[i][0],
				startMinute: 0,
				endHour: timesAndModifiers[i][1],
				endMinute: 0,
				bidModifier: Math.round(100 * (1 + timesAndModifiers[i][3])) / 100
			});
		}
	}
}

/**
* Function to remove ad schedules from all campaigns referenced in the passed array
* which do not correspond to schedules specified in the passed timesAndModifiers array.
*
* @param array campaignIds array of campaign IDs to remove ad scheduling from
* @param array timesAndModifiers array of [hour, day, bid modifier] of the wanted schedules
* @return array existingWantedSchedules array of strings identifying the existing undeleted schedules
*/
function checkAndRemoveAdSchedules(campaignIds, timesAndModifiers) {

	var adScheduleIds = [];

	var report = AdWordsApp.report(
		"SELECT CampaignId, Id " +
		"FROM CAMPAIGN_AD_SCHEDULE_TARGET_REPORT " +
		"WHERE CampaignId IN [\"" + campaignIds.join("\",\"") + "\"]"
	);

	var rows = report.rows();
	while (rows.hasNext()) {
		var row = rows.next();
		var adScheduleId = row["Id"];
		var campaignId = row["CampaignId"];
		if (adScheduleId == "--") {
			continue;
		}
		adScheduleIds.push([campaignId, adScheduleId]);
	}

	var chunkedArray = [];
	var chunkSize = 10000;

	for (var i = 0; i < adScheduleIds.length; i += chunkSize) {
		chunkedArray.push(adScheduleIds.slice(i, i + chunkSize));
	}

	var wantedSchedules = [];
	var existingWantedSchedules = [];

	for (var j = 0; j < timesAndModifiers.length; j++) {
		wantedSchedules.push(timesAndModifiers[j][0] + "|" + (timesAndModifiers[j][1]) + "|" + timesAndModifiers[j][2] + "|" + Utilities.formatString("%.2f", timesAndModifiers[j][3] + 1));
	}

	for (var k = 0; k < chunkedArray.length; k++) {
		var unwantedSchedules = [];

		var adScheduleIterator = AdWordsApp.targeting()
			.adSchedules()
			.withIds(chunkedArray[k])
			.get();
		while (adScheduleIterator.hasNext()) {
			var adSchedule = adScheduleIterator.next();
			var key = adSchedule.getStartHour() + "|" + adSchedule.getEndHour() + "|" + adSchedule.getDayOfWeek() + "|" + Utilities.formatString("%.2f", adSchedule.getBidModifier());

			if (wantedSchedules.indexOf(key) > -1) {
				existingWantedSchedules.push(key + "|" + adSchedule.getCampaign().getId());
			} else {
				unwantedSchedules.push(adSchedule);
			}
		}

		for (var p = 0; p < unwantedSchedules.length; p++) {
			unwantedSchedules[j].remove();
		}
	}

	return existingWantedSchedules;
}

/**
* Function to construct an iterator for shopping campaigns or regular campaigns.
*
* @param bool shoppingCampaigns Using shopping campaigns?
* @return AdWords iterator Returns the corresponding AdWords iterator
*/
function ConstructIterator(shoppingCampaigns) {
	if (shoppingCampaigns === true) {
		return AdWordsApp.shoppingCampaigns();
	}
	else {
		return AdWordsApp.campaigns();
	}
}

/**
* Function to set a mobile bid modifier for a set of campaigns
*
* @param array campaignIds An array of the campaign IDs to be affected
* @param Float bidModifier The multiplicative mobile bid modifier
* @return void
*/
function ModifyMobileBidAdjustment(campaignIds, bidModifier) {
	var platformIds = [];
	var newBidModifier = Math.round(100 * (1 + bidModifier)) / 100;

	for (var q = 0; q < campaignIds.length; q++) {
		platformIds.push([campaignIds[q], 30001]);
	}

	var platformIterator = AdWordsApp.targeting()
		.platforms()
		.withIds(platformIds)
		.get();
	while (platformIterator.hasNext()) {
		var platform = platformIterator.next();
		platform.setBidModifier(newBidModifier);
	}
}
