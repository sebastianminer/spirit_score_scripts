var errno = 0;

function updateForm() {
	log("running updateForm()");
	var spr = SpreadsheetApp.getActiveSpreadsheet();
	var tournamentURL = spr.getRange("B1").getValue();
	var teamSet = getTeamsFromURL(tournamentURL);
	var teams = Array.from(teamSet).sort();
	var rawScoreSheet = spr.getSheetByName('Raw Scores');
	var form = FormApp.openByUrl(rawScoreSheet.getFormUrl());
	var listItems = form.getItems(FormApp.ItemType.LIST);
	var yourTeamItem = listItems[0].asListItem();
	var opponentTeamItem = listItems[1].asListItem();

	setListItemChoices(yourTeamItem, teams);
	setListItemChoices(opponentTeamItem, teams);

	spr.getRange("A13").setValue("Teams last updated on form:");
	spr.getRange("B13").setValue(formatDate(new Date(Date.now())));
	log("updateForm() success!");
}

function initializeForm() {
	log("running initializeForm()");
	var spr = SpreadsheetApp.getActiveSpreadsheet();
	shareSpreadsheet(spr);
	var formURL = spr.getRange("B2").getValue();
	var form = FormApp.openByUrl(formURL);
	linkSheetToForm(form, spr);

	var bankUrl = spr.getRange("B3").getValue();
	linkSheetToBank(bankUrl, spr);

	if (errno) {
		log("initializeForm() completed with one or more errors.");
	} else {
		log("initializeForm() success!");
	}
}

function linkSheetToBank(bankUrl, spr) {
	var bankSheet = SpreadsheetApp.openByUrl(bankUrl).getSheetByName("Tournament Control Panel IDs");
	var tournamentSheetId = spr.getId();
	if (sheetIdInBank(tournamentSheetId, bankSheet)) {
		return;
	}
	var row = getFirstEmptyRow(bankSheet);
	bankSheet.getRange(row, 1).setValue(tournamentSheetId);
}

function sheetIdInBank(tournamentSheetId, bankSheet) {
	var existingIds = new Set(bankSheet.getRange("A:A").getValues().map(row => row[0]).filter(id => id.trim() != ""));
	return existingIds.has(tournamentSheetId);
}

function shareSpreadsheet(spr) {
	var emails = getEmailAddresses(spr);
	var spreadsheetId = spr.getId();
	for (email of emails) {
		try {
			DriveApp.getFileById(spreadsheetId).addViewer(email);
			log(`File successfully shared with ${email}.`);
		} catch (e) {
			log(`Error sharing with ${email}. ${e.name}: ${e.message}`);
			errno |= 1;
		}
	}
}

function getEmailAddresses(spr) {
	var controlPanelSheet = spr.getSheetByName("Control Panel");
	var values = controlPanelSheet.getRange("E:E").getValues().map(row => row[0]).slice(1);
	var emails = values.filter(value => value && value.trim() != "");
	return emails;
}

function linkSheetToForm(form, spr) {
	var formDestId;
	try {
		formDestId = form.getDestinationId();
	}
	catch (e) {
		formDestId = null;
	}
	if (formDestId != spr.getId()) {
		form.setDestination(FormApp.DestinationType.SPREADSHEET, spr.getId());
	}
	SpreadsheetApp.flush();
	for (sheet of spr.getSheets()) {
		if (sheet.getFormUrl() != null) {
			sheet.setName("Raw Scores");
			break;
		}
	}
}

function copyResponseToSheet(response, sheet) {
	var itemResponses = response.getItemResponses();
	var responseLength = itemResponses.length;
	var rawResponses = itemResponses.map(item => item.getResponse());
	var firstEmptyRow = getFirstEmptyRow(sheet);
	var range = `R${firstEmptyRow}C1:R${firstEmptyRow}C${responseLength+1}`;
	sheet.getRange(range).setValues([[response.getTimestamp(), ...rawResponses]]);
}

function getFirstEmptyRow(sheet) {
	var column = sheet.getRange('A:A');
	var values = column.getValues();
	var ct = 0;
	while (values[ct] && values[ct][0] != "") {
		ct++;
	}
	return (ct+1);
}

function setListItemChoices(listItem, arr) {
	listItem.setChoices(arr.map(a => listItem.createChoice(a)));
}

function getItemByTitle(title, form) {
	var items = form.getItems();
	for (var item of items) {
		if (item.getTitle() == title) {
			return item;
		}
	}
	return null;
}

function getTeamsFromDoc(doc_id) {
	var doc = DocumentApp.openById(doc_id);
	var html = doc.getBody().getText();
	return getTeamsFromString(html);
}

function getTeamsFromURL(url) {
	var http = UrlFetchApp.fetch(url);
	var html = http.getContentText();
	return getTeamsFromString(html);
}

function getTeamsFromString(s) {
	var teams = new Set();
	var lines = s.split("\n");
	for (var line of lines) {
		var match = line.match(/<a href=\"\/events\/teams\/\?.*>(.*)\([0-9]+\)<\/a>/);
		if (match && match.length) {
			team = match[1];
			teams.add(team);
		}
	}
	return teams;
}

function formatDate(date) {
	var month = String(date.getMonth() + 1).padStart(2, '0');
	var dateInMonth = String(date.getDate()).padStart(2, '0');
	var year = String((date.getYear()) + 1900).padStart(2, '0');
	var hours = String(date.getHours()).padStart(2, '0');
	var minutes = String(date.getMinutes()).padStart(2, '0');
	var seconds = String(date.getSeconds()).padStart(2, '0');
	return `${month}/${dateInMonth}/${year} ${hours}:${minutes}:${seconds}`;
}

function log(obj, omitDate) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
	var range = sheet.getRange("A1");
	var cellContents = range.getValue();
	var now = new Date(Date.now());
	var timeStamp = `[${formatDate(now)}]`;
	var cellContents = `${omitDate ? "" : timeStamp} ${String(obj)}\n${cellContents}`;
	range.setValue(cellContents);
}