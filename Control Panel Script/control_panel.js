MASTER_SHEET = "1EMwldjMQ_TWb8i0qGcpyUN2n31O4lmZK2sdw38ecdCY";

function updateForm() {
	log("running updateForm()");
	var spr = SpreadsheetApp.getActiveSpreadsheet();
	var tournamentURL = spr.getRange("B1").getValue();
	var teamSet = getTeamsFromURL(tournamentURL);
	var teams = Array.from(teamSet).sort();
	var formURL = spr.getRange("B2").getValue();
	var form = FormApp.openByUrl(formURL);
	var listItems = form.getItems(FormApp.ItemType.LIST);
	var yourTeamItem = listItems[0].asListItem();
	var opponentTeamItem = listItems[1].asListItem();

	setListItemChoices(yourTeamItem, teams);
	setListItemChoices(opponentTeamItem, teams);

	spr.getRange("B13").setValue(formatDate(new Date(Date.now())));
	log("updateForm() success!");
}

function onFormSubmit(e) {
	log("running onFormSubmit()");
	var masterSheet = SpreadsheetApp.openById(MASTER_SHEET);
	SpreadsheetApp.setActiveSpreadsheet(masterSheet);
	var masterRawScoresSheet = masterSheet.getSheetByName("Raw Scores");
	copyResponseToSheet(e.response, masterRawScoresSheet);
	log("onFormSubmit() success!");
}

function initializeForm() {
	log("running initializeForm()");
	var spr = SpreadsheetApp.getActiveSpreadsheet();
	var formURL = spr.getRange("B2").getValue();
	var form = FormApp.openByUrl(formURL);
	masterSheet = DriveApp.getFileById(MASTER_SHEET);
	createFormSubmitTrigger(form);
	linkSheetToForm(form, spr);
	spr.getRange("B14").setValue(formatDate(new Date(Date.now())));
	log("initializeForm() success!");
}

function createFormSubmitTrigger(form) {
	var triggers = ScriptApp.getProjectTriggers();
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT && triggers[i].getHandlerFunction() == 'onFormSubmit') {
			return;
		}
	}
	ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create();
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