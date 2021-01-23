MASTER_SHEET = "1EMwldjMQ_TWb8i0qGcpyUN2n31O4lmZK2sdw38ecdCY";

function onFormSubmit(e) {
	var masterSheet = SpreadsheetApp.openById(MASTER_SHEET);
	SpreadsheetApp.setActiveSpreadsheet(masterSheet);
	var rawScoresSheet = masterSheet.getSheetByName("Raw Scores");
	copyResponseToSheet(e.response);
}

function set_log(s) {
	out = "{ ";
	for (element of s) {
		out += String(element) + ", ";
	}
	out += " }";
	Logger.log(out);
}

function copyResponseToSheet(response) {
	var spr = SpreadsheetApp.getActiveSpreadsheet();
	var itemResponses = response.getItemResponses();
	var responseLength = itemResponses.length;
	var rawResponses = itemResponses.map(item => item.getResponse());
	var firstEmptyRow = getFirstEmptyRow();
	var range = `R${firstEmptyRow}C1:R${firstEmptyRow}C${responseLength+1}`;
	spr.getRange(range).setValues([[response.getTimestamp(), ...rawResponses]]);
}

function getFirstEmptyRow() {
	var spr = SpreadsheetApp.getActiveSpreadsheet();
	var column = spr.getRange('A:A');
	var values = column.getValues();
	var ct = 0;
	while (values[ct] && values[ct][0] != "") {
		ct++;
	}
	return (ct+1);
}