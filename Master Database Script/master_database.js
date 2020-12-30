const scoreKeys = ["rules", "fouls", "fairMind", "attitude", "communication", "total"];
const teamDataColumnHeadings = ["Team",
								"Number of Scores",
								"Total",
								"Rules Knowledge and Use",
								"Fouls and Body Contact",
								"Fair Mindedness",
								"Positive Attitude and Self-Control",
								"Communication",
								"Comments (Rules Knowledge and Use)",
								"Comments (Fouls and Body Contact)",
								"Comments (Fair Mindedness)",
								"Comments (Positive Attitude and Self-Control)",
								"Comments (Communication)",
								"Additional Comments"];

function updateMasterDatabase() {
	log("running updateMasterDatabase()");
	var controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control Panel");
	var rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Scores");
	var teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aggregate Team Data");

	teamDataSheet.clearContents();
	var columnNames = getColumnNames(rawScoreSheet);
	var rowData = getRowData(rawScoreSheet, columnNames.length);
	var teamData = compileTeamData(rowData);
	importTeamsIntoDatabase(teamData, teamDataSheet);

	controlPanel.getRange("B13").setValue(formatDate(new Date(Date.now())));
	log("updateMasterDatabase() success!");
}

function importTeamsIntoDatabase(teamData, teamDataSheet) {
	var teamAverages = compileTeamAverages(teamData);
	var teamComments = compileTeamComments(teamData);
	createColumnHeadings(teamDataSheet);
	var sortedTeamList = Object.keys(teamAverages).sort();
	var numRows = sortedTeamList.length;
	var numColumns = teamDataColumnHeadings.length;
	var values = new Array(numRows);
	sortedTeamList.forEach(function(team, index) {
		values[index] = [
			team,
			teamData[team].length,
			teamAverages[team].total,
			teamAverages[team].rules,
			teamAverages[team].fouls,
			teamAverages[team].fairMind,
			teamAverages[team].attitude,
			teamAverages[team].communication,
			JSON.stringify(teamComments[team].rules),
			JSON.stringify(teamComments[team].fouls),
			JSON.stringify(teamComments[team].fairMind),
			JSON.stringify(teamComments[team].attitude),
			JSON.stringify(teamComments[team].communication),
			JSON.stringify(teamComments[team].total),
		];
	});
	var range = teamDataSheet.getRange(2, 1, numRows, numColumns);
	range.setValues(values);
}

function createColumnHeadings(sheet) {
	var numColumns = teamDataColumnHeadings.length;
	sheet.getRange(`R1C1:R1C${numColumns}`).setValues([teamDataColumnHeadings]);
}

function compileTeamComments(teamData) {
	var teamComments = {};
	for (var team of Object.keys(teamData)) {
		teamComments[team] = {};
		for (commentCategory of scoreKeys) {
			teamComments[team][commentCategory] = [];
		}

		var comments = teamData[team].map(function(score) {
			return score.comments;
		});

		for (var row of comments) {
			for (var commentCategory of Object.keys(row)) {
				if (row[commentCategory] && row[commentCategory].trim() != "") {
					teamComments[team][commentCategory].push(row[commentCategory]);
				}
			}
		}
	}
	return teamComments;
}

function compileTeamAverages(teamData) {
	var averages = {};
	for (var team of Object.keys(teamData)) {
		averages[team] = {};
		var numScores = teamData[team].length;
		var scoresTotal = {};
		for (var key of scoreKeys) {
			scoresTotal[key] = 0;
		}
		for (var score of teamData[team]) {
			for (key of Object.keys(scoresTotal)) {
				scoresTotal[key] += score[key];
			}
		}

		for (var key of Object.keys(scoresTotal)) {
			scoresTotal[key] /= numScores;
		}
		averages[team] = scoresTotal;
	}

	return averages;
}

// return an object containing each team's scores received
function compileTeamData(rowData) {
	teamData = {};
	for (var row of rowData) {
		var scoredTeam = row[2];

		if (!teamData.hasOwnProperty(scoredTeam)) {
			teamData[scoredTeam] = [];
		}

		var score = {
			time: row[0],
			opponent: row[1], // this is the "your team" item in the form because this function returns scores in the perspective of the team being scored
			comments: {}
		};

		var total = 0;
		for (var i = 0; i < scoreKeys.length-1; i+=1) {
			score[scoreKeys[i]] = row[2*i+3];
			total += row[2*i+3];
			score.comments[scoreKeys[i]] = row[2*i+4];
		}
		score[scoreKeys[scoreKeys.length-1]] = total;
		score.comments[scoreKeys[scoreKeys.length-1]] = row[scoreKeys.length*2 + 1];

		teamData[scoredTeam].push(score);
	}
	return teamData;
}

function getRowData(sheet, numColumns) {
	var numRows = getFirstEmptyRow(sheet) - 2;
	return sheet.getRange(2, 1, numRows, numColumns).getValues();
}

function getColumnNames(sheet) {
	var row = sheet.getRange("1:1");
	var values = row.getValues()[0].filter(value => value);
	return values;
}

function getColumnData(sheet, numColumns) {
	var numRows = getFirstEmptyRow(sheet) - 2;
	var range = sheet.getRange(2, 1, numRows, numColumns);
	var values = range.getValues();
	var columns = new Array(numColumns);
	for (var i = 0; i < numColumns; i++) {
		columns[i] = new Array(numRows);
	}
	values.forEach(function(row, rowIndex) {
		row.forEach(function(cell, colIndex) {
			columns[colIndex][rowIndex] = cell;
		});
	});
	return columns;
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