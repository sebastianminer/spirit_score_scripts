const scoreKeys = ["rules", "fouls", "fairMind", "attitude", "communication", "total"];
const teamDataColumnHeadings = [
	"Team",
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
	"Additional Comments",
	"(Self) Number of Scores",
	"(Self) Total",
	"(Self) Rules Knowledge and Use",
	"(Self) Fouls and Body Contact",
	"(Self) Fair Mindedness",
	"(Self) Positive Attitude and Self-Control",
	"(Self) Communication",
	"(Self) Comments (Rules Knowledge and Use)",
	"(Self) Comments (Fouls and Body Contact)",
	"(Self) Comments (Fair Mindedness)",
	"(Self) Comments (Positive Attitude and Self-Control)",
	"(Self) Comments (Communication)",
	"(Self) Additional Comments"
];
const rawScoreColumnHeadings = [
	"Timestamp",
	"Your Team Name",
	"Opponent Team Name",
	"Day",
	"Round",
	"Rules Knowledge and Use",
	"Comments (Rules Knowledge and Use)",
	"Fouls and Body Contact",
	"Comments (Fouls and Body Contact)",
	"Fair Mindedness",
	"Comments (Fair Mindedness)",
	"Positive Attitude and Self-Control",
	"Comments (Positive Attitude and Self-Control",
	"Communication",
	"Comments (Communication)",
	"Additional Comments",
	"(Self) Rules Knowledge and Use",
	"(Self) Comments (Rules Knowledge and Use)",
	"(Self) Fouls and Body Contact",
	"(Self) Comments (Fouls and Body Contact)",
	"(Self) Fair Mindedness",
	"(Self) Comments (Fair Mindedness)",
	"(Self) Positive Attitude and Self-Control",
	"(Self) Comments (Positive Attitude and Self-Control",
	"(Self) Communication",
	"(Self) Comments (Communication)",
	"(Self) Additional Comments",
];
var errno = 0;

function pullScoresFromBank() {
	log("running pullScoresFromBank()");

	var lock = LockService.getScriptLock();
	try {
		lock.waitLock(10000);
	} catch (e) {
		log("It appears that this function is already being run by another user. Please wait for that operation to finish before calling this one again.");
		return;
	}

	var controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control Panel");
	var rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Scores");

	rawScoreSheet.clearContents();
	createColumnHeadings(rawScoreSheet, rawScoreColumnHeadings);
	var numCols = getFirstEmptyColumn(rawScoreSheet);

	var bankSheetUrl = controlPanel.getRange("B1").getValue();
	var bankSheet = SpreadsheetApp.openByUrl(bankSheetUrl).getSheetByName("Tournament Control Panel IDs");
	var controlPanelIds = getControlPanelIds(bankSheet);
	for (id of controlPanelIds) {
		log(`importing scores from sheet ${id}`);
		try {
			var tournamentScoresSheet = SpreadsheetApp.openById(id).getSheetByName("Raw Scores");
			var numRows = getFirstEmptyRow(tournamentScoresSheet) - 1;
			var range = tournamentScoresSheet.getRange(2, 1, numRows, numCols);
			var startRow = getFirstEmptyRow(rawScoreSheet);
			rawScoreSheet.getRange(startRow, 1, numRows, numCols).setValues(range.getValues());
			log(`import succeeded for sheet ${id}`);
		} catch (e) {
			log(`import failed for sheet ${id}. ${e.name}: ${e.message}`);
			errno |= 1;
		}
	}

	controlPanel.getRange("A13").setValue("Scores last pulled from tournament bank:");
	controlPanel.getRange("B13").setValue(formatDate(new Date(Date.now())));

	if (errno) {
		log("pullScoresFromBank() completed with one or more errors.");
	} else {
		log("pullScoresFromBank() success!");
	}
}

function getControlPanelIds(bankSheet) {
	var rows = getFirstEmptyRow(bankSheet);
	return bankSheet.getRange("A:A").getValues().map(row => row[0]).filter(id => id.trim() != "");
}

function aggregateScores() {
	log("running updateMasterDatabase()");
	var controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control Panel");
	var rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Scores");
	var teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aggregate Team Data");

	teamDataSheet.clearContents();
	var columnNames = getColumnNames(rawScoreSheet);
	var rowData = getRowData(rawScoreSheet, columnNames.length);
	var teamData = compileTeamData(rowData);
	importTeamsIntoDatabase(teamData, teamDataSheet);

	controlPanel.getRange("A14").setValue("Scores last aggregated:");
	controlPanel.getRange("B14").setValue(formatDate(new Date(Date.now())));
	log("updateMasterDatabase() success!");
}

function importTeamsIntoDatabase(teamData, teamDataSheet) {
	var teamAverages = compileTeamAverages(teamData);
	var teamComments = compileTeamComments(teamData);
	createColumnHeadings(teamDataSheet, teamDataColumnHeadings);
	var sortedTeamList = Object.keys(teamAverages).sort();
	var numRows = sortedTeamList.length;
	var numColumns = teamDataColumnHeadings.length;
	var scores = new Array(numRows);
	sortedTeamList.forEach(function(team, index) {
		scores[index] = [
			team,
			teamData[team].scores.length,
			teamAverages[team].scores.total,
			teamAverages[team].scores.rules,
			teamAverages[team].scores.fouls,
			teamAverages[team].scores.fairMind,
			teamAverages[team].scores.attitude,
			teamAverages[team].scores.communication,
			JSON.stringify(teamComments[team].comments.rules),
			JSON.stringify(teamComments[team].comments.fouls),
			JSON.stringify(teamComments[team].comments.fairMind),
			JSON.stringify(teamComments[team].comments.attitude),
			JSON.stringify(teamComments[team].comments.communication),
			JSON.stringify(teamComments[team].comments.total),
			teamData[team].self_scores.length,
			teamAverages[team].self_scores.total,
			teamAverages[team].self_scores.rules,
			teamAverages[team].self_scores.fouls,
			teamAverages[team].self_scores.fairMind,
			teamAverages[team].self_scores.attitude,
			teamAverages[team].self_scores.communication,
			JSON.stringify(teamComments[team].self_comments.rules),
			JSON.stringify(teamComments[team].self_comments.fouls),
			JSON.stringify(teamComments[team].self_comments.fairMind),
			JSON.stringify(teamComments[team].self_comments.attitude),
			JSON.stringify(teamComments[team].self_comments.communication),
			JSON.stringify(teamComments[team].self_comments.total)
		];
	});
	var range = teamDataSheet.getRange(2, 1, numRows, numColumns);
	range.setValues(scores);
}

function createColumnHeadings(sheet, columnHeadings) {
	var numColumns = columnHeadings.length;
	sheet.getRange(`R1C1:R1C${numColumns}`).setValues([columnHeadings]);
}

function compileTeamComments(teamData) {
	var teamComments = {};
	for (var team of Object.keys(teamData)) {
		teamComments[team] = {
			comments: {},
			self_comments: {}
		};
		for (commentCategory of scoreKeys) {
			teamComments[team].comments[commentCategory] = [];
			teamComments[team].self_comments[commentCategory] = [];
		}

		var comments = teamData[team].scores.map(function(score) {
			return score.comments;
		});
		var self_comments = teamData[team].self_scores.map(function(self_score) {
			return self_score.comments;
		});

		for (var row of comments) {
			for (var commentCategory of Object.keys(row)) {
				if (row[commentCategory] && row[commentCategory].trim() != "") {
					teamComments[team].comments[commentCategory].push(row[commentCategory]);
				}
			}
		}

		for (var row of self_comments) {
			for (var commentCategory of Object.keys(row)) {
				if (row[commentCategory] && row[commentCategory].trim() != "") {
					teamComments[team].self_comments[commentCategory].push(row[commentCategory]);
				}
			}
		}
	}
	return teamComments;
}

function compileTeamAverages(teamData) {
	var averages = {};
	for (var team of Object.keys(teamData)) {
		averages[team] = {
			scores: {},
			self_scores: {}
		};

		for (var direction of ["scores", "self_scores"]) {
			var numScores = teamData[team][direction].length;
			var scoresTotal = {};
			for (var key of scoreKeys) {
				scoresTotal[key] = 0;
			}
			for (var score of teamData[team][direction]) {
				for (key of Object.keys(scoresTotal)) {
					scoresTotal[key] += score[key];
				}
			}

			for (var key of Object.keys(scoresTotal)) {
				if (!numScores) {
					scoresTotal[key] = "-";
				} else {
					scoresTotal[key] /= numScores;
				}
			}
			averages[team][direction] = scoresTotal;
		}
	}

	return averages;
}

// return an object containing each team's scores received
function compileTeamData(rowData) {
	teamData = {};
	for (var row of rowData) {
		var scoringTeam = row[1]
		var scoredTeam = row[2];

		if (!teamData.hasOwnProperty(scoringTeam)) {
			teamData[scoringTeam] = {
				scores: [],
				self_scores: []
			};
		}
		if (!teamData.hasOwnProperty(scoredTeam)) {
			teamData[scoredTeam] = {
				scores: [],
				self_scores: []
			};
		}

		var score = {
			time: row[0],
			opponent: scoringTeam, // this is the "your team" item in the form because this function returns scores in the perspective of the team being scored
			day: row[3],
			round: row[4],
			comments: {}
		};

		var self_score = {
			time: row[0],
			opponent: scoredTeam, // contrast this with scoringTeam in score object
			day: row[3],
			round: row[4],
			comments: {}
		}

		var numNonTotalKeys = scoreKeys.length - 1;
		var selfScoreOffset = 2*numNonTotalKeys + 1; // offset number of columns between a score and its corresponding self-score
		var total = 0;
		var selfTotal = 0;
		for (var i = 0; i < numNonTotalKeys; i+=1) {
			score[scoreKeys[i]] = row[2*i+5];

			// 2: each key has a score and a comment. This is the number of columns created for each key.
			// 5: number of hardcoded columns before scores (i.e. timestamp, team name, opponent name, day, round)
			self_score[scoreKeys[i]] = row[2*i + 5 + selfScoreOffset];
			total += row[2*i+5];
			selfTotal += row[2*i + 5 + selfScoreOffset];
			score.comments[scoreKeys[i]] = row[2*i+6];
			self_score.comments[scoreKeys[i]] = row[2*i + 6 + selfScoreOffset];
		}
		score[scoreKeys[scoreKeys.length-1]] = total;
		self_score[scoreKeys[scoreKeys.length-1]] = selfTotal;
		score.comments[scoreKeys[scoreKeys.length-1]] = row[(scoreKeys.length-1)*2 + 5]; // additional comments
		self_score.comments[scoreKeys[scoreKeys.length-1]] = row[(scoreKeys.length-1)*2 + 5 + selfScoreOffset]; // self additional comments

		teamData[scoredTeam].scores.push(score);
		teamData[scoringTeam].self_scores.push(self_score);
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

function getFirstEmptyColumn(sheet) {
	var row = sheet.getRange('1:1');
	var values = row.getValues();
	var ct = 0;
	while (values[0][ct] && values[0][ct] != "") {
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