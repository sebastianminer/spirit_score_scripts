const SCORE_KEYS = ['rules', 'fouls', 'fairMind', 'attitude', 'communication', 'total']
const TEAM_DATA_COLUMN_HEADINGS = [
	'Team',
	'Number of Scores',
	'Total',
	'Rules Knowledge and Use',
	'Fouls and Body Contact',
	'Fair Mindedness',
	'Positive Attitude and Self-Control',
	'Communication',
	'Comments (Rules Knowledge and Use)',
	'Comments (Fouls and Body Contact)',
	'Comments (Fair Mindedness)',
	'Comments (Positive Attitude and Self-Control)',
	'Comments (Communication)',
	'Additional Comments',
	'(Self) Number of Scores',
	'(Self) Total',
	'(Self) Rules Knowledge and Use',
	'(Self) Fouls and Body Contact',
	'(Self) Fair Mindedness',
	'(Self) Positive Attitude and Self-Control',
	'(Self) Communication',
	'(Self) Comments (Rules Knowledge and Use)',
	'(Self) Comments (Fouls and Body Contact)',
	'(Self) Comments (Fair Mindedness)',
	'(Self) Comments (Positive Attitude and Self-Control)',
	'(Self) Comments (Communication)',
	'(Self) Additional Comments'
]
const RAW_SCORE_COLUMN_HEADINGS = [
	'Timestamp',
	'Tournament', // note: not present in RAW_SCORE_COLUMN_HEADINGS in form_factory_control_panel.js
	'Your Team Name',
	'Opponent Team Name',
	'Date',
	'Round',
	'Rules Knowledge and Use',
	'Comments (Rules Knowledge and Use)',
	'(Self) Rules Knowledge and Use',
	'(Self) Comments (Rules Knowledge and Use)',
	'Fouls and Body Contact',
	'Comments (Fouls and Body Contact)',
	'(Self) Fouls and Body Contact',
	'(Self) Comments (Fouls and Body Contact)',
	'Fair Mindedness',
	'Comments (Fair Mindedness)',
	'(Self) Fair Mindedness',
	'(Self) Comments (Fair Mindedness)',
	'Attitude',
	'Comments (Attitude)',
	'(Self) Attitude',
	'(Self) Comments (Attitude)',
	'Communication',
	'Comments (Communication)',
	'(Self) Communication',
	'(Self) Comments (Communication)',
	'Additional Comments',
	'(Self) Additional Comments'
]

const RAW_SCORE_ENUM = RAW_SCORE_COLUMN_HEADINGS
	.map((heading, index) => ({ [heading]: index }))
	.reduce((previous, current) => ({ ...previous, ...current }), {})

// each category has a score and a comment for both the scoring team and the scored team. This is the number of columns created for each key.
const COLUMNS_PER_CATEGORY = 4

// number of hardcoded columns before scores (i.e. timestamp, tournament, team name, opponent name, date, round)
const NUM_INITIAL_COLUMNS = 6

let errno = 0

function pullScoresFromTournaments() {
	log('running pullScoresFromTournaments()')

	let lock = LockService.getScriptLock()
	try {
		lock.waitLock(10000)
	} catch (e) {
		log('It appears that this function is already being run by another user. Please wait for that operation to finish before calling this one again.')
		return
	}

	let controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel')
	let rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')

	let numTournamentCols = RAW_SCORE_COLUMN_HEADINGS.length
	let numMasterCols = numTournamentCols + 1
	let numMasterRowsForFormatClear = rawScoreSheet.getMaxRows() - 1
	rawScoreSheet.clearContents()
	rawScoreSheet.getRange(2, 1, numMasterRowsForFormatClear, numMasterCols).clearFormat()

	createColumnHeadings(rawScoreSheet, RAW_SCORE_COLUMN_HEADINGS)

	let thisFileId = SpreadsheetApp.getActiveSpreadsheet().getId()
	let thisFile = DriveApp.getFileById(thisFileId)
	let parentFolder = thisFile.getParents().next()
	let tournamentFolders = parentFolder.getFolders()

	while (tournamentFolders.hasNext()) {
		let tournamentFolder = tournamentFolders.next()
		let folderName = tournamentFolder.getName()

		log(`importing scores from folder ${folderName}`)
		try {
			let files = tournamentFolder.getFiles()
			let spreadsheetSeen = false
			while (files.hasNext()) {
				let file = files.next()
				if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
					continue
				} else if (spreadsheetSeen) {
					log(`unknown file seen in folder ${folderName}! please ensure that each folder contains only a form and a spreadsheet.`)
				}
				spreadsheetSeen = true
				let tournamentScoresSheet = SpreadsheetApp.openById(file.getId()).getSheetByName('Raw Scores')
				let numSourceRows = getFirstEmptyRow(tournamentScoresSheet) - 2
				let range = tournamentScoresSheet.getRange(2, 1, numSourceRows, numTournamentCols)
				let startRow = getFirstEmptyRow(rawScoreSheet)

				let values = range.getValues().filter(row => row[0])

				// insert tournament name at index 1 in each row of data.
				// not ideal to insert a value into a bunch of arrays, but it beats adding the tournament name to the form
				for (let i = 0; i < values.length; i++) {
					values[i].splice(1, 0, folderName)
				}

				let numTargetRows = values.length
				rawScoreSheet.getRange(startRow, 1, numTargetRows, numMasterCols).setValues(values)
			}
			log(`import succeeded for folder ${folderName}`)
		} catch (e) {
			log(`import failed for folder ${folderName}. ${e.name}: ${e.message}`)
			errno |= 1
		}
	}

	controlPanel.getRange('A13').setValue('Scores last pulled from tournament folders:')
	controlPanel.getRange('B13').setValue(formatDate(new Date(Date.now())))

	if (errno) {
		log('pullScoresFromTournaments() completed with one or more errors.')
	} else {
		log('pullScoresFromTournaments() success!')
	}
}

function aggregateScores() {
	log('running aggregateScores()')
	let controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel')
	let rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')
	let teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aggregate Team Data')

	teamDataSheet.clearContents()
	teamDataSheet.getRange('2:2').clearFormat() // clear green formatting on winner row, if the sheet had been sorted previously
	let columnNames = getColumnNames(rawScoreSheet)
	let rowData = getRowData(rawScoreSheet, columnNames.length)
	let teamData = compileTeamData(rowData)
	importTeamsIntoDatabase(teamData, teamDataSheet)

	controlPanel.getRange('A14').setValue('Scores last aggregated:')
	controlPanel.getRange('B14').setValue(formatDate(new Date(Date.now())))
	log('aggregateScores() success!')
}

function sortScores() {
	log('running sortScores()')
	let teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aggregate Team Data')
	let numColumns = TEAM_DATA_COLUMN_HEADINGS.length
	let numRows = getFirstEmptyRow(teamDataSheet) - 2
	let range = teamDataSheet.getRange(2, 1, numRows, numColumns)
	range.sort({
		column: 3,
		ascending: false
	})
	let winnerRange = teamDataSheet.getRange(2, 1, 1, numColumns)
	winnerRange.setBackground('#B7E1CD')
	log('sortScores() success!')
}

function colorFormattingButtonClick() {
	log('running colorFormattingButtonClick()')
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')
	if (sheet) {
		addConditionalFormatting(sheet)
		addDuplicateFormatting(sheet)
	}
	log('colorFormattingButtonClick() success!')
}

function addDuplicateFormatting(sheet) {
	let numRows = getFirstEmptyRow(sheet) - 2
	let numCols = RAW_SCORE_COLUMN_HEADINGS.length
	let range = sheet.getRange(2, 1, numRows, numCols)
	let rows = range.getValues()
	let possibleDuplicates = {}
	rows.forEach((row, index) => {
		let team = row[RAW_SCORE_ENUM['Your Team Name']]
		let opponent = row[RAW_SCORE_ENUM['Opponent Team Name']]
		let date = row[RAW_SCORE_ENUM['Date']]
		let tupleStr = team + opponent + date
		if (possibleDuplicates[tupleStr]) {
			possibleDuplicates[tupleStr].push(index)
		} else {
			possibleDuplicates[tupleStr] = [index]
		}
	})
	possibleDuplicates = Object.entries(possibleDuplicates)
		.filter(([key, value]) => key && value.length > 1)
		.reduce((cumulativeObj, [key, value]) => ({ ...cumulativeObj, [key]: value }), {})

	let teamColNum = RAW_SCORE_ENUM['Your Team Name'] + 1
	let dateColNum = RAW_SCORE_ENUM['Date'] + 1
	numCols = dateColNum - teamColNum + 1
	Object.keys(possibleDuplicates).forEach(key => {
		possibleDuplicates[key].forEach(rowIndex => {
			let rowNum = rowIndex + 2
			let range = sheet.getRange(rowNum, teamColNum, 1, numCols)
			range.clearFormat()
			range.setBackground('#A8DFFF')
		})
	})
}

function addConditionalFormatting(sheet) {
	let range = sheet.getRange('A2:AA1000')
	range.clearFormat()
	let zeroRule = SpreadsheetApp.newConditionalFormatRule()
		.setRanges([range])
		.whenNumberEqualTo(0)
		.setBackground('#F4C7C3')
		.build()
	let fourRule = SpreadsheetApp.newConditionalFormatRule()
		.setRanges([range])
		.whenNumberEqualTo(4)
		.setBackground('#84D6AF')
		.build()
	let sixRule = SpreadsheetApp.newConditionalFormatRule()
		.setRanges([range])
		.whenFormulaSatisfied('=AND(SUM($G2,$I2,$K2,$M2,$O2) <= 6, $A2 <> "")')
		.setBackground('#FCE8B2')
		.build()
	let fourteenRule = SpreadsheetApp.newConditionalFormatRule()
		.setRanges([range])
		.whenFormulaSatisfied('=AND(SUM($G2,$I2,$K2,$M2,$O2) >= 14, $A2 <> "")')
		.setBackground('#B7E1CD')
		.build()
	sheet.setConditionalFormatRules([zeroRule, fourRule, sixRule, fourteenRule])
}

function importTeamsIntoDatabase(teamData, teamDataSheet) {
	let teamAverages = compileTeamAverages(teamData)
	let teamComments = compileTeamComments(teamData)
	createColumnHeadings(teamDataSheet, TEAM_DATA_COLUMN_HEADINGS)
	let sortedTeamList = Object.keys(teamAverages).sort()
	let numRows = sortedTeamList.length
	let numColumns = TEAM_DATA_COLUMN_HEADINGS.length
	let scores = new Array(numRows)
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
		]
	})
	let range = teamDataSheet.getRange(2, 1, numRows, numColumns)
	range.setValues(scores)
}

function createColumnHeadings(sheet, columnHeadings) {
	let numColumns = columnHeadings.length
	sheet.getRange(`R1C1:R1C${numColumns}`).setValues([columnHeadings])
}

function compileTeamComments(teamData) {
	let teamComments = {}
	for (let team of Object.keys(teamData)) {
		teamComments[team] = {
			comments: {},
			self_comments: {}
		}
		for (commentCategory of SCORE_KEYS) {
			teamComments[team].comments[commentCategory] = []
			teamComments[team].self_comments[commentCategory] = []
		}

		let comments = teamData[team].scores.map(function(score) {
			return score.comments
		})
		let self_comments = teamData[team].self_scores.map(function(self_score) {
			return self_score.comments
		})

		for (let row of comments) {
			for (let commentCategory of Object.keys(row)) {
				if (row[commentCategory] && row[commentCategory].trim() != '') {
					teamComments[team].comments[commentCategory].push(row[commentCategory])
				}
			}
		}

		for (let row of self_comments) {
			for (let commentCategory of Object.keys(row)) {
				if (row[commentCategory] && row[commentCategory].trim() != '') {
					teamComments[team].self_comments[commentCategory].push(row[commentCategory])
				}
			}
		}
	}
	return teamComments
}

function compileTeamAverages(teamData) {
	let averages = {}
	for (let team of Object.keys(teamData)) {
		averages[team] = {
			scores: {},
			self_scores: {}
		}

		for (let direction of ['scores', 'self_scores']) {
			let numScores = teamData[team][direction].length
			let scoresTotal = {}
			for (let key of SCORE_KEYS) {
				scoresTotal[key] = 0
			}
			for (let score of teamData[team][direction]) {
				for (key of Object.keys(scoresTotal)) {
					scoresTotal[key] += score[key]
				}
			}

			for (let key of Object.keys(scoresTotal)) {
				if (!numScores) {
					if (key === 'total') {
						scoresTotal[key] = 0
					} else {
						scoresTotal[key] = '-'
					}
				} else {
					scoresTotal[key] /= numScores
				}
			}
			averages[team][direction] = scoresTotal
		}
	}

	return averages
}

// return an object containing each team's scores received
function compileTeamData(rowData) {
	teamData = {}
	for (let row of rowData) {
		let scoringTeam = row[RAW_SCORE_ENUM['Your Team Name']]
		let scoredTeam = row[RAW_SCORE_ENUM['Opponent Team Name']]
		let time = row[RAW_SCORE_ENUM['Timestamp']]
		let date = row[RAW_SCORE_ENUM['Date']]
		let round = row[RAW_SCORE_ENUM['Round']]

		if (!teamData.hasOwnProperty(scoringTeam)) {
			teamData[scoringTeam] = {
				scores: [],
				self_scores: []
			}
		}
		if (!teamData.hasOwnProperty(scoredTeam)) {
			teamData[scoredTeam] = {
				scores: [],
				self_scores: []
			}
		}

		let score = {
			time,
			opponent: scoringTeam, // this is the 'your team' item in the form because this function returns scores in the perspective of the team being scored
			date,
			round,
			comments: {}
		}

		let self_score = {
			time,
			opponent: scoredTeam, // contrast this with scoringTeam in score object
			date,
			round,
			comments: {}
		}

		let numNonTotalKeys = SCORE_KEYS.length - 1
		let total = 0
		let selfTotal = 0

		for (let i = 0; i < numNonTotalKeys; i++) {
			score[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + NUM_INITIAL_COLUMNS]
			self_score[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + NUM_INITIAL_COLUMNS + 2]
			total += score[SCORE_KEYS[i]]
			selfTotal += self_score[SCORE_KEYS[i]]
			score.comments[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + NUM_INITIAL_COLUMNS + 1]
			self_score.comments[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + NUM_INITIAL_COLUMNS + 3]
		}
		score[SCORE_KEYS[SCORE_KEYS.length-1]] = total
		self_score[SCORE_KEYS[SCORE_KEYS.length-1]] = selfTotal
		score.comments[SCORE_KEYS[SCORE_KEYS.length-1]] = row[(SCORE_KEYS.length-1)*COLUMNS_PER_CATEGORY + NUM_INITIAL_COLUMNS]; // additional comments
		self_score.comments[SCORE_KEYS[SCORE_KEYS.length-1]] = row[(SCORE_KEYS.length-1)*COLUMNS_PER_CATEGORY + NUM_INITIAL_COLUMNS + 1]; // self additional comments

		teamData[scoredTeam].scores.push(score)
		teamData[scoringTeam].self_scores.push(self_score)
	}
	return teamData
}

function getRowData(sheet, numColumns) {
	let numRows = getFirstEmptyRow(sheet) - 2
	return sheet.getRange(2, 1, numRows, numColumns).getValues().filter(row => row[0])
}

function getColumnNames(sheet) {
	let row = sheet.getRange('1:1')
	let values = row.getValues()[0].filter(value => value)
	return values
}

function getColumnData(sheet, numColumns) {
	let numRows = getFirstEmptyRow(sheet) - 2
	let range = sheet.getRange(2, 1, numRows, numColumns)
	let values = range.getValues()
	let columns = new Array(numColumns)
	for (let i = 0; i < numColumns; i++) {
		columns[i] = new Array(numRows)
	}
	values.forEach(function(row, rowIndex) {
		row.forEach(function(cell, colIndex) {
			columns[colIndex][rowIndex] = cell
		})
	})
	return columns
}

function getFirstEmptyRow(sheet) {
	return sheet.getLastRow() + 1
}

function formatDate(date) {
	let month = String(date.getMonth() + 1).padStart(2, '0')
	let dateInMonth = String(date.getDate()).padStart(2, '0')
	let year = String((date.getYear()) + 1900).padStart(2, '0')
	let hours = String(date.getHours()).padStart(2, '0')
	let minutes = String(date.getMinutes()).padStart(2, '0')
	let seconds = String(date.getSeconds()).padStart(2, '0')
	return `${month}/${dateInMonth}/${year} ${hours}:${minutes}:${seconds}`
}

function log(obj, omitDate) {
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log')
	let range = sheet.getRange('A1')
	let cellContents = range.getValue()
	let now = new Date(Date.now())
	let timeStamp = `[${formatDate(now)}]`
	cellContents = `${omitDate ? '' : timeStamp} ${String(obj)}\n${cellContents}`
	range.setValue(cellContents)
}