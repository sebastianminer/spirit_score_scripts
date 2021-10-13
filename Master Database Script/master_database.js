const SCORE_KEYS = ['rules', 'fouls', 'fairMind', 'attitude', 'communication', 'total']
const TEAM_DATA_COLUMN_HEADINGS = [
	'Team',
	'Number of Scores Submitted',
	'Number of Scores Received',
	'Teams Who Need to Be Scored',
	'Teams from Whom a Score Is Needed',
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
	'Additional Comments'
]
const RAW_SCORE_COLUMN_HEADINGS = [
	'Timestamp',
	'Email',
	'Tournament', // note: not present in RAW_SCORE_COLUMN_HEADINGS in form_factory_control_panel.js or control_panel.js
	'Your Team Name',
	'Opponent Team Name',
	'Day',
	'Round',
	'Rules Knowledge and Use',
	'Comments (Rules Knowledge and Use)',
	'Fouls and Body Contact',
	'Comments (Fouls and Body Contact)',
	'Fair Mindedness',
	'Comments (Fair Mindedness)',
	'Attitude',
	'Comments (Attitude)',
	'Communication',
	'Comments (Communication)',
	'Additional Comments',
	'(Self) Rules Knowledge and Use',
	'(Self) Comments (Rules Knowledge and Use)',
	'(Self) Fouls and Body Contact',
	'(Self) Comments (Fouls and Body Contact)',
	'(Self) Fair Mindedness',
	'(Self) Comments (Fair Mindedness)',
	'(Self) Attitude',
	'(Self) Comments (Attitude)',
	'(Self) Communication',
	'(Self) Comments (Communication)',
	'(Self) Additional Comments'
]

const RAW_SCORE_ENUM = enumify(RAW_SCORE_COLUMN_HEADINGS)

// each category has a score and a comment. This is the number of columns created for each key for each team.
const COLUMNS_PER_CATEGORY = 2

// number of hardcoded columns before scores (i.e. timestamp, email, tournament, team name, opponent name, date, round)
const NUM_INITIAL_COLUMNS = RAW_SCORE_ENUM['Rules Knowledge and Use']

let errno = 0

function enumify(obj) {
	return obj.map((heading, index) => ({ [heading]: index }))
		.reduce((previous, current) => ({ ...previous, ...current }), {})
}

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
	let tournamentsFolder = parentFolder.getFoldersByName('Tournaments').next()
	let tournamentFolders = tournamentsFolder.getFolders()

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
					values[i].splice(RAW_SCORE_ENUM['Tournament'], 0, folderName)
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

function sortAggregateScoreSheet() {
	log('running sortAggregateScoreSheet()')
	let teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Aggregate Team Data')
	let numColumns = TEAM_DATA_COLUMN_HEADINGS.length
	let numRows = getFirstEmptyRow(teamDataSheet) - 2
	let range = teamDataSheet.getRange(2, 1, numRows, numColumns)
	range.sort({
		column: 6,
		ascending: false
	})
	let winnerRange = teamDataSheet.getRange(2, 1, 1, numColumns)
	winnerRange.setBackground('#B7E1CD')
	log('sortAggregateScoreSheet() success!')
}

function sortRawScoreSheet() {
	log('running sortRawScoreSheet()')
	let rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')
	let numColumns = RAW_SCORE_COLUMN_HEADINGS.length
	let numRows = getFirstEmptyRow(rawScoreSheet) - 2
	let range = rawScoreSheet.getRange(2, 1, numRows, numColumns)
	range.sort({
		column: 1,
	})
	addColorFormatting()
	log('sortRawScoreSheet() success!')
}

function colorFormattingButtonClick() {
	log('running colorFormattingButtonClick()')
	addColorFormatting()
	log('colorFormattingButtonClick() success!')
}

function addColorFormatting() {
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')
	if (sheet) {
		addConditionalFormatting(sheet)
		addDuplicateFormatting(sheet)
		addSelfScoreFormatting(sheet)
	}
}

function addSelfScoreFormatting(sheet) {
	let numRows = getFirstEmptyRow(sheet) - 2
	let numCols = RAW_SCORE_COLUMN_HEADINGS.length
	let range = sheet.getRange(2, 1, numRows, numCols)
	let rows = range.getValues()
	let duplicateRowIndices = []
	rows.forEach((row, index) => {
		let team = row[RAW_SCORE_ENUM['Your Team Name']]
		let opponent = row[RAW_SCORE_ENUM['Opponent Team Name']]
		if (team === opponent && team !== '') {
			duplicateRowIndices.push(index)
		}
	})

	let teamColNum = RAW_SCORE_ENUM['Your Team Name'] + 1
	let opponentColNum = RAW_SCORE_ENUM['Opponent Team Name'] + 1
	numCols = opponentColNum - teamColNum + 1
	duplicateRowIndices.forEach(rowIndex => {
		let rowNum = rowIndex + 2
		let range = sheet.getRange(rowNum, teamColNum, 1, numCols)
		range.clearFormat()
		range.setBackground('#B57924')
	})
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
		let date = row[RAW_SCORE_ENUM['Day']]
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
	let dateColNum = RAW_SCORE_ENUM['Day'] + 1
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
	let columnsToSum = ['Rules Knowledge and Use', 'Fouls and Body Contact', 'Fair Mindedness', 'Attitude', 'Communication']
		.map(key => columnToLetter(RAW_SCORE_ENUM[key] + 1))
	let sumArgumentsString = columnsToSum.map(letter => `$${letter}2`).join(',')
	let numColumns = RAW_SCORE_COLUMN_HEADINGS.length
	let range = sheet.getRange(`A2:${numColumns}1000`)
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
		.whenFormulaSatisfied(`=AND(SUM(${sumArgumentsString}) <= 6, $A2 <> "")`)
		.setBackground('#FCE8B2')
		.build()
	let fourteenRule = SpreadsheetApp.newConditionalFormatRule()
		.setRanges([range])
		.whenFormulaSatisfied(`=AND(SUM(${sumArgumentsString}) >= 14, $A2 <> "")`)
		.setBackground('#B7E1CD')
		.build()
	sheet.setConditionalFormatRules([zeroRule, fourRule, sixRule, fourteenRule])
}

function importTeamsIntoDatabase(teamData, teamDataSheet) {
	let teamAverages = compileTeamAverages(teamData)
	let teamComments = compileTeamComments(teamData)
	let missedTeams = compileMissedTeams(teamData)
	createColumnHeadings(teamDataSheet, TEAM_DATA_COLUMN_HEADINGS)
	let sortedTeamList = Object.keys(teamAverages).sort()
	let numRows = sortedTeamList.length
	let numColumns = TEAM_DATA_COLUMN_HEADINGS.length
	let scores = new Array(numRows)
	sortedTeamList.forEach(function(team, index) {
		scores[index] = [
			team,
			teamAverages[team].scoresSubmitted,
			teamAverages[team].scoresReceived,
			missedTeams[team].scoresNeededFor,
			missedTeams[team].scoresNeededFrom,
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
			JSON.stringify(teamComments[team].total)
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
		teamComments[team] = {}
		for (commentCategory of SCORE_KEYS) {
			teamComments[team][commentCategory] = []
		}

		let comments = teamData[team].map(function(score) {
			return score.comments
		})

		for (let row of comments) {
			for (let commentCategory of Object.keys(row)) {
				if (row[commentCategory] && row[commentCategory].trim() != '') {
					teamComments[team][commentCategory].push(row[commentCategory])
				}
			}
		}
	}
	return teamComments
}

function compileTeamAverages(teamData) {
	let averages = {}

	// initialize averages[team] for each team
	for (let team of Object.keys(teamData)) {
		averages[team] = {}
		averages[team].scoresReceived = teamData[team].length
		averages[team].scoresSubmitted = 0
		for (let key of SCORE_KEYS) {
			averages[team][key] = 0
		}
	}

	// set values for each averages[team]
	for (let team of Object.keys(teamData)) {
		let scoresTotal = averages[team]
		let numScores = scoresTotal.scoresReceived
		for (let score of teamData[team]) {
			for (key of SCORE_KEYS) {
				scoresTotal[key] += score[key]
			}
			let scoringTeam = score.opponent
			averages[scoringTeam].scoresSubmitted++
		}

		for (let key of SCORE_KEYS) {
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
		averages[team] = scoresTotal
	}

	return averages
}

function compileMissedTeams(teamData) {
	let opponentQuantities = Object.keys(teamData).reduce((cumulativeObj, teamName) => ({
			...cumulativeObj,
			[teamName]: {}
		}), {}
	)

	Object.keys(teamData).forEach(teamName => {
		let teamScores = teamData[teamName]
		teamScores.forEach(score => {
			if (!opponentQuantities[teamName][score.opponent]) {
				opponentQuantities[teamName][score.opponent] = {
					scoresFor: 0,
					scoresFrom: 0
				}
				opponentQuantities[score.opponent][teamName] = {
					scoresFor: 0,
					scoresFrom: 0
				}
			}
			opponentQuantities[score.opponent][teamName].scoresFor++
			opponentQuantities[teamName][score.opponent].scoresFrom++
		})
	})

	let missedTeams = {}
	Object.keys(teamData).forEach(teamName => {
		missedTeams[teamName] = {}

		Object.keys(opponentQuantities[teamName]).forEach(opponentName => {
			let quantity = opponentQuantities[teamName][opponentName]
			let scoresNeededFor = quantity.scoresFrom - quantity.scoresFor
			let scoresNeededFrom = -scoresNeededFor

			if (scoresNeededFor) {
				if (!missedTeams[teamName][opponentName]) {
					missedTeams[teamName][opponentName] = {
						scoresNeededFor: 0,
						scoresNeededFrom: 0
					}
				}
			}
			if (scoresNeededFor > 0) {
				missedTeams[teamName][opponentName].scoresNeededFor = scoresNeededFor
			} else if (scoresNeededFrom > 0) {
				missedTeams[teamName][opponentName].scoresNeededFrom = scoresNeededFrom
			}
		})
	})

	return getMissedTeamsAsString(missedTeams)
}

function getMissedTeamsAsString(missedTeamsObj) {
	let missedTeams = {}
	Object.keys(missedTeamsObj).forEach(teamName => {
		missedTeams[teamName] = {
			scoresNeededFor: '',
			scoresNeededFrom: ''
		}
		Object.keys(missedTeamsObj[teamName]).sort().forEach(opponentName => {
			let scoresNeededFor = missedTeamsObj[teamName][opponentName].scoresNeededFor
			let scoresNeededFrom = missedTeamsObj[teamName][opponentName].scoresNeededFrom
			if (scoresNeededFor) {
				missedTeams[teamName].scoresNeededFor += `${opponentName} (${scoresNeededFor})\n`
			} else if (scoresNeededFrom) {
				missedTeams[teamName].scoresNeededFrom += `${opponentName} (${scoresNeededFrom})\n`
			}
		})
		let scoresNeededFor = missedTeams[teamName].scoresNeededFor
		let scoresNeededFrom = missedTeams[teamName].scoresNeededFrom
		missedTeams[teamName].scoresNeededFor = scoresNeededFor.substring(0, scoresNeededFor.length - 1)
		missedTeams[teamName].scoresNeededFrom = scoresNeededFrom.substring(0, scoresNeededFrom.length - 1)
	})
	return missedTeams
}

// return an object containing each team's scores received
function compileTeamData(rowData) {
	teamData = {}
	for (let row of rowData) {
		let scoringTeam = row[RAW_SCORE_ENUM['Your Team Name']]
		let scoredTeam = row[RAW_SCORE_ENUM['Opponent Team Name']]
		let time = row[RAW_SCORE_ENUM['Timestamp']]
		let date = row[RAW_SCORE_ENUM['Day']]
		let round = row[RAW_SCORE_ENUM['Round']]

		if (!teamData.hasOwnProperty(scoringTeam)) {
			teamData[scoringTeam] = []
		}
		if (!teamData.hasOwnProperty(scoredTeam)) {
			teamData[scoredTeam] = []
		}

		let score = {
			time,
			opponent: scoringTeam, // this is the 'your team' item in the form because this function returns scores in the perspective of the team being scored
			date,
			round,
			comments: {}
		}

		let numNonTotalKeys = SCORE_KEYS.length - 1
		let total = 0

		for (let i = 0; i < numNonTotalKeys; i++) {
			score[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + NUM_INITIAL_COLUMNS]
			total += score[SCORE_KEYS[i]]
			score.comments[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + NUM_INITIAL_COLUMNS + 1].toString()
		}
		score[SCORE_KEYS[SCORE_KEYS.length-1]] = total
		score.comments[SCORE_KEYS[SCORE_KEYS.length-1]] = row[(SCORE_KEYS.length-1)*COLUMNS_PER_CATEGORY + NUM_INITIAL_COLUMNS].toString() // additional comments

		teamData[scoredTeam].push(score)
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

function columnToLetter(column) {
	var temp, letter = '';
	while (column > 0) {
		temp = (column - 1) % 26;
		letter = String.fromCharCode(temp + 65) + letter;
		column = (column - temp - 1) / 26;
	}
	return letter;
}

function letterToColumn(letter) {
	var column = 0, length = letter.length;
	for (var i = 0; i < length; i++) {
		column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
	}
	return column;
}