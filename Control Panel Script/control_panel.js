const SCORE_KEYS = ['rules', 'fouls', 'communication', 'total']
const SCORE_KEYS_TO_COLUMN_HEADING = {
	rules: 'Rules Knowledge and Use',
	fouls: 'Fouls and Body Contact',
	communication: 'Communication and Conduct',
	total: 'Total'
}
const TEAM_DATA_COLUMN_HEADINGS = [
	'Team',
	'# of Scores Submitted',
	'# of Scores Received',
	'Teams Scored',
	'Teams They Need to Score',
	'Teams from Whom a Score Is Needed',
	'Total',
	'Rules Knowledge and Use',
	'Fouls and Body Contact',
	'Communication and Conduct',
	'Comments (Rules Knowledge and Use)',
	'Comments (Fouls and Body Contact)',
	'Comments (Communication and Conduct)',
	'Additional Comments'
]
const RAW_SCORE_COLUMN_HEADINGS = [
	'Timestamp',
	'Email',
	'Your Team Name',
	'Opponent Team Name',
	'Day',
	'Round',
	'Rules Knowledge and Use',
	'Comments (Rules Knowledge and Use)',
	'Fouls and Body Contact',
	'Comments (Fouls and Body Contact)',
	'Communication and Conduct',
	'Comments (Communication and Conduct)',
	'Additional Comments',
  'Game Had Observers',
	'Observer Score',
	'Observer Comments',
	'(Self) Rules Knowledge and Use',
	'(Self) Comments (Rules Knowledge and Use)',
	'(Self) Fouls and Body Contact',
	'(Self) Comments (Fouls and Body Contact)',
	'(Self) Communication and Conduct',
	'(Self) Comments (Communication and Conduct)',
	'(Self) Additional Comments'
]
const MAIL_MERGE_COLUMN_HEADINGS = [
	'Email Addresses',
	'Subject',
	'Body'
]

const RAW_SCORE_ENUM = enumify(RAW_SCORE_COLUMN_HEADINGS)
const TEAM_DATA_ENUM = enumify(TEAM_DATA_COLUMN_HEADINGS)
const MAIL_MERGE_ENUM = enumify(MAIL_MERGE_COLUMN_HEADINGS)

const PROCESSED_SCORE_SHEET_NAME = 'Processed Scores'
const PROCESSED_SCORE_COLUMN_HEADINGS = [...RAW_SCORE_COLUMN_HEADINGS]
PROCESSED_SCORE_COLUMN_HEADINGS.splice(RAW_SCORE_ENUM['Rules Knowledge and Use'], 0, 'Total Score')
PROCESSED_SCORE_COLUMN_HEADINGS.push('Game Key', 'Game First Timestamp', 'Game Pair Order')

const PROCESSED_SCORE_ENUM = enumify(PROCESSED_SCORE_COLUMN_HEADINGS)

// each category has a score and a comment. This is the number of columns created for each key for each team.
const COLUMNS_PER_CATEGORY = 2

// number of hardcoded columns before scores (i.e. timestamp, email, team name, opponent name, date, round)
const NUM_INITIAL_COLUMNS = RAW_SCORE_ENUM['Rules Knowledge and Use']
		
// number of columns for observer scores
const NUM_OBSERVER_COLUMNS = 3

function enumify(obj) {
	return obj.map((heading, index) => ({ [heading]: index }))
		.reduce((previous, current) => ({ ...previous, ...current }), {})
}

function updateForm() {
	log('running updateForm()')
	let spr = SpreadsheetApp.getActiveSpreadsheet()
	let tournamentURL = spr.getRange('TournamentPageURL').getValue()
	let teamSet = getTeamsFromURL(tournamentURL)
	let teams = Array.from(teamSet).sort()
	let rawScoreSheet = spr.getSheetByName('Raw Scores')
	let form = FormApp.openByUrl(rawScoreSheet.getFormUrl())
	let listItems = form.getItems(FormApp.ItemType.LIST)
	let yourTeamItem = listItems[0].asListItem()
	let opponentTeamItem = listItems[1].asListItem()

	setListItemChoices(yourTeamItem, teams)
	setListItemChoices(opponentTeamItem, teams)

	spr.getRange('LiveFormURL').setValue(rawScoreSheet.getFormUrl())

	spr.getRange('TeamNamesLastUpdated').setValue(formatDate(new Date(Date.now())))
// Convert the array of teams to a 2D array for setValues
	let teams2D = teams.map(team => [team.trim()])
// Set the values starting from cell H6
	spr.getRange('TeamNames').offset(0, 0, teams.length, 1).setValues(teams2D)
	log('updateForm() success!')
}

function aggregateScoresAndGenerateMailMerge() {
	log('running aggregateScoresAndGenerateMailMerge()')
	let controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel')
	let rawScoreSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')
	let teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Team Data')

  teamDataSheet.clearContents()
	teamDataSheet.getRange('2:2').clearFormat() // clear green formatting on winner row, if the sheet had been sorted previously

	let columnNames = getColumnNames(rawScoreSheet)
	let rowData = getRowData(rawScoreSheet, columnNames.length)
	let teamData = compileTeamData(rowData)
	let selfScoreTeamData = compileTeamDataForSelfScores(rowData)
	importTeamsIntoDatabase(teamData, teamDataSheet)
	let mailMergeData = generateMailMerge(teamData, selfScoreTeamData, rowData)
	importMailMergeIntoSheet(mailMergeData)

	controlPanel.getRange('ScoresLastCalculated').setValue(formatDate(new Date(Date.now())))
	log('aggregateScoresAndGenerateMailMerge() success!')
}

function doItAll() {
	colorFormattingButtonClick()
  sortRawScoreSheet()
	addProcessedScoreSheet()
  aggregateScoresAndGenerateMailMerge()
}

function colorFormattingButtonClick() {
	log('running colorFormattingButtonClick()')
	addColorFormattingAndColumnHeadings()
	log('colorFormattingButtonClick() success!')
}

function addColorFormattingAndColumnHeadings() {
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Scores')
	if (sheet) {
		addConditionalFormatting(sheet)
		addDuplicateFormatting(sheet)
		addSelfScoreFormatting(sheet)
		createColumnHeadings(sheet, RAW_SCORE_COLUMN_HEADINGS)
	}
}

function sortAggregateScoreSheet() {
	log('running sortAggregateScoreSheet()')
	let teamDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Team Data')
	let numColumns = TEAM_DATA_COLUMN_HEADINGS.length
	let numRows = getFirstEmptyRow(teamDataSheet) - 2
	let range = teamDataSheet.getRange(2, 1, numRows, numColumns)
	range.sort({
		column: TEAM_DATA_ENUM.Total + 1,
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
		column: PROCESSED_SCORE_ENUM.Timestamp + 1,
	})
	addColorFormattingAndColumnHeadings()
	log('sortRawScoreSheet() success!')
}

function addProcessedScoreSheet() {
	log('running addProcessedScoreSheet()')

	const ss = SpreadsheetApp.getActiveSpreadsheet()
	const rawScoreSheet = ss.getSheetByName('Raw Scores')
	const numRawColumns = RAW_SCORE_COLUMN_HEADINGS.length
	const rawRows = getRowData(rawScoreSheet, numRawColumns)

	const processedRows = rawRows.map(row => {
		const total =
			Number(row[RAW_SCORE_ENUM['Rules Knowledge and Use']] || 0) +
			Number(row[RAW_SCORE_ENUM['Fouls and Body Contact']] || 0) +
			Number(row[RAW_SCORE_ENUM['Communication and Conduct']] || 0)

		const team = row[RAW_SCORE_ENUM['Your Team Name']]
		const opponent = row[RAW_SCORE_ENUM['Opponent Team Name']]
		const day = row[RAW_SCORE_ENUM['Day']]
		const round = row[RAW_SCORE_ENUM['Round']]

		const teamsSorted = [team, opponent].sort()
		const gameKey = `${day} | ${round} | ${teamsSorted[0]} | ${teamsSorted[1]}`

		const totalColumnIndex = RAW_SCORE_ENUM['Rules Knowledge and Use']

		return [
			...row.slice(0, totalColumnIndex),
			total,
			...row.slice(totalColumnIndex),
			gameKey,
			'', // Game First Timestamp, filled below
			''  // Game Pair Order, filled below
		]
	})

	const gameFirstTimestamp = {}
	processedRows.forEach(row => {
		const gameKey = row[PROCESSED_SCORE_ENUM['Game Key']]
		const timestamp = row[PROCESSED_SCORE_ENUM['Timestamp']]
		if (!gameFirstTimestamp[gameKey] || timestamp < gameFirstTimestamp[gameKey]) {
			gameFirstTimestamp[gameKey] = timestamp
		}
	})

	const gameCounts = {}
	processedRows.forEach(row => {
		const gameKey = row[PROCESSED_SCORE_ENUM['Game Key']]
		gameCounts[gameKey] = (gameCounts[gameKey] || 0) + 1

		row[PROCESSED_SCORE_ENUM['Game First Timestamp']] = gameFirstTimestamp[gameKey]
		row[PROCESSED_SCORE_ENUM['Game Pair Order']] = gameCounts[gameKey]
	})

	processedRows.sort((a, b) => {
		const firstTimeA = a[PROCESSED_SCORE_ENUM['Game First Timestamp']]
		const firstTimeB = b[PROCESSED_SCORE_ENUM['Game First Timestamp']]
		const pairOrderA = a[PROCESSED_SCORE_ENUM['Game Pair Order']]
		const pairOrderB = b[PROCESSED_SCORE_ENUM['Game Pair Order']]
		const timestampA = a[PROCESSED_SCORE_ENUM['Timestamp']]
		const timestampB = b[PROCESSED_SCORE_ENUM['Timestamp']]

		return firstTimeA - firstTimeB ||
			pairOrderA - pairOrderB ||
			timestampA - timestampB
	})

	let processedSheet = ss.getSheetByName(PROCESSED_SCORE_SHEET_NAME)
	if (processedSheet) {
		processedSheet.clearContents()
    processedSheet
    	.getRange(2, 1, processedSheet.getMaxRows() - 1, processedSheet.getMaxColumns())
	    .setBackground(null).setWrap(false)
      .setBorder(false, false, false, false, false, false)
    // resetRowsToModalHeight(processedSheet)
	} else {
		processedSheet = ss.insertSheet(PROCESSED_SCORE_SHEET_NAME)
	}

	createColumnHeadings(processedSheet, PROCESSED_SCORE_COLUMN_HEADINGS)

	if (processedRows.length) {
		processedSheet
			.getRange(2, 1, processedRows.length, PROCESSED_SCORE_COLUMN_HEADINGS.length)
			.setValues(processedRows)
	}

  formatProcessedScoreSheet(processedSheet)
  addConditionalFormatting(processedSheet)
  addDuplicateFormatting(processedSheet)
  addSelfScoreFormatting(processedSheet)

	log('addProcessedScoreSheet() success!')
}

function formatProcessedScoreSheet(processedSheet) {
	const totalColumnLetter = columnToLetter(PROCESSED_SCORE_ENUM['Total Score'] + 1)
	const numRows = Math.max(processedSheet.getLastRow() - 1, 1)

	const totalsColumnRange = processedSheet.getRange(
		2,
		PROCESSED_SCORE_ENUM['Total Score'] + 1,
		numRows,
		1
	)

	processedSheet.setFrozenRows(1)

  const gamePairOrderCol = PROCESSED_SCORE_ENUM['Game Pair Order'] + 1
  const lastRow = processedSheet.getLastRow()
  const lastCol = processedSheet.getLastColumn()

  if (lastRow > 1) {
    const pairOrderValues = processedSheet
      .getRange(2, gamePairOrderCol, lastRow - 1, 1)
      .getValues()

    pairOrderValues.forEach((row, index) => {
      const pairOrder = row[0]
      if (pairOrder === 1) {
        const sheetRow = index + 2
        processedSheet
          .getRange(sheetRow, 1, 1, lastCol)
          .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      }
    })
  }
}

function addSelfScoreFormatting(sheet) {
  const { enum: scoreEnum, headings } = getScoreSheetConfig(sheet);

  const numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) return;

  const numCols = headings.length;
  const range = sheet.getRange(2, 1, numRows, numCols);
  const rows = range.getValues();

  const selfScoreRowIndices = [];

  rows.forEach((row, index) => {
    const team = row[scoreEnum['Your Team Name']];
    const opponent = row[scoreEnum['Opponent Team Name']];

    if (team === opponent && team !== '') {
      selfScoreRowIndices.push(index);
    }
  });

  const teamColNum = scoreEnum['Your Team Name'] + 1;
  const opponentColNum = scoreEnum['Opponent Team Name'] + 1;
  const colsToColor = opponentColNum - teamColNum + 1;

  selfScoreRowIndices.forEach(rowIndex => {
    const rowNum = rowIndex + 2;
    sheet
      .getRange(rowNum, teamColNum, 1, colsToColor)
      .setBackground('#B57924');
  });
}

function addDuplicateFormatting(sheet) {
  const { enum: scoreEnum, headings } = getScoreSheetConfig(sheet);

  const numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) return;

  const numCols = headings.length;
  const range = sheet.getRange(2, 1, numRows, numCols);
  const rows = range.getValues();

  let possibleDuplicates = {};

  rows.forEach((row, index) => {
    const team = row[scoreEnum['Your Team Name']];
    const opponent = row[scoreEnum['Opponent Team Name']];
    const date = row[scoreEnum['Day']];
    const tupleStr = `${team}|${opponent}|${date}`;

    if (!team || !opponent || !date) return;

    if (possibleDuplicates[tupleStr]) {
      possibleDuplicates[tupleStr].push(index);
    } else {
      possibleDuplicates[tupleStr] = [index];
    }
  });

  possibleDuplicates = Object.entries(possibleDuplicates)
    .filter(([key, value]) => key && value.length > 1)
    .reduce((obj, [key, value]) => ({ ...obj, [key]: value }), {});

  const teamColNum = scoreEnum['Your Team Name'] + 1;
  const dateColNum = scoreEnum['Day'] + 1;
  const colsToColor = dateColNum - teamColNum + 1;

  Object.keys(possibleDuplicates).forEach(key => {
    possibleDuplicates[key].forEach(rowIndex => {
      const rowNum = rowIndex + 2;
      sheet
        .getRange(rowNum, teamColNum, 1, colsToColor)
        .setBackground('#A8DFFF');
    });
  });
}

function addConditionalFormatting(sheet) {
  const { enum: scoreEnum, headings } = getScoreSheetConfig(sheet)

  const scoreColumnNames = [
    'Rules Knowledge and Use',
    'Fouls and Body Contact',
    'Communication and Conduct',
    'Observer Score',
    '(Self) Rules Knowledge and Use',
    '(Self) Fouls and Body Contact',
    '(Self) Communication and Conduct'
  ]

  const numRows = Math.max(sheet.getLastRow() - 1, 1)
  const numColumns = headings.length

  const dataRange = sheet.getRange(2, 1, numRows, numColumns)

  const scoreRanges = scoreColumnNames.map(columnName =>
    sheet.getRange(2, scoreEnum[columnName] + 1, numRows, 1)
  )

  const scoreColumnLetters = scoreColumnNames.slice(0, 3).map(columnName =>
    columnToLetter(scoreEnum[columnName] + 1)
  )

  const sumFormula = scoreColumnLetters.map(letter => `$${letter}2`).join(',')

  const zeroRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges(scoreRanges)
    .whenNumberEqualTo(0)
    .setBackground('#F4C7C3')
    .build()

  const fourRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges(scoreRanges)
    .whenNumberEqualTo(4)
    .setBackground('#84D6AF')
    .build()

  const lowTotalRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([dataRange])
    .whenFormulaSatisfied(`=AND($A2 <> "", SUM(${sumFormula}) <= 4)`)
    .setBackground('#FCE8B2')
    .build()

  const highTotalRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([dataRange])
    .whenFormulaSatisfied(`=AND($A2 <> "", SUM(${sumFormula}) >= 8)`)
    .setBackground('#B7E1CD')
    .build()

  sheet.setConditionalFormatRules([
    zeroRule,
    fourRule,
    lowTotalRule,
    highTotalRule
  ])
}

function getColumnNames(sheet) {
	let row = sheet.getRange('1:1')
	let values = row.getValues()[0].filter(value => value)
	return values
}

function getRowDataIncludingEmpty(sheet, numColumns) {
	let numRows = getFirstEmptyRow(sheet) - 2
	return sheet.getRange(2, 1, numRows, numColumns).getValues()
}

function getRowData(sheet, numColumns) {
	return getRowDataIncludingEmpty(sheet, numColumns).filter(row => row[0])
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

function compileTeamDataForSelfScores(rowData) {
	teamData = {}
	for (let row of rowData) {
		let scoringTeam = row[RAW_SCORE_ENUM['Your Team Name']]
		let opponent = row[RAW_SCORE_ENUM['Opponent Team Name']]
		let time = row[RAW_SCORE_ENUM['Timestamp']]
		let date = row[RAW_SCORE_ENUM['Day']]
		let round = row[RAW_SCORE_ENUM['Round']]

		if (!teamData.hasOwnProperty(scoringTeam)) {
			teamData[scoringTeam] = []
		}

		let score = {
			time,
			opponent,
			date,
			round,
			comments: {}
		}

		let numNonTotalKeys = SCORE_KEYS.length - 1
		let total = 0
		let columnOffset = NUM_INITIAL_COLUMNS
				+ COLUMNS_PER_CATEGORY * numNonTotalKeys // all scored categories
				+ 1 // additional comments
				+ NUM_OBSERVER_COLUMNS;

		for (let i = 0; i < numNonTotalKeys; i++) {
			score[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + columnOffset]
			total += score[SCORE_KEYS[i]] || 0
			score.comments[SCORE_KEYS[i]] = row[COLUMNS_PER_CATEGORY*i + columnOffset + 1].toString()
		}
		score[SCORE_KEYS[SCORE_KEYS.length-1]] = total
		score.comments[SCORE_KEYS[SCORE_KEYS.length-1]] = row[(SCORE_KEYS.length-1)*COLUMNS_PER_CATEGORY + columnOffset].toString() // additional comments


		// Only add a score if it's been filled out at all
		if (SCORE_KEYS.some(key => key !== 'total' && score[key] !== '')) {
			teamData[scoringTeam].push(score)
		}
	}
	return teamData
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
			missedTeams[team].scoresFor,
			missedTeams[team].scoresNeededFor,
			missedTeams[team].scoresNeededFrom,
			teamAverages[team].total,
			teamAverages[team].rules,
			teamAverages[team].fouls,
			teamAverages[team].communication,
			JSON.stringify(teamComments[team].rules),
			JSON.stringify(teamComments[team].fouls),
			JSON.stringify(teamComments[team].communication),
			JSON.stringify(teamComments[team].total)
		]
	})
	let range = teamDataSheet.getRange(2, 1, numRows, numColumns)
	range.setValues(scores)
}

function getTeamToEmailAddressesDictionary(rowData) {
	const teamToEmailAddressesDictionary = rowData.reduce((dict, row) => {
		const teamName = row[RAW_SCORE_ENUM['Your Team Name']]
		const email = row[RAW_SCORE_ENUM['Email']].toLowerCase()
		if (!dict.hasOwnProperty(teamName)) {
			dict[teamName] = new Set()
		}
		dict[teamName].add(email)
		return dict
	}, {})
	Object.keys(teamToEmailAddressesDictionary).forEach(teamName => {
		teamToEmailAddressesDictionary[teamName] = Array.from(teamToEmailAddressesDictionary[teamName])
	})
	return teamToEmailAddressesDictionary
}

function generateMailMerge(teamData, selfScoreTeamData, rowData) {
	const controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel') 
	const header = controlPanel.getRange('MailMergeHeaderText').getValue()
	const footer = controlPanel.getRange('MailMergeFooterText').getValue()
			
	// The same as teamData, but the scoring team is now the key, and the
	// receiving team is the opponent.
	const inverseTeamData = Object.entries(teamData).reduce((dict, [receivingTeam, scores]) => {
		scores.forEach(score => {
			const scoringTeam = score.opponent
			if (!dict.hasOwnProperty(scoringTeam)) {
				dict[scoringTeam] = []
			}
			let inverseScore = {...score}
			inverseScore.opponent = receivingTeam
			dict[scoringTeam].push(inverseScore)
		})
		return dict
	}, {})

	const teamToEmailAddressesDictionary = getTeamToEmailAddressesDictionary(rowData)
	const mailMergeByTeam = Object.keys(teamData).reduce((dict, teamName) => ({
		...dict,
		[teamName]: {
			emailAddressesString: teamToEmailAddressesDictionary[teamName]?.join(', ') || '',
			emailSubject: `Spirit Scores for ${teamName}`,
			emailBody: `${header}\n`
	 }}), {})
	Object.entries(teamData).forEach(([teamName, scores]) => {
		mailMergeByTeam[teamName].emailBody += '\n'
		mailMergeByTeam[teamName].emailBody += 'Scores given to you by other teams:'
		
		scores.forEach(score => {
			const scoringTeam = score.opponent
			const round = score.round || 'N/A'
			const day = score.date || 'N/A'
			
			mailMergeByTeam[teamName].emailBody += "\n\n"
			mailMergeByTeam[teamName].emailBody += `Team: ${scoringTeam}\n`
			mailMergeByTeam[teamName].emailBody += `Round: ${round}\n`
			mailMergeByTeam[teamName].emailBody += `Day: ${day}`
			mailMergeByTeam[teamName].emailBody += SCORE_KEYS.reduce((str, scoreKey) =>
				str + `\n${SCORE_KEYS_TO_COLUMN_HEADING[scoreKey]}: ${score[scoreKey]}`, ''
			)
      mailMergeByTeam[teamName].emailBody += SCORE_KEYS.reduce((str, scoreKey) => {
        if (score.comments && score.comments[scoreKey] !== '') {
          return str + `\n${SCORE_KEYS_TO_COLUMN_HEADING[scoreKey]} Comments: ${score.comments[scoreKey]}`;
        }
        return str;
      }, '');
		})
				
		const inverseScores = inverseTeamData[teamName]
		if (inverseScores) {
			mailMergeByTeam[teamName].emailBody += '\n\n\n\n'
			mailMergeByTeam[teamName].emailBody += 'Scores that you gave other teams:'
					
			inverseScores.forEach(score => {
				const receivingTeam = score.opponent
				const round = score.round || 'N/A'
				const day = score.date || 'N/A'
				
				mailMergeByTeam[teamName].emailBody += "\n\n"
				mailMergeByTeam[teamName].emailBody += `Team: ${receivingTeam}\n`
				mailMergeByTeam[teamName].emailBody += `Round: ${round}\n`
				mailMergeByTeam[teamName].emailBody += `Day: ${day}`
				
				mailMergeByTeam[teamName].emailBody += SCORE_KEYS.reduce((str, scoreKey) =>
					str + `\n${SCORE_KEYS_TO_COLUMN_HEADING[scoreKey]}: ${score[scoreKey]}`, ''
				)
				mailMergeByTeam[teamName].emailBody += SCORE_KEYS.reduce((str, scoreKey) => {
					if (score.comments && score.comments[scoreKey] !== '') {
						return str + `\n${SCORE_KEYS_TO_COLUMN_HEADING[scoreKey]} Comments: ${score.comments[scoreKey]}`;
					}
					return str;
				}, '');
			})
		}
				
		const selfScores = selfScoreTeamData[teamName]
		if (selfScores && selfScores.length) {
			mailMergeByTeam[teamName].emailBody += '\n\n\n\n'
			mailMergeByTeam[teamName].emailBody += 'Scores given to you by your own team:'
			
			selfScores.forEach(score => {
				const opponent = score.opponent
				const round = score.round || 'N/A'
				const day = score.date || 'N/A'
				mailMergeByTeam[teamName].emailBody += "\n\n"
				mailMergeByTeam[teamName].emailBody += `Team: ${opponent}\n`
				mailMergeByTeam[teamName].emailBody += `Round: ${round}\n`
				mailMergeByTeam[teamName].emailBody += `Day: ${day}`
				mailMergeByTeam[teamName].emailBody += SCORE_KEYS.reduce((str, scoreKey) =>
					str + `\n${SCORE_KEYS_TO_COLUMN_HEADING[scoreKey]}: ${score[scoreKey]}`, ''
				)
				mailMergeByTeam[teamName].emailBody += SCORE_KEYS.reduce((str, scoreKey) => {
					if (score.comments && score.comments[scoreKey] !== '') {
						return str + `\n${SCORE_KEYS_TO_COLUMN_HEADING[scoreKey]} Comments: ${score.comments[scoreKey]}`;
					}
					return str;
				}, '');
			})
		}

		mailMergeByTeam[teamName].emailBody += `\n\n${footer}`
	})
	
	return mailMergeByTeam
}

function importMailMergeIntoSheet(mailMergeData) {
	let mailMergePreviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mail Merge Preview')
	if (mailMergePreviewSheet) {
		mailMergePreviewSheet.clear()
	} else {
		mailMergePreviewSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Mail Merge Preview')
	}
	createColumnHeadings(mailMergePreviewSheet, MAIL_MERGE_COLUMN_HEADINGS)

	const rowData = Object.entries(mailMergeData).map(([teamName, { emailAddressesString, emailSubject, emailBody }]) => [
		emailAddressesString,
		emailSubject,
		emailBody
	])
	const numRows = rowData.length
	const numColumns = MAIL_MERGE_COLUMN_HEADINGS.length
	mailMergePreviewSheet.getRange(2, 1, numRows, numColumns).setValues(rowData)
}

function sendEmails() {
	const ui = SpreadsheetApp.getUi()

  // Check for placeholder name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel');
  const nameCell = sheet.getRange('B38');
  const nameValue = nameCell.getValue().trim();
  if (/^e\.?g\.?/i.test(nameValue)) {
    ui.alert('Please add name', 'Before sending emails, please put the Spirit Director\'s name in cell B38.\nThen regenerate the Team Data & Mail Merge sheets.', ui.ButtonSet.OK);
    return;
  }

  // Make sure the person really meant to send.
	const result = ui.alert('Please confirm', 'Are you sure you want to send emails to your teams?', ui.ButtonSet.YES_NO)
	if (result != ui.Button.YES) {
		return
	}
	const mailMergePreviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mail Merge Preview')
	const emailRows = mailMergePreviewSheet.getDataRange().getValues()
	emailRows.shift()
	emailRows
		.filter(row => row[MAIL_MERGE_ENUM['Email Addresses']])
		.forEach(row => {
			try {
				GmailApp.sendEmail(
					row[MAIL_MERGE_ENUM['Email Addresses']],
					row[MAIL_MERGE_ENUM['Subject']],
					row[MAIL_MERGE_ENUM['Body']]
				)
				log(`sent email with subject '${row[MAIL_MERGE_ENUM['Subject']]}' to ${row[MAIL_MERGE_ENUM['Email Addresses']]}`)
			} catch (e) {
				log(`FAILED sending email with subject '${row[MAIL_MERGE_ENUM['Subject']]}' to ${row[MAIL_MERGE_ENUM['Email Addresses']]}:\n${e}`)
			}
		})
	const controlPanel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel')
	controlPanel.getRange('EmailsLastSent').setValue(formatDate(new Date(Date.now())))
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

			if (!missedTeams[teamName][opponentName]) {
				missedTeams[teamName][opponentName] = {
					scoresNeededFor: 0,
					scoresNeededFrom: 0,
					scoresFor: 0
				}
			}

			if (scoresNeededFor > 0) {
				missedTeams[teamName][opponentName].scoresNeededFor = scoresNeededFor
			} else if (scoresNeededFrom > 0) {
				missedTeams[teamName][opponentName].scoresNeededFrom = scoresNeededFrom
			}
			missedTeams[teamName][opponentName].scoresFor = quantity.scoresFor
		})
	})

	return getMissedTeamsAsString(missedTeams)
}

function getTeamsScoredAsString(teamsScoredObj) {
	teamsScored = ''
	Object.keys(teamsScoredObj).sort().forEach(opponentName => {
		let numScoresForOpponent = teamsScoredObj[opponentName]
		teamsScored += `${opponentName} (${numScoresForOpponent})\n`
	})
	teamsScored = teamsScored.substring(0, teamsScored.length - 1)
	return teamsScored
}

function getMissedTeamsAsString(missedTeamsObj) {
	let missedTeams = {}
	Object.keys(missedTeamsObj).forEach(teamName => {
		missedTeams[teamName] = {
			scoresNeededFor: '',
			scoresNeededFrom: '',
			scoresFor: ''
		}
		Object.keys(missedTeamsObj[teamName]).sort().forEach(opponentName => {
			let scoresNeededFor = missedTeamsObj[teamName][opponentName].scoresNeededFor
			let scoresNeededFrom = missedTeamsObj[teamName][opponentName].scoresNeededFrom
			let scoresFor = missedTeamsObj[teamName][opponentName].scoresFor
			if (scoresNeededFor) {
				missedTeams[teamName].scoresNeededFor += `${opponentName} (${scoresNeededFor})\n`
			} else if (scoresNeededFrom) {
				missedTeams[teamName].scoresNeededFrom += `${opponentName} (${scoresNeededFrom})\n`
			}

			if (scoresFor > 0) {
				missedTeams[teamName].scoresFor += `${opponentName} (${scoresFor})\n`
			}
		})
		let scoresNeededFor = missedTeams[teamName].scoresNeededFor
		let scoresNeededFrom = missedTeams[teamName].scoresNeededFrom
		let scoresFor = missedTeams[teamName].scoresFor
		missedTeams[teamName].scoresNeededFor = scoresNeededFor.substring(0, scoresNeededFor.length - 1)
		missedTeams[teamName].scoresNeededFrom = scoresNeededFrom.substring(0, scoresNeededFrom.length - 1)
		missedTeams[teamName].scoresFor = scoresFor.substring(0, scoresFor.length - 1)
	})
	return missedTeams
}

function createColumnHeadings(sheet, columnHeadings) {
	let numColumns = columnHeadings.length
	sheet.getRange(`R1C1:R1C${numColumns}`).setValues([columnHeadings])
}

function getFirstEmptyRow(sheet) {
	return sheet.getLastRow() + 1
}

function setListItemChoices(listItem, arr) {
	listItem.setChoices(arr.map(a => listItem.createChoice(a)))
}

function getTeamsFromURL(url) {
	let http = UrlFetchApp.fetch(url)
	let html = http.getContentText()
	return getTeamsFromString(html)
}

function getTeamsFromString(s) {
	let teams = new Set()
	let lines = s.split('\n')
	for (let line of lines) {
		let match = line.match(/<a href=\"\/events\/teams\/\?.*>(.*)\([0-9]+\)<\/a>/)
		if (match && match.length) {
			team = match[1]
			teams.add(team)
		}
	}
	return teams
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

function resetRowsToModalHeight(sheet, startRow = 2) {
  const numRows = sheet.getMaxRows() - startRow + 1;
  if (numRows <= 0) return;

  const heights = Array.from(
    { length: numRows },
    (_, i) => sheet.getRowHeight(startRow + i)
  );

  const modalHeight = mode(heights);

  sheet.setRowHeights(startRow, numRows, modalHeight);
}

function mode(values) {
  if (!values.length) {
    throw new Error('Cannot calculate mode of an empty array.');
  }

  const counts = new Map();
  let bestValue = values[0];
  let bestCount = 0;

  for (const value of values) {
    const count = (counts.get(value) || 0) + 1;
    counts.set(value, count);

    if (count > bestCount) {
      bestValue = value;
      bestCount = count;
    }
  }

  return bestValue;
}

function getScoreSheetConfig(sheet) {
  if (sheet.getName() === PROCESSED_SCORE_SHEET_NAME) {
    return {
      enum: PROCESSED_SCORE_ENUM,
      headings: PROCESSED_SCORE_COLUMN_HEADINGS
    }
  }

  return {
    enum: RAW_SCORE_ENUM,
    headings: RAW_SCORE_COLUMN_HEADINGS
  }
}
