let errno = 0

function updateForm() {
	log('running updateForm()')
	let spr = SpreadsheetApp.getActiveSpreadsheet()
	let tournamentURL = spr.getRange('B1').getValue()
	let teamSet = getTeamsFromURL(tournamentURL)
	let teams = Array.from(teamSet).sort()
	let rawScoreSheet = spr.getSheetByName('Raw Scores')
	let form = FormApp.openByUrl(rawScoreSheet.getFormUrl())
	let listItems = form.getItems(FormApp.ItemType.LIST)
	let yourTeamItem = listItems[0].asListItem()
	let opponentTeamItem = listItems[1].asListItem()

	setListItemChoices(yourTeamItem, teams)
	setListItemChoices(opponentTeamItem, teams)

	spr.getRange('A13').setValue('Teams last updated on form:')
	spr.getRange('B13').setValue(formatDate(new Date(Date.now())))
	log('updateForm() success!')
}

function initializeForm() {
	log('running initializeForm()')
	let spr = SpreadsheetApp.getActiveSpreadsheet()
	shareSpreadsheet(spr)
	let formURL = spr.getRange('B2').getValue()
	let form = FormApp.openByUrl(formURL)
	linkSheetToForm(form, spr)

	let bankUrl = spr.getRange('B3').getValue()
	linkSheetToBank(bankUrl, spr)

	if (errno) {
		log('initializeForm() completed with one or more errors.')
	} else {
		log('initializeForm() success!')
	}
}

function linkSheetToBank(bankUrl, spr) {
	let bankSheet = SpreadsheetApp.openByUrl(bankUrl).getSheetByName('Tournament Control Panel IDs')
	let tournamentSheetId = spr.getId()
	if (sheetIdInBank(tournamentSheetId, bankSheet)) {
		return
	}
	let row = getFirstEmptyRow(bankSheet)
	bankSheet.getRange(row, 1).setValue(tournamentSheetId)
}

function sheetIdInBank(tournamentSheetId, bankSheet) {
	let existingIds = new Set(bankSheet.getRange('A:A').getValues().map(row => row[0]).filter(id => id.trim() != ''))
	return existingIds.has(tournamentSheetId)
}

function shareSpreadsheet(spr) {
	let emails = getEmailAddresses(spr)
	let spreadsheetId = spr.getId()
	for (email of emails) {
		try {
			DriveApp.getFileById(spreadsheetId).addViewer(email)
			log(`File successfully shared with ${email}.`)
		} catch (e) {
			log(`Error sharing with ${email}. ${e.name}: ${e.message}`)
			errno |= 1
		}
	}
}

function getEmailAddresses(spr) {
	let controlPanelSheet = spr.getSheetByName('Control Panel')
	let values = controlPanelSheet.getRange('E:E').getValues().map(row => row[0]).slice(1)
	let emails = values.filter(value => value && value.trim() != '')
	return emails
}

function linkSheetToForm(form, spr) {
	let formDestId
	try {
		formDestId = form.getDestinationId()
	}
	catch (e) {
		formDestId = null
	}
	if (formDestId != spr.getId()) {
		form.setDestination(FormApp.DestinationType.SPREADSHEET, spr.getId())
	}
	SpreadsheetApp.flush()
	for (sheet of spr.getSheets()) {
		if (sheet.getFormUrl() != null) {
			sheet.setName('Raw Scores')
			break
		}
	}
}

function copyResponseToSheet(response, sheet) {
	let itemResponses = response.getItemResponses()
	let responseLength = itemResponses.length
	let rawResponses = itemResponses.map(item => item.getResponse())
	let firstEmptyRow = getFirstEmptyRow(sheet)
	let range = `R${firstEmptyRow}C1:R${firstEmptyRow}C${responseLength+1}`
	sheet.getRange(range).setValues([[response.getTimestamp(), ...rawResponses]])
}

function getFirstEmptyRow(sheet) {
	let column = sheet.getRange('A:A')
	let values = column.getValues()
	let ct = 0
	while (values[ct] && values[ct][0] != '') {
		ct++
	}
	return (ct+1)
}

function setListItemChoices(listItem, arr) {
	listItem.setChoices(arr.map(a => listItem.createChoice(a)))
}

function getItemByTitle(title, form) {
	let items = form.getItems()
	for (let item of items) {
		if (item.getTitle() == title) {
			return item
		}
	}
	return null
}

function getTeamsFromDoc(doc_id) {
	let doc = DocumentApp.openById(doc_id)
	let html = doc.getBody().getText()
	return getTeamsFromString(html)
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
	let cellContents = `${omitDate ? '' : timeStamp} ${String(obj)}\n${cellContents}`
	range.setValue(cellContents)
}