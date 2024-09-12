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

const SHEETS_TO_REMAKE = ['_rawScores']

let errno = 0

function onFormSubmit(e) {
	log('running onFormSubmit()')
	let thisFileId = SpreadsheetApp.getActiveSpreadsheet().getId()
	let thisFile = DriveApp.getFileById(thisFileId)
	let parentFolder = thisFile.getParents().next()
	createFolder(parentFolder, 'Tournaments')
	let tournamentsFolder = parentFolder.getFoldersByName('Tournaments').next()
	let itemResponses = e.response.getItemResponses()
	let rawResponses = itemResponses.map(item => item.getResponse())
	let tournamentName = rawResponses[0]
	let emailResponse = rawResponses[1]

	let tournamentFolder = createFolder(tournamentsFolder, tournamentName)
	let templateFolderContents = getTemplateFolderContents(parentFolder)
	copyFilesToFolder(templateFolderContents, tournamentFolder)
	let emails = getEmailsFromResponse(emailResponse)
	if (emails.length) {
		tournamentFolder.addEditors(emails)
	}
	log('onFormSubmit() success!')
}

function initializeForm() {
	log('running initializeForm()')
	let spr = SpreadsheetApp.getActiveSpreadsheet()
	let formUrl = spr.getRange('B1').getValue()
	let form = FormApp.openByUrl(formUrl)
	linkSheetToForm(form, spr, 'Responses')
	createFormSubmissionTrigger(form)
	log('initializeForm() success!')
}

function copyFilesToFolder(fileIterator, folder) {

	// don't copy any files if the folder is not empty
	if (folder.getFiles().hasNext()) {
		return
	}

	let formUrl
	let controlPanelSpreadsheet
	while (fileIterator.hasNext()) {
		let file = fileIterator.next()
		let mimeType = file.getMimeType()
		let name
		if (mimeType === MimeType.GOOGLE_SHEETS) {
			name = folder.getName() + ' Spirit Scorekeeping'
		} else if (mimeType === MimeType.GOOGLE_FORMS) {
			name = folder.getName() + ' Spirit Scoring Form'
		} else {
			name = file.getName()
		}
		log(`creating ${name}`)
		let newFile = file.makeCopy(name, folder)

		// if the file is the form, set its title
		if (mimeType === MimeType.GOOGLE_FORMS) {
			FormApp.openById(newFile.getId()).setTitle(name)
			formUrl = newFile.getUrl()
		} else if (mimeType === MimeType.GOOGLE_SHEETS) {
			controlPanelSpreadsheet = SpreadsheetApp.openById(newFile.getId())
		}
	}

	linkSheetToForm(FormApp.openByUrl(formUrl), controlPanelSpreadsheet, 'Raw Scores', RAW_SCORE_COLUMN_HEADINGS)
	createColumnHeadings(controlPanelSpreadsheet.getSheetByName('Raw Scores'), RAW_SCORE_COLUMN_HEADINGS)

	// refresh sheet references in the formulas in these sheets, because the raw scores sheet didn't exist before,
	//  causing the references to break until we refresh them
	remakeSheets(controlPanelSpreadsheet, SHEETS_TO_REMAKE)
}

function createColumnHeadings(sheet, columnHeadings) {
	let numColumns = columnHeadings.length
	sheet.getRange(`R1C1:R1C${numColumns}`).setValues([columnHeadings])
}

function remakeSheets(spreadsheet, sheetNames) {
	sheetNames.forEach(sheetName => {
		let sheet = spreadsheet.getSheetByName(sheetName)

		let range = sheet.getDataRange()
		let formulas = range.getFormulas()
		range.clearContent()
		range.setFormulas(formulas)
	})
}

function getEmailsFromResponse(rawResponse) {
	if (rawResponse.trim() === '') {
		return []
	}
	return rawResponse.split('\n').map(line => line.trim()).filter(email => email)
}

function createFolder(parentFolder, name) {
	let existingFolders = parentFolder.getFoldersByName(name)
	if (existingFolders.hasNext()) {
		return existingFolders.next()
	}
	return parentFolder.createFolder(name)
}

function getTemplateFolderContents(parentFolder) {
	let spiritScoresFolder = parentFolder.getParents().next().getParents().next()
	while (!spiritScoresFolder.getFoldersByName('Templates (do not edit contents)').hasNext()) {
		spiritScoresFolder = spiritScoresFolder.getParents().next()
	}
	let templateFolder = spiritScoresFolder.getFoldersByName('Templates (do not edit contents)').next().getFoldersByName('Tournament Templates').next()
	return templateFolder.getFiles()
}

function linkSheetToForm(form, spr, responseSheetName, responseColumnHeadings) {
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
			sheet.setName(responseSheetName)
			if (responseColumnHeadings) {
				let numCols = responseColumnHeadings.length
				let range = sheet.getRange(1, 1, 1, numCols)
				range.setValues([responseColumnHeadings])
				range.setFontWeight('bold')
				range.setWrap(true)
			}
			break
		}
	}
	SpreadsheetApp.flush()
}

function createFormSubmissionTrigger(form) {
	let triggers = ScriptApp.getUserTriggers(form)
	for (let i = 0; i < triggers.length; i++) {
		let eventType = triggers[i].getEventType()
		let handlerFunction = triggers[i].getHandlerFunction()
		if (eventType === ScriptApp.EventType.ON_FORM_SUBMIT && handlerFunction === 'onFormSubmit') {
			return
		}
	}
	ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create()
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
