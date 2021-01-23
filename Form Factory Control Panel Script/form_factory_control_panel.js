var errno = 0

function onFormSubmit(e) {
	log('running onFormSubmit()')
	let thisFileId = SpreadsheetApp.getActiveSpreadsheet().getId()
	let thisFile = DriveApp.getFileById(thisFileId)
	let parentFolder = thisFile.getParents().next()
	let itemResponses = e.response.getItemResponses()
	let rawResponses = itemResponses.map(item => item.getResponse())
	let tournamentName = rawResponses[0]
	let emailResponse = rawResponses[1]

	let tournamentFolder = createFolder(parentFolder, tournamentName)
	let templateFolderContents = getTemplateFolderContents(parentFolder)
	copyFilesToFolder(templateFolderContents, tournamentFolder)
	let emails = getEmailsFromResponse(emailResponse)
	tournamentFolder.addEditors(emails)
	log('onFormSubmit() success!')
}

function initializeForm() {
	log("running initializeForm()")
	var spr = SpreadsheetApp.getActiveSpreadsheet()
	var formUrl = spr.getRange("B1").getValue()
	var form = FormApp.openByUrl(formUrl)
	linkSheetToForm(form, spr, 'Responses')
	createFormSubmissionTrigger(form)
	log("initializeForm() success!")
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
			name = folder.getName() + ' Spirit Score Control Panel'
		} else if (mimeType === MimeType.GOOGLE_FORMS) {
			name = folder.getName() + ' Spirit Score Form'
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

	linkSheetToForm(FormApp.openByUrl(formUrl), controlPanelSpreadsheet, 'Raw Scores')
}

function getEmailsFromResponse(rawResponse) {
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
	let templateFolder = parentFolder.getParents().next().getFoldersByName('Templates').next()
	return templateFolder.getFiles()
}

function linkSheetToForm(form, spr, responseSheetName) {
	var formDestId
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
			break
		}
	}
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
	var month = String(date.getMonth() + 1).padStart(2, '0')
	var dateInMonth = String(date.getDate()).padStart(2, '0')
	var year = String((date.getYear()) + 1900).padStart(2, '0')
	var hours = String(date.getHours()).padStart(2, '0')
	var minutes = String(date.getMinutes()).padStart(2, '0')
	var seconds = String(date.getSeconds()).padStart(2, '0')
	return `${month}/${dateInMonth}/${year} ${hours}:${minutes}:${seconds}`
}

function log(obj, omitDate) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log")
	var range = sheet.getRange("A1")
	var cellContents = range.getValue()
	var now = new Date(Date.now())
	var timeStamp = `[${formatDate(now)}]`
	var cellContents = `${omitDate ? "" : timeStamp} ${String(obj)}\n${cellContents}`
	range.setValue(cellContents)
}
