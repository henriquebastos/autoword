-- autoword.applescript
-- autoword
--  Created by Henrique Bastos on 10/04/09.

property templates : {"letter-template.docx"}
property tempDir : "tmp:"

-- Dynamically generate the destination filename
on generateFilename(templateFile, clientName)
	set sep to "."
	set today to current date
	set datePart to (year of today) & (month of today) & (day of today) as text
	set clientPart to ""
	if clientName is not equal to "" then set clientPart to first word of clientName
	set filename to tempDir & datePart & sep & clientPart & sep & first word of templateFile & ".pdf"
	return filename
end generateFilename

-- Create new doc from template, update dynamic values and save a pdf
on buildDocument(templateFile, rec)
	set pdffile to generateFilename(templateFile, (get clientName of rec))
	tell application "Microsoft Word"
		set doc to (create new document attached template templateFile)
		set variable value of (get variable "clientName" of doc) to clientName of rec
		set variable value of (get variable "clientStreet" of doc) to clientStreet of rec
		set variable value of (get variable "clientProvince" of doc) to clientProvince of rec
		set variable value of (get variable "clientCity" of doc) to clientCity of rec
		set variable value of (get variable "clientState" of doc) to clientState of rec
		set variable value of (get variable "clientZip" of doc) to clientZip of rec
		set variable value of (get variable "clientOrder" of doc) to clientOrder of rec
		repeat with i from 1 to number of fields in doc
			set thisField to field i of doc
			update field thisField
		end repeat
		save as doc file name pdffile file format format PDF
		close active window of doc saving no
	end tell
end buildDocument

-- Open Finder on the destination path
on showGeneratedFiles()
	set destFolder to POSIX file "/tmp"
	tell application "Finder" to open destFolder
end showGeneratedFiles

-- Handles the generate button click building our documents from form values
on clicked theObject
	set rec to {clientName:"", clientStreet:"", clientZip:"", clientProvince:"", clientCity:"", clientState:"", clientOrder:""}
	tell (window of theObject)
		set clientName of rec to contents of text field "clientName"
		set clientStreet of rec to contents of text field "clientStreet"
		set clientZip of rec to contents of text field "clientZip"
		set clientProvince of rec to contents of text field "clientProvince"
		set clientCity of rec to contents of text field "clientCity"
		set clientState of rec to contents of text field "clientState"
		set clientOrder of rec to contents of text field "clientOrder"
	end tell
	repeat with doc in templates
		buildDocument(doc, rec)
	end repeat
	showGeneratedFiles()
end clicked
