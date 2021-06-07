<%

sub UploadFile(ByRef formFields, strFolderLocation, blnOverwrite, blnCustomErrors)
	set objFso = CreateObject("Scripting.FileSystemObject") 
	Server.ScriptTimeout = 36000
	intUploadLimit = 209715200 ' 200MB + (52428800 is 50MB)
	strUploadLimittxt = "200 MB"
	strTempPath = "E:\egovdoctemp\"
	blnIsPhysicalPath = false
	if instr(strFolderLocation,"[") = 0 then strFolderLocation = Server.MapPath(strFolderLocation)

	'Initiate the object and set max upload size
	Set objUpload = Server.CreateObject("Dundas.Upload.2")
	objUpload.MaxUploadSize = intUploadLimit
	
	
	'Save file(s) to temp directory
	on error resume next
		objUpload.Save strTempPath

		if Err.number <> 0 then
			if not blnCustomErrors then
				response.write "Your request exceeded the system's upload size limit of " & strUploadLimittxt & ".  Please reduce your file's size and try again."
				response.write err.number
				response.write "<br />" & strTempPath
				response.end
			else
				formFields.Add "sSuccess","tb"
			end if
		end if
	on error goto 0


	'Add Filecount to dictionary
	formFields.Add "filecount", objUpload.Files.Count

	'Load Form Fields into dictionary
	For Each objFormItem In objUpload.Form
		formFields.Add objFormItem & "", objFormItem.Value & ""
		'response.write "ADDED Name: " & objFormItem & ", Value: " & objFormItem.Value & "<br />"
	Next



	'Override folder location with form field
	if instr(strFolderLocation,"[FORM]") > 0 then
		strFolderLocation = formFields(replace(strFolderLocation,"[FORM]",""))
	elseif instr(strFolderLocation,"[FORMPHYSICAL]") > 0 then
		blnIsPhysicalPath = true
		strFolderLocation = "e:" & formFields(replace(strFolderLocation,"[FORMPHYSICAL]",""))
	end if

	'Override file overwrite
	if formFields("chkOverwrite") = "on" then blnOverwrite = true



	'Process Uploaded file(s)
	intFileCount = 0
	For Each objUploadedFile in objUpload.Files
		intFileCount = intFileCount + 1

		tempFileName = objUploadedFile.Path
		origFileName = objUploadedFile.OriginalPath
 		origFileName = RIGHT(origFileName,LEN(origFileName) - instrrev(origFileName,"\"))
	
	
		'Add filename to dictionary
		key = "filename"
		if objUpload.Files.Count > 1 then key = key & intFileCount
		formFields.Add key, origFileName


		'Add filesize to dictionary
		key = "filesize"
		if objUpload.Files.Count > 1 then key = key & intFileCount
		formFields.Add key, objUploadedFile.Size
	

		'Determine where the file will go
		strFullPath = strFolderLocation & "\" & origFileName

		'Delete existing file if overwriting
		If objUpload.FileExists( strFullPath ) and blnOverwrite Then objUpload.FileDelete( strFullPath )

		'write file if file doesn't exist
		If not objUpload.FileExists( strFullPath ) Then 
			mvsuccessful = false
			iCount = 0
			Do while not mvsuccessful and iCount < 1000
				on error resume next
				objFso.MoveFile tempFileName, strFullPath
				if Err.number = 0 then mvsuccessful = true
				on error goto 0
				iCount = iCount + 1
			loop
			if mvsuccessful then
				formFields.Add "sSuccess","true"
			else
				if not blnCustomErrors then
					response.write "Sorry, your file could not be uploaded.  Please contact support and report that your file could not be moved."
					response.end
				else
					formFields.Add "sSuccess","nm"
				end if
			end if
		else
			objFso.DeleteFile tempFileName

			if not blnCustomErrors then
				response.write "Sorry, this file already exists. Please either choose to overwrite the existing file or choose a new file name."
				response.end
			else
				formFields.Add "sSuccess","df"
			end if
		end if

	next




	Set objUpload = Nothing
	set objFso = Nothing 


end sub
%>
