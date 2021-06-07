<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: renamedocumentdo.asp
' AUTHOR: Steve Loar
' CREATED: 09/07/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Renames a document.
'
' MODIFICATION HISTORY
' 1.0   09/07/2010	Steve Loar - INITIAL VERSION
' 1.1	06/23/2011	Steve Loar - Wrapped the table update with dbsafe() calls around the strings
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sNewFileName, sTargetFile, iFileId, sFileName, sPhysicalPath, sDBFile, oFile
Dim sNewPhysicalPath, sNewDBFile, sSql, sNewName, re, IsValidName


sTargetFile = dbsafe(request("path"))

iFileId = CLng(request("fileid"))
sFileName = Trim(request("filename"))
sNewName = Trim(request("NewName"))

' Block San Luis from inputing bad file names by disabling their JavaScript
Set re = new RegExp
re.IgnoreCase = false
re.global = false
re.Pattern = "^[A-Za-z0-9 _-]+\.{1}[A-Za-z0-9]{2}[A-Za-z0-9]{0,2}$"
IsValidName = re.Test(sNewName)

If Not IsValidName Then
	response.redirect "renamedocument.asp?path=" & request("path") & "&filename=" & sFileName & "&fileid=" & iFileId & "&sf=bf"
End If 

'sPhysicalPath = "e:" & Replace(sTargetFile, "/", "\") ' This is the physical file we need to check and rename
sPhysicalPath = Application("DocumentsDrive") & Replace(sTargetFile, "/", "\") ' This is the physical file we need to check and rename

'sDBFile = Replace(sTargetFile, "egovlink300_docs", "public_documents300")
sDBFile = Replace(sTargetFile, Application("DocumentsRootDirectory"), "public_documents300")

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists( sPhysicalPath ) Then
	
	Set oFile = objFSO.GetFile(sPhysicalPath)
	sNewPhysicalPath = Left(sPhysicalPath,Len(sPhysicalPath) - Len(sFileName)) & sNewName
	sNewDBFile = Left(sDBFile,Len(sDBFile) - Len(sFileName)) & sNewName

	If objFSO.FileExists(sNewPhysicalPath) Then
		sSuccessFlag = "fe" ' file by that name exists
	Else
		'response.write sNewPhysicalPath & "<br />"
		' Rename the physical file
		oFile.Move sNewPhysicalPath

		' Do the DB rename
		sSql = "UPDATE Documents SET "
		sSql = sSql & " DocumentTitle = '" & dbsafe( sNewName ) & "', "
		sSql = sSql & " DocumentURL = '" & dbsafe( sNewDBFile ) & "' "
		sSql = sSql & " WHERE DocumentID = " & iFileId
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		sSuccessFlag = "fr"
	End If 

	Set oFile = Nothing

Else
	sSuccessFlag = "nr" ' original file not found
End If 

Set objFSO = Nothing 

'response.write "sSuccessFlag = " & sSuccessFlag & "<br /><br />"

If sSuccessFlag= "fe" Then
	' file by that name exists so go back to the rename page
	response.redirect "renamedocument.asp?path=" & request("path") & "&filename=" & sFileName & "&fileid=" & iFileId & "&sf=" & sSuccessFlag
Else
	' successful rename or file deleted so go to the default page 
	response.redirect "default.asp?filename=" & sFileName & "&newname=" & sNewName & "&sf=" & sSuccessFlag
End If 


%>