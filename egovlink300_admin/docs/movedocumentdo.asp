<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: movedocumentdo.asp
' AUTHOR: Steve Loar
' CREATED: 09/09/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Moves a document.
'
' MODIFICATION HISTORY
' 1.0   09/09/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sNewFileName, sTargetFile, iFileId, sFileName, sPhysicalPath, sDBFile, oFile
Dim sNewPhysicalPath, sNewDBFile, sSql, iFolderId, objFSO, sNewDBFolderPath


sTargetFile = dbsafe(request("path"))
iFileId = CLng(request("fileid"))
sFileName = Trim(request("filename"))
iFolderId = CLng(request("folderid"))

sPhysicalPath = Application("DocumentsDrive") & Replace(sTargetFile, "/", "\") ' This is the physical file we need to check and move

sNewDBFolderPath = GetFolderPath( iFolderId )
sNewDBFile = sNewDBFolderPath & "/" & sFileName

'sNewPhysicalPath = Replace(sNewDBFile, "public_documents300", "egovlink300_docs")
sNewPhysicalPath = Replace(sNewDBFile, "public_documents300", Application("DocumentsRootDirectory"))
sNewPhysicalPath = Application("DocumentsDrive") & Replace(sNewPhysicalPath, "/", "\") ' Where we are moving it to

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists( sPhysicalPath ) Then
	
	Set oFile = objFSO.GetFile(sPhysicalPath)

	If objFSO.FileExists(sNewPhysicalPath) Then
		sSuccessFlag = "fe" ' file by that name exists
	Else
		' Rename the physical file
		oFile.Move sNewPhysicalPath

		' Do the DB rename
		sSql = "UPDATE Documents SET "
		sSql = sSql & " DocumentURL = '" & sNewDBFile & "', "
		sSql = sSql & " ParentFolderID = " & iFolderId 
		sSql = sSql & " WHERE DocumentID = " & iFileId
		response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		sSuccessFlag = "fm"
	End If 

	Set oFile = Nothing

Else
	sSuccessFlag = "nm" ' original file not found
End If 

Set objFSO = Nothing 

response.write "sSuccessFlag = " & sSuccessFlag & "<br /><br />"

If sSuccessFlag= "fe" Then
	' file by that name exists so go back to the move page
	response.redirect "movedocument.asp?path=" & request("path") & "&filename=" & sFileName & "&fileid=" & iFileId & "&folderid="& iFolderId & "&sf=" & sSuccessFlag
Else
	' successful move or file deleted so go to the default page 
	response.redirect "default.asp?filename=" & sFileName & "&sf=" & sSuccessFlag
End If 



%>