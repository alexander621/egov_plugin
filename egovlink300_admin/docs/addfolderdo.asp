<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addfolderdo.asp
' AUTHOR: Steve Loar
' CREATED: 08/30/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Creates new folders for documents.
'
' MODIFICATION HISTORY
' 1.0   08/30/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sParentFolder, sFolderName, sPhysicalPath, objFSO, sSql, sFolderPath, iParentFolderId

' initially the paths are set for physically checking and creating the folder
sParentFolder = dbsafe(request("path"))	' /egovlink300_docs/custom/pub/eclink/published_documents/
'response.write sParentFolder & "<br /><br />"

sFolderName = dbsafe(trim(request("foldername")))

sFolderPath = sParentFolder & sFolderName

'sPhysicalPath = "e:" & Replace(sFolderPath, "/", "\") ' This is the physical folder we need to check and make
sPhysicalPath = Application("DocumentsDrive") & Replace(sFolderPath, "/", "\") ' This is the physical folder we need to check and make

' now alter the paths for the DB Insert
'sFolderPath = Replace(sFolderPath, "egovlink300_docs", "public_documents300")
sFolderPath = Replace(sFolderPath, Application("DocumentsRootDirectory"), "public_documents300")

'sParentFolder = Replace(sParentFolder, "egovlink300_docs", "public_documents300")
sParentFolder = Replace(sParentFolder, Application("DocumentsRootDirectory"), "public_documents300")

iParentFolderId = GetFolderId( Left(sParentFolder, (Len(sParentFolder) - 1)) )
	
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
If Not objFSO.FolderExists( sPhysicalPath ) Then
	objFSO.CreateFolder( sPhysicalPath )

	sSql = "INSERT INTO DocumentFolders ( OrgID, CreatorUserID, FolderName, FolderPath ,ParentFolderID) "
	sSql = sSql & "VALUES ( " & session("orgid") & ", " & session("userid") & ", '" & sFolderName & "', '"
	sSql = sSql & sFolderPath & "', " & iParentFolderId & " )"
	'response.write sSql & "<br /><br />"

	RunSQLStatement sSql

	sSuccessFlag = "fa"
Else
	sSuccessFlag = "fe"
End If

Set objFSO = Nothing 

'response.write "sSuccessFlag = " & sSuccessFlag & "<br /><br />"
' Return to the add folder page
response.redirect "addfolder.asp?path=" & request("path") & "&foldername=" & sFolderName & "&sf=" & sSuccessFlag




%>
