<%
Server.ScriptTimeout = 5000
%>
<!-- #include file="../includes/adovbs.inc" -->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: docsync.asp
' AUTHOR: Steve Loar
' CREATED: 11/07/2007
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds new documents to the docs tables in a non-destructive way.
'
' MODIFICATION HISTORY
' 1.0   11/07/2007	Steve Loar - Initial Version
' 1.1	06/23/2011	Steve Loar - Added more dfsave calls for file names with apostrophes
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sCity, iDocOrgID, sPath, sVirtualPath, iUserid, FSO, iFolderId

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

iDocOrgID = CLng(request("orgid") )
sCity = GetCityVirtualDirectory( iDocOrgID )
'response.write sCity & "<br />"
iFolderId = 0

If sCity = "" Then
	' no virtual directory, so bail on this
	response.redirect "../admin/manage_features.asp?orgid=" & iDocOrgID
End If 

' Production path
'sPath = "C:\wwwroot\www.cityegov.com\egovlink300_admin\custom\pub\" & sCity
'sPath = "E:\egovlink300_docs\custom\pub\" & sCity
sPath = Application("DocumentsDrive") & "\" & Application("DocumentsRootDirectory") & "\custom\pub\" & sCity
' Development path
'sPath = "C:\www_server_root\egovlink\egovlink release 4.0.0\egovlink300_admin\custom\pub\" & sCity

sVirtualPath = "/public_documents300/custom/pub/" & sCity
iUserid = session("userid")

' See if the root level folder exists and if not, create it in the table
If Not FolderInDB( iDocOrgID, sVirtualPath ) Then
	'response.write "Root Not Found<br />"
	iFolderId = AddFolder( iDocOrgID, sVirtualPath, " NULL ", sCity )
Else
	'response.write "Root Found<br />"
	iFolderId = GetFolderId( iDocOrgID, sVirtualPath )
End If 

'response.write "Root Folder iFolderId = " & iFolderId & "<br />"
'response.End 

Set FSO = CreateObject("Scripting.FileSystemObject")

' ENUMERATE FOLDERS AND FILES ADDING TO DATABASE the new files
SyncDocuments iDocOrgID, FSO.GetFolder(sPath), iFolderId

' Walk the docs and remove those not found in the FSO
cleanDocumentsTable iDocOrgID, sPath, sVirtualPath

' walk the folder and remove those not found in the FSO
cleanFoldersTable iDocOrgID, sPath, sVirtualPath 

Set FSO = Nothing 

' Back to the edit page
'response.redirect "../admin/manage_features.asp?orgid=" & iDocOrgID
response.redirect "../admin/featureselection.asp?orgid=" & iDocOrgID


'-------------------------------------------------------------------------------------------------------
' SUB SYNCDOCUMENTS(FOLDER)
'-------------------------------------------------------------------------------------------------------
Sub SyncDocuments( iDocOrgID, Folder, iParentFolderId )
	Dim iFolderId, sSubfolderName, bValidFolder
   
	For Each Subfolder In Folder.SubFolders
		sSubPath = replace(Subfolder.Path,sPath,"")
		sSubPath = replace(sSubPath,"\","/")
		sSubPath = sVirtualPath & sSubPath
		sSubfolderName = FSO.GetBaseName( Subfolder ) 
		
		If Not FolderInDB( iDocOrgID, sSubPath ) Then 
			If isValidFolder( sVirtualPath, sSubPath ) Then 
				bValidFolder = True 
				response.write "<b>Adding Folder: </b>" & sSubfolderName & "<br />"
				iFolderId = AddFolder( iDocOrgID, sSubPath, iParentFolderId, sSubfolderName )
			Else
				bValidFolder = False  
			End If
		Else
			bValidFolder = True 
			'response.write "<b>Found Folder: </b>" & sSubPath & "<br />"
			iFolderId = GetFolderId( iDocOrgID, sSubPath )
		End If 
		'response.write sSubfolderName & " - iFolderId = " & iFolderId & "<br />"
		
		If bValidFolder Then 
			For Each File In Subfolder.Files 
				sDocumentPath = sSubPath & "/" & File.Name 
				session("lastfile") = sDocumentPath
				If Not FileInDB( iDocOrgID, sDocumentPath ) Then 
					If isValidFileName( File.Name ) Then 
						'response.write  "<b>Adding File: </b>" &sDocumentPath & "<br />"
						AddDocument iDocOrgID, sDocumentPath, iParentFolderId, File.Name, File.Size
					End If 
				Else
					iDocumentId = GetDocumentId( iDocOrgID, sDocumentPath )
					UpdateFileSize iDocumentId, File.Size 
				End If 
			Next 
			
			SyncDocuments iDocOrgID, Subfolder, iFolderId
		End If 

    Next
End Sub


'--------------------------------------------------------------------------------------------------
' string GetCityVirtualDirectory( iDocOrgID )
'--------------------------------------------------------------------------------------------------
Function GetCityVirtualDirectory( ByVal iDocOrgID )
	Dim sSql, oRs

	sSql = "SELECT OrgVirtualSiteName FROM organizations WHERE orgid = " & iDocOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		GetCityVirtualDirectory = oRs("OrgVirtualSiteName")
	Else
		GetCityVirtualDirectory = ""
	End If
	
	oRs.close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean FolderInDB( iDocOrgID, sFolderPath )
'--------------------------------------------------------------------------------------------------
Function FolderInDB( ByVal iDocOrgID, ByVal sFolderPath )
	Dim sSql, oRs

	sSql = "SELECT FolderID FROM documentfolders WHERE folderpath = '" & DBsafe( sFolderPath ) & "' AND orgid = " & iDocOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		FolderInDB = True 
	Else
		FolderInDB = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean FileInDB( iDocOrgID, sDocumentPath )
'--------------------------------------------------------------------------------------------------
Function FileInDB( ByVal iDocOrgID, ByVal sDocumentPath )
	Dim sSql, oRs

	sSql = "SELECT documentid FROM documents WHERE documenturl = '" & DBsafe( sDocumentPath ) & "' AND orgid = " & iDocOrgID

	'session("sql") =  sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 0, 1
	'session("sql") =  ""
	
	If Not oRs.EOF Then
		FileInDB = True 
	Else
		FileInDB = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetDocumentId( iDocOrgID, sDocumentPath )
'--------------------------------------------------------------------------------------------------
Function GetDocumentId( ByVal iDocOrgID, ByVal sDocumentPath )
	Dim sSql, oRs

	sSql = "SELECT documentid FROM documents WHERE documenturl = '" & DBsafe(  sDocumentPath ) & "' AND orgid = " & iDocOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		GetDocumentId = oRs("documentid") 
	Else
		GetDocumentId = 0 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void UpdateFileSize iDocumentId, File.Size 
'--------------------------------------------------------------------------------------------------
Sub UpdateFileSize( ByVal iDocumentId, ByVal sSize )
	Dim sSql

	sSql = "UPDATE documents SET documentsize = " & sSize & " WHERE documentid = " & iDocumentId

	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetFolderId( iDocOrgID, sFolderPath )
'--------------------------------------------------------------------------------------------------
Function GetFolderId( ByVal iDocOrgID, ByVal sFolderPath )
	Dim sSql, oRs

	sSql = "SELECT FolderID FROM documentfolders WHERE folderpath = '" & DBsafe( sFolderPath ) & "' AND orgid = " & iDocOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then
		GetFolderId = CLng(oRs("FolderID"))
	Else
		GetFolderId = CLng(0) 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean UserIsRootAdmin( iUserID )
'--------------------------------------------------------------------------------------------------
Function UserIsRootAdmin( ByVal iUserID )
	Dim sSql, oUsers

	UserIsRootAdmin = False 
	sSql = "Select isnull(isrootadmin,0) as isrootadmin from users where userid = " & iUserID

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSQL, Application("DSN"), 3, 1

	If NOT oUsers.EOF Then
		If oUsers("isrootadmin") Then 
			UserIsRootAdmin = True 
		End If 
	End If

	oUsers.close
	Set oUsers = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer AddFolder( iDocOrgID, sFolderPath, iParentFolderId, sFolderName )
'--------------------------------------------------------------------------------------------------
Function AddFolder( ByVal iDocOrgID, ByVal sFolderPath, ByVal iParentFolderId, ByVal sFolderName )
	Dim sSql, oInsert, iFolderId

	iFolderId = 0
	'sFolderName = Mid(sFolderPath, InStrRev(sFolderPath, "/")+1 )

	sSql = "INSERT INTO documentfolders ( orgid, CreatorUserID, FolderName, FolderPath ,ParentFolderID) VALUES ( "
	sSql = sSql & iDocOrgID & ", " & Session("UserID") & ", '" & DBsafe( sFolderName ) & "', '" & DBsafe( sFolderPath ) & "', " & iParentFolderId & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
	'response.write sSQL & "<br /><br />"
	'response.End 

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3

	iFolderId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	AddFolder = iFolderId

End Function 


'--------------------------------------------------------------------------------------------------
' void AddDocument iDocOrgID, sDocumentPath, iParentFolderId, sFileName, sFileSize 
'--------------------------------------------------------------------------------------------------
Sub AddDocument( ByVal iDocOrgID, ByVal sDocumentPath, ByVal iParentFolderId, ByVal sFileName, ByVal sFileSize )
	Dim oCmd, sSql

	sSql = "INSERT INTO documents ( orgid, documenturl, documenttitle, parentfolderid, creatoruserid, LinkTargetsNew, documentsize ) VALUES ( "
	sSql = sSql & iDocOrgID & ", '" & DBsafe( sDocumentPath ) & "', '" & DBsafe( sFileName ) & "', " & iParentFolderId & ", " & Session("UserID") & ", 0, " & sFileSize & " )"

	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean isValidFolder( sVirtualPath, sSubPath )
'--------------------------------------------------------------------------------------------------
Function isValidFolder( ByVal sVirtualPath, ByVal sSubPath )
	Dim sFolderPath

	isValidFolder = True 
	sFolderPath = Replace(sSubPath, (sVirtualPath & "/"), "" )
	'response.write "sFolderPath = " & sFolderPath & "<br />"

	If Left(sFolderPath, 19) <> "published_documents" Then
		If Left(sFolderPath, 21) <> "unpublished_documents" Then
			isValidFolder = False 
		End If
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean isValidFileName( sFileName )
'--------------------------------------------------------------------------------------------------
Function isValidFileName( ByVal sFileName )

	If LCase( sFileName ) = "thumbs.db" Then 
		isValidFileName = False 
	Else
		isValidFileName = True 
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' void cleanDocumentsDB( iDocOrgID, sPath, sVirtualPath )
'--------------------------------------------------------------------------------------------------
Sub cleanDocumentsTable( ByVal iDocOrgID, ByVal sPath, ByVal sVirtualPath )
	Dim sSql, oRs, sFilePath

	sSql = "SELECT documentid, documenturl FROM documents where orgid = " & iDocOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sFilePath = oRs("documenturl")
		sFilePath = Replace(sFilePath, sVirtualPath, "" )
		sFilePath = sPath & sFilePath
		sFilePath = Replace(sFilePath, "/", "\" )
		If Not FSO.FileExists( sFilePath ) Then 
			RemoveDocumentFromTable oRs("documentid")
		End If 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' void RemoveDocumentFromTable( iDocumentId )
'--------------------------------------------------------------------------------------------------
Sub RemoveDocumentFromTable( ByVal iDocumentId )
	Dim oCmd, sSql

	sSql = "DELETE FROM documents WHERE documentid = " & iDocumentId

	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' void cleanFoldersTable( iDocOrgID, sPath, sVirtualPath )
'--------------------------------------------------------------------------------------------------
Sub cleanFoldersTable( ByVal iDocOrgID, ByVal sPath, ByVal sVirtualPath )
	Dim sSql, oRs, sFolderPath

	sSql = "SELECT folderid, folderpath FROM documentfolders where orgid = " & iDocOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sFolderPath = oRs("folderpath")
		sFolderPath = Replace(sFolderPath, sVirtualPath, "" )
		sFolderPath = sPath & sFolderPath
		sFolderPath = Replace(sFolderPath, "/", "\" )
		If Not FSO.FolderExists( sFolderPath ) Then 
			'response.write "<b>Not Found: " & sFolderPath & "</b><br />"
			RemoveFolderFromTable oRs("folderid")
'		Else
'			response.write "Found: " & sFolderPath & "<br />"
		End If 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void RemoveFolderFromTable iFolderId 
'--------------------------------------------------------------------------------------------------
Sub RemoveFolderFromTable( ByVal iFolderId )
	Dim oCmd, sSql

	sSql = "DELETE FROM documentfolders WHERE folderid = " & iFolderId

	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' string DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString

End Function


'-------------------------------------------------------------------------------------------------
' void RunSQLStatement sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQLStatement( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub
%>
