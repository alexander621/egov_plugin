<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: docscommon.asp
' AUTHOR: Steve Loar
' CREATED: 08/31/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared functions for documents. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   08/31/2010   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' integer GetFileId( sFileURL )
'--------------------------------------------------------------------------------------------------
Function GetFileId( ByVal sFileURL )
	Dim sSql, oRs 

	sSql = "SELECT DocumentID FROM Documents WHERE DocumentURL = '" & sFileURL
	sSql = sSql & "' AND OrgID = " & session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFileId = oRs("DocumentID")
	Else
		GetFileId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFileName( iFileId )
'--------------------------------------------------------------------------------------------------
Function GetFileName( ByVal iFileId )
	Dim sSql, oRs 

	sSql = "SELECT DocumentTitle FROM Documents WHERE DocumentID = " & iFileId
	sSql = sSql & " AND OrgID = " & session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFileName = oRs("DocumentTitle")
	Else
		GetFileName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



'--------------------------------------------------------------------------------------------------
' integer GetFolderId( sFolder )
'--------------------------------------------------------------------------------------------------
Function GetFolderId( ByVal sFolder )
	Dim sSql, oRs 

	sSql = "SELECT FolderID FROM DocumentFolders WHERE FolderPath = '" & sFolder
	sSql = sSql & "' AND OrgID = " & session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFolderId = oRs("FolderID")
	Else
		GetFolderId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFolderName( iFolderId )
'--------------------------------------------------------------------------------------------------
Function GetFolderName( ByVal iFolderId )
	Dim sSql, oRs 

	sSql = "SELECT FolderName FROM DocumentFolders WHERE FolderID = " & iFolderId
	sSql = sSql & " AND OrgID = " & session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFolderName = oRs("FolderName")
	Else
		GetFolderName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetFolderPath( iFolderId )
'--------------------------------------------------------------------------------------------------
Function GetFolderPath( ByVal iFolderId )
	Dim sSql, oRs 

	sSql = "SELECT FolderPath FROM DocumentFolders WHERE FolderID = " & iFolderId
	sSql = sSql & " AND OrgID = " & session("orgid")
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFolderPath = oRs("FolderPath")
	Else
		GetFolderPath = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>