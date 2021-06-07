<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: editpublicaccessdo.asp
' AUTHOR: Steve Loar
' CREATED: 09/28/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Edits public access.
'
' MODIFICATION HISTORY
' 1.0   09/28/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFolder, iFolderId, sTask, iSelectID

sTargetFolder = request("path")
iFolderId = CLng(request("folderid"))
sTask = Request("t")

sSuccessFlag = "pc"

'response.write "sTask = " & sTask & "<br />"
If sTask = "add" Then

	Set oCnn = Server.CreateObject("ADODB.Connection")
	oCnn.Open Application("DSN")

	For Each iSelectID in Request.Form("RemainingList")
		If CLng(iSelectID) <> -1 Then
			sSql = "EXEC NewFolderAccessCitizen '" & sTargetFolder & "', " & iSelectID
			response.write sSql & "<br />"
			oCnn.Execute sSql
		End If
	Next

	oCnn.Close
	Set oCnn = Nothing

ElseIf sTask = "del" Then

	Set oCnn = Server.CreateObject("ADODB.Connection")
	oCnn.Open Application("DSN")

	For Each iSelectID in Request.Form("ExistingList")
		If CLng(iSelectID) <> -1 Then
			sSql = "EXEC DelFolderAccessCitizen '" & sTargetFolder & "', " & iSelectID
			response.write sSql & "<br />"
			oCnn.Execute sSql
		End If
	Next

	oCnn.Close
	Set oCnn = Nothing

End If

response.write "sSuccessFlag = " & sSuccessFlag & "<br /><br />"

' back to the edit public access page
response.redirect "editpublicaccess.asp?path=" & request("path") & "/&sf=" & sSuccessFlag


%>