<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitattachmentremove.asp
' AUTHOR: Steve Loar
' CREATED: 07/08/2008
' COPYRIGHT: COPYRIGHT 2008 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION: remove permit attachments. Called via AJAX from permitedit.asp
'
' MODIFICATION HISTORY
' 1.0	07/08/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitAttachmentId, sSql, oRs, bFound, oFSO, sAttachmentName, sAttachmentPath

iPermitAttachmentId = CLng(request("permitattachmentid"))

' GET FILE EXTENSION FROM THE ATTACHMENT TABLE
sSql = "SELECT attachmentname, attachmentpath FROM egov_permitattachments "
sSql = sSql & " WHERE orgid = " & session("orgid") & " AND permitattachmentid = " & iPermitAttachmentId

Set oRs =  Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL,Application("DSN"), 3, 1

If NOT oRs.EOF Then
	sAttachmentName = oRs("attachmentname")
	sAttachmentPath = oRs("attachmentpath")
	bFound = True 
Else 
	bFound = False 
End If

oRs.Close
Set oRs = Nothing

If bFound Then 
	' DELETE FROM THE ATTACHMENT TABLE
	sSql = "DELETE FROM egov_permitattachments WHERE permitattachmentid = " & iPermitAttachmentId
	RunSQL sSql

	' DELETE FROM SERVER FILESYSTEM
	sServerPath = server.mappath(sAttachmentPath)
	sServerPath = sServerPath & "\" & iPermitAttachmentId & "_" & sAttachmentName

	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(sServerPath)  Then
		'DELETE FILE
		oFSO.DeleteFile(sServerPath)
		response.write "SUCCESS"
	Else 
		response.write "Failed. File not found."
	End If

	Set oFSO = Nothing

	
Else
	response.write "Failed. Attachment file not found."
End If 

%>