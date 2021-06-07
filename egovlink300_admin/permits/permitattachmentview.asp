<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitattachmentview.asp
' AUTHOR: Steve Loar
' CREATED: 07/08/2008
' COPYRIGHT: COPYRIGHT 2008 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  VIEW PERMIT ATTACHMENTS
'
' MODIFICATION HISTORY
' 1.0	07/08/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oStream, sSql, iPermitAttachmentId, oRs, bFound, oFSO, sExt

iPermitAttachmentId = CLng(request("permitattachmentid"))

' GET FILE EXTENSION FROM THE ATTACHMENT TABLE
sSql = "SELECT attachmentname, fileextension FROM egov_permitattachments WHERE orgid = " & session("orgid") & " AND permitattachmentid = " & iPermitAttachmentId

Set oRs =  Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL,Application("DSN"), 3, 1

If NOT oRs.EOF Then
	' FILE FOUND IN SQL TABLE
	sExt = oRs("fileextension")
	sFileName = oRs("attachmentname")
	bFound = True 
Else
	' ERROR FILE NOT FOUND
	response.write "<p>Attachment Not Found!</p>"
	bFound = false
End If
oRs.Close	
Set oRs = Nothing

If bFound Then 
	' FIND IN SERVER FILESYSTEM
	sServerPath = server.mappath("..\permitattachments") 
	sServerPath = sServerPath & "\" & iPermitAttachmentId & "." & sExt

	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If Not oFSO.FileExists(sServerPath)  Then
		' ERROR FILE NOT FOUND
		response.write "<p>Attachment Not Found!</p>"
		response.end
	End If

	Set oFSO = Nothing

	' OPEN DOCUMENT IN BINARY AND STREAM TO BROWSER
	response.clear
	Response.Charset = "UTF-8"
	Select Case LCase(sExt)
		Case "jpg"
			response.contentType = "image/jpeg"
		Case "jpeg"
			response.contentType = "image/jpeg"
		Case "gif"
			response.contentType = "image/gif"
		Case "pdf"
			response.contentType = "application/pdf"
		Case "xls"
			response.contentType = "application/x-excel"
		Case Else 
			response.contentType = "application/octet-stream"
	End Select 

	response.addheader "content-disposition", "attachment; filename=" & sFileName


	Const adTypeBinary = 1
	Set oStream = Server.CreateObject("ADODB.Stream")
	oStream.Open
	oStream.Type = adTypeBinary
	oStream.LoadFromFile sServerPath
	Response.BinaryWrite oStream.Read
	Response.Flush

	oStream.Close
	Set objStream = Nothing

End If 	

%>
