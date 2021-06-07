<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: ATTACHMENT_VIEW.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/29/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  VIEW ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	01/29/07	JOHN STULLENBERGER - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'CALL SUB TO VIEW SELECTED ATTACHMENT
 Call subViewAttachment(request("attachmentid"))

'------------------------------------------------------------------------------
sub subViewAttachment(iattachmentid)

	'GET FILE EXTENSION FROM THE ATTACHMENT TABLE
 	sSQL = "Select attachment_name FROM egov_submitted_request_attachments WHERE attachmentid = '" & iattachmentid  & "'"
 	set oExt = Server.CreateObject("ADODB.Recordset")
 	oExt.Open sSQL,Application("DSN"),1,3

 	if not oExt.eof then
	  	'FILE FOUND IN SQL TABLE
   		sExt = LCASE(RIGHT(oExt("attachment_name"),LEN(oExt("attachment_name")) - instrrev(oExt("attachment_name"),".")))
   		sFileName = oExt("attachment_name")
  else
		  'ERROR FILE NOT FOUND
   		response.write "Attachment Not Found!"
   		response.end
 	end if

 	set oExt = nothing

	'FIND IN SERVER FILESYSTEM
	 'sServerPath = server.mappath("../") & "\custom\pub\" & session("virtualdirectory") & "\attachments"
  sServerPath = server.mappath("\public_documents300\") & "\" & session("virtualdirectory") & "\attachments"
	 sServerPath = sServerPath & "\" & iattachmentid & "." & sExt

 	set oFSO = Server.CreateObject("Scripting.FileSystemObject")

 	if oFSO.FileExists(sServerPath)  Then
  		'FILE FOUND IN SERVER FILESYSTEM
  else
  		'ERROR FILE NOT FOUND
   		response.write "Attachment Not Found!"
   		response.end
  end if

  set oFSO = nothing

	'OPEN DOCUMENT IN BINARY AND STREAM TO BROWSER
 	response.clear
 	Response.Charset = "UTF-8"

	 if LCase(sExt) = "jpg" or LCase(sExt) = "jpeg" then
   		response.contentType = "image/jpeg"
  else
   		response.contentType = "application/octet-stream"
  end if

 	response.addheader "content-disposition", "attachment; filename=" & sFileName
	 Response.Buffer = true
	
 	Const adTypeBinary = 1
 	Set oStream = Server.CreateObject("ADODB.Stream")
 	oStream.Open
 	oStream.Type = adTypeBinary
 	oStream.LoadFromFile sServerPath
 	Response.BinaryWrite oStream.Read
 	Response.Flush

 	oStream.Close
 	set objStream = nothing

end sub
%>