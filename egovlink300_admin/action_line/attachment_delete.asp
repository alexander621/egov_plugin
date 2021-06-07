<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: ATTACHMENT_DELETE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  DELETE ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	01/29/07	John Stullenberger - Initial Version
' 1.1 12/14/09 David Boyer - Fixed Timezone issue when deleting an attachment.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

 call subDeleteAttachment(request("attachmentid"),request("irequestid"))

'------------------------------------------------------------------------------
sub subDeleteAttachment(iattachmentid,irequestid)

	'Get the file extension from the attachment table
 	sSQL = "SELECT attachment_name,attachment_desc "
  sSQL = sSQL & " FROM egov_submitted_request_attachments "
  sSQL = sSQL & " WHERE attachmentid = '" & iattachmentid  & "'"

 	set oExt =  Server.CreateObject("ADODB.Recordset")
 	oExt.Open sSQL,Application("DSN"),1,3

 	if not oExt.eof then
   		sExt      = LCASE(RIGHT(oExt("attachment_name"),LEN(oExt("attachment_name")) - instrrev(oExt("attachment_name"),".")))
   		sFileName = oExt("attachment_name")
 	  	sFileDesc = oExt("attachment_desc")
 	end if
 	set oExt = nothing

	'Delete from the attachment table
 	sSQL = "DELETE FROM egov_submitted_request_attachments WHERE attachmentid = '" & iattachmentid  & "'"

  set oDelAttachment = Server.CreateObject("ADODB.Recordset")
  oDelAttachment.Open sSQL, Application("DSN"), 3, 1
 	'Set oCmd = Server.CreateObject("ADODB.Command")
 	'With oCmd
  '		.ActiveConnection = Application("DSN")
	 ' 	.CommandText = sSql
	 ' 	.Execute
	 'End With
	 'Set oCmd = Nothing


	'Delete from server filesystem
 	sServerPath = server.mappath("../") & "\custom\pub\" & session("virtualdirectory") & "\attachments"
 	sServerPath = sServerPath & "\" & iattachmentid & "." & sExt

 	set oFSO = Server.CreateObject("Scripting.FileSystemObject")

 'Delete the file
 	if oFSO.FileExists(sServerPath) then
	   	oFSO.DeleteFile(sServerPath)
 	end if

 	set oFSO = nothing

	'Record in the log - Delete Activity
 	sInternalMsg = sFileName  & " deleted. File Description: " & sFileDesc

	 'Call AddCommentTaskComment(sInternalMsg,sExternalMsg,request("status"),iRequestId,session("userid"),session("orgid"),request("substatusid"))
  AddCommentTaskComment sInternalMsg, sExternalMsg, request("status"), iRequestId, session("userid"), session("orgid"), request("substatusid"), "", ""

 	response.redirect "action_respond.asp?control=" & iRequestID

end sub

'--------------------------------------------------------------------------------------------------
' ADDCOMMENTTASKCOMMENT(SINTERNALMSG,SEXTERNALMSG)
'--------------------------------------------------------------------------------------------------
'Function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID)
'		sSQL = "INSERT egov_action_responses (action_status,action_internalcomment,action_externalcomment,action_userid,action_orgid,action_autoid) VALUES ('" & sStatus & "','" & DBsafe(sInternalMsg) & "','" & DBsafe(sExternalMsg) & "','" & iUserID & "','" & iOrgID & "','" &iFormID & "')"
'		Set oComment = Server.CreateObject("ADODB.Recordset")
'		oComment.Open sSQL, Application("DSN") , 3, 1
'		Set oComment = Nothing
'End Function


'----------------------------------------
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'----------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function
%>