<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: updateAttachmentSecurity.asp
' AUTHOR: David Boyer
' CREATED: 08/10/2009
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates the "isSecure" column for an attachment.
'
' MODIFICATION HISTORY
' 1.0  08/10/09 	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success   = "Y"
 sAttachmentID = 0
 sIsSecure     = 0

 if request("attachmentid") <> "" then
    sAttachmentID = request("attachmentid")
 end if

 if request("isSecure") = "on" then
    sIsSecure = 1
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

'Update the attachment
 sSQL = "UPDATE egov_submitted_request_attachments SET "
 sSQL = sSQL & " isSecure = " & sIsSecure
 sSQL = sSQL & " WHERE attachmentid = " & sAttachmentID

 set oSaveOpt = Server.CreateObject("ADODB.Recordset")
 oSaveOpt.Open sSQL, Application("DSN"), 3, 1

 set oSaveOpt = nothing 

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    response.write "Attachment Changes Saved"
 end if
%>