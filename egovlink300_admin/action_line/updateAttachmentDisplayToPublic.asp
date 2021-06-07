<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: updateAttachmentDisplayToPublic.asp
' AUTHOR: David Boyer
' CREATED: 06/29/2010
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates the "DisplayToPublic" column for an attachment.
'
' MODIFICATION HISTORY
' 1.0 06/29/2010	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 lcl_success      = "Y"
 sAttachmentID    = 0
 sDisplayToPublic = 0

 if request("attachmentid") <> "" then
    sAttachmentID = request("attachmentid")
 end if

 if request("DisplayToPublic") = "on" then
    sDisplayToPublic = 1
 end if

 if request("isAjaxRoutine") = "Y" then
    lcl_isAjaxRoutine = True
 else
    lcl_isAjaxRoutine = False
 end if

'Update the attachment
 sSQL1 = "UPDATE egov_submitted_request_attachments SET "
 sSQL1 = sSQL1 & " displayToPublic = " & sDisplayToPublic
 sSQL1 = sSQL1 & " WHERE attachmentid = " & sAttachmentID

 set oDisplayToPublic = Server.CreateObject("ADODB.Recordset")
 oDisplayToPublic.Open sSQL1, Application("DSN"), 3, 1

 set oDisplayToPublic = nothing 

 if lcl_success = "Y" AND lcl_isAjaxRoutine then
    response.write "Attachment Changes Saved"
 end if

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "insert into my_table_dtb(notes) values ('" & replace(p_value,"'","''") & "') "

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

  set oDTB = nothing

end sub
%>