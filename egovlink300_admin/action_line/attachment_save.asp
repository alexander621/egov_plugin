<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: ATTACHMENT_SAVE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE ATTACHMENT FOR REQUEST 
'
' MODIFICATION HISTORY
' 1.0 01/29/07	John Stullenbeger - Initial Version
' 1.1 08/10/09 David Boyer - Added "isSecure"
' 2.0	04/25/2013	Terry Foster - Using centralized upload sub
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'Save the attachment
 subSaveAttachment()

'------------------------------------------------------------------------------
sub subSaveAttachment()

 sServerPath         = "\public_documents300\" & session("virtualdirectory") & "\attachments"
'Make sure that Attachment folder exists.  If not then create it.
 Call subAttachmentFolderCheck(Server.MapPath(sServerPath))

Set formFields = Server.CreateObject("Scripting.Dictionary")
call UploadFile(formFields, sServerPath, true, false)

'Get Variables
 sDesc               = formFields("attachmentdesc")
 sFileName           = formFields("filename")
 sFileExt            = LCASE(RIGHT(sFileName,LEN(sFileName) - instrrev(sFileName,".")))
 iRequestId          = formFields("irequestid")
 iScreenType         = formFields("screentype")
 iAttachmentIsSecure = formFields("attachmentIsSecure")
 iAdminUserId        = session("userid")
 iOrgId              = session("orgid")

'------------------------------------------------------------------------------
'N = New ActionLine Request
'E = Existing ActionLine Request
 if iScreenType <> "" then
    lcl_screenType = UCASE(iScreenType)
 else
    lcl_screenType = "E"
 end if

 if UCASE(iAttachmentIsSecure) = "ON" then
    iIsSecure = 1
 else
    iIsSecure = 0
 end if


 sServerFileName = FnStoreAttatchmentInfo(iRequestId,sFileName,sDesc,iAdminUserId,iIsSecure)

 'rename uploaded file
set objFso = CreateObject("Scripting.FileSystemObject") 
objFso.MoveFile sServerPath & "\" & sFileName, sServerPath & "\" & sServerFileName & "." & sFileExt
set objFso = Nothing 


'Return to request page.
 response.redirect "action_respond.asp?control=" & iRequestId & "&success=ATTACHMENT_ADDED"

end sub

'------------------------------------------------------------------------------
function fnStoreAttatchmentInfo(irequestid,sAttachment_Name,sAttachment_Desc,iadminuserid,iIsSecure)

 lcl_submit_date = ConvertDateTimetoTimeZone()

 sSQL = "SELECT attachmentid, submitted_request_id, attachment_name, attachment_desc, adminuserid,date_added, isSecure "
 sSQL = sSQL & " FROM egov_submitted_request_attachments "
 sSQL = sSQL & " WHERE 1 = 2"

	set oAttachment = Server.CreateObject("ADODB.Recordset")
	oAttachment.CursorLocation = 3
	oAttachment.Open sSQL, Application("DSN"), 1, 2

'Add database row
	oAttachment.AddNew
	oAttachment("submitted_request_id") = irequestid
	oAttachment("attachment_name")      = sAttachment_Name
	oAttachment("attachment_desc")      = sAttachment_Desc
	oAttachment("adminuserid")          = iadminuserid
	'oAttachment("date_added")           = Now()
	oAttachment("date_added")           = lcl_submit_date
 oAttachment("isSecure")             = iIsSecure
	oAttachment.Update
	iReturnValue = oAttachment("attachmentid")

'Close recordset
	oAttachment.Close
	set oAttachment = nothing

	fnStoreAttatchmentInfo = iReturnValue

end function

'------------------------------------------------------------------------------
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'------------------------------------------------------------------------------
function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
end function

'------------------------------------------------------------------------------
sub subAttachmentFolderCheck(sFolderPath)

	set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	if oFSO.FolderExists(sFolderPath) <> True then
   'Create Attachments folder
  		set oFolder = oFSO.CreateFolder(sFolderPath)
		  set oFolder = nothing
 end if

	set oFSO = nothing

end sub
%>
