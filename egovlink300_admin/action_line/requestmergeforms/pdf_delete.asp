<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: PDF_DELETE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/28/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  DELETE ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	02/28/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------




' CALL SUB TO DELETE SELECTED ATTACHMENT
Call subDeletePDF(request("ipdfid"))




'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBDELETEPDF(IPDFID)
'--------------------------------------------------------------------------------------------------
Sub subDeletePDF(ipdfid)
	

	' DELETE FROM THE PDF TABLE
	sSQL = "DELETE FROM egov_action_request_pdfforms WHERE pdfid = '" & ipdfid  & "'"
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing


	' DELETE FROM SERVER FILESYSTEM
	sServerPath = server.mappath("../../") & "\custom\pub\" & session("virtualdirectory") & "\pdf_forms"
	sServerPath = sServerPath & "\" & ipdfid & ".pdf" 
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(sServerPath)  Then
		'DELETE FILE
		oFSO.DeleteFile(sServerPath)
	End If
	Set oFSO = Nothing


	' RECORD IN LOG DELETE ACTIVITY
	'sInternalMsg = sFileName  & " deleted. File Description: " & sFileDesc
	'Call AddCommentTaskComment(sInternalMsg,sExternalMsg,request("status"),iRequestId,session("userid"),session("orgid"))


	' RETURN REQUEST PAGE
	response.redirect("requestmergeforms_manage.asp")

End Sub


'----------------------------------------------------------------------------------------------------------------------
' ADDCOMMENTTASKCOMMENT(SINTERNALMSG,SEXTERNALMSG)
'----------------------------------------------------------------------------------------------------------------------
Function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID)
		sSQL = "INSERT egov_action_responses (action_status,action_internalcomment,action_externalcomment,action_userid,action_orgid,action_autoid) VALUES ('" & sStatus & "','" & DBsafe(sInternalMsg) & "','" & DBsafe(sExternalMsg) & "','" & iUserID & "','" & iOrgID & "','" &iFormID & "')"
		Set oComment = Server.CreateObject("ADODB.Recordset")
		oComment.Open sSQL, Application("DSN") , 3, 1
		Set oComment = Nothing
End Function
%>