<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_REQUEST_FORM_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/13/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE CONTACT INFORMATION
'
' MODIFICATION HISTORY
' 1.0	02/13/07	JOHN STULLENBERGER - INITIAL VERSION
'
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' CALL ROUTINE TO SAVE CONTACT INFORMATION
If request("formtype") = "blob" Then
	
	' SAVE Blob INFORMATION
	Call subSaveFormInfoBlob()

Else

	' SAVE NON-BLOB INFORMATION
	Call subSaveFormInfoNonBlob()

End If


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' SUB SUBSAVEFORMINFOBLOB() 
'------------------------------------------------------------------------------------------------------------
Sub subSaveFormInfoBlob()

	' GET INFORMATION FROM DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	sSQL = "SELECT comment FROM  egov_actionline_requests WHERE action_autoid= " & request("irequestid")
	oSave.Open sSQL, Application("DSN"), 1, 2

	' UPDATE INFORMATION
	If NOT oSave.EOF Then

		sLogEntry = ""
		

		' CREATE ARRAYS TO HANDLE DATA
		Dim arrQuestions(50) 
		Dim arrAnswers(50)


		' GET SUBMITTED QUESTIONS AND ANSWERS
		For Each oItem In Request.Form
			' QUESTIONS
			If Left(oItem,8) = "question" Then
				iQuesCount = iQuesCount + 1
				arrQuestions(iQuesCount) = request(oItem) 
			End If

			' ANSWERS
			If Left(oItem,6) = "answer" Then
				iAnsCount = iAnsCount + 1
				arrAnswers(iAnsCount) = request(oItem) 
			End If
		Next

		' BUILD NEW BLOB
		For iLoop = 1 To 50
			If arrQuestions(iLoop) <> "" Then
				sBlob = sBlob & "<p><b>" & arrQuestions(iLoop) & "</b><br>" & arrAnswers(iLoop) & "</p>"
			End If
		Next
		

		' COMPARE VALUES - IF CHANGED UPDATE AND LOG
		If UCase(oSave("comment")) <> UCase(sBlob) Then
			' SAVE CHANGES
			sLogEntry = "Edit Form Information: " & oSave("comment") & " changed to " & sBlob
			oSave("comment") = dbsafe(sBlob)
			oSave.Update
		End If

		' CLOSE RECORDSET
		oSave.Close
	
	End If


	Set oSave = Nothing

	' RECORD IN LOG THE SAVE ACTIVITY
	Call AddCommentTaskComment(sLogEntry,sExternalMsg,request("status"),request("irequestid"),session("userid"),session("orgid"))


	' RETURN REQUEST PAGE
	response.redirect("correction_request_form.asp?irequestid=" & request("irequestid") & "&r=save&status="&request("status"))

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SUBSAVEFORMINFONONBLOB()
'------------------------------------------------------------------------------------------------------------
Sub subSaveFormInfoNonBlob()


	' GET INFORMATION FROM DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	sSQL = "SELECT * FROM egov_submitted_request_fields WHERE submitted_request_id= " & request("irequestid")
	oSave.Open sSQL, Application("DSN"), 1, 2

	' UPDATE INFORMATION
	If NOT oSave.EOF Then

		sLogEntry = ""
		
		' GET VALUES FOR QUESTIONS
		For Each oItem In Request.Form
			 
			' FIND ANSWERS
			If Left(oItem,9) = "frmanswer" Then
				ifield = Replace(oItem,"frmanswer","")
				
				response.write ifield

				' COMPARE VALUES

				' DELETE CURRENT ANSWER VALUES
				DeleteCurrentAnswers(ifield)

				' SAVE NEW VALUES
				For iResponse = 1 to request.form(oItem).count
					Call SaveNewAnswers(iField,request(oItem)(iResponse))
				Next

			End If
			
		Next


	End If


	Set oSave = Nothing

	' RECORD IN LOG THE SAVE ACTIVITY
	sLogEntry = "Edit Form Information: This form was editted."
	Call AddCommentTaskComment(sLogEntry,sExternalMsg,request("status"),request("irequestid"),session("userid"),session("orgid"))


	' RETURN REQUEST PAGE
	response.redirect("correction_request_form.asp?irequestid=" & request("irequestid") & "&r=save&status="&request("status"))

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


'----------------------------------------
'  FUNCTION DBSAFE( STRDB )
'----------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION DISPLAYCONTACTMETHOD(ISELECTED)
'------------------------------------------------------------------------------------------------------------
Function DisplayContactMethod(iValue)

	sSQL = "SELECT * FROM egov_contactmethods WHERE rowid='" & iValue & "'"

	Set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oMethods.EOF Then
		iReturnValue = oMethods("contactdescription") 
	Else
		iReturnValue = "NOT SPECIFIED"
	End If


	Set oMethods = Nothing
	
	DisplayContactMEthod = iReturnValue
	

End Function


'------------------------------------------------------------------------------------------------------------
' SUB DELETECURRENTANSWERS(IFIELD)
'------------------------------------------------------------------------------------------------------------
Sub DeleteCurrentAnswers(iField)
	
	' GET INFORMATION FROM DATABASE
	Set oDelete = Server.CreateObject("ADODB.Recordset")
	sSQL = "DELETE FROM egov_submitted_request_field_responses WHERE submitted_request_field_id= " & iField
	oDelete.Open sSQL, Application("DSN"), 1, 3
	Set oDelete = Nothing

End Sub


'------------------------------------------------------------------------------------------------------------
' SUB SAVENEWANSWERS(IFIELD,SVALUE)
'------------------------------------------------------------------------------------------------------------
Sub SaveNewAnswers(iField,sValue)
	
	' CONNECT TO DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE submitted_request_field_id= " & iField
	oSave.Open sSQL, Application("DSN"), 1, 2
	
	oSave.AddNew
	oSave("submitted_request_field_id") = iField
	oSave("submitted_request_field_response") = DBSAfe(sValue)
	oSave.Update
	oSave.Close

	Set oSave = Nothing

End Sub
%>
