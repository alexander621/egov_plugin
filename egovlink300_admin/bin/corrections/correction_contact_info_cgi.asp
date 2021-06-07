<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_CONTACT_INFO_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/12/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE CONTACT INFORMATION
'
' MODIFICATION HISTORY
' 1.0	02/12/07	JOHN STULLENBERGER - INITIAL VERSION
'
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' CALL ROUTINE TO SAVE CONTACT INFORMATION
Call subSaveContactInfo()


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' SUB SUBSAVECONTACTINFO()
'------------------------------------------------------------------------------------------------------------
Sub subSaveContactInfo()

	' GET INFORMATION FROM DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	sSQL = "Select userfname,userlname,userbusinessname,useremail,userhomephone,userfax,useraddress,usercity,userstate,userzip,contactmethodid From  egov_actionline_requests INNER JOIN egov_users ON  egov_actionline_requests.userid = egov_users.userid where egov_actionline_requests.action_autoid= " & request("irequestid")
	oSave.Open sSQL, Application("DSN"), 1, 2

	' UPDATE INFORMATION
	If NOT oSave.EOF Then

		sLogEntry = ""
		
		' CHECK FOR DIFFERENT VALUES AND SAVE
		For Each oColumn in oSave.Fields
			
			' COMPARE VALUES - IF CHANGED UPDATE AND LOG
			If trim(oColumn.Value) <> Trim(request(oColumn.Name)) Then

				' LOG CHANGES
				If oColumn.Name <> "contactmethodid" Then
					' ALL CHANGES OTHER THAN CONTACT ID
					If sLogEntry = "" Then
						sLogEntry = "Edit Contact Information: " & sLogEntry & chr(34) & trim(oColumn.Value)  & chr(34) &  " changed to " & chr(34) &  Trim(request(oColumn.Name)) & chr(34) 
					Else
						sLogEntry = sLogEntry & "," & chr(34) &  trim(oColumn.Value)  & chr(34) &  " changed to " & chr(34) &  Trim(request(oColumn.Name)) & chr(34) 
					End If
				Else
					' SPECIAL HANDLING TO GET DISPLAY NAME FOR CONTACT METHOD
					If sLogEntry = "" Then
						sLogEntry = "Edit Contact Information: " & sLogEntry & chr(34) & DisplayContactMethod(trim(oColumn.Value))  & chr(34) &  " changed to " & chr(34) &  DisplayContactMethod(Trim(request(oColumn.Name))) & chr(34) 
					Else
						sLogEntry = sLogEntry & "," & chr(34) &   DisplayContactMethod(trim(oColumn.Value))   & chr(34) &  " changed to " & chr(34) &  DisplayContactMethod(Trim(request(oColumn.Name))) & chr(34) 
					End If


				End If

				' SAVE
				oSave(oColumn.Name) = Trim(request(oColumn.Name))

			End If

		Next

		
		' SAVE CHANGES
		oSave.Update
	
	End If

	' CLOSE RECORDSET
	oSave.Close
	Set oSave = Nothing

	' RECORD IN LOG THE SAVE ACTIVITY
	Call AddCommentTaskComment(sLogEntry,sExternalMsg,request("status"),request("irequestid"),session("userid"),session("orgid"))


	' RETURN REQUEST PAGE
	response.redirect("correction_contact_info.asp?irequestid=" & request("irequestid") & "&r=save&status="&request("status"))

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
%>
