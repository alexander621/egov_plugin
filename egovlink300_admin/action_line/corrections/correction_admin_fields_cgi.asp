<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../action_line_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: CORRECTION_REQUEST_FORM_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/13/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE REQUEST FORM INFORMATION
'
' MODIFICATION HISTORY
' 1.0	 02/13/07	 JOHN STULLENBERGER - INITIAL VERSION
' 2.0  01/17/08  David Boyer - Added code to only update activity log for fields that have changed.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' CALL ROUTINE TO SAVE FORM INFORMATION
Select Case request("formtype")

 	Case "adminfields"

    	 	'SAVE ADMIN FIELD INFORMATION
       'Performed the very first time internal fields is opened for updating.
       'This process takes the data from the "setup tables" and inserts the data on the "live tables"
      		Call subSaveFormAdminFields()

  Case "nonblob"

     	 'SAVE NON-BLOB INFORMATION
       'Performed every time after the first time the internal fields screen is opened.
       'The "live tables" are accessed during this maintenance.
    		  Call subSaveFormInfoNonBlob()

End Select

'------------------------------------------------------------------------------
Sub subSaveFormInfoNonBlob()

'GET INFORMATION FROM DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_submitted_request_fields "
 sSQL = sSQL & " WHERE submitted_request_id = " & request("irequestid")
	oSave.Open sSQL, Application("DSN"), 1, 2

'UPDATE INFORMATION IF REQUEST IS FOUND
	If NOT oSave.EOF Then

  		sLogEntry   = ""
  		sBlob       = ""
		  iFieldCount = 0

  	'GET VALUES FOR QUESTIONS
  		For Each oItem In Request.Form

    			'FIND ANSWERS
     			If Left(oItem,9) = "frmanswer" Then
       				ifield      = Replace(oItem,"frmanswer","")
       				iFieldCount = iFieldCount + 1

      				'COMPARE VALUES
       				sLogEntry   = sLogEntry & fnCompareFieldValues(request(oItem),ifield,fnGetFieldPrompt(iField))

      				'DELETE CURRENT ANSWER VALUES
         		DeleteCurrentAnswers(ifield)

      				'SAVE NEW VALUES
        			For iResponse = 1 to request.form(oItem).count
          					Call SaveNewAnswers(iField,request(oItem)(iResponse),request.form("pdfformname")(iFieldCount))
       				Next
     			End If
  		Next

 		'Record in log the save activity
  		if sLogEntry <> "" then
		    	sLogEntry = "Edit Administrative Fields: "  & sLogEntry
    			AddCommentTaskComment sLogEntry, sExternalMsg, request("status"), request("irequestid"), session("userid"), session("orgid"), request("substatus"), "", ""
    end if
	end if

	set oSave = nothing

	response.redirect "../action_respond.asp?control=" & request("irequestid") & "&r=save&status="&request("status")

end sub

'------------------------------------------------------------------------------
Sub subSaveFormAdminFields()

	'DELETE EXISTING DATA

	'REPLACES BLOB FUNCTIONALITY - STORES DATA IN PROMPT ANSWER FORMAT
	 Call InsertRequestFieldsandResponses(request("irequestid"))

  sLogEntry = ""

	'BUILD LOG ENTRY
	 For Each oItem In Request.Form

      'if left(oItem,10) = "answerlist" OR left(oItem,10) = "fmquestion" then
      if left(oItem,6) = "fmname" then
         lcl_question = request(oItem)
         lcl_id       = replace(oItem,"fmname","")
         sValue       = trim(REPLACE(REPLACE(request("fmquestion" & lcl_id),"default_novalue",""),chr(13),""))

         if sValue <> "" then
            'iQuesNum = replace(oItem,"fmname","")
            'sLogEntry = sLogEntry & request(oItem) & " " & chr(34) & sCurrentValue & chr(34) & " changed to " & chr(34) & request("fmquestion" & iQuesNum) & chr(34) & "<BR>"
            sLogEntry = sLogEntry & lcl_question & " " & chr(34)
            sLogEntry = sLogEntry & sCurrentValue & chr(34)
            sLogEntry = sLogEntry & " changed to " & chr(34) & sValue & chr(34) & "<br />"

         end if
      end if
  next

	'Update the Activity Log
  if sLogEntry <> "" then
    	sLogEntry = "Edit Administrative Fields: "  & sLogEntry
    	AddCommentTaskComment sLogEntry, sExternalMsg, request("status"), request("irequestid"), session("userid"), session("orgid"), request("substatus"), "", ""
  end if

	 response.redirect "../action_respond.asp?control=" & request("irequestid") & "&r=save&status="&request("status")

end sub

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
Sub DeleteCurrentAnswers(iField)
	
	' GET INFORMATION FROM DATABASE
	Set oDelete = Server.CreateObject("ADODB.Recordset")
	sSQL = "DELETE FROM egov_submitted_request_field_responses WHERE submitted_request_field_id= " & iField
	oDelete.Open sSQL, Application("DSN"), 1, 3
	Set oDelete = Nothing

End Sub

'------------------------------------------------------------------------------
Sub SaveNewAnswers(iField,sValue,sPDFName)
	
	' CONNECT TO DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_submitted_request_field_responses "
 sSQL = sSQL & " WHERE submitted_request_field_id = " & iField
	oSave.Open sSQL, Application("DSN"), 1, 2
	
	oSave.AddNew
	oSave("submitted_request_field_id") = iField
	oSave("submitted_request_field_response") = DBSAfe(sValue)
	oSave("submitted_request_form_field_name") = sPDFName
	oSave.Update
	oSave.Close

	Set oSave = Nothing

End Sub

'------------------------------------------------------------------------------
Sub InsertRequestFieldsandResponses(iRequestID)
	
	iFieldCount = 0

	' ENUMERATE FIELDS AND ENTERED RESPONSES
	For Each oField in Request.Form
		
		' GET ONLY FIELDS AND THEIR ASSOCIATED VALUES
		If Left(oField,10) = "fmquestion" Then
			
			iFieldCount = iFieldCount + 1

			' GET FIELD PROMPT
			sFieldPrompt = request.form("fmname" & replace(oField,"fmquestion",""))
			iFieldID =  InsertFieldPrompt(sFieldPrompt,iRequestID,request.form("fieldtype")(iFieldCount),request.form("answerlist" & iFieldCount),request.form("isrequired")(iFieldCount),request.form("sequence")(iFieldCount),request.form("pdfformname")(iFieldCount))

			' ENUMERATE AND GET FIELD RESPONSES
'     sSQL1 = "INSERT INTO my_table_dtb VALUES ('" & sFieldPrompt & " - " & iFieldID & "') "

'     Set rs = Server.CreateObject("ADODB.Recordset")
'     rs.Open sSQL1, Application("DSN") , 3, 1


			For iResponse = 1 to request.form(oField).count
   				Call InsertFieldResponse(request.form(oField)(iResponse),iFieldID,request.form("pdfformname")(iFieldCount))
			Next
			
		End If

	Next

End Sub

'------------------------------------------------------------------------------
Function InsertFieldPrompt(sPrompt,iRequestID,iFieldType,sAnswerList,blnIsRequired,iSequence,sPDFFormName)
	  
	  iReturnValue = 0 

	  sSQL = "SELECT * FROM egov_submitted_request_fields WHERE 1=2"
	  Set oAddFieldPrompt = Server.CreateObject("ADODB.Recordset")
	  oAddFieldPrompt.CursorLocation = 3
      oAddFieldPrompt.Open sSQL, Application("DSN") , 1,3
	  
	  ' ADD NEW ROW
	  oAddFieldPrompt.AddNew
	  
	  oAddFieldPrompt("submitted_request_field_prompt")     = sPrompt
	  oAddFieldPrompt("submitted_request_field_type_id")    = iFieldType
	  oAddFieldPrompt("submitted_request_field_answerlist") = sAnswerList
	  oAddFieldPrompt("submitted_request_field_isrequired") = blnIsRequired
	  oAddFieldPrompt("submitted_request_field_pdf_name")   = sPDFFormName
	  oAddFieldPrompt("submitted_request_field_sequence")   = iSequence
	  oAddFieldPrompt("submitted_request_id")               = iRequestID
	  oAddFieldPrompt("submitted_request_field_isinternal") = True
	 
	 ' SAVE ADDED INFORMATION
	  oAddFieldPrompt.Update
	  
	  ' SET NEW ROW ID
	  iReturnValue = oAddFieldPrompt("submitted_request_field_id")
	  
	  ' CLOSE 
	  oAddFieldPrompt.Close

	  InsertFieldPrompt = iReturnValue


End Function

'------------------------------------------------------------------------------
Function InsertFieldResponse(sResponse,iFieldID,sPDFName)
	  
	  iReturnValue = 0 

	  sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE 1=2"
	  Set oAddFieldPrompt = Server.CreateObject("ADODB.Recordset")
	  oAddFieldPrompt.CursorLocation = 3
      oAddFieldPrompt.Open sSQL, Application("DSN") , 1,3
	  
	  ' ADD NEW ROW
	  oAddFieldPrompt.AddNew
	  
	  oAddFieldPrompt("submitted_request_field_id") = iFieldID
	  oAddFieldPrompt("submitted_request_field_response") = sResponse
	  oAddFieldPrompt("submitted_request_form_field_name") = sPDFName
	 
	 ' SAVE ADDED INFORMATION
	  oAddFieldPrompt.Update
	  
	  
	  ' CLOSE 
	  oAddFieldPrompt.Close

	  InsertFieldResponse = iReturnValue


End Function

'------------------------------------------------------------------------------
Sub SaveNewBlob(sBlob,irequestid)
	
	' CONNECT TO DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT comment FROM egov_actionline_requests WHERE action_autoid= " & irequestid
	oSave.Open sSQL, Application("DSN"), 1, 2
	oSave("comment") = DBSAfe(sBlob)
	oSave.Update
	oSave.Close

	Set oSave = Nothing

End Sub

'------------------------------------------------------------------------------
Function fnCompareFieldValues(sValue,ifieldid,sfieldname)
	
	 sReturnValue  = ""
	 sCurrentValue = ""

	 sSQL = "SELECT submitted_request_field_response "
  sSQL = sSQL & " FROM egov_submitted_request_field_responses "
  sSQL = sSQL & " WHERE submitted_request_field_id = " & ifieldid

	 set oCompare = Server.CreateObject("ADODB.Recordset")
  oCompare.Open sSQL, Application("DSN") , 1,3

	'IF VALUES FOUND FOR FIELD ID GET AND COMPARE
	 If NOT oCompare.EOF Then
		  	do while NOT oCompare.EOF
				 			sCurrentValue = sCurrentValue & oCompare("submitted_request_field_response") & ", "
    				oCompare.MoveNext
   		loop

 			'REMOVE TRAILING COMMA
  			if Len(sCurrentValue) > 0 then
		    		sCurrentValue = Mid(sCurrentValue,1,Len(sCurrentValue) - 2 )
  			end if

    'Clean up the values so we can compare them.
     sValue        = trim(REPLACE(REPLACE(sValue,"default_novalue",""),chr(13),""))
   		sCurrentValue = trim(REPLACE(REPLACE(sCurrentValue,"default_novalue",""),chr(13),""))

			  if sValue <> sCurrentValue then
       'Only document the history if the valuehas changed and either the current and/or new values are not null.
        if sValue <> "" AND sCurrentValue <> "" _
        OR sValue <> "" AND sCurrentValue = "" _
        OR sValue = ""  AND sCurrentValue <> "" then
       				sReturnValue = sReturnValue & sfieldname & " " & chr(34) & sCurrentValue & chr(34) & " changed to " & chr(34) & sValue & chr(34) & "<BR>"
        end if
  			end if

 			'CLOSE
  			oCompare.Close
	 end if

  set oCompare = Nothing

  fnCompareFieldValues = sReturnValue

End Function

'------------------------------------------------------------------------------
Function fnGetFieldPrompt(iField)

	  sReturnValue = "UNKNOWN" 

	  sSQL = "SELECT submitted_request_field_prompt FROM  egov_submitted_request_fields WHERE submitted_request_field_id=" & iField
	  Set oFieldName = Server.CreateObject("ADODB.Recordset")
      oFieldName.Open sSQL, Application("DSN") , 1,3
	  
	  If NOT oFieldName.EOF Then
		sReturnValue = oFieldName("submitted_request_field_prompt") 
		oFieldName.Close
	  End If

	  Set oFieldName = Nothing
	
	  fnGetFieldPrompt = sReturnValue

End Function
%>