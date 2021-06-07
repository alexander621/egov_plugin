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
' 1.0 02/13/07	John Stullenberger - Initial Version
' 1.1 09/06/07 David Boyer - Added Sub-Status Activity Log Tracking
' 1.2 08/04/10 David Boyer - Combined "admin" fields.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

 if request("ftype") <> "" then
    lcl_fieldtype = request("ftype")
 else
    lcl_fieldtype = "PUB"
 end if

'Determine which routine is needed to save the information
 select case request("formtype")

  	case "adminfields"

     'SAVE ADMIN FIELD INFORMATION
     'Performed the very first time internal fields is opened for updating.
     'This process takes the data from the "setup tables" and inserts the data on the "live tables"
     	subSaveFormAdminFields lcl_fieldtype

	  case "blob"

  		 'SAVE BLOB INFORMATION
		    subSaveFormInfoBlob lcl_fieldtype

	  case "nonblob"

   	 'SAVE NON-BLOB INFORMATION
		    subSaveFormInfoNonBlob lcl_fieldtype
 end select

'------------------------------------------------------------------------------
sub subSaveFormInfoBlob(iFieldType)

	sSQL = "SELECT comment "
 sSQL = sSQL & " FROM egov_actionline_requests "
 sSQL = sSQL & " WHERE action_autoid= " & request("irequestid")

 set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	oSave.Open sSQL, Application("DSN"), 1, 2

	if not oSave.eof then
  		sLogEntry = ""

 		'Create arrays to handle data
	  	dim arrQuestions(50)
		  dim arrAnswers(50)

 		'Get submitted questions and answers
  		for each oItem in request.form
     		'Questions
     			if left(oItem,8) = "question" then
       				iQuesCount = iQuesCount + 1
       				arrQuestions(iQuesCount) = request(oItem) 
     			end if

    			'Answers
     			if left(oItem,6) = "answer" then
       				iAnsCount = iAnsCount + 1
       				arrAnswers(iAnsCount) = request(oItem) 
     			end if
    next

 		'Build new blob
  		for iLoop = 1 to 50
		     	if arrQuestions(iLoop) <> "" then
    	   			sBlob = sBlob & "<p><b>" & arrQuestions(iLoop) & "</b><br>" & REPLACE(arrAnswers(iLoop),"default_novalue","") & "</p>"
     			end if
  		next

		'Compare Values: If changed then update and log
  	if ucase(oSave("comment")) <> ucase(sBlob) then
   			'sLogEntry = "Edit Request Form: " & REPLACE(oSave("comment"),"default_novalue","") & " changed to " & sBlob
      lcl_fieldtype_label = getFieldTypeLabel(iFieldType)
   			sLogEntry           = lcl_fieldtype_label & ": " & REPLACE(oSave("comment"),"default_novalue","") & " changed to " & sBlob

   			oSave("comment") = dbsafe(sBlob)
   			oSave.Update
 		end if

 		oSave.Close

 end if

	set oSave = nothing

'If there are change, record them
	if sLogEntry <> "" then
 		'Record in Log the save activity
  		AddCommentTaskComment sLogEntry, sExternalMsg, request("status"), request("irequestid"), session("userid"), _
                          session("orgid"), request("substatus"), "", ""
 end if

	response.redirect "../action_respond.asp?control=" & request("irequestid") & "&r=save&status="&request("status")

end sub

'------------------------------------------------------------------------------
sub subSaveFormInfoNonBlob(iFieldType)

 	sSQL = "SELECT * FROM egov_submitted_request_fields WHERE submitted_request_id= " & request("irequestid")

	 set oSave = Server.CreateObject("ADODB.Recordset")
	 oSave.CursorLocation = 3
	 oSave.Open sSQL, Application("DSN"), 1, 2

	'Update information if request is found
 	if not oSave.eof then
   		sLogEntry = ""
   		sBlob     = ""
		
   		for each oItem in request.form
       'Find answers
     			if Left(oItem,9) = "frmanswer" then
       				ifield = Replace(oItem,"frmanswer","")

      				'Compare Values
           lcl_value = request(oItem)
           lcl_value = replace(lcl_value,"&quot;","""")

       				sLogEntry = sLogEntry & fnCompareFieldValues(lcl_value,ifield,fnGetFieldPrompt(iField))
				
        			DeleteCurrentAnswers(ifield)

      				'Save new values
       				for iResponse = 1 to request.form(oItem).count
          					SaveNewAnswers iField, _
                              request(oItem)(iResponse), _
                              request("submitted_request_form_field_name"&iField), _
                              request("submitted_request_pushfieldid"&iField)
           next

     				 'Build new blob for display
           'sBlob = sBlob & "<p><b>" & fnGetFieldPrompt(iField) & "</b><br>" & request(oItem) & "</p>" & vbcrlf & vbcrlf
        			sBlob = sBlob & "<p><b>" & fnGetFieldPrompt(iField) & "</b><br>" & request(oItem) & "</p>" & vbcrlf
     			end if
     next

  		'Update blob for display
     if iFieldType <> "INT" then
   		   SaveNewBlob sBlob, request("irequestid")
     end if

   	'Record in the log, save activity
   		if sLogEntry <> "" then
  	   		'sLogEntry = "Edit Request Form: "  & sLogEntry
        lcl_fieldtype_label = getFieldTypeLabel(iFieldType)
  	   		sLogEntry           = lcl_fieldtype_label & ": "  & sLogEntry

   			  AddCommentTaskComment sLogEntry, sExternalMsg, request("status"), request("irequestid"), session("userid"), _
                              session("orgid"),request("substatus"), "", ""
     end if
  end if

 	set oSave = nothing

  response.redirect "../action_respond.asp?control=" & request("irequestid") & "&r=save&status="&request("status")

end sub

'------------------------------------------------------------------------------
sub subSaveFormAdminFields(iFieldType)

	'REPLACES BLOB FUNCTIONALITY - STORES DATA IN PROMPT ANSWER FORMAT
	 InsertRequestFieldsandResponses request("irequestid")

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
sub SaveNewAnswers(iField, sValue, sFormName, sPushFieldID)

  if sValue <> "" then
     lcl_value = replace(sValue,"&quot;","""")
'     lcl_value = replace(lcl_value,chr(10),"")
'     lcl_value = replace(lcl_value,chr(13),"")
     lcl_value = dbsafe(lcl_value)
  else
     lcl_value = ""
  end if

 	set oSave = Server.CreateObject("ADODB.Recordset")

 	sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_submitted_request_field_responses "
  sSQL = sSQL & " WHERE submitted_request_field_id= " & iField
 	oSave.Open sSQL, Application("DSN"), 1, 2
	
 	oSave.AddNew
 	oSave("submitted_request_field_id")        = iField
 	oSave("submitted_request_field_response")  = lcl_value
  oSave("submitted_request_form_field_name") = DBSafe(sFormName)

  if sPushFieldID <> "" then
     oSave("submitted_request_pushfieldid") = sPushFieldID
  end if

 	oSave.Update
 	oSave.Close

 	set oSave = nothing

end sub

'------------------------------------------------------------------------------
sub InsertRequestFieldsandResponses(iRequestID)
	
	iFieldCount = 0

	' ENUMERATE FIELDS AND ENTERED RESPONSES
	For Each oField in Request.Form

	'Get only fields and their associated values
		'If Left(oField,10) = "fmquestion" Then
'dtb_debug(oField)
		if left(oField,9) = "frmanswer" then
  			iFieldCount = iFieldCount + 1

  		'Get Field Prompt
  			'sFieldPrompt = request.form("fmname" & replace(oField,"fmquestion",""))
		  	sFieldPrompt = request.form("fmname" & iFieldCount)
'dtb_debug("sFieldPrompt: [" & sFieldPrompt & "] - fieldcount: [" & iFieldCount & "]")
'dtb_debug("fieldtype: ["    & request.form("fieldtype"   & iFieldCount) & "]")
'dtb_debug("frmanswer: ["   & request.form(oField) & "] - answerslist: [" & request.form("answerslist" & iFieldCount) & "]")
'dtb_debug("pdfname: ["      & request.form("pdfname"     & iFieldCount) & "]")
'dtb_debug("sequence: ["     & request.form("sequence"    & iFieldCount) & "]")
'dtb_debug("pushfieldid: ["  & request.form("pushfieldid" & iFieldCount) & "]")

  			iFieldID     = InsertFieldPrompt(sFieldPrompt, iRequestID, request.form("fieldtype"   & iFieldCount), _
                                                                request.form("answerslist" & iFieldCount), _
                                                                request.form("isrequired"  & iFieldCount), _
                                                                request.form("pdfname"     & iFieldCount), _
                                                                request.form("sequence"    & iFieldCount), _
                                                                request.form("pushfieldid" & iFieldCount))

 			'Enumerate and get field responses
   		for iResponse = 1 to request.form(oField).count
'dtb_debug("submitted_request_field_id: [" & iFieldID & "]")
'dtb_debug("submitted_request_field_response: [" & request.form(oField)(iResponse) & "]")

      			InsertFieldResponse iFieldID, _
                             request.form(oField)(iResponse), _
                             request.form("pdfname" & iFieldCount), _
                             request.form("pushfieldid" & iFieldCount)
  			next
		end if

	next

end sub

'------------------------------------------------------------------------------
function InsertFieldPrompt(sPrompt, _
                           iRequestID, _
                           iFieldType, _
                           sAnswerList, _
                           blnIsRequired, _
                           sPDFName, _
                           iSequence, _
                           iPushFieldID)
	  
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
	  oAddFieldPrompt("submitted_request_field_pdf_name")   = sPDFName
	  oAddFieldPrompt("submitted_request_field_sequence")   = iSequence

   if iPushFieldID <> "" then
      oAddFieldPrompt("submitted_request_field_pushfieldid") = iPushFieldID
   end if

	  oAddFieldPrompt("submitted_request_id")                = iRequestID
	  oAddFieldPrompt("submitted_request_field_isinternal")  = True
	 
	 ' SAVE ADDED INFORMATION
	  oAddFieldPrompt.Update
	  
	  ' SET NEW ROW ID
	  iReturnValue = oAddFieldPrompt("submitted_request_field_id")
	  
	  ' CLOSE 
	  oAddFieldPrompt.Close

	  InsertFieldPrompt = iReturnValue


End Function

'------------------------------------------------------------------------------
function InsertFieldResponse(iFieldID, sResponse, sPDFName, sPushFieldID)

	  iReturnValue = 0 

	  sSQL = "SELECT * FROM egov_submitted_request_field_responses WHERE 1=2"
	  Set oAddFieldPrompt = Server.CreateObject("ADODB.Recordset")
	  oAddFieldPrompt.CursorLocation = 3
   oAddFieldPrompt.Open sSQL, Application("DSN") , 1,3
	  
	  oAddFieldPrompt.AddNew
	  oAddFieldPrompt("submitted_request_field_id")        = iFieldID
	  oAddFieldPrompt("submitted_request_field_response")  = sResponse
   oAddFieldPrompt("submitted_request_form_field_name") = sPDFName

   if sPushFieldID <> "" then
      oAddFieldPrompt("submitted_request_pushfieldid")     = sPushFieldID
   end if

	  oAddFieldPrompt.Update
	  oAddFieldPrompt.Close

	  InsertFieldResponse = iReturnValue


end function

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
	
	 sReturnValue = ""
	 sCurrentValue = ""

	  sSQL = "SELECT submitted_request_field_response FROM egov_submitted_request_field_responses WHERE submitted_request_field_id=" & ifieldid
	  Set oCompare = Server.CreateObject("ADODB.Recordset")
      oCompare.Open sSQL, Application("DSN") , 1,3

	  ' IF VALUES FOUND FOR FIELD ID GET AND COMPARE
	  If NOT oCompare.EOF Then

			Do While NOT oCompare.EOF
				
				sCurrentValue = sCurrentValue & oCompare("submitted_request_field_response") & ", "

				oCompare.MoveNext
			Loop

			' REMOVE TRAILING COMMA
			If Len(sCurrentValue) > 0 Then
				sCurrentValue = Mid(sCurrentValue,1,Len(sCurrentValue) - 2 )
			End If

			If trim(sValue) <> trim(sCurrentValue) Then
				sReturnValue = sReturnValue & sfieldname & " " & chr(34) & sCurrentValue & chr(34) & " changed to " & chr(34) & sValue & chr(34) & "<BR>"
			End If

			' CLOSE 
			oCompare.Close

	  End If

	  Set oCompare = Nothing
	  

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

end function

'------------------------------------------------------------------------------
function getFieldTypeLabel(p_fieldType)

 'Default label is for the "public" fields
  lcl_return    = ""
  lcl_fieldtype = p_fieldType

  if lcl_fieldtype <> "" then
     lcl_fieldtype = UCASE(lcl_fieldtype)
  else
     lcl_fieldtype = "PUB"
  end if

 'Determine if the field type is for the "internal only" fields.
  if lcl_fieldtype = "INT" then
     lcl_return = "Edit Administrative Fields"
  else
     lcl_return = "Edit Request Form"
  end if

  getFieldTypeLabel = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  lcl_value = p_value

  if p_value <> "" then
     lcl_value = replace(lcl_value,"'","''")
  end if

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & lcl_value & "') "

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN") , 1,3

  set oDTB = nothing

end sub
%>