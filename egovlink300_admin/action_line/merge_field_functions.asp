<%
'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLDYNAMICFIELDS(SVALUE,IREQUESTID)
'------------------------------------------------------------------------------------------------------------
Function fnFillDynamicFields( sValue, irequestid )

	'SET DEFAULT RETURN VALUE
	 sReturnValue = sValue

	'Retrieve values for form fields associated with this request
	 sSQL = "SELECT submitted_request_field_id, submitted_request_field_type_id "
  sSQL = sSQL & " FROM egov_submitted_request_fields "
  sSQL = sSQL & " WHERE submitted_request_id = '" & irequestid & "' "
  sSQL = sSQL & " ORDER BY submitted_request_field_sequence "

	 set oDynamicFields = Server.CreateObject("ADODB.Recordset")
	 oDynamicFields.Open sSQL, Application("DSN"), 3, 1

  if not oDynamicFields.eof then
     lcl_value = ""
     while not oDynamicFields.eof

       'Cycle through and build answer list of all of the answers that have been selected/entered for each field on the form.
        sSQLr = "SELECT submitted_request_field_id, submitted_request_field_response, submitted_request_form_field_name "
        sSQLr = sSQLr & " FROM egov_submitted_request_field_responses "
        sSQLr = sSQLr & " WHERE submitted_request_field_id = " & oDynamicFields("submitted_request_field_id")
        sSQLr = sSQLr & " AND submitted_request_form_field_name IS NOT NULL "
        sSQLr = sSQLr & " AND submitted_request_form_field_name <> '' "
        sSQLr = sSQLr & " ORDER BY submitted_request_field_id "

      	 set rsr = Server.CreateObject("ADODB.Recordset")
      	 rsr.Open sSQLr, Application("DSN"), 3, 1

        if not rsr.eof then
          'If the field type is a radio (2), select (4), or checkbox (6) then we need to cycle through the answers to before converting
          'the place holder.
           if clng(oDynamicFields("submitted_request_field_type_id")) = 2 _
           OR clng(oDynamicFields("submitted_request_field_type_id")) = 4 _
           OR clng(oDynamicFields("submitted_request_field_type_id")) = 6 then

              lcl_value = ""
              while not rsr.eof
                'An answer of "default_novalue" is the hidden field that has been set up to fix the bug
                'with radio, select, and checkbox fields and the user not selecting an option in these lists.
                 if rsr("submitted_request_field_response") <> "default_novalue" then
                    if lcl_value <> "" then
                       lcl_value = lcl_value & ", " & cleanup_display_value(rsr("submitted_request_field_response"))
                    else
                       lcl_value = cleanup_display_value(rsr("submitted_request_field_response"))
                    end if
                 else
                    lcl_value = lcl_value
                 end if

                'Get the form name.  It should be the same for all of the answers.
                 lcl_submitted_request_form_field_name = rsr("submitted_request_form_field_name")

                 rsr.movenext
              wend

              if lcl_submitted_request_form_field_name <> "" then
              			sFieldPlaceHolder = "[*" & lcl_submitted_request_form_field_name & "*]"
                	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_value)
              end if
          '----------------------------------------------------------------------------------
           else  'Any other type of field can simply convert the place holder with the answer
          '----------------------------------------------------------------------------------
              'lcl_value = cleanup_display_value(rsr("submitted_request_field_response"))
              lcl_value = rsr("submitted_request_field_response")

             'Get the form name.  It should be the same for all of the answers.
              lcl_submitted_request_form_field_name = rsr("submitted_request_form_field_name")

              if lcl_submitted_request_form_field_name <> "" then
              			sFieldPlaceHolder = "[*" & lcl_submitted_request_form_field_name & "*]"
                	sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_value)
              end if
           end if
        end if

        oDynamicFields.movenext
     wend
  end if

 'RETURN VALUE
 	fnFillDynamicFields = sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLSTANDARDFIELDS(SVALUE,IUSERID)
'------------------------------------------------------------------------------------------------------------
Function fnFillStandardFields(sValue,iUserId)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

'CONNECT TO DATABASE RETRIEVE VALUES FOR FORM FIELDS ASSOCIATED WITH THIS REQUEST
	sSQL ="Select * From egov_users where userid='" & iUserId & "'"

	Set oFields = Server.CreateObject("ADODB.Recordset")
	oFields.Open sSQL, Application("DSN"), 3, 1

	If NOT oFields.EOF Then
  	'LOOP THRU FORM FIELDS
  		For Each Field In oFields.Fields

    			'REPLACE FORM FIELD PLACE HOLDER WITH ACTUAL DATA
     			sFieldPlaceHolder = "[*" & Field.Name & "*]"
     			if oFields(Field.Name) <> "" AND NOT IsNull(oFields(Field.Name)) then
       				sReturnValue = replace(sReturnValue,sFieldPlaceHolder,oFields(Field.Name))
     			end if
  		Next

	End If

'RETURN VALUE
	fnFillStandardFields = sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLISSUELOCATIONFIELDS(SVALUE,IREQUESTID)
'------------------------------------------------------------------------------------------------------------
Function fnFillIssueLocationFields(sValue,irequestid)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

'CONNECT TO DATABASE RETRIEVE VALUES FOR ISSUE LOCATION FIELDS ASSOCIATED WITH THIS REQUEST
'	sSQL ="SELECT streetnumber, streetaddress, city, state, zip, comments "
	sSQL = "SELECT streetnumber, "
 sSQL = sSQL & " dbo.fn_buildAddress('', "
 sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetprefix, ''), "
 sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetaddress, ''), "
 sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetsuffix, ''), "
 sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetdirection, '')) AS streetaddress, "
 sSQL = sSQL & " city, state, zip, comments, legaldescription,listedowner,parcelidnumber "
 sSQL = sSQL & " FROM egov_action_response_issue_location "
 sSQL = sSQL & " WHERE actionrequestresponseid = '" & irequestid & "'"

	Set oFields = Server.CreateObject("ADODB.Recordset")
	oFields.Open sSQL, Application("DSN"), 3, 1

	If NOT oFields.EOF Then
 		'LOOP THRU FORM FIELDS
   	For Each Field In oFields.Fields

			    'REPLACE FORM FIELD PLACE HOLDER WITH ACTUAL DATA
     			sFieldPlaceHolder = "[*" & Field.Name & "*]"
     			if oFields(Field.Name) <> "" AND NOT IsNull(oFields(Field.Name)) then
       				sReturnValue = replace(sReturnValue,sFieldPlaceHolder,oFields(Field.Name))
      		end if
   	Next
	End If

'RETURN VALUE
	fnFillIssueLocationFields = sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLMERGEFIELDS(SVALUE,IREQUESTID,IUSERID,SADDTEXT)
'------------------------------------------------------------------------------------------------------------
Function fnFillMergeFields(sValue,irequestid,iuserid,sAddText,itrackid)

	'Fill General Purpose Fields
	 sValue = fnFillGeneralFields( sValue )

	'FILL STANDARD FIELDS
 	sValue = fnFillStandardFields(sValue,iuserid)

	'FILL ISSUE LOCATION FIELDS
	 sValue = fnFillIssueLocationFields(sValue,itrackid)

	'FILL DYNAMIC FIELDS
 	sValue = fnFillDynamicFields(sValue,itrackid)

	'FILL ADDITIONAL COMMENT FIELD
	 sValue = fnFillAdditionalCommentField(sValue,sAddText)

	'FILL CODE SECTIONS FIELDS
  sValue = fnFillCodeSectionsField(sValue,itrackid)

 'FILL TRACKING NUMBER FIELD
  sValue = fnFillTrackingNumberField(sValue,itrackid)

	'CLEAR ANY REMAINING UNFILLED MERGE FIELDS
 	sValue = fnClearMergeFields(sValue)

	'RETURN VALUE
 	fnFillMergeFields = sValue

End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION fnFillGeneralFields(SVALUE)
'------------------------------------------------------------------------------------------------------------
Function fnFillGeneralFields( ByVal sValue )
	Dim sNewValue, sDate

	sDate = MonthName(Month(Date())) & " " & Day(Date()) & ", " & Year(Date())
	
	' Today's Date
	sNewValue = Replace(sValue,"[*TodaysDate*]", sDate )
	
	fnFillGeneralFields = sNewValue 
	
End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLADDITIONALCOMMENTFIELD(SVALUE)
'------------------------------------------------------------------------------------------------------------
Function fnFillAdditionalCommentField(sValue,sAddText)
	
	sValue = replace(sValue,"[*Admin_Additional_Comments*]",sAddText)
	
	' RETURN VALUE
	fnFillAdditionalCommentField = sValue 
	
End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNCLEARMERGEFIELDS( SSTRING )
'------------------------------------------------------------------------------------------------------------
Function fnClearMergeFields( sString )
	
	' DEFAULT RETURN VALUE
	sReturnValue = sString
	sPattern = "([\[\*].*[\*\]])"

	' CREATE REGULAR EXPRESSION TO FIND MATCHING TAG SYNTAX OF [* any merge field name *]
	Dim oMergeFields
	Set oMergeFields = New RegExp
	oMergeFields.Pattern = sPattern
	oMergeFields.Global = True
	oMergeFields.IgnoreCase = True
		
	' REPLACE MATCHING TAGS
	sReturnValue = oMergeFields.Replace(sString, " ")

	' DESTROY OBJECT
	Set oMergeFields = Nothing
	
	' RETURN VALUE
	fnClearMergeFields =  sReturnValue

End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION FNFILLCODESECTIONSFIELD(SVALUE,IUSERID)
'------------------------------------------------------------------------------------------------------------
Function fnFillCodeSectionsField(sValue,itrackid)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

 'sSql = "SELECT code_sections FROM egov_actionline_requests WHERE action_autoid = " & itrackid

 sSql = "SELECT scs.submitted_action_code_id, cs.code_name, cs.description "
	sSql = sSql & " FROM egov_submitted_request_code_sections scs, egov_actionline_code_sections cs "
	sSql = sSql & " WHERE scs.submitted_action_code_id = cs.action_code_id "
	sSql = sSql & " AND scs.submitted_request_id = " & itrackid
	sSql = sSql & " ORDER BY upper(cs.code_name), upper(cs.description) "

	Set oGetCodes = Server.CreateObject("ADODB.Recordset")
	oGetCodes.Open sSql, Application("DSN"), 3, 1

	If NOT oGetCodes.EOF Then

 		'LOOP THRU CODE SECTIONS
    lcl_display_codes = ""
	  	while not oGetCodes.eof
		    'REPLACE FORM FIELD PLACE HOLDER WITH ACTUAL DATA
       if lcl_display_codes = "" then
          lcl_display_codes = "<b>" & oGetCodes("code_name") & "</b><br />&nbsp;&nbsp;&nbsp;" & oGetCodes("description") & "<br /><br />"
       else
          lcl_display_codes = lcl_display_codes & "<b>" & oGetCodes("code_name") & "</b><br />&nbsp;&nbsp;&nbsp;" & oGetCodes("description") & "<br /><br />"
       end if
       oGetCodes.movenext
		  wend
 else
	   lcl_display_codes = ""
	end if

'RETURN VALUE
 sFieldPlaceHolder = "[*Code_Sections*]"
 sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_display_codes)

 fnFillCodeSectionsField = sReturnValue

end function

'------------------------------------------------------------------------------------------------------------
function fnFillTrackingNumberField(sValue,itrackid)

'SET DEFAULT RETURN VALUE
	sReturnValue = sValue

 sSQL = "SELECT [tracking number] "
	sSQL = sSQL & " FROM egov_rpt_actionline "
	sSQL = sSQL & " WHERE action_autoid = " & itrackid

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSQL, Application("DSN"), 3, 1

	if NOT rs.eof then
    lcl_tracking_number = rs("tracking number")
 else
	   lcl_tracking_number = ""
	end if

'RETURN VALUE
 sFieldPlaceHolder = "[*Tracking Number*]"
 sReturnValue      = replace(sReturnValue,sFieldPlaceHolder,lcl_tracking_number)

 fnFillTrackingNumberField = sReturnValue

end function

'----------------------------------------------------------------------
function cleanup_display_value(p_value)
  if p_value <> "" then
     lcl_value = p_value
     lcl_value = REPLACE(lcl_value,vbcrlf,"")
     lcl_value = REPLACE(lcl_value,chr(10),"")
     lcl_value = REPLACE(lcl_value,chr(13),"")

     cleanup_display_value = lcl_value
  end if

end function
%>
