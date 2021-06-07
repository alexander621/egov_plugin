<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="merge_field_functions.asp" //-->
<%
 response.buffer = True 

'Get variables
 'lcl_system_id  = request("sys")
 lcl_request_id            = ""
 lcl_form_id               = ""
 lcl_documentid            = ""
 lcl_filename              = ""
 lcl_comment_title         = ""
 lcl_pdf_error_msg         = ""
 lcl_status                = ""
 lcl_substatus             = ""
 lcl_public_actionline_pdf = ""
 lcl_pdf_action            = ""
 lcl_pdf_action_text       = ""
 lcl_fillRequestData       = "N"
 lcl_mapServerPath         = "N"
 lcl_openExternalPDFviaURL = "N"

 if request("hideActLog") <> "" then
    lcl_hide_activitylog = UCASE(request("hideActLog"))
 else
    lcl_hide_activitylog = "N"
 end if

 if dbready_number(request("iRequestID")) then
    lcl_request_id = request("iRequestID")

   'Get the request information
    sSQLp = "SELECT r.status, r.sub_status_id, "
    sSQLp = sSQLp & " isnull(r.public_actionline_pdf,f.public_actionline_pdf) as public_actionline_pdf "
    sSQLp = sSQLp & " FROM egov_actionline_requests r, egov_action_request_forms f "
    sSQLp = sSQLp & " WHERE r.orgid = " & session("orgid")
    sSQLp = sSQLp & " AND r.category_id = f.action_form_id "
    sSQLp = sSQLp & " AND r.action_autoid = " & lcl_request_id

    set oPath = Server.CreateObject("ADODB.Recordset")
    oPath.Open sSQLp, Application("DSN"), 3, 1

    if not oPath.eof then
       lcl_filename      = oPath("public_actionline_pdf")
       lcl_comment_title = lcl_filename
       lcl_status        = oPath("status")
       lcl_substatus     = oPath("sub_status_id")
    end if

    oPath.close
    set oPath = nothing

 end if

 if dbready_number(request("docID")) then
    lcl_documentid = request("docID")
 end if

'Determine which PDF the user is attempting to open
 lcl_pdf_action = UCASE(request("pdfaction"))

'Begin building the path
 if lcl_documentid <> "" then
    sSQL = "SELECT documenturl, documenttitle "
    sSQL = sSQL & " FROM documents "
    sSQL = sSQL & " WHERE orgid = " & session("orgid")
    sSQL = sSQL & " AND documentid = " & lcl_documentid

    set oPDF = Server.CreateObject("ADODB.Recordset")
    oPDF.Open sSQL, Application("DSN"), 3, 1

    if not oPDF.eof then
       'lcl_filename        = ".." & REPLACE(UCASE(oPDF("documenturl")),"/PUBLIC_DOCUMENTS300","")
       'lcl_filename              = ".." & replace(ucase(oPDF("documenturl")),"/CUSTOM/PUB","")
       lcl_filename              = replace(ucase(oPDF("documenturl")),"/CUSTOM/PUB","")
       lcl_mapServerPath         = "Y"
       lcl_fillRequestData       = "Y"
       lcl_openExternalPDFviaURL = "Y"
    end if

    oPDF.close
    set oPDF = nothing

    'lcl_filename = Application("CommunityLink_DocUrl") & replace(lcl_filename,"..","")

 else
    if lcl_pdf_action = "WORKORDER" then
       'lcl_filename = "http://secure.eclink.com/egovlink/work_order_pdf_old.asp"
	   lcl_filename = "work_order_pdf_old.asp"
       lcl_filename = lcl_filename & "?sys="        & Application("INSTANCE")
       lcl_filename = lcl_filename & "&irequestID=" & lcl_request_id
       lcl_filename = lcl_filename & "&hideActLog=" & lcl_hide_activitylog

       lcl_comment_title         = "The ""Work Order"" has been viewed/printed."
       lcl_openExternalPDFviaURL = "Y"

    elseif lcl_pdf_action = "WORKORDER_CONDENSED" then
       'lcl_filename = "http://secure.eclink.com/egovlink/work_order_pdf.asp"
	   lcl_filename = "work_order_pdf.asp"
       lcl_filename = lcl_filename & "?sys="        & Application("INSTANCE")
       lcl_filename = lcl_filename & "&irequestID=" & lcl_request_id
       lcl_filename = lcl_filename & "&iOrgID="     & session("orgid")
       lcl_filename = lcl_filename & "&hideActLog=" & lcl_hide_activitylog

       lcl_comment_title         = "The ""Work Order (Condensed)"" has been viewed/printed."
       lcl_openExternalPDFviaURL = "Y"

    else
       'sPath = "/custom/pub/" & session("virtualdirectory") & "/unpublished_documents"
       sPath = "/public_documents300/" & session("virtualdirectory") & "/unpublished_documents"

      'Determine if there is a PDF associated to the form on the request
      'OR if an actual PDF filename has been passed in.
       if request("pdf") <> "" then
          lcl_filename = request("pdf")
       else
          lcl_filename      = lcl_public_actionline_pdf
          lcl_comment_title = lcl_filename
       end if

       if lcl_filename <> "" then
'          if left(lcl_filename,1) <> "/" AND right(sPath,1) <> "/" then
'             lcl_filename = "/" & lcl_filename
'          end if

         'Combine the path and filename
          'lcl_filename = ".." & sPath & lcl_filename
          lcl_filename = sPath & lcl_filename

       end if

      'Set the variable so that the filename is "mapped" on the server properly
       lcl_mapServerPath = "Y"

      'Set the variable so that the Action Line Request data is filled-in
       lcl_fillRequestData = "Y"
    end if
 end if

'Validate the file extension, Open the PDF, or show an error message.
 if lcl_filename <> "" then
   'Check the extension to ensure it is a PDF file.
    if lcl_openExternalPDFviaURL <> "Y" then
       if right(ucase(lcl_filename),4) <> ".PDF" then
          lcl_pdf_error_msg = "This is not a valid PDF file."
       end if
    end if

    if lcl_pdf_error_msg = "" then

      'Determine if we need to map the server path
       if lcl_mapServerPath = "Y" then
          sPath = server.mappath(lcl_filename)
          'sPath = server.mappath("../custom/pub/eclink/unpublished_documents/PDFs/mypdf.pdf")
       end if

      'Retrieve the Action Line Request data
       if lcl_fillRequestData = "Y" then
          fillForm sPath, lcl_request_id
       end if

      'Determine if we are to open an URL to a PDF
       if lcl_openExternalPDFviaURL = "Y" then
          response.redirect lcl_filename
       end if
    end if
 else
    lcl_pdf_error_msg = "No PDF has been associated to this request."
 end if

'Determine if there was an error
 if lcl_pdf_error_msg <> "" then
    response.write "<html>" & vbcrlf
    response.write "<head>" & vbcrlf
    response.write "  <script language=""javascript"">" & vbcrlf
    response.write "    function closeWindow() {" & vbcrlf
    response.write "      alert('" & lcl_pdf_error_msg & "');" & vbcrlf
    response.write "      parent.close();" & vbcrlf
    response.write "    }" & vbcrlf
    response.write "  </script>" & vbcrlf
    response.write "</head>" & vbcrlf
    response.write "<body onload=""closeWindow()"">" & vbcrlf
    response.write "</body>" & vbcrlf
    response.write "</html>" & vbcrlf
 end if

'------------------------------------------------------------------------------
 sub fillForm(sPDFPath,iRequestID)

  'Create PDF object
  	set oPDF  = Server.CreateObject("APToolkit.Object")
  	oDocument = oPDF.OpenOutputFile("MEMORY") 'CREATE THE OUTPUT INMEMORY

 	'Build PDF
  	oPDF.OutputPageWidth  = 612 ' 8.5 inches
  	oPDF.OutputPageHeight = 792 ' 11 inches

 	'Add form
   r = oPDF.OpenInputFile(sPDFPath)

	 'Add data to form
  	Call PopulateFormwithData(oPDF,iRequestID)
  	oPDF.FlattenRemainingFormFields = True 
  	r = oPDF.CopyForm(0, 0)

 	'Close PDF
  	oPDF.CloseOutputFile
  	oDocument = oPDF.binaryImage 

 	'Stream PDF to browser
  	response.expires = 0
  	response.Clear
  	response.ContentType = "application/pdf"
  	response.AddHeader "Content-Type", "application/pdf"
  	response.AddHeader "Content-Disposition", "inline;filename=FORMS.PDF"
  	response.BinaryWrite oDocument  

 	'Destory Objects
  	set oPDF      = nothing
  	set oDocument = nothing

 end sub

'------------------------------------------------------------------------------
sub setPDFFormFieldData(oPDF,iFieldName,iFieldValue,iReadOnly)

 'Object Properties: object.SetFormFieldData "FieldName", "FieldData", LeaveReadOnlyFlag
  r = oPDF.SetFormFieldData(iFieldName,iFieldValue,iReadOnly)

end sub

'------------------------------------------------------------------------------
sub PopulateFormwithData(oPDF,iRequestID)

  fnFillGeneralFields       oPDF
  fnFillStandardFields      oPDF,iRequestID
  fnFillIssueLocationFields oPDF,iRequestID
  fnFillDynamicFields       oPDF,iRequestID
  fnFillCodeSectionsField   oPDF,iRequestID
  fnFillTrackingNumberField oPDF,iRequestID

end sub

'------------------------------------------------------------------------------
sub fnFillGeneralFields(oPDF)

  lcl_readonly = 1

 	lcl_field_value = MonthName(Month(Date())) & " " & Day(Date()) & ", " & Year(Date())

  setPDFFormFieldData oPDF,"TodaysDate",lcl_field_value,lcl_readonly
	
end sub

'------------------------------------------------------------------------------
sub fnFillStandardFields(oPDF,iRequestID)

  lcl_readonly = 1

 'Get the userid from the request
  lcl_userid = getRequestUserID(iRequestID)

  if lcl_userid <> "" then
    	sSQL = "SELECT * FROM egov_users WHERE userid = '" & lcl_userid & "'"

    	set oFields = Server.CreateObject("ADODB.Recordset")
    	oFields.Open sSQL, Application("DSN"), 3, 1

    	if not oFields.eof then
      		for each field in oFields.fields

       			'Replace form field place holder with actual data
      		   sFieldPlaceHolder = field.name

        			if oFields(field.name) <> "" AND NOT IsNull(oFields(field.name)) then
              setPDFFormFieldData oPDF,sFieldPlaceHolder,oFields(field.name),lcl_readonly
        			end if
        next
     end if

     oFields.close
     set oFields = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub fnFillIssueLocationFields(oPDF,iRequestID)

 	sSQL = "SELECT streetnumber, "
  sSQL = sSQL & " dbo.fn_buildAddress('', "
  sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetprefix, ''), "
  sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetaddress, ''), "
  sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetsuffix, ''), "
  sSQL = sSQL &                      "ISNULL(dbo.egov_action_response_issue_location.streetdirection, '')) AS streetaddress, "
  sSQL = sSQL & " city, state, zip, comments "
  sSQL = sSQL & " FROM egov_action_response_issue_location "
  sSQL = sSQL & " WHERE actionrequestresponseid = '" & iRequestID & "'"

 	set oIssueFields = Server.CreateObject("ADODB.Recordset")
 	oIssueFields.Open sSQL, Application("DSN"), 3, 1

  lcl_readonly = 1
 	if not oIssueFields.eof then
     for each Field in oIssueFields.fields

			     'Replace form field place holder with actual data
      			sFieldPlaceHolder = field.name

      			if oIssueFields(field.name) <> "" AND NOT IsNull(oIssueFields(field.name)) then
            setPDFFormFieldData oPDF,sFieldPlaceHolder,oIssueFields(field.name),lcl_readonly
      	 	end if
    	next
 	end if

  oIssueFields.close
  set oIssueFields = nothing

end sub

'------------------------------------------------------------------------------
sub fnFillDynamicFields(oPDF,iRequestID)
  sSQL = "SELECT r.submitted_request_field_response, r.submitted_request_form_field_name, f.submitted_request_field_type_id "
  sSQL = sSQL & " FROM egov_submitted_request_fields f, egov_submitted_request_field_responses r "
  sSQL = sSQL & " WHERE r.submitted_request_field_id = f.submitted_request_field_id "
  sSQL = sSQL & " AND f.submitted_request_id = " & iRequestID
  sSQL = sSQL & " AND (r.submitted_request_form_field_name <> '' OR r.submitted_request_form_field_name IS NOT NULL) "
  sSQL = sSQL & " ORDER BY f.submitted_request_field_sequence "

  set oData = Server.CreateObject("ADODB.Recordset")
  oData.Open sSQL, Application("DSN"), 3, 1

  lcl_readonly = 1
  if not oData.eof then
     while not oData.eof
        lcl_field_value = ""

       'Remove any "returns" from the value and set the field to read-only.
        if oData("submitted_request_field_response") <> "" then
           lcl_field_value = oData("submitted_request_field_response")
           lcl_field_value = replace(lcl_field_value,chr(13),"")
           lcl_field_value = replace(lcl_field_value,chr(10),"")
           'lcl_field_value = replace(lcl_field_value,vbcrlf,"")

          'Setup the parameters to pass the PDF field.
          '  1. if the field type is either a checkbox then we need to pass the "value" as the "name" of the field instead 
               'and a "Yes" if the value exists to "check" the field.  This allows us to select multiple checkbox options.
          '  2. Checkbox fields type = 6
          '  3. If a record exists in this loop then the value has been selected.  Unselected values do NOT have records on the table.
           if oData("submitted_request_field_type_id") = 6 then
              setPDFFormFieldData oPDF,lcl_field_value,"Yes",lcl_readonly
           else
              setPDFFormFieldData oPDF,oData("submitted_request_form_field_name"),lcl_field_value,lcl_readonly
           end if
        end if

    			 oData.MoveNext
	 	  wend
 	end if

		oData.Close
 	set oData = nothing

end sub

'------------------------------------------------------------------------------
sub fnFillCodeSectionsField(oPDF,iRequestID)

  sSQL = "SELECT scs.submitted_action_code_id, cs.code_name, cs.description "
	 sSQL = sSQL & " FROM egov_submitted_request_code_sections scs, egov_actionline_code_sections cs "
 	sSQL = sSQL & " WHERE scs.submitted_action_code_id = cs.action_code_id "
	 sSQL = sSQL & " AND scs.submitted_request_id = " & iRequestID
 	sSQL = sSQL & " ORDER BY upper(cs.code_name), upper(cs.description) "

 	set oGetCodes = Server.CreateObject("ADODB.Recordset")
	 oGetCodes.Open sSQL, Application("DSN"), 3, 1

  lcl_readonly = 1
  if not oGetCodes.eof then

     lcl_display_codes = ""

	  	 while not oGetCodes.eof
       'Replace form field place holder with actual data
        if lcl_display_codes = "" then
           lcl_display_codes = "<strong>" & oGetCodes("code_name") & "</strong><br />&nbsp;&nbsp;&nbsp;" & oGetCodes("description") & "<br /><br />"
        else
           lcl_display_codes = lcl_display_codes & "<strong>" & oGetCodes("code_name") & "</strong><br />&nbsp;&nbsp;&nbsp;" & oGetCodes("description") & "<br /><br />"
        end if

        oGetCodes.movenext
 		  wend
  else
	    lcl_display_codes = ""
 	end if

  oGetCodes.close
  set oGetCodes = nothing

  sFieldPlaceHolder = "Code_Sections"

 'Set the PDF form field with the data
  setPDFFormFieldData oPDF,sFieldPlaceHolder,lcl_display_codes,lcl_readonly

end sub

'------------------------------------------------------------------------------
sub fnFillTrackingNumberField(oPDF,iRequestID)

  lcl_readonly = 1

  sSQL = "SELECT [tracking number] "
	 sSQL = sSQL & " FROM egov_rpt_actionline "
 	sSQL = sSQL & " WHERE action_autoid = " & iRequestID

 	set oRequestID = Server.CreateObject("ADODB.Recordset")
	 oRequestID.Open sSQL, Application("DSN"), 3, 1

 	if not oRequestID.eof then
     lcl_tracking_number = oRequestID("tracking number")
  else
	    lcl_tracking_number = ""
 	end if

  oRequestID.close
  set oRequestID = nothing

  sFieldPlaceHolder = "Tracking Number"

 'Set the PDF form field with the data
  setPDFFormFieldData oPDF,sFieldPlaceHolder,lcl_tracking_number,lcl_readonly

end sub

'------------------------------------------------------------------------------
function getRequestFormID(iRID)
  lcl_return = ""

  if dbready_number(iRID) then
     sSQL = "SELECT category_id "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE action_autoid = " & iRID

     set oFormID = Server.CreateObject("ADODB.Recordset")
     oFormID.Open sSQL, Application("DSN"), 3, 1

     if not oFormID.eof then
        lcl_return = oFormID("category_id")
     end if

     oFormID.close
     set oFormID = nothing

  end if

  getRequestFormID = lcl_return

end function

'------------------------------------------------------------------------------
function getRequestUserID(iRID)
  lcl_return = ""

  if dbready_number(iRID) then
     sSQL = "SELECT userid "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE action_autoid = " & iRID

     set oUserID = Server.CreateObject("ADODB.Recordset")
     oUserID.Open sSQL, Application("DSN"), 3, 1

     if not oUserID.eof then
        lcl_return = oUserID("userid")
     end if

     oUserID.close
     set oUserID = nothing

  end if

  getRequestUserID = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQLi, Application("DSN"), 3, 1
end sub
%>