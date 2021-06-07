<!-- #include file="includes/common.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewXMLPDF.asp
' AUTHOR: Steve Loar
' CREATED: 02/13/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays a PDF.
'
' MODIFICATION HISTORY
' 1.0   05/06/2013	Steve Loar - Initial Version
'
' The calling JavaScript
' window.open('viewXMLPDF.asp?iRequestID=' + iRequestID + '&docID=' + iDocumentID + '&pdfaction=' + iAction + '&hideActLog=<%=lcl_hide_activitylog%')
' window.open('viewXMLPDF.asp?iRequestID=' + iRequestID + '&pdf=' + lcl_filename + '&pdfaction=' + iAction + '&hideActLog=<%=lcl_hide_activitylog%')
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'response.buffer = True 
Dim lcl_request_id, oRs, lcl_documentid

'Get variables
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
lcl_openExternalPDFviaURL = "Y" ' for the XML PDF to work, this needs to be a URL. SJL 2/7/2013

If dbready_number(request("iRequestID")) Then 
	lcl_request_id = request("iRequestID")

	'Get the request information
	sSql = "SELECT r.status, r.sub_status_id, "
	sSql = sSql & "ISNULL(r.public_actionline_pdf,f.public_actionline_pdf) AS public_actionline_pdf "
	sSql = sSql & "FROM egov_actionline_requests r, egov_action_request_forms f "
	sSql = sSql & "WHERE r.orgid = " & iorgid
	sSql = sSql & " AND r.category_id = f.action_form_id "
	sSql = sSql & " AND r.action_autoid = " & lcl_request_id

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		lcl_filename = oRs("public_actionline_pdf")
		lcl_comment_title = lcl_filename
		lcl_status = oRs("status")
		lcl_substatus = oRs("sub_status_id")
	End If 

	oRs.Close
	Set oRs = Nothing 

End If 


'Begin building the path
sPath = Application("CommunityLink_DocUrl") & "public_documents300/" & sorgVirtualSiteName & "/unpublished_documents" 

'Validate the file extension, Open the PDF, or show an error message.
If lcl_filename <> "" Then 
	'Check the extension to ensure it is a PDF file.
	If Right(UCase(lcl_filename),4) <> ".PDF" Then 
		lcl_pdf_error_msg = "<h1>No PDF has been associated to this request.</h1>"
	Else
		lcl_filename = sPath & lcl_filename
	End If 

	If lcl_pdf_error_msg = "" Then 
		'Retrieve the Action Line Request data - this is the pull to fill the PDF
		showForm lcl_filename, lcl_request_id
	End If 
Else 
	lcl_pdf_error_msg = "<h1>No PDF has been associated to this request.</h1>"
End If 


'Determine if there was an error
If lcl_pdf_error_msg <> "" Then 

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

End If 


'------------------------------------------------------------------------------
Sub ShowForm( ByVal sPDFPath, ByVal iRequestId )
	Dim bShow, iUserId

	bShow = True	' this controls whether this shows raw XML(false) or a PDF(true)
	
'	response.write "sPDFPath = " & sPDFPath & "<br />"
'	response.write "iRequestId = " & iRequestId & "<br /><br />"
'	response.End 

	If bShow Then 
		Response.ContentType = "application/vnd.adobe.xdp+xml"
		response.write "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
		response.write "<?xfa generator='AdobeDesigner_V7.0' APIVersion='2.2.4333.0'?>" & vbcrlf
		response.write "<xdp:xdp xmlns:xdp='http://ns.adobe.com/xdp/'>" & vbcrlf
		response.write "<xfa:datasets xmlns:xfa='http://www.xfa.org/schema/xfa-data/1.0/'>" & vbcrlf
		response.write "<xfa:data>" & vbcrlf
	End If 

	response.write "<form1>" & vbcrlf

	WriteXMLLine "TodaysDate", MonthName(Month(Date())) & " " & Day(Date()) & ", " & Year(Date())

	iUserId = getRequestUserID( iRequestId )
	getUserFields iUserId
	getIssueLocationFields iRequestId
	getDynamicFields iRequestId
	getCodeSectionsField iRequestId
	getTrackingNumber iRequestId

	response.write "</form1>" & vbcrlf 

	If bShow Then
		response.write "</xfa:data>" & vbcrlf
		response.write "</xfa:datasets>" & vbcrlf

		response.write "<pdf href='" & sPDFPath & "' xmlns='http://ns.adobe.com/xdp/pdf/' />" & vbcrlf
		response.write "</xdp:xdp>" & vbcrl
	End If 

	'response.write "<br /><br />All Done"
End Sub 


'------------------------------------------------------------------------------
Sub getUserFields( ByVal iUserId )
	Dim sSql, oRs

	If iUserId <> "" Then 
		sSql = "SELECT * FROM egov_users WHERE userid = " & iUserId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			For Each field In oRs.fields
				If oRs(field.name) <> "" And Not IsNull(oRs(field.name)) Then 
					WriteXMLLine field.name, oRs(field.name)
				End If 
			Next 
		End If 

		oRs.Close
		Set oRs = Nothing 

	End If 

End Sub 


'------------------------------------------------------------------------------
Sub getIssueLocationFields( ByVal iRequestID )
	Dim sSql, oRs

	sSql = "SELECT streetnumber, dbo.fn_buildAddress('', "
	sSql = sSql & "ISNULL(dbo.egov_action_response_issue_location.streetprefix, ''), "
	sSql = sSql & "ISNULL(dbo.egov_action_response_issue_location.streetaddress, ''), "
	sSql = sSql & "ISNULL(dbo.egov_action_response_issue_location.streetsuffix, ''), "
	sSql = sSql & "ISNULL(dbo.egov_action_response_issue_location.streetdirection, '')) AS streetaddress, "
	sSql = sSql & "city, state, zip, comments "
	sSql = sSql & "FROM egov_action_response_issue_location "
	sSql = sSql & "WHERE actionrequestresponseid = " & iRequestID 

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		For Each Field In oRs.fields
			If oRs(field.name) <> "" And Not IsNull(oRs(field.name)) Then 
				WriteXMLLine field.name, oRs(field.name)
			End If 
		Next 
	End If 

	oRs.Close
	Set oRs = Nothing  

End Sub 


'------------------------------------------------------------------------------
Sub getDynamicFields( ByVal iRequestID )
	Dim sSql, oRs, lcl_field_value

	sSql = "SELECT r.submitted_request_field_response, ISNULL(r.submitted_request_form_field_name,'') AS submitted_request_form_field_name, f.submitted_request_field_type_id "
	sSql = sSql & "FROM egov_submitted_request_fields f, egov_submitted_request_field_responses r "
	sSql = sSql & "WHERE r.submitted_request_field_id = f.submitted_request_field_id "
	sSql = sSql & "AND f.submitted_request_id = " & iRequestID
	sSql = sSql & " AND (r.submitted_request_form_field_name <> '' OR r.submitted_request_form_field_name IS NOT NULL) "
	sSql = sSql & "ORDER BY f.submitted_request_field_sequence"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		lcl_field_value = ""

		'Remove any "returns" from the value and set the field to read-only.
		If oRs("submitted_request_field_response") <> "" Then 
			lcl_field_value = oRs("submitted_request_field_response")
			lcl_field_value = Replace(lcl_field_value,Chr(13),"")
			lcl_field_value = Replace(lcl_field_value,Chr(10),"")

			'Setup the parameters to pass the PDF field.
			'  1. if the field type is either a checkbox then we need to pass the "value" as the "name" of the field instead 
			'and a "Yes" if the value exists to "check" the field.  This allows us to select multiple checkbox options.
			'  2. Checkbox fields type = 6
			'  3. If a record exists in this loop then the value has been selected.  Unselected values do NOT have records in the table.
			If oRs("submitted_request_field_type_id") = 6 Then 
				'WriteXMLLine LCase(Replace(lcl_field_value," ","")), LCase(Replace(lcl_field_value," ",""))
				WriteXMLLine LCase(Replace(lcl_field_value," ","")), "Yes"
			ElseIf oRs("submitted_request_field_type_id") = 2 Then 
				' Radio Buttons
				WriteXMLLine oRs("submitted_request_form_field_name"), LCase(Replace(ReplaceBreaksForPDFs( lcl_field_value )," ",""))
			Else 
				' These are the merge field names set on on the form creator, the value is the answer from the form
				If oRs("submitted_request_form_field_name") <> "" Then 
					WriteXMLLine oRs("submitted_request_form_field_name"), ReplaceBreaksForPDFs( lcl_field_value )
				End If 
			End If 
		End If 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub getCodeSectionsField( ByVal iRequestId )
	Dim sSql, oRs, lcl_display_codes

	lcl_display = ""

	sSql = "SELECT ISNULL(cs.code_name,'') AS code_name, ISNULL(cs.description, '') AS description "
	sSql = sSql & "FROM egov_submitted_request_code_sections scs, egov_actionline_code_sections cs "
	sSql = sSql & "WHERE scs.submitted_action_code_id = cs.action_code_id "
	sSql = sSql & "AND scs.submitted_request_id = " & iRequestId
	sSql = sSql & " ORDER BY UPPER(cs.code_name), UPPER(cs.description)"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		lcl_display = lcl_display & oRs("code_name") & vbcrlf & ReplaceBreaksForPDFs( oRs("description") ) & vbcrlf

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	If lcl_display <> "" Then 
		WriteXMLLine "Code_Sections", lcl_display
	End If 

End Sub 


'------------------------------------------------------------------------------
Sub getTrackingNumber( ByVal iRequestId )
	Dim sSql, oRs

	sSql = "SELECT [tracking number] "
	sSql = sSql & "FROM egov_rpt_actionline "
	sSql = sSql & "WHERE action_autoid = " & iRequestId

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		WriteXMLLine "Tracking Number", oRs("tracking number")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Function getRequestUserID( ByVal iRequestId )
	Dim sSql, oRs, iUserId

	iUserId = "0"

	If dbready_number( iRequestId ) Then 
		sSql = "SELECT userid FROM egov_actionline_requests "
		sSql = sSql & " WHERE action_autoid = " & iRequestId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			iUserId = oRs("userid")
		End If 

		oRs.Close
		Set oRs = Nothing 

	End If 

	getRequestUserID = iUserId

End Function 


'--------------------------------------------------------------------------------------------------
' string ReplaceBreaksForPDFs sFieldValue 
'--------------------------------------------------------------------------------------------------
Function ReplaceBreaksForPDFs( ByVal sFieldValue )
	
	If Not IsNull(sFieldValue) Then 
		sFieldValue = Replace( sFieldValue, "<br />", vbcrlf )
		sFieldValue = Replace( sFieldValue, "<br/>", vbcrlf )
		sFieldValue = Replace( sFieldValue, "<br>", vbcrlf )
	Else
		sFieldValue = ""
	End If 

	ReplaceBreaksForPDFs = sFieldValue

End Function 


'--------------------------------------------------------------------------------------------------
' void WriteXMLLine sNodeName, sValue
'--------------------------------------------------------------------------------------------------
Sub WriteXMLLine( ByVal sNodeName, ByVal sValue )

	' handle reserved XML characters
	sValue = Replace(sValue, "&", "&amp;")
	sValue = Replace(sValue, ">", "&gt;")
	sValue = Replace(sValue, "<", "&lt;")
	sValue = Replace(sValue, "'", "&apos;")
	sValue = replace(sValue, "’", "&apos;")
	sValue = replace(sValue, "‘", "&apos;")
	sValue = Replace(sValue, "%", "&#37;")
	'sValue = Replace(sValue, "(", "&#40;")
	'sValue = Replace(sValue, ")", "&#41;")
	'sValue = Replace(sValue, "(", "")
	'sValue = Replace(sValue, ")", "")
	'sValue = Replace(sValue, "-", "")
	sValue = Trim(sValue)

	' This part is for the custom fields
	sNodeName = Replace(sNodeName, "(", "")
	sNodeName = Replace(sNodeName, ")", "")
	sNodeName = Replace(sNodeName, "-", "")
	sNodeName = Replace(sNodeName, " ", "")
	sNodeName = Replace(sNodeName, "<", "")
	sNodeName = Replace(sNodeName, ">", "")
	sNodeName = Replace(sNodeName, "&", "")
	sNodeName = Replace(sNodeName, "'", "")
	sNodeName = Replace(sNodeName, "%", "")
	sNodeName = Trim(sNodeName)

	response.write "<" & sNodeName & ">" & sValue & "</" & sNodeName & ">" & vbcrlf

End Sub 



%>