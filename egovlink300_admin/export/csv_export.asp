<%
'SET UP PAGE OPTIONS
' sDate = Month(Date()) & Day(Date()) & Year(Date())
sDate = year(date()) & month(date()) & day(date())
sTime = hour(time()) & minute(time()) & second(time())
server.scripttimeout = 4800
response.ContentType = "application/msexcel"
response.AddHeader "Content-Disposition", "attachment;filename=EGOV_EXPORT_" & sDate & "_" & sTime & ".CSV"

'CREATE CSV FILE FOR DOWNLOAD
' CreateDownload

'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------
' FUNCTION CREATEDOWNLOAD()
'------------------------------------------------------------------------------------------------------------
' Sub CreateDownload()

Set oSchema = Server.CreateObject("ADODB.Recordset")

sSQL = session("DISPLAYQUERY")
oSchema.Open sSQL, Application("DSN"), 3, 1
'response.write sSQL
'response.end

If Not oSchema.EOF Then

	'WRITE COLUMN HEADINGS
	For Each fldLoop in oSchema.Fields
		response.write chr(34) & fldLoop.Name & chr(34) & ","
	Next
	response.write vbcrlf
	response.flush

	'WRITE DATA
	lcl_issue_street_number = ""
	do while not oSchema.eof
		for each fldLoop in oSchema.Fields
	 if fldLoop.name = "E-Gov Link Tracking #" then
		lcl_action_autoid = LEFT(fldLoop.value,LEN(fldLoop.value)-4)
	 else
		lcl_action_autoid = lcl_action_autoid
	 end if

				sFieldValue = trim(fldLoop.Value)

				'REMOVE LINE BREAKS
			if not isnull(sFieldValue) then
					sFieldValue = replace(sFieldValue,chr(10),"")
					sFieldValue = replace(sFieldValue,chr(13),"")
		sFieldValue = replace(sFieldValue,"default_novalue","")
		sFieldValue = replace(sFieldValue,"<p><b>","")
		sFieldValue = replace(sFieldValue,"</b><br></p>"," [] ")
		sFieldValue = replace(sFieldValue,"</b><br>"," [")
		sFieldValue = replace(sFieldValue,"</p>","] ")

	   '-- Asterisk for non-listed issue location address -----------------------
		if fldLoop.Name = "Non-Listed Address" then
		   if sFieldValue <> "Y" then
			  sFieldValue = "*"
		   else
			  sFieldValue = ""
		   end if
		end if

	   '-- Preferred Contact Method ---------------------------------------------
		if fldLoop.Name = "Preferred Contact Method" then
			sSQL1 = "SELECT rowid, contactdescription "
		   sSQL1 = sSQL1 & " FROM egov_contactmethods "
		   sSQL1 = sSQL1 & " WHERE rowid = " & sFieldValue

			Set rs1 = Server.CreateObject("ADODB.Recordset")
			rs1.Open sSQL1, Application("DSN"), 3, 1

		   if not rs1.eof then
			  sFieldValue = rs1("contactdescription")
		   else
			  sFieldValue = ""
		   end if
		end if

	   '-- Format FirstActionDate -----------------------------------------------
		if fldLoop.Name = "FirstActionDate" then
		   sFieldValue = FormatDateTime(sFieldValue,vbshortdate) & " " & FormatDateTime(sFieldValue,3)
		end if

	   '-- Issue Street Number --------------------------------------------------
		if fldLoop.Name = "Issue Street Number" then
		   lcl_issue_street_number = fldLoop.Value
		else
		   lcl_issue_street_number = lcl_issue_street_number
		end if

	   '-- Issue Street Name ----------------------------------------------------
		if fldLoop.Name = "Issue Street Name" then
		   lcl_sn_length = LEN(lcl_issue_street_number)
		   if isnumeric(LEFT(fldLoop.Value,lcl_sn_length)) then
			  if CLng(LEFT(fldLoop.Value,lcl_sn_length)) = CLng(lcl_issue_street_number) then
				 sFieldValue = REPLACE(fldLoop.Value,lcl_issue_street_number&" ","")
			  end if
		   end if
		end if

	   '-- Internal Only field questions/answers --------------------------------
		if fldLoop.Name = "Internal Values" then
		  'Retrieve all of the questions
		   sSQL = "SELECT submitted_request_field_id, submitted_request_field_prompt "
		   sSQL = sSQL & " FROM egov_submitted_request_fields "
		   sSQL = sSQL & " WHERE submitted_request_field_isinternal = 1 "
		   sSQL = sSQL & " AND submitted_request_id = " & lcl_action_autoid
		   sSQL = sSQL & " ORDER BY submitted_request_field_sequence "

		   set rs = Server.CreateObject("ADODB.Recordset")
		   rs.Open sSQL, Application("DSN"), 3, 1

		   if not rs.eof then
			  lcl_internal_values = ""
			  while not rs.eof
				'Retrieve all of the answers that have been selected for this question
				 sSQLa = "SELECT submitted_request_field_response "
				 sSQLa = sSQLa & " FROM egov_submitted_request_field_responses "
				 sSQLa = sSQLa & " WHERE submitted_request_field_id = " & rs("submitted_request_field_id")

				 set rsa = Server.CreateObject("ADODB.Recordset")
				 rsa.Open sSQLa, Application("DSN"), 3, 1

				 if not rsa.eof then
					lcl_answer      = ""
					lcl_answer_list = ""
					while not rsa.eof
					   lcl_answer = rsa("submitted_request_field_response")
									lcl_answer = replace(lcl_answer,chr(10),"")
								lcl_answer = replace(lcl_answer,chr(13),"")
					   lcl_answer = replace(lcl_answer,vbcrlf,", ")
					   lcl_answer = replace(lcl_answer,"default_novalue","")

					   if lcl_answer_list = "" then
						  lcl_answer_list = lcl_answer
					   else
						  lcl_answer_list = lcl_answer_list & ", " & lcl_answer
					   end if

					   rsa.movenext
					wend
				 else
					lcl_answers = ""
				 end if

				 lcl_internal_values = lcl_internal_values & rs("submitted_request_field_prompt") & " [" & lcl_answer_list & "] "
				 rs.movenext
			  wend
		   end if

		   sFieldValue = lcl_internal_values

		end if
				end if

'LEFT OFF HERE!!!!!
'             if sFieldValue <> "" then
'                sFieldValue = replace(sFieldValue,"&quot;","""")
'             end if

			response.write chr(34) & sFieldValue & chr(34) & ","
		next

		response.write vbcrlf
		response.flush

		oSchema.MoveNext
	Loop
Else

 'NO DATA

End If
Set oSchema = Nothing

' End Sub

sub dtb_debug(p_value)

  if p_value <> "" then
     lcl_value = replace(p_value,"'","''")
  else
     lcl_value = ""
  end if

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & lcl_value & "')"
  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

end sub
%>
