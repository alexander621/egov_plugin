<!-- #include file="../includes/common.asp" //-->
<%
'TWFStart = Timer()
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: EVAL_EXPORT.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 06/02/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	 06/02/06	John Stullenberger - Initial Version
' 2.0  11/12/07 David Boyer - Added ALL fields to export
' 2.1  08/18/08 David Boyer - Added Status and Sub-Status to WHERE Clause
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

' SET AS CSV FILE
Server.ScriptTimeout = 600  'in secs.  10 min.

sFileName = replace(replace(replace(replace(replace(replace(Now(),":",""),"\","")," ",""),"AM",""),"PM",""),"/","_") & ".csv"
Response.ContentType = "application/msexcel"
Response.AddHeader "Content-Disposition", "attachment;filename=" & sFileName

'Build FilesystemName
fsFileName = Server.MapPath("tempcsvexport\" & sFileName)
Set objFSO=Server.CreateObject("Scripting.FileSystemObject")

'Open File
Set oReport = objFSO.OpenTextFile(fsFileName,2,True)

'GENERATE REPORT
 Call SubListEvaluationResponses(REQUEST("iFormID"))
' oReport.WriteLine Timer() - TWFStart

oReport.Close

response.BinaryWrite ReadBinaryFile(fsFileName)

objFSO.DeleteFile(fsFileName)


Set oReport = Nothing
Set objFSO = Nothing

'------------------------------------------------------------------------------
Sub SubListEvaluationResponses(iFormID)
	Dim sToDate

'Check for org features
 lcl_orghasfeature_action_line_substatus       = orghasfeature("action_line_substatus")
 lcl_orghasfeature_issue_location              = orghasfeature("issue location")
 lcl_orghasfeature_actionline_maintain_duedate = orghasfeature("actionline_maintain_duedate")

'Check for user permissions
 lcl_userhaspermission_actionline_maintain_duedate = userhaspermission(session("userid"),"actionline_maintain_duedate")

'Build WHERE Clause for Status/Sub-Status
 lcl_where_clause_status    = ""
 lcl_where_clause_substatus = ""

 for each oField in request.form
     if UCASE(LEFT(oField,9)) = "P_STATUS_" then
        if lcl_where_clause_status = "" then
           lcl_where_clause_status = " AND UPPER(status) IN ('" & REPLACE(UCASE(request.form(oField)),", ","', '")
        else
           lcl_where_clause_status = lcl_where_clause_status & "','" & REPLACE(UCASE(request.form(oField)),", ","', '")
        end if
     elseif UCASE(LEFT(oField,12)) = "P_SUBSTATUS_" then
        if lcl_where_clause_substatus = "" then
           lcl_where_clause_substatus = " OR sub_status_id IN (" & request.form(oField)
        else
           lcl_where_clause_substatus = lcl_where_clause_substatus & "," & request.form(oField)
        end if
     end if
 next

 if lcl_where_clause_status <> "" then
    lcl_where_clause_status = lcl_where_clause_status & "')"
 end if

 if lcl_where_clause_substatus <> "" then
    if lcl_where_clause_status = "" then
       lcl_where_clause_substatus = REPLACE(lcl_where_clause_substatus," OR ", " AND ")
    end if

    lcl_where_clause_substatus = lcl_where_clause_substatus & ")"
 end if

	sToDate = DateAdd("d", 1, request("todate"))
	sSQL = "SELECT [Tracking Number], [Form Name], status, comment, [Date Submitted] as submit_date, due_date, department, assignedname, "
 sSQL = sSQL & " [Submitted By], streetnumber, streetprefix, streetaddress, streetsuffix, streetdirection, sortstreetname, "
 sSQL = sSQL & " streetname AS COMPLETED_ISSUE_ADDRESS, city, state, zip, parcelidnumber, comments as issue_comments, validstreet, "
 sSQL = sSQL & " userfname, userlname, useraddress, useraddress2, usercity, userstate, usercity, userzip, useremail, "
 sSQL = sSQL & " action_autoid, sub_status_desc, last_edit_date "
 sSQL = sSQL & " FROM egov_rpt_actionline "
 sSQL = sSQL & " WHERE action_formid='" & iFormID & "' "
 sSQL = sSQL & " AND [Date Submitted]  BETWEEN '" & request("fromdate") & "' AND '" & sToDate & "' "
 sSQL = sSQL & " AND orgid='" & session("orgid") & "'"
 sSQL = sSQL & lcl_where_clause_status
 sSQL = sSQL & lcl_where_clause_substatus

'OPEN RECORDSET
	set oData = Server.CreateObject("ADODB.Recordset")
	oData.Open sSQL, Application("DSN"), 3, 1

'IF NOT EMPTY PROCESS RESULT SET
	if NOT oData.eof then

 		'WRITE COLUMN HEADINGS
		  ColumnHeadings = ""
   	ColumnHeadings = ColumnHeadings & "TRACKING NUMBER,FORM NAME,STATUS,"

   'Check for Org Feature: Sub-Status
    if lcl_orghasfeature_action_line_substatus then
     		ColumnHeadings = ColumnHeadings & "SUB-STATUS,"
    end if

   	ColumnHeadings = ColumnHeadings & "SUBMIT DATE,SUBMITTED BY,LAST EDIT DATE,ASSIGNED TO,DEPARTMENT,"

   'Check for Org Feature: Due Date
    if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
       ColumnHeadings = ColumnHeadings & "DUE DATE,"
    end if

   'Check for Org Feature: Issue Location
    if lcl_orghasfeature_issue_location then
    			ColumnHeadings = ColumnHeadings & "ISSUE STREET NUMBER,ISSUE STREET NAME,ISSUE CITY,ISSUE STATE,ISSUE ZIP,ISSUE PARCEL ID,ISSUE COMMENTS,NON-LISTED ADDRESS,"
    end if

   	ColumnHeadings = ColumnHeadings & "FIRST NAME,LAST NAME,ADDRESS,ADDRESS 2,CITY,STATE,ZIP,EMAIL,"

 		'WRITE CUSTOM COLUM HEADINGS
     		'subSeparateColumnFields(oData("comment")) 
     		ColumnHeadings = ColumnHeadings & subSeparateColumnFields(iFormID,"U")
     		ColumnHeadings = ColumnHeadings & subSeparateColumnFields(iFormID,"I")
     		oReport.WriteLine ColumnHeadings

     		lcl_street_name = ""

   		do while NOT oData.eof

     			lcl_street_name = buildStreetAddress("", oData("streetprefix"), oData("streetaddress"), oData("streetsuffix"), oData("streetdirection"))

    			'WRITE BASIC FORM INFORMATION
     			RowData = chr(34) & removelinebreaks(oData("Tracking Number")) & chr(34) & "," _
            				& chr(34) & removelinebreaks(oData("Form Name"))       & chr(34) & "," _
            				& chr(34) & oData("status")          & chr(34) & ","

       'OrgFeature: Sub-Status
     			if lcl_orghasfeature_action_line_substatus then
       				RowData = RowData & chr(34) & oData("sub_status_desc") & chr(34) & ","
     			end if

     			RowData = RowData & chr(34) & oData("submit_date")    & chr(34) & "," _
                      				& chr(34) & oData("Submitted By")   & chr(34) & "," _
                      				& chr(34) & oData("last_edit_date") & chr(34) & "," _
                      				& chr(34) & oData("assignedname")   & chr(34) & "," _
                      				& chr(34) & oData("department")     & chr(34) & ","

       'OrgFeature: Due Date
        if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
           RowData = RowData & chr(34) & oData("due_date") & chr(34) & ","
        end if

       'OrgFeature: Issue Location
     			if lcl_orghasfeature_issue_location then
  		     		if oData("validstreet") <> "Y" then
       		  			lcl_valid_street = "*"
       				else
         					lcl_valid_street = ""
       				end if

       				RowData = RowData & chr(34) & oData("streetnumber")   & chr(34) & "," _
  	                      				& chr(34) & lcl_street_name         & chr(34) & "," _
                        					& chr(34) & oData("city")           & chr(34) & "," _
                        					& chr(34) & oData("state")          & chr(34) & "," _
                        					& chr(34) & oData("zip")            & chr(34) & "," _
                         				& chr(34) & oData("parcelidnumber") & chr(34) & "," _
                         				& chr(34) & oData("issue_comments") & chr(34) & "," _
                        					& chr(34) & lcl_valid_street        & chr(34) & ","
     			end if

    			'WRITE USER INFORMATION
  	   		RowData = RowData & chr(34) & oData("userfname")    & chr(34) & "," _
                      				& chr(34) & oData("userlname")    & chr(34) & "," _
                      				& chr(34) & oData("useraddress")  & chr(34) & "," _
                      				& chr(34) & oData("useraddress2") & chr(34) & "," _
                      				& chr(34) & oData("usercity")     & chr(34) & "," _
                      				& chr(34) & oData("userstate")    & chr(34) & "," _
                      				& chr(34) & oData("userzip")      & chr(34) & "," _
                      				& chr(34) & oData("useremail")    & chr(34) & ","

   			 'WRITE FORM FIELD INFORMATION
     			RowData = RowData & subSeparateFormFields_new(oData("comment"),oData("action_autoid"),iFormID,"U")
  	   		RowData = RowData & subSeparateFormFields_new(oData("comment"),oData("action_autoid"),iFormID,"I")

       	oReport.WriteLine RowData

     			oData.MoveNext
   		loop
 else
     RowData = chr(34) & chr(34)

    	oReport.WriteLine RowData
 end If


'CLEAN UP OBJECTS
	set oData = nothing

end sub

'------------------------------------------------------------------------------
Function subSeparateFormFields_new(sText,p_request_id,p_form_id,p_column_type)
	ReturnString = ""
	if sText <> "" AND p_form_id <> "" then
		'p_column_type values:
		'  1. I = Internal Questions
		'  2. U = User Questions
		
		'Retrieve all of the "submitted" questions that exist on request for a specific form (i.e. Appliance Collection)
		'that match the "setup" questions for the same form.
 		sSQL = "SELECT DISTINCT r.submitted_request_field_id, submitted_request_field_prompt AS prompt, "
 		sSQL = sSQL & " submitted_request_field_sequence AS sequence, "  ','answers go here' AS answers
 		sSQL = sSQL & " CAST(resp.submitted_request_field_response as VARCHAR(MAX)) AS Response "
 		sSQL = sSQL & " FROM egov_submitted_request_fields r "
		 sSQL = sSQL & " INNER JOIN egov_action_form_questions q ON upper(q.prompt) = upper(r.submitted_request_field_prompt) "
	 	sSQL = sSQL & " LEFT JOIN egov_submitted_request_field_responses resp ON resp.submitted_request_field_id = r.submitted_request_field_id "
 		sSQL = sSQL & " WHERE   (submitted_request_id IN (SELECT action_autoid "
 		sSQL = sSQL &                               " FROM egov_actionline_requests "
 		sSQL = sSQL &                               " WHERE orgid = " & session("orgid")
 		sSQL = sSQL &                               " AND category_id = " & p_form_id & ")) "
		
 		if p_column_type = "I" then
			sSQL = sSQL & " AND submitted_request_field_isinternal = 1 "
 		else
			sSQL = sSQL & " AND (submitted_request_field_isinternal = 0 OR "
			sSQL = sSQL &      " submitted_request_field_isinternal IS NULL) "
 		end if
		
	 	sSQL = sSQL & " AND submitted_request_id = " & p_request_id
	
		'-----------------------------
 		sSQL = sSQL & " UNION ALL "
		'-----------------------------
		'Retrieve all remaining "setup" questions for the form that do NOT exist on the "submitted" table
 		sSQL = sSQL & " SELECT DISTINCT '' AS submitted_request_field_id, prompt, sequence, NULL AS Response " ','' AS answers
 		sSQL = sSQL & " FROM egov_action_form_questions "
 		sSQL = sSQL & " WHERE orgid = " & session("orgid")
 		sSQL = sSQL & " AND formid = " & p_form_id
	
		if p_column_type = "I" then
		  	sSQL = sSQL & " AND isinternalonly = 1 "
		else
		  	sSQL = sSQL & " AND (isinternalonly = 0 OR isinternalonly IS NULL) "
		end if
	
		sSQL = sSQL & " AND (UPPER(prompt) NOT IN "
		sSQL = sSQL &             " (SELECT DISTINCT UPPER(submitted_request_field_prompt) " 'Same query that is the first part of this UNION
		sSQL = sSQL &              " FROM egov_submitted_request_fields "
		sSQL = sSQL &              " WHERE (submitted_request_id IN "
		sSQL = sSQL &                        " (SELECT action_autoid "
		sSQL = sSQL &                         " FROM egov_actionline_requests "
		sSQL = sSQL &                         " WHERE orgid = " & session("orgid")
		sSQL = sSQL &                         " AND category_id = " & p_form_id & ")) "
		
		if p_column_type = "I" then
		  	sSQL = sSQL &                      " AND submitted_request_field_isinternal = 1 "
		else
  			sSQL = sSQL &                      " AND (submitted_request_field_isinternal = 0 OR "
	  		sSQL = sSQL &                            "submitted_request_field_isinternal IS NULL) "
		end if
	
		sSQL = sSQL &                         " AND (submitted_request_id = " & p_request_id & "))) "
		sSQL = sSQL & " ORDER BY 3, 2 "

		PROBLEMCHILD = False
		'if p_request_id = "3549" then PROBLEMCHILD = TRUE
		'response.write sSQL & "<p>"
		'response.end
	
		set rs = Server.CreateObject("ADODB.Recordset")
    		rs.Open sSQL, Application("DSN"), 3, 1
	
		if not rs.eof then
			testEOF = False
			do while not rs.eof and NOT testEOF
				'Retrieve all of the internal questions that have been submitted on the current form.
         			'We need this because to compare what has been submitted to what is currently ON the form.
				lcl_answers = ""
	
          			'Compare the internal questions that have been submitted and the questions that are currently ON the form.
				'Retrieve all of the answers for each question, submitted, and build the answer to display
	
      				if not IsNull(rs("Response")) then
					Question = rs("Prompt")
					SameQuestion = True
					IF PROBLEMCHILD then response.write "ORIGINAL QUESTION: " & Question & "<br>"

        				do while not rs.EOF and SameQuestion
						if rs("Response") = "default_novalue" then
							lcl_field_response = ""
						else
							lcl_field_response = rs("Response")
						end if
	
						if lcl_field_response <> "" then
							if lcl_answers = "" then
								lcl_answers = lcl_field_response
							else
								lcl_answers = lcl_answers & ", " & lcl_field_response
							end if
						end if

						'Check to see if the next question 
						rs.movenext
						if not rs.EOF then
							If rs("Prompt") <> Question then 
								SameQuestion = False
								rs.MovePrevious
							end if
							IF PROBLEMCHILD then response.write "NEXT Question (" & SameQuestion & ")" & rs("Prompt") & "<br>"
						end if
					loop
					if rs.EOF then 
						rs.MovePrevious
						testEOF = TRUE
					end if
					IF PROBLEMCHILD then response.write "ANSWER(s) FOR REPORT " & lcl_answers & "<br>"
					IF PROBLEMCHILD then response.write "<br>"

	
					ReturnString = ReturnString & chr(34) & RemoveLineBreaks(StripHTML(lcl_answers)) & chr(34) & ","
				else
					if INSTR(REPLACE(REPLACE(REPLACE(UCASE(sText),"<B>","<B>[")," </B>","</B>"),"</B>","]</B>"),UCASE("["&TRIM(rs("prompt"))&"]")) > 0 then
						'This may be an older version, so double-check the previous way of saving these values
						'BREAK LIST INTO SEPARATE LINES
						arrInfo = SPLIT(UCASE(sText),"<P><B>")
						
						'BREAK LINES INTO FIELD NAME AND VALUE
						For w = 1 to UBOUND(arrInfo)
							if INSTR("["&REPLACE(REPLACE(UCASE(arrInfo(w))," </B>","</B>"),"</B>","]</B>"),UCASE("["&TRIM(rs("prompt"))&"]")) > 0 then
								arrNamedPair = SPLIT(UCASE(arrInfo(w)),"<BR>")
						
								If ISARRAY(arrNamedPair) Then
									'WRITE DATA
									ReturnString = ReturnString & chr(34) & RemoveLineBreaks(StripHTML(arrNamedPair(1))) & chr(34) & ","
								else
									ReturnString = ReturnString & chr(34) & chr(34) & ","
								End If
							end if
						Next
					else
   						ReturnString = ReturnString & chr(34) & chr(34) & ","
 					end if
				end if

      rs.movenext
			loop
			set rs = nothing
		end if
	end if
	subSeparateFormFields_new = ReturnString
End Function

'------------------------------------------------------------------------------
Function subSeparateColumnFields(p_form_id, p_column_type)
	ReturnString = ""
 'p_column_type values:
 '  1. I = Internal Questions
 '  2. U = User Questions

  if p_form_id <> "" then
    'Retrieve all of the questions for the form
     sSQL = "SELECT prompt "
     sSQL = sSQL & " FROM egov_action_form_questions "
     sSQL = sSQL & " WHERE orgid = " & session("orgid")
     sSQL = sSQL & " AND formid = "  & p_form_id
     if p_column_type = "I" then
        sSQL = sSQL & " AND isinternalonly = 1 "
     else
        sSQL = sSQL & " AND (isinternalonly is NULL OR isinternalonly = 0) "
     end if
     sSQL = sSQL & " ORDER BY sequence, prompt "

    	set oColumn = Server.CreateObject("ADODB.Recordset")
    	oColumn.Open sSQL, Application("DSN"), 3, 1

     if not oColumn.eof then
        if p_column_type = "I" then
           lcl_column_header = "INTERNAL: "
        else
           lcl_column_header = ""
        end if

        while not oColumn.eof
		  ReturnString = ReturnString & chr(34) & lcl_column_header & oColumn("prompt") & chr(34) & ","
          oColumn.movenext
        wend
     end if

     set oColumn = nothing

  end if
	subSeparateColumnFields = ReturnString
end Function

'------------------------------------------------------------------------------
Function StripHTML(asHTML) 
	
	Dim loRegExp           'Regular Expression Object 
	Dim theOutString       'string for output
	Dim theLastStringVal   'out string copy for loop comparison 
	Dim filteringComplete  'flag for filtering loop 

'Create built In Regular Expression object to look for HTML tags 
	Set loRegExp = New RegExp
	loRegExp.Pattern = "<[^>]*>" 

'Set the out string 
	theOutString = asHTML 
	
'Loop through the out string looking for HTML and strip it filtering
	Complete = FALSE 
	while filteringComplete = FALSE 
    theOutString = loRegExp.Replace(theOutString, "") 
	 	 If theLastStringVal = theOutString Then 
		    	filteringComplete = TRUE 
  	 End If 
		  theLastStringVal = theOutString 
 wend 'Return the original String stripped of HTML 

	StripHTML = theOutString 

'Release object from memory 
	Set loRegExp = Nothing 

End Function 

'------------------------------------------------------------------------------
Function RemoveLineBreaks( sText )
 	Dim sNewText 
 	sNewText = REPLACE(sText,vbcrlf,"")
 	sNewText = REPLACE(sNewText,chr(10),"")
 	sNewText = REPLACE(sNewText,chr(13),"")
 	RemoveLineBreaks = sNewText
End Function

'------------------------------------------------------------------------------
Function ReadBinaryFile(FileName)
  Const adTypeBinary = 1
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To get binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream
  BinaryStream.Open
  
  'Load the file data from disk To stream object
  BinaryStream.LoadFromFile FileName
  
  'Open the stream And get binary data from the object
  ReadBinaryFile = BinaryStream.Read
End Function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
 	set oInsert = Server.CreateObject("ADODB.Recordset")
 	oInsert.Open sSQLi, Application("DSN"), 3, 1

  set oInsert = nothing

end sub
%>
