<!-- #include file="../includes/common.asp" //-->
<%
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
' 1.0	 06/02/06	 JOHN STULLENBERGER - INITIAL VERSION
' 2.0  11/12/07  David Boyer - Added ALL fields to export
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

' SET AS CSV FILE
response.buffer = false
Server.ScriptTimeout = 600  'in secs.  10 min.

'Build FileName
sFileName = replace(replace(replace(replace(replace(replace(Now(),":",""),"\","")," ",""),"AM",""),"PM",""),"/","_") & ".csv"
Response.ContentType = "application/msexcel"
Response.AddHeader "Content-Disposition", "attachment;filename=" & sFileName

'Build FilesystemName
fsFileName = Server.MapPath("tempcsvexport\" & sFileName)
Set objFSO=Server.CreateObject("Scripting.FileSystemObject")

'Open File
Set oReport = objFSO.OpenTextFile(fsFileName,2,True)

' GENERATE REPORT
Call SubListEvaluationResponses(REQUEST("iFormID"))

oReport.Close

if fsFileName <> "" then
   response.BinaryWrite ReadBinaryFile(fsFileName)
end if

objFSO.DeleteFile(fsFileName)

set oReport = nothing
set objFSO  = nothing

'------------------------------------------------------------------------------
Sub SubListEvaluationResponses(iFormID)
	Dim sToDate

'Check for org features
 lcl_orghasfeature_actionline_maintain_duedate = orghasfeature("actionline_maintain_duedate")

'Check for user permissions
 lcl_userhaspermission_actionline_maintain_duedate = userhaspermission(session("userid"),"actionline_maintain_duedate")

'Build WHERE Clause for Status/Sub-Status
 lcl_where_clause_status    = ""
' lcl_where_clause_substatus = ""

 for each oField in request.form
     if UCASE(LEFT(oField,9)) = "P_STATUS_" then
        if lcl_where_clause_status = "" then
           lcl_where_clause_status = " AND UPPER(status) IN ('" & REPLACE(UCASE(request.form(oField)),", ","', '")
        else
           lcl_where_clause_status = lcl_where_clause_status & "','" & REPLACE(UCASE(request.form(oField)),", ","', '")
        end if
'     elseif UCASE(LEFT(oField,12)) = "P_SUBSTATUS_" then
'        if lcl_where_clause_substatus = "" then
'           lcl_where_clause_substatus = " OR sub_status_id IN (" & request.form(oField)
'        else
'           lcl_where_clause_substatus = lcl_where_clause_substatus & "," & request.form(oField)
'        end if
     end if
 next

 if lcl_where_clause_status <> "" then
    lcl_where_clause_status = lcl_where_clause_status & "')"
 end if

' if lcl_where_clause_substatus <> "" then
'    if lcl_where_clause_status = "" then
'       lcl_where_clause_substatus = REPLACE(lcl_where_clause_substatus," OR ", " AND ")
'    end if

'    lcl_where_clause_substatus = lcl_where_clause_substatus & ")"
' end if

	sToDate = DateAdd("d", 1, request("todate"))

	sSQL = "SELECT [Tracking Number], [Form Name], status, comment, [Date Submitted] as submit_date, due_date, department, assignedname, [Submitted By], "
 sSQL = sSQL & " streetnumber, streetprefix, streetaddress, streetsuffix, streetdirection, sortstreetname, streetname AS COMPLETE_ISSUE_ADDRESS, "
 sSQL = sSQL & " city, state, zip, userfname, userlname, useraddress, useraddress2, usercity, userstate, usercity, userzip, useremail "
 sSQL = sSQL & " FROM egov_rpt_actionline "
 sSQL = sSQL & " WHERE action_formid='" & iFormID & "' "
 sSQL = sSQL & " AND [Date Submitted]  BETWEEN '" & request("fromdate") & "' AND '" & sToDate & "' "
 sSQL = sSQL & " AND orgid='" & session("orgid") & "' "
 sSQL = sSQL & lcl_where_clause_status
' sSQL = sSQL & lcl_where_clause_substatus

'dtb_debug(sSQL)

	' OPEN RECORDSET
	Set oData = Server.CreateObject("ADODB.Recordset")
	oData.Open sSQL, Application("DSN"), 3, 1

	' IF NOT EMPTY PROCESS RESULT SET
	If NOT oData.EOF Then
		CustomColumns = ""
		CustomColumns = subSeparateColumnFields(oData("comment")) 
		
		' WRITE COLUMN HEADINGS
  lcl_columnHeadings = ""
  lcl_columnHeadings = lcl_columnHeadings & "TRACKING NUMBER,"
  lcl_columnHeadings = lcl_columnHeadings & "FORM NAME,"
  lcl_columnHeadings = lcl_columnHeadings & "STATUS,"
  lcl_columnHeadings = lcl_columnHeadings & "SUBMIT DATE,"
  lcl_columnHeadings = lcl_columnHeadings & "SUBMITTED BY,"
  lcl_columnHeadings = lcl_columnHeadings & "ASSIGNED TO,"
  lcl_columnHeadings = lcl_columnHeadings & "DEPARTMENT,"

  if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
     lcl_columnHeadings = lcl_columnHeadings & "DUE DATE,"
  end if

  lcl_columnHeadings = lcl_columnHeadings & "ISSUE STREET NUMBER,"
  lcl_columnHeadings = lcl_columnHeadings & "ISSUE STREET NAME,"
  lcl_columnHeadings = lcl_columnHeadings & "ISSUE CITY,"
  lcl_columnHeadings = lcl_columnHeadings & "ISSUE STATE,"
  lcl_columnHeadings = lcl_columnHeadings & "ISSUE ZIP,"
  lcl_columnHeadings = lcl_columnHeadings & "FIRST NAME,"
  lcl_columnHeadings = lcl_columnHeadings & "LAST NAME,"
  lcl_columnHeadings = lcl_columnHeadings & "ADDRESS,"
  lcl_columnHeadings = lcl_columnHeadings & "ADDRESS 2,"
  lcl_columnHeadings = lcl_columnHeadings & "CITY,"
  lcl_columnHeadings = lcl_columnHeadings & "STATE,"
  lcl_columnHeadings = lcl_columnHeadings & "ZIP,"
  lcl_columnHeadings = lcl_columnHeadings & "EMAIL,"
  lcl_columnHeadings = lcl_columnHeadings & CustomColumns
		oReport.WriteLine lcl_columnHeadings

		' WRITE CUSTOM COLUM HEADINGS
		'response.write vbcrlf

  		lcl_street_name = ""
    lcl_columnData  = ""

		do while NOT oData.eof

  		 lcl_street_name = buildStreetAddress("", oData("streetprefix"), oData("streetaddress"), oData("streetsuffix"), oData("streetdirection"))
  	 	CustomFields    = ""
   		CustomFields    = subSeparateFormFields(oData("comment")) 

  		'WRITE BASIC FORM INFORMATION
  			lcl_columnData = chr(34) & oData("Tracking Number") & chr(34) & "," _
		   		& chr(34) & RemoveLineBreaks(oData("Form Name"))  & chr(34) & "," _
  					& chr(34) & oData("status")       & chr(34) & "," _
		  			& chr(34) & oData("submit_date")  & chr(34) & "," _
  					& chr(34) & oData("Submitted By") & chr(34) & "," _
		  			& chr(34) & oData("assignedname") & chr(34) & "," _
  					& chr(34) & oData("department")   & chr(34) & ","

       if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
          lcl_columnData = lcl_columnData & chr(34) & oData("due_date") & chr(34) & ","
       end if

  			lcl_columnData = lcl_columnData _
       & chr(34) & oData("streetnumber") & chr(34) & "," _
  					& chr(34) & lcl_street_name       & chr(34) & "," _
		  			& chr(34) & oData("city")         & chr(34) & "," _
  					& chr(34) & oData("state")        & chr(34) & "," _
		  			& chr(34) & oData("zip")          & chr(34) & "," _
  					& chr(34) & oData("userfname")    & chr(34) & "," _
		  			& chr(34) & oData("userlname")    & chr(34) & "," _
  					& chr(34) & oData("useraddress")  & chr(34) & "," _
		  			& chr(34) & oData("useraddress2") & chr(34) & "," _
  					& chr(34) & oData("usercity")     & chr(34) & "," _
		  			& chr(34) & oData("userstate")    & chr(34) & "," _
  					& chr(34) & oData("userzip")      & chr(34) & "," _
  					& chr(34) & oData("useremail")    & chr(34) & "," _
  					& CustomFields

		 	'WRITE FORM FIELD INFORMATION
     oReport.WriteLine lcl_columnData

		  	oData.MoveNext
		loop
 else
  lcl_columnData = chr(34) & chr(34)
  oReport.WriteLine lcl_columnData
	end if



'CLEAN UP OBJECTS
	set oData = nothing

end sub

'------------------------------------------------------------------------------
Function subSeparateFormFields(sText)
	ReturnString = ""
	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(UCASE(sText), "<P><B>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 1 to UBOUND(arrInfo)
		
			arrNamedPair = SPLIT(UCASE(arrInfo(w)),"<BR>")
			
			If ISARRAY(arrNamedPair) Then
				' WRITE DATA
				ReturnString = ReturnString & chr(34)
				ReturnString = ReturnString & RemoveLineBreaks(StripHTML(arrNamedPair(1)))
				ReturnString = ReturnString & chr(34) 
				ReturnString = ReturnString & ","
			End If

		Next

	End If
	subSeparateFormFields = ReturnString
End Function

'------------------------------------------------------------------------------
Function subSeparateColumnFields(sText)

	' MAKE SURE THERE IS INFORMATION TO PARSE
	If sText <> "" Then
	
		' BREAK LIST INTO SEPARATE LINES
		arrInfo = SPLIT(UCASE(sText), "<P><B>")

		' BREAK LINES INTO FIELD NAME AND VALUE
		For w = 1 to UBOUND(arrInfo)
		
			arrNamedPair = SPLIT(UCASE(arrInfo(w)),"<BR>")
			
			If ISARRAY(arrNamedPair) Then
				' WRITE DATA
				ReturnString = ReturnString & chr(34) & RemoveLineBreaks(StripHTML(arrNamedPair(0))) & chr(34) & ","
			End If

		Next

	End If
	subSeparateColumnFields = ReturnString
End Function

'------------------------------------------------------------------------------
Function StripHTML(asHTML) 
	
	Dim loRegExp ' Regular Expression Object 
	Dim theOutString ' string for output
	Dim theLastStringVal ' out string copy for loop comparison 
	Dim filteringComplete ' flag for filtering loop 
	' Create built In Regular Expression object to look for HTML tags 
	Set loRegExp = New RegExp
	loRegExp.Pattern = "<[^>]*>" 
	' Set the out string 
	theOutString = asHTML 
	
	' Loop through the out string looking for HTML and strip it filtering
	Complete = FALSE 
	While filteringComplete = FALSE 
 		theOutString = loRegExp.Replace(theOutString, "") 
	 	If theLastStringVal = theOutString Then 
		   	filteringComplete = TRUE 
  	End If 
		 theLastStringVal = theOutString 
 wend ' Return the original String stripped of HTML 

	StripHTML = theOutString 
	' Release object from memory 
	
	Set loRegExp = Nothing 

End Function 

'------------------------------------------------------------------------------
Function RemoveLineBreaks( sText )
	Dim sNewText 
	sNewText = Replace(sText,vbcrlf,"")
	sNewText = Replace(sNewText,Chr(10),"")
	sNewText = Replace(sNewText,Chr(13),"")
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
