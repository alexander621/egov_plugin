<%
'SET UP PAGE OPTIONS
' sDate = Month(Date()) & Day(Date()) & Year(Date())
sDate = year(date()) & month(date()) & day(date())
sTime = hour(time()) & minute(time()) & second(time())
server.scripttimeout = 4800
response.ContentType = "application/msexcel"
response.AddHeader "Content-Disposition", "attachment;filename=EGOV_ROCK_EXPORT_" & sDate & "_" & sTime & ".CSV"

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

'response.write request.cookies("User")("UserID")
'response.end

sSQL = session("DISPLAYQUERY")

sSQL = "SELECT  action_autoid, egov_action_request_view.orgid, userfname as [First Name], userlname as [Last Name], useremail as [Email], userhomephone as [Daytime Phone] " & vbcrlf _
	& mid(sSQL, instr(sSQL,"FROM egov_action_request_view"))

sSQL = LEFT(sSQL, instr(sSQL, "ORDER BY")-1)

sSQL = "SELECT [First Name], [Last Name], Email, [Daytime Phone], rm.answer as [Type of Removal], ad.answer as Address,sd.answer as StartDate, " & vbcrlf _
       	& " DATEADD(d,29,CONVERT(datetime,sd.answer)) as EndDate,MIN(a.parcelidnumber) as [Parcel ID (S/B/L)] " & vbcrlf _
	& " FROM (" & sSQL & ") ar " & vbcrlf _
	& " LEFT JOIN action_submitted_questions_and_answers ad ON ad.action_autoid = ar.action_autoid and ad.question = 'Address' " & vbcrlf _
	& " LEFT JOIN egov_residentaddresses a ON a.orgid=ar.orgid AND ad.answer = a.residentstreetnumber + ' ' + a.residentstreetname " & vbcrlf _
	& " LEFT JOIN action_submitted_questions_and_answers sd ON sd.action_autoid = ar.action_autoid and sd.question = 'Start Date' " & vbcrlf _
	& " LEFT JOIN action_submitted_questions_and_answers rm ON rm.action_autoid = ar.action_autoid and rm.question = 'Type of Removal' " & vbcrlf _
	& " GROUP BY ar.action_autoid, ar.[First Name], ar.[Last Name], ar.Email, ar.[Daytime Phone], sd.answer, ad.answer, rm.answer "

'response.write replace(sSQL,vbcrlf,"<br />")
'response.end
oSchema.Open sSQL, Application("DSN"), 3, 1

If Not oSchema.EOF Then

	'WRITE COLUMN HEADINGS
	For Each fldLoop in oSchema.Fields
		if fldLoop.Name <> "assigned_userID" and fldLoop.Name <> "assignedname" then response.write chr(34) & fldLoop.Name & chr(34) & ","
	Next
	response.write vbcrlf
	response.flush

	'WRITE DATA
	do while not oSchema.eof
		for each fldLoop in oSchema.Fields

				sFieldValue = fldLoop.Value
					if sFieldValue <> "" and not isnull(sFieldValue) then
						sFieldValue = replace(replace(trim(sFieldValue),chr(13),""), chr(10), "")
					end if
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

%>
