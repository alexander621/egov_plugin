<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: icalMonthExport.asp
' AUTHOR: SteveLoar
' CREATED: 1/24/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description: This pulls a 1 month of calendar events into an iCal format file
'
' MODIFICATION HISTORY
' 1.0   1/24/2014		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, dEventDate, sUTCEventDate, iEventDuration, sEndDate, dEndDate, sMessage
Dim sIcalDateStamp, dIcalDateStamp, sStartDate, iCategoryId

If request("startdate") <> "" Then
	sStartDate = request("startdate")
	If IsDate(sStartDate) Then
		sStartDate = CDate(request("startdate"))
	Else	
		sStartDate = Date()
	End If
Else
	sStartDate = Date()
End If

sStartDate = CDate(Month(sStartDate) & "/1/" & Year(sStartDate))

'response.write "startdate passed = " & request("startdate") & "<br />"
'response.write "startdate set to = " & request("startdate") & "<br /><br />"

If request("categoryid") <> "" Then
	iCategoryId = CLng(request("categoryid"))
Else
	iCategoryId = 0 
End If

' SET UP PAGE OPTIONS
server.scripttimeout = 9000
Response.ContentType = "text/calendar"
Response.AddHeader "Content-Disposition", "attachment;filename=events.ics"

' Get and output the data
sSql = "SELECT eventdate, eventduration, dbo.GetUTCTime( " & iorgid & ", eventdate ) AS utctime, "
sSql = sSql & "dbo.GetUTCTime( " & iorgid & ", getdate() ) AS icaldatestamp, "
sSql = sSql & "subject, message "
sSql = sSql & "FROM events WHERE orgid = " & iorgid & " AND (calendarfeature = '' OR calendarfeature IS NULL) "
sSql = sSql & " AND eventdate > '" & DateValue(sStartDate) & "' AND eventdate < '" & DateValue(DateAdd("m", 1, sStartDate)) & "' "
If iCategoryId > CLng(0) Then
	sSql = sSql & " AND categoryid = " & iCategoryId
End If
sSql = sSql & " ORDER BY eventdate"
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	response.write "BEGIN:VCALENDAR" & vbcrlf
	response.write "VERSION:2.0" & vbcrlf
	response.write "PRODID:-//EGOVLINK//EN" & vbcrlf
	response.write "METHOD:PUBLISH" & vbcrlf
	'Response.write "X-MS-OLK-FORCEINSPECTOROPEN:TRUE" & vbcrlf
	response.flush

	Do While Not oRs.EOF
		response.write "BEGIN:VEVENT" & vbcrlf
		' 20101119T220000Z
		dEventDate = CDate(oRs("utctime"))
		sUTCEventDate = Year(dEventDate) & Right("0" & Month(dEventDate),2) & Right("0" & Day(dEventDate),2) & "T" & Right("0" & Hour(dEventDate),2) & Right("0" & Minute(dEventDate),2) & "00Z"
		iEventDuration = CLng(oRs("eventduration"))
		If iEventDuration > CLng(0) Then
			dEndDate = DateAdd("n", iEventDuration, dEventDate)
			sEndDate = Year(dEndDate) & Right("0" & Month(dEndDate),2) & Right("0" & Day(dEndDate),2) & "T" & Right("0" & Hour(dEndDate),2) & Right("0" & Minute(dEndDate),2) & "00Z"
		Else
			sEndDate = sUTCEventDate
		End If 
		response.write "DTSTART:" & sUTCEventDate & vbcrlf
		response.write "DTEND:" & sEndDate & vbcrlf
		response.write "TRANSP:OPAQUE" & vbcrlf
		response.write "SEQUENCE:0" & vbcrlf
		response.write "UID:" & generateRequestID( 200) & "@eclink.com" & vbcrlf
		dIcalDateStamp = CDate(oRs("icaldatestamp"))
		sIcalDateStamp = Year(dIcalDateStamp) & Right("0" & Month(dIcalDateStamp),2) & Right("0" & Day(dIcalDateStamp),2) & "T" & Right("0" & Hour(dEventDate),2) & Right("0" & Minute(dIcalDateStamp),2) & "00Z"
		response.write "DTSTAMP:" & sIcalDateStamp & vbcrlf
		
		If oRs("message") = "" Then
			sMessage = "\n"
		Else
			sMessage = Replace(Replace(Replace(oRs("message"),Chr(44),"\,"), Chr(10),"\n"),Chr(13),"")
		End If 
		response.write "DESCRIPTION:" & sMessage & vbcrlf
		response.write "SUMMARY:" & Replace(oRs("subject"), chr(44),"\,") & vbcrlf
		response.write "PRIORITY:5" & vbcrlf
		response.write "X-MICROSOFT-CDO-IMPORTANCE:1" & vbcrlf
		response.write "CLASS:PUBLIC" & vbcrlf
		response.write "END:VEVENT" & vbcrlf
		response.flush
		oRs.MoveNext
	Loop

	response.write "END:VCALENDAR" & vbcrlf
	response.flush
End If 

oRs.Close
Set oRs = Nothing 


'------------------------------------------------------------------------------
' string generateRequestID( tmpLength )
'------------------------------------------------------------------------------
Function generateRequestID( ByVal tmpLength )

	Randomize Timer

  	Dim tmpCounter, tmpGUID
  	Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

  	For tmpCounter = 1 To tmpLength
    	tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  	Next

  	generateRequestID = tmpGUID

End Function


%>
