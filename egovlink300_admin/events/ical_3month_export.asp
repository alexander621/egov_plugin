<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: ical_3month_export.asp
' AUTHOR: SteveLoar
' CREATED: 11/29/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: This pulls a static 3 month of calendar events into an iCal format file
'
' MODIFICATION HISTORY
' 1.0   11/29/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, dEventDate, sUTCEventDate, iEventDuration, sEndDate, dEndDate, sMessage
Dim sIcalDateStamp, dIcalDateStamp

' SET UP PAGE OPTIONS
server.scripttimeout = 9000
Response.ContentType = "text/calendar"
Response.AddHeader "Content-Disposition", "attachment;filename=events.ics"

' Get and output the data
sSql = "SELECT eventdate, eventduration, dbo.GetUTCTime( " & session("orgid") & ", eventdate ) AS utctime, "
sSql = sSql & "dbo.GetUTCTime( " & session("orgid") & ", getdate() ) AS icaldatestamp, "
sSql = sSql & "subject, message "
sSql = sSql & "FROM events WHERE orgid = " & session("orgid") & " AND (calendarfeature = '' OR calendarfeature IS NULL) "
sSql = sSql & " AND eventdate > '" & DateValue(Date()) & "' AND eventdate < '" & DateValue(DateAdd("m", 3, Date())) & "' "
sSql = sSql & "ORDER BY eventdate"
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	response.write "BEGIN:VCALENDAR" & vbcrlf
	response.write "VERSION:2.0" & vbcrlf
	response.write "PRODID:-//EGOVLINK//EN" & vbcrlf
	response.write "METHOD:PUBLISH" & vbcrlf
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
			sMessage = Replace(Replace(oRs("message"), Chr(10),"\n"),Chr(13),"")
		End If 
		response.write "DESCRIPTION:" & sMessage & vbcrlf
		response.write "SUMMARY:" & oRs("subject") & vbcrlf
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
