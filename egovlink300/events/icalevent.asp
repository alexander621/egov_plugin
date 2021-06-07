<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: icalevent.asp
' AUTHOR: SteveLoar
' CREATED: 12/10/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: This pulls a calendar event into an iCal format file
'
' MODIFICATION HISTORY
' 1.0   12/10/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, dEventDate, sUTCEventDate, iEventDuration, sEndDate, dEndDate, iEventId, sIcalDateStamp
Dim dIcalDateStamp, sMessage, sSubject

If Trim(request("e")) <> "" Then 
	If Not IsNumeric(Trim(request("e"))) Then 
		response.End 
    Else
		iEventId = CLng(request("e"))
	End If 
Else
	response.End
End If 

' SET UP PAGE OPTIONS
server.scripttimeout = 9000
Response.ContentType = "text/calendar"
Response.AddHeader "Content-Disposition", "attachment;filename=event_" & iEventId & ".ics"

' Get and output the data
sSql = "SELECT eventdate, eventduration, dbo.GetUTCTime( " & iorgid & ", eventdate ) AS utctime, "
sSql = sSql & "dbo.GetUTCTime( " & iorgid & ", getdate() ) AS icaldatestamp, "
sSql = sSql & "subject, message "
sSql = sSql & "FROM events WHERE eventid = " & iEventId
'response.write sSql

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	response.write "BEGIN:VCALENDAR" & vbcrlf
	response.write "VERSION:2.0" & vbcrlf
	response.write "PRODID:-//EGOVLINK//EN" & vbcrlf
	'response.write "CALSCALE:GREGORIAN" & vbcrlf
	response.write "METHOD:PUBLISH" & vbcrlf
	response.flush

	response.write vbcrlf & "BEGIN:VEVENT" & vbcrlf
	dEventDate = CDate(oRs("utctime"))
	sUTCEventDate = Year(dEventDate) & Right("0" & Month(dEventDate),2) & Right("0" & Day(dEventDate),2) & "T" & Right("0" & Hour(dEventDate),2) & Right("0" & Minute(dEventDate),2) & "00Z"
	iEventDuration = CLng(oRs("eventduration"))
	If iEventDuration > CLng(0) Then
		dEndDate = DateAdd("n", iEventDuration, dEventDate)
		sEndDate = Year(dEndDate) & Right("0" & Month(dEndDate),2) & Right("0" & Day(dEndDate),2) & "T" & Right("0" & Hour(dEndDate),2) & Right("0" & Minute(dEndDate),2) & "00Z"
	Else
		sEndDate = sUTCEventDate
	End If 
	dIcalDateStamp = CDate(oRs("icaldatestamp"))
	sIcalDateStamp = Year(dIcalDateStamp) & Right("0" & Month(dIcalDateStamp),2) & Right("0" & Day(dIcalDateStamp),2) & "T" & Right("0" & Hour(dEventDate),2) & Right("0" & Minute(dIcalDateStamp),2) & "00Z"
	response.write "DTSTART:" & sUTCEventDate & vbcrlf
	response.write "DTEND:" & sEndDate & vbcrlf
	response.write "TRANSP:OPAQUE" & vbcrlf
	response.write "SEQUENCE:0" & vbcrlf
	'response.write "UID:040000008200E00074C5B7101A82E00800000000307BB6556840CB01000000000000000010000000A3335DB70FEACD439C9E5455BBDC552D" & vbcrlf
	response.write "UID:" & generateRequestID( 200) & "@eclink.com" & vbcrlf
	response.write "DTSTAMP:" & sIcalDateStamp & vbcrlf
	If oRs("message") = "" Then
		sMessage = "\n"
	Else
		sMessage = Replace(Replace(oRs("message"), Chr(10),"\n"),Chr(13),"")
		sMessage = Replace(sMessage, "<br />", "\n" )
	End If 
	response.write "DESCRIPTION:" & sMessage & vbcrlf
	
	sSubject = oRs("subject")
	sSubject = Replace(sSubject, "<br />", "" )
	response.write "SUMMARY:" & sSubject & vbcrlf

	response.write "PRIORITY:5" & vbcrlf
	response.write "X-MICROSOFT-CDO-IMPORTANCE:1" & vbcrlf
	response.write "CLASS:PUBLIC" & vbcrlf
	response.write "END:VEVENT" & vbcrlf

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
