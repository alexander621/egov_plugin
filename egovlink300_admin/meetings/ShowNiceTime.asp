<%
Function ShowNiceTime(objTime)
Dim sDate, sHour, sMin, sAmPm
	sDate = FormatDateTime(oRst("MeetingTime"),VBShortDate)
	sHour = Hour(oRst("MeetingTime")) 
	sMin = Minute(oRst("MeetingTime"))
	If sMin < 10 then sMin = Right("00" & sMin, 2)
	If sHour > 11 then
		If sHour > 12 then sHour = sHour - 12
		sAmPm = langPM
	Else
		sAmPm = langAM
	End If
	ShowNiceTime = sDate & " at " & sHour & ":" & sMin & " " &  sAmPm
End Function
%>