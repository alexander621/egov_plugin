<%
Function MyFormatDateTime( LongDate, Delimeter )
  Dim iHour, sTime

  iHour = Hour(LongDate)
  If iHour > 12 Then iHour = iHour - 12
  sTime = iHour & ":" & Right("00" & Minute(LongDate),2) & " " & Right(LongDate, 2)
  
  'check if 12 am
  If iHour = 0 then
    sTime ="12:00 AM"
  end if

  MyFormatDateTime = FormatDateTime(LongDate, vbShortDate) & Delimeter & sTime
End Function
%>