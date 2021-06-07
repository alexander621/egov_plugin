<%

Function GetFOILDueDate( ByRef sDate)
	for x = 0 to 19 'sDate + 20 "Business Days"
       		sDate = FindNextRyeBusinessDay(DateAdd("d",1,sDate))
	next

	GetFOILDueDate = sDate
End Function

Function FindNextRyeBusinessDay(sDate)


	
	if (Weekday(sDate) = 1 or Weekday(sDate) = 7) _ 
		or (Day(sDate) = 1 and Month(sDate) = 1) _ 
		or (DateSerial(Year(sDate), Month(sDate), Day(sDate)) = NthXday("2/1/" & Year(sDate),3,2))  _ 
		or sDate & " " = LastMonday("5/31/" & Year(sDate)) & " " _
		or (Day(sDate) = 4 and Month(sDate) = 7) _ 
		or (DateSerial(Year(sDate), Month(sDate), Day(sDate)) = NthXday("9/1/" & Year(sDate),1,2))  _ 
		or (DateSerial(Year(sDate), Month(sDate), Day(sDate)) = NthXday("11/1/" & Year(sDate),4,5))  _ 
		or (Day(sDate) = 25 and Month(sDate) = 12) _ 
		 then
		sDate = FindNextRyeBusinessDay(DateAdd("d",1,sDate))
	end if

	FindNextRyeBusinessDay = sDate

End Function

Function LastMonday(sDate)
	Do While Not WeekDay(sDate) = 2
		sDate = DateAdd("d",-1,sDate)
	loop
	LastMonday = sDate
End Function

Function NthXDay(sDate,N,xDay)
nYear = Year(sDate)
nMonth = Month(sDate)


' Get the first of the month as a date
yourDate = DateSerial(nYear, nMonth, 1)

' Let us say first of the week is a Sunday (as in the US)
firstOfWeek = 1

' Find out what day is that
nDay = WeekDay(yourDate, firstOfWeek)

' First "xDay" will be as follows
' See if the xDay is beyond the first of the month
' Else you have to go the xDay of next week
If (xDay - nDay) >= 0 Then
NthXDay = DateSerial(nYear, nMonth, 1 + (xDay - nDay) + (N-1) * 7)
Else
NthXDay = DateSerial(nYear, nMonth, 1 + (xDay - nDay) + (N * 7))
End If
End Function
%>
