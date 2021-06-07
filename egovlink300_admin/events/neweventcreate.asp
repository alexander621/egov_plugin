<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="events_global_functions.asp" //-->
<%
Dim dDate, lDuration, sDurationInterval, lcl_customcategory, iCategoryID, lcl_isHiddenCL, lcl_pushedfrom_requestid, lcl_displayHistoryToPublic, lcl_displayHistoryOption
Dim lcl_insert_eventdate, lcl_insert_subject, lcl_insert_message, lcl_insert_calendarfeature, lcl_insert_displayHistoryOption, lcl_calendarfeature, sEventTime
Dim lcl_calendarfeatureid, lcl_feature_rssfeeds, lcl_userhaspermission_rssfeeds_events, lcl_orghasfeature_rssfeeds_events, sSql, lcl_newEventID, iTimeZoneId
Dim dDates, dNext, dStart, dEnd, sRecurCode, sHowLong, iHowMany, x, iWeeks, iDays, sDays, iDay, iMonth, iMonths, sMonthOften, iDayLike, sMonthOrdinal
Dim sSubject, sMessage, sDayOften, sYearOften, iYearPick, iMonthPick, sDayPick

sEventTime                      =  Request.Form("Hour") & ":" & Request.Form("Minute") & " " & Request.Form("AMPM")
dDate                           = CDate(Request.Form("DatePicker") & " " & sEventTime)
session("eventdate")			= Request.Form("DatePicker")  ' This is for Mason, OH
lcl_calendarfeatureid			= ""
lcl_calendarfeature				= ""
lDuration                       = -1
sDurationInterval               = request("DurationInterval")
lcl_customcategory              = ""
iCategoryID                     = 0
lcl_isHiddenCL                  = 1
lcl_pushedfrom_requestid        = "NULL"
lcl_displayHistoryToPublic      = "0"
lcl_displayHistoryOption        = ""
lcl_insert_eventdate            = "''"
lcl_insert_subject              = "''"
sSubject						= ""
lcl_insert_message              = "''"
sMessage						= ""
lcl_insert_calendarfeature      = "NULL"
lcl_insert_displayHistoryOption = "NULL"
lcl_feature_rssfeeds			= "rssfeeds_events_communitycalendar"
iTimeZoneId						= CLng(request("timezone"))

If Trim(request("cal")) <> "" Then 
	lcl_calendarfeatureid = CLng(Trim(request("cal")))
	lcl_calendarfeature   = getFeatureByID( session("orgid"), lcl_calendarfeatureid )
End If 

'Check for org features
lcl_orghasfeature_rssfeeds_events    = orghasfeature( lcl_feature_rssfeeds )

'Check for user permissions
lcl_userhaspermission_rssfeeds_events = userhaspermission( session("userid"), lcl_feature_rssfeeds )

If request("duration") <> "" Then 
	lDuration = request("Duration")
	lDuration = CLng(lDuration) * clng(sDurationInterval)
End If 

'Create a New Category for this Organization
If request("CustomCategory") <> "" Then 
	lcl_customcategory = request("CustomCategory")

	' Add the new category. The function returns the iCategoryId of the new category. This is in events_global_functions.asp
	newCategory session("orgid"), lcl_customcategory, "#000000", lcl_calendarfeature, iCategoryID
Else 
	iCategoryID = request("Category")
End If 

If request("isHiddenCL") = "on" Then 
	lcl_isHiddenCL = 0
End If 

If request("displayHistoryToPublic") = "Y" Then 
	lcl_displayHistoryToPublic = "1"
End If 

if request("displayHistoryOption") <> "" then
	lcl_displayHistoryOption = request("displayHistoryOption")
end if

'Check to see if this record is being created from a request.
if request("requestid") <> "" then
	lcl_pushedfrom_requestid = request("requestid")
end if

'Set up the fields to be inserted into the table
If dDate <> "" Then 
	lcl_insert_eventdate = dDate
	lcl_insert_eventdate = dbsafe(lcl_insert_eventdate)
	lcl_insert_eventdate = "'" & lcl_insert_eventdate & "'"
End If 

If request("subject") <> "" Then 
	lcl_insert_subject = request("subject")
	sSubject = lcl_insert_subject
	lcl_insert_subject = dbsafe(lcl_insert_subject)
	lcl_insert_subject = Left(lcl_insert_subject,50)
	lcl_insert_subject = "'" & lcl_insert_subject & "'"
End If 

If request("message") <> "" Then 
	lcl_insert_message = request("message")
	sMessage = lcl_insert_message
	lcl_insert_message = dbsafe(lcl_insert_message)
	lcl_insert_message = Left(lcl_insert_message,1500)
	
	lcl_insert_message = "'" & lcl_insert_message & "'"
End If 

If lcl_calendarfeature <> "" Then 
	lcl_insert_calendarfeature = lcl_calendarfeature
	lcl_insert_calendarfeature = dbsafe(lcl_insert_calendarfeature)
	lcl_insert_calendarfeature = "'" & lcl_insert_calendarfeature & "'"
End If 

If lcl_displayHistoryOption <> "" Then 
	lcl_insert_displayHistoryOption = lcl_displayHistoryOption
	lcl_insert_displayHistoryOption = dbsafe(lcl_insert_displayHistoryOption)
	lcl_insert_displayHistoryOption = "'" & lcl_insert_displayHistoryOption & "'"
End If 

'Create the event
sSql = "INSERT INTO Events ( "
sSql = sSql & "OrgID, "
sSql = sSql & "CreatorUserID, "
sSql = sSql & "EventDate, "
sSql = sSql & "EventTimeZoneID, "
sSql = sSql & "EventDuration, "
sSql = sSql & "[Subject], "
sSql = sSql & "[Message], "
sSql = sSql & "ModifierUserID, "
sSql = sSql & "CategoryID, "
sSql = sSql & "calendarfeature, "
sSql = sSql & "isHiddenCL, "
sSql = sSql & "pushedfrom_requestid, "
sSql = sSql & "displayHistoryToPublic, "
sSql = sSql & "displayHistoryOption "
sSql = sSql & ") VALUES ( "
sSql = sSql & session("orgid")           & ", "
sSql = sSql & session("userid")          & ", "
sSql = sSql & lcl_insert_eventdate       & ", "
sSql = sSql & iTimeZoneId		         & ", "
sSql = sSql & lDuration                  & ", "
sSql = sSql & lcl_insert_subject         & ", "
sSql = sSql & lcl_insert_message         & ", "
sSql = sSql & session("userid")          & ", "
sSql = sSql & iCategoryID                & ", "
sSql = sSql & lcl_insert_calendarfeature & ", "
sSql = sSql & lcl_isHiddenCL             & ", "
sSql = sSql & lcl_pushedfrom_requestid   & ", "
sSql = sSql & lcl_displayHistoryToPublic & ", "
sSql = sSql & lcl_insert_displayHistoryOption
sSql = sSql & " ) "

session("eventsql") = sSql
'response.write sSql & "<br /><br />"

' create the event
lcl_newEventID = RunIdentityInsertStatement( sSql )

If request("isrepeating") = "on" Then 
	' now we can do the repeating stuff here
	dNext = dDate ' this is the event date as a date type
	dStart = dNext
	dDates = ""

	' if there is an end by date, then set that here
	If IsDate(Request("endbydate")) Then 
		dEnd = CDate(Request("endbydate"))
	Else 
		dEnd = dStart
	End If 
	'response.write "dEnd = " & dEnd & "<br /><br />"

	sHowLong = Request("HowLong")
	iHowMany = Request("HowMany")
	'response.write "sHowLong = " & sHowLong & "<br /><br />"
	'response.write "iHowMany = " & iHowMany & "<br /><br />"

	sRecurCode = Request( "recur" )
	'response.write "sRecurCode = " & sRecurCode & "<br /><br />"
	'response.write "Request(""dayoften"") = " & Request("dayoften") & "<br /><br />"

	' calculate the dates that the event repeats on
	Select Case sRecurCode
		Case "dd"
			sDayOften = request("dayoften")
			'response.write "in case dd<br />"
			If sDayOften = "days" Then 
				If sHowLong = "endby" Then 
					Do While DateAdd("d",Request("Days"),dNext) < dEnd
						dNext = DateAdd("d",Request("Days"),dNext)
						If dNext > dStart Then 
							dDates = dNext & "," & dDates
						End If 
					Loop 
				ElseIf  sHowLong = "till" Then 
					For x = 1 To iHowMany
						dNext = DateAdd("d",Request("Days"),dNext)
						If dNext > dStart Then
							dDates = dNext & "," & dDates
						End If 
					Next 
				End If 
			ElseIf sDayOften = "weekdays" Then 
				If sHowLong = "till" Then 
					For x = 1 To iHowMany
						dNext = AddWeekDays( 1, dNext )
						If dNext > dStart Then
							dDates = dNext & "," & dDates
						End If 
					Next 
				ElseIf sHowLong = "endby" Then 
					Do While AddWeekDays( 1, dNext ) < dEnd
						dNext = AddWeekDays(1,dNext)
						If dNext > dStart Then
							dDates = dNext & "," & dDates
						End If 
					Loop 
				End If 
			End If 

		Case "ww"
			'response.write "in weekly<br />"
			iWeeks = Request("Weeks") ' the number of times
			sDays = Request("WeekDayNum") ' the days it happens on

			If sHowLong = "till" Then
				iDays = Split(sDays,",")
				For Each iDay In iDays
					dNext = dStart
					dNext = GetNextWeekDay( clng(iDay), dNext )
					For x = 1 To iHowMany
						dNext = DateAdd("ww",iWeeks,dNext)
						If dNext > dStart Then
							dDates = dNext & "," & dDates
						End If 
					Next
				Next
			ElseIf sHowLong = "endby" Then 
				iDays = Split(sDays,",")
				For Each iDay In iDays
					dNext = dStart
					dNext = GetNextWeekDay( clng(iDay), dNext )
					Do While DateAdd("ww",iWeeks,dNext) < dEnd
						dNext = DateAdd("ww",iWeeks,dNext)
						If dNext > dStart Then
							dDates = dNext & "," & dDates
						End If 
					Loop
				Next
			End If 

		Case "mm"
			'response.write "in monthly<br />"
			iDay = Request("monthDay")
			iMonth = Request("month")
			iMonths = Request("monthQty")
			sMonthOften = request("monthOften")
			iDayLike = Request.Form("DayLike")
			sMonthOrdinal = Request.Form("monthOrdinal")

			If sHowLong = "till" Then
				If sMonthOften = "absolute" Then
					dNext = NextDayOfMonth( clng(iDay), dNext )
					For x = 1 To clng(iHowMany)
						dNext = NextDayOfMonth( clng(iDay), dNext )
						If dNext > dStart Then
							'dDates = dNext & " " & sEventTime & "," & dDates
							dDates = dNext & "," & dDates
						End If 
						dNext = DateAdd("m",clng(iMonths),dNext)
					Next
				ElseIf sMonthOften = "relative" Then
					For x = 1 To clng(iHowMany)
						dNext = OrdinalDate( iDayLike, sMonthOrdinal, dNext )
						If dNext > dStart Then
							dDates = dNext & " " & sEventTime & "," & dDates
						End If 
						dNext = DateAdd("m",clng(iMonth),dNext)
					Next
				End If
			ElseIf sHowLong = "endby" Then 
				If sMonthOften = "absolute" Then
					Do While NextDayOfMonth(clng(iDay),dNext) < dEnd
						dNext = NextDayOfMonth(clng(iDay),dNext)
						If dNext > dStart Then
							'dDates = dNext & " " & sEventTime & "," & dDates
							dDates = dNext & "," & dDates
						End if
						dNext = DateAdd("m",iMonths,dNext)
					Loop
				ElseIf sMonthOften = "relative" Then
					Do While OrdinalDate( iDayLike, sMonthOrdinal, dNext ) < dEnd
						dNext = OrdinalDate( iDayLike, sMonthOrdinal, dNext )
						If dNext > dStart Then
							dDates = dNext & " " & sEventTime & "," & dDates
						End If 
						dNext = DateAdd("m",iMonth,dNext)
					Loop
				End If
			End If 

		Case "yy"
			'response.write "in yearly<br />"
			sYearOften = request("yearOften")

			If sHowLong = "till" Then
				Select Case sYearOften
					Case "every"
						iMonth = request("yearMonth")
						iDay = request("yearDay")
						dNext = NextDate( clng(iMonth), clng(iDay), dNext )
						For x = 1 To clng(iHowMany)
							If dNext > dStart Then
								dDates = dNext & " " & sEventTime & "," & dDates
							End If  
							dNext = DateAdd("yyyy",1,dNext)
						Next
					Case "absolute"
						iMonths = request("yearMonths")
						iDay = request("yearDayNum")
						dNext = NextDayOfMonth( clng(iDay), dNext )
						Do While DatePart("yyyy",NextDayOfMonth( clng(iDay), dNext))-DatePart("yyyy",dStart) <= clng(iHowMany)
							dNext = NextDayOfMonth(clng(iDay),dNext)
							If dNext > dStart Then
								dDates = dNext & "," & dDates
							End If 
							dNext = DateAdd("m",clng(iMonths),dNext)
						Loop
					Case "relative"
						iYearPick = request("yearOrdinal")
						iMonthPick = request("yearMonthPick")
						sDayPick = request("yearDayPick")
						dNext = NextDate( iMonthPick,1,dNext )
						For x = 1 To clng(iHowMany)
							dNext = OrdinalDate( sDayPick, iYearPick, dNext )
							If dNext > dStart Then
								dDates = dNext & " " & sEventTime & "," & dDates
							End If 
							dNext = DateAdd("yyyy",1,dNext)
						Next
				End Select
			ElseIf sHowLong = "endby" Then
				Select Case sYearOften
					Case "every"
						iMonth = request("yearMonth")
						iDay = request("yearDay")
						dNext = NextDate(clng(iMonth),clng(iDay),dNext)
						Do While dNext < dEnd
							If dNext > dStart Then
								dDates = dNext & " " & sEventTime & "," & dDates
							End If 
							dNext = DateAdd("yyyy",1,dNext)
						Loop
					Case "absolute"
						iMonths = request("yearMonths")
						iDay = request("yearDayNum")
						Do While NextDayOfMonth(clng(iDay),dNext) < dEnd
							dNext = NextDayOfMonth(clng(iDay),dNext)
							If dNext > dStart Then
								dDates = dNext & "," & dDates
							End If 
							dNext = DateAdd("m",clng(iMonths),dNext)
						Loop
					Case "relative"
						iYearPick = request("yearOrdinal")
						iMonthPick = request("yearMonthPick")
						sDayPick = request("yearDayPick")
						dNext = NextDate(iMonthPick,1,dNext)
						Do While OrdinalDate(sDayPick,iYearPick,dNext) < dEnd
							dNext = OrdinalDate(sDayPick,iYearPick,dNext)
							If dNext > dStart Then
								dDates = dNext & " " & sEventTime & "," & dDates
							End If 
							dNext = DateAdd("yyyy",1,dNext)
						Loop
				End Select
			End If 
	End Select

	' create the repeating events
	RecurEvent lcl_newEventID, dDates, lcl_isHiddenCL, dDate, iTimeZoneId, lDuration, sSubject, sMessage, iCategoryID, lcl_calendarfeature

End If 

lcl_return_parameters = ""

If lcl_orghasfeature_rssfeeds_events And lcl_userhaspermission_rssfeeds_events And request("sendTo_RSS") = "on" Then 
	lcl_return_parameters = "&sendTo_RSS=" & lcl_newEventID
End If 

'response.write "Done"
response.redirect "default.asp?success=SA&useSessions=Y&cal=" & lcl_calendarfeatureid & lcl_return_parameters




'------------------------------------------------------------------------------
'Function dbsafe( ByVal p_value )
'	Dim lcl_return
'
'	lcl_return = ""
'	lcl_return = Replace(p_value,"'","''")
'	dbsafe = lcl_return
'
'End Function 


'------------------------------------------------------------------------------
Sub RecurEvent( ByVal iEventID, ByVal dRepeatingDates, ByVal p_isHiddenCL, ByVal dEventDate, ByVal iTimeZoneId, ByVal iDuration, ByVal sSubject, ByVal sMessage, ByVal iCategoryID, ByVal sCalendarFeature )
	Dim dDate, dDates, lcl_identity

	response.write "dRepeatingDates = " & dRepeatingDates & "<br /><br />"

	dDates = Split(dRepeatingDates,",")

	For Each dDate In dDates 
		If dDate <> "" Then 
			If DateDiff("d", dDate, dEventDate) <> 0 Then 
				'dDate = dDate '& " " & FormatDateTime(dEventDate,vbLongTime)

				'Create the New Recurring Event. This is in events_global_functions.asp
				newRecurEvent session("orgid"), iEventID, dDate, session("userid"), iTimeZoneId, iDuration, sSubject, sMessage, iCategoryID, sCalendarFeature, p_isHiddenCL, lcl_identity
			End If 
		End If 
	Next 

End Sub 


'------------------------------------------------------------------------------
Function AddWeekDays( ByVal nDaystoAdd, ByVal dtStartDate )
	'Adds working days based on a five day week

	Dim dtEndDate, iLoop

	'First add whole weeks
	dtEndDate = DateAdd("ww",Int(nDaysToAdd/5),dtStartDate)

	'Add any odd days
	For iLoop = 1 To (nDaysToAdd Mod 5)
		dtEndDate = DateAdd("d",1,dtEndDate)
		'If Saturday increment to following Monday
		If WeekDay(dtEndDate) = vbSaturday Then
			'Increment date to following Monday
			dtEndDate = DateAdd("d",2,dtEndDate)
		End If
	Next

	AddWeekDays = dtEndDate

End Function


'------------------------------------------------------------------------------
Function GetNextWeekDay( ByVal iWeekDay, ByVal dTemp)
	'Increments a date until the next specified weekday

	'Add any odd days
	Do While Not DatePart("w",dTemp) = iWeekDay
		dTemp = DateAdd("d",1,dTemp)
	Loop

	GetNextWeekDay = dTemp

End Function


'------------------------------------------------------------------------------
Function GetAWeekday( ByVal iNumDays, ByVal iWeekDay, ByVal dStart )
	'Increments a date until the next specified weekday
	Dim x, dEnd

	dEnd = dStart

	For x = 1 To iNumDays
		'Add any odd days
		Do While Not DatePart("w",dEnd) = iWeekDay
			dEnd = DateAdd("d",1,dEnd)
		Loop

		If x <> iNumDays Then
			dEnd = DateAdd("d",1,dEnd)
		End If
	Next

	GetAWeekday = dEnd

End Function


'------------------------------------------------------------------------------
Function GetLastWeekday( ByVal dStart )
	'Increments a date until the next specified weekday
	Dim dEnd

	dEnd = dStart
	dEnd = GetLastDayInMonth( dEnd )


	'Add any odd days
	Do While Weekend(DatePart("w",dEnd))
		dEnd = DateAdd("d",-1,dEnd)
	Loop

	GetLastWeekDay = dEnd

End Function


'------------------------------------------------------------------------------
Function GetLastWeekend( ByVal dStart )
	'Increments a date until the next specified weekday
	Dim dEnd

	dEnd = dStart
	dEnd = GetLastDayInMonth( dEnd )

	'Add any odd days
	Do While Not Weekend( dEnd )
		dEnd = DateAdd("d",-1,dEnd)
	Loop

	GetLastWeekend = dEnd

End Function


'------------------------------------------------------------------------------
Function GetALastWeekday( ByVal iDay, ByVal dStart )
	'Increments a date until the next specified weekday
	Dim dEnd

	dEnd = dStart
	dEnd = GetLastDayInMonth( dEnd )

	'Add any odd days
	Do While Not DatePart("w",dEnd) = iDay
		dEnd = DateAdd("d",-1,dEnd)
	Loop

	GetALastWeekday = dEnd

End Function


'------------------------------------------------------------------------------
Function GetAnyWeekday( ByVal iNumDays, ByVal dStart )
	'Increments a date until the next weekday
	Dim x, dEnd

	dEnd = CDate(dStart)

	'Add any odd days
	For x = 1 To iNumDays
		Do While Weekend( dEnd )
			dEnd = DateAdd("d",1,dEnd)
		Loop
		
		If x <> iNumDays Then
			dEnd = DateAdd("d",1,dEnd)
		End If
	Next

	GetAnyWeekDay = dEnd

End Function


'------------------------------------------------------------------------------
Function NextDate( ByVal iMonth, ByVal iDay, ByVal dEnd )
	'Increments a date until the next weekday
	Dim dTemp

	If GetDaysInMonth( dEnd ) < iDay Then 
		iDay = GetDaysInMonth( dEnd )
	End If 

	dTemp = DateSerial(DatePart("yyyy",dEnd),iMonth,iDay)
	If dTemp < dEnd Then
		dTemp = DateAdd("yyyy",1,dTemp)
	End If

	NextDate = dTemp

End Function


'------------------------------------------------------------------------------
Function GetAnyWeekend( ByVal iNumDays, ByVal dStart )
	'Increments a date until the next weekend
	Dim x, dEnd

	dEnd = dStart

	For x = 1 To iNumDays
		'Add any odd days
		Do While Not Weekend( dEnd )
			dEnd = DateAdd("d",1,dEnd)
		Loop

		If x <> iNumDays Then
			dEnd = DateAdd("d",1,dEnd)
		End If
	Next

	GetAnyWeekend = dEnd

End Function


'------------------------------------------------------------------------------
Function GetDaysInMonth( ByVal dDate )
	Dim dTemp, iYear, iMonth

	iYear = DatePart("yyyy",dDate)
	iMonth = DatePart("m",dDate)
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))

	GetDaysInMonth = Day(dTemp)

End Function


'------------------------------------------------------------------------------
Function GetLastDayInMonth( ByVal dDate )
	Dim dTemp, iYear, iMonth

	iYear = DatePart("yyyy",dDate)
	iMonth = DatePart("m",dDate)
	dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))

	GetLastDayInMonth = dTemp

End Function


'------------------------------------------------------------------------------
Function FirstDayInMonth( ByVal dDate )
	Dim iYear, iMonth, dTemp

	iYear = DatePart("yyyy",dDate)
	iMonth = DatePart("m",dDate)
	dTemp = DateSerial(iYear, iMonth, 1)

	FirstDayInMonth = CDate(dTemp)

End Function


'------------------------------------------------------------------------------
Function Weekend( ByVal dDate )

	If WeekDay( dDate ) = VBSaturday or WeekDay( dDate ) = VBSunday Then
		Weekend = True
	Else
		Weekend = False
	End If

End Function


'------------------------------------------------------------------------------
Function NextDayOfMonth( ByVal iDay, ByVal dDate)

	If iDay <= GetDaysInMonth( dDate ) Then
		Do While (Not DatePart("d",dDate) = iDay)
			dDate = DateAdd("d", 1, dDate)
		Loop
	Else
		dDate = GetLastDayInMonth( dDate )
	End If

	NextDayOfMonth = dDate

End Function


'------------------------------------------------------------------------------
Function OrdinalDate( ByVal tType, ByVal iOrd, ByVal dDate )

	Select Case tType
		Case "d"   'Day
			Select Case clng(iOrd)
				Case 1
				  dDate = FirstDayInMonth(dDate)
				Case 2
				  dDate = DateAdd("d",1,FirstDayInMonth(dDate))
				Case 3
				  dDate = DateAdd("d",2,FirstDayInMonth(dDate))
				Case 4
				  dDate = DateAdd("d",3,FirstDayInMonth(dDate))
				Case 5
				  dDate = GetLastDayInMonth(dDate)
			End Select

	  	Case "wd"  'Weekday
			Select Case iOrd
				Case 1
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekday(1,dDate)
				Case 2
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekday(2,dDate)
				Case 3
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekday(3,dDate)
				Case 4
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekday(4,dDate)
				Case 5
				  dDate = GetLastWeekDay(dDate)
			End Select

	  	Case "wed" 'Weekend day
			Select Case iOrd
				Case 1
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekend(1,dDate)
				Case 2
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekend(2,dDate)
				Case 3
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekend(3,dDate)
				Case 4
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAnyWeekend(4,dDate)
				Case 5
				  dDate = GetLastWeekend(dDate)
			End Select

	    Case "1","2","3","4","5","6","7"
			Select Case iOrd
				Case 1
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAWeekday(1,clng(tType),dDate)
				Case 2
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAWeekday(2,clng(tType),dDate)
				Case 3
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAWeekday(3,clng(tType),dDate)
				Case 4
				  dDate = FirstDayInMonth(dDate)
				  dDate = GetAWeekday(4,clng(tType),dDate)
				Case 5
				  dDate = GetALastWeekday(clng(tType),dDate)
			End Select
    End Select

	OrdinalDate = dDate

End Function




%>
