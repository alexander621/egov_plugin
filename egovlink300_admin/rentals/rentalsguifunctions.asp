<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalsguifunctions.asp
' AUTHOR: Steve Loar
' CREATED: 08/21/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared gui functions for rentals. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   08/21/2009   Steve Loar - INITIAL VERSION
' 1.1	03/24/2011	Steve Loar - hide deactivated rentals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' voidDrawDateChoices sName
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""16"">Today</option>"
	response.write vbcrlf & "<option value=""17"">Yesterday</option>"
	response.write vbcrlf & "<option value=""18"">Tomorrow</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""14"">Next Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""13"">Next Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""15"">Next Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""19"">This Year</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""20"">Next Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Today</option>"
	response.write vbcrlf & "<option value=""21"">Today through Next Year</option>"
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowAccountPicks sSelectName, iAccountNo, bShowAllPick
'--------------------------------------------------------------------------------------------------
Sub ShowAccountPicks( ByVal sSelectName, ByVal iAccountNo, ByVal bShowAllPick )
	Dim sSql, oRs

	sSql = "SELECT accountid, accountname FROM egov_accounts WHERE orgid = " & session("orgid")
	sSql = sSql & " AND accountstatus = 'A' ORDER BY accountname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	If bShowAllPick Then 
		response.write "<option value=""0"">Include All GL Accounts</option>"
	End If 
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("accountid") & """"

		If iAccountNo <> "" Then 
   			If CLng(oRS("accountid")) = CLng(iAccountNo) Then
				response.write " selected=""selected"" "
   			End If
		End If 

		response.write ">" & oRs("accountname") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowAdminPicks iRentalUserid, sSearchName
'--------------------------------------------------------------------------------------------------
Sub ShowAdminPicks( ByVal iRentalUserid, ByVal sName )
	Dim sSql, oRs, sSearchName, sAdminIncl

	sSearchName = dbsafe(sName)

	If UserIsRootAdmin( Session("UserId") ) Then
		sAdminIncl = ""
	Else
		sAdminIncl = " AND isrootadmin <> 1 "
	End If 

	sSql = "SELECT 1 AS foo, userid, firstname, lastname, ISNULL(email,'') AS email, "
	sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname "
	sSql = sSql & " FROM users WHERE orgid = " & session("orgid")
	sSql = sSql & " AND lastname LIKE '" & sSearchName & "%' " & sAdminIncl
	sSql = sSql & " UNION "
	sSql = sSql & " SELECT 2 AS foo, userid, firstname, lastname, ISNULL(email,'') AS email, "
	sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname "
	sSql = sSql & " FROM users WHERE orgid = " & session("orgid")
	sSql = sSql & " AND ( firstname LIKE '%" & sSearchName & "%' OR lastname LIKE '%" & sSearchName & "%' ) " & sAdminIncl
	sSql = sSql & " AND userid NOT IN ( SELECT userid FROM users WHERE orgid = " & session("orgid")
	sSql = sSql & " AND lastname LIKE '" & sSearchName & "%' )"
	sSql = sSql & " ORDER BY foo, sortname, email, userid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write "Select a Name: <select name='rentaluserid' id='rentaluserid'>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value='" & oRs("userid") & "'"
			If iRentalUserid = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			response.write oRs("lastname") & ", " & oRs("firstname")
			If oRs("email") <> "" Then
				response.write " - " & oRs("email")
			End If 
			response.write "</option>"
			oRs.MoveNext 
		Loop
		response.write vbcrlf & "</select>"
	Else
		response.write "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />No Matching Names Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowAmPmPicks sSelectName, sAmPm, sDisabledOption
'--------------------------------------------------------------------------------------------------
Sub ShowAmPmPicks( ByVal sSelectName, ByVal sAmPm, ByVal sDisabledOption )

	If sDisabledOption = "disabled" Then
		response.write "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & sAmPm & """ />"
		response.write sAmPm
	Else 
		response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
		response.write vbcrlf & "<option value=""AM""" 
		If sAmPm = "AM" Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">AM</option>"
		response.write vbcrlf & "<option value=""PM"""
		If sAmPm = "PM" Then 
			response.write " selected=""selected"" "
		End If
		response.write ">PM</option>"
		response.write vbcrlf & "</select>"
	End If 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowCitizenPicks iRentalUserid, sSearchName
'--------------------------------------------------------------------------------------------------
Sub ShowCitizenPicks( ByVal iRentalUserid, ByVal sName )
	Dim sSearchName, sSql, oRs

	sSearchName = dbsafe(sName)

	' This query finds last names that start with the search ahead of those that match anywhere
	sSql = "SELECT 1 AS foo, userid AS userid, userfname AS firstname, userlname AS lastname, "
	sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
	sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
	sSql = sSql & " AND userlname LIKE '" & sSearchName & "%' "
	sSql = sSql & " UNION "
	sSql = sSql & " SELECT 2 AS foo, userid, userfname AS firstname, userlname AS lastname, "
	sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
	sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
	sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' ) "
	sSql = sSql & " AND userid NOT IN ( SELECT userid FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 "
	sSql = sSql & " AND headofhousehold = 1 AND userregistered = 1 AND userlname LIKE '" & sSearchName & "%' ) "
	sSql = sSql & " ORDER BY foo, sortname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRS.EOF Then
		response.write "Select a Name: <select name='rentaluserid' id='rentaluserid'>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value='" & oRs("userid") & "'"
			If iRentalUserid = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			response.write oRs("lastname") & ", " & oRs("firstname")
			If oRs("address") <> "" Then
				response.write " - " & oRs("address")
			End If 

			response.write "</option>"

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />No Matching Names Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowClassActivityNoAndName iTimeId, bShowEdit
'--------------------------------------------------------------------------------------------------
Sub ShowClassActivityNoAndName( ByVal iTimeId, ByVal bShowEdit )
	Dim oRs, sSql 

	sSql = "SELECT C.classid, C.classname, T.activityno "
	sSql = sSql & " FROM egov_class C, egov_class_time T "
	sSql = sSql & " WHERE C.classid = T.classid and T.timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write oRs("activityno") & " &nbsp; " & oRs("classname")
		If bShowEdit Then 
			response.write " &nbsp; <input type=""button"" class=""button"" value=""Edit Class"" onclick=""location.href='../classes/edit_class.asp?classid=" & oRs("classid") & "';"" />"
			response.write " &nbsp; <input type=""button"" class=""button"" value=""View Roster"" onclick=""location.href='../classes/view_roster.asp?classid=" & oRs("classid") & "&timeid=" & iTimeId & "';"" />"
		End If 
	Else
		response.write ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowDOWPicks sSelectName, iDOW
'--------------------------------------------------------------------------------------------------
Sub ShowDOWPicks( ByVal sSelectName, ByVal iDOW )
	Dim x

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	For x = 1 To 7
		response.write vbcrlf & "<option value=""" & x & """"
		If clng(x) = clng(iDOW) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & WeekDayName(x) & "</option>"
	Next 
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowHourPicks sSelectName, iHour, sDisabledOption
'--------------------------------------------------------------------------------------------------
Sub ShowHourPicks( ByVal sSelectName, ByVal iHour, ByVal sDisabledOption )
	Dim x

	If sDisabledOption = "disabled" Then
		response.write "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & iHour & """ />"
		response.write iHour
	Else 
		response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
		For x = 1 To 12
			response.write vbcrlf & "<option value=""" & x & """"
			If clng(x) = clng(iHour) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & x & "</option>"
		Next 
		response.write vbcrlf & "</select>"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowLocationPicks iLocationId, bIncludeAllPick
'--------------------------------------------------------------------------------------------------
Sub ShowLocationPicks( ByVal iLocationId, ByVal bIncludeAllPick )
	Dim sSql, oRs

	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""locationid"">"

	If bIncludeAllPick Then
		response.write vbcrlf & vbtab & "<option value=""0"">All Locations</option>"
	End If 
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("locationid") & """ "
			If clng(oRs("locationid")) = clng(iLocationId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowMinutePicks sSelectName, iMinute, sDisabledOption
'--------------------------------------------------------------------------------------------------
Sub ShowMinutePicks( ByVal sSelectName, ByVal iMinute, ByVal sDisabledOption )
	Dim x, sMinutePrefix

	If sDisabledOption = "disabled" Then
		If clng(iMinute) < clng(10) Then
			sMinutePrefix = "0"
		Else 
			sMinutePrefix = ""
		End If
		response.write Right(sMinutePrefix & iMinute,2)
		response.write "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & sMinutePrefix & iMinute & """ />"
	Else 
		response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
		For x = 0 To 59
			If clng(x) < clng(10) Then
				sMinutePrefix = "0"
			Else 
				sMinutePrefix = ""
			End If
			response.write vbcrlf & "<option value="""
			response.write sMinutePrefix & x & """"
			If clng(x) = clng(iMinute) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			response.write sMinutePrefix & x & "</option>"
		Next 
		response.write vbcrlf & "</select>"
	End If 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowMonthlyPeriodPicks iMonthlyPeriod
'--------------------------------------------------------------------------------------------------
Sub ShowMonthlyPeriodPicks( iMonthlyPeriod )
	
	response.write "<select name='monthlyperiodid' id='monthlyperiodid'>"

	response.write vbcrlf & "<option value='1'"
	If iMonthlyPeriod = CLng(1) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">First</option>"

	response.write vbcrlf & "<option value='2'"
	If iMonthlyPeriod = CLng(2) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Second</option>"

	response.write vbcrlf & "<option value='3'"
	If iMonthlyPeriod = CLng(3) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Third</option>"

	response.write vbcrlf & "<option value='4'"
	If iMonthlyPeriod = CLng(4) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Fourth</option>"

	response.write vbcrlf & "<option value='5'"
	If iMonthlyPeriod = CLng(5) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Last</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOtherReservationsForDate iRentalId, dDate
'--------------------------------------------------------------------------------------------------
Sub ShowOtherReservationsForDate( ByVal iRentalId, ByVal dStartDateTime )
	Dim oRs, sSql, dWantedEndTime, sFirstName, sLastName, sPhone

	' This gets anything that starts anytime on the passed date
	dStartDateTime = dStartDateTime & " 0:00 AM" ' Add the time of midnight to the passed in date
	dWantedEndTime = DateAdd("d", 1, CDate(dStartDateTime)) ' set this to midnight of the next day

	sSql = "SELECT D.reservationid, D.reservationstarttime, D.billingendtime, D.reservationendtime, T.reservationtype, R.timeid, T.reservationtypeselector, "
	sSql = sSql & " ISNULL(R.rentaluserid,0) AS rentaluserid, ISNULL(R.adminuserid,'') AS adminuserid, T.isreservation, T.isclass "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservations R, egov_rentalreservationtypes T "
	sSql = sSql & " WHERE D.reservationid = R.reservationid AND R.reservationtypeid = T.reservationtypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND D.reservationstarttime BETWEEN '" & dStartDateTime & "' AND '" & dWantedEndTime & "' AND R.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY D.reservationstarttime"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' GetTimeFormat is in common.asp
		response.write "<br /><a href=""reservationedit.asp?reservationid=" & oRs("reservationid") & """ target=""_blank"">"
		response.write GetTimeFormat(oRs("reservationstarttime")) & " to " & GetTimeFormat(oRs("billingendtime")) & " &ndash; "
		response.write oRs("reservationtype") & " &ndash; "
		If oRs("isreservation") Then
			If oRs("reservationtypeselector") = "public" Then 
				' Show the citizen name
				ShowShortCitizenName oRs("rentaluserid")
			Else
				response.write GetAdminName(oRs("rentaluserid"))
			End If 
		Else
			If oRs("isclass") Then
				' Show the class name
				ShowShortClassName oRs("timeid")
			Else 
				' the admin who made the hold or block or whatever - in rentlascommonfunctions.asp
				GetAdminNameAndPhone oRs("adminuserid"), sFirstName, sLastName, sPhone
				If sLastName <> "" And sFirstName <> "" Then 
					response.write Left(UCase(Left(sFirstName,1)) & ". " & sLastName,30)
				End If 
			End If 
		End If 
		response.write "</a>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowPaymentLocations
'------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oRs

	sSql = "SELECT paymentlocationid, paymentlocationname FROM egov_paymentlocations "
	sSql = sSql & " WHERE isadminmethod = 1 ORDER BY paymentlocationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""paymentlocationid"" name=""paymentlocationid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("paymentlocationid") & """>" & oRs("paymentlocationname") & "</option>"
		oRs.movenext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRecreationCategories iRecreationCategoryId
'--------------------------------------------------------------------------------------------------
Sub ShowRecreationCategories( ByVal iRecreationCategoryId )
	Dim oRs, sSql

	sSql = "SELECT recreationcategoryid, categorytitle FROM egov_recreation_categories "
	sSql = sSql & "WHERE isforrentals = 1 AND orgid = " & session("orgid") & " ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select name=""recreationcategoryid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("recreationcategoryid") & """ "
		If CLng(oRs("recreationcategoryid")) = CLng(iRecreationCategoryId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("categorytitle") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowRentalAvailabilityFlag iRentalid, aWantedDates, sPeriodType, bIsForClass
'--------------------------------------------------------------------------------------------------
Sub ShowRentalAvailabilityFlag( ByVal iRentalid, ByRef aWantedDates, ByVal sPeriodType, ByVal bIsForClass )
	Dim xbOffSeasonFlag, sFlag, dWantedStartTime, dWantedEndTime

	sFlag = "Yes"

	' Check each day
	For x = 0 To UBound(aWantedDates,2)
		' Set the start and end times
		'dWantedStartTime = CDate(aWantedDates(0,x))
		'dWantedEndTime = CDate(aWantedDates(1,x))
		sTestStart = aWantedDates(0,x)
		if not IsDate(sTestStart) then sTestStart = "1/1/1900"
		sTestEnd = aWantedDates(1,x)
		if not IsDate(sTestEnd) then sTestEnd = "1/1/1900"
		dWantedStartTime = CDate(sTestStart)
		dWantedEndTime = CDate(sTestEnd)

		' Is this date in season or off season
		bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(dWantedStartTime) )
		'response.write bOffSeasonFlag

		' Check if the rental is only available all day on this DOW and if so then adjust the times
		If RentalIsAllDay( iRentalid, bOffSeasonFlag, Weekday(DateValue(dWantedStartTime)) ) Then 
			GetAllDayHours iRentalid, bOffSeasonFlag, Weekday(DateValue(dWantedStartTime)), dWantedStartTime, dWantedEndTime
		End If 

		' Now check that the wanted date fits into the hours of the rental itself
		sFlag = CheckRentalHours( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType, bOffSeasonFlag )

		If UCase(sFlag) = "YES" Then
			' Finally check if anyone else has a conflicting reservation - buffer time is now optional so do not check it
			sFlag = CheckForExistingReservations( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType, bOffSeasonFlag, bIsForClass, False )
		End If 

		If UCase(sFlag) = "NO" Then
			' When we hit a no We are done with looking
			Exit For 
		End If 
	Next 
	response.write sFlag

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalLocationPicks iRentalId, bIncludeAllPick, bIncludeLocationAllPicks
'--------------------------------------------------------------------------------------------------
Sub ShowRentalLocationPicks( ByVal iRentalId, ByVal bIncludeAllPicks, ByVal bIncludeLocationAllPicks )
	Dim sSql, oRs, iLocationId

	iLocationId = CLng(0)

	sSql = "SELECT L.locationid , R.rentalid, L.name AS locationname, R.rentalname "
	sSql = sSql & "FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.locationid = L.locationid AND R.isdeactivated = 0 AND R.orgid = " & session("orgid")
	sSql = sSql & "ORDER BY L.name, R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	'response.write iRentalId & "<br /><br />"
	response.write vbcrlf & "<select id=""rentalid"" name=""rentalid"">"

	If bIncludeAllPicks Then
		response.write vbcrlf & vbtab & "<option value=""0"">All Locations &ndash; All Rentals</option>"
	End If 
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			' This is all rentals for a location
			If bIncludeLocationAllPicks And iLocationId <> oRs("locationid") Then
				iLocationId = CLng(oRs("locationid"))
				response.write vbcrlf & vbtab & "<option value=""L" & oRs("locationid") & """ "
				If "L" & oRs("locationid") = iRentalId Then 
					response.write " selected=""selected"" "
				End If 
				response.write ">" & oRs("locationname") & " &ndash; All Rentals</option>"
			End If 

			' These are just the actual rentals
			response.write vbcrlf & vbtab & "<option value=""R" & oRs("rentalid") & """ "
			If "R" & oRs("rentalid") = iRentalId Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("locationname") & " &ndash; " & oRs("rentalname") & "</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalLocations iLocationId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalLocations( ByVal iLocationId )
	Dim oRs, sSql

	sSql = "SELECT locationid, name FROM egov_class_location "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select name=""locationid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">Any Location</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("locationid") & """ "
		If CLng(oRs("locationid")) = CLng(iLocationId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("name") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalNameAndLocation iRentalId 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalNameAndLocation( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalname, L.name AS location FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.locationid = L.locationid AND R.orgid = " & session("orgid") & " AND R.rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write "<strong>" & oRs("rentalname") & ", " & oRs("location") & "</strong>"
	Else
		response.write "Rental Not Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub

'--------------------------------------------------------------------------------------------------
function getRentalNameAndLocation( ByVal iRentalId)
	Dim sSql, oRs, lcl_return

 lcl_return = "Rental Not Found"

	sSQL = "SELECT "
 sSQL = sSQL & " R.rentalname, "
 sSQL = sSQL & " L.name AS location "
 sSQL = sSQL & " FROM egov_rentals R, "
 sSQL = sSQL &      " egov_class_location L "
	sSQL = sSQL & " WHERE R.locationid = L.locationid "
 sSQL = sSQL & " AND R.orgid = " & session("orgid")
 sSQL = sSQL & " AND R.rentalid = " & iRentalId

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	if not oRs.eof then
  		lcl_return = oRs("rentalname") & ", " & oRs("location")
	end if

	oRs.Close
	set oRs = nothing 

 getRentalNameAndLocation = lcl_return

end function


'--------------------------------------------------------------------------------------------------
' void ShowRentalPeriods iPeriodTypeId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalPeriods( ByVal iPeriodTypeId )
	Dim oRs, sSql

	sSql = "SELECT periodtypeid, periodtype FROM egov_rentalperiodtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select name=""periodtypeid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("periodtypeid") & """ "
		If CLng(oRs("periodtypeid")) = CLng(iPeriodTypeId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("periodtype") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalSupervisors iSupervisorUserId, sZeroOption 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalSupervisors( ByVal iSupervisorUserId, ByVal sZeroOption )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid")
	sSql = sSql & " AND isrentalsupervisor = 1 ORDER BY lastname, firstname, userid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""supervisoruserid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">" & sZeroOption & "</option>"
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("userid") & """ "
			If CLng(oRs("userid")) = CLng(iSupervisorUserId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReservationInfoContainer sReservationType, sRenterName, sRenterPhone, sReservationStatus, sReservedDate, sAdminName, iReservationId, sReservationTypeSelector, iTimeId
'--------------------------------------------------------------------------------------------------
Sub ShowReservationInfoContainer( ByVal sReservationType, ByVal sRenterName, ByVal sRenterPhone, ByVal sReservationStatus, ByVal sReservedDate, ByVal sAdminName, ByVal iReservationId, ByVal sReservationTypeSelector, ByVal iTimeId )

	response.write vbcrlf & "<p id=""rentalnamecontainer"">"
	response.write vbcrlf & "<table id=""reservationgeneral"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr><td class=""leftcell""><strong>Reservation Id:</strong></td><td>" & iReservationId & "</td></tr>"
	response.write vbcrlf & "<tr><td class=""leftcell""><strong>Reservation Type:</strong></td><td nowrap=""nowrap"">" & sReservationType
	If sReservationTypeSelector = "class" Then
		' Show the class name and the activity number
		response.write " &nbsp; "
		ShowClassActivityNoAndName iTimeId, True 
	End If 
	response.write "</td></tr>"

	If sRenterName <> "" Then	
		response.write vbcrlf & "<tr><td class=""leftcell"" valign=""top""><strong>For:</strong></td><td>" & sRenterName
		If sRenterPhone <> "" Then	
			response.write "<br />" & sRenterPhone
		End If						
		response.write "</td></tr>"
	End If		
	response.write vbcrlf & "<tr><td class=""leftcell""><strong>Status:</strong></td><td>" & sReservationStatus & "</td></tr>"
	response.write vbcrlf & "<tr><td class=""leftcell""><strong>Reserved:</strong></td><td>" & sReservedDate & "</td></tr>"
	If sAdminName <> "" Then	
		response.write vbcrlf & "<tr><td class=""leftcell""><strong>Made By:</strong></td><td>" & sAdminName & "</td></tr>"
	End If		
	response.write vbcrlf & "</table>"
	response.write vbcrlf & "</p>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReservationTypeFilter iReservationTypeId, bIsReservationOnly
'--------------------------------------------------------------------------------------------------
Sub ShowReservationTypeFilter( ByVal iReservationTypeId, ByVal bIsReservationOnly )
	Dim sSql, oRs, sIsReservationOnlyPick

	If bIsReservationOnly Then
		sIsReservationOnlyPick = " AND isreservation = 1 "
	Else 
		sIsReservationOnlyPick = ""
	End If 

	sSql = "SELECT reservationtypeid, reservationtype FROM egov_rentalreservationtypes WHERE orgid = " & session("orgid")
	sSql = sSql & sIsReservationOnlyPick & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""reservationtypeid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">All Reservation Types</option>"
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("reservationtypeid") & """ "
			If CLng(oRs("reservationtypeid")) = CLng(iReservationTypeId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("reservationtype") & " Only</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowReservationTypePicks iReservationTypeId
'--------------------------------------------------------------------------------------------------
Sub ShowReservationTypePicks( ByVal iReservationTypeId )
	Dim sSql, oRs

	sSql = "SELECT reservationtypeid, reservationtype FROM egov_rentalreservationtypes WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""reservationtypeid"">"
	response.write vbcrlf & vbtab & "<option value=""0"">All Reservation Types</option>"
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("reservationtypeid") & """ "
			If CLng(oRs("reservationtypeid")) = CLng(iReservationTypeId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("reservationtype") & "</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowRestrictedPeriods iRestrictedPeriodId
'--------------------------------------------------------------------------------------------------
Sub ShowRestrictedPeriods( ByVal iRestrictedPeriodId )
	Dim sSql, oRs

	sSql = "SELECT restrictionperiodid, restrictionperiod FROM egov_rentalrestrictionperiods"
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""restrictedperiodid"" name=""restrictedperiodid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("restrictionperiodid") & """"
		If CLng(oRS("restrictionperiodid")) = CLng(iRestrictedPeriodId) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("restrictionperiod") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowSameNextDayPick sSelectName, sDay, sDisabledOption
'--------------------------------------------------------------------------------------------------
Sub ShowSameNextDayPick( ByVal sSelectName, ByVal sDay, ByVal sDisabledOption )

	If sDisabledOption = "disabled" Then
		response.write "<input type=""hidden"" id=""" & sSelectName & """ name=""" & sSelectName & """ value=""" & sDay & """ />"
		If clng(sDay) = clng(0) Then 
			response.write "That Day"
		Else
			response.write "The Next Day"
		End If 
	Else 
		response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
		response.write vbcrlf & "<option value=""0""" 
		If clng(sDay) = clng(0) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">That Day</option>"
		response.write vbcrlf & "<option value=""1"""
		If clng(sDay) = clng(1) Then 
			response.write " selected=""selected"" "
		End If
		response.write ">The Next Day</option>"
		response.write vbcrlf & "</select>"
	End If 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowSameNextYearPick sSelectName, iYear
'--------------------------------------------------------------------------------------------------
Sub ShowSameNextYearPick( ByVal sSelectName, ByVal iYear )
	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	response.write vbcrlf & "<option value=""0""" 
	If clng(iYear) = clng(0) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">The Same Year</option>"
	response.write vbcrlf & "<option value=""1"""
	If clng(iYear) = clng(1) Then 
		response.write " selected=""selected"" "
	End If
	response.write ">The Next Year</option>"
	response.write vbcrlf & "</select>"

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowShortCitizenName iUserId
'--------------------------------------------------------------------------------------------------
Sub ShowShortCitizenName( ByVal iUserId )
	Dim oRs, sSql

	sSql = "SELECT ISNULL(userlname,' ') AS userlname, ISNULL(userfname,' ') AS userfname FROM egov_users WHERE userid = " & iUserid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write Left(UCase(Left(oRs("userfname"),1)) & ". " & oRs("userlname"),30)
	Else
		response.write  ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowShortClassName iTimeId
'--------------------------------------------------------------------------------------------------
Sub ShowShortClassName( ByVal iTimeId )
	Dim oRs, sSql

	sSql = "SELECT C.classname FROM egov_class_time T, egov_class C "
	sSql = sSql & " WHERE T.classid = C.classid AND T.timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write Left(ors("classname"),30)
	Else
		response.write ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowTimeTypePicks sSelectName, iTimeTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowTimeTypePicks( ByVal sSelectName, ByVal iTimeTypeId, ByVal bIsBuffer )
	Dim sSql, oRs, sWhere

	If bIsBuffer Then 
		sWhere = " AND isforbuffering = 1 "
	End If 

	sSql = "SELECT timetypeid, timetype FROM egov_rentaltimetypes WHERE orgid = " & session("orgid") & sWhere
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("timetypeid") & """"
		If CLng(oRS("timetypeid")) = CLng(iTimeTypeId) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("timetype") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' DrawTimeChoices element_name, selection
'------------------------------------------------------------------------------------------------------------
Sub DrawTimeChoices( ByVal element_name, ByVal selection )
	Dim display_hour, value_hour

	response.write vbcrlf & "<select name=""" & element_name & """ id=""" & element_name & """ class=""time_pick"">"
	response.write vbcrlf & "<option value=""none""" & SetSelection( selection, "none") & ">All Day</option>"
	
	For x = 0 To 23
		value_hour = CStr(x) & ":00"
		If x < 10 Then
			value_hour = "0" & value_hour
		End If 
		Select Case x
			Case 0
				display_hour = "12 AM"
			Case 12
				display_hour = "12 PM"
			Case Else 
				display_hour = CStr(x) 
				If x > 11 Then 
					display_hour = CStr(x - 12) & " PM"
				Else
					display_hour = CStr(x) & " AM"
				End If 
		End Select
		response.write vbcrlf & "<option value=""" & value_hour & """" & SetSelection( selection, value_hour) & ">" & display_hour & "</option>"
	Next 

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' SetSelection selection, pick_value
'------------------------------------------------------------------------------------------------------------
Function SetSelection( ByVal selection, ByVal pick_value )

	If selection = pick_value Then
		SetSelection = " selected=""selected""" 
	Else 
		SetSelection = ""
	End If 

End Function 




%>
