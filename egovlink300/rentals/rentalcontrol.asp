<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcontrol.asp
' AUTHOR: Steve Loar
' CREATED: 01/27/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Controls the flow of rental reservations. This just controls routing, and does not
'				have a gui component.
'
' MODIFICATION HISTORY
' 1.0   01/27/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSource, iRti, iCategoryId, iRentalId, iRid, sViewType, sSelectedDate, sSelectDate, x
Dim sStartDate, sEndDate, sWantedDOWs, iCitizenUserId, bOkToContinue, iDayIntervals, sCategories
Dim dLastReservationMade, bOk, sMessage, sStartDateTime, sEndDateTime, iWeekDay, bOffSeasonFlag

If request("src") = "" Then
	response.redirect "rentalcategories.asp"
Else
	sSource = request("src")
End If


Select Case sSource
	Case "dp"
		' this is from the date picking page - rentalavailability.asp
		' save the data that came across and get the rti number
		iCategoryId = CLng(request("cid"))
		iRid = CLng(request("rid"))
		sViewType = "'" & dbready_string(request("viewtype"),50) & "'"
		iRentalId = 0
		if isnumeric(request("selectedrid")) then iRentalId = CLng(request("selectedrid"))
		sSelectedDate = "'" & dbready_string(request("selecteddate"),20) & "'"
		sSelectDate = "'" & dbready_string(request("selectdate"),20) & "'"
		sStartDate = "'" & dbready_string(request("startdate"),20) & "'"
		sEndDate = "'" & dbready_string(request("enddate"),20) & "'"
		sWantedDOWs = "'" & dbready_string(request("wanteddows"),25) & "'"

		sSql = "INSERT INTO egov_rentalreservationstemppublic ( orgid, cid, rid, viewtype, rentalid, "
		sSql = sSql & " selecteddate, selectdate, startdate, enddate, wanteddows ) VALUES ( "
		sSql = sSql & iOrgId & ", " & iCategoryId & ", " & iRid & ", " & sViewType & ", " 
		sSql = sSql & iRentalId & ", " & sSelectedDate & ", " & sSelectDate & ", " & sStartDate & ", " 
		sSql = sSql & sEndDate & ", " & sWantedDOWs & " )"

		iRti = RunIdentityInsertStatement( sSql )

		'If they do not have a userid set, take them to the login page automatically
		If request.cookies("userid") = "" Or request.cookies("userid") = "-1" Then
			session("RedirectPage") = "rentals/rentalcontrol.asp?rti=" & iRti & "&src=lp"
			session("RedirectLang") = "Return to " & GetOrgDisplay( iOrgId, "rentalscategorypagetop" )
			session("ManageURL")    = ""
			session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
			response.redirect "../user_login.asp"
		Else
			' Update the citizenUserid
			iCitizenUserId = request.cookies("userid")
			sSql = "UPDATE egov_rentalreservationstemppublic SET citizenuserid = " & iCitizenUserId
			sSql = sSql & " WHERE reservationtempid = " & iRti
			RunSQLStatement sSql
		End If 

		' do the address required check here for those already logged in. see also 'lp'
		If OrgHasFeature( iOrgId, "requires address check" ) Then
			If CitizenAddressIsMissing( iCitizenUserId ) Then
				session("RedirectPage") = "rentals/rentalcontrol.asp?rti=" & iRti & "&src=lp"
				response.redirect "../getuseraddress.asp?userid=" & iCitizenUserId
			End If
		End If 

		bOkToContinue = True 
		' see if any of the categories this rental is in has restriced public reservations and if so for what period.
		If CategoryHasRestricedPeriod( iRti, iDayIntervals, sCategories, sPeriod ) Then 
			' Find out when their last reservation was for anything in the categories
			dLastReservationMade = GetLastReservationMade( iCitizenUserId, sCategories )
			'response.write "dLastReservationMade = " & dLastReservationMade & "<br />"
			'response.write "iDayIntervals = " & iDayIntervals & "<br />"
			'response.write "datediff = " & DateDiff("d", dLastReservationMade, Date()) & "<br />"
			'response.end

			' See if this user has made a reservation within the restricted period
			If DateDiff("d", dLastReservationMade, Date()) < iDayIntervals Then
				dNextDate = DateAdd("d", iDayIntervals, dLastReservationMade)
				bOkToContinue = False 
			End If 
		End If 

		If bOkToContinue Then 
			' send them to the time selection page
 			LogThePage
			response.redirect "rentaltimeselection.asp?rti=" & iRti & "&pk=1"
		Else
			' Send them to a page telling them when they can make another reservation
 			LogThePage
			response.redirect "reservationperiodmessage.asp?rti=" & iRti & "&pt=" & sPeriod
		End If 

	Case "lp"
		' this is the return from the login page
		iRti = CLng(request("rti"))
		' double check that they are logged in now
		If request.cookies("userid") = "" Or request.cookies("userid") = "-1" Then
			session("RedirectPage") = "rentals/rentalcontrol.asp?rti=" & iRti & "&src=lp"
			session("RedirectLang") = "Return to " & GetOrgDisplay( iOrgId, "rentalscategorypagetop" )
			session("ManageURL")    = ""
			session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
 			LogThePage
			response.redirect "../user_login.asp"
		Else
			' Update the citizenUserid
			iCitizenUserId = request.cookies("userid")
			sSql = "UPDATE egov_rentalreservationstemppublic SET citizenuserid = " & iCitizenUserId
			sSql = sSql & " WHERE reservationtempid = " & iRti
			RunSQLStatement sSql
		End If 

		' see if org has feature to check if the user address is filled in. See also 'dp'
		' OrgHasFeature is in include_top_functions.asp
		If OrgHasFeature( iOrgId, "requires address check" ) Then
			' check if user has an address populated. CitizenAddressIsMissing is in include_top_functions.asp
			 If CitizenAddressIsMissing( iCitizenUserId ) Then 
				' if they are missing an address, take them to a page to enter it.
				session("RedirectPage") = "rentals/rentalcontrol.asp?rti=" & iRti & "&src=lp"
				'session("RedirectLang") = "Return to " & GetOrgDisplay( iOrgId, "rentalscategorypagetop" )
				'session("ManageURL")    = ""
 				LogThePage
				response.redirect "../getuseraddress.asp?userid=" & iCitizenUserId
			End If 
		End If 
		
		bOkToContinue = True 
		' see if any of the categories this rental is in has restriced public reservations and if so for what period.
		If CategoryHasRestricedPeriod( iRti, iDayIntervals, sCategories, sPeriod ) Then 
			' Find out when their last reservation was for anything in the categories
			dLastReservationMade = GetLastReservationMade( iCitizenUserId, sCategories )

			' See if this user has made a reservation within the restricted period
			If DateDiff("d", dLastReservationMade, Date()) < iDayIntervals Then
				dNextDate = DateAdd("d", iDayIntervals, dLastReservationMade)
				bOkToContinue = False 
			End If 
		End If 

		If bOkToContinue Then 
			' send them to the time selection page
 			LogThePage
			response.redirect "rentaltimeselection.asp?rti=" & iRti & "&pk=1"
		Else
			' Send them to a page telling them when they can make another reservation
 			LogThePage
			response.redirect "reservationperiodmessage.asp?rti=" & iRti & "&pt=" & sPeriod
		End If 

	Case "ts"
		' this is from the time selection page
		iRti = CLng(request("rti"))
		isallday = clng(request("isallday"))
		bOk = True 

		GetSelectedDate iRti, iRentalId, sSelectedDate

		iIncludePriceTypeId = "NULL"
		If CLng(request("maxrentalcharges")) > CLng(0) Then
			For x = CLng(1) To CLng(request("maxrentalcharges"))
				If request("includepricetype" & x) = "on" Then 
					iIncludePriceTypeId = request("pricetypeid" & x)
					Exit For 
				End If 
			Next 
		End If 

		If clng(isallday) = clng(0) Then 

			iStartHour = request("startinghour")
			'response.write "iStartHour = " & iStartHour & "<br />"
			iStartMinute = request("startingminute")
			sStartAmPm = request("startingampm")
			iEndHour = request("endinghour")
			iEndMinute = request("endingminute")
			sEndAmPm = request("endingampm")
			iArrivalHour = iStartHour
			iArrivalMinute = iStartMinute
			sArrivalAmPm = sStartAmPm 
			iDepartureHour = iEndHour
			iDepartureMinute = iEndMinute
			sDepartureAmPm = sEndAmPm

			If iStartHour = iEndHour And iStartMinute = iEndMinute And sStartAmPm = sEndAmPm Then 
				sMessage = "sm"
				bOk = False 
			End If 
		Else
			bOffSeasonFlag =GetOffSeasonFlag( iRentalid, sSelectedDate )
			iWeekDay = Weekday( sSelectedDate )
			sStartDateTime = DateValue(sSelectedDate)
			sEndDateTime = DateValue(sSelectedDate)

			' Get the opening and closing hours and set the start and end to those times
			GetAllDayHours iRentalid, bOffSeasonFlag, iWeekDay, sStartDateTime, sEndDateTime, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm

			iArrivalHour = request("arrivalhour")
			iArrivalMinute = request("arrivalminute")
			sArrivalAmPm = request("arrivalampm")
			iDepartureHour = request("departurehour")
			iDepartureMinute = request("departureminute")
			sDepartureAmPm = request("departureampm")

			If iArrivalHour = iDepartureHour And iArrivalMinute = iDepartureMinute And sArrivalAmPm = sDepartureAmPm Then 
				sMessage = "sm"
				bOk = False 
			End If 
		End If 

		' update the temp record
		sSql = "UPDATE egov_rentalreservationstemppublic "
		sSql = sSql & "SET starthour = " & iStartHour
		sSql = sSql & ", startminute = " & iStartMinute
		sSql = sSql & ", startampm = '" & sStartAmPm & "'"
		sSql = sSql & ", endhour = " & iEndHour
		sSql = sSql & ", endminute = " & iEndMinute
		sSql = sSql & ", endampm = '" & sEndAmPm & "'"
		sSql = sSql & ", arrivalhour = " & iArrivalHour
		sSql = sSql & ", arrivalminute = " & iArrivalMinute
		sSql = sSql & ", arrivalampm = '" & sArrivalAmPm & "'"
		sSql = sSql & ", departurehour = " & iDepartureHour
		sSql = sSql & ", departureminute = " & iDepartureMinute
		sSql = sSql & ", departureampm = '" & sDepartureAmPm & "'"
		sSql = sSql & ", isallday = " & isallday
		sSql = sSql & ", includepricetypeid = " & iIncludePriceTypeId
		sSql = sSql & " WHERE reservationtempid = " & iRti
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		If bOk Then
			If isallday = clng(0) Then
				sStartDateTime = CDate(sSelectedDate & " " & iStartHour & ":" & iStartMinute & " " & sStartAmPm )
				sEndDateTime = CDate(sSelectedDate & " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm )

				' if the end date is less than the start date then it must end the next day
				If sEndDateTime < sStartDateTime Then
					sEndDateTime = DateAdd("d", 1, sEndDateTime)
				End If 

				' Round up as required by the org to the next wanted interval
				CheckOrgRentalRoundUp iOrgId, sStartDateTime, sEndDateTime, iEndHour, iEndMinute, sEndAmPm

				' set the arrival and departure times to the start and end
				iArrivalHour = iStartHour
				iArrivalMinute = iStartMinute
				sArrivalAmPm = sStartAmPm
				iDepartureHour = iEndHour
				iDepartureMinute = iEndMinute
				sDepartureAmPm = sEndAmPm
				
				' update the record with the rounded time
				sSql = "UPDATE egov_rentalreservationstemppublic "
				sSql = sSql & "SET starthour = " & iStartHour
				sSql = sSql & ", startminute = " & iStartMinute
				sSql = sSql & ", startampm = '" & sStartAmPm & "'"
				sSql = sSql & ", endhour = " & iEndHour
				sSql = sSql & ", endminute = " & iEndMinute
				sSql = sSql & ", endampm = '" & sEndAmPm & "'"
				sSql = sSql & ", arrivalhour = " & iArrivalHour
				sSql = sSql & ", arrivalminute = " & iArrivalMinute
				sSql = sSql & ", arrivalampm = '" & sArrivalAmPm & "'"
				sSql = sSql & ", departurehour = " & iDepartureHour
				sSql = sSql & ", departureminute = " & iDepartureMinute
				sSql = sSql & ", departureampm = '" & sDepartureAmPm & "'"
				sSql = sSql & " WHERE reservationtempid = " & iRti
				'response.write sSql & "<br /><br />"
				RunSQLStatement sSql
		
				' check if the time period is too short
				If Not CheckIfMinimumRentalTimeMet( iRentalId, sStartDateTime, sEndDateTime ) Then
					sMessage = "st"
					bOk = False
				End If 
			End If 
		
			If bOk Then
				' check the availability
				bOk = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, sMessage )	' In rentalcommonfunctions.asp
			End If 
		End If 

		If bOk Then
			' go to the summary page
 			LogThePage
			response.redirect "reservationsummary.asp?rti=" & iRti
		Else
			' go back to the time selection page with the message.
 			LogThePage
			response.redirect "rentaltimeselection.asp?rti=" & iRti & "&msg=" & sMessage
		End If 

	Case "sp"
		' this is the summary page
		iRti = CLng(request("rti"))


		' if terms were not checked
		If request("agreetoterms") = "" Then 
 			LogThePage
			response.redirect "reservationsummary.asp?rti=" & iRti & "&msg=nt"
		Else 
			GetWantedDateAndTimes iRti, iRentalId, sStartDateTime, sEndDateTime
			' check availability before going off to the payment form
			bOk = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, sMessage )	' In rentalcommonfunctions.asp

			If bOk Then
				sTotalAmount = GetTotalAmount( iRti )
				If RentalHasNoCosts( iRentalId ) or sTotalAmount = 0 Then 
					' go to reservation making script - this is all there is in Phase 1
 					LogThePage
					response.redirect "rentalreservationmake.asp?rti=" & iRti & "&src=rc"
				Else
					' send them to the secure payment page - Phase 2
 					LogThePage
					response.redirect Application("PAYMENTURL") & "/" & sorgVirtualSiteName & "/rentals/paymentform.asp?rti=" & iRti
				End If 
			Else
				' go to time unavailable page.
 				LogThePage
				response.redirect "rentalunavailable.asp?rti=" & iRti
			End If 
		End If 

	Case Else
		' this is something else that is not part of rentals, like a spider, robot, or hack
 		LogThePage
		response.redirect "rentalcategories.asp"

End Select 

Sub LogThePage( )
	Dim sSql, oCmd, sScriptName, sVirtualDirectory, aVirtualDirectory, sPage, arr, sUserAgent, sUserAgentGroup

	sScriptName = Request.ServerVariables("SCRIPT_NAME")

	If request.servervariables("http_user_agent") <> "" Then 
		sUserAgent = "'" & Track_DBsafe(Trim(Left(request.servervariables("http_user_agent"),480))) & "'"
	Else
		sUserAgent = "NULL"
	End If 

	If Len(Trim(request.servervariables("http_user_agent"))) > 0 Then 
		sUserAgentGroup = "'" & GetUserAgentGroup( LCase(request.servervariables("http_user_agent")) ) & "'"
	Else
		sUserAgentGroup = "'" & GetUntrackedUserAgentGroup( ) & "'"
	End If 

	' Get the virtual directory
	aVirtualDirectory = Split(sScriptName, "/", -1, 1) 
	sVirtualDirectory = "/" & aVirtualDirectory(1) 
	sVirtualDirectory = "'" & Replace(sVirtualDirectory,"/","") & "'"

	' Get the page
	For Each arr in aVirtualDirectory 
		sPage = arr 
	Next 

	sSql = "INSERT INTO egov_pagelog ( virtualdirectory, applicationside, page, loadtime, scriptname, querystring, "
	sSql = sSql & " servername, remoteaddress, requestmethod, orgid, userid, username, sectionid, documenttitle, useragent, useragentgroup, requestformcollection, cookiescollection, sessioncollection, sessionid  ) VALUES ( "
	sSql = sSql & sVirtualDirectory & ", "
	sSql = sSql & "'public', "
	sSql = sSql & "'" & sPage & "', "
	sSql = sSql & FormatNumber(iLoadTime,3,,,0) & ", "
	sSql = sSql & "'" & sScriptName & "', "

	If Request.ServerVariables("QUERY_STRING") <> "" Then 
		sSql = sSql & "'" & Track_DBsafe(Left(Request.ServerVariables("QUERY_STRING"),500)) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 
	' our server name
	sSql = sSql & "'" & Request.ServerVariables("SERVER_NAME") & "', "

	' remote address
	sSql = sSql & "'" & Request.ServerVariables("REMOTE_ADDR") & "', "

	' request method - GET or POST
	sSql = sSql & "'" & Request.ServerVariables("REQUEST_METHOD") & "', "

	' orgid
	If iorgid <> "" Then 
		sSql = sSql & iorgid & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Userid
	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" and isnumeric(request.cookies("userid")) Then
		sSql = sSql & request.cookies("userid") & ", "
	Else
		sSql = sSql & "NULL, "
		response.cookies("userid") = ""
	End If 

	' Get username
	If sUserName <> "" Then
		sSql = sSql & "'" & Track_DBsafe(sUserName) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Section Id for the old LogPageVisit functionality
	If iSectionID <> "" Then 
		sSql = sSql & iSectionID & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Document Title for the old LogPageVisit functionality
	If sDocumentTitle <> "" Then 
		sSql = sSql & "'" & Track_DBsafe(sDocumentTitle) & "',  "
	Else
		sSql = sSql & "NULL, "
	End If 

	' User Agent
	sSql = sSql & sUserAgent & ", "

	' User Agent Group
	sSql = sSql & sUserAgentGroup & ", "

	sSql = sSql & "'" & Track_DBsafe(GetRequestformInformation()) & "',"
	sSql = sSql & "'" & GetCookiesCollection() & "',"
	sSql = sSql & "'" & GetSessionCollection() & "',"


	sSql = sSql & "'" & Session.SessionID & "'"

	sSql = sSql & " )"
	'response.write sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql

	session("sSql") = sSql
	oCmd.Execute
	session("sSql") = ""

	Set oCmd = Nothing


End Sub 
Function GetCookiesCollection()
	Collection = ""
	on error resume next
	For Each Item in Request.Cookies
		Collection = Collection & Item & ":  " & request.cookies(Item) & vbcrlf
	Next
	on error goto 0
	GetCookiesCollectionCollection = track_dbsafe(Collection)
End Function
Function GetSessionCollection()
	sSessionLog = ""
	on error resume next
	For each session_name in Session.Contents
		sSessionLog = sSessionLog & session_name & ":  " & session(session_name) & vbcrlf
	Next
	on error goto 0

	GetSessionCollection = track_dbsafe(sSessionLog)
End Function


'------------------------------------------------------------------------------
Function GetUserAgentGroup( ByVal sUserAgent )
	Dim sSql, oRs, sUserAgentGroup

	sUserAgentGroup = GetUntrackedUserAgentGroup()

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 0 AND isactive = 1 ORDER BY checkorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If clng(InStr( 1, sUserAgent, LCase(oRs("useragentgroup")), 1 )) > clng(0) Then
			sUserAgentGroup = oRs("useragentgroup")
			Exit Do 
		End If 
		oRs.MoveNext
	Loop 
	
	oRs.Close
	Set oRs = Nothing 
	
	GetUserAgentGroup = sUserAgentGroup

End Function 


'------------------------------------------------------------------------------
Function GetUntrackedUserAgentGroup( )
	Dim sSql, oRs

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetUntrackedUserAgentGroup = oRs("useragentgroup")
	Else
		GetUntrackedUserAgentGroup = "untracked"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 
'--------------------------------------------------------------------------------------------------
' FUNCTION GETREQUESTFORMINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetRequestFormInformation()
	Dim sReturnValue, key
	
	sReturnValue = ""

	For each key in request.Form
		If key <> "accountnumber" And key <> "cvv2" Then 
			sReturnValue = sReturnValue & key & ":" & request.form(key) & "<br />" & vbcrlf
		End If 
	Next 
	
	GetRequestFormInformation = sReturnValue

End Function



%>
