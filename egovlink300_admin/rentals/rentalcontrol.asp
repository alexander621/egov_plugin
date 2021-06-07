<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcontrol.asp
' AUTHOR: Steve Loar
' CREATED: 10/12/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Controls the flow of simple rental reservations. This just controls routing, and does not
'				have a gui component. Taken from the public side script of the same name
'
' MODIFICATION HISTORY
' 1.0   10/12/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSource, iRti, iCategoryId, iRentalId, iRid, sViewType, sSelectedDate, sSelectDate, x
Dim sStartDate, sEndDate, sWantedDOWs, iCitizenUserId, bOkToContinue, iDayIntervals, sCategories
Dim dLastReservationMade, bOk, sMessage, sStartDateTime, sEndDateTime, iWeekDay, bOffSeasonFlag
Dim iAdminUserId, iPeriodTypeId

If request("src") = "" Then
	response.redirect "rentalcategoryselection.asp"
Else
	sSource = request("src")
End If


Select Case sSource
	Case "dp"
		' this is from the date picking page - rentalavailability.asp
		' save the data that came across and get the rti number
		iCategoryId   = CLng(request("cid"))
		iRid          = CLng(request("rid"))
		sViewType     = "'" & dbready_string(request("viewtype"),50) & "'"
		iRentalId     = CLng(request("selectedrid"))
		sSelectedDate = "'" & dbready_string(request("selecteddate"),20) & "'"
		sSelectDate   = "'" & dbready_string(request("selectdate"),20) & "'"
		sStartDate    = "'" & dbready_string(request("startdate"),20) & "'"
		sEndDate      = "'" & dbready_string(request("enddate"),20) & "'"
		sWantedDOWs   = "'" & dbready_string(request("wanteddows"),25) & "'"
		iAdminUserId  = session("userid")
		iPeriodTypeId = GetSelectedPeriodTypeId( "anytime" )

		sSQL = "INSERT INTO egov_rentalreservationstemp ("
  sSQL = sSQL & "orgid, "
  sSQL = sSQL & "sessionid, "
  sSQL = sSQL & "cid, "
  sSQL = sSQL & "rid, "
  sSQL = sSQL & "viewtype, "
  sSQL = sSQL & "rentalid, "
		sSQL = sSQL & "requestedstartdate, "
  sSQL = sSQL & "selectdate, "
  sSQL = sSQL & "startdate, "
  sSQL = sSQL & "enddate, "
  sSQL = sSQL & "weeklydays, "
  sSQL = sSQL & "adminuserid, "
  sSQL = sSQL & "periodtypeid"
  sSQL = sSQL & ") VALUES ("
		sSQL = sSQL & session("orgid") & ", "
  sSQL = sSQL & "'" & Session.SessionID  & "', "
  sSQL = sSQL & iCategoryId   & ", "
  sSQL = sSQL & iRid          & ", "
  sSQL = sSQL & sViewType     & ", "
  sSQL = sSQL & iRentalId     & ", "
  sSQL = sSQL & sSelectedDate & ", "
  sSQL = sSQL & sSelectDate   & ", "
  sSQL = sSQL & sStartDate    & ", " 
		sSQL = sSQL & sEndDate      & ", "
  sSQL = sSQL & sWantedDOWs   & ", "
  sSQL = sSQL & iAdminUserId  & ", "
  sSQL = sSQL & iPeriodTypeId
  sSQL = sSQL & ")"
		response.write sSQL & "<br /><br />"

		iRti = RunInsertStatement( sSQL )

		' Add the dates to the tempdates table
		sSQL = "INSERT INTO egov_rentalreservationdatestemp ("
  sSQL = sSQL & "reservationtempid, "
  sSQL = sSQL & "sessionid, "
  sSQL = sSQL & "orgid, "
		sSQL = sSQL & "position, "
  sSQL = sSQL & "reservationstarttime, "
  sSQL = sSQL & "reservationendtime, "
  sSQL = sSQL & "endday"
  sSQL = sSQL & ") VALUES ("
		sSQL = sSQL & iRti & ", "
  sSQL = sSQL & "'" & Session.SessionID & "', "
  sSQL = sSQL & session("orgid") & ", "
  sSQL = sSQL & "1, "
		sSQL = sSQL & sSelectedDate & ", "
  sSQL = sSQL & sSelectedDate & ", "
  sSQL = sSQL & "0"
  sSQL = sSQL & ")"
		response.write sSQL & "<br /><br />"

		RunSQLStatement sSQL

	'send them to the time selection page
  lcl_url = "rentaldateselection.asp"
  lcl_url = lcl_url & "?rti=" & iRti
  lcl_url = lcl_url & "&selected_rentalids=" & iRentalId
  lcl_url = lcl_url & "&createPath=SIMPLE"

  response.redirect lcl_url
		'response.redirect "rentaltimeselection.asp?rti=" & iRti & "&pk=1"


'	Case "ts"
'		' this is from the time selection page
'		iRti = CLng(request("rti"))
'		isallday = clng(request("isallday"))
'		bOk = True 
'
'		GetSelectedDate iRti, iRentalId, sSelectedDate
'
'		iIncludePriceTypeId = "NULL"
'		If CLng(request("maxrentalcharges")) > CLng(0) Then
'			For x = CLng(1) To CLng(request("maxrentalcharges"))
'				If request("includepricetype" & x) = "on" Then 
'					iIncludePriceTypeId = request("pricetypeid" & x)
'					Exit For 
'				End If 
'			Next 
'		End If 
'
'		If clng(isallday) = clng(0) Then 
'
'			iStartHour = request("startinghour")
'			'response.write "iStartHour = " & iStartHour & "<br />"
'			iStartMinute = request("startingminute")
'			sStartAmPm = request("startingampm")
'			iEndHour = request("endinghour")
'			iEndMinute = request("endingminute")
'			sEndAmPm = request("endingampm")
'			iArrivalHour = iStartHour
'			iArrivalMinute = iStartMinute
'			sArrivalAmPm = sStartAmPm 
'			iDepartureHour = iEndHour
'			iDepartureMinute = iEndMinute
'			sDepartureAmPm = sEndAmPm
'
'			If iStartHour = iEndHour And iStartMinute = iEndMinute And sStartAmPm = sEndAmPm Then 
'				sMessage = "sm"
'				bOk = False 
'			End If 
'		Else
'			bOffSeasonFlag =GetOffSeasonFlag( iRentalid, sSelectedDate )
'			iWeekDay = Weekday( sSelectedDate )
'			sStartDateTime = DateValue(sSelectedDate)
'			sEndDateTime = DateValue(sSelectedDate)
'
'			' Get the opening and closing hours and set the start and end to those times
'			GetAllDayHours iRentalid, bOffSeasonFlag, iWeekDay, sStartDateTime, sEndDateTime, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm
'
'			iArrivalHour = request("arrivalhour")
'			iArrivalMinute = request("arrivalminute")
'			sArrivalAmPm = request("arrivalampm")
'			iDepartureHour = request("departurehour")
'			iDepartureMinute = request("departureminute")
'			sDepartureAmPm = request("departureampm")
'
'			If iArrivalHour = iDepartureHour And iArrivalMinute = iDepartureMinute And sArrivalAmPm = sDepartureAmPm Then 
'				sMessage = "sm"
'				bOk = False 
'			End If 
'		End If 
'
'		' update the temp record
'		sSql = "UPDATE egov_rentalreservationstemppublic "
'		sSql = sSql & "SET starthour = " & iStartHour
'		sSql = sSql & ", startminute = " & iStartMinute
'		sSql = sSql & ", startampm = '" & sStartAmPm & "'"
'		sSql = sSql & ", endhour = " & iEndHour
'		sSql = sSql & ", endminute = " & iEndMinute
'		sSql = sSql & ", endampm = '" & sEndAmPm & "'"
'		sSql = sSql & ", arrivalhour = " & iArrivalHour
'		sSql = sSql & ", arrivalminute = " & iArrivalMinute
'		sSql = sSql & ", arrivalampm = '" & sArrivalAmPm & "'"
'		sSql = sSql & ", departurehour = " & iDepartureHour
'		sSql = sSql & ", departureminute = " & iDepartureMinute
'		sSql = sSql & ", departureampm = '" & sDepartureAmPm & "'"
'		sSql = sSql & ", isallday = " & isallday
'		sSql = sSql & ", includepricetypeid = " & iIncludePriceTypeId
'		sSql = sSql & " WHERE reservationtempid = " & iRti
'		response.write sSql & "<br /><br />"
'		RunSQLStatement sSql
'
'		If bOk Then
'			If isallday = clng(0) Then
'				sStartDateTime = CDate(sSelectedDate & " " & iStartHour & ":" & iStartMinute & " " & sStartAmPm )
'				sEndDateTime = CDate(sSelectedDate & " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm )
'
'				' if the end date is less than the start date then it must end the next day
'				If sEndDateTime < sStartDateTime Then
'					sEndDateTime = DateAdd("d", 1, sEndDateTime)
'				End If 
'
'				' Round up as required by the org to the next wanted interval
'				CheckOrgRentalRoundUp iOrgId, sStartDateTime, sEndDateTime, iEndHour, iEndMinute, sEndAmPm
'
'				' set the arrival and departure times to the start and end
'				iArrivalHour = iStartHour
'				iArrivalMinute = iStartMinute
'				sArrivalAmPm = sStartAmPm
'				iDepartureHour = iEndHour
'				iDepartureMinute = iEndMinute
'				sDepartureAmPm = sEndAmPm
'				
'				' update the record with the rounded time
'				sSql = "UPDATE egov_rentalreservationstemppublic "
'				sSql = sSql & "SET starthour = " & iStartHour
'				sSql = sSql & ", startminute = " & iStartMinute
'				sSql = sSql & ", startampm = '" & sStartAmPm & "'"
'				sSql = sSql & ", endhour = " & iEndHour
'				sSql = sSql & ", endminute = " & iEndMinute
'				sSql = sSql & ", endampm = '" & sEndAmPm & "'"
'				sSql = sSql & ", arrivalhour = " & iArrivalHour
'				sSql = sSql & ", arrivalminute = " & iArrivalMinute
'				sSql = sSql & ", arrivalampm = '" & sArrivalAmPm & "'"
'				sSql = sSql & ", departurehour = " & iDepartureHour
'				sSql = sSql & ", departureminute = " & iDepartureMinute
'				sSql = sSql & ", departureampm = '" & sDepartureAmPm & "'"
'				sSql = sSql & " WHERE reservationtempid = " & iRti
'				response.write sSql & "<br /><br />"
'				RunSQLStatement sSql
'		
'				' check if the time period is too short
'				If Not CheckIfMinimumRentalTimeMet( iRentalId, sStartDateTime, sEndDateTime ) Then
'					sMessage = "st"
'					bOk = False
'				End If 
'			End If 
'		
'			If bOk Then
'				' check the availability
'				bOk = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, sMessage )	' In rentalcommonfunctions.asp
'			End If 
'		End If 
'
'		If bOk Then
'			' go to the summary page
'			response.redirect "reservationsummary.asp?rti=" & iRti
'		Else
'			' go back to the time selection page with the message.
'			response.redirect "rentaltimeselection.asp?rti=" & iRti & "&msg=" & sMessage
'		End If 
'
'	Case "sp"
'		' this is the summary page
'		iRti = CLng(request("rti"))
'
'		' if terms were not checked
'		If request("agreetoterms") = "" Then 
'			response.redirect "reservationsummary.asp?rti=" & iRti & "&msg=nt"
'		Else 
'			GetWantedDateAndTimes iRti, iRentalId, sStartDateTime, sEndDateTime
'			' check availability before going off to the payment form
'			bOk = CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, sMessage )	' In rentalcommonfunctions.asp
'			If bOk Then
'				If RentalHasNoCosts( iRentalId ) Then 
'					' go to reservation making script - this is all there is in Phase 1
'					response.redirect "rentalreservationmake.asp?rti=" & iRti & "&src=rc"
'				Else
'					' send them to the secure payment page - Phase 2
'					response.redirect Application("PAYMENTURL") & "/" & sorgVirtualSiteName & "/rentals/paymentform.asp?rti=" & iRti
'				End If 
'			Else
'				' go to time unavailable page.
'				response.redirect "rentalunavailable.asp?rti=" & iRti
'			End If 
'		End If 

	Case Else
		' this is something else that is not part of rentals
		response.redirect "rentalcategoryselection.asp"

End Select 


%>
