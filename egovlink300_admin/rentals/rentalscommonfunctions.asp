<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalscommonfunctions.asp
' AUTHOR: Steve Loar
' CREATED: 08/21/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared functions for rentals. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   08/21/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' boolean AccountLedgerFeeRowExists( iReservationFeeTypeId, sReservationFeeType )
'--------------------------------------------------------------------------------------------------
Function AccountLedgerFeeRowExists( ByVal iReservationFeeTypeId, ByVal sReservationFeeType )
	Dim sSql, oRs

	sSql = "SELECT COUNT(ledgerid) AS hits FROM egov_accounts_ledger "
	sSql = sSql & "WHERE reservationfeetype = '" & sReservationFeeType & "' AND reservationfeetypeid = " & iReservationFeeTypeId
	'response.write sSql & "<br /><br />"


	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			AccountLedgerFeeRowExists = True 
		Else
			AccountLedgerFeeRowExists = False 
		End If 
	Else
		AccountLedgerFeeRowExists = False 
	End If
	
	oRs.Close 
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' datetime AddPostBufferTime( iRentalid, bOffSeasonFlag, dEndDateTime, dStartDateTime)
'--------------------------------------------------------------------------------------------------
Function AddPostBufferTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByRef dEndDateTime, ByVal dStartDateTime )
	Dim sSql, oRs, iWeekday

	iWeekday = Weekday(dStartDateTime)

	sSql = "SELECT ISNULL(postbuffer,0) AS postbuffer, P.dateaddstring AS postdateaddstring "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes P "
	sSql = sSql & " WHERE D.postbuffertimetypeid = P.timetypeid "
	sSql = sSql & " AND rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag
	sSql = sSql & " AND dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("postbuffer")) > CLng(0) Then 
			AddPostBufferTime = DateAdd(oRs("postdateaddstring"), oRs("postbuffer"), CDate(dEndDateTime))
		Else
			AddPostBufferTime = CDate(dEndDateTime)
		End If 
	Else
		AddPostBufferTime = CDate(dEndDateTime)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function  


'--------------------------------------------------------------------------------------------------
' integer CalculateDurationInHours( sStartDateTime, sEndDateTime )
'--------------------------------------------------------------------------------------------------
Function CalculateDurationInHours( ByVal sStartDateTime, ByVal sEndDateTime )
	Dim iDuration

	iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
	iDuration = CDbl(iDuration / 60)

	CalculateDurationInHours = iDuration

End Function 


'--------------------------------------------------------------------------------------------------
' double CalculateReservationTotal( iReservationId )
'--------------------------------------------------------------------------------------------------
Function CalculateReservationTotal( ByVal iReservationId, ByVal sField )
	Dim oRs, sSql, dTotalReservationFees, dTotalReservationDateFees, dTotalReservationDateItemFees

	' Get the total reservation fees (deposits and alcohol)
	dTotalReservationFees = GetReservationFeesTotal( iReservationId, sField )

	'Get the total reservation date fees (hourly rate fees)
	dTotalReservationDateFees = GetReservationDateFeesTotal( iReservationId, sField )

	' Get the total reservation date item fees (Chairs and tables)
	dTotalReservationDateItemFees = GetReservationDateItemFeesTotal( iReservationId, sField )

	CalculateReservationTotal = CDbl(dTotalReservationFees + dTotalReservationDateFees + dTotalReservationDateItemFees)

End Function 


'--------------------------------------------------------------------------------------------------
' void CancelReservationDate iReservationDateId
'--------------------------------------------------------------------------------------------------
Sub CancelReservationDate( ByVal iReservationDateId )
	Dim iCancelStatusId, sSql

	' Set the date fees to $0
	sSql = "UPDATE egov_rentalreservationdatefees SET feeamount = 0.0000 WHERE reservationdateid = " & iReservationDateId
	RunSQLStatement sSql

	' Set the date items to 0 qty and $0 fee
	sSql = "UPDATE egov_rentalreservationdateitems SET feeamount = 0.0000, quantity = 0 WHERE reservationdateid = " & iReservationDateId
	RunSQLStatement sSql

	' get the cancel status id
	iCancelStatusId = GetReservationStatusId( "iscancelled" )

	' set the reservation date to cancel status
	sSql = "UPDATE egov_rentalreservationdates SET statusid = " & iCancelStatusId & " WHERE reservationdateid = " & iReservationDateId
	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' string CheckForExistingReservations( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType, bOffSeasonFlag, bIncludeBuffer )
'--------------------------------------------------------------------------------------------------
Function CheckForExistingReservations( ByVal iRentalid, ByVal dWantedStartTime, ByVal dWantedEndTime, ByVal sPeriodType, ByVal bOffSeasonFlag, ByVal bIsForClass, ByVal bIncludeBuffer )
	Dim sSql, oRs, sCompareEndTime

	If sPeriodType <> "anytime" Then 
		If sPeriodType = "selectedperiod" Then
			If Not bIsForClass Then 
				If EndTimeIsNotClosingTime( iRentalId, dWantedEndTime, bOffSeasonFlag, dWantedStartTime ) Then
					If bIncludeBuffer Then 
						' Add on the end buffer
						dWantedEndTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, dWantedEndTime, dWantedStartTime )
					End If 
				End If 
			End If 
			' we will add a minute to this so start time can be the same as the end of bufferend time of another reservation
			dWantedStartTime = DateAdd("n", 1, dWantedStartTime)
			' we will remove a minute so the end of the buffer can be the same minute as the start of another reservation
			dWantedEndTime = DateAdd("n", -1, dWantedEndTime)

			If bIncludeBuffer Then
				sCompareEndTime = "reservationendtime"
			Else
				sCompareEndTime = "billingendtime"
			End If 

			' set sql to look for conflicting times
			sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
			sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
			sSql = sSql & " AND (reservationstarttime BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
			sSql = sSql & " OR " & sCompareEndTime & " BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
			sSql = sSql & " OR (reservationstarttime <= '" & dWantedStartTime & "' AND " & sCompareEndTime & " >= '" & dWantedEndTime & "'))"
		Else
			' allday - set sql to look for any starting time on that day
			sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
			sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
			sSql = sSql & " AND reservationstarttime > '" & DateValue(dWantedStartTime) & " 0:00 AM' "
			sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dWantedStartTime)) & " 0:00 AM' "
		End If 
		'response.write sSql
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then
			If CLng(oRs("hits")) > CLng(0) Then
				' conflicting reservation times
				CheckForExistingReservations = "No"
			Else
				' No conflicts found
				CheckForExistingReservations = "Yes"
			End If 
		Else
			' No rows returned - not likely using count(), but still no conflicts
			CheckForExistingReservations = "Yes"
		End If 

		oRs.Close
		Set oRs = Nothing 
	Else
		If RentalHasAnyAvailableTime( iRentalid, bOffSeasonFlag, Weekday(dWantedStartTime), dWantedStartTime ) Then 
			CheckForExistingReservations = "Yes"
		Else
			CheckForExistingReservations = "No"
		End If 
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean CheckIfMinimumRentalTimeMet( iRentalId, dStartTime, dEndTime )
'--------------------------------------------------------------------------------------------------
Function CheckIfMinimumRentalTimeMet( ByVal iRentalId, ByVal dStartTime, ByRef dEndTime )
	Dim sOffSeasonFlag, sSql, oRs, iWeekday, sReturn

	sOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(dStartTime) )

	iWeekday = Weekday(dStartTime)

	sReturn = True 

	sSql = "SELECT ISNULL(D.minimumrental,0) AS minimumrental, M.dateaddstring "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes M "
	sSql = sSql & " WHERE D.minimumrentaltimetypeid = M.timetypeid "
	sSql = sSql & " AND D.rentalid = " & iRentalid & " AND D.isoffseason = " & sOffSeasonFlag & " AND D.dayofweek = " & iWeekday

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sInterval = oRs("dateaddstring")
		If sInterval = "h" Then 
			iMinimumTime = CDbl(oRs("minimumrental")) * CDbl(60)
		Else 
			iMinimumTime = CDbl(oRs("minimumrental"))
		End If 
		' see if the interval is too small for the org
		'response.write CDbl(DateDiff("n", dStartTime, dEndTime)) & "<br />"
		'response.write iMinimumTime & "<br />"
		If CDbl(DateDiff("n", dStartTime, dEndTime)) < CDbl(iMinimumTime) Then
			sReturn = False 
		Else
			sReturn = True 
		End If 
	Else
		sReturn = False  
	End If 

	oRs.Close 
	Set oRs = Nothing 

	CheckIfMinimumRentalTimeMet = sReturn

	'response.write sReturn & "<br /><br />"

End Function 


'--------------------------------------------------------------------------------------------------
' void CheckMinimumRentalInterval dStartTime, dEndTime, iEndHour, iEndMinute, sEndAmPm
'--------------------------------------------------------------------------------------------------
Sub CheckMinimumRentalInterval( ByVal dStartTime, ByRef dEndTime, ByRef iEndHour, ByRef iEndMinute, ByRef sEndAmPm )
	' THis is wrong , but should be used for the minimal time for each rental instead.
	Dim sSql, oRs, sInterval, iMinimumTime

	' Get the minimum interval for the org
	sSql = "SELECT ISNULL(O.rentalroundup,0) AS rentalroundup, ISNULL(T.dateaddstring,'n') AS dateaddstring "
	sSql = sSql & " FROM Organizations O, egov_rentaltimetypes T "
	sSql = sSql & " WHERE O.rentalrounduptimetypeid = T.timetypeid AND O.orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sInterval = oRs("dateaddstring")
		iMinimumTime = CLng(oRs("rentalroundup"))
		' see if the interval is too small for the org
		If CLng(DateDiff(sInterval, dStartTime, dEndTime)) < iMinimumTime Then
			' if too small change the end time to be the minimum interval
			dEndTime = DateAdd(sInterval, iMinimumTime, dStartTime)

			' now set the end time pick to be the new end time
			SetEndingTimes dEndTime, iEndHour, iEndMinute, sEndAmPm
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void CheckOrgRentalRoundUp dStartTime, dEndTime, iEndHour, iEndMinute, sEndAmPm
'--------------------------------------------------------------------------------------------------
Sub CheckOrgRentalRoundUp( ByVal dStartTime, ByRef dEndTime, ByRef iEndHour, ByRef iEndMinute, ByRef sEndAmPm )
	Dim sSql, oRs, sInterval, iRoundUp, iTimeInterval, iIntervals

	' Get the minimum interval for the org
	sSql = "SELECT ISNULL(O.rentalroundup,0) AS rentalroundup, ISNULL(T.dateaddstring,'n') AS dateaddstring "
	sSql = sSql & " FROM Organizations O, egov_rentaltimetypes T "
	sSql = sSql & " WHERE O.rentalrounduptimetypeid = T.timetypeid AND O.orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sInterval = oRs("dateaddstring")
		iRoundUp = CLng(oRs("rentalroundup"))
		If iRoundUp > CLng(0) Then 
			iTimeInterval = DateDiff(sInterval, dStartTime, dEndTime)
			iIntervals = iTimeInterval \ iRoundUp		' integer whole intervals in the wanted time.
			iNewTime = iIntervals * iRoundUp
			If iTimeInterval Mod iRoundUp > 0 Then
				iNewTime = iNewTime + iRoundUp		' round up to the next whole interval
			End If 
			dEndTime = DateAdd(sInterval, iNewTime, dStartTime)  ' Apply the roundup
			' now set the end time pick to be the new end time
			SetEndingTimes dEndTime, iEndHour, iEndMinute, sEndAmPm
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, bIsForClass )
'--------------------------------------------------------------------------------------------------
Function CheckRentalAvailability( ByVal iRentalid, ByVal sStartDateTime, ByVal sEndDateTime, ByVal bIsForClass )
	Dim sOffSeasonFlag, sCheckReturn, sReturn

	sReturn        = ""
	sOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )

	' Check if rental is open that day and that we are not looking for a time when it is not open
	sCheckReturn = CheckRentalHours( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag )
	
	If Left(sCheckReturn,2) = "No" Then 
		'sReturn = sStartDateTime & "closed" & sCheckReturn
		sReturn = "closed"
	End If 

	If sReturn <> "closed" Then 
		' Check if the time is available without the end buffer
		If CheckForExistingReservations( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag, bIsForClass, False ) = "No" Then
			sReturn = "conflict" 
		End If 
	End If 

	If sReturn = "" Then 
		' Then set it to the OK flag
		sReturn = "OK"
		If Not bIsForClass Then 
			' Check if the time is available with the buffer included
			If CheckForExistingReservations( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag, bIsForClass, True ) = "No" Then
				sReturn = "buffer" 
			End If 
			' Check the minimum rental period met
			If Not CheckIfMinimumRentalTimeMet( iRentalId, sStartDateTime, sEndDateTime ) Then
				If sReturn = "OK" Then 
					sReturn = "short"
				Else
					sReturn = "buffershort"
				End If 
			End If 
		End If 
	End If 

	'response.write sReturn & "<br /><br />"

	CheckRentalAvailability = sReturn

End Function 


'--------------------------------------------------------------------------------------------------
' string CheckRentalHours( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType, bOffSeasonFlag )
'--------------------------------------------------------------------------------------------------
Function CheckRentalHours( ByVal iRentalid, ByRef dWantedStartTime, ByRef dWantedEndTime, ByVal sPeriodType, ByVal bOffSeasonFlag )
	Dim sSql, oRs, iWeekday, iWantedDuration, iMinimumDuration, sInterval, dOpeningTime, dLatestStartTime
	Dim dClosingTime

	iWeekday = Weekday(dWantedStartTime)

	sSql = "SELECT isavailabletopublic, isopen, openinghour, dbo.AddLeadingZeros(openingminute,2) AS openingminute, openingampm, "
	sSql = sSql & " closinghour, dbo.AddLeadingZeros(closingminute,2) AS closingminute, closingampm, closingday, "
	sSql = sSql & " lateststarthour, lateststartminute, lateststartampm, "
	sSql = sSql & " ISNULL(postbuffer,0) AS postbuffer, P.dateaddstring AS postdateaddstring, minimumrental, M.dateaddstring AS minimumdateaddstring "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes P, egov_rentaltimetypes M "
	sSql = sSql & " WHERE D.postbuffertimetypeid = P.timetypeid AND D.minimumrentaltimetypeid = M.timetypeid "
	sSql = sSql & " AND rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' First see if the rental is open on this day of the week
		If oRs("isopen") Then 
			If sPeriodType = "selectedperiod" Then 
				' See if the start time is before the opening time
				'response.write "dWantedStartTime: " & dWantedStartTime & "<br /><br />"
				dOpeningTime = CDate(DateValue(dWantedStartTime) & " " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm"))
				'response.write DateValue(dWantedStartTime) & " " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm")
				If CDate(dOpeningTime) <= CDate(dWantedStartTime) Then 
					' If there is a latest start time see it we are past that time
					dLatestStartTime = CDate(DateValue(dWantedStartTime) & " " & oRs("lateststarthour") & ":" & oRs("lateststartminute") & " " & oRs("lateststartampm"))
					If CDate(dLatestStartTime) >= CDate(dWantedStartTime) Then
						' Finally see if we want to leave after the closing time
						If CLng(oRs("closingday")) = CLng(0) Then
							dClosingTime = CDate(DateValue(dWantedStartTime) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
						Else
							dClosingTime = CDate(DateAdd("d", 1, DateValue(dWantedStartTime)) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
						End If 
'						If CDate(dWantedEndTime) > CDate(dClosingTime) Then
						If DateDiff( "n", CDate(dWantedEndTime), CDate(dClosingTime)) < 0 Then 
							' Time is after the rental closes
'							CheckRentalHours = "No5 " & dWantedEndTime & " " & DateDiff( "n", CDate(dWantedEndTime), CDate(dClosingTime) )
							CheckRentalHours = "No"
						Else
							' Finally is it OK for rental during the wanted time
'							CheckRentalHours = "Yes1 " & dWantedEndTime & " " & DateDiff( "n", CDate(dWantedEndTime), CDate(dClosingTime) )
							CheckRentalHours = "Yes"
						End If 
					Else
						' The start time is too late in the day
						CheckRentalHours = "No" '& dLatestStartTime & "|" & dWantedStartTime
					End If 
				Else
					' the start time is before the place opens
					CheckRentalHours = "No"
				End If 
			Else
				' for allday and any time period pick we do not have a start and end time for the reservation so it is "OK" if the rental is open that day.
				CheckRentalHours = "Yes"
			End If 
		Else
			' It is not open on that day of the week.
			CheckRentalHours = "No"
		End If 
	Else
		' It is not open on that day, or at least no time has been entered for this DOW.
		CheckRentalHours = "No"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' void ClearClassTimeReservation iReservationId
'--------------------------------------------------------------------------------------------------
Sub ClearClassTimeReservation( ByVal iReservationId )
	Dim sSql, oRs

	' Get the timeid from the reservation row
	sSql = "SELECT ISNULL(timeid,0) AS timeid FROM egov_rentalreservations WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("timeid")) > CLng(0) Then 
			' Set the reservationid to NULL for the classtime row
			sSql = "UPDATE egov_class_time SET reservationid = NULL WHERE timeid = " & oRs("timeid")
			RunSQLStatement sSql
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub


'--------------------------------------------------------------------------------------------------
' void ClearTempReservation iReservationTempId
'--------------------------------------------------------------------------------------------------
Sub ClearTempReservation( ByVal iReservationTempId )
	Dim sSql
		
	sSql = "DELETE FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & session("orgid")

	RunSQLStatement sSql		' in ../includes/common.asp

End Sub 


'--------------------------------------------------------------------------------------------------
' void CreateRentalReservationDateFees iReservationDateId, iReservationId, iRentalid, bOffSeasonFlag, iWeekday, sUserType, sStartDateTime, sEndDateTime, sReservationTypeSelection
'--------------------------------------------------------------------------------------------------
Sub CreateRentalReservationDateFees( ByVal iReservationDateId, ByVal iReservationId, ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal sUserType, ByVal sStartDateTime, ByVal sEndDateTime, ByRef dTotalAmount, ByVal sReservationTypeSelection )
	Dim sSql, oRs, sAccount, sHour, sMinute, sAmPm, iDuration, iFeeAmount, bAddFee

	' Get the Rental Reservation Fees for the day
	sSql = "SELECT R.pricetypeid, P.pricetypename, ISNULL(accountid,0) AS accountid, R.ratetypeid, ISNULL(amount,0.00) AS amount, "
	sSql = sSql & "ISNULL(R.starthour,0) AS starthour, dbo.AddLeadingZeros(ISNULL(R.startminute,0),2) AS startminute, "
	sSql = sSql & "ISNULL(R.startampm,'AM') AS startampm, P.pricetype, P.isbaseprice, P.isfee, P.hasstarttime, P.isweekendsurcharge, "
	sSql = sSql & "ISNULL(P.basepricetypeid,0) AS basepricetypeid, P.checkresidency, P.isresident, T.datediffstring, alwaysadd "
	sSql = sSql & "FROM egov_rentaldayrates R, egov_rentaldays D, egov_price_types P, egov_rentalratetypes T "
	sSql = sSql & "WHERE D.dayid = R.dayid AND D.rentalid = R.rentalid AND R.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND T.ratetypeid = R.ratetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If oRs("accountid") = CLng(0) Then 
			sAccount = "NULL"
		Else
			sAccount = oRs("accountid")
		End If 

		sHour = "NULL"
		sMinute = "NULL"
		sAmPm = "NULL"

		If oRs("isfee") Then	
			' These should be the weekend surcharge
			If oRs("hasstarttime") Then
				sHour = oRs("starthour")
				sMinute = oRs("startminute")
				sAmPm = "'" & oRs("startampm") & "'"
				sAmPmValue = oRs("startampm")
				'response.write "sSurchargeStart = " & DateValue(sStartDateTime) & " " & sHour & ":" & sMinute & " " & sAmPmValue & "<br /><br />"
				sSurchargeStart = CDate(DateValue(sStartDateTime) & " " & sHour & ":" & sMinute & " " & sAmPmValue)
			Else
				sSurchargeStart = sStartDateTime
			End If 

			'response.write "sStartDateTime: " & sStartDateTime & "<br />"
			'response.write "sEndDateTime: " & sEndDateTime & "<br />"
			'response.write DateDiff( "n", CDate(sEndDateTime), CDate(sSurchargeStart)) & "<br /><br />"
			If DateDiff( "n", CDate(sEndDateTime), CDate(sSurchargeStart)) < 0 Then
				bAddFee = True 
				'response.write DateDiff( "n", CDate(sStartDateTime), CDate(sSurchargeStart)) & "<br /><br />"
				If DateDiff( "n", CDate(sStartDateTime), CDate(sSurchargeStart)) < 0 Then
					iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
					If oRs("datediffstring") = "h" Then
						iDuration = CDbl(iDuration / 60)
					End If 
					sRate = CDbl(oRs("amount"))
					iFeeAmount = CDbl(iDuration) * sRate
				Else
					iDuration = DateDiff("n", sSurchargeStart, CDate(sEndDateTime))
					If oRs("datediffstring") = "h" Then
						iDuration = CDbl(iDuration / 60)
					ElseIf oRs("datediffstring") = "d" Then
						iDuration = CDbl(1.00)
					End If 
					sRate = CDbl(oRs("amount"))
					iFeeAmount = CDbl(iDuration) * sRate
				End If 
			Else
				bAddFee = False 
			End If 
		Else 
			' These should be the resident, non-resident or everyone fees
			If oRs("checkresidency") Then 
				If oRs("alwaysadd") Then
					' This is Menlo Park's Resident Rate
					bAddFee = True
					iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
					If oRs("datediffstring") = "h" Then
						iDuration = CDbl(iDuration / 60)
					ElseIf oRs("datediffstring") = "d" Then
						iDuration = CDbl(1.00)
					End If 
					sRate = CDbl(oRs("amount"))
					iFeeAmount = CDbl(iDuration) * sRate
				Else 
					' This should handle the Montgomery type rates and the Menlo Park Non-Resident rates
					If sUserType = oRs("pricetype") Then 
						bAddFee = True
						iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
						If oRs("datediffstring") = "h" Then
							iDuration = CDbl(iDuration / 60)
						ElseIf oRs("datediffstring") = "d" Then
							iDuration = CDbl(1.00)
						End If 
						sRate = CDbl(oRs("amount"))
						iFeeAmount = CDbl(iDuration) * sRate
					Else
						bAddFee = False
					End If 
				End If 
			Else
				' The everyone fees
				iDuration = DateDiff("n", CDate(sStartDateTime), CDate(sEndDateTime))
				If oRs("datediffstring") = "h" Then
					iDuration = CDbl(iDuration / 60)
				ElseIf oRs("datediffstring") = "d" Then
					iDuration = CDbl(1.00)
				End If 
				sRate = CDbl(oRs("amount"))
				iFeeAmount = CDbl(iDuration) * sRate
				bAddFee = True
			End If 
		End If 

		If bAddFee Then 
			' The internal reservations are always $0.00
			If sReservationTypeSelection = "admin" Then 
				iFeeAmount = CDbl(0.00)
				sRate = CDbl(0.00) 
			End If 

			dTotalAmount = dTotalAmount + CDbl(iFeeAmount)

			sSql = "INSERT INTO egov_rentalreservationdatefees (reservationdateid, reservationid, rentalid, pricetypeid, "
			sSql = sSql & "accountid, ratetypeid, amount, starthour, startminute, startampm, feeamount, paidamount, "
			sSql = sSql & "refundamount, duration, datediffstring ) VALUES ( " & iReservationDateId & ", " & iReservationId & ", " & iRentalid & ", "
			sSql = sSql & oRs("pricetypeid") & ", " & sAccount & ", " & oRs("ratetypeid") & ", " & sRate & ", "
			sSql = sSql & sHour & ", " & sMinute & ", " & sAmPm & ", " & iFeeAmount & ", 0.0000, 0.0000, "
			sSql = sSql & iDuration & ", '" & oRs("datediffstring") & "' )"
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql
		End If 

		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void CreateRentalReservationDateItems iReservationDateId, iReservationId, iRentalid, sReservationTypeSelection
'--------------------------------------------------------------------------------------------------
Sub CreateRentalReservationDateItems( ByVal iReservationDateId, ByVal iReservationId, ByVal iRentalid, ByVal sReservationTypeSelection )
	Dim sSql, oRs, sAccount, sAmount

	' Add the Rental Reservation Fees
	sSql = "SELECT rentalitemid, ISNULL(rentalitem,'') AS rentalitem, ISNULL(accountid,0) AS accountid, "
	sSql = sSql & "ISNULL(maxavailable,0) AS maxavailable, ISNULL(amount,0.00) AS amount "
	sSql = sSql & "FROM egov_rentalitems WHERE rentalid = " & iRentalid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If oRs("accountid") = CLng(0) Then 
			sAccount = "NULL"
		Else
			sAccount = oRs("accountid")
		End If 
		If sReservationTypeSelection = "public" Then
			sAmount = CDbl(oRs("amount"))
		Else
			sAmount = CDbl(0.00)
		End If 

		sSql = "INSERT INTO egov_rentalreservationdateitems ( reservationdateid, reservationid, rentalid, "
		sSql = sSql & "rentalitemid, rentalitem, accountid, maxavailable, amount, quantity, feeamount, paidamount, refundamount ) "
		sSql = sSql & " VALUES ( " & iReservationDateId & ", " & iReservationId & ", " & iRentalid & ", " & oRs("rentalitemid") & ", '"
		sSql = sSql & dbsafe(oRs("rentalitem")) &"', " & sAccount & ", " & oRs("maxavailable") & ", " & sAmount & ", "
		sSql = sSql & "0, 0.00, 0.00, 0.00 )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void CreateRentalReservationFees iReservationId, iRentalid, dTotalAmount
'--------------------------------------------------------------------------------------------------
Sub CreateRentalReservationFees( ByVal iReservationId, ByVal iRentalid, ByRef dTotalAmount )
	Dim sSql, oRs, sAccount, sPrompt, dFeeAmount

	' Add the Rental Reservation Fees
	sSql = "SELECT F.pricetypeid, ISNULL(accountid,0) AS accountid, ISNULL(amount,0.00) AS amount, ISNULL(prompt,'') AS prompt, P.isoptional "
	sSql = sSql & " FROM egov_rentalfees F, egov_price_types P WHERE F.pricetypeid = P.pricetypeid AND rentalid = " & iRentalid
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If oRs("accountid") = CLng(0) Then 
			sAccount = "NULL"
		Else
			sAccount = oRs("accountid")
		End If 
		If oRs("prompt") = "" Then 
			sPrompt = "NULL"
		Else
			sPrompt = "'" & dbsafe(oRs("prompt")) & "'"
		End If 
		If oRs("isoptional") Then 
			dFeeAmount = "0.0000"
		Else
			dFeeAmount = CDbl(oRs("amount"))
			dTotalAmount = dTotalAmount + dFeeAmount
		End If 

		sSql = "INSERT INTO egov_rentalreservationfees ( reservationid, rentalid, pricetypeid, amount, accountid, feeamount, prompt, paidamount, refundamount ) "
		sSql = sSql & " VALUES ( " & iReservationId & ", " & iRentalid & ", " & oRs("pricetypeid") & ", "
		sSql = sSql & CDbl(oRs("amount")) & ", " & sAccount & ", " & dFeeAmount & ", " &  sPrompt & ", 0.00, 0.00 )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Double displayReservationRefunds( iReservationId, sClass )
'--------------------------------------------------------------------------------------------------
Function displayReservationRefunds( ByVal iReservationId, ByVal sClass )
	Dim sSql, oRs, dRefundTotal

	dRefundTotal = CDbl(0.00)
	
	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) as refundamount "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_payment J "
	sSql = sSql & " WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 0 AND A.entrytype = 'debit' "
	sSql = sSql & " AND A.reservationid = " & iReservationId
	sSql = sSql & " GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"" colspan=""2"">Date</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td>&nbsp;" & oRs("paymentid") & "</td>"
			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td align=""right"">"
			response.write FormatNumber(oRs("refundamount"),2,,,0) 
			' changed to sum the refund amounts. 11/9/2012 SJL
			dRefundTotal = dRefundTotal + CDbl(oRs("refundamount"))
			response.write "</td></tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	displayReservationRefunds = dRefundTotal

End Function  


'--------------------------------------------------------------------------------------------------
' Boolean EndTimeIsNotClosingTime( iRentalId, dEndDateTime, bOffSeasonFlag, dWantedStartTime )
'--------------------------------------------------------------------------------------------------
Function EndTimeIsNotClosingTime( ByVal iRentalId, ByVal dEndDateTime, ByVal bOffSeasonFlag, ByVal dStartDateTime )
	Dim sSql, oRs, iWeekday, dClosingTime

	iWeekday = Weekday(CDate(dStartDateTime))

	sSql = "SELECT ISNULL(closinghour,0) AS closinghour, dbo.AddLeadingZeros(ISNULL(closingminute,0),2) AS closingminute, "
	sSql = sSql & " ISNULL(closingampm,'AM') AS closingampm, ISNULL(closingday,0) AS closingday "
	sSql = sSql & " FROM egov_rentaldays "
	sSql = sSql & " WHERE rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("closingday")) = CLng(0) Then
			dClosingTime = CDate(DateValue(dStartDateTime) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		Else
			dClosingTime = CDate(DateAdd("d", 1, DateValue(dStartDateTime)) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		End If 
		If dClosingTime = CDate(dEndDateTime) Then
			' It is closing time
			EndTimeIsNotClosingTime = False  
		Else
			EndTimeIsNotClosingTime = True 
		End If 
	Else
		EndTimeIsNotClosingTime = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string FormatTimeString( dDateTimeString )
'--------------------------------------------------------------------------------------------------
Function FormatTimeString( ByVal dDateTimeString )
	Dim sAmPm, sHour, sMinute

	sAmPm = "AM"
	sHour = Hour(dDateTimeString)
	If clng(sHour) = clng(0) Then
		sHour = 12
		sAmPm = "AM"
	Else
		If clng(sHour) > clng(12) Then
			sHour = clng(sHour) - clng(12)
			sAmPm = "PM"
		End If 
		If clng(sHour) = clng(12) Then
			sAmPm = "PM"
		End If 
	End If 
	sMinute = Minute(dDateTimeString)
	If sMinute < 10 Then
		sMinute = "0" & sMinute
	End If 

	FormatTimeString = sHour & ":" & sMinute & " " & sAmPm

End Function 


'--------------------------------------------------------------------------------------------------
' string GetAccountName( iAccountId )
'--------------------------------------------------------------------------------------------------
Function GetAccountName( ByVal iAccountId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountname,'') AS accountname FROM egov_accounts "
	sSql = sSql & "WHERE accountid = " & iAccountId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAccountName = oRs("accountname")
	Else
		GetAccountName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetActivityNo( iTimeId )
'--------------------------------------------------------------------------------------------------
Function GetActivityNo( ByVal iTimeId )
	Dim sSql, oRs 

	sSql = "SELECT activityno FROM egov_class_time WHERE timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetActivityNo = oRs("activityno")
	Else
		GetActivityNo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' void GetAdminNameAndPhone iUserId, sRenterFirstname, sRenterLastName, sRenterPhone
'------------------------------------------------------------------------------------------------------------
Sub GetAdminNameAndPhone( ByVal iUserId, ByRef sRenterFirstname, ByRef sRenterLastName, ByRef sRenterPhone )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(FirstName,'') AS FirstName, ISNULL(LastName,'') AS LastName, "
	sSql = sSql & " ISNULL(BusinessNumber,'') AS BusinessNumber "
	sSql = sSql & " FROM users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sRenterFirstname = oRs("FirstName")
		sRenterLastName = oRs("LastName")
		sRenterPhone = FormatPhoneNumber(oRs("BusinessNumber"))
	Else
		sRenterFirstname = ""
		sRenterLastName = ""
		sRenterPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' String GetAdminPhone( iAdminUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminPhone( ByVal iAdminUserId )
	Dim oRs, sSql

	sSql = "SELECT businessnumber FROM users WHERE userid = " & iAdminUserId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAdminPhone = oRs("BusinessNumber") ' Should be formatted already
	Else
		GetAdminPhone = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetAllDayHours iRentalid, bOffSeasonFlag, iWeekDay, dWantedStartTime, dWantedEndTime
'--------------------------------------------------------------------------------------------------
Sub GetAllDayHours( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekDay, ByRef dWantedStartTime, ByRef dWantedEndTime )
	Dim sSql, oRs, iIsOffSeason

	If bOffSeasonFlag Then 
		iIsOffSeason = 1
	Else
		iIsOffSeason = 0
	End If 

	sSql = "SELECT openinghour, dbo.AddLeadingZeros(openingminute,2) AS openingminute, openingampm, "
	sSql = sSql & " closinghour, dbo.AddLeadingZeros(closingminute,2) AS closingminute, closingampm, closingday "
	sSql = sSql & " FROM egov_rentaldays WHERE rentalid = " & iRentalid & " AND isoffseason = " & iIsOffSeason
	sSql = sSql & " AND dayofweek = " & iWeekDay

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' Set the start time to the opening time
		dWantedStartTime = CDate(DateValue(dWantedStartTime) & " " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm"))
		
		' Set the end time to the closing time 
		If CLng(oRs("closingday")) = CLng(0) Then
			dWantedEndTime = CDate(DateValue(dWantedStartTime) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		Else
			dWantedEndTime = CDate(DateAdd("d", 1, DateValue(dWantedStartTime)) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetCategoryTitle( iCategoryId )
'--------------------------------------------------------------------------------------------------
Function GetCategoryTitle( ByVal iCategoryId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(categorytitle,'') AS categorytitle "
	sSql = sSql & "FROM egov_recreation_categories "
	sSql = sSql & "WHERE recreationcategoryid = " & iCategoryId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetCategoryTitle = oRs("categorytitle")
	Else
		GetCategoryTitle = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' void GetCitizenNameAndPhone iUserId, sRenterFirstname, sRenterLastName, sRenterPhone
'------------------------------------------------------------------------------------------------------------
Sub GetCitizenNameAndPhone( ByVal iUserId, ByRef sRenterFirstname, ByRef sRenterLastName, ByRef sRenterPhone )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sRenterFirstname = oRs("userfname")
		sRenterLastName = oRs("userlname")
		sRenterPhone = FormatPhoneNumber(oRs("userhomephone"))
	Else
		sRenterFirstname = ""
		sRenterLastName = ""
		sRenterPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' String GetCitizenPhone( iCitizenUserId )
'--------------------------------------------------------------------------------------------------
Function GetCitizenPhone( ByVal iCitizenUserId )
	Dim oRs, sSql 

	sSql = "SELECT userhomephone FROM egov_users WHERE userid = " & iCitizenUserId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCitizenPhone = FormatPhoneNumber(oRs("userhomephone")) ' In common.asp
	Else
		GetCitizenPhone = ""
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' void GetClassName( iTimeId )
'--------------------------------------------------------------------------------------------------
Function GetClassName( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT C.classname FROM egov_class C, egov_class_time T "
	sSql = sSql & "WHERE C.classid = T.classid AND T.timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetClassName = oRs("classname")
	Else
		GetClassName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetCurrentPaidAmount( iKeyFieldValue, sKeyField, sTable )	
'--------------------------------------------------------------------------------------------------
Function GetCurrentPaidAmount( ByVal iKeyFieldValue, ByVal sKeyField, ByVal sTable )	
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(paidamount,0.0000) AS paidamount FROM " & sTable & " WHERE " & sKeyField & " = " & iKeyFieldValue

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCurrentPaidAmount = CDbl(oRs("paidamount"))
	Else
		GetCurrentPaidAmount = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End function


'--------------------------------------------------------------------------------------------------
' double GetCurrentRefundAmount( iKeyFieldValue, sKeyField, sTable )	
'--------------------------------------------------------------------------------------------------
Function GetCurrentRefundAmount( ByVal iKeyFieldValue, ByVal sKeyField, ByVal sTable )	
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(refundamount,0.0000) AS refundamount FROM " & sTable & " WHERE " & sKeyField & " = " & iKeyFieldValue

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCurrentRefundAmount = CDbl(oRs("refundamount"))
	Else
		GetCurrentRefundAmount = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'--------------------------------------------------------------------------------------------------
' void GetGeneralReservationData iReservationId
'--------------------------------------------------------------------------------------------------
Sub GetGeneralReservationData( ByVal iReservationId )
	Dim sSql, oRs
	
	sSql = "SELECT R.reservationtypeid, R.reserveddate, R.servingalcohol, ISNULL(R.organization,'') AS organization, "
	sSql = sSql & " ISNULL(R.pointofcontact,'') AS pointofcontact, ISNULL(R.numberattending,'') AS numberattending, "
	sSql = sSql & " ISNULL(R.purpose,'') AS purpose, ISNULL(R.receiptnotes,'') AS receiptnotes, S.iscancelled, ISNULL(R.timeid,0) AS timeid, "
	sSql = sSql & " ISNULL(R.privatenotes,'') AS privatenotes, T.reservationtype, T.reservationtypeselector, T.isreservation, "
	sSql = sSql & " S.reservationstatus, ISNULL(R.rentaluserid,0) AS rentaluserid, ISNULL(R.adminuserid,0) AS adminuserid, R.isonhold, R.iscall, "
	sSql = sSql & " CAST(ISNULL(u.facilityabuse,0) as bit) as facilityabuse, u.facilityabusenote "
	sSql = sSql & " FROM egov_rentalreservations R "
	sSql = sSql & " INNER JOIN egov_rentalreservationtypes T ON R.reservationtypeid = T.reservationtypeid "
	sSql = sSql & " INNER JOIN egov_rentalreservationstatuses S ON R.reservationstatusid = S.reservationstatusid "
	sSql = sSql & " LEFT JOIN egov_users u ON u.userid = r.rentaluserid "
	sSql = sSql & " WHERE R.orgid = " & session("OrgId") & " AND R.reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iReservationTypeId = oRs("reservationtypeid")
		sReservationType = oRs("reservationtype")
		sReservationStatus = oRs("reservationstatus")
		sOrganization = oRs("organization")
		sPointOfContact = oRs("pointofcontact")
		sNumberAttending = oRs("numberattending")
		sPurpose = oRs("purpose")
		sReceiptNotes = oRs("receiptnotes")
		sPrivateNotes = oRs("privatenotes")
		sReservedDate = DateValue(oRs("reserveddate"))
		sReservationTypeSelector = oRs("reservationtypeselector")
		iRentalUserId = oRs("rentaluserid")
		iTimeId = oRs("timeid")
		bFacilityAbuse = oRs("facilityabuse")
		sFacilityAbuseNote = oRs("facilityabusenote")
		If oRs("isreservation") Then 
			bIsReservation = True 
			If LCase(sReservationTypeSelector) = "admin" Then
				sRenterName = GetAdminName( oRs("rentaluserid") )
				sRenterPhone = GetAdminPhone( oRs("rentaluserid") )
			Else
				sRenterName = GetCitizenName( oRs("rentaluserid") )
				sRenterPhone = GetCitizenPhone( oRs("rentaluserid") )
			End If 
		Else
			bIsReservation = False 
			sRenterName = ""
			sRenterPhone = ""
		End If 
		If CLng(oRs("adminuserid")) > CLng(0) Then 
			sAdminName = GetAdminName( oRs("adminuserid") )
		Else
			' This is a public side reservation
			sAdminName = sRenterName
		End If 
		If oRs("servingalcohol") Then
			sServingAlcohol = " checked=""checked"" "
		Else
			sServingAlcohol = ""
		End If 
		If oRs("iscancelled") then
			bReservationIsCancelled = True 
		Else
			bReservationIsCancelled = False 
		End If 
		If oRs("isonhold") Then
			sHoldFlag = " checked=""checked"" "
		Else
			sHoldFlag = ""
		End If 
		If oRs("iscall") Then
			sCallFlag = " checked=""checked"" "
		Else
			sCallFlag = ""
		End If
	Else
		iReservationTypeId = 0
		sReservationType = ""
		sReservationStatus = ""
		sOrganization = ""
		sPointOfContact = ""
		sNumberAttending = ""
		sPurpose = ""
		sReceiptNotes = ""
		sPrivateNotes = ""
		sReservedDate = ""
		sReservationTypeSelector = ""
		sRenterName = ""
		sRenterPhone = ""
		sAdminName = ""
		sServingAlcohol = ""
		bReservationIsCancelled = True 
		iRentalUserId = 0
		iTimeId = 0
		bIsReservation = False 
		sCallFlag = ""
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' Integer GetHourFromDateTime( dDateTime, sAmPm )
'--------------------------------------------------------------------------------------------------
Function GetHourFromDateTime( ByVal dDateTime, ByRef sAmPm )
	Dim iHour

	session("dDateTime") = dDateTime
	sAmPm = "AM"
	iHour = clng(Hour(dDateTime))
	
	If iHour = clng(0) Then
		iHour = clng(12)
	ElseIf iHour = clng(12) Then
		sAmPm = "PM"
	ElseIf iHour > clng(12) Then
		iHour = iHour - clng(12)
		sAmPm = "PM"
	End If 
	session("dDateTime") = ""

	GetHourFromDateTime = iHour

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetInitialReservationStatusId()
'--------------------------------------------------------------------------------------------------
Function GetInitialReservationStatusId()
	Dim sSql, oRs

	' Get the initial status to make a reservation so we do not hard code this
	sSql = "SELECT reservationstatusid FROM egov_rentalreservationstatuses "
	sSql = sSql & " WHERE isinitialstatus = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetInitialReservationStatusId = oRs("reservationstatusid")
	Else
		GetInitialReservationStatusId = 0	' This would be a problem
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetItemTypeId( sItemType )
'------------------------------------------------------------------------------
Function GetItemTypeId( ByVal sItemType )
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sItemType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetItemTypeId = CLng(oRs("itemtypeid"))
	Else
		GetItemTypeId = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetJournalEntryTypeID( sType )
'------------------------------------------------------------------------------
Function GetJournalEntryTypeID( ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT journalentrytypeid FROM egov_journal_entry_types WHERE journalentrytype = '" & sType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetJournalEntryTypeID = oRs("journalentrytypeid") 
	Else 
		GetJournalEntryTypeID = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' date GetLastReservationDate( iReservationId ) 
'--------------------------------------------------------------------------------------------------
Function GetLastReservationDate( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT MAX(reservationstarttime) AS reservationstarttime "
	sSql = sSql & "FROM egov_rentalreservationdates WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetLastReservationDate = oRs("reservationstarttime") 
	Else 
		GetLastReservationDate = Now()
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' datetime GetLatestReservationTime( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
'--------------------------------------------------------------------------------------------------
Function GetLatestReservationTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dStartDate )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(lateststarthour,0) AS lateststarthour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(lateststartminute,0),2) AS lateststartminute, "
	sSql = sSql & " ISNULL(lateststartampm,'AM') AS lateststartampm "
	sSql = sSql & " FROM egov_rentaldays "
	sSql = sSql & " WHERE rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag
	sSql = sSql & " AND dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then	
		' there will always be a latest starting time and it will be on the startdate before midnight
		GetLatestReservationTime = CDate(DateValue(dStartDate) & " " & oRs("lateststarthour") & ":" & oRs("lateststartminute") & " " & oRs("lateststartampm"))
	Else
		' this would be a problem, so send back the start time
		GetLatestReservationTime = dStartDate
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetLocationName( iLocationId )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationId )
	Dim sSql, oRs

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetLocationName = oRs("name") 
	Else 
		GetLocationName = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean GetMinimalTimeInfo( iRentalid, bOffSeasonFlag, iWeekday, iInterval, sDateAddString, bIsAllDay )
'--------------------------------------------------------------------------------------------------
Function GetMinimalTimeInfo( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByRef iInterval, ByRef sDateAddString, ByRef bIsAllDay )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(D.minimumrental,0) AS minimumrental, T.dateaddstring, T.isallday "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes T "
	sSql = sSql & " WHERE D.minimumrentaltimetypeid = T.timetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isallday") Then
			bIsAllDay = True 
			iInterval = 1
			sDateAddString = "d"
		Else
			bIsAllDay = False 
			iInterval = clng(oRs("minimumrental"))
			sDateAddString = oRs("dateaddstring")
		End If 
		GetMinimalTimeInfo = True 
	Else
		bIsAllDay = False 
		iInterval = 0
		sDateAddString = "h"
		GetMinimalTimeInfo = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' date GetNextOrdinalDayMonth( iWeekDay, iPos, iMonth, iYear ) - by JOHN STULLENBERGER
'--------------------------------------------------------------------------------------------------
Function GetNextOrdinalDayMonth( ByVal iWeekDay, ByVal iPos, ByVal iMonth, ByVal iYear )
	Dim iCount, dTemp, dReturnValue

	' This is from the facilities feature and was created by John Stullenberger

	' INITIALIZE DATE VALUES
	dTemp = CDate(iMonth & "/1/" & iYear)
	dReturnValue = dTemp
	iCount = 0

	' PERFORM DATE LOOKUP BASED ON ORDINAL POSITION
	Select Case iPos
		Case 1
			' FIRST OCCURRENCE
			 Do While Not  clng(WeekDay(dTemp)) = clng(iWeekDay) 
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d", 1, dTemp)
			 Loop
			 dReturnValue = dTemp

		Case 2
			' SECOND OCCURRENCE
			 Do While iCount < 2
				' FOUND DAY OF WEEK MATCH
				If clng(WeekDay(dTemp) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d", 1, dTemp)
			 Loop

		Case 3
			' THIRD OCCURRENCE
			 Do While clng(iCount) < clng(3)
				' FOUND DAY OF WEEK MATCH
				If clng(WeekDay(dTemp) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d", 1, dTemp)
			 Loop

		Case 4
			' FOURTH OCCURRENCE
			 Do While iCount < 4
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d", 1, dTemp)
			 Loop

		Case 5
			datNextMonth = dateAdd("m", 1, dTemp)
			' LAST OCCURRENCE
			 Do While iCount < 5 AND (CDate(dtemp) < CDate(datNextMonth))
				' FOUND DAY OF WEEK MATCH
				If (clng(WeekDay(dTemp)) = clng(iWeekDay)) Then
					iCount = iCount + 1 ' ADD 1 TO OCCURENCE COUNT
					dReturnValue = dTemp
				End If
				' ADD 1 DAY TO CURRENT DATE
				dTemp = DateAdd("d", 1, dTemp)
			 Loop
			
	End Select

	' RETURN DATE VALUE
	GetNextOrdinalDayMonth = dReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' string GetOffSeasonFlag( iRentalid, dCheckDate )
'--------------------------------------------------------------------------------------------------
Function GetOffSeasonFlag( ByVal iRentalid, ByVal dCheckDate )
	Dim sSql, oRs, dOffSeasonStartDate, dOffSeasonEndDate

	' See if we have an off season and if so is the date in the off season.
	' You want to return a flag indicating which season (in, off) to use.

	sSql = "SELECT hasoffseason, ISNULL(offseasonstartmonth, 1) AS offseasonstartmonth, "
	sSql = sSql & "ISNULL(offseasonstartday,1) AS offseasonstartday, ISNULL(offseasonendmonth,1) AS offseasonendmonth, "
	sSql = sSql & "ISNULL(offseasonendday,1) AS offseasonendday, ISNULL(offseasonendyear,1) AS offseasonendyear "
	sSql = sSql & "FROM egov_rentals WHERE orgid = " & session("orgid") & " AND rentalid = " & iRentalid
	'Response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		dOffSeasonStartDate = CDate(oRs("offseasonstartmonth") & "/" & oRs("offseasonstartday") & "/" & Year(dCheckDate))
		dOffSeasonEndDate = CDate(oRs("offseasonendmonth") & "/" & oRs("offseasonendday") & "/" & Year(dCheckDate))
		If oRs("hasoffseason") Then
			If clng(oRs("offseasonendyear")) = clng(1) Then
				' the start and end are in different calendar years - ouch!
				If CDate("1/1/" & Year(dCheckDate)) <= dCheckDate And dOffSeasonEndDate > dCheckDate Then
					' This is in one of the off season periods
					GetOffSeasonFlag = "1"
				Else
					If CDate("12/31/" & Year(dCheckDate)) >= dCheckDate And dOffSeasonStartDate <= dCheckDate Then
						' This is in the other off season period
						GetOffSeasonFlag = "1"
					Else
						' This is the in season
						GetOffSeasonFlag = "0"
					End If 
				End If 
			Else
				' the start and end are in the same calendar year
				If dOffSeasonStartDate <= dCheckDate AND dOffSeasonEndDate > dCheckDate Then
					' This is in the off season period
					GetOffSeasonFlag = "1"
				Else
					' This is the in season
					GetOffSeasonFlag = "0"
				End If 
			End If 
		Else
			' No off season set up
			GetOffSeasonFlag = "0"
		End If 
	Else
		' could not find the rental- this would be a problem
		GetOffSeasonFlag = "0"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetOpeningAndClosingTimes iRentalid, bOffSeasonFlag, iWeekday, dStartDate, sStartTime, sEndTime
'--------------------------------------------------------------------------------------------------
Sub  GetOpeningAndClosingTimes( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dStartDate, ByRef sStartTime, ByRef sEndTime )
	Dim sSql, oRs 
	
	sSql = "SELECT isopen, isavailabletopublic, ISNULL(openinghour,0) AS openinghour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(openingminute,0),2) AS openingminute, ISNULL(openingampm,'AM') AS openingampm, "
	sSql = sSql & " ISNULL(closinghour,0) AS closinghour, dbo.AddLeadingZeros(ISNULL(closingminute,0),2) AS closingminute, "
	sSql = sSql & " ISNULL(closingampm,'AM') AS closingampm, ISNULL(closingday,0) AS closingday "
	sSql = sSql & " FROM egov_rentaldays "
	sSql = sSql & " WHERE rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isopen") Then 
			sStartTime = CDate(DateValue(dStartDate) & " " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm"))
			If oRs("closingday") = "0" Then 
				sEndTime = dStartDate
			Else
				sEndTime = DateAdd("d", 1, dStartDate)
			End If 
			sEndTime = CDate(DateValue(sEndTime) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' void GetPostBufferTime iRentalid, bOffSeasonFlag, iWeekday, iPostBuffer, sDateAddString
'--------------------------------------------------------------------------------------------------
Sub GetPostBufferTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByRef iPostBuffer, ByRef sDateAddString )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(postbuffer,0) AS postbuffer, P.dateaddstring "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes P "
	sSql = sSql & " WHERE D.postbuffertimetypeid = P.timetypeid "
	sSql = sSql & " AND rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag
	sSql = sSql & " AND dayofweek = " & iWeekday

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("postbuffer")) > CLng(0) Then 
			iPostBuffer = oRs("postbuffer")
			sDateAddString = oRs("dateaddstring")
		Else
			iPostBuffer = 0
			sDateAddString = "h"
		End If 
	Else
		iPostBuffer = 0
		sDateAddString = "h"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub   


'--------------------------------------------------------------------------------------------------
' string GetRentalHours( iRentalId, dCheckDate )
'--------------------------------------------------------------------------------------------------
Function GetRentalHours( ByVal iRentalId, ByVal dCheckDate )
	Dim sSql, oRs, bOffSeasonFlag, iWeekday, sRentalHours

	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, dCheckDate )
	iWeekday = Weekday(dCheckDate)

	sSql = "SELECT D.dayid, isopen, isavailabletopublic, openinghour, dbo.AddLeadingZeros(openingminute,2) AS openingminute, "
	sSql = sSql & " openingampm, closinghour, dbo.AddLeadingZeros(closingminute,2) AS closingminute, closingampm, closingday, "
	sSql = sSql & " ISNULL(lateststarthour,0) AS lateststarthour, dbo.AddLeadingZeros(lateststartminute,2) AS lateststartminute, "
	sSql = sSql & " lateststartampm, ISNULL(minimumrental,0) AS minimumrental, M.timetype, ISNULL(postbuffer,0) AS postbuffer, "
	sSql = sSql & " B.timetype AS postbuffertimetype, M.isallday "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes M, egov_rentaltimetypes B "
	sSql = sSql & " WHERE D.minimumrentaltimetypeid = M.timetypeid AND D.postbuffertimetypeid = B.timetypeid AND rentalid = " & iRentalid 
	sSql = sSql & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("isopen") Then
			' Open and closing hours
			sRentalHours = "<br />Open: " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm") & " to "
			sRentalHours = sRentalHours & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm")

			' Last Reservation time
			If clng(oRs("lateststarthour")) > clng(0) Then 
				sRentalHours = sRentalHours & "<br />Latest Reservation: " & oRs("lateststarthour") & ":" & oRs("lateststartminute") & " " & oRs("lateststartampm")
			End If 

			' Find any weekend surcharge Start times
			sWeekendSurchargeStart = GetWeekendSurchargeStart( oRs("dayid") )
			If sWeekendSurchargeStart <> "" Then
				sRentalHours = sRentalHours & "<br />Weekend Surcharge Starts: " & sWeekendSurchargeStart
			End If 

			' Minimum Rental time 
			If clng(oRs("minimumrental")) > clng(0) Then
				sRentalHours = sRentalHours & "<br />Minimum Time: " & oRs("minimumrental") & " " & oRs("timetype")
			Else
				If oRs("isallday") Then 
					sRentalHours = sRentalHours & "<br />Minimum Time: " & oRs("timetype")
				End If 
			End If 

			' Post Buffer
			If clng(oRs("postbuffer")) > clng(0) Then
				sRentalHours = sRentalHours & "<br />End Buffer: " & oRs("postbuffer") & " " & oRs("postbuffertimetype")
			End If 
		Else
			sRentalHours = "<br />Closed on " & WeekDayName(iWeekday) & "s"
		End If 
	Else
		sRentalHours = "<br />Closed on " & WeekDayName(iWeekday) & "s"
	End If
	
	oRs.Close
	Set oRs = Nothing 

	GetRentalHours = sRentalHours

End Function 


'--------------------------------------------------------------------------------------------------
' string GetRentalName( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetRentalName( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT rentalname FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRentalName = oRs("rentalname")
	Else
		GetRentalName = ""
	End If
	
	oRs.Close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' integer GetRentalPaymentId( iReservationId )
'--------------------------------------------------------------------------------------------------
Function GetRentalPaymentId( ByVal iReservationId )
	Dim sSql, oRs

	' Pulls the first paymentid of a reservation. Used when adding dates to a reservation
	sSql = "SELECT paymentid FROM egov_class_payment "
	sSql = sSql & "WHERE isforrentals = 1 AND reservationid = " & iReservationId
	sSql = sSql & " ORDER BY paymentid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRentalPaymentId = oRs("paymentid")
	Else
		GetRentalPaymentId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetRentalPaymentTypeId( iOrgid, sType )
'--------------------------------------------------------------------------------------------------
Function GetRentalPaymentTypeId( ByVal iOrgid, ByVal sType )
	Dim sSql, oRs

	sSql = "SELECT P.paymenttypeid FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & "WHERE P.paymenttypeid = O.paymenttypeid AND P.isforrentals = 1 AND P." & sType & " = 1 "
	sSql = sSql & "AND O.orgid = " & iOrgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRentalPaymentTypeId = oRs("paymenttypeid")
	Else
		GetRentalPaymentTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetRentalSeason( iRentalId, dCheckDate )
'--------------------------------------------------------------------------------------------------
Function GetRentalSeason( ByVal iRentalId, ByVal dCheckDate )
	Dim sSql, oRs, dOffSeasonStartDate, dOffSeasonEndDate

	' See if we have an off season and if so is the date in the off season.
	sSql = "SELECT hasoffseason, ISNULL(offseasonstartmonth,1) AS offseasonstartmonth, "
	sSql = sSql & "ISNULL(offseasonstartday,1) AS offseasonstartday, ISNULL(offseasonendmonth,1) AS offseasonendmonth, "
	sSql = sSql & "ISNULL(offseasonendday,1) AS offseasonendday, ISNULL(offseasonendyear,1) AS offseasonendyear "
	sSql = sSql & "FROM egov_rentals WHERE orgid = " & session("orgid") & " AND rentalid = " & iRentalid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		dOffSeasonStartDate = CDate(oRs("offseasonstartmonth") & "/" & oRs("offseasonstartday") & "/" & Year(dCheckDate))
		dOffSeasonEndDate = CDate(oRs("offseasonendmonth") & "/" & oRs("offseasonendday") & "/" & Year(dCheckDate))
		If oRs("hasoffseason") Then
			If clng(oRs("offseasonendyear")) = clng(1) Then
				' the start and end are in different calendar years - ouch!
				If CDate("1/1/" & Year(dCheckDate)) <= dCheckDate And dOffSeasonEndDate >= dCheckDate Then
					' This is in one of the off season periods
					GetRentalSeason = "Off Season"
				Else
					If CDate("12/31/" & Year(dCheckDate)) >= dCheckDate And dOffSeasonStartDate <= dCheckDate Then
						' This is in the other off season period
						GetRentalSeason = "Off Season"
					Else
						' This is the in season
						GetRentalSeason = "In Season"
					End If 
				End If 
			Else
				' the start and end are in the same calendar year
				If dOffSeasonStartDate <= dCheckDate AND dOffSeasonEndDate >= dCheckDate Then
					' This is in the off season period
					GetRentalSeason = "Off Season"
				Else
					' This is the in season
					GetRentalSeason = "In Season"
				End If 
			End If 
		Else
			' No off season set up
			GetRentalSeason = "In Season"
		End If 
	Else
		' could not find the rental- this would be a problem
		GetRentalSeason = "In Season"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetRentalShortDescription( iRentalId )
'------------------------------------------------------------------------------
Function GetRentalShortDescription( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(shortdescription,'') AS shortdescription FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRentalShortDescription = oRs("shortdescription")
	Else
		GetRentalShortDescription = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetRentalSupervisorPhone( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetRentalSupervisorPhone( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(U.businessnumber,'') AS businessnumber "
	sSql = sSql & " FROM egov_rentals R, Users U "
	sSql = sSql & " WHERE R.supervisoruserid = U.UserID "
	sSql = sSql & " AND R.rentalid = " & iRentalId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRentalSupervisorPhone = oRs("businessnumber")
	Else
		GetRentalSupervisorPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetRenterName( iReservationId )
'--------------------------------------------------------------------------------------------------
Function GetRenterName( ByVal iReservationId )
	Dim sSql, oRs, sRenterName

	sSql = "SELECT T.isreservation,  T.reservationtypeselector, ISNULL(R.rentaluserid,0) AS rentaluserid "
	sSql = sSql & " FROM egov_rentalreservations R, egov_rentalreservationtypes T, egov_rentalreservationstatuses S "
	sSql = sSql & " WHERE R.reservationtypeid = T.reservationtypeid AND R.reservationstatusid = S.reservationstatusid "
	sSql = sSql & " AND R.orgid = " & session("OrgId") & " AND R.reservationid = " & iReservationId

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isreservation") Then 
			If LCase( oRs("reservationtypeselector")) = "admin" Then
				sRenterName = GetAdminName( oRs("rentaluserid") )
			Else
				sRenterName = GetCitizenName( oRs("rentaluserid") )
			End If 
		Else
			sRenterName = ""
		End If 
	Else
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetRenterName = sRenterName

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetReservationAccountId( iKeyFieldValue, sKeyField, sTable )
'--------------------------------------------------------------------------------------------------
Function GetReservationAccountId( ByVal iKeyFieldValue, ByVal sKeyField, ByVal sTable )
	Dim sSql, oRs 

	sSql = "SELECT ISNULL(accountid,0) AS accountid FROM " & sTable & " WHERE " & sKeyField & " = " & iKeyFieldValue
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationAccountId = oRs("accountid")
	Else
		GetReservationAccountId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetReservationDateFees( iReservationDateId, sField )
'--------------------------------------------------------------------------------------------------
Function GetReservationDateFees( ByVal iReservationDateId, ByVal sField )
	Dim oRs, sSql

	' Get the reservation fees for a specific date
	sSql = "SELECT ISNULL(SUM(" & sField & "),0.0000) AS someamount "
	sSql = sSql & " FROM egov_rentalreservationdatefees"
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationDateFees = CDbl(oRs("someamount"))
	Else
		GetReservationDateFees = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetReservationDateFeesTotal( iReservationId, sField )
'--------------------------------------------------------------------------------------------------
Function GetReservationDateFeesTotal( ByVal iReservationId, ByVal sField )
	Dim oRs, sSql

	' Get the total reservation fees
	sSql = "SELECT ISNULL(SUM(" & sField & "),0.0000) AS someamount "
	sSql = sSql & " FROM egov_rentalreservationdatefees"
	sSql = sSql & " WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationDateFeesTotal = CDbl(oRs("someamount"))
	Else
		GetReservationDateFeesTotal = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetReservationDateId( iKeyFieldValue, sKeyField, sTable )
'--------------------------------------------------------------------------------------------------
Function GetReservationDateId( ByVal iKeyFieldValue, ByVal sKeyField, ByVal sTable )
	Dim sSql, oRs 

	sSql = "SELECT reservationdateid FROM " & sTable & " WHERE " & sKeyField & " = " & iKeyFieldValue
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationDateId = oRs("reservationdateid")
	Else
		GetReservationDateId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' double GetReservationDateItemFees( iReservationDateId, sField )
'--------------------------------------------------------------------------------------------------
Function GetReservationDateItemFees( ByVal iReservationDateId, ByVal sField )
	Dim oRs, sSql

	' Get the total reservation date item fees
	sSql = "SELECT ISNULL(SUM(" & sField & "),0.0000) AS someamount "
	sSql = sSql & " FROM egov_rentalreservationdateitems"
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationDateItemFees = CDbl(oRs("someamount"))
	Else
		GetReservationDateItemFees = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetReservationDateItemFeesTotal( iReservationId, sField )
'--------------------------------------------------------------------------------------------------
Function GetReservationDateItemFeesTotal( ByVal iReservationId, ByVal sField )
	Dim oRs, sSql

	' Get the total reservation date item fees
	sSql = "SELECT ISNULL(SUM(" & sField & "),0.0000) AS someamount "
	sSql = sSql & " FROM egov_rentalreservationdateitems"
	sSql = sSql & " WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationDateItemFeesTotal = CDbl(oRs("someamount"))
	Else
		GetReservationDateItemFeesTotal = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetReservationDateItemKeyValues iReservationDateItemId, iReservationDateId, iPaymentId, iAccountId
'--------------------------------------------------------------------------------------------------
Sub GetReservationDateItemKeyValues( ByVal iReservationDateItemId, ByRef iReservationDateId, ByRef iPaymentId, ByRef iAccountId )
	Dim sSql, oRs

	sSql = "SELECT I.reservationdateid, ISNULL(I.accountid,0) AS accountid, P.paymentid "
	sSql = sSql & " FROM egov_rentalreservationdateitems I, egov_class_payment P "
	sSql = sSql & " WHERE P.isforrentals = 1 AND P.reservationid = I.reservationid AND "
	sSql = sSql & " I.reservationdateitemid = " & iReservationDateItemId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iReservationDateId = oRs("reservationdateid")
		iPaymentId = oRs("paymentid")
		If CLng(oRs("accountid")) > CLng(0) Then 
			iAccountId = oRs("accountid")
		Else
			iAccountId = "NULL"
		End If 
	Else
		iReservationDateId = 0
		iPaymentId = 0
		iAccountId = "NULL"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' double GetReservationFeesTotal( iReservationId, sField )
'--------------------------------------------------------------------------------------------------
Function GetReservationFeesTotal( ByVal iReservationId, ByVal sField )
	Dim oRs, sSql

	' Get the total reservation fees
	sSql = "SELECT ISNULL(SUM(" & sField & "),0.0000) AS someamount "
	sSql = sSql & " FROM egov_rentalreservationfees"
	sSql = sSql & " WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationFeesTotal = CDbl(oRs("someamount"))
	Else
		GetReservationFeesTotal = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetReservationIdFromDateId( iReservationDateId ) 
'------------------------------------------------------------------------------
Function GetReservationIdFromDateId( ByVal iReservationDateId ) 
	Dim oRs, sSql

	' Get the total reservation fees
	sSql = "SELECT reservationid FROM egov_rentalreservationdates "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationIdFromDateId = CLng(oRs("reservationid"))
	Else
		GetReservationIdFromDateId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetReservationIdFromPaymentId( iPaymentId ) 
'------------------------------------------------------------------------------
Function GetReservationIdFromPaymentId( ByVal iPaymentId ) 
	Dim oRs, sSql

	' Get the total reservation fees
	sSql = "SELECT ISNULL(reservationid,0) AS reservationid FROM egov_class_payment "
	sSql = sSql & " WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationIdFromPaymentId = CLng(oRs("reservationid"))
	Else
		GetReservationIdFromPaymentId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double GetReservationRefundDue( iReservationId )
'------------------------------------------------------------------------------
Function GetReservationRefundDue( ByVal iReservationId )
	Dim sSql, oRs, iRefundDue

	iRefundDue = CDbl(0)

	sSql = "SELECT reservationid, ISNULL(SUM( paidamount - (feeamount + refundamount)),0) AS refunddue "
	sSql = sSql & "FROM egov_rentalreservationdatefees  "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND paidamount > 0 AND paidamount > (feeamount + refundamount) "
	sSql = sSql & "GROUP BY reservationid "
	sSql = sSql & "UNION SELECT reservationid, ISNULL(SUM( paidamount - (feeamount + refundamount)),0) AS refunddue  "
	sSql = sSql & "FROM egov_rentalreservationdateitems  "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND paidamount > 0 AND paidamount > (feeamount + refundamount) "
	sSql = sSql & "GROUP BY reservationid "
	sSql = sSql & "UNION SELECT reservationid, ISNULL(SUM( paidamount - (feeamount + refundamount)),0) AS refunddue  "
	sSql = sSql & "FROM egov_rentalreservationfees  "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND paidamount > 0 AND paidamount > (feeamount + refundamount) "
	sSql = sSql & "GROUP BY reservationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRefundDue = iRefundDue + CDbl(oRs("refunddue"))
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

	GetReservationRefundDue = iRefundDue

End Function 


'------------------------------------------------------------------------------
' integer GetReservationRentalUserId( iReservationId )
'------------------------------------------------------------------------------
Function GetReservationRentalUserId( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(rentaluserid,0) AS rentaluserid "
	sSql = sSql & "FROM egov_rentalreservations "
	sSql = sSql & "WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationRentalUserId = CLng(oRs("rentaluserid"))
	Else
		GetReservationRentalUserId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetReservationStatusId( sStatusFlag )
'--------------------------------------------------------------------------------------------------
Function GetReservationStatusId( ByVal sStatusFlag )
	Dim sSql, oRs

	sSql = "SELECT reservationstatusid FROM egov_rentalreservationstatuses "
	sSql = sSql & " WHERE " & sStatusFlag & " = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationStatusId = oRs("reservationstatusid")
	Else
		GetReservationStatusId = 0	' This would be a problem
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double GetReservationTotalAmount( iReservationId, sField )
'------------------------------------------------------------------------------
Function GetReservationTotalAmount( ByVal iReservationId, ByVal sField )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(" & sField & ",0.0000) AS someamount "
	sSql = sSql & "FROM egov_rentalreservations "
	sSql = sSql & "WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationTotalAmount = CDbl(oRs("someamount"))
	Else
		GetReservationTotalAmount = CDbl(0.0000)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' double GetReservationTotalPaid( iReservationId )
'--------------------------------------------------------------------------------------------------
Function GetReservationTotalPaid( ByVal iReservationId )
	Dim oRs, sSql

	sSql = "SELECT ISNULL(totalpaid,0.00) AS totalpaid "
	sSql = sSql & "FROM egov_rentalreservations WHERE reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationTotalPaid = CDbl(oRs("totalpaid"))
	Else
		GetReservationTotalPaid = CDbl(0.00)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetReservationTotalPaid( iReservationId )
'--------------------------------------------------------------------------------------------------
Function GetTotalPaidForReservation( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM( P.paymenttotal),0) AS totalpaid FROM egov_class_payment P, egov_journal_entry_types J "
	sSql = sSql & "WHERE P.reservationid = " & iReservationId
	sSql = sSql & " AND P.journalentrytypeid = J.journalentrytypeid AND J.journalentrytype = 'rentalpayment'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetTotalPaidForReservation = CDbl(oRs("totalpaid"))
	Else
		GetTotalPaidForReservation = CDbl(0)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetReservationType( iReservationTypeId )
'------------------------------------------------------------------------------
Function GetReservationType( ByVal iReservationTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(reservationtype,'') AS reservationtype "
	sSql = sSql & "FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE reservationtypeid = " & iReservationTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationType = oRs("reservationtype")
	Else
		GetReservationType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' integer GetReservationTypeId( iReservationId )
'------------------------------------------------------------------------------
Function GetReservationTypeId( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT reservationtypeid "
	sSql = sSql & "FROM egov_rentalreservations "
	sSql = sSql & "WHERE reservationid = " & iReservationId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationTypeId = oRs("reservationtypeid")
	Else
		GetReservationTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' integer GetReservationTypeIdBySelector( sReservationTypeSelector )
'------------------------------------------------------------------------------
Function GetReservationTypeIdBySelector( ByVal sReservationTypeSelector )
	Dim sSql, oRs

	sSql = "SELECT reservationtypeid FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND reservationtypeselector = '" & sReservationTypeSelector & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationTypeIdBySelector = oRs("reservationtypeid")
	Else
		GetReservationTypeIdBySelector = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetReservationTypeSelection( iReservationTypeId )
'------------------------------------------------------------------------------
Function GetReservationTypeSelection( ByVal iReservationTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(reservationtypeselector,'') AS reservationtypeselector "
	sSql = sSql & "FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE reservationtypeid = " & iReservationTypeId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetReservationTypeSelection = oRs("reservationtypeselector")
	Else
		GetReservationTypeSelection = "none"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetResidentTypeByAddress iUserid, iorgid
'------------------------------------------------------------------------------
Function GetResidentTypeByAddress( ByVal iUserid, ByVal iorgid )
	' Try to match the person's address to one of the resident addresses
	Dim sSql, oCount
	
	GetResidentTypeByAddress = "N"

	sSql = "SELECT COUNT(R.residentaddressid) AS hits FROM egov_residentaddresses R, egov_users U"
	sSql = sSql & " WHERE R.orgid = U.orgid AND "
	sSql = sSql & " R.residentstreetnumber + ' ' + R.residentstreetname = U.useraddress AND "
	sSql = sSql & " R.residenttype = 'R' AND "
	sSql = sSql & " R.orgid = " & iorgid & " AND U.userid = " & iUserid

	Set oCount = Server.CreateObject("ADODB.Recordset")
	oCount.Open sSql, Application("DSN"), 3, 1
	
	If Not oCount.EOF Then 
		If CLng(oCount("hits")) > CLng(0) Then
			' Match found
			GetResidentTypeByAddress = "R"
		End If 
	End if

	oCount.Close
	Set oCount = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetResidentTypeDesc( sUserType )
'------------------------------------------------------------------------------
Function GetResidentTypeDesc( ByVal sUserType )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetResidentTypeDesc"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@sResidentType", 129, 1, 1, sUserType)
		.Parameters.Append oCmd.CreateParameter("@sDescription", 200, 2, 20)
	    .Execute
	End With

	GetResidentTypeDesc = oCmd.Parameters("@sDescription").Value

	Set oCmd = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetSelectedPeriodType( iPeriodTypeId )
'--------------------------------------------------------------------------------------------------
Function GetSelectedPeriodType( ByVal iPeriodTypeId )
	Dim sSql, oRs

	sSql = "SELECT periodtypeselector FROM egov_rentalperiodtypes WHERE periodtypeid = " & iPeriodTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetSelectedPeriodType = oRs("periodtypeselector")
	Else
		GetSelectedPeriodType = "selectedperiod"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetSelectedPeriodTypeId( sPeriodTypeSelector )
'--------------------------------------------------------------------------------------------------
Function GetSelectedPeriodTypeId( ByVal sPeriodTypeSelector )
	Dim sSql, oRs

	sSql = "SELECT periodtypeid FROM egov_rentalperiodtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND periodtypeselector = '" & sPeriodTypeSelector & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetSelectedPeriodTypeId = oRs("periodtypeid")
	Else
		GetSelectedPeriodTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' String GetTimePortion( iDateTimeValue )
'------------------------------------------------------------------------------
Function GetTimePortion( ByVal iDateTimeValue )
	Dim iEndHour, iEndMinute, sEndAmPm

	' we already have something close so use that
	SetEndingTimes iDateTimeValue, iEndHour, iEndMinute, sEndAmPm

	GetTimePortion = iEndHour & ":" & iEndMinute & " " & sEndAmPm

End Function 


'------------------------------------------------------------------------------
' String GetUserResidentType( iUserId )
'------------------------------------------------------------------------------
Function GetUserResidentType( ByVal iUserId )
	Dim oCmd

	If iUserid = "" Then
		GetUserResidentType = "N"
	Else
		' iUserId = clng(iUserId)
		Set oCmd = Server.CreateObject("ADODB.Command")
		With oCmd
			.ActiveConnection = Application("DSN")
		    .CommandText = "GetUserResidentType"
		    .CommandType = 4
			.Parameters.Append oCmd.CreateParameter("@iUserid", 3, 1, 4, iUserId)
			.Parameters.Append oCmd.CreateParameter("@ResidentType", 129, 2, 1)
		    .Execute
		End With
		
		GetUserResidentType = oCmd.Parameters("@ResidentType").Value

		Set oCmd = Nothing

		If IsNull(GetUserResidentType) Or GetUserResidentType = "" Then
			GetUserResidentType = "N"
		End if
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' Integer GetWantedDates( aWantedDates, sStartDate, sEndDate, sPeriodType, sStartTime, sEndTime, iEndDay, sOccurs, sWantedDOWs, iMonthlyPeriod, iMonthlyDOW )
'--------------------------------------------------------------------------------------------------
Function GetWantedDates( ByRef aWantedDates, ByVal sStartDate, ByVal sEndDate, ByVal sPeriodType, ByVal sStartTime, ByVal sEndTime, ByVal iEndDay, ByVal sOccurs, ByVal sWantedDOWs, ByVal iMonthlyPeriod, ByVal iMonthlyDOW )
	Dim iTotalDays


		' Find the dates based on the period they picked.
		Select Case sOccurs
			Case "o"
				' This is a single date (date range of 1 day)
				iTotalDays = SetDailyDates( aWantedDates, sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime )
			Case "d"
				' This is a date range
				iTotalDays = SetDailyDates( aWantedDates, sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime )
			Case "w"
				' This is weekly on selected days of the week
				iTotalDays = SetWeeklyDates( aWantedDates, sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime, sWantedDOWs )
			Case "m"
				' This is monthly on a selected day of the week on a selected week of the month
				iTotalDays = SetMonthlyDates( aWantedDates, sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime, iMonthlyPeriod, iMonthlyDOW )
		End Select 

		GetWantedDates = iTotalDays

End Function 


'--------------------------------------------------------------------------------------------------
' string GetWeekendSurchargeStart( iDayid )
'--------------------------------------------------------------------------------------------------
Function GetWeekendSurchargeStart( ByVal iDayid )
	Dim sSql, oRs

	sSql = "SELECT D.starthour, dbo.AddLeadingZeros(D.startminute,2) AS startminute, D.startampm "
	sSql = sSql & " FROM egov_rentaldayrates D, egov_price_types P "
	sSql = sSql & " WHERE D.pricetypeid = P.pricetypeid AND isweekendsurcharge = 1 AND D.dayid = " & iDayid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetWeekendSurchargeStart= oRs("starthour") & ":" & oRs("startminute") & " " & oRs("startampm")
	Else
		GetWeekendSurchargeStart = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean IsReservation( sReservationTypeId )
'--------------------------------------------------------------------------------------------------
Function IsReservation( ByVal sReservationTypeId )
	Dim sSql, oRs

	sSql = "SELECT isreservation FROM egov_rentalreservationtypes WHERE reservationtypeid = " & sReservationTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isreservation") Then 
			IsReservation = True 
		Else
			IsReservation = False 
		End If 
	Else
		IsReservation = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void RemoveReservationDate iReservationDateId 
'--------------------------------------------------------------------------------------------------
Sub RemoveReservationDate( ByVal iReservationDateId )
	Dim sSql

	' Remove any date items
	sSql = "DELETE FROM egov_rentalreservationdateitems WHERE reservationdateid = " & iReservationDateId
	RunSQLStatement sSql

	' Remove any date fees
	sSql = "DELETE FROM egov_rentalreservationdatefees WHERE reservationdateid = " & iReservationDateId
	RunSQLStatement sSql

	' Remove the date row
	sSql = "DELETE FROM egov_rentalreservationdates WHERE reservationdateid = " & iReservationDateId
	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean RentalFeeIsAlreadyPaid( iReservationDateId, sReservationFeeType, iReservationFeeTypeId )
'--------------------------------------------------------------------------------------------------
Function RentalFeeIsAlreadyPaid( ByVal iReservationDateId, ByVal sReservationFeeType, ByVal iReservationFeeTypeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(ledgerid) AS hits FROM egov_accounts_ledger "
	sSql = sSql & " WHERE reservationfeetype = '" & sReservationFeeType & "' AND reservationdateid = " & iReservationDateId
	sSql = sSql & " AND entrytype = 'credit' AND reservationfeetypeid = " & iReservationFeeTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			RentalFeeIsAlreadyPaid = True 
		Else
			RentalFeeIsAlreadyPaid = False 
		End If 
	Else
		RentalFeeIsAlreadyPaid = False 
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasAnyAvailableTime( iRentalId, dStartTime, dEndTime )
'--------------------------------------------------------------------------------------------------
Function RentalHasAnyAvailableTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dStartDate )
	Dim sSql, oRs, dOpeningTime, dClosingTime, dLastStart, iCount, iMinInterval, sDateAddString, bIsAllDay
	Dim iPostBuffer, iAvailableTimeBlock, dLatestAllowed, dReservationStartTime

	GetOpeningAndClosingTimes iRentalid, bOffSeasonFlag, iWeekday, dStartDate, dOpeningTime, dClosingTime

	sSql = "SELECT reservationstarttime, reservationendtime, billingendtime " 
	sSql = sSql & " FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
	sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND reservationstarttime > '" & DateValue(dStartDate) & " 0:00 AM' "
	sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dStartDate)) & " 0:00 AM' "
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	dLastStart = dOpeningTime
	iCount = clng(0)
	Do While Not oRs.EOF
		dReservationStartTime = CDate(oRs("reservationstarttime"))
		If dLastStart < CDate(oRs("reservationstarttime")) Then 
			iCount = iCount + 1
		End If 
		dLastStart = CDate(oRs("billingendtime"))
		oRs.MoveNext
	Loop

	If dLastStart < dClosingTime Then
		iCount = iCount + 1
	End If 

	If clng(iCount) = clng(0) Then 
		RentalHasAnyAvailableTime = False 
	Else 
		RentalHasAnyAvailableTime = True 
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasItems( iRentalId )
'--------------------------------------------------------------------------------------------------
Function RentalHasItems( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(rentalitemid) AS hits FROM egov_rentalitems WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			RentalHasItems = True 
		Else
			RentalHasItems = False 
		End If 
	Else
		RentalHasItems = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' boolean RentalHasNoCosts( iRentalId )
'--------------------------------------------------------------------------------------------------
Function RentalHasNoCosts( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT nocosttorent FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("nocosttorent") Then
			RentalHasNoCosts = True 
		Else
			RentalHasNoCosts = False 
		End If 
	Else
		RentalHasNoCosts = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasReservations( iRentalId )
'--------------------------------------------------------------------------------------------------
Function RentalHasReservations( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			RentalHasReservations = True 
		Else
			RentalHasReservations = False 
		End If 
	Else
		RentalHasReservations = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalIsAllDay( iRentalid, bOffSeasonFlag, iWeekDay )
'--------------------------------------------------------------------------------------------------
Function RentalIsAllDay( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekDay )
	Dim sSql, oRs, iIsOffSeason

	If bOffSeasonFlag Then 
		iIsOffSeason = 1
	Else
		iIsOffSeason = 0
	End If 

	sSql = "SELECT T.isallday FROM egov_rentaltimetypes T, egov_rentaldays D "
	sSql = sSql & " WHERE T.timetypeid = D.minimumrentaltimetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND isoffseason = " & iIsOffSeason & " AND dayofweek = " & iWeekDay

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isallday") Then
			RentalIsAllDay = True 
		Else
			RentalIsAllDay = False 
		End If 
	Else
		RentalIsAllDay = False 
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalIsClosed( iRentalid, bOffSeasonFlag, iWeekday )
'--------------------------------------------------------------------------------------------------
Function RentalIsClosed( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday )
	Dim sSql, oRs

	sSql = "SELECT isopen FROM egov_rentaldays WHERE rentalid = " & iRentalid
	sSql = sSql & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' First see if the rental is open on this day of the week
		If oRs("isopen") Then 
			RentalIsClosed = False 
		Else
			RentalIsClosed = True 
		End If
	Else
		' we do not have a day record, that is bad
		RentalIsClosed = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ReservationHasReservedDates( iReservationId )
'--------------------------------------------------------------------------------------------------
Function ReservationHasReservedDates( ByVal iReservationId )
	Dim sSql, oRs, iReservedStatusId

	iReservedStatusId = GetReservationStatusId( "isreserved" )

	sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates "
	sSql = sSql & " WHERE reservationid = " & iReservationId & " AND statusid = " & iReservedStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			ReservationHasReservedDates = True 
		Else
			ReservationHasReservedDates = False 
		End If 
	Else
		ReservationHasReservedDates = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ReservationIsForAClass( iReservationId )
'--------------------------------------------------------------------------------------------------
Function ReservationIsForAClass( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT isclass FROM egov_rentalreservationtypes T, egov_rentalreservations R "
	sSql = sSql & "WHERE T.reservationtypeid = R.reservationtypeid AND R.reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isclass") Then
			ReservationIsForAClass = True 
		Else
			ReservationIsForAClass = False 
		End If 
	Else
		ReservationIsForAClass = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ReservationNeedsBufferTimeAdded( iReservationId )
'--------------------------------------------------------------------------------------------------
Function ReservationNeedsBufferTimeAdded( ByVal iReservationId )
	Dim sSql, oRs

	' by this, public and internal get the buffer check, while block and class do not
	sSql = "SELECT T.isreservation FROM egov_rentalreservationtypes T, egov_rentalreservations R "
	sSql = sSql & "WHERE T.reservationtypeid = R.reservationtypeid AND R.reservationid = " & iReservationId
	'response.write sSql & " " 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isreservation") Then 
			ReservationNeedsBufferTimeAdded = True 
		Else
			ReservationNeedsBufferTimeAdded = False 
		End If 
	Else
		ReservationNeedsBufferTimeAdded = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void SaveClassWantedDate iReservationTempId, sStartDateTime, sEndDateTime, iEndDay, iPosition, iTimeDayId 
'--------------------------------------------------------------------------------------------------
Sub SaveClassWantedDate( ByVal iReservationTempId, ByVal sStartDateTime, ByVal sEndDateTime, ByVal iEndDay, ByVal iPosition, ByVal iTimeDayId )
	Dim sSql

	sSql = "INSERT INTO egov_rentalreservationdatestemp ( reservationtempid, sessionid, orgid, "
	sSql = sSql & "position, reservationstarttime, reservationendtime, endday, timedayid ) VALUES ( "
	sSql = sSql & iReservationTempId & ", '" & Session.SessionID & "', " & session("orgid") & ", "
	sSql = sSql & iPosition & ", '" & sStartDateTime & "', '" & sEndDateTime & "', "
	sSql = sSql & iEndDay & ", " & iTimeDayId & " )"

	'response.write sSql
	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' void SaveTempWantedDates iReservationTempId, sStartDateTime, sEndDateTime, iEndDay, iPosition 
'--------------------------------------------------------------------------------------------------
Sub SaveTempWantedDates( ByVal iReservationTempId, ByVal sStartDateTime, ByVal sEndDateTime, ByVal iEndDay, ByVal iPosition )
	Dim sSql

	sSql = "INSERT INTO egov_rentalreservationdatestemp ( reservationtempid, sessionid, orgid, "
	sSql = sSql & "position, reservationstarttime, reservationendtime, endday ) VALUES ( "
	sSql = sSql & iReservationTempId & ", '" & Session.SessionID & "', " & session("orgid") & ", "
	sSql = sSql & iPosition & ", '" & sStartDateTime & "', '" & sEndDateTime & "', " & iEndDay & " )"
	'response.write sSql
	RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' void setCurrentPaidAmount idFieldValue, idField, tableName, newAmount, iPaymentId '
'------------------------------------------------------------------------------
Sub setCurrentPaidAmount( ByVal idFieldValue, ByVal idField, ByVal tableName, ByVal newAmount, ByVal iPaymentId )
	Dim sSql

	sSql = "UPDATE " & tableName & " SET paidamount = " & FormatNumber(newAmount,2,,,0)
      sSql = sSql & " WHERE " & idField & " = " & idFieldValue

      RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' void setCurrentRefundAmount idFieldValue, idField, tableName, newAmount, iPaymentId '
'------------------------------------------------------------------------------
Sub setCurrentRefundAmount( ByVal idFieldValue, ByVal idField, ByVal tableName, ByVal newAmount, ByVal iPaymentId )
	Dim sSql

	sSql = "UPDATE " & tableName & " SET refundamount = " & FormatNumber(newAmount,2,,,0)
      sSql = sSql & " WHERE " & idField & " = " & idFieldValue
      'response.write sSql & "<br /><br />"

      RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' integer SetDailyDates( aWantedDates, sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime )
'--------------------------------------------------------------------------------------------------
Function SetDailyDates( ByRef aWantedDates, ByVal sStartDate, ByVal sEndDate, ByVal sPeriodType, ByVal iEndDay, ByVal sStartTime, ByVal sEndTime )
	Dim dTempDate 
	
	' There will always be at least one date, so put that in the array
	iTotalDays = 0
	ReDim aWantedDates(1,0)
	dTempDate = CDate(sStartDate)

	Do While dTempDate <= CDate(sEndDate)
		ReDim Preserve aWantedDates(1,iTotalDays)
		If sPeriodType = "selectedperiod" Then 
			' we have set time periods
			aWantedDates(0,iTotalDays) = dTempDate & sStartTime
			If iEndDay = clng(0) Then 
				aWantedDates(1,iTotalDays) = dTempDate & sEndTime
			Else
				aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate))) & sEndTime
			End If 
		Else
			' No set time periods selected (any time period picked)
			aWantedDates(0,iTotalDays) = dTempDate
			aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate)))
		End If 
		iTotalDays = iTotalDays + 1
		dTempDate = DateAdd("d",1,dTempDate)
	Loop 

	SetDailyDates = iTotalDays

End Function 


'--------------------------------------------------------------------------------------------------
' void SetEndingTimes dEndTime, iEndHour, iEndMinute, sEndAmPm 
'--------------------------------------------------------------------------------------------------
Sub SetEndingTimes( ByVal dEndTime, ByRef iEndHour, ByRef iEndMinute, ByRef sEndAmPm )

	iEndHour = Hour(dEndTime)  ' range 0 - 23

	If iEndHour > 11 Then 
		sEndAmPm = "PM"
	Else
		sEndAmPm = "AM"
	End If 

	If iEndHour = 0 Then
		iEndHour = 12
	ElseIf iEndHour > 12 Then
		iEndHour = iEndHour - 12
	End If 
	iEndHour = CStr(iEndHour)

	iEndMinute = Minute(dEndTime)   ' range 0 - 59
	If iEndMinute < 10 Then
		iEndMinute = "0" & CStr(iEndMinute)
	End If 
	iEndMinute = CStr(iEndMinute)

End Sub 


'--------------------------------------------------------------------------------------------------
' void SetHoursToOpenAndClose iRentalId, dStartDate, sStartTime, sEndTime, iEndDay
'--------------------------------------------------------------------------------------------------
Sub SetHoursToOpenAndClose( ByVal iRentalId, ByVal dStartDate, ByRef sStartTime, ByRef sEndTime, ByRef iEndDay )
	Dim sSql, oRs, iWeekday, bOffSeasonFlag

	iWeekday = Weekday(dStartDate)

	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, dStartDate )

	sSql = "SELECT isopen, isavailabletopublic, ISNULL(openinghour,0) AS openinghour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(openingminute,0),2) AS openingminute, ISNULL(openingampm,'AM') AS openingampm, "
	sSql = sSql & " ISNULL(closinghour,0) AS closinghour, dbo.AddLeadingZeros(ISNULL(closingminute,0),2) AS closingminute, "
	sSql = sSql & " ISNULL(closingampm,'AM') AS closingampm, ISNULL(closingday,0) AS closingday "
	sSql = sSql & " FROM egov_rentaldays "
	sSql = sSql & " WHERE rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isopen") Then 
			sStartTime = dStartDate & " " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm")
			iEndDay = oRs("closingday")
			If oRs("closingday") = "0" Then 
				sEndTime = dStartDate
			Else
				sEndTime = DateAdd("d", 1, dStartDate)
			End If 
			sEndTime = sEndTime & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm")
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer SetMonthlyDates( aWantedDates, sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime, iMonthlyPeriod, iMonthlyDOW )
'--------------------------------------------------------------------------------------------------
Function SetMonthlyDates( ByRef aWantedDates, ByVal sStartDate, ByVal sEndDate, ByVal sPeriodType, ByVal iEndDay, ByVal sStartTime, ByVal sEndTime, ByVal iMonthlyPeriod, ByVal iMonthlyDOW )
	Dim dTempDate, dIncrementDate
	
	' There will always be at least one date, so put that in the array
	iTotalDays = 0
	ReDim aWantedDates(1,0)
	dTempDate = GetNextOrdinalDayMonth( iMonthlyDOW, iMonthlyPeriod, Month(CDate(sStartDate)),  Year(CDate(sStartDate)) )
	dIncrementDate = CDate(Month(CDate(sStartDate)) & "/1/" & Year(CDate(sStartDate)))

	Do While dTempDate <= CDate(sEndDate)
		ReDim Preserve aWantedDates(1,iTotalDays)
		If sPeriodType = "selectedperiod" Then 
			' we have set time periods
			aWantedDates(0,iTotalDays) = dTempDate & sStartTime
			If iEndDay = clng(0) Then 
				aWantedDates(1,iTotalDays) = dTempDate & sEndTime
			Else
				aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate))) & sEndTime
			End If 
		Else
			' No set time periods selected
			aWantedDates(0,iTotalDays) = dTempDate
			aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate)))
		End If 
		iTotalDays = iTotalDays + 1
		dIncrementDate = DateAdd("m", 1, dIncrementDate)
		dTempDate = GetNextOrdinalDayMonth( iMonthlyDOW, iMonthlyPeriod, Month(dIncrementDate),  Year(dIncrementDate) )
	Loop

	SetMonthlyDates = iTotalDays
	
End Function 


'--------------------------------------------------------------------------------------------------
' integer SetWeeklyDates( ByRef aWantedDates, ByVal sStartDate, sEndDate, sPeriodType, iEndDay, sStartTime, sEndTime, sWantedDOWs )
'--------------------------------------------------------------------------------------------------
Function SetWeeklyDates( ByRef aWantedDates, ByVal sStartDate, ByVal sEndDate, ByVal sPeriodType, ByVal iEndDay, ByVal sStartTime, ByVal sEndTime, ByVal sWantedDOWs )
	Dim dTempDate 
	
	' There will always be at least one date, so put that in the array
	iTotalDays = 0
	ReDim aWantedDates(1,0)
	dTempDate = CDate(sStartDate)

	Do While dTempDate <= CDate(sEndDate)
		sWeekDay = CStr(Weekday(dTempDate)) ' get the DOW number 1-7
		If InStr(sWantedDOWs, sWeekDay) > 0 Then 
			' If the dow Is a wanted one Then keep it
			ReDim Preserve aWantedDates(1,iTotalDays)
			If sPeriodType = "selectedperiod" Then 
				' we have set time periods
				aWantedDates(0,iTotalDays) = dTempDate & sStartTime
				If iEndDay = clng(0) Then 
					aWantedDates(1,iTotalDays) = dTempDate & sEndTime
				Else
					aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate))) & sEndTime
				End If 
			Else
				' No set time periods selected
				aWantedDates(0,iTotalDays) = dTempDate
				aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate)))
			End If 
			iTotalDays = iTotalDays + 1
		End If 
		dTempDate = DateAdd("d",1,dTempDate)
	Loop 
	SetWeeklyDates = iTotalDays

End Function 


'--------------------------------------------------------------------------------------------------
Sub ShowReservationPayments( ByVal iReservationId, ByVal sClass )
	Dim sSql, oRs

	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) AS paidamount "
	sSql = sSql & "FROM egov_accounts_ledger A, egov_class_payment J "
	sSql = sSql & "WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 1 AND A.entrytype = 'debit' "
	sSql = sSql & "AND A.reservationid = " & iReservationId
	sSql = sSql & "GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"">Date</td><td class=""subheadercell"">&nbsp;</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td>" & oRs("paymentid") & "</td>"
			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td align=""right"">"
			response.write FormatNumber(oRs("paidamount"),2,,,0) 
			response.write "</td></tr>"
			oRs.MoveNext 
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
Sub ShowReservationRefunds( ByVal iReservationId, ByVal sClass )
	Dim sSql, oRs
	
	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) as refundamount "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_payment J "
	sSql = sSql & " WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 0 AND A.entrytype = 'debit' "
	sSql = sSql & " AND A.reservationid = " & iReservationId
	sSql = sSql & " GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"" colspan=""2"">Date</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td>&nbsp;" & oRs("paymentid") & "</td>"
			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td align=""right"">"
			response.write FormatNumber(oRs("refundamount"),2,,,0) 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Function buildReservationDateTime( ByVal iReservationDate, ByVal iReservationHour, ByVal iReservationMinute, ByVal iReservationAMPM ) 
	Dim lcl_return, sReservationDate, sReservationHour, sReservationMinute, sResevationAMPM

	lcl_return         = ""
	sReservationDate   = ""
	sReservationHour   = ""
	sReservationMinute = ""
	sReservationAMPM   = ""

	If iReservationDate <> "" Then 
		sReservationDate = iReservationDate
		lcl_return       = sReservationDate

		If Trim(iReservationHour) <> "" Then 
			sReservationHour = trim(iReservationHour)
			lcl_return       = lcl_return & " " & sReservationHour
		End If 

		If Trim(iReservationMinute) <> "" Then 
			sReservationMinute = trim(iReservationMinute)
			lcl_return         = lcl_return & ":" & sReservationMinute
		End If 

		If Trim(iReservationAMPM) <> "" Then 
			sReservationAMPM = trim(iReservationAMPM)
			lcl_return       = lcl_return & " " & sReservationAMPM
		End If 
	End If 

	buildReservationDateTime = lcl_return

End Function 



%>
