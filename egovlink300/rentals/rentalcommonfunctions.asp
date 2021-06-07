<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalcommonfunctions.asp
' AUTHOR: Steve Loar
' CREATED: 01/14/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Collection of common rentals functions. 
'               Try to keep them in alphabetical order, please.
'
' MODIFICATION HISTORY
' 1.0   01/14/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' datetime AddPostBufferTime( iRentalid, bOffSeasonFlag, dEndDateTime, dStartDateTime )
'--------------------------------------------------------------------------------------------------
Function AddPostBufferTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal dEndDateTime, ByVal dStartDateTime )
	Dim sSql, oRs, iWeekday

	iWeekday = Weekday(dStartDateTime)

	sSql = "SELECT ISNULL(postbuffer,0) AS postbuffer, P.dateaddstring AS postdateaddstring "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes P "
	sSql = sSql & " WHERE D.postbuffertimetypeid = P.timetypeid "
	sSql = sSql & " AND rentalid = " & iRentalid & " AND isoffseason = " & bOffSeasonFlag
	sSql = sSql & " AND dayofweek = " & iWeekday

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

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


'------------------------------------------------------------------------------
' void AddToPaymentLog iPaymentControlNumber, sLogEntry 
'------------------------------------------------------------------------------
Sub AddToPaymentLog( ByVal iPaymentControlNumber, ByVal sLogEntry  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & iOrgID & ", 'public', 'rentals', '" & dbready_string(sLogEntry, 500) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

End Sub 


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
' boolean CategoryHasRestricedPeriod( iRti, iDayIntervals, sCategories, sPeriod )
'--------------------------------------------------------------------------------------------------
Function CategoryHasRestricedPeriod( ByVal iRti, ByRef iDayIntervals, ByRef sCategories, ByRef sPeriod )
	Dim sSql, oRs, sMatchDateString

	sCategories = ""
	iDayIntervals = 0
	sPeriod = ""

	' want to find the longest period they have to wait
	sSql = "SELECT C.recreationcategoryid, P.restrictionperiod, dateaddstring "
	sSql = sSql & " FROM egov_recreation_categories C, egov_rentals_to_categories RTC, "
	sSql = sSql & " egov_rentalreservationstemppublic R, egov_rentalrestrictionperiods P "
	sSql = sSql & " WHERE C.recreationcategoryid = RTC.recreationcategoryid "
	sSql = sSql & " AND RTC.rentalid = R.rentalid AND R.reservationtempid = " & iRti 
	sSql = sSql & " AND C.hasrestrictedperiod = 1 AND P.restrictionperiodid = C.restrictedperiodid "
	sSql = sSql & " ORDER BY P.dateaddstring DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		CategoryHasRestricedPeriod = True 
		sMatchDateString = oRs("dateaddstring")
		sPeriod = sMatchDateString
		' ' one week interval
		If oRs("dateaddstring") = "ww" Then 
			iDayIntervals = 7
		Else
			' one day interval
			iDayIntervals = 1
		End If 
		' build the list of categories that have that same period restriction
		' if there are more than daily and weekly, then this may not work right
		Do While Not oRs.EOF
			If sMatchDateString = oRs("dateaddstring") Then 
				If sCategories <> "" Then 
					sCategories = ", " & sCategories
				End If 
				sCategories = sCategories & oRs("recreationcategoryid")
			End If 
			oRs.MoveNext 
		Loop 
	Else
		CategoryHasRestricedPeriod = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string sFlag = CheckForExistingReservations( iRentalid, dWantedStartTime, dWantedEndTime, sPeriodType )
'--------------------------------------------------------------------------------------------------
Function CheckForExistingReservations( ByVal iRentalid, ByVal dWantedStartTime, ByVal dWantedEndTime, ByVal sPeriodType, ByVal bOffSeasonFlag )
	Dim sSql, oRs

	If sPeriodType = "selectedperiod" Then
		If EndTimeIsNotClosingTime( iRentalId, dWantedEndTime, bOffSeasonFlag, dWantedStartTime ) Then 
			' Add on the end buffer
			dWantedEndTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, dWantedEndTime, dWantedStartTime )
		End If 
		' we will add a minute to this so start time can be the same as the end of bufferend time of another reservation
		dWantedStartTime = DateAdd("n", 1, dWantedStartTime)
		' we will remove a minute so the end of the buffer can be the same minute as the start of another reservation
		dWantedEndTime = DateAdd("n", -1, dWantedEndTime)
		' set sql to look for conflicting times
		sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
		sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
		sSql = sSql & " AND (reservationstarttime BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
		sSql = sSql & " OR reservationendtime BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
		sSql = sSql & " OR (reservationstarttime <= '" & dWantedStartTime & "' AND reservationendtime >= '" & dWantedEndTime & "'))"
	Else
		' allday - set sql to look for any starting time on that day
		sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
		sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
		sSql = sSql & " AND reservationstarttime > '" & DateValue(dWantedStartTime) & " 0:00 AM' "
		sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dWantedStartTime)) & " 0:00 AM' "
	End If 
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	session("overlapSQL") = sSql
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
' void CheckOrgRentalRoundUp dStartTime, dEndTime, iEndHour, iEndMinute, sEndAmPm
'--------------------------------------------------------------------------------------------------
Sub CheckOrgRentalRoundUp( ByVal iOrgId, ByVal dStartTime, ByRef dEndTime, ByRef iEndHour, ByRef iEndMinute, ByRef sEndAmPm )
	Dim sSql, oRs, sInterval, iRoundUp, iTimeInterval, iIntervals

	' Get the minimum interval for the org
	sSql = "SELECT ISNULL(O.rentalroundup,0) AS rentalroundup, ISNULL(T.dateaddstring,'n') AS dateaddstring "
	sSql = sSql & " FROM Organizations O, egov_rentaltimetypes T "
	sSql = sSql & " WHERE O.rentalrounduptimetypeid = T.timetypeid AND O.orgid = " & iOrgId

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
' boolean CheckRentalAvailability( iRentalid, sStartDateTime, sEndDateTime, sMessage )
'--------------------------------------------------------------------------------------------------
Function CheckRentalAvailability( ByVal iRentalid, ByVal sStartDateTime, ByVal sEndDateTime, ByRef sMessage )
	Dim sOffSeasonFlag

	sMessage = ""

	sOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )

	' Check if rental is open that day and that we are not looking for a time when it is not open
	sCheckReturn = CheckRentalHours( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag )
	
	If Left(sCheckReturn,2) = "No" Then 
		sMessage = "closed"
	End If 

	If sCheckReturn = "toolate" Then
		sMessage = "toolate"
	End If 

	If sMessage = "" Then 
		' Check if the time is available
		If CheckForExistingReservations( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag ) = "No" Then
			sMessage = "conflict" 
		End If 
	End If 

	If sMessage = "" Then 
		CheckRentalAvailability = True 
	Else
		CheckRentalAvailability = False 
	End If 


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
				'response.write "dOpeningTime: " & DateValue(dWantedStartTime) & " " & oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm") & "<br />"
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
							CheckRentalHours = "No5 " & dWantedEndTime & " " & DateDiff( "n", CDate(dWantedEndTime), CDate(dClosingTime) )
							CheckRentalHours = "No"
						Else
							' Finally is it OK for rental during the wanted time
							CheckRentalHours = "Yes1 " & dWantedEndTime & " " & DateDiff( "n", CDate(dWantedEndTime), CDate(dClosingTime) )
							CheckRentalHours = "Yes"
						End If 
					Else
						' The start time is too late in the day
						CheckRentalHours = "toolate" '& dLatestStartTime & "|" & dWantedStartTime
					End If 
				Else
					' the start time is before the place opens
					'response.write "Opening is after starttime"
					CheckRentalHours = "No"
				End If 
			Else
				' for allday and any time we do not have a start and end time for the reservation so it is OK if they just open
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
' void ClearTempReservation iReservationTempId, iOrgId
'--------------------------------------------------------------------------------------------------
Sub ClearTempReservation( ByVal iReservationTempId, ByVal iOrgId )
	Dim sSql
		
	sSql = "DELETE FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & iOrgId

	RunSQLStatement sSql		' in ../includes/common.asp

'	sSql = "DELETE FROM egov_rentalreservationdatestemp WHERE reservationtempid = " & iReservationTempId
'	sSql = sSql & " AND orgid = " & iOrgId
	'response.write sSql & "<br /><br />"
'	RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' integer CreatePaymentControlRow( sLogEntry )
'------------------------------------------------------------------------------
Function CreatePaymentControlRow( ByVal sLogEntry )
	Dim sSql, iPaymentControlNumber

	sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iOrgID & ", 'public', 'rentals', '" & sLogEntry & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunIdentityInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

	CreatePaymentControlRow = iPaymentControlNumber

End Function 


'--------------------------------------------------------------------------------------------------
' void CreateRentalReservationDateFees iReservationDateId, iReservationId, iRentalid, bOffSeasonFlag, iWeekday, sUserType, sStartDateTime, sEndDateTime
'--------------------------------------------------------------------------------------------------
Sub CreateRentalReservationDateFees( ByVal iReservationDateId, ByVal iReservationId, ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal sUserType, ByVal sStartDateTime, ByVal sEndDateTime )
	Dim sSql, oRs, sAccount, sHour, sMinute, sAmPm, iDuration, iFeeAmount, bAddFee, sSurchargeStart

	' Get the Rental Reservation Fees by user type (Resident, Nonresident, Everyone) for the day
	sSql = "SELECT R.pricetypeid, P.pricetypename, ISNULL(accountid,0) AS accountid, R.ratetypeid, ISNULL(amount,0.00) AS amount, "
	sSql = sSql & "ISNULL(R.starthour,0) AS starthour, dbo.AddLeadingZeros(ISNULL(R.startminute,0),2) AS startminute, "
	sSql = sSql & "ISNULL(R.startampm,'AM') AS startampm, P.pricetype, P.isbaseprice, P.isfee, P.hasstarttime, P.isweekendsurcharge, "
	sSql = sSql & "ISNULL(P.basepricetypeid,0) AS basepricetypeid, P.checkresidency, P.isresident, T.datediffstring, alwaysadd "
	sSql = sSql & "FROM egov_rentaldayrates R, egov_rentaldays D, egov_price_types P, egov_rentalratetypes T "
	sSql = sSql & "WHERE D.dayid = R.dayid AND D.rentalid = R.rentalid AND R.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND T.ratetypeid = R.ratetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday
	sSql = sSql & " ORDER BY P.displayorder"
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
			If oRS("hasstarttime") Then
				sHour = oRs("starthour")
				sMinute = oRs("startminute")
				sAmPm = "'" & oRs("startampm") & "'"
				sAmPmValue = oRs("startampm")
				'response.write "sSurchargeStart = " & DateValue(sStartDateTime) & " " & sHour & ":" & sMinute & " " & sAmPmValue & "<br /><br />"
				sSurchargeStart = CDate(DateValue(sStartDateTime) & " " & sHour & ":" & sMinute & " " & sAmPmValue)
			Else
				sSurchargeStart = sStartDateTime
			End If 

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
			sSql = "INSERT INTO egov_rentalreservationdatefees (reservationdateid, reservationid, rentalid, pricetypeid, "
			sSql = sSql & "accountid, ratetypeid, amount, starthour, startminute, startampm, feeamount, paidamount, "
			sSql = sSql & "refundamount, duration, datediffstring ) VALUES ( " & iReservationDateId & ", " & iReservationId & ", " & iRentalid & ", "
			sSql = sSql & oRs("pricetypeid") & ", " & sAccount & ", " & oRs("ratetypeid") & ", " & sRate & ", "
			sSql = sSql & sHour & ", " & sMinute & ", " & sAmPm & ", " & iFeeAmount & "," & iFeeAmount & ", 0.0000, "
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
' void CreateRentalReservationDateItems iReservationDateId, iReservationId, iRentalid
'--------------------------------------------------------------------------------------------------
Sub CreateRentalReservationDateItems( ByVal iReservationDateId, ByVal iReservationId, ByVal iRentalid )
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

		sAmount = CDbl(oRs("amount"))

		sSql = "INSERT INTO egov_rentalreservationdateitems ( reservationdateid, reservationid, rentalid, "
		sSql = sSql & "rentalitemid, rentalitem, accountid, maxavailable, amount, quantity, feeamount, paidamount, refundamount ) "
		sSql = sSql & " VALUES ( " & iReservationDateId & ", " & iReservationId & ", " & iRentalid & ", " & oRs("rentalitemid") & ", '"
		sSql = sSql & SQLText(oRs("rentalitem")) &"', " & sAccount & ", " & oRs("maxavailable") & ", " & sAmount & ", "
		sSql = sSql & "0, 0.00, 0.00, 0.00 )"
		'response.write sSql & "<br /><br />"

		RunSQLStatement sSql

		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void CreateRentalReservationFees iReservationId, iRentalid, iSelectedPriceTypeId
'--------------------------------------------------------------------------------------------------
Sub CreateRentalReservationFees( ByVal iReservationId, ByVal iRentalid, ByVal iSelectedPriceTypeId )
	Dim sSql, oRs, sAccount, sPrompt, dFeeAmount, bUpdateReservationAlcoholCheck

	' Add the Rental Reservation Fees
	sSql = "SELECT F.pricetypeid, ISNULL(accountid,0) AS accountid, ISNULL(amount,0.00) AS amount, "
	sSql = sSql & " ISNULL(prompt,'') AS prompt, P.isoptional "
	sSql = sSql & " FROM egov_rentalfees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.rentalid = " & iRentalid
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
			sPrompt = "'" & SQLText(oRs("prompt")) & "'"
		End If 
		If oRs("isoptional") Then 
			' This would be the alcohol surcharge. See if they picked that
			If CLng(oRs("pricetypeid")) = CLng(iSelectedPriceTypeId) Then 
				dFeeAmount = CDbl(oRs("amount"))
				' This is the serving Alcohol Fee so need to update the reservation.
				bUpdateReservationAlcoholCheck = True 
			Else 
				dFeeAmount = "0.0000" 
				bUpdateReservationAlcoholCheck = False 
			End If 
		Else
			' This would be the deposit
			dFeeAmount = CDbl(oRs("amount"))
			bUpdateReservationAlcoholCheck = False 
		End If 
		sSql = "INSERT INTO egov_rentalreservationfees ( reservationid, rentalid, pricetypeid, amount, accountid, feeamount, prompt, paidamount, refundamount ) "
		sSql = sSql & " VALUES ( " & iReservationId & ", " & iRentalid & ", " & oRs("pricetypeid") & ", "
		sSql = sSql & CDbl(oRs("amount")) & ", " & sAccount & ", " & dFeeAmount & ", " & sPrompt & ", " & dFeeAmount & ", 0.00 )"
		'response.write sSql & "<br /><br />"

		RunSQLStatement sSql

		If bUpdateReservationAlcoholCheck Then 
			sSql = "UPDATE egov_rentalreservations SET servingalcohol = 1 WHERE reservationid = " & iReservationId
			RunSQLStatement sSql
		End If 

		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean DateIsInCurrentSeason( iRentalId, dCheckDate, dAvailableDate )
'--------------------------------------------------------------------------------------------------
Function DateIsInCurrentSeason( ByVal iRentalId, ByVal dCheckDate, ByRef dAvailableDate )
	Dim sSql, oRs, dSeasonStartDate, dSeasonEndDate

	' See if we have an off season and if so is the date in the off season.
	' You want to return a flag indicating which season (in, off) to use.

	sSql = "SELECT hasoffseason, offseasonstartmonth, offseasonstartday, offseasonendmonth, offseasonendday "
	sSql = sSql & "FROM egov_rentals WHERE rentalid = " & iRentalid
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("hasoffseason") Then
			If GetOffSeasonFlag( iRentalid, Date() ) = "1" Then 
'				response.write "Off Season<Br />"
				' we are in an off season so set the start and end dates for that
				dSeasonStartDate = CDate(oRs("offseasonstartmonth") & "/" & oRs("offseasonstartday") & "/" & Year(Date()))
				dSeasonEndDate = CDate(oRs("offseasonendmonth") & "/" & oRs("offseasonendday") & "/" & Year(Date()))
			Else
'				response.write "In Season<Br />"
				' current season start and end dates for currently being in season
				dSeasonStartDate = CDate(oRs("offseasonendmonth") & "/" & oRs("offseasonendday") & "/" & Year(Date()))
				dSeasonEndDate = CDate(oRs("offseasonstartmonth") & "/" & oRs("offseasonstartday") & "/" & Year(Date()))
			End If 

			If dSeasonStartDate > Date() Then
				' we need to set the current start date to last year
				dSeasonStartDate = DateAdd("yyyy", -1, dSeasonStartDate)
			End If 
			If dSeasonEndDate < Date() Then 
				' we need to set the current end date to next year
				dSeasonEndDate = DateAdd("yyyy", 1, dSeasonEndDate)
			End If 
'			response.write "dSeasonStartDate = " & dSeasonStartDate & "<Br />"
'			response.write "dSeasonEndDate = " & dSeasonEndDate & "<Br />"
'			response.write "dCheckDate = " & dCheckDate & "<Br />"
			If dSeasonStartDate <= dCheckDate Then
				If dSeasonEndDate > dCheckDate Then 
					DateIsInCurrentSeason = True 
				Else
					DateIsInCurrentSeason = False 
				End If 
			Else
				DateIsInCurrentSeason = False 
			End If 
			dAvailableDate = dSeasonEndDate
		Else
			' if no off season, then the date is always in the current season
			DateIsInCurrentSeason = True 
		End If 
	Else
		' something is wrong
		DateIsInCurrentSeason = False 
	End If 
	

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean DisplayAvailability iRentalId, dWantedDate, bOffSeasonFlag
'--------------------------------------------------------------------------------------------------
Function DisplayAvailability( ByVal iRentalId, ByVal dWantedDate, ByVal bOffSeasonFlag )
	Dim dAvailableDate, bHasHours, sPhoneNumber, bShowCall, bShowCallMsg
	
	bShowCall = orghasfeature( iorgid, "set_call_flag" )
	bShowCallMsg = False 

	sPhoneNumber = GetRentalSupervisorPhone( iRentalId )
	If sPhoneNumber = "" Then
		' Use the City Default Phone Number 
		' sDefaultPhone is set in include_top_functions
		sPhoneNumber = FormatPhoneNumber( sDefaultPhone )	' in common.asp
	End If 
	sPhoneNumber = Trim(sPhoneNumber)


	' see if the date is in the past
	If dWantedDate < CDate(DateValue(Date())) Then
		response.write "<p class=""noreservemsg"">Reservations cannot be made for past dates.</p>"
		bHasHours = False
	Else 
		' check if closed on this date
		If RentalIsClosed( iRentalId, bOffSeasonFlag, Weekday(dWantedDate) ) Then
			response.write "<p class=""noreservemsg"">Closed</p>"
			bHasHours = False
		Else 
			' check that the public is allows to rent this date
			If RentalIsAvailableToPublic( iRentalId, bOffSeasonFlag, Weekday(dWantedDate) ) Then
				' if rental has in season restriction then
				If RentalHasCurrentSeasonRestriction( iRentalId ) Then 
					' if date is in the current season then
					If DateIsInCurrentSeason( iRentalId, dWantedDate, dAvailableDate ) Then 	' in rentalcommonfunctions.asp
						bOkToShow = True 
						'response.write "In curent season<br />"
					Else 
						bOkToShow = False 
					End If 
				Else 
					bOkToShow = True 
				End If 
				bHasHours = False 
				If bOkToShow = True Then 
					' Go find the available times here
					bHasHours = ShowAvailability( iRentalId, bOffSeasonFlag, Weekday(dWantedDate), dWantedDate, bShowCall, bShowCallMsg )
					'If bHasHours Then
					'	ShowMinimumReservationTime oRs("rentalid"), bOffSeasonFlag, Weekday(CDate(aWantedDates(0,x)))
					'End If 
					' if bHasHours = false and org has call feature the display message'
					If bHasHours = false And bShowCall And bShowCallMsg Then 
						' THis is too general.  Needs to be reservation specific'
						response.write "<p class=""noreservemsg"">Contact us at " & sPhoneNumber & " to inquire about reservations on this date.</p>"
					End If 
				Else
					bHasHours = ShowAvailability( iRentalId, bOffSeasonFlag, Weekday(dWantedDate), dWantedDate, bShowCall, bShowCallMsg )
					'response.write "<p class=""noreservemsg"">You can only make reservations for dates in the current season.<br />The next season starts " & Month(dAvailableDate) & "/" & Right(("0" & CStr(Day(dAvailableDate))),2) & ".<br />Contact our office to make reservations for this date.</p>"
					response.write "<p class=""noreservemsg"">You can only make reservations for dates in the current season.<br />Contact us at " & sPhoneNumber & " to make reservations for this date.</p>"
					bHasHours = False 
				End If 
			Else
				bHasHours = ShowAvailability( iRentalId, bOffSeasonFlag, Weekday(dWantedDate), dWantedDate, bShowCall, bShowCallMsg )
				response.write "<p class=""noreservemsg"">Contact us at " & sPhoneNumber & " to make reservations for this date.</p>"
				bHasHours = False
			End If 
		End If 
	End If 

	DisplayAvailability = bHasHours

End Function 


'--------------------------------------------------------------------------------------------------
' void DisplayCategoryMenu
'--------------------------------------------------------------------------------------------------
Sub DisplayCategoryMenu( ByVal iOrgId )
	Dim sSql, oRs, blnFirst

	blnFirst = True 

	sSql = "SELECT recreationcategoryid, categorytitle, isforrentals FROM egov_recreation_categories "
	sSql = sSql & "WHERE isroot = 0 and hidefrompublic = 0 AND orgid = " & iOrgId & " ORDER BY categorytitle"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	if oRs.RecordCount > 1 then

	Response.Write vbcrlf & "<p>" & vbcrlf & "<div class=""subcategorymenu"">" 
	response.Write vbcrlf & "<a class=""subcategorymenu featurename"" style=""display:table;"" href=""../rentals/rentalcategories.asp"" >" & GetOrgFeatureName( "rentals" ) & "</a>"

	Do While Not oRs.EOF
		If oRs("isforrentals") then
			sGoTo = "rentalofferings"
		Else
			sGoTo = "../recreation/facilitycategory"
		End If 

		If Not blnFirst Then
			Response.Write(" | ")
		Else
			blnFirst = False
		End If

		Response.Write vbcrlf & "<a class=""subcategorymenu"" href=""" & sGoTo & ".asp?categoryid=" & oRs("recreationcategoryid") & """ >" & oRs("categorytitle") & "</a>"
		oRs.MoveNext
	Loop
	
	Response.Write vbcrlf & "</div>"

	end if

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayRentalDocuments iRentalId
'--------------------------------------------------------------------------------------------------
Sub DisplayRentalDocuments( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT documenttitle, documenturl FROM egov_rentaldocuments WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<p id=""documentcollection""><strong>Documents:</strong><br />"
		Do While Not oRs.EOF
			response.write "<span class=""documentitle""><a href=""" & oRs("documenturl") & """ target=""_blank"">" & oRs("documenttitle") & "</a></span><br />"
			oRs.MoveNext
		Loop
		response.write "</p>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub


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
' string sTime = FormatTimeString( dDateTimeString )
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
' Function GetAccountName( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAccountName( iPaymentId, iPaymentTypeId )
	Dim sSql, oName

	sSql = "SELECT userfname, userlname FROM egov_verisign_payment_information, egov_users "
	sSql = sSql & " WHERE paymentid = " & iPaymentId & " AND paymenttypeid = " & iPaymentTypeId
	sSql = sSql & " AND citizenuserid = userid "

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		GetAccountName = oName("userfname") & " " & oName("userlname")
	Else
		GetAccountName = ""
	End If 

	oName.Close
	Set oName = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetAdminLocation( iAdminLocationId )
'--------------------------------------------------------------------------------------------------
Function GetAdminLocation( ByVal iAdminLocationId )
	Dim sSql, oRs

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iAdminLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetAdminLocation = oRs("name")
	Else
		GetAdminLocation = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string  GetAdminName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT firstname + ' ' + lastname AS username FROM users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetAdminName = oRs("username")
	Else
		GetAdminName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetAllDayHours iRentalid, bOffSeasonFlag, iWeekDay, dWantedStartTime, dWantedEndTime, iStartHour, iStartMinute, sStartAmPm, iEndHour iEndMinute, sEndAmPm
'--------------------------------------------------------------------------------------------------
Sub GetAllDayHours( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekDay, ByRef dWantedStartTime, ByRef dWantedEndTime, ByRef iStartHour, ByRef iStartMinute, ByRef sStartAmPm, ByRef iEndHour, ByRef iEndMinute, ByRef sEndAmPm )
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
		iStartHour = oRs("openinghour")
		iStartMinute = oRs("openingminute")
		sStartAmPm = oRs("openingampm")
		
		' Set the end time to the closing time 
		If CLng(oRs("closingday")) = CLng(0) Then
			dWantedEndTime = CDate(DateValue(dWantedStartTime) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		Else
			dWantedEndTime = CDate(DateAdd("d", 1, DateValue(dWantedStartTime)) & " " & oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm"))
		End If 
		iEndHour = oRs("closinghour")
		iEndMinute = oRs("closingminute")
		sEndAmPm = oRs("closingampm")

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


'--------------------------------------------------------------------------------------------------
' string GetFamilyEmail( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyEmail( ByVal iUserId )
	Dim sSql, oRs, iFamilyId

	iFamilyId = GetFamilyId( iUserId )

	sSql = "SELECT useremail FROM egov_users "
	sSql = sSql & " WHERE useremail IS NOT NULL AND headofhousehold = 1 AND familyid = " & iFamilyId 
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetFamilyEmail = oRs("useremail")
	Else
		GetFamilyEmail = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetFamilyId( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT familyid FROM egov_users WHERE userid = " & iUserId 
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFamilyId = oRs("familyid")
	Else
		GetFamilyId = iUserId
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' datetime GetFirstAvailableTime( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
'--------------------------------------------------------------------------------------------------
Function GetFirstAvailableTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dStartDate )
	Dim sSql, oRs, dOpeningTime, dClosingTime, dLastStart, iMinInterval, sDateAddString, bIsAllDay
	Dim iPostBuffer, iAvailableTimeBlock, dLatestAllowed, bFound

	bFound = False 

	GetOpeningAndClosingTimes iRentalid, bOffSeasonFlag, iWeekday, dStartDate, dOpeningTime, dClosingTime

	' Get Minimal Time interval Info for this day
	bHasMinimum = GetMinimalTimeInfo( iRentalid, bOffSeasonFlag, iWeekday, iMinInterval, sDateAddString, bIsAllDay )

	If Not bIsAllDay Then 
		' convert this interval into minutes if needed
		If sDateAddString = "h" Then
			iMinInterval = CLng(iMinInterval) * 60
		End If 

		' get the post buffer for this day, if any
		GetPostBufferTime iRentalid, bOffSeasonFlag, iWeekday, iPostBuffer, sDateAddString

		If sDateAddString = "h" Then
			' convert to minutes
			iPostBuffer = CLng(iPostBuffer) * 60
		End If 

		' add the post buffer to the minimum allowed time
		'iMinInterval = CLng(iMinInterval) + CLng(iPostBuffer)
	End If 

	sSql = "SELECT reservationstarttime, reservationendtime " 
	sSql = sSql & " FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
	sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND reservationstarttime > '" & DateValue(dStartDate) & " 0:00 AM' "
	sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dStartDate)) & " 0:00 AM' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	dLastStart = dOpeningTime
	Do While Not oRs.EOF
		If dLastStart = dOpeningTime Then
			dReservationStartTime = CDate(oRs("reservationstarttime"))
		Else
			If clng(iPostBuffer) > clng(0) Then 
				dReservationStartTime = DateAdd("n", -(iPostBuffer), CDate(oRs("reservationstarttime")) )
			Else
				dReservationStartTime = CDate(oRs("reservationstarttime"))
			End If 
		End If 
		'response.write "dReservationStartTime = " & dReservationStartTime & "<br />"
		If dLastStart < CDate(oRs("reservationstarttime")) Then 
			If Not bIsAllDay Then 
				'response.write "dLastStart = " & dLastStart & "<br />"
				iAvailableTimeBlock = DateDiff("n", dLastStart, dReservationStartTime)
				' if the available time >= minimum allowed time then we have a time
				If (bHasMinimum = False) Or (bHasMinimum = True And iAvailableTimeBlock >= iMinInterval) Then 
					If Not bFound Then 
						GetFirstAvailableTime = dLastStart
						bFound = True 
					End If 
				End If 
			End If 
		End If 
		dLastStart = CDate(oRs("reservationendtime"))
		oRs.MoveNext
	Loop

	If bFound = False And dLastStart < dClosingTime Then 
		iAvailableTimeBlock = DateDiff("n", dLastStart, dClosingTime)
		If (bHasMinimum = False) Or (bHasMinimum = True And iAvailableTimeBlock >= iMinInterval) Then 
			' check that the last start time is before the latest reservation start time
			dLatestAllowed = GetLatestReservationTime( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
			If DateDiff("n", dLastStart, dLatestAllowed) >= 0 Then 
				GetFirstAvailableTime = dLastStart
			End If 
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetInitialReservationStatusId( iOrgid )
'--------------------------------------------------------------------------------------------------
Function GetInitialReservationStatusId( ByVal iOrgid )
	Dim sSql, oRs

	' Get the initial status to make a reservation so we do not hard code this
	sSql = "SELECT reservationstatusid FROM egov_rentalreservationstatuses "
	sSql = sSql & " WHERE isinitialstatus = 1 AND orgid = '" & iOrgid & "'"
	'response.write sSql & "<br /><br />"

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
' date GetLastReservationMade( iCitizenUserId, sCategories )
'--------------------------------------------------------------------------------------------------
Function GetLastReservationMade( ByVal iCitizenUserId, ByVal sCategories )
	Dim sSql, oRs

	' Get the latest date they made a reservations in the restricted categories
	sSql = "SELECT TOP 1 R.reserveddate "
	sSql = sSql & " FROM egov_rentalreservations R, egov_rentalreservationstatuses S, "
	sSql = sSql & " egov_rentalreservationtypes T, egov_rentalreservationdates D, egov_rentals_to_categories C "
	sSql = sSql & " WHERE R.adminuserid IS NULL AND R.rentaluserid = " & iCitizenUserId
	sSql = sSql & " AND R.reservationid = D.reservationid AND D.rentalid = C.rentalid "
	sSql = sSql & " AND C.recreationcategoryid IN ( " & sCategories & " ) "
	sSql = sSql & " AND R.reservationtypeid = T.reservationtypeid AND reservationtypeselector = 'public' "
	sSql = sSql & " AND R.reservationstatusid = S.reservationstatusid AND S.isreserved = 1 "
	sSql = sSql & " ORDER BY R.reserveddate DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetLastReservationMade = CDate(DateValue(CDate(oRs("reserveddate"))))
	Else
		' nothing found so set to a month ago, so all is ok on the test
		GetLastReservationMade = DateAdd("m", -1, Date() )
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetLatestReservationHour( iRentalid, bOffSeasonFlag, iWeekday )
'--------------------------------------------------------------------------------------------------
Function GetLatestReservationHour( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday )
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
		GetLatestReservationHour = oRs("lateststarthour") & ":" & oRs("lateststartminute") & " " & oRs("lateststartampm")
	Else
		GetLatestReservationHour = ""
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
	oRs.Open sSql, Application("DSN"), 3, 1

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
' boolean GetMinimalTimeInterval iRentalid, bOffSeasonFlag, iWeekday, iInterval, sDateAddString
'--------------------------------------------------------------------------------------------------
Function GetMinimalTimeInterval( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByRef iInterval, ByRef sDateAddString )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(D.minimumrental,0) AS minimumrental, T.dateaddstring "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes T "
	sSql = sSql & " WHERE D.minimumrentaltimetypeid = T.timetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iInterval = clng(oRs("minimumrental"))
		sDateAddString = oRs("dateaddstring")
		GetMinimalTimeInterval = True 
	Else
		iInterval = 0
		sDateAddString = "h"
		GetMinimalTimeInterval = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetOffSeasonFlag( iRentalid, dCheckDate )
'--------------------------------------------------------------------------------------------------
Function GetOffSeasonFlag( ByVal iRentalid, ByVal dCheckDate )
	Dim sSql, oRs, dOffSeasonStartDate, dOffSeasonEndDate

	' See if we have an off season and if so is the date in the off season.
	' You want to return a flag indicating which season (in, off) to use.

	sSql = "SELECT hasoffseason, offseasonstartmonth, offseasonstartday, offseasonendmonth, offseasonendday, offseasonendyear "
	sSql = sSql & "FROM egov_rentals WHERE rentalid = " & iRentalid
	'response.write sSql & "<br /><br />"

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
' string GetOrgRentalRoundUpString( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetOrgRentalRoundUpString( ByVal iOrgId )
	Dim sSql, oRs, sRoundUp

	' Get the minimum interval for the org
	sSql = "SELECT ISNULL(O.rentalroundup,0) AS rentalroundup, ISNULL(T.dateaddstring,'n') AS dateaddstring "
	sSql = sSql & " FROM Organizations O, egov_rentaltimetypes T "
	sSql = sSql & " WHERE O.rentalrounduptimetypeid = T.timetypeid AND O.orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("rentalroundup")) > CLng(0) Then 
			sRoundUp = oRs("rentalroundup") & " "
			If oRs("dateaddstring") = "n" Then 
				sRoundUp = sRoundUp & " Minutes"
			Else
				sRoundUp = sRoundUp & " Hours"
			End If 
		Else
			sRoundUp = ""
		End If 
		GetOrdRentalRoundUpString = sRoundUp
	Else
		' No Round up
		GetOrdRentalRoundUpString = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetPaymentAccountId( iOrgId, iPaymentTypeId )
'------------------------------------------------------------------------------
Function GetPaymentAccountId( ByVal iOrgId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountid,0) AS accountid FROM egov_organizations_to_paymenttypes "
	sSql = sSql & "WHERE orgid = " & iOrgId & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPaymentAccountId = CLng(oRs("accountid"))
	Else
		GetPaymentAccountId = CLng(0) 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


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


'------------------------------------------------------------------------------
' string sRefundName = GetRefundName()
'------------------------------------------------------------------------------
Function GetRefundName( )
	Dim sSql, oRs

	sSql = "SELECT T.paymenttypename FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE T.isrefundmethod = 1 AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " & Session("OrgID") 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundName = oRs("paymenttypename")
	Else
		GetRefundName = "Refund Voucher"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' string GetRentalLocation( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetRentalLocation( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(L.name,'') AS location FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.locationid = L.locationid AND R.rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetRentalLocation = oRs("location")
	Else
		GetRentalLocation = ""
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' string GetRentalNameByRTI( iReservationTempId )
'--------------------------------------------------------------------------------------------------
Function GetRentalNameByRTI( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalname "
	sSql = sSql & " FROM egov_rentals R, egov_rentalreservationstemppublic P "
	sSql = sSql & " WHERE R.rentalid = P.rentalid AND P.reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetRentalNameByRTI = oRs("rentalname")
	Else
		GetRentalNameByRTI = ""
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
' integer GetRentalsInCategoryCount( iRecreationCategoryId )
'--------------------------------------------------------------------------------------------------
Function GetRentalsInCategoryCount( ByVal iRecreationCategoryId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(rentalid) AS hits FROM egov_rentals_to_categories WHERE recreationcategoryid = " & iRecreationCategoryId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRentalsInCategoryCount = CLng(oRs("hits"))
	Else
		GetRentalsInCategoryCount = CLng(0)
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
' string GetRentalSupervisorPhoneByRTI( iReservationTempId )
'--------------------------------------------------------------------------------------------------
Function GetRentalSupervisorPhoneByRTI( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(U.businessnumber,'') AS businessnumber "
	sSql = sSql & " FROM egov_rentalreservationstemppublic P, egov_rentals R, Users U "
	sSql = sSql & " WHERE P.rentalid = R.rentalid AND R.supervisoruserid = U.UserID "
	sSql = sSql & " AND P.reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRentalSupervisorPhoneByRTI = oRs("businessnumber")
	Else
		GetRentalSupervisorPhoneByRTI = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

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
' integer GetReservationTypeIdBySelector( sReservationTypeSelector )
'------------------------------------------------------------------------------
Function GetReservationTypeIdBySelector( ByVal sReservationTypeSelector )
	Dim sSql, oRs

	sSql = "SELECT reservationtypeid FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE orgid = " & iOrgId & " AND reservationtypeselector = '" & sReservationTypeSelector & "'"

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
' string GetResidentTypeByAddress( iUserid, iorgid )
'------------------------------------------------------------------------------
Function GetResidentTypeByAddress( ByVal iUserid, ByVal iorgid )
	' Try to match the person's address to one of the resident addresses
	Dim sSql, oRs
	
	GetResidentTypeByAddress = "N"

	sSql = "SELECT COUNT(R.residentaddressid) AS hits FROM egov_residentaddresses R, egov_users U"
	sSql = sSql & " WHERE R.orgid = U.orgid AND "
	sSql = sSql & " R.residentstreetnumber + ' ' + R.residentstreetname = U.useraddress AND "
	sSql = sSql & " R.residenttype = 'R' AND "
	sSql = sSql & " R.orgid = " & iorgid & " AND U.userid = " & iUserid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			' Match found
			GetResidentTypeByAddress = "R"
		End If 
	End if

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetResidentTypeDesc( sUserType )
'--------------------------------------------------------------------------------------------------
Function GetResidentTypeDesc( ByVal sUserType )
	Dim oCmd

	If Trim(sUserType <> "") Then 
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
	Else
		GetResidentTypeDesc = ""
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetRentalName( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetRentalName( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT rentalname FROM egov_rentals "
	sSql = sSql & "WHERE rentalid = " & iRentalId

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
' string GetRentalTerms( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetRentalTerms( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(terms,'') AS terms FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRentalTerms = oRs("terms")
	Else
		GetRentalTerms = "" 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' void GetSelectedDate iRti, iRentalId, sSelectedDate
'--------------------------------------------------------------------------------------------------
Sub GetSelectedDate( ByVal iRti, ByRef iRentalId, ByRef sSelectedDate )
	Dim sSql, oRs

	sSql = "SELECT rentalid, selecteddate "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iRti

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iRentalId = oRs("rentalid")
		sSelectedDate = oRs("selecteddate")
	Else
		' this is unlikely and a real problem
		iRentalId = 0
		sSelectedDate = Date()
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' String GetUserResidentType( iUserId )
'------------------------------------------------------------------------------
Function GetUserResidentType( ByVal iUserId )
  lcl_return = "N"

  sSQL = "SELECT isnull(residenttype,'N') AS residenttype,useraddress "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = '" & iUserId & "'"

 	set rs = Server.CreateObject("ADODB.Recordset")
	 rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_return = rs("residenttype")
     if rs("useraddress") = "" or isnull(rs("useraddress")) then lcl_return = "Z"
  end if

  set rs = nothing

  getUserResidentType = lcl_return

End Function 


'--------------------------------------------------------------------------------------------------
' void GetWantedDateAndTimes iRti, iRentalId, sStartDateTime, sEndDateTime
'--------------------------------------------------------------------------------------------------
Sub GetWantedDateAndTimes( ByVal iRti, ByRef iRentalId, ByRef sStartDateTime, ByRef sEndDateTime )
	Dim sSql, oRs

	sSql = "SELECT rentalid, selecteddate, starthour, startminute, startampm, endhour, endminute, endampm "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iRti

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iRentalId = oRs("rentalid")
		sStartDateTime = CDate(oRs("selecteddate") & " " & oRs("starthour") & ":" & oRs("startminute") & " " & oRs("startampm"))
		sEndDateTime = CDate(oRs("selecteddate") & " " & oRs("endhour") & ":" & oRs("endminute") & " " & oRs("endampm"))
		If sEndDateTime < sStartDateTime Then
			sEndDateTime = DateAdd("d", 1, sEndDateTime)
		End If 
	Else
		' this is unlikely and a real problem
		iRentalId = 0
		sStartDateTime = Date()
		sEndDateTime = Date()
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean PublicCanReserveRental( iRentalId )
'--------------------------------------------------------------------------------------------------
Function PublicCanReserveRental( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT publiccanreserve FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("publiccanreserve") Then 
			PublicCanReserveRental = True 
		Else
			PublicCanReserveRental = False 
		End If 
	Else
		' this is an error
		PublicCanReserveRental = False 
	End If 

	oRs.Close
	Set oRs = Nothing 


End Function 


'--------------------------------------------------------------------------------------------------
' boolean PublicCanReserveRentalsInCategory( iRecreationCategoryId )
'--------------------------------------------------------------------------------------------------
Function PublicCanReserveRentalsInCategory( ByVal iRecreationCategoryId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(R.rentalid) AS hits FROM egov_rentals R, egov_rentals_to_categories C "
	sSql = sSql & " WHERE R.rentalid = C.rentalid AND R.publiccanreserve = 1 AND publiccanview = 1 "
	sSql = sSql & " AND recreationcategoryid = " & iRecreationCategoryId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PublicCanReserveRentalsInCategory = True 
		Else
			PublicCanReserveRentalsInCategory = False 
		End If 
	Else
		' this is an error
		PublicCanReserveRentalsInCategory = False 
	End If 

	oRs.Close
	Set oRs = Nothing 


End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasCurrentSeasonRestriction( iRentalId )
'--------------------------------------------------------------------------------------------------
Function RentalHasCurrentSeasonRestriction( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT reservationsduringseason FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("reservationsduringseason") Then 
			RentalHasCurrentSeasonRestriction = True 
		Else
			RentalHasCurrentSeasonRestriction = False 
		End If 
	Else
		' this is an error
		RentalHasCurrentSeasonRestriction = False 
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
' boolean RentalIsAvailableToPublic( iRentalid, bOffSeasonFlag, iWeekday )
'--------------------------------------------------------------------------------------------------
Function RentalIsAvailableToPublic( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday )
	Dim sSql, oRs

	sSql = "SELECT isavailabletopublic FROM egov_rentaldays WHERE rentalid = " & iRentalid
	sSql = sSql & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekday

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' First see if the rental is available on this day of the week
		If oRs("isavailabletopublic") Then 
			RentalIsAvailableToPublic = True  
		Else
			RentalIsAvailableToPublic = False  
		End If
	Else
		' we do not have a day record, that is bad
		RentalIsAvailableToPublic = False  
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
' boolean ShowAvailability( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
'--------------------------------------------------------------------------------------------------
Function ShowAvailability( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dStartDate, ByVal bShowCall, ByRef bShowCallMsg )
	Dim sSql, oRs, dOpeningTime, dClosingTime, dLastStart, iCount, iMinInterval, sDateAddString, bIsAllDay
	Dim iPostBuffer, iAvailableTimeBlock, dLatestAllowed
	
	bShowCallMsg = false 

	GetOpeningAndClosingTimes iRentalid, bOffSeasonFlag, iWeekday, dStartDate, dOpeningTime, dClosingTime

	' Get Minimal Time interval Info for this day
	bHasMinimum = GetMinimalTimeInfo( iRentalid, bOffSeasonFlag, iWeekday, iMinInterval, sDateAddString, bIsAllDay )

	If Not bIsAllDay Then 
		' convert this interval into minutes if needed
		If sDateAddString = "h" Then
			iMinInterval = CLng(iMinInterval) * 60
		End If 

		' get the post buffer for this day, if any
		GetPostBufferTime iRentalid, bOffSeasonFlag, iWeekday, iPostBuffer, sDateAddString

		If sDateAddString = "h" Then
			' convert to minutes
			iPostBuffer = CLng(iPostBuffer) * 60
		End If 

		' add the post buffer to the minimum allowed time
		'iMinInterval = CLng(iMinInterval) + CLng(iPostBuffer)
	End If 

	sSql = "SELECT reservationstarttime, reservationendtime, reservationid " 
	sSql = sSql & " FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
	sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND reservationstarttime > '" & DateValue(dStartDate) & " 0:00 AM' "
	sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dStartDate)) & " 0:00 AM' "
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	dLastStart = dOpeningTime
	iCount = clng(0)
	Do While Not oRs.EOF
		If clng(iPostBuffer) > clng(0) Then 
			dReservationStartTime = DateAdd("n", -(iPostBuffer), CDate(oRs("reservationstarttime")) )
		Else
			dReservationStartTime = CDate(oRs("reservationstarttime"))
		End If 

		'response.write "dReservationStartTime = " & dReservationStartTime & "<br />"
		If dLastStart < CDate(oRs("reservationstarttime")) Then 
			If Not bIsAllDay Then 
				'response.write "dLastStart = " & dLastStart & "<br />"
				iAvailableTimeBlock = DateDiff("n", dLastStart, dReservationStartTime)
				'response.write "iAvailableTimeBlock = " & iAvailableTimeBlock & "<br />"
				'response.write "iMinInterval = " & iMinInterval & "<br />"
				' if the available time >= minimum allowed time then output the string
				If (bHasMinimum = False) Or (bHasMinimum = True And iAvailableTimeBlock >= iMinInterval) Then 
					iCount = iCount + 1
					If iCount > clng(1) Then
						response.write "<br />"
					End If 
					' display time slot available
					response.write FormatTimeString( dLastStart ) & " to " & FormatTimeString( dReservationStartTime ) '& " - " & iAvailableTimeBlock & " Min"
				End If 
			End If 
		End If 
		dLastStart = CDate(oRs("reservationendtime"))
		
		' Look for the show call message flag'
		If bShowCall And bShowCallMsg = false Then 
			' Once this is set to true, we do not want to unset it'
			bShowCallMsg = getReservationCallFlag( CLng(oRs("reservationid")))
		End If
		
		oRs.MoveNext
	Loop

	If dLastStart < dClosingTime Then 
		iAvailableTimeBlock = DateDiff("n", dLastStart, dClosingTime)
		If (bHasMinimum = False) Or (bHasMinimum = True And iAvailableTimeBlock >= iMinInterval) Then 
			' check that the last start time is before the latest reservation start time
			dLatestAllowed = GetLatestReservationTime( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
'			response.write "<br />" & dLatestAllowed & "<br />"
'			response.write dClosingTime & "<br />"
'			response.write DateDiff("n", dLastStart, dLatestAllowed) & "<br />"
			If DateDiff("n", dLastStart, dLatestAllowed) >= 0 Then 
				iCount = iCount + 1
				If iCount > clng(1) Then
					response.write "<br />"
				End If 
				' Display end of day time slot available
				response.write FormatTimeString( dLastStart ) & " to " & FormatTimeString( dClosingTime ) '& " - " & iAvailableTimeBlock & " Min"
			End If 
		End If 
	End If 

	If clng(iCount) = clng(0) Then 
		response.write "<span class=""noreservemsg"">Unavailable</span>"
		ShowAvailability = False 
	Else 
		ShowAvailability = True 
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


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
' boolean getReservationCallFlag( iReservationId )
'--------------------------------------------------------------------------------------------------
Function getReservationCallFlag( ByVal iReservationId )
	Dim sSql, oRs
	
	sSql = "SELECT iscall FROM egov_rentalreservations WHERE reservationid = " & iReservationId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("iscall") Then 
			getReservationCallFlag = True 
		Else
			getReservationCallFlag = False
		End If 
	Else
		getReservationCallFlag = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ShowMinimumReservationTime( iRentalid, bOffSeasonFlag, iWeekday )
'--------------------------------------------------------------------------------------------------
Function ShowMinimumReservationTime( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(D.minimumrental,0) AS minimumrental, T.timetype, T.isallday, "
	sSql = sSql & " ISNULL(D.postbuffer,0) AS postbuffer, R.timetype AS buffertimetype "
	sSql = sSql & " FROM egov_rentaldays D, egov_rentaltimetypes T, egov_rentaltimetypes R "
	sSql = sSql & " WHERE D.minimumrentaltimetypeid = T.timetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.postbuffertimetypeid = R.timetypeid "
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "<p>The minimum reservation time is "
		If oRs("isallday") Then
			response.write "all day."
			ShowMinimumReservationTime = True 
		Else
			response.write oRs("minimumrental") & " " & oRs("timetype") & "."
			ShowMinimumReservationTime = False 
		End If 
		' Will move the post buffer onto the front of available times so the citizen does not have to worry about it.
'		If clng(oRs("postbuffer")) > clng(0) Then 
'			response.write "<br />There is a minimum period of " & oRs("postbuffer") & " " & oRs("buffertimetype")
'			response.write " required between each reservation.<br />You need to allow for this time, but not include it, in your reservation."
'		End If 
		response.write "</p>"
	Else
		ShowMinimumReservationTime = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


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
' void ShowRentalNameAndLocation iRentalId 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalNameAndLocation( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalname, L.name AS location FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.locationid = L.locationid AND R.rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write "<span class=""schedulerentalname"">" & oRs("location") & " &ndash; " & oRs("rentalname") & "</span>"
	Else
		response.write "Rental Not Found"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub

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
'--------------------------------------------------------------------------------------------------
' double GetTotalAmount( iReservationTempId )
'--------------------------------------------------------------------------------------------------
Function GetTotalAmount( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(feetotal,0.0000) AS feetotal FROM egov_rentalreservationstemppublic "
	sSql = sSql & "WHERE reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetTotalAmount = CDbl(oRs("feetotal"))
	Else
		GetTotalAmount = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

%>
