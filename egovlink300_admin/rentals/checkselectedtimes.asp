<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkselectedtimes.asp
' AUTHOR: Steve Loar
' CREATED: 06/15/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed dates and times are OK, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   06/`5/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iRentalId, iMaxRows, x, iReservationTempId, iEndDay, sStartDateTime, sEndDateTime
Dim iPosition, iEndHour, iEndMinute, sEndAmPm, sReturn, sOffSeasonFlag, iReservationTypeId, iRentalUserid

iReservationId = CLng(request("reservationid"))

iReservationDateId = CLng(request("reservationdateid"))

iRentalId = CLng(request("rentalid"))

iReservationTypeId = GetReservationTypeId( iReservationId )

sReturn = "OK"

sStartDateTime = request("startdate") & " " & request("starthour") & ":" & request("startminute") & " " & request("startampm")
iEndDay = request("endday" & x)
sEndDateTime = request("startdate") & " " & request("endhour") & ":" & request("endminute") & " " & request("endampm")
If iEndDay = "1" Then 
	sEndDateTime = CStr(DateAdd("d", 1, CDate(sEndDateTime)))
End If 

If IsReservation( iReservationTypeId ) Then 
	' Round up as required by the org to the next wanted interval
	CheckOrgRentalRoundUp sStartDateTime, sEndDateTime, iEndHour, iEndMinute, sEndAmPm
Else
	SetEndingTimes sEndDateTime, iEndHour, iEndMinute, sEndAmPm 
End If 

If sReturn = "OK" Then
	sReturn = CheckRentalTimeAvailability( iRentalid, sStartDateTime, sEndDateTime, iReservationDateId, iReservationId )	
End If 

If sReturn = "" Then 
	sReturn = "OK"
End If 
 
response.write sReturn



'--------------------------------------------------------------------------------------------------
' string CheckRentalTimeAvailability( iRentalid, sStartDateTime, sEndDateTime, iReservationDateId, iReservationId )
'--------------------------------------------------------------------------------------------------
Function CheckRentalTimeAvailability( ByVal iRentalid, ByVal sStartDateTime, ByVal sEndDateTime, ByVal iReservationDateId, ByVal iReservationId )
	Dim sOffSeasonFlag, sCheckReturn, sReturn, dWantedEndTime

	sReturn = ""

	sOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )

	' Check if rental is open that day and that we are not looking for a time when it is not open
	sCheckReturn = CheckRentalHours( iRentalid, sStartDateTime, sEndDateTime, "selectedperiod", sOffSeasonFlag )
	
	If Left(sCheckReturn,2) = "No" Then 
		'sReturn = sStartDateTime & "closed" & sCheckReturn
		sReturn = "closed"
	End If 

	If sReturn <> "closed" Then 
		' Check if the time is available without the end buffer
		If CheckForConflictingReservations( iRentalid, sStartDateTime, sEndDateTime, sOffSeasonFlag, iReservationDateId, iReservationId, False ) Then
			sReturn = "conflict" 
		End If 
	End If 

	If sReturn = "" Then 
		' Then set it to the OK flag
		sReturn = "OK"
		'response.write ReservationNeedsBufferTimeAdded( iReservationId ) & " "
		If ReservationNeedsBufferTimeAdded( iReservationId ) Then 
			dWantedEndTime = AddPostBufferTime( iRentalid, sOffSeasonFlag, sEndDateTime, sStartDateTime )
			' Check if the time is available with the buffer included
			If CheckForConflictingReservations( iRentalid, sStartDateTime, dWantedEndTime, sOffSeasonFlag, iReservationDateId, iReservationId, True ) Then
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

	CheckRentalTimeAvailability = sReturn

End Function 


'--------------------------------------------------------------------------------------------------
' string CheckForConflictingReservations( iRentalid, dWantedStartTime, dWantedEndTime, bOffSeasonFlag, iReservationDateId, iReservationId, bIncludeBuffer )
'--------------------------------------------------------------------------------------------------
Function CheckForConflictingReservations( ByVal iRentalid, ByVal dWantedStartTime, ByVal dWantedEndTime, ByVal bOffSeasonFlag, ByVal iReservationDateId, ByVal iReservationId, ByVal bIncludeBuffer )
	Dim sSql, oRs, sCompareEndTime

'	bIncludeBuffer = ReservationNeedsBufferTimeAdded( iReservationId )

'	If EndTimeIsNotClosingTime( iRentalId, dWantedEndTime, bOffSeasonFlag, dWantedStartTime ) Then
'		If bIncludeBuffer Then 
			' Add on the end buffer
'			dWantedEndTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, dWantedEndTime, dWantedStartTime )
'		End If 
'	End If 

	If bIncludeBuffer Then
		sCompareEndTime = "reservationendtime"
	Else
		sCompareEndTime = "billingendtime"
	End If 

	' we will add a minute to this so start time can be the same as the end of bufferend time of another reservation
	dWantedStartTime = DateAdd("n", 1, dWantedStartTime)
	' we will remove a minute so the end of the buffer can be the same minute as the start of another reservation
	dWantedEndTime = DateAdd("n", -1, dWantedEndTime)

	' set sql to look for conflicting times
	sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
	sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND (reservationstarttime BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
	sSql = sSql & " OR " & sCompareEndTime & " BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
	sSql = sSql & " OR (reservationstarttime <= '" & dWantedStartTime & "' AND " & sCompareEndTime & " >= '" & dWantedEndTime & "'))"
	sSql = sSql & " AND reservationdateid != " & iReservationDateId

	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			' conflicting reservation times
			CheckForConflictingReservations = True 
		Else
			' No conflicts found
			CheckForConflictingReservations = False  
		End If 
	Else
		' No rows returned - not likely using count(), but still no conflicts
		CheckForConflictingReservations = False  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>