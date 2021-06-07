<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: classreservationmake.asp
' AUTHOR: Steve Loar
' CREATED: 10/21/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Creates The initial Rental reservation.
'
' MODIFICATION HISTORY
' 1.0   10/21/2009	Steve Loar - INITIAL VERSION
' 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, sSql, oRs, iReservationId, iReservationTempId, iRentalUserid, sReservationTypeId
Dim iInitialStatusId, iAdminUserId, iMaxRows, x, sStartDateTime, iEndDay, sBillingEndDateTime
Dim sEndDateTime, iReservationDateId, iWeekday, bIsReservation, sUserType, bPublicReservation
Dim dTotalAmount, iTimeId, iTimeDayId

iRentalId = CLng(request("rentalid"))

iTimeId = CLng(request("timeid"))

iReservationTempId = CLng(request("rti"))

sReservationTypeId = CLng(request("reservationtypeid"))

bIsReservation = IsReservation( sReservationTypeId ) ' should always be false

dTotalAmount = CDbl(0.0000) ' who cares about this for classes

sReservationTypeSelection = GetReservationTypeSelection( sReservationTypeId ) ' should always be "class"
response.write "sReservationTypeSelection = " & sReservationTypeSelection & "<br /><br />"

'If bIsReservation Then 
'	' Need to know if they are admin or public
'	iRentalUserid = CLng(request("rentaluserid"))
'	If sReservationTypeSelection = "public" Then 
'		sUserType = GetUserResidentType( iRentalUserid )
'		'If they are not one of these (R, N), we have to figure which they are
'		If sUserType <> "R" And sUserType <> "N" Then 
'			'This leaves E and B - See if they are a resident, also
'			sUserType = GetResidentTypeByAddress( iRentalUserid, Session("OrgID") )
'		End If 
'	Else
'		' Admin type
'		sUserType = "E" ' employee
'	End If 
'Else 
	' Blocked type
'	iRentalUserid = "NULL"
'	sUserType = "E"
'End If 

iInitialStatusId = GetInitialReservationStatusId()

iAdminUserId = session("userid")

' Create the reservation row
sSql = "INSERT INTO egov_rentalreservations ( orgid, reservationtypeid, reservationstatusid, rentaluserid, "
sSql = sSql & "adminuserid, reserveddate, timeid, originalrentalid ) VALUES ( " & session("OrgID") & ", " & sReservationTypeId & ", "
sSql = sSql & iInitialStatusId & ", NULL, " & iAdminUserId & ", dbo.GetLocalDate("
sSql = sSql & Session("OrgID") & ",getdate()), " & iTimeId & ", " & iRentalId & " )"
response.write sSql & "<br /><br />"
iReservationId = RunInsertStatement( sSql )


' Update the class time table with the reservationid 
sSql = "UPDATE egov_class_time SET reservationid = " & iReservationId & " WHERE timeid = " & iTimeId
response.write sSql & "<br /><br />"
RunSQLStatement sSql


' Get the dates from the passed values
iMaxRows = CLng(request("maxrows"))

For x = 1 To iMaxRows
	If request("startdate" & x) <> "" Then 
		If request("includereservationtime" & x) = "on" Then 
			sStartDateTime = request("startdate" & x) & " " & request("starthour" & x) & ":" & request("startminute" & x) & " " & request("startampm" & x)
			iEndDay = request("endday" & x)
			bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )
			sBillingEndDateTime = request("startdate" & x) & " " & request("endhour" & x) & ":" & request("endminute" & x) & " " & request("endampm" & x)
			iTimeDayId = request("timedayid" & x)
			If iEndDay = "1" Then 
				sBillingEndDateTime = CStr(DateAdd("d", 1, CDate(sBillingEndDateTime)))
			End If 
			' The real end time includes the end buffer if this is not to closing time for the rental
	'		If EndTimeIsNotClosingTime( iRentalId, sBillingEndDateTime, bOffSeasonFlag, sStartDateTime ) Then 
				' Add on the end buffer
	'			sEndDateTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, sBillingEndDateTime, sStartDateTime )
	'		Else
				' For classes the end is the end, but maybe not??
			sEndDateTime = sBillingEndDateTime
	'		End If 

			sSql = "INSERT INTO egov_rentalreservationdates ( reservationid, rentalid, orgid, statusid, reservationstarttime, "
			sSql = sSql & "reservationendtime, billingendtime, actualstarttime, actualendtime, adminuserid, reserveddate, timedayid ) VALUES ( "
			sSql = sSql & iReservationId & ", " & iRentalid & ", " & Session("OrgID") & ", " & iInitialStatusId & ", '" & sStartDateTime & "', '"
			sSql = sSql & sEndDateTime & "', '" & sBillingEndDateTime & "', '" & sStartDateTime & "', '" & sBillingEndDateTime & "', "
			sSql = sSql & iAdminUserId & ", " & "dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), " & iTimeDayId & " )"
			'response.write sSql & "<br /><br />"
			iReservationDateId = RunInsertStatement( sSql )

			' Create the rental reservation date items rows - These are things like tables and chairs
			CreateRentalReservationDateItems iReservationDateId, iReservationId, iRentalid, sReservationTypeSelection
		End If 
	End If 
Next 


' delete the temp data
sSql = "DELETE FROM egov_rentalreservationdatestemp WHERE reservationtempid = " & iReservationTempId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

sSql = "DELETE FROM egov_rentalreservationstemp WHERE reservationtempid = " & iReservationTempId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Take them to the edit page for this reservation
response.redirect "reservationedit.asp?reservationid=" & iReservationId & "&sf=rc"


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------



%>

