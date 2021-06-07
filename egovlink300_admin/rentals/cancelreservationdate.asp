<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: cancelreservationdate.asp
' AUTHOR: Steve Loar
' CREATED: 10/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed dates and times are OK, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/20/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iReservationDateId, iReservationId, dTotalCharges, iCancelStatusId, sReservationTypeSelection

'iReservationDateId = CLng(request("reservationdateid"))
arrReservationDateIDs = split(request("reservationdateid"),",")

for x = 0 to UBOUND(arrReservationDateIDs)

	iReservationDateId = CLng(arrReservationDateIDs(x))

	' Get the reservation id
	iReservationId = GetReservationIdFromDateId( iReservationDateId )
	
	CancelReservationDate iReservationDateId
	
	' Recalculate the total charges
	dTotalCharges = CalculateReservationTotal( iReservationId, "feeamount" )	' in rentalscommonfunctions.asp
	
	' Update the total for the reservation
	sSql = "UPDATE egov_rentalreservations SET totalamount = " & dTotalCharges
	sSql = sSql & " WHERE reservationid = " & iReservationId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	
	' See if there are any more dates and if not then cancel the reservation
	If Not ReservationHasReservedDates( iReservationId ) Then 
		' Zero out any reservation level fees
		sSql = "UPDATE egov_rentalreservationfees SET feeamount = 0.0000 WHERE reservationid = " & iReservationId
		RunSQLStatement sSql
	
		' get the cancel status id
		iCancelStatusId = GetReservationStatusId( "iscancelled" )
	
		' Set the Reservation status to cancelled
		sSql = "UPDATE egov_rentalreservations SET reservationstatusid = " & iCancelStatusId
		sSql = sSql & ", isonhold = 0 "
		sSql = sSql & " WHERE reservationid = " & iReservationId
		RunSQLStatement sSql
	
		' Clear the reservationid on any class time rows
		ClearClassTimeReservation iReservationId
	Else 
		' The reservation has other days so if this date has not been paid for, or is class or internal then just remove it totally from the reservation
	'	sReservationTypeSelection = GetReservationTypeSelection( GetReservationTypeId( iReservationId ) )
		'If sReservationTypeSelection = "public" Or sReservationTypeSelection = "admin" Or sReservationTypeSelection = "class" Or sReservationTypeSelection = "block" Then
			' If the amount of fees paid for the date is $0.00 then delete the rows
			If ( CDbl(GetReservationDateFees( iReservationDateId, "paidamount" )) = CDbl(0) ) And ( CDbl(GetReservationDateItemFees( iReservationDateId, "paidamount" )) = CDbl(0) ) Then 
				RemoveReservationDate iReservationDateId 
			End If 
		'End If 
	
	End If 
	
next
response.write "cd"

%>
