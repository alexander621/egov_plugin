<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: cancelreservation.asp
' AUTHOR: Steve Loar
' CREATED: 10/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This cancels a reservation, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/20/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iReservationId, dTotalCharges, iCancelStatusId, oRs, iReservedStatusId

arrReservationIDs = split(request("reservationid"),",")

for x = 0 to UBOUND(arrReservationIDs)

	iReservationId = CLng(arrReservationIDs(x))
	
	' get the reserved status id
	iReservedStatusId = GetReservationStatusId( "isreserved" )
	
	' Loop through the reservation dates that are not already cancelled and cancel them
	sSql = "SELECT reservationdateid FROM egov_rentalreservationdates "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND statusid = " & iReservedStatusId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	Do While Not oRs.EOF
		CancelReservationDate oRs("reservationdateid")
		oRs.MoveNext
	Loop
	
	oRs.Close 
	Set oRs = Nothing 
	
	' Zero out any reservation level fees
	sSql = "UPDATE egov_rentalreservationfees SET feeamount = 0.0000 WHERE reservationid = " & iReservationId
	RunSQLStatement sSql
	
	' Recalculate the total charges - THese should be $0 at this point
	dTotalCharges = CalculateReservationTotal( iReservationId, "feeamount" )	' in rentalscommonfunctions.asp
	
	' Update the total for the reservation
	sSql = "UPDATE egov_rentalreservations SET totalamount = " & dTotalCharges
	sSql = sSql & ", isonhold = 0 "
	sSql = sSql & " WHERE reservationid = " & iReservationId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	
	' get the cancel status id
	iCancelStatusId = GetReservationStatusId( "iscancelled" )
	
	' Set the Reservation status to cancelled
	sSql = "UPDATE egov_rentalreservations SET reservationstatusid = " & iCancelStatusId
	sSql = sSql & " WHERE reservationid = " & iReservationId
	RunSQLStatement sSql
	
	' Clear the reservationid on any class time rows
	ClearClassTimeReservation iReservationId
	
next

response.write "cr"

%>
