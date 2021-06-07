<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationreinstate.asp
' AUTHOR: Steve Loar
' CREATED: 08/22/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This reinstates a reservation, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/22/2011	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, iReservationId, iReservedStatusId

iReservationId = CLng(request("reservationid"))

' get the reserved status id
iReservedStatusId = GetReservationStatusId( "isreserved" )

' Set the Reservation status to reserved
sSql = "UPDATE egov_rentalreservations SET reservationstatusid = " & iReservedStatusId
sSql = sSql & " WHERE reservationid = " & iReservationId & " AND orgid = " & session("orgid")
RunSQLStatement sSql

' set the reservation dates to reserved 
sSql = "UPDATE egov_rentalreservationdates SET statusid = " & iReservedStatusId 
sSql = sSql & " WHERE reservationid = " & iReservationId & " AND orgid = " & session("orgid")
RunSQLStatement sSql

' Send back the completed flag to the calling script
response.write "re"

%>

