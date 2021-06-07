<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addnewdatesave.asp
' AUTHOR: Steve Loar
' CREATED: 11/10/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This adds new dates to a reservation, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   11/10/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iRentalId, iReservationId, iInitialStatusId, iAdminUserId, sStartDateTime, iEndDay
Dim bOffSeasonFlag, sBillingEndDateTime, sEndDateTime, iReservationDateId, bIsReservation, iWeekday
Dim sReservationTypeSelection, iReservationTypeId, dTotalAmount, dTotalPayment, dTotalRefunds

iReservationId = CLng(request("reservationid"))

iRentalId = CLng(Mid(request("rentalid"),2))

iInitialStatusId = GetInitialReservationStatusId()

iAdminUserId = session("userid")

iReservationTypeId = GetReservationTypeId( iReservationId )

bIsReservation = IsReservation( iReservationTypeId ) 

sReservationTypeSelection = GetReservationTypeSelection( iReservationTypeId )

If bIsReservation Then 
	' Need to know if they are admin or public
	iRentalUserid = GetReservationRentalUserId( iReservationId )
	If sReservationTypeSelection = "public" Then 
		sUserType = GetUserResidentType( iRentalUserid )
		'If they are not one of these (R, N), we have to figure which they are
		If sUserType <> "R" And sUserType <> "N" Then 
			'This leaves E and B - See if they are a resident, also
			sUserType = GetResidentTypeByAddress( iRentalUserid, Session("OrgID") )
		End If 
	Else
		' Admin type
		sUserType = "E" ' employee
	End If 
Else 
	' Blocked type
	iRentalUserid = "NULL"
	sUserType = "E"
End If 

'dTotalAmount = GetReservationTotalAmount( iReservationId )

sStartDateTime = request("startdate" & x) & " " & request("starthour" & x) & ":" & request("startminute" & x) & " " & request("startampm" & x)
iEndDay = request("endday" & x)
bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(sStartDateTime)) )
sBillingEndDateTime = request("startdate" & x) & " " & request("endhour" & x) & ":" & request("endminute" & x) & " " & request("endampm" & x)
If iEndDay = "1" Then 
	sBillingEndDateTime = CStr(DateAdd("d", 1, CDate(sBillingEndDateTime)))
End If 
' The real end time includes the end buffer if this is not to closing time for the rental
If EndTimeIsNotClosingTime( iRentalId, sBillingEndDateTime, bOffSeasonFlag, sStartDateTime ) Then 
	' Add on the end buffer
	sEndDateTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, sBillingEndDateTime, sStartDateTime )
Else
	sEndDateTime = sBillingEndDateTime
End If 

sSql = "INSERT INTO egov_rentalreservationdates ( reservationid, rentalid, orgid, statusid, reservationstarttime, "
sSql = sSql & "reservationendtime, billingendtime, actualstarttime, actualendtime, adminuserid, reserveddate ) VALUES ( "
sSql = sSql & iReservationId & ", " & iRentalid & ", " & Session("OrgID") & ", " & iInitialStatusId & ", '" & sStartDateTime & "', '"
sSql = sSql & sEndDateTime & "', '" & sBillingEndDateTime & "', '" & sStartDateTime & "', '" & sBillingEndDateTime & "', "
sSql = sSql & iAdminUserId & ", " & "dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) )"
'response.write sSql & "<br /><br />"
iReservationDateId = RunInsertStatement( sSql )

If bIsReservation Then 
	iWeekday = Weekday(CDate(sStartDateTime))
	If sReservationTypeSelection <> "block" And sReservationTypeSelection <> "class" Then 
		' Create the rental reservation date fees rows
		CreateRentalReservationDateFees iReservationDateId, iReservationId, iRentalid, bOffSeasonFlag, iWeekday, sUserType, sStartDateTime, sBillingEndDateTime, dTotalAmount, sReservationTypeSelection

		If sReservationTypeSelection = "admin" Then
			' Add to the account ledger for the existing payment
			CreateReservationDateAccountLedgerEntries iReservationDateId, iReservationId
		End If 
	End If 
End If 

If sReservationTypeSelection <> "block" Then 
	' Create the rental reservation date items rows - These are things like tables and chairs
	CreateRentalReservationDateItems iReservationDateId, iReservationId, iRentalid, sReservationTypeSelection
End If 

' Get the fee total amount from the tables, this gives a more accurate value
dTotalCharges = CalculateReservationTotal( iReservationId, "feeamount" )	' in rentalscommonfunctions.asp
dTotalRefunds = CalculateReservationTotal( iReservationId, "paidamount" )
dTotalPayment = CalculateReservationTotal( iReservationId, "refundamount" )

' Update the total amount due on the reservation
sSql = "UPDATE egov_rentalreservations SET totalamount = " & dTotalCharges
sSql = sSql & ", totalpaid = " & dTotalPayment
sSql = sSql & ", totalrefunded = " & dTotalRefunds
sSql = sSql & " WHERE reservationid = " & iReservationId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql


response.write "Success"


'--------------------------------------------------------------------------------------------------
'  void CreateReservationDateAccountLedgerEntries iReservationDateId, iReservationId
'--------------------------------------------------------------------------------------------------
Sub CreateReservationDateAccountLedgerEntries( ByVal iReservationDateId, ByVal iReservationId )
	Dim sSql, oRs, iAccountId, iItemTypeID

	iItemTypeID = GetItemTypeId( "rentals" )

	sSql = "SELECT F.reservationdatefeeid, P.paymentid "
	sSql = sSql & "FROM egov_rentalreservationdatefees F, egov_class_payment P "
	sSql = sSql & "WHERE F.reservationid = P.reservationid AND P.isforrentals = 1 "
	sSql = sSql & "AND F.reservationdateid = " & iReservationDateId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF

		iAccountId = GetReservationAccountId( oRs("reservationdatefeeid"), "reservationdatefeeid", "egov_rentalreservationdatefees" )	' In rentalscommonfunctions.asp

		' Add to Accounts Ledger Row
		sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
		sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, reservationfeetypeid, reservationfeetype, reservationdateid ) VALUES ( "
		sSql = sSql & oRs("paymentid") & ", " & session("orgid") & ", 'credit', " & iAccountId & ", 0.00, " & iItemTypeId & ", '+', " 
		sSql = sSql & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", " & oRs("reservationdatefeeid") & ", 'reservationdatefeeid', " & iReservationDateId & " )"
		response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>