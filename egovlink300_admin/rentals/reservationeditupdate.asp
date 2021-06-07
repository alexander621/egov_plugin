<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationeditupdate.asp
' AUTHOR: Steve Loar
' CREATED: 11/09/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates Reservations. Called from reservationedit.asp
'
' MODIFICATION HISTORY
' 1.0   11/09/2009	Steve Loar - INITIAL VERSION
' 1.1	06/14/2010	Steve Loar - changed internal reservations to automatically adust the account ledger 
'								 to have the new counts (add or update) or delete the entry if set to 0.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationId, iReservationItemCount, iRentalRateCount, iReservationFeeCount, x, sSql
Dim dTotalCharges, sOrganization, sPointOfContact, sNumberAttending, sPurpose, sReceiptNotes
Dim sPrivateNotes, iReservationDateCount, dTotalPayment, dTotalRefunds, sHoldFlag, sCallFlag
Dim sReservationTypeSelection, iReservationDateId, iPaymentId, iAccountId

iReservationId = CLng(request("reservationid"))
dTotalCharges = CDbl(0.0000)

iReservationItemCount = CLng(request("maxreservationitems"))
iRentalRateCount = CLng(request("maxrentalrates"))
iReservationFeeCount = CLng(request("maxreservationfees"))
iReservationDateCount = CLng(request("maxreservationdates"))

If request("organization") <> "" Then
	sOrganization = "'" & dbsafe(request("organization")) & "'"
Else
	sOrganization = "NULL"
End If 

If request("pointofcontact") <> "" Then
	sPointOfContact = "'" & dbsafe(request("pointofcontact")) & "'"
Else
	sPointOfContact = "NULL"
End If 

If request("numberattending") <> "" Then
	sNumberAttending = "'" & dbsafe(request("numberattending")) & "'"
Else
	sNumberAttending = "NULL"
End If 

If request("purpose") <> "" Then
	sPurpose = "'" & dbsafe(request("purpose")) & "'"
Else
	sPurpose = "NULL"
End If 

If request("receiptnotes") <> "" Then
	sReceiptNotes = "'" & dbsafe(request("receiptnotes")) & "'"
Else
	sReceiptNotes = "NULL"
End If 

If request("privatenotes") <> "" Then
	sPrivateNotes = "'" & dbsafe(request("privatenotes")) & "'"
Else
	sPrivateNotes = "NULL"
End If 

If LCase(request("isonhold")) = "on" Then 
	sHoldFlag = "1"
Else
	sHoldFlag = "0"
End If 

If LCase(request("isabusive")) = "on" Then 
	sAbuseFlag = "1"
Else
	sAbuseFlag = "0"
End If 
sAbuseNote = DBSafe(request("abusenote"))
if request("abuseuserid") <> "" then
	sAbuseUserID = request("abuseuserid")
else
	sAbuseUserID = "0"
end if

If LCase(request("iscall")) = "on" Then 
	sCallFlag = "1"
Else
	sCallFlag = "0"
End If

sReservationTypeSelection = GetReservationTypeSelection( GetReservationTypeId( iReservationId ) )


' Save the abuse info
sSQL = "UPDATE egov_users SET facilityabuse = '" & sAbuseFlag & "', facilityabusenote = '" & sAbuseNote & "' WHERE userid = '" & sAbuseUserID & "'"
RunSQLStatement sSql

' Save the reservation information
sSql = "UPDATE egov_rentalreservations SET organization = " & sOrganization
sSql = sSql & ", pointofcontact = " & sPointOfContact
sSql = sSql & ", numberattending = " & sNumberAttending
sSql = sSql & ", purpose = " & sPurpose
sSql = sSql & ", receiptnotes = " & sReceiptNotes
sSql = sSql & ", privatenotes = " & sPrivateNotes
sSql = sSql & ", isonhold = " & sHoldFlag
sSql = sSql & ", iscall = " & sCallFlag
sSql = sSql & " WHERE reservationid = " & iReservationId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Save the Arrival and departure times
If iReservationDateCount > CLng(0) Then
	For x = 1 To iReservationDateCount
		sSql = "UPDATE egov_rentalreservationdates SET actualstarttime = '" & request("reservationarrivaldate" & x) & " " & request("arrivalhour" & x) & ":" & request("arrivalminute" & x) & " " & request("arrivalampm" & x)
		sSql = sSql & "', actualendtime = '" & request("reservationdeparturedate" & x) & " " & request("departurehour" & x) & ":" & request("departureminute" & x) & " " & request("departureampm" & x)
		sSql = sSql & "' WHERE reservationdateid = " & CLng(request("reservationdateid" & x))
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	Next 
End If 


' Save the Rental Rates (deposits and alcohol)
If iReservationFeeCount > CLng(0) Then
	For x = 1 To iReservationFeeCount
		sSql = "UPDATE egov_rentalreservationfees SET feeamount = " & CDbl(request("reservationfeeamount" & x))
		sSql = sSql & " WHERE reservationfeeid = " & CLng(request("reservationfeeid" & x))
		'dTotalCharges = dTotalCharges + CDbl(request("reservationfeeamount" & x))
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	Next 
End If 

' Save the date fee amounts
If iRentalRateCount > CLng(0) Then 
	For x = 1 To iRentalRateCount
		sSql = "UPDATE egov_rentalreservationdatefees SET feeamount = " & CDbl(request("datefeeamount" & x))
		sSql = sSql & " WHERE reservationdatefeeid = " & CLng(request("reservationdatefeeid" & x))
		'dTotalCharges = dTotalCharges + CDbl(request("datefeeamount" & x))
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	Next 
End If 

' Save the item quantities and fee amounts
If iReservationItemCount > CLng(0) Then
	For x = 1 To iReservationItemCount
		sSql = "UPDATE egov_rentalreservationdateitems SET feeamount = " & CDbl(request("itemfeeamount" & x))
		sSql = sSql & ", quantity = " & clng(request("quantity" & x))
		sSql = sSql & " WHERE reservationdateitemid = " & CLng(request("reservationdateitemid" & x))
		'dTotalCharges = dTotalCharges + CDbl(request("itemfeeamount" & x))
		response.write sSql & "<br /><br />"
		RunSQLStatement sSql

		'response.write "sReservationTypeSelection = " & sReservationTypeSelection & "<br /><br />"

		If sReservationTypeSelection = "admin" Then
			' if it is internal, then change the Account Ledger Row for that item.
			response.write "quantity = " & clng(request("quantity" & x)) & "<br /><br />"
			If clng(request("quantity" & x)) = clng(0) Then
				' Remove the account ledger row since the quantity is now 0
				sSql = "DELETE FROM egov_accounts_ledger WHERE reservationfeetype = 'reservationdateitemid' "
				sSql = sSql & "AND reservationfeetypeid = " & CLng(request("reservationdateitemid" & x))
				'response.write sSql & "<br /><br />"
				RunSQLStatement sSql
			Else
				' update or add a row 
				If AccountLedgerFeeRowExists( CLng(request("reservationdateitemid" & x)), "reservationdateitemid" ) Then	
					' update it
					sSql = "UPDATE egov_accounts_ledger SET itemquantity = " & clng(request("quantity" & x))
					sSql = sSql & " WHERE reservationfeetype = 'reservationdateitemid' "
					sSql = sSql & " AND reservationfeetypeid = " & CLng(request("reservationdateitemid" & x))
					'response.write sSql & "<br /><br />"
					RunSQLStatement sSql
				Else
					GetReservationDateItemKeyValues CLng(request("reservationdateitemid" & x)), iReservationDateId, iPaymentId, iAccountId			' in rentalscommonfunctions.asp
					iItemTypeID = GetItemTypeId( "rentals" )	' in rentalscommonfunctions.asp
					' create it 
					sSql = "INSERT INTO egov_accounts_ledger ( orgid, paymentid, entrytype, accountid, itemquantity, amount, "
					sSql = sSql & "itemtypeid, plusminus, itemid, ispaymentaccount, paymenttypeid, isccrefund, reservationid, "
					sSql = sSql & "reservationdateid, reservationfeetype, reservationfeetypeid ) VALUES ( " & session("orgid") & ", "
					sSql = sSql & iPaymentId & ", 'credit', " & iAccountId & ", " & clng(request("quantity" & x)) & ", 0.00, "
					sSql = sSql & iItemTypeID & ", '+', " & iReservationId & ", 0, NULL, 0, " & iReservationId & ", " 
					sSql = sSql & iReservationDateId & ", 'reservationdateitemid', " & CLng(request("reservationdateitemid" & x)) & " )"
					'response.write sSql & "<br /><br />"
					RunSQLStatement sSql
				End If 
			End If 
		End If 
	Next 
End If 

' Get the fee total amount from the tables, this gives a more accurate value
dTotalCharges = CalculateReservationTotal( iReservationId, "feeamount" )	' in rentalscommonfunctions.asp
dTotalPayment = CalculateReservationTotal( iReservationId, "paidamount" )	' in rentalscommonfunctions.asp
dTotalRefunds = CalculateReservationTotal( iReservationId, "refundamount" )	' in rentalscommonfunctions.asp

' Update the total for the reservation
sSql = "UPDATE egov_rentalreservations SET totalamount = " & dTotalCharges
sSql = sSql & ", totalpaid = " & dTotalPayment
sSql = sSql & ", totalrefunded = " & dTotalRefunds
If request("servingalcohol") = "on" Then
	sSql = sSql & ", servingalcohol = 1"
Else
	sSql = sSql & ", servingalcohol = 0"
End If 
sSql = sSql & " WHERE reservationid = " & iReservationId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Take them back to the edit page for this reservation
response.redirect "reservationedit.asp?reservationid=" & iReservationId & "&sf=cs"


%>
