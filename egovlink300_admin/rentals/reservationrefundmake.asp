<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationrefundmake.asp
' AUTHOR: Steve Loar
' CREATED: 11/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Makes the refunds for Rental reservations.
'
' MODIFICATION HISTORY
' 1.0   11/20/2009	Steve Loar - INITIAL VERSION
' 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationId, x, iItemTypeID, iPaymentLocationId, sRefundNotes, iJournalEntryTypeID, iRentalUserid
Dim dRefundAmount, iAdminLocationId, iAccountId, iReservationDateId, iMaxRentalRates, iMaxReservationItems
Dim dNewRefundAmount, dOldRefundAmount, iMaxRefundFees, dRefundFeeTotal, iPaymentTypeId
Dim iIsCCRefund, dTotalRefunded, dGrossRefundAmount

iReservationId = CLng(request("reservationid"))

iItemTypeID = GetItemTypeId( "rentals" )

iPaymentLocationId = request("paymentlocationid")

'this is where the admin person is working today
If session("LocationId") <> "" Then 
	iAdminLocationId = session("LocationId")
Else 
	iAdminLocationId = 0 
End If 

' citizen userid
iRentalUserid = GetReservationRentalUserId( iReservationId )

dRefundAmount = CDbl(request("refundtotal")) ' Refund total from the refund page. It is already less fees and damages
'response.write "dRefundAmount: " & dRefundAmount & "<br />"

'dGrossRefundAmount = CDbl(request("grossrefundamount")) ' The total refund before taking out fees and damages
'response.write "dGrossRefundAmount: " & dGrossRefundAmount & "<br />"

dTotalRefunded = CDbl(0)
dRefundFeeTotal = CDbl(0.00) 

iJournalEntryTypeID = GetJournalEntryTypeID( "refund" )

sRefundNotes = dbsafe(request("refundnotes"))


' Make the journal entry
'Insert the egov_class_payment row (Journal entry) - this needs to be the total refund before fees
sSql = "INSERT INTO egov_class_payment (paymentdate, paymentlocationid, orgid, adminlocationid, "
sSql = sSql & " userid, adminuserid, paymenttotal, journalentrytypeid, notes, isforrentals, reservationid) VALUES (dbo.GetLocalDate(" & Session("orgid") & ",GetDate()), " 
sSql = sSql & iPaymentLocationId & ", " & Session("orgid") & ", " & iAdminLocationId & ", "
sSql = sSql & iRentalUserid & ", " & Session("UserID") & ", " & dRefundAmount & ", " & iJournalEntryTypeID & ", '" & sRefundNotes & "', 1, " & iReservationId & " )"
'response.write sSql & "<br /><br />"
iPaymentId = RunInsertStatement( sSql )


' make account ledger entries for the day fees being refunded
iMaxRentalRates = clng(request("maxrentalrates"))
If iMaxRentalRates > clng(0) Then 
	x = 1
	Do While x <= iMaxRentalRates
		dNewRefundAmount = CDbl(0)
		If CDbl(request("datefeeamount" & x)) > CDbl(0.00) Then 
			' Get the account for the rate
			iAccountId = GetReservationAccountId( request("reservationdatefeeid" & x), "reservationdatefeeid", "egov_rentalreservationdatefees" )	' In rentalscommonfunctions.asp
			iReservationDateId = GetReservationDateId( request("reservationdatefeeid" & x), "reservationdatefeeid", "egov_rentalreservationdatefees" )	' In rentalscommonfunctions.asp

			' Add to Accounts Ledger Row
			sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
			sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, reservationfeetypeid, reservationfeetype, reservationdateid ) VALUES ( "
			sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'debit', " & iAccountId & ", " & CDbl(request("datefeeamount" & x)) & ", " & iItemTypeId & ", '-', " 
			sSql = sSql & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", " & request("reservationdatefeeid" & x) & ", 'reservationdatefeeid', " & iReservationDateId & " )"
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql

			' Update the date fee with the new refund amount incase there was a partial refund on this already
			dOldRefundAmount = GetCurrentRefundAmount( request("reservationdatefeeid" & x), "reservationdatefeeid", "egov_rentalreservationdatefees" )	' In rentalscommonfunctions.asp
			'response.write "dOldRefundAmount = " & dOldRefundAmount & " ****<br /><br />"
			dNewRefundAmount = dOldRefundAmount + CDbl(request("datefeeamount" & x))
			'response.write "dNewRefundAmount = " & dNewRefundAmount & " ****<br /><br />"
			sSql = "UPDATE egov_rentalreservationdatefees SET refundamount = " & dNewRefundAmount & " WHERE reservationdatefeeid = " & request("reservationdatefeeid" & x)
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql
		End If 

		x = x + 1
	Loop
End If 

' make account ledger entries for the item fees being refunded
iMaxReservationItems = clng(request("maxreservationitems"))
If iMaxReservationItems > clng(0) Then 
	x = 1
	Do While x <= iMaxReservationItems
		dNewRefundAmount = CDbl(0)
		If CDbl(request("itemfeeamount" & x)) > CDbl(0.00) Then 
			' Get the account for the item
			iAccountId = GetReservationAccountId( request("reservationdateitemid" & x), "reservationdateitemid", "egov_rentalreservationdateitems" )	' In rentalscommonfunctions.asp
			iReservationDateId = GetReservationDateId( request("reservationdateitemid" & x), "reservationdateitemid", "egov_rentalreservationdateitems" )	' In rentalscommonfunctions.asp

			' Add to Accounts Ledger Row
			sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
			sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, reservationfeetypeid, reservationfeetype, reservationdateid ) VALUES ( "
			sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'debit', " & iAccountId & ", " & CDbl(request("itemfeeamount" & x)) & ", " & iItemTypeId & ", '-', " 
			sSql = sSql & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", " & request("reservationdateitemid" & x) & ", 'reservationdateitemid', " & iReservationDateId & " )"
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql

			' Update the date item with the new refund amount
			dOldRefundAmount = GetCurrentRefundAmount( request("reservationdateitemid" & x), "reservationdateitemid", "egov_rentalreservationdateitems" )	' In rentalscommonfunctions.asp
			dNewRefundAmount = dOldRefundAmount + CDbl(request("itemfeeamount" & x))
			sSql = "UPDATE egov_rentalreservationdateitems SET refundamount = " & dNewRefundAmount & " WHERE reservationdateitemid = " & request("reservationdateitemid" & x)
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql
		End If 

		x = x + 1
	Loop
End If 


' make account ledger entries for the reservation fees being refunded
iMaxReservationFees = clng(request("maxreservationfees"))
If iMaxReservationFees > clng(0) Then 
	x = 1
	Do While x <= iMaxReservationFees
		dNewRefundAmount = CDbl(0)
		If CDbl(request("reservationfeeamount" & x)) > CDbl(0.00) Then 
			' Get the account for the item
			iAccountId = GetReservationAccountId( request("reservationfeeid" & x), "reservationfeeid", "egov_rentalreservationfees" )	' In rentalscommonfunctions.asp

			' Add to Accounts Ledger Row
			sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
			sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, reservationfeetypeid, reservationfeetype ) VALUES ( "
			sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'debit', " & iAccountId & ", " & CDbl(request("reservationfeeamount" & x)) & ", " & iItemTypeId & ", '-', " 
			sSql = sSql & iReservationId & ", 0, NULL, NULL, " & iReservationId & ", " & request("reservationfeeid" & x) & ", 'reservationfeeid' )"
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql

			' Update the date item with the new refund amount
			dOldRefundAmount = GetCurrentRefundAmount( request("reservationfeeid" & x), "reservationfeeid", "egov_rentalreservationfees" )	' In rentalscommonfunctions.asp
			'response.write "dOldRefundAmount = " & dOldRefundAmount & " ****<br /><br />"
			dNewRefundAmount = dOldRefundAmount + CDbl(request("reservationfeeamount" & x))
			'response.write "dNewRefundAmount = " & dNewRefundAmount & " ****<br /><br />"
			sSql = "UPDATE egov_rentalreservationfees SET refundamount = " & dNewRefundAmount & " WHERE reservationfeeid = " & request("reservationfeeid" & x)
			'response.write "***** " & sSql & " ****<br /><br />"
			RunSQLStatement sSql
		End If 

		x = x + 1
	Loop
End If 


' Handle the refund fee entries into the account ledger
iMaxRefundFees = clng(request("maxrefundfees"))
If iMaxRefundFees > clng(0) Then 
	x = 1
	dRefundFeeTotal = CDbl(0)
	Do While x <= iMaxRefundFees
		If CDbl(request("refundfeeamount" & x)) > CDbl(0.00) Then 
			' Get the account for the item
			iAccountId = GetPaymentAccountId( Session("OrgId"), request("paymenttypeid" & x) )  ' In common.asp

			dRefundFeeTotal = dRefundFeeTotal + CDbl(request("refundfeeamount" & x))

			' Add to Accounts Ledger Row
			sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
			sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid ) VALUES ( "
			sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', " & iAccountId & ", " & CDbl(request("refundfeeamount" & x)) & ", " & iItemTypeId & ", '+', " 
			sSql = sSql & iReservationId & ", 1, " & request("paymenttypeid" & x) & ", NULL, " & iReservationId & " )"
			'response.write sSql & "<br /><br />"
			RunSQLStatement sSql
		End If 

		x = x + 1
	Loop
End If 


' Calculate any left over to give back
'dRefundAmount = dRefundAmount - dRefundFeeTotal
If CDbl(dRefundAmount) > CDbl(0.00) Then 
	' Handle any extra to give back via a refund voucher or a CC refund
	If CLng(request("accountid")) = CLng(0) Then
		' if Refund voucher was selected so create a ledger row for that
		iPaymentTypeId = GetRefundPaymentTypeId( )  ' In common.asp
		iAccountId = GetPaymentAccountId( Session("OrgId"), iPaymentTypeId )  ' In common.asp
		
		If request("isccrefund") <> "" Then
			iIsCCRefund = request("isccrefund")
		Else
			iIsCCRefund = 0
		End If 

		' Add to Accounts Ledger Row
		sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
		sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid, isccrefund ) VALUES ( "
		sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', " & iAccountId & ", " & dRefundAmount & ", " & iItemTypeId & ", '-', " 
		sSql = sSql & iReservationId & ", 1, " & iPaymentTypeId & ", NULL, " & iReservationId & ", " & iIsCCRefund & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	Else 
		' Else Create Ledger entry for crediting (-) user account 
		iAccountId = CLng(request("accountid"))
		cPriorBalance = GetCitizenCurrentBalance( iAccountId )		' In Common.asp
		response.write "cPriorBalance: " & cPriorBalance & "<br />"
		response.write "dRefundAmount: " & dRefundAmount & "<br />"

		' Credit the Citizen account that is to get the refund
		AdjustCitizenAccountBalance iAccountId, "credit", dRefundAmount

		' Add to Accounts Ledger Row
		sSql = "INSERT Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
		sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, reservationid ) VALUES ( "
		sSql = sSql & iPaymentId & ", " & session("orgid") & ", 'credit', " & iAccountId & ", " & dRefundAmount & ", " & iItemTypeId & ", '-', " 
		sSql = sSql & iReservationId & ", 0, 4, " & cPriorBalance & ", " & iReservationId & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	End If 
End If 


' Update the registration with the refunded total
dTotalRefunded = CalculateReservationTotal( iReservationId, "refundamount" )
sSql = "UPDATE egov_rentalreservations SET totalrefunded = " & dTotalRefunded & " WHERE reservationid = " & iReservationId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql


' Go to the receipt page
response.redirect "viewpaymentreceipt.asp?paymentid=" & iPaymentId & "&rt=r"


%>
