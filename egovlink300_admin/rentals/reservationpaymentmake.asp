<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationpaymentmake.asp
' AUTHOR: Steve Loar
' CREATED: 11/12/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Makes the payments for Rental reservations.
'
' MODIFICATION HISTORY
' 1.0   11/12/2009	Steve Loar - INITIAL VERSION
' 
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
dim iReservationId, sSql, oRs, iPaymentLocationId, iAdminLocationId, iRentalUserid, dPaymentTotal
dim iJournalEntryTypeID, sPurchaseNotes, iPaymentId, iMaxPayments, x, iMaxRentalRates, dOldPaidAmount
dim iItemTypeID, dNewPaidAmount, iMaxReservationItems, iMaxReservationFees, iReservationDateId
dim bOkToProcess, bIncludeZeroFees

iReservationId      = CLng(request("reservationid"))
iItemTypeID         = GetItemTypeId( "rentals" )
iPaymentLocationId  = request("paymentlocationid") 
bIncludeZeroFees    = false 
iAdminLocationId    = 0 
iRentalUserid       = GetReservationRentalUserId( iReservationId )  'citizen userid
dPaymentTotal       = CDbl(request("paymenttotal"))  'Payment total
iJournalEntryTypeID = GetJournalEntryTypeID( "rentalpayment" )
sPurchaseNotes      = dbsafe(request("purchasenotes"))

if request("includezerocharges") = "on" then
  bIncludeZeroFees = true
end if

'this is where the admin person is working today
if session("LocationId") <> "" then
 	iAdminLocationId = session("LocationId")
end if

'Insert the egov_class_payment row (Journal entry)
sSql = "INSERT INTO egov_class_payment ("
sSql = sSql & "paymentdate, "
sSql = sSql & "paymentlocationid, "
sSql = sSql & "orgid, "
sSql = sSql & "adminlocationid, "
sSql = sSql & "userid, "
sSql = sSql & "adminuserid, "
sSql = sSql & "paymenttotal, "
sSql = sSql & "journalentrytypeid, "
sSql = sSql & "notes, "
sSql = sSql & "isforrentals, "
sSql = sSql & "reservationid"
sSql = sSql & ") VALUES ("
sSql = sSql & "dbo.GetLocalDate(" & session("orgid") & ",GetDate()), " 
sSql = sSql & iPaymentLocationId   & ", "
sSql = sSql & session("orgid")     & ", "
sSql = sSql & iAdminLocationId     & ", "
sSql = sSql & iRentalUserid        & ", "
sSql = sSql & session("UserID")    & ", "
sSql = sSql & dPaymentTotal        & ", "
sSql = sSql & iJournalEntryTypeID  & ", "
sSql = sSql & "'" & sPurchaseNotes & "', "
sSql = sSql & "1, "
sSql = sSql & iReservationId
sSql = sSql & ")"
'response.write sSql & "<br /><br />"

iPaymentId = RunInsertStatement( sSql )

'Insert the payment rows
iMaxPayments = clng(request("maxpayments"))

if iMaxPayments > clng(0) then
   	x = 1

   	do while x <= iMaxPayments
     		if trim(request("paymentamount" & x)) <> "" then
       			if CDbl(request("paymentamount" & x)) > CDbl(0.00) then
         				bOkToProcess = true
       			else
         				if bIncludeZeroFees then
           					bOkToProcess = true  
         				else
           					bOkToProcess = false 
          			end if
       			end if

       			if bOkToProcess then 
         				if request("hascheckno" & x) = "yes" then
           				'Check payment
           					sCheck            = "'" & dbsafe(request("checkno" & x)) & "'"
      										iCitizenAccountId = "NULL"
      										sPlusMinus        = "+"
      										cPriorBalance     = "NULL"
      										iAccountId        = GetPaymentAccountId( Session("Orgid"), request("paymenttypeid" & x) )		' In common.asp
      							else
      										sCheck = "NULL"

      										if HasCitizensAccounts( request("paymenttypeid" & x) ) then
      											 'Paying with the citizen balance on account
      					  						iCitizenAccountId = request("accountid")	' this is the pick of family member accounts to get the payment from
					  												iAccountId        = iCitizenAccountId
					  												sPlusMinus        = "-"
					  												cPriorBalance     = GetCitizenCurrentBalance( iCitizenAccountId )		' In common.asp

				  												'Debit the account that was the source of the funds
					  												AdjustCitizenAccountBalance iCitizenAccountId, "debit", request("paymentamount" & x)
					  									else
				  												'Charge, Cash and Other
					  												iCitizenAccountId = "NULL"
					  												sPlusMinus        = "+"
					  												cPriorBalance     = "NULL"
					  												iAccountId        = GetPaymentAccountId( Session("Orgid"), request("paymenttypeid" & x) )		' In common.asp
					  									end if
					  						end if

					  					  'Make the ledger entry for the payment
					  						'iLedgerId = MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid )
					  						'iLedgerId = MakeLedgerEntry( Session("Orgid"), iAccountId, iPaymentId, CDbl(request("paymentamount" & x)), "NULL", "debit", sPlusMinus, "NULL", 1, x, cPriorBalance, "NULL" )
					  						sSql = "INSERT INTO egov_accounts_ledger ("
                        sSql = sSql & "paymentid, "
                        sSql = sSql & "orgid, "
                        sSql = sSql & "entrytype, "
                        sSql = sSql & "accountid, "
                        sSql = sSql & "amount, "
                        sSql = sSql & "itemtypeid, "
                        sSql = sSql & "plusminus, "
                        sSql = sSql & "itemid, "
                        sSql = sSql & "ispaymentaccount, "
                        sSql = sSql & "paymenttypeid, "
                        sSql = sSql & "priorbalance, "
                        sSql = sSql & "pricetypeid, "
                        sSql = sSql & "reservationid"
                        sSql = sSql & ") VALUES ("
                        sSql = sSql & iPaymentId                         & ", "
                        sSql = sSql & Session("Orgid")                   & ", "
                        sSql = sSql & "'debit', "
                        sSql = sSql & iAccountId                         & ", "
                        sSql = sSql & CDbl(request("paymentamount" & x)) & ", "
                        sSql = sSql & "NULL, '"
                        sSql = sSql & sPlusMinus                         & "', "
                        sSql = sSql & "NULL, "
                        sSql = sSql & "1, "
                        sSql = sSql & request("paymenttypeid" & x)       & ", "
                        sSql = sSql & cPriorBalance                      & ", "
                        sSql = sSql & "NULL, "
                        sSql = sSql & iReservationId
                        sSql = sSql & ")"

                        'response.write sSql & "<br /><br />"

                        iLedgerId = RunInsertStatement( sSql )

                        'Make the entry in the egov_verisign_payment_information table
                        'InsertPaymentInformation iPaymentId, iLedgerId, x, CDbl(request("amount" & x)), "APPROVED", sCheck, iCitizenAccountId
                        'InsertPaymentInformation iPaymentId, iLedgerId, iPaymentTypeId, sAmount, sStatus, sCheckNo, iAccountId
                        sSql = "INSERT INTO egov_verisign_payment_information ("
                        sSql = sSql & "paymentid, "
                        sSql = sSql & "ledgerid, "
                        sSql = sSql & "paymenttypeid, "
                        sSql = sSql & "amount, "
                        sSql = sSql & "paymentstatus, "
                        sSql = sSql & "checkno, "
                        sSql = sSql & "citizenuserid"
                        sSql = sSql & ") VALUES ("
                        sSql = sSql & iPaymentId                         & ", "
                        sSql = sSql & iLedgerId                          & ", " 
                        sSql = sSql & request("paymenttypeid" & x)       & ", "
                        sSql = sSql & CDbl(request("paymentamount" & x)) & ", "
                        sSql = sSql & "'APPROVED', "
                        sSql = sSql & sCheck                             & ", "
                        sSql = sSql & iCitizenAccountId
                        sSql = sSql & ")"

                        'response.write sSql & "<br /><br />"

                        RunSQLStatement sSql
						  		end if
       end if

     		x = x + 1
   	loop
end if

'BEGIN: Loop through the DAY FEE ROWS and record them -------------------------
iMaxRentalRates = clng(request("maxrentalrates"))

if iMaxRentalRates > clng(0) then
   	x = 1
   	do while x <= iMaxRentalRates
     		iReservationDateId = GetReservationDateId(request("reservationdatefeeid" & x), "reservationdatefeeid", "egov_rentalreservationdatefees")	 'In rentalscommonfunctions.asp

     		if CDbl(request("datefeeamount" & x)) > CDbl(0.00) then
       			bOkToProcess = True 
     		else
       			if bIncludeZeroFees then
         				bOkToProcess = true
          else
         				bOkToProcess = false
          end if
       end if

     		if bOkToProcess then
      			'BEGIN: Get the account for the rate ---------------------------------
        		iAccountId = GetReservationAccountId(request("reservationdatefeeid" & x), "reservationdatefeeid", "egov_rentalreservationdatefees")	 'In rentalscommonfunctions.asp
      			'END: Get the account for the rate -----------------------------------

      			'BEGIN: Add to Accounts Ledger Row -----------------------------------
            sSql = "INSERT INTO egov_accounts_ledger ("
            sSql = sSql & "paymentid, "
            sSql = sSql & "orgid, "
            sSql = sSql & "entrytype, "
            sSql = sSql & "accountid, "
            sSql = sSql & "amount, "
            sSql = sSql & "itemtypeid, "
            sSql = sSql & "plusminus, "
            sSql = sSql & "itemid, "
            sSql = sSql & "ispaymentaccount, "
            sSql = sSql & "paymenttypeid, "
            sSql = sSql & "priorbalance, "
            sSql = sSql & "reservationid, "
            sSql = sSql & "reservationfeetypeid, "
            sSql = sSql & "reservationfeetype, "
            sSql = sSql & "reservationdateid"
            sSql = sSql & ") VALUES ("
            sSql = sSql & iPaymentId & ", "
            sSql = sSql & session("orgid") & ", "
            sSql = sSql & "'credit', "
            sSql = sSql & iAccountId & ", "
            sSql = sSql & CDbl(request("datefeeamount" & x)) & ", "
            sSql = sSql & iItemTypeId & ", "
            sSql = sSql & "'+', "
            sSql = sSql & iReservationId & ", "
            sSql = sSql & "0, "
            sSql = sSql & "NULL, "
            sSql = sSql & "NULL, "
            sSql = sSql & iReservationId & ", "
            sSql = sSql & request("reservationdatefeeid" & x) & ", "
            sSql = sSql & "'reservationdatefeeid', "
            sSql = sSql & iReservationDateId
            sSql = sSql & ")"

            'response.write sSql & "<br /><br />"

            RunSQLStatement sSql
            'END: Add to Accounts Ledger Row -------------------------------------

					'BEGIN: Update the date fee with the new paid amount -----------------
						dOldPaidAmount = GetCurrentPaidAmount(request("reservationdatefeeid" & x), "reservationdatefeeid", "egov_rentalreservationdatefees")	 'In rentalscommonfunctions.asp

						dNewPaidAmount = dOldPaidAmount + CDbl(request("datefeeamount" & x))

						sSql = "UPDATE egov_rentalreservationdatefees SET "
            sSql = sSql & " paidamount = " & dNewPaidAmount
            sSql = sSql & " WHERE reservationdatefeeid = " & request("reservationdatefeeid" & x)

						'response.write sSql & "<br /><br />"

						RunSQLStatement sSql
					'END: Update the date fee with the new paid amount -------------------
 			 		end if

     		x = x + 1
   	loop
end if
'END: Loop through the DAY FEE ROWS and record them ---------------------------

'BEGIN: Loop through any DAY ITEMS and record them ----------------------------
iMaxReservationItems = clng(request("maxreservationitems"))

if iMaxReservationItems > clng(0) then
 	x = 1

   	do while x <= iMaxReservationItems
     		iReservationDateId = GetReservationDateId(request("reservationdateitemid" & x), "reservationdateitemid", "egov_rentalreservationdateitems")	 'In rentalscommonfunctions.asp

     		if CDbl(request("itemfeeamount" & x)) > CDbl(0.00) then
       			bOkToProcess = true
     		else
       			if bIncludeZeroFees then
         				bOkToProcess = true
          else
         				bOkToProcess = false
       			end if
       end if

     		if bOkToProcess then
      			'BEGIN: Get the account for the item ---------------------------------
        		iAccountId = GetReservationAccountId(request("reservationdateitemid" & x), "reservationdateitemid", "egov_rentalreservationdateitems")	 'In rentalscommonfunctions.asp
      			'END: Get the account for the item -----------------------------------

      			'BEGIN: Add to Accounts Ledger Row -----------------------------------
            sSql = "INSERT INTO egov_accounts_ledger ("
            sSql = sSql & "paymentid, "
            sSql = sSql & "orgid, "
            sSql = sSql & "entrytype, "
            sSql = sSql & "accountid, "
            sSql = sSql & "itemquantity, "
            sSql = sSql & "amount, "
            sSql = sSql & "itemtypeid, "
            sSql = sSql & "plusminus, "
            sSql = sSql & "itemid, "
            sSql = sSql & "ispaymentaccount, "
            sSql = sSql & "paymenttypeid, "
            sSql = sSql & "priorbalance, "
            sSql = sSql & "reservationid, "
            sSql = sSql & "reservationfeetypeid, "
            sSql = sSql & "reservationfeetype, "
            sSql = sSql & "reservationdateid"
            sSql = sSql & ") VALUES ("
            sSql = sSql & iPaymentId                           & ", "
            sSql = sSql & session("orgid")                     & ", "
            sSql = sSql & "'credit', "
            sSql = sSql & iAccountId                           & ", "
            sSql = sSql & CDbl(request("itemquantity" & x))    & ", "
            sSql = sSql & CDbl(request("itemfeeamount" & x))   & ", "
            sSql = sSql & iItemTypeId                          & ", "
            sSql = sSql & "'+', "
            sSql = sSql & iReservationId                       & ", "
            sSql = sSql & "0, "
            sSql = sSql & "NULL, "
            sSql = sSql & "NULL, "
            sSql = sSql & iReservationId                       & ", "
            sSql = sSql & request("reservationdateitemid" & x) & ", "
            sSql = sSql & "'reservationdateitemid', "
            sSql = sSql & iReservationDateId
            sSql = sSql & ")"

            'response.write sSql & "<br /><br />"

    				RunSQLStatement sSql
      			'END: Add to Accounts Ledger Row -------------------------------------

            'BEGIN: Update the date item with the new paid amount ----------------
            dOldPaidAmount = GetCurrentPaidAmount(request("reservationdateitemid" & x), "reservationdateitemid", "egov_rentalreservationdateitems")	 'In rentalscommonfunctions.asp

            dNewPaidAmount = dOldPaidAmount + CDbl(request("itemfeeamount" & x))

            sSql = "UPDATE egov_rentalreservationdateitems SET "
            sSql = sSql & " paidamount = " & dNewPaidAmount
            sSql = sSql & " WHERE reservationdateitemid = " & request("reservationdateitemid" & x)

            'response.write sSql & "<br /><br />"

            RunSQLStatement sSql
            'END: Update the date item with the new paid amount ------------------
 			 		end if

     		x = x + 1
   	loop
end if
'END: Loop through any DAY ITEMS and record them ------------------------------

'BEGIN: Loop through the RESERVATION FEES and record them ---------------------
iMaxReservationFees = clng(request("maxreservationfees"))

if iMaxReservationFees > clng(0) then
   	x = 1

   	do while x <= iMaxReservationFees
     		if CDbl(request("reservationfeeamount" & x)) > CDbl(0.00) then
      			'BEGIN: Get the account for the item ---------------------------------
       			iAccountId = GetReservationAccountId(request("reservationfeeid" & x), "reservationfeeid", "egov_rentalreservationfees") 	'In rentalscommonfunctions.asp
      			'END: Get the account for the item -----------------------------------

      			'BEGIN: Add to Accounts Ledger Row -----------------------------------
            sSql = "INSERT INTO egov_accounts_ledger ("
            sSql = sSql & "paymentid, "
            sSql = sSql & "orgid, "
            sSql = sSql & "entrytype, "
            sSql = sSql & "accountid, "
            sSql = sSql & "amount, "
            sSql = sSql & "itemtypeid, "
            sSql = sSql & "plusminus, "
            sSql = sSql & "itemid, "
            sSql = sSql & "ispaymentaccount, "
            sSql = sSql & "paymenttypeid, "
            sSql = sSql & "priorbalance, "
            sSql = sSql & "reservationid, "
            sSql = sSql & "reservationfeetypeid, "
            sSql = sSql & "reservationfeetype"
            sSql = sSql & ") VALUES ("
            sSql = sSql & iPaymentId                                & ", "
            sSql = sSql & session("orgid")                          & ", "
            sSql = sSql & "'credit', "
            sSql = sSql & iAccountId                                & ", "
            sSql = sSql & CDbl(request("reservationfeeamount" & x)) & ", "
            sSql = sSql & iItemTypeId                               & ", "
            sSql = sSql & "'+', " 
            sSql = sSql & iReservationId                            & ", "
            sSql = sSql & "0, "
            sSql = sSql & "NULL, "
            sSql = sSql & "NULL, "
            sSql = sSql & iReservationId                            & ", "
            sSql = sSql & request("reservationfeeid" & x)           & ", "
            sSql = sSql & "'reservationfeeid'"
            sSql = sSql & " )"

            'response.write sSql & "<br /><br />"

            RunSQLStatement sSql
            'END: Add to Accounts Ledger Row -------------------------------------

            'BEGIN: Update the date item with the new paid amount ----------------
       			dOldPaidAmount = GetCurrentPaidAmount(request("reservationfeeid" & x), "reservationfeeid", "egov_rentalreservationfees")	 'In rentalscommonfunctions.asp

    				dNewPaidAmount = dOldPaidAmount + CDbl(request("reservationfeeamount" & x))

    				sSql = "UPDATE egov_rentalreservationfees SET "
            sSql = sSql & " paidamount = " & dNewPaidAmount
            sSql = sSql & " WHERE reservationfeeid = " & request("reservationfeeid" & x)

    				'response.write sSql & "<br /><br />"

    				RunSQLStatement sSql
			      'END: Update the date item with the new paid amount ------------------
       end if

     		x = x + 1
    loop
end if
'END: Loop through the RESERVATION FEES and record them ---------------------

'When done calculate a new paid total for the reservation, then save that to the table
dTotalPaid = CalculateReservationTotal(iReservationId, "paidamount" )

sSql = "UPDATE egov_rentalreservations SET "
sSql = sSql & " totalpaid = " & dTotalPaid
sSql = sSql & " WHERE reservationid = " & iReservationId
response.write sSql & "<br /><br />"

RunSQLStatement sSql

' see if the org has the undo feature and set the session variable'
If OrgHasFeature("undo on rental receipt") Then
  ' In ../includes/common.asp'
  SetUnDoBtnDisplay iPaymentId, True  
End If 

'Go to the receipt page
response.redirect "viewpaymentreceipt.asp?paymentid=" & iPaymentId & "&rt=r"


%>
