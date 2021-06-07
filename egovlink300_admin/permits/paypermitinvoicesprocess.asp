<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: paypermitinvoicesprocess.asp
' AUTHOR: Steve Loar
' CREATED: 06/03/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Process invoice payments for one permit
'
' MODIFICATION HISTORY
' 1.0   06/03/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitContactId, dTotalDue, dPaymentTotal, sSql, dCurrentPaid, dTotalPaid
Dim iPaymentLocationId, sPermitNo, iJournalEntryTypeID, iMaxPayments, x, sCheck, cAmount
Dim iAccountId, iPaymentTypeId, iJournalId, iMaxInvoices, iItemTypeId, iInvoiceStatusId

iPermitId = CLng(request("permitid"))

sPermitNo = GetPermitNumber( iPermitId )  ' in permitcommonfunctions.asp

iPermitContactId = CLng(request("permitcontactid"))

dTotalDue = CDbl(request("totaldue"))

dPaymentTotal = CDbl(request("paymenttotal"))

iPaymentLocationId = CLng(request("paymentlocationid"))

iMaxPayments = CLng(request("maxpayments"))

iMaxInvoices = CLng(request("maxinvoices"))

iInvoiceStatusId = GetInvoiceStatusId( "ispaid" )

x = 1


' Update the total paid for the permit
dCurrentPaid = GetPermitTotalPaid( iPermitId )
dTotalPaid = dCurrentPaid + dPaymentTotal

' Update the Total Paid column on the permit
sSql = "UPDATE egov_permits SET totalpaid = " & dTotalPaid & " WHERE permitid = " & iPermitId
'response.write sSQL & "<br /><br />"
RunSQL sSql

' Create the Journal Entry in egov_class_payment
iJournalEntryTypeID = GetJournalEntryTypeID( "permitpayment" )
iJournalId = MakeJournalEntry( iPaymentLocationId, "NULL", "NULL", session("userid"), dPaymentTotal, iJournalEntryTypeID, "Admin payment towards permitid " & iPermitId & ", permit# " & sPermitNo )

' Create the Account Ledger rows for the payment types
Do While x <= iMaxPayments
	'response.write "Amount field = " & request("amount" & x)
	If request("amount" & x) <> "" and isnumeric(request("amount" & x)) Then 
		iPaymentTypeId = CLng(request("paymenttypeid" & x))
		iAccountId = GetPaymentAccountId( session("orgid"), iPaymentTypeId )
		cAmount = CDbl(request("amount" & x))
		If request("checkno" & x) <> "" Then 
			sCheck = "'" & dbsafe(request("checkno" & x)) & "'"
		Else
			sCheck = "NULL"
		End If 

		'           MakeLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iPriceTypeid, iPermitId, iInvoiceId, iPermitFeeId )
		iLedgerId = MakeLedgerEntry( session("orgid"), iAccountId, iJournalId, cAmount, "NULL", "debit", "+", "NULL", 1, iPaymentTypeId, "NULL", "NULL", iPermitId, "NULL", "NULL" )

		' Make the entry in the egov_verisign_payment_information table - This is in ../includes/common.asp
		InsertPaymentInformation iJournalId, iLedgerId, x, cAmount, "APPROVED", sCheck, "NULL"
	End If 
	x = x + 1
Loop 

iItemTypeId = GetItemTypeId( "permit" )

x = 1
' For each invoice update the paymentid from the Journal Entry
Do While x <= iMaxInvoices
	If request("includeinvoice" & x) = "on" Then
		' For each fee of each invoice update the payment id and amount paid (full amount)
		CreateInvoiceLedgerEntries iPermitId, request("invoiceid" & x), iJournalId, iItemTypeId
		' Set the back link so the invoice knows how it was paid
		sSql = "UPDATE egov_permitinvoices SET paymentid = " & iJournalId & ", invoicestatusid = " & iInvoiceStatusId & " WHERE invoiceid = " & request("invoiceid" & x)
		RunSQL sSql
	End If 
	x = x + 1
Loop

' Push out the expiration date
PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp

' Go to the permit invoice summary page for this permit and contact
'response.redirect "viewinvoicesummary.asp?permitid=" & iPermitId & "&permitcontactid=" & iPermitContactId
response.write "<script>parent.RefreshPageAfterVoid( 'PaidInvoice|" & iPermitContactId & "' );</script>"


'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub CreateInvoiceLedgerEntries( iPermitId, iInvoiceId, iJournalId, iItemTypeId )
'--------------------------------------------------------------------------------------------------
Sub CreateInvoiceLedgerEntries( iPermitId, iInvoiceId, iJournalId, iItemTypeId )
	Dim oRs, sSql, iAccountId, iLedgerId

	' Pull the fees for the invoice
	sSql = "SELECT permitfeeid, invoicedamount FROM egov_permitinvoiceitems WHERE invoiceid = " &  iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		' get account for each fee
		iAccountId = GetFeeAccountId( oRs("permitfeeid") )
		
		' Make ledger entry for each fee
		iLedgerId = MakeLedgerEntry( session("orgid"), iAccountId, iJournalId, CDbl(oRs("invoicedamount")), iItemTypeId, "credit", "+", "NULL", 0, "NULL", "NULL", "NULL", iPermitId, iInvoiceId, oRs("permitfeeid") )
		
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetFeeAccountId( iPermitFeeId )
'--------------------------------------------------------------------------------------------------
Function GetFeeAccountId( iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountid,0) AS accountid FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("accountid")) > CLng(0) Then 
			GetFeeAccountId = CLng(oRs("accountid"))
		Else
			GetFeeAccountId = "NULL"
		End If 
	Else 
		GetFeeAccountId = "NULL"
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetItemTypeId( sType )
'--------------------------------------------------------------------------------------------------
Function GetItemTypeId( sType )
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetItemTypeId = CLng(oRs("itemtypeid"))
	Else 
		GetItemTypeId = CLng(0)
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function ShowPaymentChoices()
'--------------------------------------------------------------------------------------------------
Function GetPermitTotalPaid( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(totalpaid, 0.00) AS totalpaid FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTotalPaid = CDbl(oRs("totalpaid"))
	Else
		GetPermitTotalPaid = CDbl(0.00)
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function



%>
