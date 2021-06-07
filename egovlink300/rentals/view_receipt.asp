<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: view_receipt.asp
' AUTHOR: Steve Loar
' CREATED: 01/13/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  rentals receipt page
'
' MODIFICATION HISTORY
' 1.0   01/13/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle, iLoggedUserid, iPaymentId, iReservationId, lcl_orghasfeature_citizen_accounts
Dim iUserId, nTotal, dPaymentDate, nRowTotal, bMultiWeeks, iAdminLocationId, iJournalEntryTypeId
Dim sJournalEntryType, bIsCCRefund, sNotes, iAdminUserId, bHasPaymentData, bHasPaymentFee
Dim dProcessingFee

iPaymentId = CLng(request("iPaymentId"))

If request.servervariables("HTTPS") <> "on" Then 
	'If they do not have a userid set, take them to the login page automatically
	If request.cookies("userid") = "" Or request.cookies("userid") = "-1" Then 
		session("RedirectPage") = "rentals/view_receipt.asp?ipaymentid=" & iPaymentId
		session("RedirectLang") = "Return to " & GetOrgDisplay( iOrgId, "rentalscategorypagetop" )
		session("ManageURL")    = ""
		session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
		response.redirect "../user_login.asp"
	End If 

	iLoggedUserid = request.cookies("userid")
Else 
	iLoggedUserid = request("userid")
End If 

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If


iReservationId = GetReservationIdFromPaymentId( iPaymentId )	' in rentalscommonfunctions.asp
'if iOrgID = "228" then response.write iReservationID

'GetGeneralReservationData iReservationId

'Check for org features
lcl_orghasfeature_citizen_accounts = OrgHasFeature( iOrgId, "citizen accounts" )

bHasPaymentData = GetPaymentDetails( iPaymentId, iUserId, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminUserId, iLoggedUserid )


%>

<html>
<head>

	<title><%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="rentalprintstyles.css" media="print" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="Javascript">
	<!--


	//-->
	</script>

<%
	If request.servervariables("HTTPS") = "on" Then 
		response.write "<style>" & vbcrlf
		response.write "  body {behavior: url('https://secure.egovlink.com/" & sorgVirtualSiteName & "/csshover.htc');}" & vbcrlf
		response.write "</style>" & vbcrlf
	End If 
%>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<br /><br />

<div id="content">
	<div id="centercontent">
		<div id="topleftbuttons">
		  <input type="button" class="button" onclick="javascript:window.print();" value="Print" />
		</div>
<%	
	If bHasPaymentData Then 

		' Display their header
		ShowReceiptHeader iPaymentId
		
		response.write "<hr />"
		response.write "<span id=""receiptadmininfo"">"
		response.write " Location: " & GetAdminLocation( iAdminLocationId ) & "&nbsp;&nbsp;"   'In ../includes/common.asp
		response.write " Administrator: " & GetAdminName( iAdminuserid )   'In ../includes/common.asp
		response.write "</span>"
		response.write vbcrlf & "Date: " & DateValue(CDate(dPaymentDate)) & "&nbsp;&nbsp;"
		response.write " Receipt #: " & iPaymentId & "&nbsp;&nbsp;"
		response.write " Reservation Id: " & iReservationId & "&nbsp;&nbsp;"
		'response.write " Location: " & GetAdminLocation( iAdminLocationId )  & " &nbsp; &nbsp; "
		'response.write " Administrator: " & GetAdminName( iAdminUserId )
		response.write "<hr />"

		' this will float to the right
		response.write vbcrlf & "<div id=""receipttopright"">"
		response.write vbcrlf & "<p id=""transactiontotal""><strong>"
		sJournalEntryType = GetJournalEntryType( iJournalEntryTypeId )

		If sJournalEntryType = "rentalpayment" Then
			' See if the gateway for this org has fees they charge the citizen
			If PaymentGatewayRequiresFeeCheck( iOrgId ) Then
				bHasPaymentFee = True 
				dProcessingFee = GetProcessingFee( iPaymentId )
			Else
				bHasPaymentFee = False 
				dProcessingFee = CDbl("0.00")
			End If 
			response.write "Transaction Total:"
		Else
			bHasPaymentFee = False 
			dProcessingFee = CDbl("0.00")
			response.write "Amount Refunded:"
		End If 

		response.write "</strong> " & FormatCurrency((CDbl(nTotal) + CDbl(dProcessingFee)),2) & "</p>"

		' this will be to the left of the last div.
		If lcl_orghasfeature_citizen_accounts Then 
  			ShowAccountChange iPaymentId, iUserId
		End If 
		response.write vbcrlf & "</div>"

		response.write ShowUserInfo( iUserId, sJournalEntryType )
		response.write "<hr />"

		If sJournalEntryType = "refund" Then 
			ShowRefundType iPaymentId
		Else
			ShowPaymentTypes iPaymentId, sJournalEntryType, bHasPaymentFee, dProcessingFee
		End If 
		response.write "<hr />"

		response.write "<strong>Transactions</strong>"
		response.write "<hr />"
		'response.write "sJournalEntryType = " & sJournalEntryType

		Select Case sJournalEntryType
			Case "rentalpayment"
				' Show purchase details
				ShowReservationDetails iReservationId, iPaymentId, "credit", sJournalEntryType, bHasPaymentFee, dProcessingFee
			Case "refund"
				' Show refund stuff
				ShowReservationDetails iReservationId, iPaymentId, "debit", sJournalEntryType, bHasPaymentFee, dProcessingFee
		End Select 

		response.write vbcrlf & "<hr />" 

		'response.write vbcrlf & "<p><strong>Receipt Notes:</strong> " & Trim(sNotes) & "</p>"

		ShowReservationReceiptNotes iReservationId, sJournalEntryType

		ShowReceiptFooter sNotes, sJournalEntryType

		response.write "<br /><br />"

	Else 
		response.write "<p id=""nopermissionreceipt"">No Details could be found for the requested receipt or you do not have permission to view this receipt.</p>"
	End If 

	If bHasPaymentData And sJournalEntryType = "rentalpayment" Then
		If ReservationHasAssociatedDocuments( iReservationId ) Then
			' Do a page break and list out the documents
			response.write vbcrlf & "<div id=""receiptdocumentstart""></div>"
			
			response.write vbcrlf & "<p id=""receiptdocumentstitle"">Documents Related To This Reservation</p>"
			
			ShowReservationDocuments iReservationId
		End If 
	End If 
%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' string GetJournalEntryType( iJournalEntryTypeId )
'------------------------------------------------------------------------------
Function GetJournalEntryType( ByVal iJournalEntryTypeId )
	Dim sSql, oRs

	sSql = "SELECT journalentrytype FROM egov_journal_entry_types WHERE journalentrytypeid = " & iJournalEntryTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetJournalEntryType = oRs("journalentrytype")
	Else
		GetJournalEntryType = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing
End Function 


'------------------------------------------------------------------------------
' double GetLedgerAmount( iPaymentId, iPaymentTypeId )
'------------------------------------------------------------------------------
Function GetLedgerAmount( ByVal iPaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs, cAmount

	sSql = "SELECT amount FROM egov_accounts_ledger "
	sSql = sSql & "WHERE ispaymentaccount = 1 AND paymentid = " & iPaymentId
	sSql = sSql & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		cAmount = CDbl(oRs("amount"))
	Else
		cAmount = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetLedgerAmount = cAmount

End Function 


'------------------------------------------------------------------------------
' boolean GetPaymentDetails( iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate )
'------------------------------------------------------------------------------
Function GetPaymentDetails( ByVal iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate, ByRef iAdminLocationId, ByRef iJournalEntryTypeId, ByRef sNotes, ByRef iAdminUserId, ByVal iLoggedUserid )
	Dim sSql, oRs

	sSql = "SELECT userid, paymenttotal, paymentdate, ISNULL(adminlocationid,0) AS adminlocationid, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, journalentrytypeid, notes "
	sSql = sSql & " FROM egov_class_payment WHERE paymentid = " & iPaymentId & " AND orgid = " & iorgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		iUserId = oRs("userid")
		If CLng(iLoggedUserid) <> CLng(iUserId) Then 
			GetPaymentDetails = False
		Else 
			nTotal = oRs("paymenttotal")
			dPaymentDate = oRs("paymentdate")
			iAdminLocationId = oRs("adminlocationid")
			iJournalEntryTypeId = oRs("journalentrytypeid")
			sNotes = oRs("notes")
			iAdminUserId = oRs("adminuserid")
			GetPaymentDetails = True 
		End If 
	Else 
		GetPaymentDetails = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' integer GetReceiptHeader( iPaymentId )
'------------------------------------------------------------------------------
Function GetReceiptHeader( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT receipttitle FROM egov_class_payment P, egov_journal_entry_types J "
	sSql = sSql & " WHERE P.journalentrytypeid = J.journalentrytypeid AND P.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetReceiptHeader = oRs("receipttitle")
	Else
		GetReceiptHeader = iPaymentId
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowAccountChange iPaymentId, iUid 
'------------------------------------------------------------------------------
Sub ShowAccountChange( ByVal iPaymentId, ByVal iUid )
	Dim sSql, oRs, cAmount, cPriorBalance, cCurrentBalance, sEntryType, cPrefix

	' Get the activities that they were part of - Should give 1 or 0 rows
	sSql = "SELECT entrytype, amount, priorbalance, plusminus "
	sSql = sSql & " FROM egov_accounts_ledger "
	sSql = sSql & " WHERE accountid = " & iUid & " AND paymentid = " & iPaymentId

	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRS.EOF Then 
		cPriorBalance = CDbl(oRs("priorbalance"))
		cAmount = CDbl(oRs("amount"))
		sEntryType = oRs("entrytype")
		cPrefix = oRs("plusminus")
	Else 
		cPriorBalance = CDbl(0.0)
		cAmount = CDbl(0.0)
		sEntryType = "credit"
		cPrefix = "+"
	End If 
	
	response.write vbcrlf & "<span class=""receipttitles"">Payee Account Information</span><br />"
	response.write "<div id=""accountchange"">"
	response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" id=""accountchange"">"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"">Prior Balance.............................</td><td align=""right"">" & FormatCurrency(cPriorBalance,2) & "</td></tr>"
	
	If sEntryType = "credit" then
		cCurrentBalance = cPriorBalance + cAmount
	Else
		cCurrentBalance = cPriorBalance - cAmount
	End If 
	
	response.write vbcrlf & "<tr><td nowrap=""nowrap"">Change.....................................</td><td id=""changecell"" align=""right"">" & cPrefix & FormatCurrency(cAmount,2) & "</td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"">Current Balance.........................</td><td align=""right"">" & FormatCurrency(cCurrentBalance,2) & "</td></tr>"
	response.write vbcrlf & "</table>"
	response.write "</div>"


	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowDateFeesPaid iReservationDateId, iPaymentId, dTotalCharge
'--------------------------------------------------------------------------------------------------
Sub ShowDateFeesPaid( ByVal iReservationDateId, ByVal iPaymentId, ByRef dTotalCharge )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(amount),0.00) AS amount FROM egov_accounts_ledger "
	sSql = sSql & "WHERE reservationfeetype = 'reservationdatefeeid' AND paymentid = " & iPaymentId
	sSql = sSql & " AND reservationdateid = " & iReservationDateId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		'If CDbl(oRs("amount")) > CDbl(0.0000) Then ' commented out to show no cost rental rates
			response.write vbcrlf & "<tr><td colspan=""3"" align=""right""><strong>Rates:</strong></td>"
			response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(CDbl(oRs("amount")),2) & "</td></tr>"
			dTotalCharge = dTotalCharge + CDbl(oRs("amount"))
		'End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowDateItemsPaid iReservationDateId, iPaymentId, dTotalCharge, sJournalEntryType
'--------------------------------------------------------------------------------------------------
Sub ShowDateItemsPaid( ByVal iReservationDateId, ByVal iPaymentId, ByRef dTotalCharge, ByVal sJournalEntryType )
	Dim sSql, oRs

	sSql = "SELECT I.rentalitem, I.quantity, ISNULL(L.itemquantity,0) AS itemquantity, L.amount "
	sSql = sSql & "FROM egov_accounts_ledger L, egov_rentalreservationdateitems I "
	sSql = sSql & "WHERE L.reservationfeetypeid = I.reservationdateitemid AND L.paymentid = " & iPaymentId
	sSql = sSql & "AND L.reservationfeetype = 'reservationdateitemid' AND L.reservationdateid = " & iReservationDateId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr><td colspan=""3"" align=""right""><strong>Items:</strong></td>&nbsp;</td></tr>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr><td colspan=""3"" align=""right"">"
			If sJournalEntryType = "rentalpayment" Then 
				If CLng(oRs("itemquantity")) > CLng(0) Then 
					response.write oRs("itemquantity") & " "
				End If 
			End If 
			response.write oRs("rentalitem") & "</td>"
			response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(CDbl(oRs("amount")),2) & "</td></tr>"
			dTotalCharge = dTotalCharge + CDbl(oRs("amount"))
			oRs.MoveNext 
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'------------------------------------------------------------------------------
' void ShowPaymentTypes iPaymentId, sJournalEntryType, bHasPaymentFee, dProcessingFee
'------------------------------------------------------------------------------
Sub ShowPaymentTypes( ByVal iPaymentId, ByVal sJournalEntryType, ByVal bHasPaymentFee, ByVal dProcessingFee )
	Dim sSql, oRs, cTotal, sWhere

	If sJournalEntryType <> "refund" Then
		sWhere = " AND isrefundmethod = 0 AND isrefunddebit = 0 "
	Else
		sWhere = ""
	End If 
	cTotal = 0.00
	sSql = "SELECT P.paymenttypeid, P.paymenttypename, P.requirescheckno, P.requirescitizenaccount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O " 
	sSql = sSql & " WHERE P.paymenttypeid = O.paymenttypeid AND O.orgid = " & iorgid & sWhere & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">"
			response.write oRs("paymenttypename") 
			response.write ": &nbsp;</td><td class=""amountcell"">"
			cAmount = GetLedgerAmount( iPaymentId, oRs("paymenttypeid") )
			cTotal = cTotal + cAmount
			If bHasPaymentFee And (CDbl(cAmount) > CDbl(0.00)) Then
				' if it has a processing fee then it is only credit card, so add in the fee
				cTotal = cTotal + dProcessingFee
				response.write FormatCurrency((CDbl(cAmount) + CDbl(dProcessingFee)), 2)
			Else 
				response.write FormatCurrency(cAmount, 2)
			End If 
			response.write "</td><td>"
			If oRs("requirescheckno") Then
				response.write " &nbsp;&nbsp;  Check # " 
				If CDbl(cAmount) > CDbl(0.00) Then 
					response.write GetCheckNo( iPaymentId, oRs("paymenttypeid") )
				End If 
			End If 
			If oRs("requirescitizenaccount") Then
				response.write " &nbsp;&nbsp; From: &nbsp; " 
				If CDbl(cAmount) > CDbl(0.00) Then 
					response.write GetAccountName( iPaymentId, oRs("paymenttypeid") )
				End If 
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">Total: &nbsp;</td><td class=""totalpayment"">" & FormatCurrency(cTotal,2) & "</td><td>&nbsp;</td><tr>"
		response.write vbcrlf & "</table>"
	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowPurchaseDetails iPaymentId, sEntryType, sJournalEntryType 
'------------------------------------------------------------------------------
Sub ShowPurchaseDetails( ByVal iPaymentId, ByVal sEntryType, ByVal sJournalEntryType )
	Dim sSql, oRs, cTotal

	cTotal = CDbl(0.00)

	' Pull a set of items purchased
	sSql = "SELECT T.cartdisplayorder, itemtype, itemid, L.itemtypeid, SUM(amount) AS amount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 0 AND entrytype = '" & sEntryType & "' "
	sSql = sSql & " AND L.paymentid = " & iPaymentId
	sSql = sSql & " GROUP BY T.cartdisplayorder, itemtype, itemid, L.itemtypeid "
	sSql = sSql & " ORDER BY T.cartdisplayorder, itemtype, itemid, L.itemtypeid"
	'sSql = sSql & " group by itemtype, itemid, L.itemtypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		oRs.MoveNext
	Loop 

	oRs.Close 
	Set oRs = Nothing

	If sJournalEntryType = "refund" Then
		cTotal = cTotal - ShowRefundFee( iPaymentId )
	End If 

	response.write vbcrlf & "<hr />"
	response.write "<div id=""receiptdetailtotal""><strong>Purchase Total:</strong> " & FormatCurrency(cTotal,2) & "</div>"

End Sub 


'------------------------------------------------------------------------------
' void ShowReceiptFooter( sNotes, sJournalEntryType )
'------------------------------------------------------------------------------
Sub ShowReceiptFooter( ByVal sNotes, ByVal sJournalEntryType )
	Dim sFooter

	response.write vbcrlf & "<p><strong>Receipt Notes:</strong> " & Trim(sNotes) & "</p>" & vbcrlf

	If sJournalEntryType = "refund" Then 
		sFooter = "rental refund footer"
	Else 
		sFooter = "rental receipt footer"
	End If 

	If OrgHasDisplay( iorgid, sFooter ) Then 
		response.write vbcrlf & "<p>" & GetOrgDisplay( iorgid, sFooter ) & "</p>"
	End If 

End Sub 


'------------------------------------------------------------------------------
' void ShowReceiptHeader( iPaymentId )
'------------------------------------------------------------------------------
Sub ShowReceiptHeader( ByVal iPaymentId )

	If OrgHasDisplay( iorgid, "rental receipt header" ) Then
		response.write vbcrlf & "<p class=""receiptheader"">" & GetOrgDisplay( iorgid, "rental receipt header" ) & "</p>"
		response.write vbcrlf & "<p class=""receiptheader"">"
		response.write "<br /><br />" & GetReceiptHeader( iPaymentId )
		response.write vbcrlf & "</p>"
	Else  
		response.write vbcrlf & "<h3 align=""center"">" & Session("sOrgName") & " " & GetReceiptHeader( iPaymentId ) & "</h3><br /><br />"
	End If 

End Sub 


'------------------------------------------------------------------------------
' double ShowRefundFee( iPaymentId )
'------------------------------------------------------------------------------
Function ShowRefundFee( ByVal iPaymentId )
	Dim sSql, oRs, dTotalAmount

	dTotalAmount = CDbl(0) 
	
	' Pull a the refund fee row
	sSql = "SELECT itemtype, itemid, ISNULL(amount,0.0000) AS amount, paymenttypename "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 AND entrytype = 'credit' "
	sSql = sSql & " AND P.isrefunddebit = 1 AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr><td colspan=""3"" align=""right""><strong>" & oRs("paymenttypename") & ":</strong></td>"
			response.write "<td align=""right"" class=""receiptamount"">-" & FormatCurrency((oRs("amount") + cRefundShortage),2)
			dTotalAmount = dTotalAmount + CDbl(oRs("amount"))
			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowRefundFee = dTotalAmount

End Function 


'------------------------------------------------------------------------------
' void ShowRefundType iPaymentId 
'------------------------------------------------------------------------------
Sub ShowRefundType( ByVal iPaymentId )
	Dim sSql, oRs, cTotal

	sSql = "SELECT ISNULL(accountid,0) AS accountid, amount, priorbalance, plusminus, itemid, ispaymentaccount, paymenttypeid, isccrefund "
	sSql = sSql & " FROM egov_accounts_ledger WHERE plusminus = '-' AND entrytype = 'credit' AND paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">"
		If oRs("ispaymentaccount") Then 
			If oRs("isccrefund") Then 
				response.write "Refund to Credit Card"
			Else 
			' This is a refund voucher
			response.write GetRefundName() 
			End If 
			response.write ": &nbsp;</td><td class=""refundamountcell"" nowrap=""nowrap"">"
			response.write FormatCurrency(oRs("amount"), 2)
		Else
			If CDbl(oRs("amount")) > CDbl(0.00) Then 
				' This is to a citizen account
				response.write "Citizen Account: &nbsp;</td><td class=""refundamountcell"" nowrap=""nowrap"">"
				response.write FormatCurrency(oRs("amount"), 2) & "</td>" 
				response.write "<td> &nbsp; To: &nbsp; " & GetCitizenName( oRs("accountid") )
			Else
				response.write "Removed From The Waitlist &nbsp;</td><td class=""refundamountcell"" nowrap=""nowrap"">&nbsp;</td><td> &nbsp;"
			End If 
		End If 
		response.write "</td></tr>"
'		response.write "<tr><td class=""label"" align=""right"" nowrap=""nowrap"" width=""40%"">Total: &nbsp;</td><td class=""totalpayment"">" & FormatCurrency(cTotal,2) & "</td><td>&nbsp;</td><tr>"
		response.write vbcrlf & "</table>"
	Else
		' No payments were credited, which can only be a removal from a waitlist
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">"
		response.write "Removed From The Waitlist &nbsp;</td><td class=""amountcell"" nowrap=""nowrap"">&nbsp;</td><td> &nbsp;"
		response.write "</td></tr>"
		response.write vbcrlf & "</table>"

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalNotes iReservationDateId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalNotes( ByVal iReservationDateId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(R.receiptnotes,'') AS receiptnotes FROM egov_rentals R, egov_rentalreservationdates D "
	sSql = sSql & "WHERE R.rentalid = D.rentalid AND D.reservationdateid = " & iReservationDateId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("receiptnotes") <> "" Then 
			'response.write vbcrlf & "<tr><td align=""right"" valign=""top""><strong>Rental Notes:</strong></td>"
			response.write vbcrlf & "<tr>"
			response.write "<td colspan=""3"" valign=""top"">" & oRs("receiptnotes") & "</td><td>&nbsp;</td></tr>"
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReservationDetails iReservationId, iPaymentId, sEntryType, sJournalEntryType, bHasPaymentFee, dProcessingFee
'--------------------------------------------------------------------------------------------------
Sub ShowReservationDetails( ByVal iReservationId, ByVal iPaymentId, ByVal sEntryType, ByVal sJournalEntryType, ByVal bHasPaymentFee, ByVal dProcessingFee )
	Dim sSql, oRs, dTotalCharge, dTotalRefundDue

	dTotalCharge = CDbl(0)
	dTotalRefundDue = CDbl(0)

	' Show the daily fees and items
	sSql = "SELECT DISTINCT D.reservationdateid, D.reservationstarttime, D.billingendtime, R.rentalname, L.name AS locationname "
	sSql = sSql & "FROM egov_rentalreservationdates D, egov_accounts_ledger A, egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE A.reservationdateid = D.reservationdateid AND R.rentalid = D.rentalid "
	sSql = sSql & "AND R.locationid = L.locationid AND A.reservationid = " & iReservationId
	sSql = sSql & " AND A.paymentid = " & iPaymentId
	sSql = sSql & " ORDER BY D.reservationstarttime"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<table id=""receiptdatesandfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr>"
		' Date
		response.write "<td><strong>"
		response.write DateValue(oRs("reservationstarttime"))
		response.write "</strong></td>"

		' From and to times
		sStartAmPm = "AM"
		sStartHour = Hour(oRs("reservationstarttime"))
		If clng(sStartHour) = clng(0) Then
			sStartHour = 12
			sStartAmPm = "AM"
		Else
			If clng(sStartHour) > clng(12) Then
				sStartHour = clng(sStartHour) - clng(12)
				sStartAmPm = "PM"
			End If 
			If clng(sStartHour) = clng(12) Then
				sStartAmPm = "PM"
			End If 
		End If 
		sStartMinute = Minute(oRs("reservationstarttime"))
		If sStartMinute < 10 Then
			sStartMinute = "0" & sStartMinute
		End If 

		sEndAmPm = "AM"
		sEndHour = Hour(oRs("billingendtime"))
		If clng(sEndHour) = clng(0) Then
			sEndHour = 12
			sEndAmPm = "AM"
		Else
			If clng(sEndHour) > clng(12) Then
				sEndHour = clng(sEndHour) - clng(12)
				sEndAmPm = "PM"
			End If 
			If clng(sEndHour) = clng(12) Then
				sEndAmPm = "PM"
			End If 
		End If 
		sEndMinute = Minute(oRs("billingendtime"))
		If sEndMinute < 10 Then
			sEndMinute = "0" & sEndMinute
		End If 

		response.write "<td><strong>" & sStartHour & ":" & sStartMinute & " " & sStartAmPm
		response.write " &mdash; " & sEndHour & ":" & sEndMinute & " " & sEndAmPm
		response.write "</strong> (" & CalculateDurationInHours( oRs("reservationstarttime"), oRs("billingendtime") ) & " hours)"
		response.write "</td>"
		
		' Location
		response.write "<td>"
		response.write "<strong>" & oRs("locationname") & " &ndash; " & oRs("rentalname") & "</strong>"
		response.write "</td>"

		response.write "<td>&nbsp;</td>"

		response.write "</tr>"

		If sJournalEntryType = "rentalpayment" Then
			ShowRentalNotes oRs("reservationdateid")
		End If 

		ShowDateFeesPaid oRs("reservationdateid"), iPaymentId, dTotalCharge

		ShowDateItemsPaid oRs("reservationdateid"), iPaymentId, dTotalCharge, sJournalEntryType
		
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	' Show the reservation fees
	bHasReservationFees = ShowReservationFees( iReservationId, iPaymentId, dTotalCharge )

	If bHasPaymentFee Then
		' Add the Processing Fee to the total charged
		dTotalCharge = dTotalCharge + dProcessingFee

		If Not bHasReservationFees Then 
			' if there are no reservation fees we still want the seperator row if there is a processing fee
			response.write vbcrlf & "<tr class=""feeseperator""><td colspan=""4"">&nbsp;</td></tr>"
		End If 

		' Show the processing fee
		response.write vbcrlf & "<tr id=""processingfeerow""><td colspan=""3"" align=""right""><strong>Processing Fee:</strong></td>"
		response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(dProcessingFee,2) & "</td></tr>"
	End If 

	' Show the total charges paid
	If sJournalEntryType = "refund" Then
		dTotalCharge = dTotalCharge - ShowRefundFee( iPaymentId )
	End If 

	response.write vbcrlf & "</table>"

	response.write "<table id=""receiptfeetotal"" cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf
	response.write "  <tr>" & vbcrlf
	response.write "      <td colspan=""4"" align=""right""><strong>Total:</strong></td>" & vbcrlf
	response.write "      <td align=""right"" class=""receiptamount"">" & FormatCurrency(dTotalCharge,2) & "</td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf

	'BEGIN: Payment History ------------------------------------------------------
	lcl_total_charges = GetReservationTotalAmount(iReservationID, "totalamount")

	'Total Charges Row
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 

	response.write "<table id=""receiptfeetotal2"" cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf
	response.write "  <tr" & sClass & ">" & vbcrlf
	response.write "      <td class=""totalscell"" colspan=""2"" align=""right"">&nbsp;</td>"
	response.write "      <td class=""totalscell"" align=""right""><strong>Total Charges</strong></td>"
	response.write "      <td class=""totalscell"" align=""right"">"
	response.write            FormatNumber(lcl_total_charges,2,,,0) 
	response.write "      </td>" & vbcrlf
	response.write "  </tr>"

	'Total Paid Row
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 

	response.write "  <tr" & sClass & ">" & vbcrlf
	response.write "      <td id=""receiptRowPayments"" class=""totalscell"" colspan=""4"" align=""center""><strong>Payments</strong></td>" & vbcrlf
	response.write "  </tr>" & vbcrlf

	'Want to show payments with link to receipt page here
	ShowReservationPayments iReservationId, sClass
	dTotalPaid = GetReservationTotalAmount( iReservationId, "totalpaid" ) ' In rentalscommonfunctions.asp

	response.write "  <tr" & sClass & ">" & vbcrlf
	response.write "      <td id=""receiptRowTotalPaid"" colspan=""3"" align=""right"" class=""totalscell""><strong>Total Paid</strong></td>" & vbcrlf
	response.write "      <td id=""receiptRowTotalPaid"" align=""right"" class=""totalscell"">" & FormatNumber(dTotalPaid,2,,,0) & "</td>" & vbcrlf
	response.write "  </tr>" & vbcrlf

	'Refund Row
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 

	response.write "<tr" & sClass & ">" 
	response.write "<td class=""totalscell"" colspan=""4"" align=""center""><strong>Refunds</strong></td>"
	response.write "</tr>" & vbcrlf

	'Want to show refunds with link to receipt page here
	ShowReservationRefunds iReservationId, sClass
	dTotalRefunded = GetReservationTotalAmount( iReservationId, "totalrefunded" ) ' In rentalscommonfunctions.asp - is just the pull of the total field

	response.write "<tr" & sClass & ">"
	response.write "<td id=""receiptRowTotalRefunds"" class=""totalscell"" colspan=""2"">&nbsp;</td>"
	response.write "<td id=""receiptRowTotalRefunds"" class=""totalscell"" align=""right""><strong>Total Refunds</strong></td>"
	response.write "<td id=""receiptRowTotalRefunds"" class=""totalscell"" align=""right"">" & FormatNumber(dTotalRefunded,2,,,0) & "</td>"
	response.write "  </tr>" & vbcrlf

	' Refund Due Row
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 
	' get the refund due amount
	dTotalRefundDue = GetReservationRefundDue( iReservationId )
	response.write vbcrlf & "<tr" & sClass & ">"
	response.write "<td class=""totalscell"" colspan=""2"">&nbsp;</td>"
	response.write "<td class=""totalscell"" align=""right""><strong>Refund Due</strong></td>"
	response.write "<td class=""totalscell"" align=""right"">"
	response.write FormatNumber(dTotalRefundDue,2,,,0) 
	response.write "</td></tr>"

	'Balance Due Row
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 

	dBalanceDue = (lcl_total_charges + dTotalRefunded) - (dTotalPaid - dTotalRefundDue)

	response.write "<tr" & sClass & ">"
	response.write "<td class=""totalscell"" colspan=""3"" align=""right""><strong>Balance Due</strong></td>"
	response.write "<td class=""totalscell"" align=""right"">" & FormatNumber(dBalanceDue,2,,,0) & "</td>"
	response.write "</tr>" & vbcrlf

	response.write "</table>" & vbcrlf
	'END: Payment History ---------------------------------------------------------

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean ShowReservationFees( iReservationId, iPaymentId, dTotalCharge )
'--------------------------------------------------------------------------------------------------
Function ShowReservationFees( ByVal iReservationId, ByVal iPaymentId, ByRef dTotalCharge )
	Dim sSql, oRs

	sSql = "SELECT P.pricetypename, L.amount "
	sSql = sSql & "FROM egov_accounts_ledger L, egov_rentalreservationfees F, egov_price_types P "
	sSql = sSql & "WHERE L.reservationfeetype = 'reservationfeeid' AND L.reservationid = " & iReservationId
	sSql = sSql & " AND L.reservationfeetypeid = F.reservationfeeid AND "
	sSql = sSql & "F.pricetypeid = p.pricetypeid AND L.paymentid = " & iPaymentId
	sSql = sSql & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		ShowReservationFees = True 
		response.write vbcrlf & "<tr class=""feeseperator""><td colspan=""4"">&nbsp;</td></tr>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr><td colspan=""3"" align=""right""><strong>" & oRs("pricetypename") & ":</strong></td>"
			response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(CDbl(oRs("amount")),2) & "</td></tr>"
			dTotalCharge = dTotalCharge + CDbl(oRs("amount"))
			oRs.MoveNext 
		Loop 
	Else
		ShowReservationFees = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowReservationReceiptNotes iReservationId, sJournalEntryType
'--------------------------------------------------------------------------------------------------
Sub ShowReservationReceiptNotes( ByVal iReservationId, ByVal sJournalEntryType )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(receiptnotes,'') AS receiptnotes FROM egov_rentalreservations "
	sSql = sSql & "WHERE reservationid = " & iReservationId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("receiptnotes") <> "" Then 
			response.write vbcrlf & "<p><strong>"
			If sJournalEntryType = "rentalpayment" Then 
				response.write "Reservation"
			Else 
				response.write "Refund"
			End If 
			response.write " Notes:</strong> "
			response.write oRs("receiptnotes") & "</p>"
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' string ShowUserInfo( iUserId, sJournalEntryType )
'------------------------------------------------------------------------------
Function ShowUserInfo( ByVal iUserId, ByVal sJournalEntryType )
	Dim oCmd, sResidentDesc, sUserType
	ShowUserInfo = ""

	sUserType = GetUserResidentType(iUserid)
	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iUserid, Session("OrgID"))
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 1, 4, iUserId)
	    Set oRs = .Execute
	End With
	
	ShowUserInfo = ShowUserInfo & "<span class=""receipttitles"">"
	If sJournalEntryType = "refund" Then
		ShowUserInfo = ShowUserInfo & "Refundee "
	Else 
		ShowUserInfo = ShowUserInfo & "Payee "
	End If 
	ShowUserInfo = ShowUserInfo & "Information</span><br />"
	ShowUserInfo = ShowUserInfo & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & oRs("userfname") & " " & oRs("userlname") & "</strong><br />"
	ShowUserInfo = ShowUserInfo & "<strong>" & oRs("useraddress") 
	If oRs("userunit") <> "" Then 
		ShowUserInfo = ShowUserInfo & "&nbsp;&nbsp;" & oRs("userunit")
	End If
	If oRs("useraddress2") <> "" Then 
		ShowUserInfo = ShowUserInfo & "<br />" & oRs("useraddress2")
	End If 
	ShowUserInfo = ShowUserInfo & "<br />" & oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "</strong></td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td colspan=""2"">&nbsp;</td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & GetFamilyEmail( iUserId ) & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhoneNumber(oRs("userhomephone")) & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "</table>"

	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing
	
End Function 


'------------------------------------------------------------------------------
' boolean ReservationHasAssociatedDocuments( iReservationId )
'------------------------------------------------------------------------------
Function ReservationHasAssociatedDocuments( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(documentid) AS hits "
	sSql = sSql & "FROM egov_rentalreservationdates R, egov_rentaldocuments D "
	sSql = sSql & "WHERE R.rentalid = D.rentalid AND R.reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(ors("hits")) > CLng(0) Then
			ReservationHasAssociatedDocuments = True 
		Else
			ReservationHasAssociatedDocuments = False 
		End If 
	Else
		ReservationHasAssociatedDocuments = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void ShowReservationDocuments iReservationId
'------------------------------------------------------------------------------
Sub ShowReservationDocuments( ByVal iReservationId )
	Dim sSql, oRs, iOldRentalId

	iOldRentalId = CLng(0)

	sSql = "SELECT DISTINCT L.name AS locationname, R.rentalname, R.rentalid, "
	sSql = sSql & "D.documentid, D.documenturl, D.documenttitle "
	sSql = sSql & "FROM egov_rentalreservationdates RD, egov_rentaldocuments D, egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.rentalid = RD.rentalid AND RD.rentalid = D.rentalid "
	sSql = sSql & "AND R.locationid = L.locationid AND RD.reservationid = " & iReservationId
	sSql = sSql & "ORDER BY L.name, R.rentalname, D.documenttitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<div id=""reservationdocs"">"
		Do While Not oRs.EOF
			If iOldRentalId <> CLng(oRs("rentalid")) Then
				iOldRentalId = CLng(oRs("rentalid"))
				response.write vbcrlf & "<div class=""receiptrentalname"">" & oRs("locationname") & " &ndash; " & oRs("rentalname") & "</div>"
			End If 
			response.write vbcrlf & "<div class=""receiptrentaldoc""><a href=""" & oRs("documenturl") & """ target=""_blank"">" & oRs("documenttitle") & "</a></div>"
			oRs.MoveNext 
		Loop
		response.write vbcrlf & "</div>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' string sCheck = GetCheckNo( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetCheckNo( ByVal iPaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT checkno FROM egov_verisign_payment_information WHERE paymentid = " & iPaymentId
	sSql = sSql & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetCheckNo = oRs("checkno")
	Else
		GetCheckNo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' double GetReservationRefundDue( iReservationId )
'------------------------------------------------------------------------------
Function GetReservationRefundDue( ByVal iReservationId )
	Dim sSql, oRs, iRefundDue

	iRefundDue = CDbl(0)

	sSql = "SELECT reservationid, ISNULL(SUM( paidamount - (feeamount + refundamount)),0) AS refunddue "
	sSql = sSql & "FROM egov_rentalreservationdatefees  "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND paidamount > 0 AND paidamount > (feeamount + refundamount) "
	sSql = sSql & "GROUP BY reservationid "
	sSql = sSql & "UNION SELECT reservationid, ISNULL(SUM( paidamount - (feeamount + refundamount)),0) AS refunddue  "
	sSql = sSql & "FROM egov_rentalreservationdateitems  "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND paidamount > 0 AND paidamount > (feeamount + refundamount) "
	sSql = sSql & "GROUP BY reservationid "
	sSql = sSql & "UNION SELECT reservationid, ISNULL(SUM( paidamount - (feeamount + refundamount)),0) AS refunddue  "
	sSql = sSql & "FROM egov_rentalreservationfees  "
	sSql = sSql & "WHERE reservationid = " & iReservationId & " AND paidamount > 0 AND paidamount > (feeamount + refundamount) "
	sSql = sSql & "GROUP BY reservationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRefundDue = iRefundDue + CDbl(oRs("refunddue"))
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

	GetReservationRefundDue = iRefundDue

End Function 




%>
