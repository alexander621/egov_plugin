<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: viewpaymentreceipt.asp
' AUTHOR: Steve Loar
' CREATED: 11/16/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the receipt for a rental reservation payment, or refund 
'
' MODIFICATION HISTORY
' 1.0	11/16/2009	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iPaymentId, iUid, sReceiptType, sJEType, iDisplayType, sNotes, iAdminuserid
Dim lcl_orghasfeature_citizen_accounts, iReservationId, iPriorPaymentId, sTotalLabel
Dim iUserId, nTotal, dPaymentDate, sSql, nRowTotal, iAdminLocationId, iJournalEntryTypeId
Dim sJournalEntryType, bIsCCRefund, sReservationTypeSelector, sRenterName, sRenterPhone
Dim dProcessingFee, bHasPaymentFee, lcl_total_display

iPriorPaymentId = ""

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit reservations", sLevel	' In common.asp

iPaymentId = CLng(request("paymentid"))

iReservationId = GetReservationIdFromPaymentId( iPaymentId )	' in rentalscommonfunctions.asp

'sReservationTypeSelector = GetReservationTypeSelection( GetReservationTypeId( iReservationId ) )

GetGeneralReservationData iReservationId

'Check for org features
lcl_orghasfeature_citizen_accounts = OrgHasFeature( "citizen accounts" )

%>
<html>
<head>
	<title>E-Gov Administration Console {Receipt}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="javascript">
	<!--

		window.onload = function()
		{
			//factory.printing.header = "Printed on &d"
			//factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
			//factory.printing.portrait     = true;
			//factory.printing.leftMargin   = 0.5;
			//factory.printing.topMargin    = 0.5;
			//factory.printing.rightMargin  = 0.5;
			//factory.printing.bottomMargin = 0.5;

			// enable control buttons
			//var templateSupported = factory.printing.IsTemplateSupported();
			//var controls = idControls.all.tags("input");
			//for ( i = 0; i < controls.length; i++ ) 
			//{
			//	controls[i].disabled = false;
			//	if (templateSupported && controls[i].className == "ie55" )
			//		controls[i].style.display = "inline";
			//}
		}

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
<%
	'<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
	'<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />&nbsp;&nbsp;
%>
<%	If request("rt") = "r" Then %>
		&nbsp;&nbsp;<input type="button" class="button" value="<< Back To Reservation" onclick="location.href='reservationedit.asp?reservationid=<%=iReservationId%>';" />	
<%	ElseIf request("rt") = "b" Then %>
		&nbsp;&nbsp;<input type="button" class="button" value="<< Back" onclick="history.go(-1);" />	
<%	End If	%>
</div>

<%
'<object id="factory" viewastext  style="display:none"
'   classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
'   codebase="../includes/smsx.cab#Version=6,3,434,12">
'</object>
%>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">
<%
	ShowReceiptHeader iPaymentId

	response.write vbcrlf & "<hr />"


	If GetPaymentDetails(iPaymentId, iUserId, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminuserid ) Then 

		sJournalEntryType = GetJournalEntryType( iJournalEntryTypeId )

		response.write "<span id=""receiptadmininfo"">"
		response.write " Location: " & GetAdminLocation( iAdminLocationId ) & "&nbsp;&nbsp;"   'In ../includes/common.asp
		response.write " Administrator: " & GetAdminName( iAdminuserid )   'In ../includes/common.asp
		response.write "</span>"
		response.write " Date: " & DateValue(CDate(dPaymentDate)) & "&nbsp;&nbsp;" 
		response.write " Receipt: " & iPaymentId & "&nbsp;&nbsp;" 
		response.write " Reservation Id: <a href=""reservationedit.asp?reservationid=" & iReservationId & """>" & iReservationId & "</a>&nbsp;&nbsp;" 

		response.write vbcrlf & "<hr />" 

 		response.write vbcrlf & "<div id=""receipttopright"">" 

		If sJournalEntryType = "refund" Then 
			bHasPaymentFee = False 
			dProcessingFee = CDbl("0.00")
			sTotalLabel = "Amount Refunded:"
			lcl_total_display = nTotal
		Else 
			' See if the gateway for this org has fees they charge the citizen
			If PaymentGatewayRequiresFeeCheck( session("orgid") ) Then
				bHasPaymentFee = True 
				dProcessingFee = GetProcessingFee( iPaymentId )
			Else
				bHasPaymentFee = False 
				dProcessingFee = CDbl("0.00")
			End If 
			sTotalLabel = "Transaction Total:"
			lcl_total_display = nTotal
		End If 

		response.write vbcrlf & "<p id=""transactiontotal"">" 
		response.write "<strong>" & sTotalLabel & "</strong> " & FormatCurrency((CDbl(lcl_total_display) + CDbl(dProcessingFee)),2)
		response.write "</p>" 
			
		If lcl_orghasfeature_citizen_accounts And sReservationTypeSelector = "public" Then 
		   ShowAccountChange iPaymentId, iUserId
		End If 

  		response.write "</div>"

		If sReservationTypeSelector = "public" Then 
  			response.write ShowUserInfo( iUserId, sJournalEntryType )
		Else 
			' Admin Renter
			response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
			response.write vbcrlf & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & sRenterName & "</strong></td></tr>"
			response.write vbcrlf & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td>" & sRenterPhone & "</td></tr>"
			response.write vbcrlf & "</table>"
		End If 
		response.write vbcrlf & "<hr />" 

  		If sJournalEntryType = "refund" Then 
    			ShowRefundType iPaymentId
  		Else 
		     ShowPaymentTypes iPaymentId, sJournalEntryType, bHasPaymentFee, dProcessingFee
  		End If 

  		response.write vbcrlf & "<hr />" 
 		response.write vbcrlf & "<strong>Reservations</strong>" 
		response.write vbcrlf & "<hr />" 
	'	'response.write "[" & sJournalEntryType & "]"

  		Select Case sJournalEntryType
			Case "rentalpayment"
				'Show purchase details
				ShowReservationDetails iReservationId, iPaymentId, "credit", sJournalEntryType, bHasPaymentFee, dProcessingFee

			Case "refund"
				'Show refund stuff
				ShowReservationDetails iReservationId, iPaymentId, "debit", sJournalEntryType, bHasPaymentFee, dProcessingFee
		End Select 

	 Else 
			response.write "<p>No details could be found for this receipt.</p>" 
	 End If 

	response.write vbcrlf & "<hr />" 

	'response.write vbcrlf & "<p><strong>Receipt Notes:</strong> " & Trim(sNotes) & "</p>"

	ShowReservationReceiptNotes iReservationId, sJournalEntryType

	ShowReceiptFooter sNotes, sJournalEntryType

%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowReservationDetails iReservationId, iPaymentId, sEntryType, sJournalEntryType, bHasPaymentFee, dProcessingFee
'--------------------------------------------------------------------------------------------------
Sub ShowReservationDetails( ByVal iReservationId, ByVal iPaymentId, ByVal sEntryType, ByVal sJournalEntryType, ByVal bHasPaymentFee, ByVal dProcessingFee )
	Dim sSql, oRs, dTotalCharge

	dTotalCharge = CDbl(0.0000)

	' Show the daily fees and items
	sSql = "SELECT DISTINCT D.reservationdateid, D.reservationstarttime, D.billingendtime, R.rentalname, L.name AS locationname "
	sSql = sSql & "FROM egov_rentalreservationdates D, egov_accounts_ledger A, egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE A.reservationdateid = D.reservationdateid AND R.rentalid = D.rentalid "
	sSql = sSql & "AND R.locationid = L.locationid AND A.reservationid = " & iReservationId
	sSql = sSql & " AND A.paymentid = " & iPaymentId
	sSql = sSql & " ORDER BY D.reservationstarttime"

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
		response.write "</strong> (" & FormatNumber(CalculateDurationInHours( oRs("reservationstarttime"), oRs("billingendtime") ),2) & " hours)"
		response.write "</td>"
		
		' Location
		response.write "<td>"
		response.write "<strong>" & oRs("rentalname") & ", " & oRs("locationname") & "</strong>"
		response.write "</td>"

		response.write "<td>&nbsp;</td>"

		response.write "</tr>"

		If sJournalEntryType = "rentalpayment" Then
			ShowRentalNotes oRs("reservationdateid")
		End If 

		ShowDateFeesPaid oRs("reservationdateid"), iPaymentId, dTotalCharge

		ShowDateItemsPaid oRs("reservationdateid"), iPaymentId, dTotalCharge, sJournalEntryType

		response.write vbcrlf & "</table>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	' Show the reservation fees
	ShowReservationFees iReservationId, iPaymentId, dTotalCharge

	If bHasPaymentFee Then
		' Add the Processing Fee to the total charged
		dTotalCharge = dTotalCharge + dProcessingFee

		' Show the processing fee
		response.write vbcrlf & "<table id=""receiptprocessingfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr id=""processingfeerow""><td colspan=""4"" align=""right""><strong>Processing Fee:</strong></td>"
		response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(dProcessingFee,2) & "</td></tr>"
		response.write vbcrlf & "</table>"
	End If 

	' Show the total charges paid
	If sJournalEntryType = "refund" Then
		dTotalCharge = dTotalCharge - ShowRefundFee( iPaymentId )
		'response.write vbcrlf & "<table id=""receiptfeetotal"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		'response.write vbcrlf & "<tr><td colspan=""4"" align=""right""><strong>Total Charges:</strong></td>"
		'response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(dTotalCharge,2) & "</td></tr>"
		'response.write vbcrlf & "</table>"
	'Else
	'	response.write vbcrlf & "<table id=""receiptfeetotal"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
	'	response.write vbcrlf & "<tr><td colspan=""4"" align=""right""><strong>Total Charges:</strong></td>"
	'	response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(dTotalCharge,2) & "</td></tr>"
	'	response.write vbcrlf & "</table>"
	End If

	response.write "<table id=""receiptfeetotal"" cellpadding=""2"" cellspacing=""0"" border=""0"">" & vbcrlf
	response.write "  <tr>" & vbcrlf
	response.write "      <td colspan=""4"" align=""right""><strong>Total:</strong></td>" & vbcrlf
	response.write "      <td align=""right"" class=""receiptamount"">" & FormatCurrency(dTotalCharge,2) & "</td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf

	If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
		lcl_total_charges = GetReservationTotalAmount(iReservationID, "totalamount")

		' Total Charges Row
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

		' Total Paid Row
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

		If sReservationTypeSelector = "public" Then
			response.write "  <tr" & sClass & ">" & vbcrlf
			'response.write "      <td class=""totalscell"" colspan=""4"" align=""center""><strong>Refunds</strong></td>" & vbcrlf
			response.write "      <td class=""totalscell"" colspan=""4"" align=""center""><strong>Refunds & Refund Fees</strong></td>" & vbcrlf
			response.write "  </tr>" & vbcrlf

			'Want to show refunds with link to receipt page here
			' ShowReservationRefunds iReservationId, sClass
			'dTotalRefunded = GetReservationTotalAmount( iReservationId, "totalrefunded" ) ' In rentalscommonfunctions.asp - is just the pull of the total field
			dTotalRefunded = displayReservationRefunds( iReservationId, sClass )  ' In rentalscommonfunctions.asp

			response.write "  <tr" & sClass & ">" & vbcrlf
			response.write "      <td id=""receiptRowTotalRefunds"" class=""totalscell"" colspan=""2"">&nbsp;</td>" & vbcrlf
			response.write "      <td id=""receiptRowTotalRefunds"" class=""totalscell"" align=""right""><strong>Total Refunds & Refund Fees</strong></td>" & vbcrlf
			response.write "      <td id=""receiptRowTotalRefunds"" class=""totalscell"" align=""right"">" & FormatNumber(dTotalRefunded,2,,,0) & "</td>" & vbcrlf
			response.write "  </tr>" & vbcrlf
		End If 

		'Balance Due Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
  			sClass = " class=""altrow"" "
		Else
  			sClass = ""
		End If 

		dBalanceDue = (lcl_total_charges + dTotalRefunded) - dTotalPaid

		response.write "  <tr" & sClass & ">" & vbcrlf
		response.write "      <td class=""totalscell"" colspan=""3"" align=""right""><strong>Balance Due</strong></td>"
		response.write "      <td class=""totalscell"" align=""right"">" & FormatNumber(dBalanceDue,2,,,0) & "</td>" & vbcrlf
		response.write "  </tr>" & vbcrlf
		response.write "</table>" & vbcrlf
	End If 

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
		'response.write vbcrlf & "<tr><td align=""right"" valign=""top""><strong>Rental Notes:</strong></td>"
		response.write vbcrlf & "<tr><td colspan=""3"" valign=""top"">" & oRs("receiptnotes") & "</td><td>&nbsp;</td></tr>"
	End If 

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


'--------------------------------------------------------------------------------------------------
' void ShowReservationFees iReservationId, iPaymentId, dTotalCharge
'--------------------------------------------------------------------------------------------------
Sub ShowReservationFees( ByVal iReservationId, ByVal iPaymentId, ByRef dTotalCharge )
	Dim sSql, oRs

	sSql = "SELECT P.pricetypename, L.amount "
	sSql = sSql & "FROM egov_accounts_ledger L, egov_rentalreservationfees F, egov_price_types P "
	sSql = sSql & "WHERE L.reservationfeetype = 'reservationfeeid' AND L.reservationid = " & iReservationId
	sSql = sSql & " AND L.reservationfeetypeid = F.reservationfeeid AND F.pricetypeid = p.pricetypeid AND L.paymentid = " & iPaymentId
	sSql = sSql & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""receiptreservationfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr><td colspan=""4"" align=""right""><strong>" & oRs("pricetypename") & ":</strong></td>"
			response.write "<td align=""right"" class=""receiptamount"">" & FormatCurrency(CDbl(oRs("amount")),2) & "</td></tr>"
			dTotalCharge = dTotalCharge + CDbl(oRs("amount"))
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' double ShowRefundFee( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function ShowRefundFee( ByVal iPaymentId )
	Dim sSql, oRs, cRefundShortage, dTotalAmount

	dTotalAmount = CDbl(0.00) 

	'cRefundShortage = cTotal - GetRefundDebit( iPaymentId )
	
	' Pull the refund fee rows
	sSql = "SELECT itemtype, itemid, ISNULL(amount,0.0000) AS amount, paymenttypename "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 AND entrytype = 'credit' "
	sSql = sSql & " AND P.isrefunddebit = 1 AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table id=""refundfeestable"" border=""0"" cellpadding=""2"" cellspacing=""0"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr><td colspan=""4"" align=""right""><strong>" & oRs("paymenttypename") & ":</strong></td>"
			response.write "<td align=""right"" class=""receiptamount"">-" & FormatCurrency((oRs("amount") + cRefundShortage),2) & "</td></tr>"
			dTotalAmount = dTotalAmount + CDbl(oRs("amount"))
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
		'response.write vbcrlf & "<p class=""receiptnotes"">&nbsp;</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowRefundFee = dTotalAmount

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetVoucherAmount( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetVoucherAmount( ByVal iPaymentId )
	Dim sSql, oRs
	
	' Pull a the refund voucher row
	sSql = "Select itemtype, itemid, amount from egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " where L.itemtypeid = T.itemtypeid and L.ispaymentaccount = 1 and entrytype = 'credit' "
	sSql = sSql & " and P.isrefunddebit = 0 and P.paymenttypeid = L.paymenttypeid and L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetVoucherAmount = CDbl(oRs("amount"))
	Else
		GetVoucherAmount = CDbl(0.00)
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetRefundFeeAmount( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetRefundFeeAmount( ByVal iPaymentId )
	Dim sSql, oRs
	
	' Pull a the refund fee row
	sSql = "Select amount from egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " where L.itemtypeid = T.itemtypeid and L.ispaymentaccount = 1 and entrytype = 'credit' "
	sSql = sSql & " and P.isrefunddebit = 1 and P.paymenttypeid = L.paymenttypeid and L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRefundFeeAmount = CDbl(oRs("amount"))
	Else
		GetRefundFeeAmount = CDbl(0.00)
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPurchaseTotal( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetPurchaseTotal( ByVal iPaymentId )
	Dim sSql, oRs
	
	' Pull a the purchase total sum
	sSql = "Select sum(amount) as amount from egov_accounts_ledger L, egov_item_types T "
	sSql = sSql & " where L.itemtypeid = T.itemtypeid and L.ispaymentaccount = 0 and entrytype = 'debit' "
	sSql = sSql & " and L.paymentid = " & iPaymentId & " group by L.paymentid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPurchaseTotal = CDbl(oRs("amount"))
	Else
		GetPurchaseTotal = CDbl(0.00)
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' string sJournalEntryType = GetJournalEntryType( iJournalEntryTypeId )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryType( ByVal iJournalEntryTypeId )
	Dim sSql, oRs

	If iJournalEntryTypeId <> "" Then 
		sSql = "SELECT ISNULL(journalentrytype,'') AS journalentrytype FROM egov_journal_entry_types "
		sSql = sSql & "WHERE journalentrytypeid = " & iJournalEntryTypeId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetJournalEntryType = oRs("journalentrytype")
		Else
			GetJournalEntryType = ""
		End If 

		oRs.Close 
		Set oRs = Nothing
	Else 
		GetJournalEntryType = ""
	End If 

End Function 


'--------------------------------------------------------------------------------------------------
' ShowReceiptHeader iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptHeader( ByVal iPaymentId )

	If OrgHasDisplay( Session("orgid"), "rental receipt header" ) Then
		response.write vbcrlf & "<p class=""receiptheader"">" & GetOrgDisplay( session("orgid"), "rental receipt header" ) 
		response.write vbcrlf & "<br /><br />" & GetReceiptHeader( iPaymentId )
		response.write vbcrlf & "</p>"
	Else  
		response.write vbcrlf & "<h3>" & Session("sOrgName") & " " & GetReceiptHeader( iPaymentId ) & "</h3><br /><br />"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' string sTitle = GetReceiptHeader( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetReceiptHeader( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT receipttitle FROM egov_class_payment P, egov_journal_entry_types J "
	sSql = sSql & " WHERE P.journalentrytypeid = J.journalentrytypeid AND P.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		GetReceiptHeader = oRs("receipttitle")
	Else
		GetReceiptHeader = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' ShowReceiptFooter sNotes, sJournalEntryType
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptFooter( ByVal sNotes, ByVal sJournalEntryType )
	Dim sFooter

	response.write vbcrlf & "<p><strong>Receipt Notes:</strong> " & Trim(sNotes) & "</p>" & vbcrlf

	If sJournalEntryType = "refund" then
  		sFooter = "rental refund footer"
	Else 
  		sFooter = "rental receipt footer"
	End If 

	If OrgHasDisplay( session("orgid"), sFooter ) Then 
  		response.write vbcrlf & "<p>" & GetOrgDisplay( session("orgid"), sFooter ) & "</p>"
	End If 

End Sub   


'--------------------------------------------------------------------------------------------------
' boolean bHasDetails = GetPaymentDetails(iPaymentId, iUserId, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminuserid)
'--------------------------------------------------------------------------------------------------
Function GetPaymentDetails( ByVal iPaymentId, ByRef iUserId, ByRef nTotal, ByRef dPaymentDate, ByRef iAdminLocationId, ByRef iJournalEntryTypeId, ByRef sNotes, ByRef iAdminuserid )
	Dim sSql, oRs

	sSql = "SELECT userid, paymenttotal, paymentdate, ISNULL(adminlocationid,0) AS adminlocationid, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, journalentrytypeid, notes "
	sSql = sSql & " FROM egov_class_payment "
	sSql = sSql & " WHERE paymentid = " & iPaymentId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.eof Then 
		iUserId = oRs("userid")
		nTotal = CDbl(oRs("paymenttotal"))
		dPaymentDate = DateValue(CDate(oRs("paymentdate")))
		iAdminLocationId = oRs("adminlocationid")
		iJournalEntryTypeId = oRs("journalentrytypeid")
		sNotes = oRs("notes")
		iAdminuserid = oRs("adminuserid")
		GetPaymentDetails = True 
	Else 
		iUserId = 0
		nTotal = CDbl(0.00)
		dPaymentDate = DateValue(Now())
		iAdminLocationId = 0
		iJournalEntryTypeId = 0
		sNotes = ""
		iAdminuserid = 0
		GetPaymentDetails = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' ShowAccountChange iPaymentId, iUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowAccountChange( ByVal iPaymentId, ByVal iUserId )
	Dim sSql, oRs, cAmount, cPriorBalance, cCurrentBalance, sEntryType, cPrefix

	' Get the activities that they were part of - Should give 1 or 0 rows
	sSql = "SELECT entrytype, amount, priorbalance, plusminus "
	sSql = sSql & " FROM egov_accounts_ledger "
	sSql = sSql & " WHERE accountid = " & iUserId & " AND paymentid = " & iPaymentId

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

	oRs.Close
	Set oRs = Nothing 
	
	response.write vbcrlf & "<span class=""receipttitles"">Payee Account Information</span><br />"
	response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" id=""accountchange"">"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"">Prior Balance.............................</td><td align=""right"">" & FormatCurrency(cPriorBalance,2) & "</td></tr>"
	
	If sEntryType = "credit" then
		cCurrentBalance = cPriorBalance + cAmount
	Else
		cCurrentBalance = cPriorBalance - cAmount
	End If 
	
	response.write vbcrlf & "<tr><td nowrap=""nowrap"">Change.....................................</td><td id=""changecell"" align=""right"">" & cPrefix & FormatCurrency(cAmount,2) & "</td></tr>"
	response.write vbcrlf & "<tr><td nowrap=""nowrap"">Current Balance.........................</td><td align=""right"">" & FormatCurrency(cCurrentBalance,2) & "</td></tr>"
	response.write vbcrlf & "</table>"

End Sub 


'--------------------------------------------------------------------------------------------------
' string sDisplay = ShowUserInfo( iUserId, sJournalEntryType )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( ByVal iUserId, ByVal sJournalEntryType )
	Dim oCmd, sResidentDesc, sUserType, sDisplayText, oRs

	sDisplayText = ""

	sUserType = GetUserResidentType( iUserId )

	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress( iUserId, Session("orgid") )
	End If 

	sResidentDesc = GetResidentTypeDesc( sUserType )

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iuserid", 3, 1, 4, iUserId)
	    Set oRs = .Execute
	End With
	
	sDisplayText = sDisplayText  & "<span class=""receipttitles"">"
	If sJournalEntryType = "refund" Then
		sDisplayText = sDisplayText & "Refundee "
	Else 
		sDisplayText = sDisplayText & "Payee "
	End If 
	sDisplayText = sDisplayText & "Information</span><br />"
	sDisplayText = sDisplayText  & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
	sDisplayText = sDisplayText  & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & oRs("userfname") & " " & oRs("userlname") & "</strong><br />"
	sDisplayText = sDisplayText & "<strong>" & oRs("useraddress") 
	If oRs("userunit") <> "" Then 
		sDisplayText = sDisplayText & "&nbsp;&nbsp;" & oRs("userunit") 
	End If
	If oRs("useraddress2") <> "" Then 
		sDisplayText = sDisplayText & "<br />" & oRs("useraddress2") 
	End If 
	sDisplayText = sDisplayText & "<br />" & oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "</strong></td></tr>"
	sDisplayText = sDisplayText  & "<tr><td colspan=""2"">&nbsp;</td></tr>"
	sDisplayText = sDisplayText  & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & GetFamilyEmail( iuserid ) & "</td></tr>"
	sDisplayText = sDisplayText  & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhoneNumber(oRs("userhomephone")) & "</td></tr>"
	'sDisplayText = sDisplayText & "<tr><td width=""85"" align=""right"" valign=""top"">Business:</td><td>" & oRs("userbusinessname") & "</td></tr>"
	sDisplayText = sDisplayText  & "</table>"

	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing

	ShowUserInfo = sDisplayText
	
End Function 


'--------------------------------------------------------------------------------------------------
' ShowPaymentTypes iPaymentId, sJournalEntryType, bHasPaymentFee, dProcessingFee
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentTypes( ByVal iPaymentId, ByVal sJournalEntryType, ByVal bHasPaymentFee, ByVal dProcessingFee)
	Dim sSql, oRs, cTotal, sWhere

	If sJournalEntryType <> "refund" Then
		sWhere = " AND isrefundmethod = 0 AND isrefunddebit = 0 "
	Else
		sWhere = ""
	End If 

	cTotal = 0.00

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, P.requirescheckno, P.requirescitizenaccount "
	sSql = sSql & "FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O " 
	sSql = sSql & "WHERE P.paymenttypeid = O.paymenttypeid AND O.orgid = " & Session("orgid") & sWhere
	sSql = sSql & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">"
			response.write oRs("paymenttypename") 
			response.write ": &nbsp;</td><td class=""amountcell"">"
			cAmount = GetAmount( iPaymentId, oRs("paymenttypeid") )
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
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">Total: &nbsp;</td>"
		response.write "<td class=""totalpayment"">" & FormatCurrency(cTotal,2) & "</td><td>&nbsp;</td><tr>"
		response.write vbcrlf & "</table>"
	End If 
	
	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRefundType( iPaymentId )
'--------------------------------------------------------------------------------------------------
Sub ShowRefundType( ByVal iPaymentId )
	Dim sSql, oRs, cTotal

	sSql = "SELECT accountid, amount, priorbalance, plusminus, itemid, ispaymentaccount, paymenttypeid, isccrefund "
	sSql = sSql & " FROM egov_accounts_ledger WHERE plusminus = '-' AND entrytype = 'credit' AND paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""10%"">"
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
				response.write "&nbsp;</td><td class=""refundamountcell"" nowrap=""nowrap"">&nbsp;</td><td> &nbsp;"
			End If 
		End If 
		response.write "</td></tr>"
		response.write vbcrlf & "</table>"
	Else
		' No payments were credited, So just show the table
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">"
		response.write "&nbsp;</td><td class=""refundamountcell"" nowrap=""nowrap"">&nbsp;</td><td> &nbsp;"
		response.write vbcrlf & "</table>"
		response.write "</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' double dAmount = GetAmount( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAmount( ByVal iPaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs, cAmount

	sSql = "SELECT amount FROM egov_accounts_ledger WHERE ispaymentaccount = 1 AND paymentid = " & iPaymentId
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

	GetAmount = cAmount

End Function 


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


'--------------------------------------------------------------------------------------------------
' string sAccount = GetAccountName( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAccountName( ByVal PaymentId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname FROM egov_verisign_payment_information, egov_users "
	sSql = sSql & " WHERE paymentid = " & iPaymentId & " AND paymenttypeid = " & iPaymentTypeId
	sSql = sSql & " AND citizenuserid = userid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetAccountName = oRs("userfname") & " " & oRs("userlname")
	Else
		GetAccountName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' ShowJournalDetails( iPaymentId, iUid, iDisplayType )
'--------------------------------------------------------------------------------------------------
Sub ShowJournalDetails( ByVal iPaymentId, ByVal iUid, ByVal iDisplayType )
	Dim sSql, oRs
	
	' Get the activities that they were part of - Should give 1 row
	sSql = "SELECT P.paymentid, P.paymentdate, L.entrytype, L.amount, U.firstname + ' ' + U.lastname AS adminname, P.notes "
	sSql = sSql & " FROM egov_class_payment P, egov_accounts_ledger L, egov_item_types I, users U "
	sSql = sSql & " WHERE P.paymentid = L.paymentid AND L.itemtypeid = I.itemtypeid AND "
	sSql = sSql & " P.adminuserid = U.userid AND L.accountid = " & iUid & " AND P.paymentid = " & iPaymentId
	'sSql = sSql & " order by paymentdate desc"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRS.EOF Then 
		Response.Write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"">"
		Response.Write vbcrlf & vbtab & "<tr><th align=""left"">Date</th><th align=""left"">Admin</th>"
		If iDisplayType = 2 Then 
			response.write "<th align=""left"">Transfered To</th>"
		End If 
		response.write "<th align=""left"">Notes</th><th align=""left"">Deposit</th><th align=""left"">Withdrawl</th></tr>"

		' LOOP AND DISPLAY THE RECORDS
		Do While Not oRs.EOF 
			response.Write vbcrlf & "<tr>"
			response.write "<td>" & oRs("paymentdate") & "</td>"
			response.write "<td>" & oRs("adminname") & "</td>"
			If iDisplayType = 2 Then 
				' show who got the transfer
				response.write "<td>" & GetTransferedTo( iPaymentId, iUid ) & "</td>"
			End If 
			response.write "<td>" & oRs("notes") & "</td>"
			response.write "<td>"
			If oRs("entrytype") = "credit" Then
				response.write FormatCurrency(oRs("amount"),2) 
			Else
				response.write "&nbsp;"
			End If
			response.write "</td>"
			response.write "<td>" 
			If oRs("entrytype") = "debit" Then
				response.write FormatCurrency(oRs("amount"),2) 
			Else 
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			response.Write vbcrlf & "</tr>"
			oRs.MoveNext
		Loop
		response.write "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


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


'--------------------------------------------------------------------------------------------------
' double dAmount = GetRefundDebit( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetRefundDebit( ByVal iPaymentId )
	Dim sSql, oRs

	' Pull a sum of what paid for prior class
	sSql = "SELECT SUM(amount) AS amount FROM egov_accounts_ledger "
	sSql = sSql & " WHERE ispaymentaccount = 0 AND entrytype = 'debit' AND paymentid = " & iPaymentId
	sSql = sSql & " GROUP BY paymentid "
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundDebit = CDbl(oRs("amount"))
	Else
		GetRefundDebit = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 
%>
