<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: view_receipt.asp
' AUTHOR: Steve Loar
' CREATED: 04/05/07
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the receipt for a purchase, or refund 
'
' MODIFICATION HISTORY
' 1.0	04/06/07	Steve Loar - Initial Version
' 1.1	05/19/08 Steve Loar - PageDisplayCheck added
' 1.2	01/08/09	David Boyer - Added "DisplayRosterPublic" fields for Craig,CO custom team registration
' 1.3	03/09/09	Steve Loar - Changes for Regatta Teams
' 1.4 11/19/09 David Boyer - Added "pants size" to team registration section
' 1.5 11/19/09 David Boyer - Now pull team/pants sizes from database
' 1.6	04/07/2010	Steve Loar - No more regatta team members, added team group size
' 1.7 	11/5/2013	Steve Loar - Added the drop reason to the receipt
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iPaymentId, iUid, sReceiptType, sJEType, iDisplayType, sNotes, iAdminuserid, iPriorPaymentId
Dim dProcessingFee, bHasPaymentFee, lcl_total_label, lcl_total_display, iReturnTo, bShowUnDoBtn

response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button
 
sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "registration", sLevel	' In common.asp

iPaymentId = CLng(request("iPaymentId"))

If request("return") <> "" Then
	iReturnTo = CLng(request("return"))
Else
	iReturnTo = CLng(0)
End If 

bShowUnDoBtn = false 
If IsUnDoBtnDisplayed( iPaymentId ) Then 
	bShowUnDoBtn = true 
	SetUnDoBtnDisplay iPaymentId, false 
Else
	' see if the person has rights to do undo on any purchase - may only want to do this on purchases from the admin side, not drops, or public side purchases '
	If UserHasPermission( Session("UserId"), "undo class purchase" ) Then
		' Only let them undo if it is an admin side purchase and is not a related payment. That would indicate the person has dropped this already. '
		' Deleting these would mess up the reports on the later transaction. Feaving you a refund for something that was never purchased.'
		'response.write "IsAdminPurchase = " & IsAdminPurchase( iPaymentId ) & "<br />"
		'Response.write "IsRelatedPayment = " & IsRelatedPayment( iPaymentId ) & "<br />"
		If IsAdminPurchase( iPaymentId ) And Not IsRelatedPayment( iPaymentId ) Then 
			bShowUnDoBtn = true
		End If 
	End If 
End If 

'Check for org features
lcl_orghasfeature_citizen_accounts = orgHasFeature("citizen accounts")
lcl_orghasfeature_custom_registration_craigco = orgHasFeature("custom_registration_CraigCO")

%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	
 	<title>E-Gov Administration Console {Receipt}</title>

 	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 	<link rel="stylesheet" href="../global.css" />
 	<link rel="stylesheet" href="classes.css" />
 	<link rel="stylesheet" href="receiptprint.css" media="print" />

	<style type="text/css">
		#content1       { border: 1px solid blue;}
		#centercontent1 { border: 1px solid red; }
		#topright1      { background-color: yellow;	}
	</style>
	
	<script src="https://code.jquery.com/jquery-1.5.min.js"></script>
	
	<script>
	
		var unDoTransaction = function( paymentId ) {
			var okToProceed = confirm("This will permanently remove this tranaction from the system. Do you wish to continue?");
			
			if ( okToProceed ) {
				'alert("Firing off undo script on " + paymentId );'
				var request = jQuery.ajax({  
					url: "./undo_purchase.asp",  
					type: "POST",  
					dataType: "text",
					data: { 
						paymentId : paymentId
				 	 },  
					contentType: 'application/x-www-form-urlencoded; charset=UTF-8'
				}); 

				request.done( function( data ) { 
					'alert(data);'
					if (data === 'Success') {
						alert("This transaction has successfully been removed. \nPrint this receipt if you need to. Once you leave this page, all related information will be gone.");
						$("#undobutton").hide();
					}
					else {
						alert("Failed: This transaction was not successfully removed.");
					}
				});
				
				request.fail( function(jqXHR, textStatus) { 
					alert( "Failed: " + textStatus );
				});
		
			}
		};
	
	</script>

</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
<%
	If iReturnTo = CLng(1) Then 
		response.write "&nbsp;&nbsp;<input type=""button"" class=""button"" id=""scanagainbtn"" value=""Scan Another Card"" onclick=""location.href='dropinregistration.asp';"" />"
	End If 
	
	If bShowUnDoBtn Then 
		response.write "&nbsp;&nbsp;<input type=""button"" class=""button"" id=""undobutton"" value=""Undo This Transaction"" onclick=""unDoTransaction(" & iPaymentId & ");"" />"
	End If 
%>
</div>

<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">
<%
	ShowReceiptHeader iPaymentId

	response.write vbcrlf & "<hr />"

	Dim iuserid, nTotal, dPaymentDate, sSql, nRowTotal, bMultiWeeks, iAdminLocationId, iJournalEntryTypeId, sJournalEntryType, bIsCCRefund, sDropReason
	iPriorPaymentId = ""

	If GetPaymentDetails(iPaymentId, iuserid, nTotal, dPaymentDate, iAdminLocationId, iJournalEntryTypeId, sNotes, iAdminuserid, _
                      iPriorPaymentId, iRosterGrade, iRosterShirtSize, iRosterPantsSize, iRosterCoachType, iRosterVolunteerCoachName, _
                      iRosterVolunteerCoachDayPhone, iRosterVolunteerCoachCellPhone, iRosterVolunteerCoachEmail, sDropReason) Then 

		sJournalEntryType = GetJournalEntryType( iJournalEntryTypeId )

		If sJournalEntryType = "purchase" Then
			' See if the gateway for this org has fees they charge the citizen
			If PaymentGatewayRequiresFeeCheck( session("orgid") ) Then
				bHasPaymentFee = True 
				dProcessingFee = GetProcessingFee( iPaymentId )
			Else
				bHasPaymentFee = False 
				dProcessingFee = CDbl("0.00")
			End If 
		Else
			bHasPaymentFee = False 
			dProcessingFee = CDbl("0.00")
		End If 

		response.write "<span id=""receiptadmininfo"">"
		response.write " Location: " & GetAdminLocation( iAdminLocationId ) & "&nbsp;&nbsp;"  'In ../includes/common.asp
		response.write " Administrator: " & GetAdminName( iAdminuserid )  'In ../includes/common.asp
		If iPriorPaymentId <> "" Then 
			response.write "&nbsp;&nbsp;Prior Receipt: <a href=""view_receipt.asp?iPaymentId=" & iPriorPaymentId & """>" & iPriorPaymentId & "</a>"
		End If 
		response.write "</span>"

		response.write " Date: " & DateValue(CDate(dPaymentDate)) & "&nbsp;&nbsp;"
		response.write " Receipt: " & iPaymentId & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "

		response.write "<hr />"
  		response.write "<div id=""receipttopright"">"

		If sJournalEntryType = "refund" Then 
			'sRefundAmount = GetPurchaseTotal( iPaymentId ) - GetRefundFeeAmount( iPaymentId )
		   lcl_total_label   = "Amount Refunded:"
		   lcl_total_display = GetPurchaseTotal( iPaymentId ) - GetRefundFeeAmount( iPaymentId )
		Else 
		   lcl_total_label   = "Transaction Total:"
		   lcl_total_display = nTotal
		End If 

		response.write "<p id=""transactiontotal"">"
		response.write "<strong>" & lcl_total_label & "</strong> " & FormatCurrency((CDbl(lcl_total_display) + CDbl(dProcessingFee)),2)
		response.write "</p>"
			
		If lcl_orghasfeature_citizen_accounts Then 
		   ShowAccountChange iPaymentId, iuserid
		End If 

  		response.write "</div>"
  		response.write ShowUserInfo( iuserid, sJournalEntryType )
		response.write "<hr />"

  		If sJournalEntryType = "refund" Then 
			ShowRefundType iPaymentId
  		Else 
		    ShowPaymentTypes iPaymentId, sJournalEntryType, bHasPaymentFee, dProcessingFee
  		End If 

  		response.write "<hr />"
  		response.write "<strong>Transactions</strong>"
  		response.write "<hr />"
		'response.write "[" & sJournalEntryType & "]"

  		Select Case sJournalEntryType
			Case "purchase"
				'Show purchase details
				ShowPurchaseDetails iPaymentId, "credit", sJournalEntryType, 0, bHasPaymentFee, dProcessingFee

			Case "refund"
				'Show refund stuff
				ShowPurchaseDetails iPaymentId, "debit", sJournalEntryType, iPriorPaymentId, bHasPaymentFee, dProcessingFee

			Case "transfer"
				'Show citizen account transfer

			Case "deposit"
				'Show citizen account deposit

			Case "withdrawl"
			'Show citizen account withdrawl

		End Select 

	 Else 
			response.write "<p>No Details could be found for the requested receipt.</p>"
	 End If 

	response.write "<hr />"

	ShowReceiptFooter sNotes, sJournalEntryType, sDropReason

%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' void ShowPurchaseDetails iPaymentId, sEntryType, sJournalEntryType, iPriorPaymentId, bHasPaymentFee, dProcessingFee
'--------------------------------------------------------------------------------------------------
Sub ShowPurchaseDetails( ByVal iPaymentId, ByVal sEntryType, ByVal sJournalEntryType, ByVal iPriorPaymentId, ByVal bHasPaymentFee, ByVal dProcessingFee )
	Dim sSql, oRs, cTotal, cAmount

	cTotal = CDbl(0.00)

	' Pull a set of items purchased
	sSql = "SELECT T.cartdisplayorder, itemtype, itemid, L.itemtypeid, SUM(amount) AS amount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_item_types T "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 0 AND entrytype = '" & sEntryType & "' "
	sSql = sSql & " AND L.paymentid = " & iPaymentId
	sSql = sSql & " GROUP BY T.cartdisplayorder, itemtype, itemid, L.itemtypeid "
	sSql = sSql & " ORDER BY T.cartdisplayorder, itemtype, itemid, L.itemtypeid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		' Change to a case statement as you add more things that can be bought; like gifts and lodges or memberships
		Select Case oRs("itemtype")
			Case "recreation activity"
				If sJournalEntryType = "refund" Then
					cAmount = GetPriorPurchaseTotal( iPriorPaymentId, oRs("itemtypeid"), oRs("itemid") )
				Else 
					cAmount = CDbl(oRs("amount"))
				End If 
				cTotal = cTotal + cAmount
				ShowActivityDetails oRs("itemid"), cAmount, iPaymentid, oRs("itemtypeid")

			Case "regatta team"
				cAmount = CDbl(oRs("amount"))
				cTotal = cTotal + cAmount
				ShowTeamDetails oRs("itemid"), cAmount

			Case "regatta member"
				cAmount = CDbl(oRs("amount"))
				cTotal = cTotal + cAmount
				ShowTeamDetails oRs("itemid"), cAmount

			Case "merchandise"
				cAmount = CDbl(oRs("amount"))
				cTotal = cTotal + cAmount
				response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptdetails"">"
				response.write vbcrlf & "<tr><td class=""merchandiselabelcolumn""><strong>Merchandise Purchased</strong></td>"
				response.write "<td align=""right"">" & FormatCurrency(cAmount,2) & "</td></tr>"
				response.write vbcrlf & "<tr><td class=""merchandiselabelcolumn"" valign=""top"">"
				ShowMerchandiseItems oRs("itemid")
				response.write "</td>"
				response.write "<td valign=""top"" class=""receiptshippinglabel"">" 
				ShowShippingLabel oRs("itemid")
				response.write "</td></tr>"
				response.write vbcrlf & "</table>"

			Case "shipping and handling fees"
				cAmount = CDbl(oRs("amount"))
				cTotal = cTotal + cAmount
				response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptdetails"">"
				response.write vbcrlf & "<tr><td class=""merchandiselabelcolumn""><strong>Shipping and Handling Fees</strong></td>"
				response.write "<td align=""right"">" & FormatCurrency(cAmount,2) & "</td></tr>"
				response.write vbcrlf & "</table>"

			Case "sales tax"
				cAmount = CDbl(oRs("amount"))
				cTotal = cTotal + cAmount
				response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptdetails"">"
				response.write vbcrlf & "<tr><td class=""merchandiselabelcolumn""><strong>Sales Tax</strong></td>"
				response.write "<td align=""right"">" & FormatCurrency(cAmount,2) & "</td></tr>"
				response.write vbcrlf & "</table>"

		End Select 
		oRs.MoveNext
	Loop 

	oRs.Close 
	Set oRs = Nothing

	If sJournalEntryType = "refund" Then
		cTotal = cTotal - ShowRefundFee( iPaymentId, cTotal )
		response.write "<div id=""receiptdetailtotal"">"
		response.write "<strong>Refund Total:</strong> "
	Else
		If bHasPaymentFee Then
			' Add the Processing Fee to the total charged
			cTotal = cTotal + dProcessingFee

			' Show the processing fee
			response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptdetails"">"
			response.write vbcrlf & "<tr><td class=""merchandiselabelcolumn""><strong>Processing Fee</strong></td>"
			response.write "<td align=""right"">" & FormatCurrency(dProcessingFee,2) & "</td></tr>"
			response.write vbcrlf & "</table>"
		End If 
		response.write "<div id=""receiptdetailtotal"">"
		response.write "<strong>Total:</strong> "	
	End If 
	response.write FormatCurrency(cTotal,2) & "</div>"
End Sub 


'--------------------------------------------------------------------------------------------------
' double GetPriorPurchaseTotal( iPriorPaymentId, iItemTypeId, iItemId )
'--------------------------------------------------------------------------------------------------
Function GetPriorPurchaseTotal( ByVal iPriorPaymentId, ByVal iItemTypeId, ByVal iItemId )
	Dim sSql, oRs

	' Pull a sum of what paid for prior class
	sSql = "SELECT SUM(amount) AS amount FROM egov_accounts_ledger "
	sSql = sSql & " WHERE ispaymentaccount = 0 AND entrytype = 'credit' AND paymentid = " & iPriorPaymentId
	sSql = sSql & " AND itemtypeid = " & iItemTypeId & " AND itemid = " & iItemId
	sSql = sSql & " GROUP BY itemtypeid, itemid "
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPriorPurchaseTotal = CDbl(oRs("amount"))
	Else
		GetPriorPurchaseTotal = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetRefundDebit( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetRefundDebit( ByVal iPaymentId )
	Dim sSql, oRs

	' Pull a sum of what paid for prior class
	sSql = "SELECT SUM(amount) AS amount FROM egov_accounts_ledger "
	sSql = sSql & " WHERE ispaymentaccount = 0 AND entrytype = 'debit' AND paymentid = " & iPaymentId
	sSql = sSql & " GROUP BY paymentid "
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRefundDebit = CDbl(oRs("amount"))
	Else
		GetRefundDebit = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double ShowRefundFee( iPaymentId, cTotal )
'--------------------------------------------------------------------------------------------------
Function ShowRefundFee( ByVal iPaymentId, ByVal cTotal )
	Dim sSql, oRs, cRefundShortage

	cRefundShortage = CDbl(0.00) 

	cRefundShortage = cTotal - GetRefundDebit( iPaymentId )
	
	' Pull a the refund fee row
	sSql = "SELECT itemtype, itemid, amount FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 AND entrytype = 'credit' "
	sSql = sSql & " AND P.isrefunddebit = 1 AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table class=""receiptdetails"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		response.write vbcrlf & "<tr><td class=""firstcell""><strong>Refund Fees</strong></td>"
		'response.write "<td align=""right""><strong>Fees:</strong> -" & FormatCurrency(oRs("amount"),2) & "</td></tr>"
		response.write "<td align=""right"">-" & FormatCurrency((oRs("amount") + cRefundShortage),2) & "</td></tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<p class=""receiptnotes"">&nbsp;</p>"
		ShowRefundFee = CDbl(oRs("amount") + cRefundShortage)
	Else
		ShowRefundFee = CDbl(cRefundShortage)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetVoucherAmount( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetVoucherAmount( ByVal iPaymentId )
	Dim sSql, oRs
	
	' Pull a the refund voucher row
	sSql = "SELECT itemtype, itemid, amount FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 AND entrytype = 'credit' "
	sSql = sSql & " AND P.isrefunddebit = 0 AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetVoucherAmount = CDbl(oRs("amount"))
	Else
		GetVoucherAmount = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetRefundFeeAmount( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetRefundFeeAmount( ByVal iPaymentId )
	Dim sSql, oRs
	
	' Pull a the refund fee row
	sSql = "SELECT amount FROM egov_accounts_ledger L, egov_item_types T, egov_paymenttypes P "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 1 AND entrytype = 'credit' "
	sSql = sSql & " AND P.isrefunddebit = 1 AND P.paymenttypeid = L.paymenttypeid AND L.paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetRefundFeeAmount = CDbl(oRs("amount"))
	Else
		GetRefundFeeAmount = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double GetPurchaseTotal( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetPurchaseTotal( ByVal iPaymentId )
	Dim sSql, oRs
	
	' Pull a the purchase total sum
	sSql = "SELECT SUM(amount) AS amount FROM egov_accounts_ledger L, egov_item_types T "
	sSql = sSql & " WHERE L.itemtypeid = T.itemtypeid AND L.ispaymentaccount = 0 AND entrytype = 'debit' "
	sSql = sSql & " AND L.paymentid = " & iPaymentId & " GROUP BY L.paymentid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPurchaseTotal = CDbl(oRs("amount"))
	Else
		GetPurchaseTotal = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void  ShowTeamDetails iItemId, cAmount 
'--------------------------------------------------------------------------------------------------
Sub ShowTeamDetails( ByVal iItemId, ByVal cAmount )
	Dim sSql, oRs

	sSql = "SELECT C.classname, C.startdate, C.notes, T.regattateam, T.captainfirstname, T.captainlastname, L.quantity, "
	sSql = sSql & " T.captainaddress, T.captaincity, T.captainstate, T.captainzip, T.captainphone, T.regattateamgroupid "
	sSql = sSql & " FROM egov_regattateams T, egov_class_list L, egov_class C "
	sSql = sSql & " WHERE T.regattateamid = L.regattateamid AND C.classid = L.classid "
	sSql = sSql & " AND L.classlistid = " & iItemId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptdetails"">"
		response.write vbcrlf & "<tr><td class=""firstcell""><strong>Activity: </strong> " & oRs("classname") & "</td>"
		response.write "<td><strong>Event Date:</strong> " & oRs("startdate") & "</td>"
		response.write "<td align=""right"">" & FormatCurrency(cAmount,2) & "</td></tr>"

		response.write vbcrlf & "<tr><td class=""firstcell"" valign=""top""><strong>Team Name: </strong> " & oRs("regattateam") & "</td>"
		response.write "<td>&nbsp;</td><td>&nbsp;</td></tr>"

		response.write vbcrlf & "<tr><td class=""teamnamecell"" valign=""top"" colspan=""2""><strong>Team Group: </strong> " & GetTeamGroupName( oRs("regattateamgroupid") ) & "</td>"
		response.write "<td>&nbsp;</td></tr>"

		response.write vbcrlf & "<tr><td class=""firstcell"" valign=""top""><strong>Captain: </strong> <br />"
		response.write vbcrlf & "<div class=""captaindetails"">" & oRs("captainfirstname") & " " & oRs("captainlastname") & "<br />"
		response.write oRs("captainaddress") & "<br />"
		response.write oRs("captaincity") & ", " & oRs("captainstate") & " " & oRs("captainzip") & "<br />"
		response.write FormatPhoneNumber(oRs("captainphone"))  & "</div>"
		response.write "</td>"
		'response.write "<td valign=""top""><strong>Members Added: </strong> " & oRs("quantity") & "</td>"
		response.write "<td valign=""top"">&nbsp;</td>"
		response.write "<td>&nbsp;</td></tr>"
		response.write vbcrlf & "</table>"
		If oRs("notes") <> "" Then 
			response.write vbcrlf & "<p class=""teamactivitynotes""><strong>Activity Notes: </strong>" & oRs("notes") & "</p>"
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowActivityDetails iItemId, dSubTotal, iPaymentid, iItemTypeid 
'--------------------------------------------------------------------------------------------------
Sub ShowActivityDetails( ByVal iItemId, ByVal dSubTotal, ByVal iPaymentid, ByVal iItemTypeid )
	Dim sSql, oRs 

	'response.write iItemId & "Put some Recreation Activity Purchase details here " & FormatCurrency(dSubTotal,2) & "<br />"

	'Need attendee name, class name, activity no, startdate, enddate, days, times, notes, location, isdropin, dropindate, quantity 
	sSql = "SELECT "
	sSql = sSql & " U.userfname, "
	sSql = sSql & " U.userlname, "
	sSql = sSql & " status, "
	sSql = sSql & " quantity, "
	sSql = sSql & " isdropin, "
	sSql = sSql & " dropindate, "
	sSql = sSql & " C.classid, "
	sSql = sSql & " classname, "
	sSql = sSql & " C.isparent, "
	sSql = sSql & " CL.name, "
	sSql = sSql & " CL.address1, "
	sSql = sSql & " C.startdate, "
	sSql = sSql & " C.enddate, "
	sSql = sSql & " notes, "
	sSql = sSql & " T.activityno, "
	sSql = sSql & " sunday, "
	sSql = sSql & " monday, "
	sSql = sSql & " tuesday, "
	sSql = sSql & " wednesday, "
	sSql = sSql & " thursday, "
	sSql = sSql & " friday, "
	sSql = sSql & " saturday, "
	sSql = sSql & " TD.starttime, "
	sSql = sSql & " TD.endtime, "
	sSql = sSql & " L.rostergrade, "
	sSql = sSql & " L.rostershirtsize, "
	sSql = sSql & " L.rosterpantssize, "
	sSql = sSql & " L.rostercoachtype, "
	sSql = sSql & " L.rostervolunteercoachname, "
	sSql = sSql & " rostervolunteercoachdayphone, "
	sSql = sSql & " rostervolunteercoachcellphone, "
	sSql = sSql & " rostervolunteercoachemail "
	sSql = sSql & " FROM egov_class_list L, "
	sSql = sSql &      " egov_class C, "
	sSql = sSql &      " egov_class_time T, "
	sSql = sSql &      " egov_class_time_days TD, "
	sSql = sSql &      " egov_class_location CL, "
	sSql = sSql &      " egov_users U "
	sSql = sSql & " WHERE C.classid = L.classid "
	sSql = sSql & " AND C.classid = T.classid "
	sSql = sSql & " AND L.classtimeid = T.timeid "
	sSql = sSql & " AND L.attendeeuserid = U.userid "
	sSql = sSql & " AND T.timeid = TD.timeid "
	sSql = sSql & " AND C.locationid = CL.locationid "
	sSql = sSql & " AND L.classlistid = " & iItemId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptdetails"">"
		response.write "  <tr>"
		response.write "      <td class=""firstcell""><strong>Activity:</strong> " & oRs("classname") & " (" & oRs("activityno") & ")</td>"
		response.write "      <td><strong>Location:</strong> " & oRs("name") & " &ndash; " & oRs("address1") & "</td>"
		'response.write "      <td align=""right""><strong>Price:</strong> " & FormatCurrency(dSubTotal,2) & "</td></tr>"
		response.write "      <td align=""right"">" & FormatCurrency(dSubTotal,2) & "</td>"
		response.write "  </tr>"
		response.write "  <tr>"
		response.write "      <td class=""firstcell"" valign=""top"">"
		response.write "          <strong>Attendee:</strong> " & oRs("userfname") & " " & oRs("userlname") & " &mdash; " & GetJournalItemStatus( iPaymentid, iItemTypeid, iItemId )
		response.write "          <br /><strong>Qty:</strong> " & oRs("quantity")
		response.write "      </td>"
		response.write "      <td class=""times"" valign=""top"">"

  		If Not oRs("isdropin") Then 
    		response.write "<strong>From: </strong>" & oRs("startdate") & "&nbsp;"
			response.write "<strong>To: </strong>"   & oRs("enddate")   & "<br />"

			i = 0
    		Do While Not oRs.EOF
				i = i + 1

				If i > 1 Then 
					response.write "<br />"
				End If 

       			response.write "<strong>Days:</strong> "

				If oRs("sunday") Then 
					response.write "Su "
				End If 
				If oRs("monday") Then 
					response.write "Mo "
				End If 
				If oRs("tuesday") Then 
					response.write "Tu "
				End If 
				If oRs("wednesday") Then 
					response.write "We "
				End If 
				If oRs("thursday") Then 
					response.write "Th "
				End If 
				If oRs("friday") Then 
					response.write "Fr "
				End If 
				If oRs("saturday") Then 
					response.write "Sa "
				End If 

      			response.write "&ndash; <strong>Times:</strong> " & oRs("starttime") & " to " & oRs("endtime")

      			oRs.MoveNext
    		Loop 
  		Else 
    		response.write "<strong>Drop In Date: </strong>" & oRs("dropindate")
    End If 

	oRs.MoveLast

	response.write "      </td>"
    response.write "      <td></td>"
    response.write "  </tr>"

    If oRs("isparent") Then 
		response.write "  <tr><td class=""firstcell"" valign=""top"" colspan=""3""><strong>Series Includes:</strong></td></tr>"

  		'Show Child activities here  - Need name, location, days, times
    	ShowSeriesChildren oRs("classid")
    End If 
	response.write "</table>"

   'Determine if org is using custom roster info and display it if any exists
    if lcl_orghasfeature_custom_registration_craigco then
      'Check for an "edit display" for the T-shirt label
       if orgHasDisplay(session("orgid"),"class_teamregistration_tshirt_label") then
          lcl_label_tshirt = getOrgDisplay(session("orgid"),"class_teamregistration_tshirt_label")
       else
          lcl_label_tshirt = "T-Shirt"
       end if

       response.write "<p class=""rosterdetails"">"

       if oRs("rostergrade") <> "" then
          response.write "<strong>Grade: </strong>" & oRs("rostergrade") & "<br />"
       end if

       if trim(oRs("rostershirtsize")) <> "" and trim(oRs("rostershirtsize")) <> "," then
          response.write "<strong>" & lcl_label_tshirt & " Size: </strong>" & oRs("rostershirtsize") & "<br />"
       end if

       if trim(oRs("rosterpantssize")) <> "" AND trim(oRs("rosterpantssize")) <> "," then
          response.write "<strong>Pants Size: </strong>" & oRs("rosterpantssize") & "<br />"
       end if

       if oRs("rostercoachtype") <> "" then
          response.write "<strong>Knows someone or would like to be a volunteer: </strong>" & oRs("rostercoachtype") & "<br />"
       end if

       if oRs("rostervolunteercoachname") <> "" then
          response.write "<strong>Coach Name: </strong>" & oRs("rostervolunteercoachname") & "<br />"
       end if

       if oRs("rostervolunteercoachdayphone") <> "" then
          response.write "<strong>Day Phone: </strong>" & formatphonenumber(oRs("rostervolunteercoachdayphone")) & "<br />"
       end if

       if oRs("rostervolunteercoachcellphone") <> "" then
          response.write "<strong>Cell Phone: </strong>" & formatphonenumber(oRs("rostervolunteercoachcellphone")) & "<br />"
       end if

       if oRs("rostervolunteercoachemail") <> "" then
          response.write "<strong>Email: </strong>" & oRs("rostervolunteercoachemail") & "<br />"
       end if

       response.write "</p>"

    end if

	If oRs("notes") <> "" Then 
		response.write "<p class=""receiptnotes""><strong>Activity Notes: </strong>" & oRs("notes") & "</p>"
	End If 

    response.write "<br />"

 end if
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowSeriesChildren iParentClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowSeriesChildren( ByVal iParentClassId )
	Dim sSql, oRs

	sSql = "SELECT  C.classid, classname, CL.name, CL.address1, C.startdate, C.enddate, notes, T.activityno, "
	sSql = sSql & " sunday, monday, tuesday, wednesday, thursday, friday, saturday, TD.starttime, TD.endtime "
	sSql = sSql & " FROM egov_class C, egov_class_time T, egov_class_time_days TD, egov_class_location CL "
	sSql = sSql & " WHERE C.classid = T.classid AND T.timeid = TD.timeid AND C.locationid = CL.locationid "
	sSql = sSql & " AND C.parentclassid = " & iParentClassId & " ORDER BY C.classid, T.activityno, T.timeid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write "<tr><td class=""firstcell"" valign=""top"">" & oRs("classname") & " (" & oRs("activityno") & ")</td>"
		response.write "<td class=""times""><strong>Location:</strong> " & oRs("name") & " &ndash; " & oRs("address1") & "<br />"
		response.write "<strong>From: </strong>" & oRs("startdate") & " &nbsp; <strong>To: </strong>" & oRs("enddate") & "<br />"
			
		response.write "<strong>Days:</strong> " 
		If oRs("sunday") Then
			response.write "Su "
		End If 
		If oRs("monday") Then
			response.write "Mo "
		End If 
		If oRs("tuesday") Then
			response.write "Tu "
		End If 
		If oRs("wednesday") Then
			response.write "We "
		End If 
		If oRs("thursday") Then
			response.write "Th "
		End If 
		If oRs("friday") Then
			response.write "Fr "
		End If 
		If oRs("saturday") Then
			response.write "Sa "
		End If 
		response.write "&ndash; <strong>Times:</strong> " & oRs("starttime") & " to " & oRs("endtime") & "<br />"
		response.write "</td><td>&nbsp;</td></tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetJournalEntryType( iJournalEntryTypeId )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryType( ByVal iJournalEntryTypeId )
	Dim sSql, oRs

 If iJournalEntryTypeId <> "" Then 
   	sSql = "SELECT journalentrytype FROM egov_journal_entry_types WHERE journalentrytypeid = " & iJournalEntryTypeId

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
' void ShowReceiptHeader iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptHeader( ByVal iPaymentId )

	If OrgHasDisplay( Session("orgid"), "receipt header" ) Then
		response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("orgid"), "receipt header" ) 
		response.write "<br /><br />" & GetReceiptHeader( iPaymentId )
		response.write "</p>"
	Else  
		response.write "<h3>" & Session("sOrgName") & " " & GetReceiptHeader( iPaymentId ) & "</h3><br /><br />"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetReceiptHeader( iPaymentId )
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
		GetReceiptHeader = iPaymentId
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowReceiptFooter sNotes, sJournalEntryType, sDropReason
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptFooter( ByVal sNotes, ByVal sJournalEntryType, ByVal sDropReason )
	Dim sFooter

	If sDropReason <> "" Then
		response.write "<p><strong>Reason: " & sDropReason & "</strong></p>"
	End If

	response.write "<p><strong>Receipt Notes:</strong> <span id=""receiptnotes"">" & Trim(sNotes)  & "</span></p>"

	If sJournalEntryType = "refund" Then 
  		sFooter = "refund footer"
	Else 
  		sFooter = "receipt footer"
	End If 

	If OrgHasDisplay( session("orgid"), sFooter ) Then  
  		response.write "<p>" & GetOrgDisplay( session("orgid"), sFooter ) & "</p>"
	End If 

End sub  


'--------------------------------------------------------------------------------------------------
' boolean GetPaymentDetails(iPaymentId, ByRef iuserid, ByRef nTotal, ByRef dPaymentDate, ByRef iAdminLocationId, iJournalEntryTypeId, _
'                           ByRef sNotes, ByRef iAdminuserid, ByRef iPriorPaymentId, ByRef iRosterGrade, ByRef iRosterShirtSize, _
'                           ByRef iRosterCoachType, ByRef iRosterVolunteerCoachName, ByRef iRosterVolunteerCoachDayPhone, _
'                           ByRef iRosterVolunteerCoachCellPhone, ByRef iRosterVolunteerCoachEmail)
'--------------------------------------------------------------------------------------------------
function GetPaymentDetails(iPaymentId, ByRef iuserid, ByRef nTotal, ByRef dPaymentDate, ByRef iAdminLocationId, iJournalEntryTypeId, _
                           ByRef sNotes, ByRef iAdminuserid, ByRef iPriorPaymentId, ByRef iRosterGrade, ByRef iRosterShirtSize, _
                           ByRef iRosterPantsSize, ByRef iRosterCoachType, ByRef iRosterVolunteerCoachName, _
                           ByRef iRosterVolunteerCoachDayPhone, ByRef iRosterVolunteerCoachCellPhone, ByRef iRosterVolunteerCoachEmail, ByRef sDropReason)
	Dim sSql, oRs

	sSql = "SELECT userid, paymenttotal, paymentdate, ISNULL(adminlocationid,0) AS adminlocationid, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, journalentrytypeid, notes, relatedpaymentid, ISNULL(dropreasonid,0) AS dropreasonid "
	sSql = sSql & " FROM egov_class_payment "
	sSql = sSql & " WHERE paymentid = " & iPaymentId & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		iuserid             = oRs("userid")
		nTotal              = oRs("paymenttotal")
		dPaymentDate        = oRs("paymentdate")
		iAdminLocationId    = oRs("adminlocationid")
		iJournalEntryTypeId = oRs("journalentrytypeid")
		sNotes              = oRs("notes")
		iAdminuserid        = oRs("adminuserid")
		iPriorPaymentId     = oRs("relatedpaymentid")
		iDropReasonId 		= oRs("dropreasonid")
		sDropReason			= GetDropReason( iDropReasonId )
		GetPaymentDetails   = True 
	Else 
  		GetPaymentDetails   = False 
  		sDropReason = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

end function


'--------------------------------------------------------------------------------------------------
' string GetDropReason( iDropReasonId )
'--------------------------------------------------------------------------------------------------
Function GetDropReason( ByVal iDropReasonId )
	Dim sSql, oRs

	sSql = "SELECT dropreason FROM egov_class_dropreasons WHERE dropreasonid = " & iDropReasonId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDropReason = oRs("dropreason")
	Else
		GetDropReason = ""
	End If

	oRs.Close
	Set oRs = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' Sub ShowAccountChange( iPaymentId, iUid )
'--------------------------------------------------------------------------------------------------
Sub ShowAccountChange( iPaymentId, iUid )
	Dim sSql, oRs, cAmount, cPriorBalance, cCurrentBalance, sEntryType, cPrefix

	' Get the activities that they were part of - Should give 1 or 0 rows
	sSql = "SELECT entrytype, amount, priorbalance, plusminus "
	sSql = sSql & " FROM egov_accounts_ledger "
	sSql = sSql & " WHERE accountid = " & iUid & " AND paymentid = " & iPaymentId

	response.write "<!--" & sSql & "-->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

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


	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowUserInfo( iuserid, sJournalEntryType )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( iuserid, sJournalEntryType )
	Dim oCmd, sResidentDesc, sUserType
	ShowUserInfo = ""

	sUserType = GetUserResidentType(iuserid)
	' If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" And sUserType <> "N" Then
		' This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iuserid, Session("orgid"))
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iuserid", 3, 1, 4, iuserid)
	    Set oUser = .Execute
	End With
	
	ShowUserInfo = ShowUserInfo & "<span class=""receipttitles"">"
	If sJournalEntryType = "refund" Then
		ShowUserInfo = ShowUserInfo & "Refundee "
	Else 
		ShowUserInfo = ShowUserInfo & "Payee "
	End If 
	ShowUserInfo = ShowUserInfo & "Information</span><br />"
	ShowUserInfo = ShowUserInfo & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & oUser("userfname") & " " & oUser("userlname") & "</strong><br />"
	ShowUserInfo = ShowUserInfo & "<strong>" & oUser("useraddress") 
	If oUser("userunit") <> "" Then 
		ShowUserInfo = ShowUserInfo & "&nbsp;&nbsp;" & oUser("userunit") 
	End If
	If oUser("useraddress2") <> "" Then 
		ShowUserInfo = ShowUserInfo & "<br />" & oUser("useraddress2") 
	End If 
	ShowUserInfo = ShowUserInfo & "<br />" & oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & "</strong></td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td colspan=""2"">&nbsp;</td></tr>"
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & GetFamilyEmail( iuserid ) & "</td></tr>"
 	if  Session("orgid") = "81" then
		ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Account&nbsp;#:</td><td>" & oUser("FamilyID") & "</td></tr>"
	end if
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhone(oUser("userhomephone")) & "</td></tr>"
	'ShowUserInfo = ShowUserInfo & "<tr><td width=""85"" align=""right"" valign=""top"">Business:</td><td>" & oUser("userbusinessname") & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "</table>"

	oUser.close
	Set oUser = Nothing
	Set oCmd = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPaymentTypes iPaymentId, sJournalEntryType, bHasPaymentFee, dProcessingFee
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentTypes( ByVal iPaymentId, ByVal sJournalEntryType, ByVal bHasPaymentFee, ByVal dProcessingFee )
	Dim sSql, oRs, cTotal, sWhere

	If sJournalEntryType <> "refund" Then
		sWhere = " and isrefundmethod = 0 and isrefunddebit = 0 "
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
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">Total: &nbsp;</td><td class=""totalpayment"">" & FormatCurrency(cTotal,2) & "</td><td>&nbsp;</td><tr>"
		response.write vbcrlf & "</table>"
	End If 
	
	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowRefundType( iPaymentId )
'--------------------------------------------------------------------------------------------------
Sub ShowRefundType( ByVal iPaymentId )
	Dim sSql, oRefund, cTotal

	sSql = "SELECT accountid, amount, priorbalance, plusminus, itemid, ispaymentaccount, paymenttypeid, isccrefund "
	sSql = sSql & " FROM egov_accounts_ledger WHERE entrytype = 'credit' AND paymentid = " & iPaymentId

	Set oRefund = Server.CreateObject("ADODB.Recordset")
	oRefund.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRefund.EOF Then 
		response.write vbcrlf & "<table id=""receiptpayments"" border=""0"" cellspacing=""2"" cellpadding=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" width=""15%"">"
		If oRefund("ispaymentaccount") Then 
			If oRefund("isccrefund") Then 
				response.write "Refund to Credit Card"
			Else 
				' This is a refund voucher
				response.write GetRefundName() 
			End If 
			response.write ": &nbsp;</td><td class=""amountcell"" nowrap=""nowrap"">"
			response.write FormatCurrency(oRefund("amount"), 2)
		Else
			If CDbl(oRefund("amount")) > CDbl(0.00) Then 
				' This is to a citizen account
				response.write "Citizen Account: &nbsp;</td><td class=""amountcell"" nowrap=""nowrap"">"
				response.write FormatCurrency(oRefund("amount"), 2) & "</td>" 
				response.write "<td> &nbsp; To: &nbsp; " & GetCitizenName( oRefund("accountid") )
			Else
				response.write "Removed From Waitlist &nbsp;</td><td class=""amountcell"" nowrap=""nowrap"">&nbsp;</td><td> &nbsp;"
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
		response.write "Removed From Waitlist &nbsp;</td><td class=""amountcell"" nowrap=""nowrap"">&nbsp;</td><td> &nbsp;"
		response.write vbcrlf & "</table>"
		response.write "</td></tr>"

	End If 

	oRefund.Close
	Set oRefund = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetAmount( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAmount( iPaymentId, iPaymentTypeId )
	Dim sSql, oAmount, cAmount

	sSql = "Select amount From egov_accounts_ledger where ispaymentaccount = 1 and paymentid = " & iPaymentId & " and paymenttypeid = " & iPaymentTypeId

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 0, 1

	If Not oAmount.EOF Then 
		cAmount = CDbl(oAmount("amount"))
	Else
		cAmount = CDbl(0.00)
	End If 

	oAmount.close
	Set oAmount = Nothing 

	GetAmount = cAmount

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetCheckNo( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetCheckNo( iPaymentId, iPaymentTypeId )
	Dim sSql, oAmount

	sSql = "Select checkno From egov_verisign_payment_information where paymentid = " & iPaymentId & " and paymenttypeid = " & iPaymentTypeId

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 0, 1

	If Not oAmount.EOF Then 
		GetCheckNo = oAmount("checkno")
	Else
		GetCheckNo = ""
	End If 

	oAmount.close
	Set oAmount = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAccountName( iPaymentId, iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetAccountName( iPaymentId, iPaymentTypeId )
	Dim sSql, oName

	sSql = "Select userfname, userlname From egov_verisign_payment_information, egov_users "
	sSql = sSql & " where paymentid = " & iPaymentId & " and paymenttypeid = " & iPaymentTypeId
	sSql = sSql & " and citizenuserid = userid "

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		GetAccountName = oName("userfname") & " " & oName("userlname")
	Else
		GetAccountName = ""
	End If 

	oName.close
	Set oName = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetTransferedTo( iPaymentId, iUid )
'--------------------------------------------------------------------------------------------------
Function GetTransferedTo( iPaymentId, iUid )
	Dim sSql, oName

	sSql = "Select userfname, userlname From egov_accounts_ledger, egov_users "
	sSql = sSql & " where accountid = userid and paymentid = " & iPaymentId & " and accountid <> " & iUid

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 0, 1

	If Not oName.EOF Then 
		GetTransferedTo = oName("userfname") & " " & oName("userlname")
	Else
		GetTransferedTo = ""
	End If 

	oName.close
	Set oName = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowJournalDetails( iPaymentId, iUid, iDisplayType )
'--------------------------------------------------------------------------------------------------
Sub ShowJournalDetails( iPaymentId, iUid, iDisplayType )
	Dim sSql, oRs
	
	' Get the activities that they were part of - Should give 1 row
	sSql = "select P.paymentid, P.paymentdate, L.entrytype, L.amount, U.firstname + ' ' + U.lastname as adminname, P.notes "
	sSql = sSql & " from egov_class_payment P, egov_accounts_ledger L, egov_item_types I, users U "
	sSql = sSql & " where P.paymentid = L.paymentid and L.itemtypeid = I.itemtypeid and "
	sSql = sSql & " P.adminuserid = U.userid and L.accountid = " & iUid & " and P.paymentid = " & iPaymentId
	'sSql = sSql & " order by paymentdate desc"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

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
	oRs.close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetFamilyMemberInfo( iFamilyMemberId )
'--------------------------------------------------------------------------------------------------
Function GetFamilyMemberInfo( iFamilyMemberId )
	Dim sSql, oName
	
	If iFamilyMemberId <> 0 Then 
		sSql = "Select firstname, lastname, birthdate From egov_familymembers Where familymemberid = " & iFamilyMemberId

		Set oName = Server.CreateObject("ADODB.Recordset")
		oName.Open sSQL, Application("DSN"), 1, 3

		GetFamilyMemberInfo = oName("firstname") & " " & oName("lastname") 
		If Not IsNull(oName("birthdate")) Then 
			GetFamilyMemberInfo = GetFamilyMemberInfo & " " & DateDiff("yyyy", oName("birthdate"), Now()) & " Years Old"
		End If 

		oName.close
		Set oName = Nothing
	End If 
	
End Function  


'--------------------------------------------------------------------------------------------------
' Function CheckIfSeries( iClassId, iPaymentId )
'--------------------------------------------------------------------------------------------------
Function CheckIfSeries( iClassId, iPaymentId )
	Dim sSql, oClass

	sSql = "Select count(classlistid) as hits From egov_class_list Where paymentid = " & iPaymentId & " and classid = " & iClassId

	Set oClass = Server.CreateObject("ADODB.Recordset")
	oClass.Open sSQL, Application("DSN"), 0, 1

	If clng(oClass("hits")) > 0 Then 
		CheckIfSeries = True
	Else 
		CheckIfSeries = False 
	End If 

	oClass.close
	Set oClass = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ListWaivers(iorgid,classid,blnshowname,blnshowdescription,blnshowlink)
'--------------------------------------------------------------------------------------------------
Sub ListWaivers(iorgid,classid,blnshowname,blnshowdescription,blnshowlink)

	sSQL = "Select * from egov_class_waivers INNER JOIN egov_class_to_waivers ON egov_class_waivers.waiverid=egov_class_to_waivers.waiverid where orgid = '" & iorgid & "' AND classid='" & classid & "' order by waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), 3, 1
	
	' LIST ALL WAIVER FOR ORGANIZATION
	If Not oWaiver.EOF Then
	
		response.write "<tr><td colspan=4>"

		' WAIVER TITLE
		response.write "<div class=waivertitle>This class requires the following waivers:"
		'response.write "<br><br><A href='http://www.adobe.com/products/acrobat/readstep2.html' target='_blank' title='Get Adobe Acrobat Reader Plug-in Here'><img border=0 src=""../images/adreader.gif"" hspace=10>Get Adobe Reader.</a>"
		response.write "</div>"

		Do While Not oWaiver.EOF
		
			response.write "<div class=waiver>" 

			' WAIVER NAME
			If blnshowname Then
				response.write "<div class=waivername>" & oWaiver("waivername") & "</div>"
			End If

			' WAIVER DESCRIPTION
			If blnshowdescription Then
				response.write "<div class=waiverdescription>" & oWaiver("waiverdescription") & "</div>"
			End If

			' WAIVER LINK
			If blnshowlink Then
				response.write "<div class=waiverlink>&bull; <a  href=""" & oWaiver("waiverurl") & """ target=""_NEW"" class=waiverlink>Click here to download " & Ucase(oWaiver("waivername")) & " waiver.</a></div>"
			End If
	
			response.write "</div>"

			oWaiver.MoveNext
		
		Loop 

		response.write "</td></tr>"

	End If 
	
	' CLOSE AND CLEAR OBJECTS
	oWaiver.close
	Set oWaiver = nothing

End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowClassDetails()
'--------------------------------------------------------------------------------------------------
Sub ShowClassDetails()
	Dim sSql, oRs 

	' Get the classes for this purchase order by start date
	sSql = "Select L.classid, C.classname, C.startdate, C.enddate, C.isparent, C.classtypeid, isnull(C.parentclassid,0) as parentclassid, "
	sSql = sSql & " C.locationid, P.name as location, isnull(P.address1,' ') as address1, isnull(P.address2,' ') as address2, "
	sSql = sSql & " T.starttime, T.endtime, T.waitlistsize, L.status, L.quantity, L.amount, isnull(L.familymemberid,0) as familymemberid, L.classtimeid "
	sSql = sSql & " From egov_class_list L, egov_class C, egov_class_time T, egov_class_location P "
	sSql = sSql & " where C.classid = L.classid and L.classtimeid = T.timeid and P.locationid = C.locationid and L.paymentid = " & iPaymentId 
	sSql = sSql & " Order By C.startdate, L.classid, L.familymemberid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<div class=""shadow"">"
		response.write vbcrlf & "<table border=""1"" cellpadding=""3"" cellspacing=""0"">"
		response.write vbcrlf & "<tr>"
		response.write "<th>Class/Event/Program</th><th>Qty</th><th>Price</th><th>Total</th>"
		response.write "</tr>"
		' Loop through the classes display details
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			'show details here
			response.write "<td><h5>" & oRs("classname") & "</h5>"
			If oRs("isparent") And oRs("classtypeid") = 1 Then 
				' This is the series level
				response.write " &ndash; Entire Series"
			End If 
			If clng(oRs("familymemberid")) <> 0 Then 
				response.write "<br /> &nbsp;  &nbsp; <strong>Attendee:</strong> " & GetFamilyMemberInfo( oRs("familymemberid") )
			End If 

			' Show the class/event dates
			response.write "<br /><br /> &nbsp;  &nbsp; <strong>Occurs:</strong><br />"
			If oRs("startdate") <> "" Then 
				response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & MonthName(Month(oRs("startdate"))) & " " & Day(oRs("startdate"))
			End If 
			' handle enddate
			bMultiWeeks = false
			If oRs("enddate") <> "" And oRs("enddate") <> oRs("startdate") Then 
				response.write " &ndash; " & MonthName(Month(oRs("enddate"))) & " " & Day(oRs("enddate"))
				If DateDiff("d", oRs("startdate"), oRs("enddate")) > 7 Then 
					bMultiWeeks = True
				Else 
					bMultiWeeks = false
				End If 
			End If 

			' Days of the week
			ShowDaysOfWeek oRs("classid"), bMultiWeeks
			' Time of the class/event
			response.write "<br /> &nbsp;  &nbsp; &nbsp; &nbsp;" & oRs("starttime") 
			If oRs("endtime") <> oRs("starttime") Then
				response.write " &ndash; " & oRs("endtime")
			End If


			response.write "<br /><br /> &nbsp;  &nbsp; <strong>Location:</strong><br />" 
			response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & oRs("location") & "<br />"
			If Trim(oRs("address1")) <> "" Then 
				response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & oRs("address1") & "<br />"
			End If 
			If Trim(oRs("address2")) <> "" Then 
				response.write " &nbsp;  &nbsp; &nbsp; &nbsp;" & oRs("address2") & "<br />"
			End If 

			response.write "</td>"
			response.write "<td align=""center"" valign=""top"">" & oRs("quantity") & "</td>"
			If oRs("status") = "ACTIVE" Then 
				If Not IsNull(oRs("amount")) Then 
					response.write "<td align=""right"" valign=""top"">"
					response.write FormatCurrency(oRs("amount")) & "</td>"
					'nRowTotal = clng(oRs("quantity")) * CDbl(oRs("amount"))
					nRowTotal = CDbl(oRs("amount"))
					response.write "<td align=""right"" valign=""top"">" & FormatCurrency(nRowTotal) & "</td>"
				Else
					' No price, see if they are part of a purchased series
					If CheckIfSeries( oRs("parentclassid"), iPaymentId ) Then
						response.write "<td colspan=""2"" align=""center"" valign=""top"">Part of Series<br />"
					Else 
						response.write "<td align=""right"" valign=""top"">&nbsp;</td><td align=""right"" valign=""top"">&nbsp;</td>"
					End If 
				End If 
			Else
				response.write "<td colspan=""2"" align=""center"" valign=""top"">"
				If CheckIfSeries( oRs("parentclassid"), iPaymentId ) Then
					response.write "Part of Series<br />"
				End If 
				response.write "On Wait List<br />"
				response.write GetWaitPosition(oRs("classid"), iuserid, oRs("familymemberid")) & " of " & oRs("waitlistsize")
				response.write "</td>"
			End If 

			' DISPLAY WAIVER LINK
			ListWaivers SESSION("orgid"),oRs("classid"),0,0,1

			oRs.movenext
			response.write "</tr>"

		Loop

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"
		
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' boolean PurchaseContainsRegattaTeams( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function PurchaseContainsRegattaTeams( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT classlistid FROM egov_class_list "
	sSql = sSql & " WHERE paymentid = " & iPaymentId
	sSql = sSql & " AND regattateamid IS NOT NULL"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		PurchaseContainsRegattaTeams = True 
	Else
		PurchaseContainsRegattaTeams = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void DisplayRegattaTeamPurchases iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub DisplayRegattaTeamPurchases( iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT classlistid, regattateamid FROM egov_class_list "
	sSql = sSql & " WHERE paymentid = " & iPaymentId
	sSql = sSql & " AND regattateamid IS NOT NULL"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' Page break here
		response.write vbcrlf & "<p class=""receiptteamdisplay"">"
		response.write vbcrlf & "<hr class=""teambreak"" />"
		ShowTeamInformation oRs("regattateamid")
		ShowTeamMembers oRs("regattateamid"), oRs("classlistid")
		response.write vbcrlf & "</p>"
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowTeamMembers iRegattaTeamId 
'--------------------------------------------------------------------------------------------------
Sub ShowTeamMembers( ByVal iRegattaTeamId, ByVal iClassListId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT regattateammember, isteamcaptain FROM egov_regattateammembers "
	sSql = sSql & " WHERE regattateamid = " & iRegattaTeamId & " AND classlistid = " & iClassListId
	sSql = sSql & " AND orgid = " & session("orgid") & " ORDER BY regattateammember"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write  vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""regattateamlist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		response.write vbcrlf & "<tr><th>Team Members Added In This Purchase</th></tr>"
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
   			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
   			End If 
			response.write "><td>" & oRs("regattateammember")
			If oRs("isteamcaptain") Then
				response.write " &nbsp; (Captain)"
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 
	Else
		response.write "<p>No members could be found for this team.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowTeamInformation iRegattaTeamId 
'--------------------------------------------------------------------------------------------------
Sub ShowTeamInformation( ByVal iRegattaTeamId )
	Dim sSql, oRs
	
	sSql = "SELECT regattateam, captainfirstname, captainlastname, captainaddress, captaincity, captainstate, captainzip, captainphone "
	sSql = sSql & " FROM egov_regattateams WHERE orgid = " & session("orgid") & " AND regattateamid = " & iRegattaTeamId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<p>"
		response.write vbcrlf & "<font size=""+1""><strong>" & oRs("regattateam") & "</strong></font><br />"
		response.write vbcrlf & "</p>"

		response.write vbcrlf & "<table id=""captaindata"" cellpadding=""3"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr>"
		response.write vbcrlf & "<td valign=""top"" id=""captainlabel""><strong>Captain:</strong><td>"
		response.write vbcrlf & "<td>"
		response.write vbcrlf & oRs("captainfirstname") & " " & oRs("captainlastname") & "<br />"
		response.write vbcrlf & oRs("captainaddress") & "<br />"
		response.write vbcrlf & oRs("captaincity") & ", " & oRs("captainstate") & "&nbsp;" & oRs("captainzip") & "<br />"
		response.write vbcrlf & formatphonenumber(oRs("captainphone"))
		response.write vbcrlf & "</td>"
		response.write vbcrlf & "</tr>"
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowShippingLabel( iMerchandiseOrderId )
'--------------------------------------------------------------------------------------------------
Sub ShowShippingLabel( ByVal iMerchandiseOrderId )
	Dim sSql, oRs
	
	sSql = "SELECT shiptoname, shiptoaddress, shiptocity, shiptostate, shiptozip "
	sSql = sSql & " FROM egov_merchandiseorders WHERE merchandiseorderid = " & iMerchandiseOrderId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "<strong>Ship To:</strong><br />"
		response.write oRs("shiptoname") & "<br />"
		response.write oRs("shiptoaddress") & "<br />"
		response.write oRs("shiptocity") &", " & oRs("shiptostate") & " " & oRs("shiptozip")
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowMerchandiseItems( iMerchandiseOrderId )
'--------------------------------------------------------------------------------------------------
Sub ShowMerchandiseItems( ByVal iMerchandiseOrderId )
	Dim sSql, oRs
	
	sSql = "SELECT merchandise, merchandisecolor, merchandisesize, quantity, isnocolor, isnosize, itemprice "
	sSql = sSql & " FROM egov_merchandiseorderitems WHERE merchandiseorderid = " & iMerchandiseOrderId
	sSql = sSql & " AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY merchandise, merchandisecolor, displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""receiptmerchandiseitems"">"
		response.write vbcrlf & "<tr class=""receiptmerchandiseitemheader""><th align=""left"">Item</th>"
		response.write "<th align=""center"">Color</th><th align=""center"">Size</th><th align=""center"">Price<br />Each</th>"
		response.write "<th align=""center"">Qty</th><th align=""center"">Item<br />Total</th></tr>"

		Do While Not oRs.EOF
			response.write "<tr>"
			response.write "<td>" & oRs("merchandise") & "</td>"
			response.write "<td align=""center"">"
			If oRs("isnocolor") Then
				response.write "&nbsp;"
			Else 
				response.write oRs("merchandisecolor")
			End If 
			response.write "</td>"
			response.write "<td align=""center"">"
			If oRs("isnosize") Then
				response.write "&nbsp;"
			Else 
				response.write oRs("merchandisesize")
			End If 
			response.write "</td>"
			response.write "<td align=""center"">" & FormatNumber(oRs("itemprice"),2) & "</td>"
			response.write "<td align=""center"">" & oRs("quantity") & "</td>"
			dItemTotal = oRs("itemprice") * oRs("quantity")
			response.write "<td align=""center"">" & FormatNumber(dItemTotal,2) & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
		response.write "</table>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub



%>
