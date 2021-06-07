<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationrefund.asp
' AUTHOR: Steve Loar
' CREATED: 11/19/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Reservation refund processing.
'
' MODIFICATION HISTORY
' 1.0   11/19/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationId, sLoadMsg, iRentalUserId, sReservationType, sRenterName, sRenterPhone
Dim sReservationStatus, sReservedDate, sAdminName, iTotalDue, iRefundCount, sReservationTypeSelector

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "make reservations", sLevel	' In common.asp

iReservationId = CLng(request("reservationid"))
iRentalUserid = GetReservationRentalUserId( iReservationId )

GetGeneralReservationData iReservationId

iRefundCount = clng(0)

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--

		function loader()
		{
			setMaxLength();
			<%=sLoadMsg%>
		}

		function processRefund()
		{
			// Make sure that they have a payment amount more than $0
			if (Number($("refundtotal").value) < 0)
			{
				alert('We cannot complete this refund.\nThe refund total cannot be less than $0.00.');
				return;
			}

			document.frmReservationRefund.submit();
		}

		function ValidateCharges( oFee )
		{
			var bValid = true;

			// Remove any extra spaces
			oFee.value = removeSpaces(oFee.value);
			//Remove commas that would cause problems in validation
			oFee.value = removeCommas(oFee.value);

			// Validate the format of the charge
			if (oFee.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oFee.value);
				if ( Ok )
				{
					oFee.value = format_number(Number(oFee.value),2);
				}
				else 
				{
					oFee.value = format_number(0,2);
					bValid = false;
				}
			}
			else
			{
				oFee.value = "0.00";
			}

			if ( bValid == false ) 
			{
				//$("reservationok").value = 'false';
				//oFee.focus();
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>Amounts should be numbers in currency format.',8,oFee.id);
				oFee.focus();
				return false;
			}

			calculateRefundTotal();
			return true;
		}

		function calculateRefundTotal()
		{
			var TotalRefund = Number(0.00);
			// Add the daily rates
			for (var t = 1; t <= parseInt($("maxrentalrates").value); t++)
			{
				TotalRefund += Number($("datefeeamount" + t).value);
			}

			// Add the Item charges
			for (t = 1; t <= parseInt($("maxreservationitems").value); t++)
			{
				TotalRefund += Number($("itemfeeamount" + t).value);
			}

			// Add the reservation fees (deposits & alcohol)
			for (t = 1; t <= parseInt($("maxreservationfees").value); t++)
			{
				TotalRefund += Number($("reservationfeeamount" + t).value);
			}

			// store this off as RefundGrossAmount needed for the journal entry
			//$("grossrefundamount").value = TotalRefund;

			// subtract the refund fees
			for (t = 1; t <= parseInt($("maxrefundfees").value); t++)
			{
				TotalRefund -= Number($("refundfeeamount" + t).value);
			}

			// Update the total refund fields
			$("refundtotal").value = TotalRefund;
			$("refundtotaldisplay").innerHTML = format_number(TotalRefund,2);

		}

	//-->
	</script>

</head>

<body onload="loader();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Reservation Refund</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p>
				<span id="screenMsg">&nbsp;</span>
				&nbsp;
				<input type="button" class="button" value="<< Back To Reservation" onclick="location.href='reservationedit.asp?reservationid=<%=iReservationId%>';" />
			</p>

<%			ShowReservationInfoContainer sReservationType, sRenterName, sRenterPhone, sReservationStatus, sReservedDate, sAdminName, iReservationId, sReservationTypeSelector, 0		%>
			
			<form name="frmReservationRefund" action="reservationrefundmake.asp" method="post">
				<input type="hidden" id="reservationid" name="reservationid" value="<%=iReservationId%>" />

				<fieldset><legend><strong> Payment Summary </strong></legend>
<%					ShowPaymentLedgerDetails iReservationId %>
				</fieldset>

				<fieldset><legend><strong> Refund Details </strong></legend>
<%					ShowOverpaymentDetails iReservationId, iRefundCount	%>
				</fieldset>

				<p>
					<strong>Citizen Location:</strong> &nbsp; <% ShowPaymentLocations  ' In rentalsguifunctions.asp %>
				</p>

				<p>
					<strong>Apply the refund to:</strong> &nbsp; <% ShowRefundChoices iRentalUserid %>
				</p>

				<p>
					<table border="0" cellpadding="0" cellspacing="0" id="refundnotes">
					<tr><td valign="top" align="right" id="notestag"><strong>Notes:</strong> &nbsp; </td><td><textarea id="refundnotes" name="refundnotes" maxlength="500" class="purchasenotes"></textarea></td></tr>
					</table>
				</p>
<%				If iRefundCount > clng(0) Then		%>
					<p>
						<input class="button" type="button" name="complete" value="Process the Refund" onclick="processRefund();" />
					</p>
<%				End If		%>
			</form>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
'  ShowPaymentLedgerDetails iReservationId
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLedgerDetails( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT A.paymentid, P.paymenttypename, A.amount, V.checkno, A.accountid, P.ispublicmethod, P.isadminmethod, P.requirescheckno, P.requirescitizenaccount, P.requirescreditcard "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_paymenttypes P, egov_verisign_payment_information V "
	sSql = sSql & " WHERE A.paymentid = V.paymentid AND A.paymenttypeid = P.paymenttypeid "
	sSql = sSql & " AND A.ledgerid = V.ledgerid AND A.paymenttypeid = V.paymenttypeid AND A.ispaymentaccount = 1 "
	sSql = sSql & " AND A.entrytype= 'debit' AND A.reservationid = " & iReservationId & " ORDER BY A.paymentid, P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""refundpayments"" border=""0"" cellpadding=""2"" cellspacing=""0"">"
		response.write vbcrlf & "<tr><td class=""displayheaderpaid"">Receipt</td><td class=""displayheaderpaid"">Media</td><td class=""displayheaderpaid"">Amount</td><td>&nbsp;</td></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap"" valign=""top"">"
			response.write "&nbsp;" & oRs("paymentid") & "</td><td>"
			response.write oRs("paymenttypename") & "</td>"
			response.write "<td>" & FormatCurrency(oRs("amount")) & "</td><td> &nbsp; "
			' look up the account name if citizen account
			If oRs("requirescitizenaccount") Then
				response.write "From: " & GetCitizenName( oRs("accountid") )
			End If 
			If oRs("requirescreditcard") Then
				bCCUsed = True 
			End If 
			response.write "</td>"
			response.write "</tr>"
			cTotalPaid = cTotalPaid + CDbl(oRs("amount"))
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "<tr><td class=""displaytotalpaid"">&nbsp;</td><td class=""displaytotalpaid""><strong>Total Paid</strong></td><td class=""displaytotalpaid"">" & FormatCurrency(CDbl(cTotalPaid),2) 
		If bCCUsed Then
			response.write "<input type=""hidden"" name=""isccrefund"" value=""1"" />"
		Else 
			response.write "<input type=""hidden"" name=""isccrefund"" value=""0"" />"
		End If 
		response.write "</td><td>&nbsp;</td></tr>"
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
'  ShowRefundChoices iHeadUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowRefundChoices( ByVal iHeadUserId )
	Dim sSql, oRs

	response.write vbcrlf & "<select name=""accountid"">"
	response.write vbcrlf & "<option value=""0"" selected=""selected"">" & GetRefundName() & "</option>"  ' In common.asp

	If OrgHasFeature( "citizen accounts" ) Then 
		sSql = "SELECT userfname, userlname, userid, ISNULL(accountbalance,0.00) AS accountbalance "
		sSql = sSql & " FROM egov_users WHERE familyid = " & GetFamilyId( iHeadUserId )
		sSql = sSql & " ORDER BY userlname, userfname"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("userid") & """>" & oRs("userfname") & " " & oRs("userlname") & " (" & FormatNumber(oRs("accountbalance"),2) & ") " & "</option>"
			oRs.MoveNext
		Loop 

		oRs.Close
		Set oRs = Nothing 
	End If 

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowOverpaymentDetails iReservationId, iRefundCount
'--------------------------------------------------------------------------------------------------
Sub ShowOverpaymentDetails( ByVal iReservationId, ByRef iRefundCount )
	Dim sSql, oRs, iTotalDue, iRentalRateCount, iReservationItemCount, iReservationFeeCount, iRefundFeeCount

	iTotalDue = CDbl(0.0000)
	iRentalRateCount = clng(0)
	iReservationItemCount = clng(0)
	iReservationFeeCount = clng(0)
	iRefundFeeCount = clng(0)

	' Get the dates and times of the reservation
	sSql = "SELECT D.reservationdateid, D.reservationid, D.reservationstarttime, D.billingendtime, D.actualstarttime, "
	sSql = sSql & " D.actualendtime, D.reserveddate, D.adminuserid, D.rentalid, S.reservationstatus, S.iscancelled "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservationstatuses S "
	sSql = sSql & " WHERE D.statusid = S.reservationstatusid AND D.orgid = " & session("OrgId")
	sSql = sSql & " AND D.reservationid = " & iReservationId
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<table id=""refunddatesandfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"

	Do While Not oRs.EOF
		If ReservationDateHasOverpayments( oRs("reservationdateid") ) Then 
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
			response.write "</td>"
			
			' Location
			response.write "<td>"
			ShowRentalNameAndLocation oRs("rentalid") 
			response.write "</td>"
			response.write "</tr>"

			' Display any hourly rates that are over paid
			ShowRentalRateOverpayments oRs("reservationdateid"), iTotalDue, iRentalRateCount
			iRefundCount = iRefundCount + iRentalRateCount

			' Display items that are over paid
			ShowItemOverpayments oRs("reservationdateid"), iTotalDue, iReservationItemCount
			iRefundCount = iRefundCount + iReservationItemCount

		End If 
		oRs.MoveNext 
	Loop

	' Display Reservation Level charges like Deposit and Alcohol Fee that are over paid
	ShowReservationFeeOverpayments iReservationId, iTotalDue, iReservationFeeCount
	iRefundCount = iRefundCount + iReservationFeeCount

	If iRefundCount > clng(0) Then 
		' Show the refund payment types
		ShowRefundFeeFields iTotalDue, iRefundFeeCount
	End If 
	

	' Show the total refund line 
	response.write vbcrlf & "<tr><td class=""totalscell""><strong>Total Refund</strong></td><td class=""totalscell"">"
	'response.write "<input type=""hidden"" id=""grossrefundamount"" name=""grossrefundamount"" value=""" & iTotalDue & """ />"
	response.write "<input type=""hidden"" id=""refundtotal"" name=""refundtotal"" value=""" & iTotalDue & """ />"
	response.write "<span id=""refundtotaldisplay"">" & FormatNumber(iTotalDue,2,,,0) & "</span>"
	response.write "</td></tr>"

	response.write vbcrlf & "</table>"

	' Write out the maxcounts
	response.write "<input type=""hidden"" id=""maxrentalrates"" name=""maxrentalrates"" value=""" & iRentalRateCount & """ />"
	response.write "<input type=""hidden"" id=""maxreservationitems"" name=""maxreservationitems"" value=""" & iReservationItemCount & """ />"
	response.write "<input type=""hidden"" id=""maxreservationfees"" name=""maxreservationfees"" value=""" & iReservationFeeCount & """ />"
	response.write "<input type=""hidden"" id=""maxrefundfees"" name=""maxrefundfees"" value=""" & iRefundFeeCount & """ />"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRefundFeeFields iTotalDue, iRefundFeeCount
'--------------------------------------------------------------------------------------------------
Sub ShowRefundFeeFields( ByRef iTotalDue, ByRef iRefundFeeCount )
	Dim sSql, oRs

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, ISNULL(O.defaultamount,0.0000) AS defaultamount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE P.paymenttypeid = O.paymenttypeid AND P.isrefunddebit = 1 AND P.isforrentals = 1 "
	sSql = sSql & " AND O.orgid = " & session("orgid") & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<tr><td colspan=""3""><strong>Refund Fees & Damages</strong></td></tr>"
	Do While Not oRs.EOF
		iRefundFeeCount = iRefundFeeCount + 1

		response.write vbcrlf & "<tr>"
		response.write "<td>" & oRs("paymenttypename") & "</td>"
		response.write "<td colspan=""2"">"
'		If CDbl(oRs("defaultamount")) > CDbl(0.0000) Then 
'			dDueAmount = CDbl(oRs("defaultamount"))
'			iTotalDue = iTotalDue - dDueAmount
'		Else
			dDueAmount = "0.0000"
'		End If 
		response.write "<input type=""hidden"" name=""paymenttypeid" & iRefundFeeCount & """ value=""" & oRs("paymenttypeid") & """ />"
		response.write "&ndash; <input type=""text"" id=""refundfeeamount" & iRefundFeeCount & """ name=""refundfeeamount" & iRefundFeeCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
		response.write " onchange=""return ValidateCharges( this );"" />"
		If CDbl(oRs("defaultamount")) > CDbl(0.0000) Then 
			response.write "&nbsp;(" & FormatNumber(dDueAmount,2,,,0) & ")"
		End If 
		response.write "</td>"
		response.write vbcrlf & "</tr>"

		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRentalRateOverpayments iReservationDateId, iTotalDue, iRentalRateCount
'--------------------------------------------------------------------------------------------------
Sub ShowRentalRateOverpayments( ByVal iReservationDateId, ByRef iTotalDue, ByRef iRentalRateCount )
	Dim sSql, oRs, dDueAmount

	sSql = "SELECT F.reservationdatefeeid, F.reservationdateid, P.pricetypename, ISNULL(F.feeamount,0.00) AS feeamount, "
	sSql = sSql & " ISNULL(F.refundamount,0.00) AS refundamount, ISNULL(F.paidamount,0.00) AS paidamount, P.pricetypename "
	sSql = sSql & " FROM egov_rentalreservationdatefees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.reservationdateid = " & iReservationDateId
	sSql = sSql & " AND F.paidamount > (F.feeamount + F.refundamount) ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRentalRateCount = iRentalRateCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td>" & oRs("pricetypename") & "</td>"
		response.write "<td colspan=""2"">"
		dDueAmount = CDbl(oRs("paidamount")) - (CDbl(oRs("feeamount")) + CDbl(oRs("refundamount")))
		response.write "<input type=""hidden"" name=""reservationdatefeeid" & iRentalRateCount & """ value=""" & oRs("reservationdatefeeid") & """ />"
		response.write "+ <input type=""text"" id=""datefeeamount" & iRentalRateCount & """ name=""datefeeamount" & iRentalRateCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
		response.write " onchange=""return ValidateCharges( this );"" />"
		iTotalDue = iTotalDue + dDueAmount
		response.write "&nbsp;(" & FormatNumber(dDueAmount,2,,,0)
		response.write ")</td>"
		response.write vbcrlf & "</tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowItemOverpayments iReservationDateId, iTotalDue, iReservationItemCount
'--------------------------------------------------------------------------------------------------
Sub ShowItemOverpayments( ByVal iReservationDateId, ByRef iTotalDue, ByRef iReservationItemCount )
	Dim sSql, oRs, dDueAmount

	sSql = "SELECT reservationdateitemid, reservationdateid, rentalitem, ISNULL(feeamount,0.00) AS feeamount, "
	sSql = sSql & " ISNULL(refundamount,0.00) AS refundamount, ISNULL(paidamount,0.00) AS paidamount "
	sSql = sSql & " FROM egov_rentalreservationdateitems "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND paidamount > (feeamount + refundamount) ORDER BY rentalitem"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iReservationItemCount = iReservationItemCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td>" & oRs("rentalitem") & "</td>"
		response.write "<td colspan=""2"">"
		dDueAmount = CDbl(oRs("paidamount")) - (CDbl(oRs("feeamount")) + CDbl(oRs("refundamount")))
		response.write "<input type=""hidden"" name=""reservationdateitemid" & iReservationItemCount & """ value=""" & oRs("reservationdateitemid") & """ />"
		response.write "+ <input type=""text"" id=""itemfeeamount" & iReservationItemCount & """ name=""itemfeeamount" & iReservationItemCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
		response.write " onchange=""return ValidateCharges( this );"" />"
		iTotalDue = iTotalDue + dDueAmount
		response.write "&nbsp;(" & FormatNumber(dDueAmount,2,,,0)
		response.write ")</td>"
		response.write vbcrlf & "</tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowReservationFeeOverpayments iReservationId, iTotalDue, iReservationFeeCount
'--------------------------------------------------------------------------------------------------
Sub ShowReservationFeeOverpayments( ByVal iReservationId, ByRef iTotalDue, ByRef iReservationFeeCount )
	Dim sSql, oRs, dDueAmount

	sSql = "SELECT F.reservationfeeid, P.pricetypename, ISNULL(F.feeamount,0.00) AS feeamount, "
	sSql = sSql & " ISNULL(F.refundamount,0.00) AS refundamount, ISNULL(F.paidamount,0.00) AS paidamount, P.pricetypename "
	sSql = sSql & " FROM egov_rentalreservationfees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.reservationid = " & iReservationId
	sSql = sSql & " AND F.paidamount > (F.feeamount + F.refundamount) ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr><td colspan=""3""><strong>Reservation Fees</strong></td></tr>"
		Do While Not oRs.EOF
			iReservationFeeCount = iReservationFeeCount + 1
			response.write vbcrlf & "<tr>"
			response.write "<td>" & oRs("pricetypename") & "</td>"
			response.write "<td colspan=""2"">"
			dDueAmount = CDbl(oRs("paidamount")) - (CDbl(oRs("feeamount")) + CDbl(oRs("refundamount")))
			response.write "<input type=""hidden"" name=""reservationfeeid" & iReservationFeeCount & """ value=""" & oRs("reservationfeeid") & """ />"
			response.write "+ <input type=""text"" id=""reservationfeeamount" & iReservationFeeCount & """ name=""reservationfeeamount" & iReservationFeeCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
			response.write " onchange=""return ValidateCharges( this );"" />"
			iTotalDue = iTotalDue + dDueAmount
			response.write "&nbsp;(" & FormatNumber(dDueAmount,2,,,0)
			response.write ")</td>"
			response.write vbcrlf & "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 


End Sub 

'--------------------------------------------------------------------------------------------------
' boolean bHasOverpayments = ReservationDateHasOverpayments( iReservationDateId )
'--------------------------------------------------------------------------------------------------
Function ReservationDateHasOverpayments( iReservationDateId )
	
	If ReservationDateHasOverpaidFees( iReservationDateId ) Then
		ReservationDateHasOverpayments = True 
	Else
		If ReservationDateHasOverpaidItems( iReservationDateId ) Then
			ReservationDateHasOverpayments = True 
		Else
			ReservationDateHasOverpayments = False 
		End If 
	End If 

End Function


'--------------------------------------------------------------------------------------------------
' boolean bHasOverpayments = ReservationDateHasOverpaidFees( iReservationDateId )
'--------------------------------------------------------------------------------------------------
Function ReservationDateHasOverpaidFees( iReservationDateId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(reservationdatefeeid) AS hits FROM egov_rentalreservationdatefees "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND paidamount > (feeamount + refundamount)"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			ReservationDateHasOverpaidFees = True 
		Else
			ReservationDateHasOverpaidFees = False 
		End If 
	Else
		ReservationDateHasOverpaidFees = False 
	End If 

	oRs.Close 
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' boolean bHasOverpayments = ReservationDateHasOverpaidItems( iReservationDateId )
'--------------------------------------------------------------------------------------------------
Function ReservationDateHasOverpaidItems( iReservationDateId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(reservationdateitemid) AS hits FROM egov_rentalreservationdateitems "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND paidamount > (feeamount + refundamount)"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			ReservationDateHasOverpaidItems = True 
		Else
			ReservationDateHasOverpaidItems = False 
		End If 
	Else
		ReservationDateHasOverpaidItems = False 
	End If 

	oRs.Close 
	Set oRs = Nothing 
End Function 



%>
