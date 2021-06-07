<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationpayment.asp
' AUTHOR: Steve Loar
' CREATED: 11/11/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Reservation payment processing.
'
' MODIFICATION HISTORY
' 1.0   11/11/2009	Steve Loar - INITIAL VERSION
' 1.1	12/31/2009	Steve Loar - Changes to handle $0 cost rentals for Montgomery
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationId, sLoadMsg, iRentalUserId, sReservationType, sRenterName, sRenterPhone
Dim sReservationStatus, sReservedDate, sAdminName, iTotalDue, sReservationTypeSelector

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "make reservations", sLevel	' In common.asp

iReservationId = CLng(request("reservationid"))
iRentalUserid  = GetReservationRentalUserId( iReservationId )

GetGeneralReservationData iReservationId
%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script src="../scripts/modules.js"></script>
	<script src="../scripts/textareamaxlength.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/removespaces.js"></script>
	<script src="../scripts/removecommas.js"></script>
	<script src="../scripts/setfocus.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>
	<script src="../scripts/ajaxLib.js"></script>

	<script>
	<!--

		function loader()
		{
			setMaxLength();
			<%=sLoadMsg%>
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
				inlineMsg(oFee.id,'<strong>Invalid Value: </strong>Charges should be numbers in currency format.',8,oFee.id);
				oFee.focus();
				return false;
			}

			calculateTotalCharges();
			return true;
		}

		function calculateTotalCharges()
		{
			var TotalCharges = Number(0.00);
			// Add the daily rates
			for (var t = 1; t <= parseInt($("maxrentalrates").value); t++)
			{
				TotalCharges += Number($("datefeeamount" + t).value);
			}

			// Add the Item charges
			for (t = 1; t <= parseInt($("maxreservationitems").value); t++)
			{
				TotalCharges += Number($("itemfeeamount" + t).value);
			}

			// Add the reservation fees
			for (t = 1; t <= parseInt($("maxreservationfees").value); t++)
			{
				TotalCharges += Number($("reservationfeeamount" + t).value);
			}

			// Update the total charges
			$("chargestotal").value = TotalCharges;
			$("chargestotaldisplay").innerHTML = format_number(TotalCharges,2);

			// Update the balance due
			var PaymentTotal = Number($("paymenttotal").value);
			var BalanceDue = TotalCharges - PaymentTotal;
			$("balancedue").value = BalanceDue;
			$("balanceduedisplay").innerHTML = format_number(BalanceDue,2);
		}

		function validatePayment( oPayment )
		{
			var bValid = true;

			// Remove any extra spaces
			oPayment.value = removeSpaces(oPayment.value);
			//Remove commas that would cause problems in validation
			oPayment.value = removeCommas(oPayment.value);

			// Validate the format of the charge
			if (oPayment.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPayment.value);
				if ( Ok )
				{
					oPayment.value = format_number(Number(oPayment.value),2);
				}
				else 
				{
					oPayment.value = '';
					bValid = false;
				}
			}
			else
			{
				oPayment.value = '';
			}

			if ( bValid == false ) 
			{
				//$("reservationok").value = 'false';
				//oPayment.focus();
				inlineMsg(oPayment.id,'<strong>Invalid Value: </strong>Payment amounts should be numbers in currency format.',8,oPayment.id);
				oPayment.focus();
				return false;
			}

			calculateTotalPayments();
			return true;
		}

		function calculateTotalPayments()
		{
			var TotalPayments = Number(0.00);
			// Add the daily rates
			console.log(TotalPayments);
			for (var t = 1; t <= parseInt($("maxpayments").value); t++)
			{
				if ($("paymentamount" + t).value != '')
				{
					//TotalPayments += Number($("paymentamount" + t).value);
					var thisPayment = (Number($("paymentamount" + t).value) * 100).toFixed(0);
					var currTotal = (Number(TotalPayments) * 100).toFixed(0);
					TotalPayments = ((Number(thisPayment) + Number(currTotal))/100).toFixed(2);
			console.log(TotalPayments);
				}
			}
			console.log("");
			// Update the total Payments
			$("paymenttotal").value = TotalPayments;
			$("paymenttotaldisplay").innerHTML = format_number(TotalPayments,2);

			// Update the balance due
			var TotalCharges = Number($("chargestotal").value);
			console.log(TotalCharges);
			console.log(TotalPayments);
			var BalanceDue = (((Number(TotalCharges)*100) - (Number(TotalPayments)*100)).toFixed(0))/100;
			console.log(BalanceDue);
			$("balancedue").value = BalanceDue;
			$("balanceduedisplay").innerHTML = format_number(BalanceDue,2);

			$("chargestotal").value = TotalCharges;
			$("chargestotaldisplay").innerHTML = format_number(TotalCharges,2);
			var PaymentTotal = Number($("paymenttotal").value);
			
			console.log("");
		}

		function validate()
		{
			// Make sure that any check payments have a check number. It is optional to have this.
			for (var t = 1; t <= parseInt($("maxpayments").value); t++)
			{
				if ($("paymentamount"+t).value != "" && $("hascheckno"+t).value == "yes")
				{
					if ($("checkno"+t).value == '')
					{
						if ( ! confirm('You have a check payment without a check number. \nDo you wish to continue?'))
						{
							$("checkno"+t).focus();
							return;
						}
					}
				}
			}

			// Make sure that they have a payment amount more than $0
			if (Number($("paymenttotal").value) == 0)
			{
				
				if ( Number($("balancedue").value) == 0 ) 
				{
					if ( ! confirm('There are no charges or payment amounts entered.\nDo you wish to continue?'))
					{
						return;
					}
					var bOkToProcess = false;
					// Make sure that at least one payment amount has something in it.
					for (t = 1; t <= parseInt($("maxpayments").value); t++)
					{
						if ( $("paymentamount"+t).value != "" )
						{
							bOkToProcess = true;
						}
					}
					if (bOkToProcess == false)
					{
						alert("To process a $0 payment, please enter 0.00 in the 'other' payment type, then try to complete this again.");
						return;
					}
				}
				else 
				{
					alert('We cannot complete this payment.\nThe payment total cannot be $0.00 when there is a balance due.');
					return;
				}
			}

			// Make sure the balance due is $0
			if (Number($("balancedue").value) == 0)
			{
				//alert('OK to Complete');
				document.frmReservationPayment.submit();
			}
			else
			{
				alert('We cannot complete this payment.\nThe payment total does not equal the total charges.');
				return;
			}
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
				<font size="+1"><strong>Reservation Payment</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p>
				<span id="screenMsg">&nbsp;</span>
				&nbsp;
				<input type="button" class="button" value="<< Back To Reservation" onclick="location.href='reservationedit.asp?reservationid=<%=iReservationId%>';" />
			</p>

<%			ShowReservationInfoContainer sReservationType, sRenterName, sRenterPhone, sReservationStatus, sReservedDate, sAdminName, iReservationId, sReservationTypeSelector, 0		%>
			
			<form name="frmReservationPayment" action="reservationpaymentmake.asp" method="post">
				<input type="hidden" id="reservationid" name="reservationid" value="<%=iReservationId%>" />

				<!-- Charges -->
<%				ShowDatesAndCharges iReservationId, iTotalDue, sReservationTypeSelector		%>

				<!-- Payment -->
				<fieldset><legend><strong>Payment&nbsp;</strong></legend><br />
					<input type="hidden" value="0.00" name="amount" />

<%					ShowPaymentChoices iRentalUserId, iTotalDue, sReservationTypeSelector		%>

					<br /><br />
					<input type="button" class="button" name="complete" value="Complete Payment" onClick="validate()" />
				</fieldset>
			</form>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' ShowDatesAndCharges iReservationId, iTotalDue, sReservationTypeSelector
'--------------------------------------------------------------------------------------------------
Sub ShowDatesAndCharges( ByVal iReservationId, ByRef iTotalDue, ByVal sReservationTypeSelector )
	Dim sSql, oRs, iRowCount, sStartHour, sStartMinute, sStartAmPm, sEndHour, sEndMinute, sEndAmPm
	Dim sAmPm, iDateTotal, iReservationFeeCount, iRentalRateCount, iReservationItemCount
	Dim iReservationDateCount, dTotalRefunded, dTotalPaid, bNoCostToRent

	iRowCount = 0
	iDateTotal = CDbl(0.00)
	iReservationFeeCount = clng(0)
	iRentalRateCount = clng(0)
	iReservationItemCount = clng(0)
	iReservationDateCount = clng(0)

	' Get the reserved dates
	sSql = "SELECT D.reservationdateid, D.reservationid, D.reservationstarttime, D.billingendtime, D.actualstarttime, "
	sSql = sSql & " D.actualendtime, D.reserveddate, D.adminuserid, D.rentalid, S.reservationstatus, S.iscancelled, R.nocosttorent "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservationstatuses S, egov_rentals R "
	sSql = sSql & " WHERE D.statusid = S.reservationstatusid AND D.rentalid = R.rentalid AND D.orgid = " & session("OrgId")
	sSql = sSql & " AND S.iscancelled = 0 AND D.reservationid = " & iReservationId
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<table id=""datesandfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"

	Do While Not oRs.EOF
		' Display the date info for all
		iRowCount = iRowCount + 1
		iReservationDateCount = iReservationDateCount + 1
		iDateTotal = CDbl(0.00)

		If sReservationTypeSelector = "admin" Then 
			bNoCostToRent = True 
		Else
			If oRs("nocosttorent") Then
				bNoCostToRent = True 
			Else
				bNoCostToRent = False 
			End If 
		End If 
		
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & ">"

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
		ShowRentalNameAndLocation oRs("rentalid") 
		response.write "</td>"

		response.write "<td>&nbsp;</td>"

		response.write "</tr>"

		' Display the hourly rates
		ShowRentalRateChargesForDate oRs("reservationdateid"), sClass, iTotalDue, iRentalRateCount, bNoCostToRent

		' Display items 
		ShowItemsChargesForDate oRs("reservationdateid"), sClass, iTotalDue, iReservationItemCount, bNoCostToRent

		oRs.MoveNext
	Loop

	' Reservation fees
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 
	' Display Reservation Level charges like Deposit and Alcohol Fee
	ShowReservationFeeCharges iReservationId, sClass, iTotalDue, iReservationFeeCount

	' Total Charges Row
	iRowCount = iRowCount + 1
	If iRowCount Mod 2 = 0 Then
		sClass = " class=""altrow"" "
	Else
		sClass = ""
	End If 
	response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""3""><strong>Charges Total</strong></td><td class=""totalscell"" align=""right"">"
	response.write "<input type=""hidden"" id=""chargestotal"" name=""chargestotal"" value=""" & iTotalDue & """ />"
	response.write "<span id=""chargestotaldisplay"">" & FormatNumber(iTotalDue,2,,,0) & "</span>"
	response.write "</td></tr>"

	response.write vbcrlf & "</table>"

	' Write out the maxcounts
	response.write "<input type=""hidden"" id=""maxreservationfees"" name=""maxreservationfees"" value=""" & iReservationFeeCount & """ />"
	response.write "<input type=""hidden"" id=""maxrentalrates"" name=""maxrentalrates"" value=""" & iRentalRateCount & """ />"
	response.write "<input type=""hidden"" id=""maxreservationitems"" name=""maxreservationitems"" value=""" & iReservationItemCount & """ />"
	response.write "<input type=""hidden"" id=""maxreservationdates"" name=""maxreservationdates"" value=""" & iReservationDateCount & """ />"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRentalRateChargesForDate iReservationDateId, sClass, iTotalDue, iRentalRateCount, bNoCostToRent
'--------------------------------------------------------------------------------------------------
Sub ShowRentalRateChargesForDate( ByVal iReservationDateId, ByVal sClass, ByRef iTotalDue, ByRef iRentalRateCount, ByVal bNoCostToRent )
	Dim sSql, oRs, dDueAmount, bOkToCharge

	sSql = "SELECT F.reservationdatefeeid, ISNULL(F.feeamount,0.00) AS feeamount, ISNULL(paidamount,0.00) AS paidamount, "
	sSql = sSql & " ISNULL(refundamount,0.00) AS refundamount, P.pricetypename "
	sSql = sSql & " FROM egov_rentalreservationdatefees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.reservationdateid = " & iReservationDateId
	sSql = sSql & " ORDER BY P.displayorder"
	' AND feeamount > 0.0000 
	response.write "<!-- " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td colspan=""3"" align=""right""><strong>Rates</strong></td>"
		response.write "<td>&nbsp;</td>"
		response.write "</tr>"
		Do While Not oRs.EOF
			dDueAmount = CDbl(oRs("feeamount")) - ( CDbl(oRs("paidamount")) - CDbl(oRs("refundamount")) )
			If CDbl(dDueAmount) > CDbl(0.0000) Then 
				bOkToCharge = True 
			Else
				If CDbl(dDueAmount) = CDbl("0.00") Then
					If bNoCostToRent Then 
						' if not already in accounts ledger table
						If RentalFeeIsAlreadyPaid( iReservationDateId, "reservationdatefeeid", oRs("reservationdatefeeid") ) Then
							bOkToCharge = False
						Else 
							bOkToCharge = True 
						End If 
					Else
						' below was added to handle no charges for a normally chargable reservation
						If CDbl(oRs("feeamount")) = CDbl("0.00") Then
							' this check is new as of 6/12/2013 as I am not sure that they want to mark these as payed more than once
							If RentalFeeIsAlreadyPaid( iReservationDateId, "reservationdatefeeid", oRs("reservationdatefeeid") ) Then
								bOkToCharge = False
							Else 
								bOkToCharge = True 
							End If 
						Else
							bOkToCharge = False 
						End If 
					End If 
				Else
					' do not charge for negative due amounts. This means that a refund is due
					bOkToCharge = False 
				End If 
			End If 

			If bOkToCharge Then 
				iRentalRateCount = iRentalRateCount + clng(1)
				response.write vbcrlf & "<tr" & sClass & ">"
				response.write "<td colspan=""3"" align=""right"">"
				response.write oRs("pricetypename")
				response.write "<td align=""right"">"
				response.write "<input type=""hidden"" name=""reservationdatefeeid" & iRentalRateCount & """ value=""" & oRs("reservationdatefeeid") & """ />"
				response.write "<input type=""text"" id=""datefeeamount" & iRentalRateCount & """ name=""datefeeamount" & iRentalRateCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
				response.write " onchange=""return ValidateCharges( this );"" />"
				iTotalDue = iTotalDue + CDbl(dDueAmount)
				response.write "</td>"
				response.write "</tr>"
			End If 
			oRs.MoveNext 
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowItemsChargesForDate iReservationDateId, sClass, iDateTotal, iReservationItemCount, bNoCostToRent
'--------------------------------------------------------------------------------------------------
Sub ShowItemsChargesForDate( ByVal iReservationDateId, ByVal sClass, ByRef iTotalDue, ByRef iReservationItemCount, ByVal bNoCostToRent )
	Dim oRs, sSql, dDueAmount, bOkToCharge

	sSql = "SELECT reservationdateitemid, rentalitem, ISNULL(maxavailable,0) AS maxavailable, ISNULL(quantity,0) AS quantity, "
	sSql = sSql & " ISNULL(paidamount,0.00) AS paidamount, ISNULL(feeamount,0.00) AS feeamount, ISNULL(refundamount,0.00) AS refundamount "
	sSql = sSql & " FROM egov_rentalreservationdateitems "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " ORDER BY rentalitem"
	' AND feeamount > 0.00 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr" & sClass & "><td colspan=""3"" align=""right""><strong>Items</strong></td>"
		response.write "<td>&nbsp;</td>"
		response.write "</tr>"
		Do While Not oRs.EOF
			dDueAmount = CDbl(oRs("feeamount")) - ( CDbl(oRs("paidamount")) - CDbl(oRs("refundamount")) )
			If CDbl(dDueAmount) > CDbl(0.0000) Then 
				bOkToCharge = True 
			Else
				If CDbl(dDueAmount) = CDbl("0.00") Then
					If bNoCostToRent Then 
						' if not already in accounts ledger table
						If RentalFeeIsAlreadyPaid( iReservationDateId, "reservationdateitemid", oRs("reservationdateitemid") ) Then
							bOkToCharge = False
						Else 
							bOkToCharge = True 
						End If 
					Else
						bOkToCharge = False 
					End If 
				Else
					' do not charge for negative due amounts. This means that a refund is due
					bOkToCharge = False
				End If 
			End If 

			If bOkToCharge Then 
				iReservationItemCount = iReservationItemCount + clng(1)
				response.write vbcrlf & "<tr" & sClass & ">"
				response.write "<td colspan=""3"" align=""right"">"
				response.write "<input type=""hidden"" name=""reservationdateitemid" & iReservationItemCount & """ value=""" & oRs("reservationdateitemid") & """ />"
				response.write oRs("rentalitem")
				response.write " (" & oRs("quantity") & ")"
				response.write "<input type=""hidden"" name=""itemquantity" & iReservationItemCount & """ value=""" & oRs("quantity") & """ />"
				response.write "</td>"
				response.write "<td align=""right"">"
				response.write "<input type=""text"" id=""itemfeeamount" & iReservationItemCount & """ name=""itemfeeamount" & iReservationItemCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
				response.write " onchange=""return ValidateCharges( this );"" />"
				iTotalDue = iTotalDue + CDbl(dDueAmount)
				response.write "</td>"
				response.write "</tr>"
			End If 
			oRs.MoveNext 
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowReservationFeeCharges iReservationId, sClass, iTotalDue, iReservationFeeCount
'--------------------------------------------------------------------------------------------------
Sub ShowReservationFeeCharges( ByVal iReservationId, ByVal sClass, ByRef iTotalDue, ByRef iReservationFeeCount )
	Dim oRs, sSql, sCellClass, dDueAmount

	sSql = "SELECT F.reservationfeeid, P.pricetypename, ISNULL(F.feeamount,0.00) AS feeamount, "
	sSql = sSql & " ISNULL(F.refundamount,0.00) AS refundamount, ISNULL(F.paidamount,0.00) AS paidamount, P.pricetypename "
	sSql = sSql & " FROM egov_rentalreservationfees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.reservationid = " & iReservationId
	sSql = sSql & " AND feeamount > 0.0000 ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		dDueAmount = CDbl(oRs("feeamount")) - ( CDbl(oRs("paidamount")) - CDbl(oRs("refundamount")) )
		If CDbl(dDueAmount) > CDbl(0.0000) Then 
			iReservationFeeCount = iReservationFeeCount + clng(1)
			If iReservationFeeCount = clng(1) Then
				sCellClass = " class=""totalscell"""
			Else
				sCellClass = ""
			End If 
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td colspan=""3"" align=""right""" & sCellClass & ">"
			response.write oRs("pricetypename")
			response.write "</td>"
			response.write "<td align=""right""" & sCellClass & ">"
			response.write "<input type=""hidden"" name=""reservationfeeid" & iReservationFeeCount & """ & value=""" &  oRs("reservationfeeid") & """ />"
			response.write "<input type=""text"" id=""reservationfeeamount" & iReservationFeeCount & """ name=""reservationfeeamount" & iReservationFeeCount & """ value=""" & FormatNumber(dDueAmount,2,,,0) & """ size=""7"" maxlength=""7"""
			response.write " return onchange=""ValidateCharges( this );"" />"
			iTotalDue = iTotalDue + CDbl(dDueAmount)
			response.write "</td>"
			response.write "</tr>"
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPaymentChoices( iUserId, sBalanceDue, sReservationTypeSelector )
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentChoices( ByVal iUserId, ByVal sBalanceDue, ByVal sReservationTypeSelector )
	Dim sSql, oRs, iRowCount, sWhere

	iRowCount = clng(0)

	If sReservationTypeSelector = "admin" Then 
		sWhere = " AND P.isothermethod = 1 "
	End If 

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, requirescheckno, requirescitizenaccount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE O.paymenttypeid = P.paymenttypeid "
	sSql = sSql & " AND isadminmethod = 1 " & sWhere
	sSql = sSql & " AND O.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""50%"">"
		
		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" valign=""top"">"
		response.write "<input type=""checkbox"" id=""includezerocharges"" name=""includezerocharges"" />"
		response.write "</td>"
		response.write "<td valign=""top"">"
		response.write "<strong>Include $0.00 charges in this payment.</strong><br />If this is for a no charge rental, also put $0.00 in the &#39;Other&#39; payment amount."
		response.write "</td>"
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Citizen Location:</td>"
		response.write "<td>"

		ShowPaymentLocations	' In rentalsguifunctions.asp

		response.write "</td>"
		response.write "</tr>"

		Do While Not oRs.EOF
			iRowCount = iRowCount + clng(1)
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">"
			response.write oRs("paymenttypename") & ": "
			response.write "</td>"
			response.write "<td>"
			response.write "<input type=""hidden"" id=""paymenttypeid" & iRowCount & """ name=""paymenttypeid" & iRowCount & """ value=""" & oRs("paymenttypeid") & """ />"
			response.write "<input type=""text"" value="""" id=""paymentamount" & iRowCount & """ name=""paymentamount" & iRowCount & """ size=""10"" maxlength=""9"" onchange=""validatePayment( this )"" />"

			If oRs("requirescheckno") Then 
				response.write "&nbsp;<strong>Check #: </strong>"
				response.write "<input type=""hidden"" id=""hascheckno" & iRowCount & """ name=""hascheckno" & iRowCount & """ value=""yes"" />"
				response.write "<input type=""text"" value="""" id=""checkno" & iRowCount & """ name=""checkno" & iRowCount & """ size=""8"" maxlength=""8"" />"
			Else
				response.write "<input type=""hidden"" id=""hascheckno" & iRowCount & """ name=""hascheckno" & iRowCount & """ value=""no"" />"
			End If 

			If oRs("requirescitizenaccount") Then 
				response.write "&nbsp; <strong>From:</strong>" 
				ShowFamilyAccounts iUserId
			End If 

			response.write "</td>"
			response.write "</tr>"

			oRs.MoveNext
		Loop

		response.write vbcrlf & "<tr>" 
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Payment Total:</td>"
		response.write "<td>"
		response.write "<input type=""hidden"" id=""maxpayments"" name=""maxpayments"" value=""" & iRowCount & """ />"
		response.write "<input type=""hidden"" id=""paymenttotal"" name=""paymenttotal"" value="""" />"
		response.write "<span id=""paymenttotaldisplay"">0.00</span></td>"
		response.write "</tr>" 
		response.write "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Balance Due:</td>"
		response.write "<td><input type=""hidden"" id=""balancedue"" name=""balancedue"" value=""" & sBalanceDue & """ />"
		response.write "<span id=""balanceduedisplay"">" & FormatNumber(sBalanceDue,2,,,0) & "</span></td>" 
		response.write "</tr>" 
		response.write "<tr>" 
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"" valign=""top"">Notes:</td>"
		response.write "<td><textarea id=""purchasenotes"" name=""purchasenotes"" class=""notes"" maxlength=""500"" wrap=""soft""></textarea></td>"
		response.write "</tr>"
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

%>
