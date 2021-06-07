<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: paypermitinvoices.asp
' AUTHOR: Steve Loar
' CREATED: 05/29/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Allows payments fo invoices for one contact and one permit.
'
' MODIFICATION HISTORY
' 1.0   05/29/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iBillToDefault, iMaxInvoices, sTotalDue, sPermitNo, iMaxPaymentChoices

iPermitId = CLng(request("permitid"))

If request("permitcontactid") = "" Then 
	iBillToDefault = GetFirstBillTo( iPermitId )
Else
	iBillToDefault = CLng(request("permitcontactid"))
End If 

iMaxInvoices = 0
sTotalDue = CDbl(0.00)
sPermitNo = GetPermitNumber( iPermitId )  ' in permitcommonfunctions.asp

%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/formatnumber.js"></script>
		<script language="JavaScript" src="../scripts/removespaces.js"></script>
		<script language="JavaScript" src="../scripts/removecommas.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="Javascript">
		<!--

			function ShowInvoices( )
			{
				document.frmPay.action = "paypermitinvoices.asp";
				document.frmPay.submit();
			}

			function calcTotalDue()
			{
				var totaldue = 0.00;

				for (var t = 1; t <= parseInt(document.frmPay.maxinvoices.value); t++)
				{
					if (document.getElementById("includeinvoice" + t).checked)
					{
						totaldue += Number(document.getElementById("invoiceamount" + t).value);
					}
				}
				document.getElementById("totaldue").value = format_number(totaldue,2);
				document.getElementById("totalduedisplay").innerHTML = format_number(totaldue,2);
				setTotalDue();
			}

			function addPaymentTotal()
			{
				var totalPayment = 0.00;
				var Ok;
				var rege;

				for (var t = 1; t <= parseInt(document.frmPay.maxpayments.value); t++)
				{
					if (document.getElementById("amount" + t).value != '')
					{
						// Remove any extra spaces
						document.getElementById("amount" + t).value = removeSpaces(document.getElementById("amount" + t).value);
						//Remove commas that would cause problems in validation
						document.getElementById("amount" + t).value = removeCommas(document.getElementById("amount" + t).value);
						rege = /^\d*\.?\d{0,2}$/
						Ok = rege.exec(document.getElementById("amount" + t).value);
						if ( Ok )
						{
							totalPayment += Number(document.getElementById("amount" + t).value);
							document.getElementById("amount" + t).value = format_number(document.getElementById("amount" + t).value,2);
						}
						else
						{
							alert('Payment values should be currency or blank.\nPlease correct this amount.');
							document.getElementById("amount" + t).focus();
							document.getElementById("amount" + t).value = '';
							return;
						}
					}
				}
				document.getElementById("paymenttotal").value = format_number(totalPayment,2);
				document.getElementById("paymenttotaldisplay").innerHTML = format_number(totalPayment,2);
				setTotalDue();
			}

			function setTotalDue()
			{
				var nBalance = 0.00;
				nBalance = Number(document.getElementById("totaldue").value) - Number(document.getElementById("paymenttotal").value);
				document.getElementById("balancedue").innerHTML = format_number(nBalance,2);
			}

			function doClose()
			{
				//window.close();
				//window.opener.focus();
				parent.hideModal(window.frameElement.getAttribute("data-close"));
			}

			function validatePayment()
			{
				var t;
				var bHasInvoices = false;

				// Check that at least one invoice is checked
				for (t = 1; t <= parseInt(document.frmPay.maxinvoices.value); t++)
				{
					if (document.getElementById("includeinvoice" + t).checked)
					{
						bHasInvoices = true;
					}
				}
				if (bHasInvoices == false)
				{
					alert("Please select at least one invoice to pay.");
					return;
				}

				// Check that the total due equals the total paid
				if (Number($("totaldue").value) == Number($("paymenttotal").value))
				{
					// look for missing check numbers
					for (t = 1; t <= parseInt(document.frmPay.maxpayments.value); t++)
					{
						if ($("amount" + t).value != '')
						{
							if ($("checkno" + t))
							{
								if ($("checkno" + t).value == '')
								{
									if ( ! confirm('You have a check payment without a check number.\nDo you wish to continue?'))
									{
										$("checkno" + t).focus();
										return;
									}
								}
								else
								{
									// validate that the check number is a number
									//rege = /^\d*$/
									rege = /^\d*[&, ]*\d*$/
									Ok = rege.exec($("checkno" + t).value);
									if ( ! Ok )
									{
										//alert('The check number must be numeric.\nPlease correct the check number and try again.');
										alert('The check number must be numeric.\nFor more than one, seperate by spaces, comma or ampersand.\nPlease correct the check number and try again.');
										$("checkno" + t).focus();
										return;
									}
								}
							}
						}
					}

					// All is OK so submit the form
					document.frmPay.action = "paypermitinvoicesprocess.asp";
					document.frmPay.submit();
				}
				else
				{
					alert("The payment total must equal the total due.\nPlease adjust the payments and try again.");
				}
			}

		//-->
		</script>

	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<script>parent.document.getElementById('modaltitle'+window.frameElement.getAttribute("data-close")).innerHTML='Pay Permit Invoices for Permit <%=sPermitNo%>';</script>

				<form name="frmPay" action="paypermitinvoices.asp" method="post">
					<input type="hidden" name="permitid" value="<%=iPermitId%>" />
					<p>
						Select the Payer: <% ShowBillingChoices iPermitId, iBillToDefault %>
					</p>
					<p>
						<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="invoicelist">
							<caption>Invoices</caption>
							<tr><th>Include</th><th>Invoice #</th><th>Date</th><th>Billed To</th><th>Invoice Total</th></tr>
<%							iMaxInvoices = ShowInvoicesToPay( iPermitId, iBillToDefault, sTotalDue ) %>
							<tr><td colspan="4" align="right" class="totalline">
								<input type="hidden" name="totaldue" id="totaldue" value="<%=FormatNumber(sTotalDue,2,,,0)%>" />
								<strong>Total Due: </strong></td>
								<td align="center" class="totalline"><strong><span id="totalduedisplay"><%=FormatNumber(sTotalDue,2,,,0)%></span></strong></td>
							</tr>
						</table>
						<input type="hidden" id="maxinvoices" name="maxinvoices" value="<%=iMaxInvoices%>" />
					</p>
					<p>
						<fieldset><legend><strong> Payment </strong></legend><br />
						<input type="hidden" value="0.00" id="paymenttotal" name="paymenttotal" />
<%						iMaxPaymentChoices = ShowPaymentChoices( sTotalDue ) %>
						<input type="hidden" id="maxpayments" name="maxpayments" value="<%=iMaxPaymentChoices%>" />
						</fieldset>
						<br /><br /><input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" name="complete" value="Complete Payment" onClick="validatePayment()" /> &nbsp; &nbsp;
						<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" />
					</p>
					
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function ShowPaymentChoices()
'--------------------------------------------------------------------------------------------------
Function ShowPaymentChoices( ByVal sBalanceDue )
	Dim sSql, oPayments, iRecCount

	iRecCount = CLng(0) 

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, requirescheckno, requirescitizenaccount FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE O.paymenttypeid = P.paymenttypeid AND isadminmethod = 1 AND isforpermits = 1 AND O.orgid = " & Session("OrgID")
	sSql = sSql & " ORDER BY displayorder"
	'response.write sSql & "<br />"

	Set oPayments = Server.CreateObject("ADODB.Recordset")
	oPayments.Open sSQL, Application("DSN"), 3, 1

	If Not oPayments.EOF Then 
		response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""50%"" id=""paymentlist"">"
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"">Payer Location:</td><td>" 
		ShowPaymentLocations  ' In permitcommonfunctions.asp
		response.write "</td></tr>"
		Do While Not oPayments.EOF 
			iRecCount = iRecCount + CLng(1)
			response.write vbcrlf & "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">"
			response.write "<input type=""hidden"" id=""paymenttypeid" & iRecCount & """ name=""paymenttypeid" & iRecCount & """ value=""" & oPayments("paymenttypeid") & """ />"
			response.write oPayments("paymenttypename") & ": "
			response.write "</td><td>"
			response.write "<input type=""text"" value="""" id=""amount" & iRecCount & """ name=""amount" & iRecCount & """ size=""10"" maxlength=""10"" onblur=""addPaymentTotal()"" />"
			If oPayments("requirescheckno") Then
				response.write " &nbsp; <strong>Check #:</strong>&nbsp;<input type=""text"" value="""" id=""checkno" & iRecCount & """ name=""checkno" & iRecCount & """ size=""18"" maxlength=""18"" />"
			End If 
			response.write "</td></tr>"
			oPayments.MoveNext
		Loop
		response.write vbcrlf & "<tr><td class=""totalline"" align=""right"" nowrap=""nowrap""><strong>Payment Total:</strong></td><td class=""totalline""><span id=""paymenttotaldisplay"">0.00</span></td></tr>"
		response.write vbcrlf & "<tr><td class=""totalline"" align=""right"" nowrap=""nowrap"">Balance Due:</td><td class=""totalline""><span id=""balancedue"">" & FormatNumber(sBalanceDue,2,,,0) & "</span></td></tr>"
		'response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"">Notes:</td><td><textarea name=""notes"" class=""purchasenotes""></textarea></td></tr>"
		response.write vbcrlf & "</table>"
	End If 
	
	oPayments.close
	Set oPayments = Nothing

	ShowPaymentChoices = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' Function ShowInvoicesToPay( ByVal iPermitId, ByVal iPermitContactId, ByRef sTotalDue )
'--------------------------------------------------------------------------------------------------
Function ShowInvoicesToPay( ByVal iPermitId, ByVal iPermitContactId, ByRef sTotalDue )
	Dim sSql, oRs, iRecCount

	iRecCount = CLng(0) 

	sSql = "SELECT I.invoiceid, I.invoicedate, I.totalamount, ISNULL(I.paymentid,0) AS paymentid, "
	sSql = sSql & " ISNULL(C.company,'') AS company, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname "
	sSql = sSql & " FROM egov_permitinvoices I, egov_permitcontacts C "
	sSql = sSql & " WHERE I.permitcontactid = C.permitcontactid AND I.isvoided = 0 AND I.allfeeswaived = 0 AND I.permitcontactid = " & iPermitContactId
	sSql = sSql & " AND I.paymentid IS NULL AND I.permitid = " & iPermitId & " ORDER BY I.invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRecCount = iRecCount + CLng(1)
			response.write vbcrlf & "<tr"
			If iRecCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write "<td align=""center"">"
			response.write "<input type=""hidden"" name=""invoiceid" & iRecCount & """ value=""" & oRs("invoiceid") & """ />"
			response.write "<input type=""checkbox"" checked=""checked"" id=""includeinvoice" & iRecCount & """ name=""includeinvoice" & iRecCount & """ onclick=""calcTotalDue();"" /></td>"
			response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("invoicedate"),2) & "</td>"
			
			response.write "<td align=""center"">"
			If oRs("firstname") <> "" Then 
				response.write oRs("firstname") & " " & oRs("lastname")
			Else
				response.write oRs("company")
			End If 
			response.write "</td>"

			response.write "<td align=""center"">" 
			response.write "<input type=""hidden"" id=""invoiceamount" & iRecCount & """ name=""invoiceamount" & iRecCount & """ value=""" & oRs("totalamount") & """ />"
			response.write FormatNumber(oRs("totalamount"),2,,,0) & "</td>"
			sTotalDue = sTotalDue + CDbl(oRs("totalamount"))
'			If CLng(oRs("paymentid")) = CLng(0) Then
'				response.write "<td align=""center"">0.00</td>"
'			Else 
'				response.write "<td align=""center"">" & GetPermitPaymentTotal( CLng(oRs("paymentid")) )   ' in permitcommonfunctions.asp
'				response.write "</td>"
'			End If 
			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowInvoicesToPay = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetFirstBillTo( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetFirstBillTo( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT C.permitcontactid, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permitinvoices I "
	sSql = sSql & " WHERE C.permitcontactid = I.permitcontactid AND C.permitid = I.permitid AND I.paymentid is NULL AND C.permitid = " & iPermitId
	sSql = sSql & " ORDER BY 2"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		' Grab the first one and pass that back
		GetFirstBillTo = CLng(oRs("permitcontactid"))
	Else
		' someting is wrong
		GetFirstBillTo = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowBillingChoices( iPermitId, iBillToDefault )
'--------------------------------------------------------------------------------------------------
Sub ShowBillingChoices( iPermitId, iBillToDefault )
	Dim sSql, oRs

	sSql = "SELECT DISTINCT C.permitcontactid, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permitinvoices I "
	sSql = sSql & " WHERE C.permitcontactid = I.permitcontactid AND C.permitid = I.permitid AND I.paymentid is NULL AND C.permitid = " & iPermitId
	sSql = sSql & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""permitcontactid"" onchange=""ShowInvoices( this );"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitcontactid") & """"
			If CLng(iBillToDefault) = CLng(oRs("permitcontactid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			If oRs("firstname") <> "" Then
				response.write oRs("firstname") & " " & oRs("lastname")
				bName = True 
			Else
				bName = False 
			End If 
			If oRs("company") <> "" Then
				If bName Then 
					response.write " ("
				End If 
				response.write oRs("company")
				If bName Then
					response.write ")"
				End If 
			End If 
			response.write "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
