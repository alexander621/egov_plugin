<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payinvoices.asp
' AUTHOR: Steve Loar
' CREATED: 08/20/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Allows payments of invoices for one contact and multiple permits and invoices.
'
' MODIFICATION HISTORY
' 1.0   08/20/2008	Steve Loar - INITIAL VERSION
' 1.1	09/25/2008	Steve Loar - Changed to handle citizen payments as well as contractors
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iBillTo, iMaxInvoices, sTotalDue, iMaxPaymentChoices, sContactType

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "pay invoices", sLevel	' In common.asp

sContactType = "C"

If request("permitcontacttypeid") <> "" Then 
	iBillTo = CLng(Mid(request("permitcontacttypeid"),2))
	sContactType = Left(request("permitcontacttypeid"),1)
Else
	iBillTo = GetFirstBillTo( sContactType )
End If 

sTotalDue = 0.00
iMaxInvoices = 0
iMaxPaymentChoices = 0

%>

<html>
	<head>
		<title>E-Gov Administration Console</title>

		<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="permits.css" />

		<script language="JavaScript" src="../scripts/formatnumber.js"></script>
		<script language="JavaScript" src="../scripts/removespaces.js"></script>
		<script language="JavaScript" src="../scripts/removecommas.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

		<script language="Javascript">
		<!--
			var w = (screen.width - 640)/2;
			var h = (screen.height - 480)/2;

			function ViewDetails( iPermitId )
			{
				//var winHandle = eval('window.open("viewpermitdetails.asp?permitid=' + iPermitId + '", "_contact", "width=900,height=700,location=0,toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1,left=' + w + ',top=' + h + '")');
				showModal('viewpermitdetails.asp?permitid=' + iPermitId, 'Permit Details', 50, 90);
			}

			function ShowInvoices( )
			{
				document.frmPay.action = "payinvoices.asp";
				document.frmPay.submit();
			}

			function calcTotalDue()
			{
				var totaldue = 0.00;

				for (var t = 1; t <= parseInt(document.frmPay.maxinvoices.value); t++)
				{
					if ($("includeinvoice" + t).checked)
					{
						totaldue += Number($("invoiceamount" + t).value);
					}
				}
				$("totaldue").value = format_number(totaldue,2);
				$("totalduedisplay").innerHTML = format_number(totaldue,2);
				setTotalDue();
			}

			function addPaymentTotal()
			{
				var totalPayment = 0.00;
				var Ok;
				var rege;

				for (var t = 1; t <= parseInt(document.frmPay.maxpayments.value); t++)
				{
					if ($("amount" + t).value != '')
					{
						// Remove any extra spaces
						$("amount" + t).value = removeSpaces($("amount" + t).value);
						//Remove commas that would cause problems in validation
						$("amount" + t).value = removeCommas($("amount" + t).value);
						rege = /^\d*\.?\d{0,2}$/
						Ok = rege.exec($("amount" + t).value);
						if ( Ok )
						{
							totalPayment += Number($("amount" + t).value);
							$("amount" + t).value = format_number($("amount" + t).value,2);
						}
						else
						{
							alert('Payment values should be currency or blank.\nPlease correct this amount.');
							$("amount" + t).focus();
							$("amount" + t).value = '';
							return;
						}
					}
				}
				$("paymenttotal").value = format_number(totalPayment,2);
				$("paymenttotaldisplay").innerHTML = format_number(totalPayment,2);
				setTotalDue();
			}

			function setTotalDue()
			{
				var nBalance = 0.00;
				nBalance = Number($("totaldue").value) - Number($("paymenttotal").value);
				$("balancedue").innerHTML = format_number(nBalance,2);
			}

			function doClose()
			{
				window.close();
				window.opener.focus();
			}

			function validatePayment()
			{
				var t;
				var bHasInvoices = false;

				// Check that at least one invoice is checked
				for (t = 1; t <= parseInt(document.frmPay.maxinvoices.value); t++)
				{
					if ($("includeinvoice" + t).checked)
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
									rege = /^\d*[&, ]*\d*$/
									Ok = rege.exec($("checkno" + t).value);
									if ( ! Ok )
									{
										alert('The check number must be numeric.\nFor more than one, seperate by spaces, comma or ampersand.\nPlease correct the check number and try again.');
										$("checkno" + t).focus();
										return;
									}
								}
							}
						}
					}
					// All is OK so submit the form
					document.frmPay.action = "payinvoicesprocess.asp";
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

		<% ShowHeader sLevel %>
		<!--#Include file="../menu/menu.asp"--> 

		<div id="content">
			<div id="centercontent">
			<div class="gutters">
				<font size="+1"><strong>Permit Invoice Payment</strong></font><br /><br />

				<form name="frmPay" action="payinvoices.asp" method="post">
					<input type="hidden" name="contacttype" value="<%=sContactType%>" />
<%					If iBillTo = CLng(0) then	%>
						<p>
							There are no invoices that need to be paid at this time.
						</p>
<%					Else	%>
						<p>
							Select the Payer: <% ShowBillingChoices iBillTo, sContactType %>
						</p>
						<p>
							<table cellpadding="2" cellspacing="0" border="0" class="feetable" id="invoicelist">
								<caption>Unpaid Invoices</caption>
								<tr><th>Include</th><th>Invoice #</th><th>Invoice<br />Date</th><th>Permit #</th><th>Address</th><th>Invoice<br />Total</th></tr>
	<%							iMaxInvoices = ShowInvoicesToPay( iBillTo, sTotalDue, sContactType ) %>
								<tr><td colspan="5" align="right" class="totalline">
									<input type="hidden" name="totaldue" id="totaldue" value="<%=FormatNumber(sTotalDue,2,,,0)%>" />
									<strong>Total Due: </strong></td>
									<td align="right" class="totalline"><strong><span id="totalduedisplay"><%=FormatNumber(sTotalDue,2,,,0)%></span></strong>&nbsp;</td>
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
							<br /><br /><input type="button" class="button ui-button ui-widget ui-corner-all" id="savebutton" name="complete" value="Complete Payment" onClick="validatePayment()" />
						</p>
<%					End If	%>					
				</form>
			</div>
			</div>
		</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

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

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, requirescheckno, requirescitizenaccount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE O.paymenttypeid = P.paymenttypeid AND isadminmethod = 1 AND isforpermits = 1 AND O.orgid = " & Session("OrgID")
	sSql = sSql & " ORDER BY displayorder"

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
' Function ShowInvoicesToPay( iContactId, sTotalDue, sContactType )
'--------------------------------------------------------------------------------------------------
Function ShowInvoicesToPay( ByVal iContactId, ByRef sTotalDue, ByVal sContactType )
	Dim sSql, oRs, iRecCount, sWhere 

	iRecCount = CLng(0) 

	If sContactType = "C" Then
		sWhere = "permitcontacttypeid"
	Else
		sWhere = "userid"
	End If 

	sSql = "SELECT I.permitid, I.invoiceid, I.invoicedate, I.totalamount, ISNULL(I.paymentid,0) AS paymentid, "
	sSql = sSql & " ISNULL(C.company,'') AS company, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & " ISNULL(A.residentunit,'') AS residentunit, ISNULL(A.streetsuffix,'') AS streetsuffix, "
	sSql = sSql & " ISNULL(A.streetdirection,'') AS streetdirection, A.residentstreetname "
	sSql = sSql & " FROM egov_permitinvoices I, egov_permitcontacts C, egov_permitaddress A "
	sSql = sSql & " WHERE I.permitcontactid = C.permitcontactid AND I.isvoided = 0 AND I.allfeeswaived = 0 AND I.paymentid IS NULL "
	sSql = sSql & " AND A.permitid = I.permitid AND C." & sWhere & " = " & iContactId
	sSql = sSql & " ORDER BY I.invoiceid"
	'response.write sSql & "<br />"

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
			response.write "<td align=""center""><a href=""javascript:ViewDetails(" & oRs("permitid") & ");"" >" & GetPermitNumber( oRs("permitid") ) & "</a></td>"

			response.write "<td align=""center"">"
			response.write oRs("residentstreetnumber")
			If oRs("residentstreetprefix") <> "" Then
				response.write " " & oRs("residentstreetprefix")
			End If 
			response.write " " & oRs("residentstreetname")
			If oRs("streetsuffix") <> "" Then
				response.write " " & oRs("streetsuffix")
			End If 
			If oRs("streetdirection") <> "" Then
				response.write " " & oRs("streetdirection")
			End If 
			If oRs("residentunit") <> "" Then 
				response.write ", " & oRs("residentunit")
			End If 
			response.write "</td>"

'			response.write "<td align=""center"">"
'			If oRs("firstname") <> "" Then 
'				response.write oRs("firstname") & " " & oRs("lastname")
'			Else
'				response.write oRs("company")
'			End If 
'			response.write "</td>"

			response.write "<td align=""right"">" 
			response.write "<input type=""hidden"" id=""invoiceamount" & iRecCount & """ name=""invoiceamount" & iRecCount & """ value=""" & oRs("totalamount") & """ />"
			response.write FormatNumber(oRs("totalamount"),2,,,0) & "&nbsp;</td>"
			sTotalDue = sTotalDue + CDbl(oRs("totalamount"))
			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowInvoicesToPay = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetFirstBillTo( )
'--------------------------------------------------------------------------------------------------
Function GetFirstBillTo( ByRef sContactType )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(C.permitcontacttypeid,0) AS permitcontacttypeid, ISNULL(C.userid,0) AS userid, ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permitinvoices I "
	sSql = sSql & " WHERE C.permitcontactid = I.permitcontactid AND C.permitid = I.permitid AND "
	sSql = sSql & " I.paymentid is NULL AND I.isvoided = 0 AND I.allfeeswaived = 0  AND C.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY sortname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		' Grab the first one and pass that back
		If CLng(oRs("permitcontacttypeid")) = CLng(0) Then 
			GetFirstBillTo = CLng(oRs("userid"))
			sContactType = "U"
		Else
			GetFirstBillTo = CLng(oRs("permitcontacttypeid"))
			sContactType = "C"
		End If 
	Else
		' Nothing to Pay
		GetFirstBillTo = CLng(0)
		sContactType = "C"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowBillingChoices( iBillTo, sContactType )
'--------------------------------------------------------------------------------------------------
Sub ShowBillingChoices( iBillTo, sContactType )
	Dim sSql, oRs

	sSql = "SELECT DISTINCT ISNULL(C.permitcontacttypeid,0) AS permitcontacttypeid, ISNULL(C.userid,0) AS userid, "
	sSql = sSql & " ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permitinvoices I "
	sSql = sSql & " WHERE C.permitcontactid = I.permitcontactid AND C.permitid = I.permitid AND "
	sSql = sSql & " I.paymentid IS NULL AND I.isvoided = 0 AND I.allfeeswaived = 0 AND C.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY 5"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""permitcontacttypeid"" onchange=""ShowInvoices( this );"">"
		Do While Not oRs.EOF
			If oRs("permitcontacttypeid") > CLng(0) Then 
				response.write vbcrlf & "<option value=""C" & oRs("permitcontacttypeid") & """"
				If sContactType = "C" And CLng(iBillTo) = CLng(oRs("permitcontacttypeid")) Then
					response.write " selected=""selected"" "
				End If 
			Else
				response.write vbcrlf & "<option value=""U" & oRs("userid") & """"
				If sContactType = "U" And CLng(iBillTo) = CLng(oRs("userid")) Then
					response.write " selected=""selected"" "
				End If
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
