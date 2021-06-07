<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitpaymentsreport.asp
' AUTHOR: Steve Loar
' CREATED: 01/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permit payments by payment media (cash, check, cc)
'
' MODIFICATION HISTORY
' 1.0   01/28/2009	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromPaymentDate, sToPaymentDate, sStreetNumber, sStreetName, sPermitNo
Dim sPayor, sInvoiceNo, iInclude, sDisplayDateRange, iPermitCategoryId

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "permit payments report", sLevel	' In common.asp

' Handle payment date range. always want some dates to limit the search
If request("topaymentdate") <> "" And request("frompaymentdate") <> "" Then
	sFromPaymentDate = request("frompaymentdate")
	sToPaymentDate = request("topaymentdate")
	sSearch = sSearch & " AND (J.paymentdate >= '" & request("frompaymentdate") & "' AND J.paymentdate < '" & DateAdd("d",1,request("topaymentdate")) & "' ) "
	sDisplayDateRange = "From: " & request("frompaymentdate") & " &nbsp;To: " & request("topaymentdate")
Else
	' initially set these to today
	sFromPaymentDate = FormatDateTime(Date,2)
	sToPaymentDate = FormatDateTime(Date,2)
	sDisplayDateRange = ""
End If 

' handle the permit number
If request("permitno") <> "" Then 
	sPermitNo = Trim(request("permitno"))
	sSearch = sSearch & BuildPermitNoSearch( sPermitNo )	' in permitcommonfunctions.asp
End If 

If request("payor") <> "" Then 
	sPayor = request("payor")
	sSearch = sSearch & " AND ( C.company LIKE '%" & dbsafe(sPayor) & "%' OR C.firstname LIKE '%" & dbsafe(sPayor) & "%' OR C.lastname LIKE '%" & dbsafe(sPayor) & "%' ) "
End If 

If request("invoiceno") <> "" Then 
	sInvoiceNo = CLng(request("invoiceno"))
	sSearch = sSearch & " AND I.invoiceid = " & sInvoiceNo
End If 

If request("include") <> "" Then
	iInclude = request("include")
	If clng(iInclude) < clng(2) Then 
		sSearch = sSearch & " AND I.isvoided = " & iInclude
	End If 
Else
	iInclude = 0
	sSearch = sSearch & " AND I.isvoided = 0 "
End If 

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = request("permitcategoryid")
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSearch = sSearch & " AND P.permitcategoryid = " & iPermitCategoryId
	End If 
Else 
	iPermitCategoryId = "0"
End If 

If sSearch <> "" Then 
	session("sSql") = sSearch
Else 
	session("sSql") = ""
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--

		function validate()
		{
			// check the payment from date
			if ($("#frompaymentdate").val() != '')
			{
				if (! isValidDate($("#frompaymentdate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#frompaymentdate").focus();
					return;
				}
			}
			// check the inspection to date
			if ($("#topaymentdate").val() != '')
			{
				if (! isValidDate($("#topaymentdate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#topaymentdate").focus();
					return;
				}
			}
			document.frmPermitSearch.submit();
		}

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<div id="idControls" class="noprint">
		<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	</div>

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p id="pagetitle">
				<span id="printdaterange"><font size="+1"><strong><%=sDisplayDateRange%></strong></font></span>
				<font size="+1"><strong>Permit Payments Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitpaymentsreport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Permit Category:</td>
									<td><%	ShowPermitCategoryPicks iPermitCategoryId	' in permitcommonfunctions.asp	%></td>
								</tr>
								<tr>
									<td>Payment Date:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" id="frompaymentdate" name="frompaymentdate" value="<%=sFromPaymentDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp; To:
										<input type="text" id="topaymentdate" name="topaymentdate" value="<%=sToPaymentDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp;
										<%DrawPriorDateChoices "paymentdate" %>
									</td>
								</tr>
<!--								<tr>
									<td>Address:</td><td><%  'DisplayLargeAddressList sStreetNumber, sStreetName %></td>
								</tr>
-->
								<tr>
									<td>Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
								</tr>
								<tr>
									<td>Payor:</td><td><input type="text" id="payor" name="payor" size="50" maxlength="50" value="<%=sPayor%>" /></td>
								</tr>
								<tr>
									<td>Invoice #:</td><td><input type="text" id="invoiceno" name="invoiceno" size="20" maxlength="20" value="<%=sInvoiceNo%>" /></td>
								</tr>
								<tr>
									<td>Include:</td>
									<td>
										<select name="include">
											<option value="0"
<%												If clng(iInclude) = clng(0) Then response.write " selected=""selected"" "	%>
											>Payments Only</option>
											<option value="1"
<%												If clng(iInclude) = clng(1) Then response.write " selected=""selected"" "	%>											
											>Voided Invoices Only</option>
											<option value="2"
<%												If clng(iInclude) = clng(2) Then response.write " selected=""selected"" "	%>											
											>Payments and Voided Invoices</option>
										</select>
									</td>	
								</tr>
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										'If request("isview") <> "" Then		%>
										<!--	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="location.href='permitissuedreportexport.asp'" /> -->
<%										'End If		%>
									</td>
								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

			<!-- Begin: Report Display -->
<%			' if they choose to view the report, then display the payments
			If request("isview") <> "" Then	
				DisplayPermitPayments sSearch
			Else 
				response.write "<strong>To view the permit payments report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
			End If 
%>
			<!-- END: Report Display -->
		</div>
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
' void DisplayLargeAddressList sStreetNumber, sStreetName 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = oRs("residentstreetprefix") & " " 
			End If 

			sCompareName = sCompareName & oRs("residentstreetname")

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetsuffix")
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetdirection")
			End If 

			response.write vbcrlf & "<option value=""" & sCompareName & """"

			If sStreetName = sCompareName Then 
				response.write " selected=""selected"" "
			End If 

			response.write " >"
			response.write sCompareName & "</option>" 
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayPermitPayments ByVal sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayPermitPayments( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dInvoiceTotal, dPaidTotal, dCheckTotal, dCashTotal, dCCAdminTotal, dCCPublicTotal
	Dim iOldPID, iOldInvoiceId, dPIDCheckTotal, dPIDCashTotal, dPIDCCAdminTotal, dPIDCCPublicTotal, dIIDInvoiceTotal
	Dim bPidLineWasPrinted, dPIDPaymentTotal, bInvoiceIsVoided

	iRowCount = CLng(0)
	dInvoiceTotal = CDbl(0.00)
	dPaidTotal = CDbl(0.00)
	dCheckTotal = CDbl(0.00)
	dCashTotal = CDbl(0.00)
	dCCAdminTotal = CDbl(0.00)
	dCCPublicTotal = CDbl(0.00)
	iOldPID = CLng(0)
	iOldInvoiceId = CLng(0)
	bPidLineWasPrinted = False 
	dPIDCheckTotal = CDbl(0.00) 
	dPIDCashTotal = CDbl(0.00)
	dPIDCCAdminTotal = CDbl(0.00)
	dPIDCCPublicTotal = CDbl(0.00)
	dIIDInvoiceTotal = CDbl(0.00)
	bInvoiceIsVoided = False 

	sSql = "SELECT J.paymentid, I.permitid, I.invoiceid, I.isvoided, J.paymentdate, J.paymenttotal, ISNULL(V.checkno,'') AS checkno, A.amount, PT.paymenttypename, "
	sSql = sSql & " PT.ispublicmethod , PT.isadminmethod, C.company, C.firstname, C.lastname, ISNULL(I.totalamount,0.00) AS totalamount, "
	sSql = sSql & " PT.requirescheckno, PT.requirescreditcard, PT.requirescash "
	sSql = sSql & " FROM egov_permitinvoices I, egov_class_payment J, egov_permits P, egov_verisign_payment_information V, "
	sSql = sSql & " egov_accounts_ledger A, egov_paymenttypes PT, egov_permitcontacts C "
	sSql = sSql & " WHERE  P.orgid = " & session("orgid") & " AND I.allfeeswaived = 0 AND I.paymentid = J.paymentid "
	sSql = sSql & " AND J.isforpermits = 1 AND I.permitid = P.permitid and P.isvoided = 0 AND A.ispaymentaccount = 1 "
	sSql = sSql & " AND J.paymentid = V.paymentid AND A.ledgerid = V.ledgerid AND J.paymentid = A.paymentid "
	sSql = sSql & " AND A.paymenttypeid = PT.paymenttypeid AND C.permitcontactid = I.permitcontactid " & sSearch
	sSql = sSql & " ORDER BY J.paymentid, I.permitid, I.invoiceid"
'	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
		response.write vbcrlf & "<tr><th>Payment #</th><th>Payment<br />Date</th><th>Payor</th><th>Permit #</th><th>Invoice #</th><th>Invoiced</th>"
		response.write "<th>Check</th><th>Check #</th><th>Cash</th><th>Admin<br />Charge</th><th>Public<br />Charge</th><th>Total Paid</th></tr>"

		Do While Not oRs.EOF
			If iOldPID <> CLng(oRs("paymentid")) Or iOldInvoiceId <> CLng(oRs("invoiceid")) Then 
				If Not bPidLineWasPrinted Then
					If iOldPID <> CLng(0) Then
						' Add to the report totals
						dCheckTotal = dCheckTotal + CDbl(dPIDCheckTotal)
						dCashTotal = dCashTotal + CDbl(dPIDCashTotal)
						dCCAdminTotal = dCCAdminTotal + CDbl(dPIDCCAdminTotal)
						dCCPublicTotal = dCCPublicTotal + CDbl(dPIDCCPublicTotal)
						dPaidTotal = dPaidTotal + CDbl(dPIDPaymentTotal)

						' Print out PID line
						PrintPIDRow iRowCount, iOldPID, sPIDDate, sPayor, sPermitNo, iOldInvoiceId, bInvoiceIsVoided, dIIDInvoiceTotal, dPIDCheckTotal, sCheckNo, dPIDCashTotal, dPIDCCAdminTotal, dPIDCCPublicTotal, dPIDPaymentTotal 
					End If
					bPidLineWasPrinted = True 
				Else
					' the PID line has already printed, so print an invoice only line
					PrintInvoiceRow iRowCount, sPermitNo, iOldInvoiceId, bInvoiceIsVoided, dIIDInvoiceTotal
					dIIDInvoiceTotal = CDbl(0.00)
				End If 
				
				sPermitNo = GetPermitNumber( oRs("permitid") )
				iOldInvoiceId = oRs("invoiceid")
				bInvoiceIsVoided = oRs("isvoided")
				dIIDInvoiceTotal = CDbl(oRs("totalamount"))
				dInvoiceTotal = dInvoiceTotal + CDbl(oRs("totalamount"))
				dPIDPaymentTotal = FormatNumber(oRs("paymenttotal"),2)

				If iOldPID <> CLng(oRs("paymentid")) Then
					iOldPID = CLng(oRs("paymentid"))
					sPIDDate = FormatDateTime(oRs("paymentdate"),2)
					If oRs("firstname") <> "" Then 
						sPayor = oRs("firstname") & " " & oRs("lastname")
					Else
						sPayor = oRs("company")
					End If 

					bPidLineWasPrinted = False 
					dPIDCheckTotal = CDbl(0.00)
					dPIDCashTotal = CDbl(0.00)
					dPIDCCAdminTotal = CDbl(0.00)
					dPIDCCPublicTotal = CDbl(0.00)
					sCheckNo = ""
				End If 
				
			End If 

			If oRs("requirescheckno") Then 
				dPIDCheckTotal = dPIDCheckTotal + CDbl(oRs("amount"))
				sCheckNo = oRs("checkno")
			End If 
			If oRs("requirescash") Then 
				dPIDCashTotal = dPIDCashTotal + CDbl(oRs("amount"))
			End If 
			If oRs("requirescreditcard") And oRs("isadminmethod") Then 
				dPIDCCAdminTotal = dPIDCCAdminTotal + CDbl(oRs("amount"))
			End If 
			If oRs("requirescreditcard") And oRs("ispublicmethod") Then 
				dPIDCCPublicTotal = dPIDCCPublicTotal + CDbl(oRs("amount"))
			End If 
			oRs.MoveNext 
		Loop
		' Print the last line of data
		If Not bPidLineWasPrinted Then
			If iOldPID <> CLng(0) Then
				' Add to the report totals
				dCheckTotal = dCheckTotal + CDbl(dPIDCheckTotal)
				dCashTotal = dCashTotal + CDbl(dPIDCashTotal)
				dCCAdminTotal = dCCAdminTotal + CDbl(dPIDCCAdminTotal)
				dCCPublicTotal = dCCPublicTotal + CDbl(dPIDCCPublicTotal)
				dPaidTotal = dPaidTotal + CDbl(dPIDPaymentTotal)

				' Print out PID line
				PrintPIDRow iRowCount, iOldPID, sPIDDate, sPayor, sPermitNo, iOldInvoiceId, bInvoiceIsVoided, dIIDInvoiceTotal, dPIDCheckTotal, sCheckNo, dPIDCashTotal, dPIDCCAdminTotal, dPIDCCPublicTotal, dPIDPaymentTotal 
			End If
		Else
			' the PID line has already printed, so print an invoice only line
			PrintInvoiceRow iRowCount, sPermitNo, iOldInvoiceId, bInvoiceIsVoided, dIIDInvoiceTotal
		End If 

		' Totals row at the bottom
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">Totals</td>"
		response.write "<td align=""right"">" & FormatNumber(dInvoiceTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCheckTotal,2) & "</td>"
		response.write "<td>&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(dCashTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCCAdminTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCCPublicTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dPaidTotal,2) & "</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table></div>"
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void  PrintPIDRow iRowCount iPID, sPIDDate, sPayor, sPermitNo, iInvoiceId, bInvoiceIsVoided, dInvoiceTotal, dCheckTotal, sCheckNo, dCashTotal, dCCAdminTotal, dCCPublicTotal, dPaidTotal
'--------------------------------------------------------------------------------------------------
Sub PrintPIDRow( ByRef iRowCount, ByVal iPID, ByVal sPIDDate, ByVal sPayor, ByVal sPermitNo, _
	ByVal iInvoiceId, ByVal bInvoiceIsVoided, ByVal dInvoiceTotal, ByVal dCheckTotal, ByVal sCheckNo, ByVal dCashTotal, _
	ByVal dCCAdminTotal, ByVal dCCPublicTotal, ByVal dPaidTotal )

	iRowCount = iRowCount + CLng(1)
	response.write vbcrlf & "<tr"
	If iRowCount Mod 2 = 0 Then
		response.write " class=""altrow"""
	End If 
	response.write ">"
	response.write "<td align=""center"" nowrap=""nowrap"">" & iPID & "</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & sPIDDate & "</td>"
	response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & sPayor & "</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & sPermitNo & "</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & iInvoiceId
	If bInvoiceIsVoided Then
		response.write "v"
	End If 
	response.write "</td>"
	response.write "<td align=""right"">" & FormatNumber(dInvoiceTotal,2) & "</td>"
	If dCheckTotal <> CDbl(0.00) Then 
		response.write "<td align=""right"">" & FormatNumber(dCheckTotal,2) & "</td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	If sCheckNo <> "" Then 
		response.write "<td align=""center"" nowrap=""nowrap"">" & sCheckNo & "</td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	If dCashTotal <> CDbl(0.00) Then 
		response.write "<td align=""right"">" & FormatNumber(dCashTotal,2) & "</td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	If dCCAdminTotal <> CDbl(0.00) Then 
		response.write "<td align=""right"">" & FormatNumber(dCCAdminTotal,2) & "</td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	If dCCPublicTotal <> CDbl(0.00) Then 
		response.write "<td align=""right"">" & FormatNumber(dCCPublicTotal,2) & "</td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If
	response.write "<td align=""right"">" & FormatNumber(dPaidTotal,2) & "</td>"
	response.write "</tr>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void PrintInvoiceRow iRowCount, sPermitNo, iInvoiceId, bInvoiceIsVoided, dInvoiceTotal
'--------------------------------------------------------------------------------------------------
Sub PrintInvoiceRow( ByRef iRowCount, ByVal sPermitNo, ByVal iInvoiceId, ByVal bInvoiceIsVoided, ByVal dInvoiceTotal )

	iRowCount = iRowCount + CLng(1)
	response.write vbcrlf & "<tr"
	If iRowCount Mod 2 = 0 Then
		response.write " class=""altrow"""
	End If 
	response.write ">"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & sPermitNo & "</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & iInvoiceId
	If bInvoiceIsVoided Then
		response.write "v"
	End If 
	response.write "</td>"
	response.write "<td align=""right"">" & FormatNumber(dInvoiceTotal,2) & "</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "<td>&nbsp;</td>"
	response.write "</tr>"

End Sub 


Sub Saveforlater()
	iRowCount = iRowCount + 1
	response.write vbcrlf & "<tr"
	If iRowCount Mod 2 = 0 Then
		response.write " class=""altrow"""
	End If 
	response.write ">"
	response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("paymentid") & "</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & FormatDateTime(oRs("paymentdate"),2) & "</td>"
	If oRs("firstname") <> "" Then 
		response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & oRs("firstname") & " " & oRs("lastname") & "</td>"
	Else
		response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & oRs("company") & "</td>"
	End If 
	response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
	response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("invoiceid") & "</td>"
	response.write "<td align=""right"">" & FormatNumber(oRs("totalamount"),2) & "</td>"
	dInvoiceTotal = dInvoiceTotal + CDbl(oRs("totalamount"))
	If oRs("requirescheckno") Then 
		response.write "<td align=""right"">" & FormatNumber(oRs("amount"),2) & "</td>"
		dCheckTotal = dCheckTotal + CDbl(oRs("amount"))
		response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("checkno") & "</td>"
	Else
		response.write "<td>&nbsp;</td>"
		response.write "<td>&nbsp;</td>"
	End If 
	If oRs("requirescash") Then 
		response.write "<td align=""right"">" & FormatNumber(oRs("amount"),2) & "</td>"
		dCashTotal = dCashTotal + CDbl(oRs("amount"))
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	If oRs("requirescreditcard") And oRs("isadminmethod") Then 
		response.write "<td align=""right"">" & FormatNumber(oRs("amount"),2) & "</td>"
		dCCAdminTotal = dCCAdminTotal + CDbl(oRs("amount"))
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	If oRs("requirescreditcard") And oRs("ispublicmethod") Then 
		response.write "<td align=""right"">" & FormatNumber(oRs("amount"),2) & "</td>"
		dCCPublicTotal = dCCPublicTotal + CDbl(oRs("amount"))
	Else
		response.write "<td>&nbsp;</td>"
	End If 
	response.write "<td align=""right"">" & FormatNumber(oRs("paymenttotal"),2) & "</td>"
	dPaidTotal = dPaidTotal + CDbl(oRs("paymenttotal"))
	response.write "</tr>"
End Sub 




%>
