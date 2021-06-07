<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: payinvoicesummary.asp
' AUTHOR: Steve Loar
' CREATED: 08/22/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the summary of invoices for a payment.
'
' MODIFICATION HISTORY
' 1.0   08/22/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iInvoiceId, sInvoiceDate, sInvoiceTotal, iPermitId, sPaymentTotal, iPermitContactId
Dim sInvoiceList, bIsWaived, bIsVoided, iPaymentId, iPermitCount, bGroupByInvoiceCategories

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "pay invoices", sLevel	' In common.asp

iPaymentId = CLng(request("paymentid"))
iPermitContactId = GetPermitContactForPayment( iPaymentId )

sInvoiceDate = ""
sInvoiceTotal = CDbl(0.00)
sPaymentTotal = CDbl(0.00 )
sInvoiceList = ""
bIsWaived = False 
bIsVoided = False 
iPermitCount = CLng(0)

%>

<html>
<head>
	<title>E-Gov Permit Invoice</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="Javascript">
	<!--

		function doClose()
		{
			window.close();
			window.opener.focus();
		}

	//-->
	</script>

</head>

<body>
		<% ShowHeader sLevel %>
		<!--#Include file="../menu/menu.asp"--> 

 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:location.href='payinvoices.asp';" value="Make Another Payment" />&nbsp;&nbsp;
	<!-- <input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> -->
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
<%	' Need to get the invoices on this payment and loop through them with a page break on each
	sSql = "SELECT I.invoiceid, I.permitid, I.invoicedate, I.totalamount, S.invoicestatus FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.orgid = " & session("orgid") & " AND paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sInvoiceTotal = 0.00
		sPaymentTotal = 0.00
		sInvoiceTotal = CDbl(oRs("totalamount"))
		sStatus = oRs("invoicestatus")
		bGroupByInvoiceCategories = GetIsGroupByInvoiceCategories( oRs("permitid") )

		If iPermitCount > CLng(0) Then 
%>
			<div id="invoicesummarypagebreak">&nbsp;</div>
<%		End If 

		iPermitCount = iPermitCount + CLng(1)

		ShowInvoiceHeader oRs("permitid")

		response.write "<hr /><br />"
		response.write "<div id=""dateline"">&nbsp; Invoice #: " & oRs("invoiceid")
		response.write " <span id=""invoicedate"">Invoice Date: " & FormatDateTime(oRs("invoicedate"),2) & "</span></div>"
		response.write "<hr />"
		response.write "<div id=""invoicestatusline"">" & UCase(sStatus) & " INVOICE</div>"
		response.write "<hr />"
%>

		<div id="permitlocation">
			PERMIT NUMBER: &nbsp; <strong><% response.write GetPermitNumber( oRs("permitid") )   ' in permitcommonfunctions.asp  %></strong>
			<hr />
<%			ShowPermitLocation oRs("permitid")	%>
		</div>

		<div id="billingcontact">
			ACCOUNT
			<hr />
<%			ShowBillingContact iPermitContactId	%>
		</div>

		<table cellpadding="0" cellspacing="0" border="0" id="invoiceitems" class="tableadmin">
			<tr><th>Fee Cat</th><th>Fee Description</th><th>&nbsp;</th><th>Status</th><th>Amount</th></tr>
<%			ShowInvoiceItems oRs("invoiceid"), bGroupByInvoiceCategories		%>
			<tr><td class="totalline" colspan="3">&nbsp;</td><td class="totalline" align="right"><strong>Invoice Total</strong></td><td class="totalline" align="right"><%=FormatNumber(sInvoiceTotal,2,,,0)%> &nbsp;</td></tr>
			<tr><th class="totalline">&nbsp;</th><th class="totalline"><strong>Payments</strong></th><th class="totalline" colspan="3">&nbsp;</th></tr>
<%			sPaymentTotal = ShowPayments( oRs("invoiceid") )	%>
			<tr><td class="totalline" colspan="3">&nbsp;</td><td class="totalline" align="right"><strong>Payment Total</strong></td><td class="totalline" align="right"> <%=FormatNumber(sPaymentTotal,2,,,0)%> &nbsp;</td></tr>
			<tr><td class="totalline" colspan="3">&nbsp;</td><td class="totalline" align="right"><strong>Balance Due</strong></td><td class="totalline" align="right"><%=FormatNumber((CDbl(sInvoiceTotal) - CDbl(sPaymentTotal)),2,,,0)%> &nbsp;</td></tr>
		</table>

<%		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
%>

	</div>
</div>
<!--END: PAGE CONTENT-->

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowInvoiceHeader_old( )
'--------------------------------------------------------------------------------------------------
Sub ShowInvoiceHeader_old( )

	response.write vbcrlf & "<div id=""invoiceheader"">"

	If OrgHasDisplay( Session("OrgID"), "invoice url" ) Then
		response.write "<img src=""" & GetOrgDisplay( Session("OrgID"), "invoice url" ) & """ border=""0"" />"
	End If 

	If OrgHasDisplay( Session("OrgID"), "invoice header" ) Then
		response.write "<div id=""invoiceheadertext"">" 
		response.write "<h3>Permit Invoice</h3><p>"
		response.write GetOrgDisplay( Session("OrgID"), "invoice header" ) 
		response.write "</p><br /><br />"
		response.write "</div>"
	Else  
		response.write "<h3>" & Session("sOrgName") & " Permit Invoice</h3><br /><br />"
	End If 

	response.write vbcrlf & "</div>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowBillingContact( iBillingId )
'--------------------------------------------------------------------------------------------------
Sub ShowBillingContact( iBillingId )
	Dim sSql, oRs

	sSql = " SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE permitcontactid = " & iBillingId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("firstname") <> "" Then 
			response.write "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			If oRs("firstname") = "" Then 
				response.write "<strong>" & oRs("company") & "</strong><br />" 
			Else 
				response.write oRs("company") & "<br />" 
			End If  
		End If 
		If Trim(oRs("address")) <> "" Then 
			response.write oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			response.write oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			response.write "<br />Phone: " & FormatPhoneNumber( oRs("phone") ) 
		End If
	End If 

	oRs.CLose
	Set oRs = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPermitLocation( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitLocation( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(residentstreetnumber,'') AS residentstreetnumber, ISNULL(residentunit,'') AS residentunit, "
	sSql = sSql & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, "
	sSql = sSql & " ISNULL(streetdirection,'') AS streetdirection, ISNULL(residentcity,'') AS residentcity, "
	sSql = sSql & " ISNULL(residentstate,'') AS residentstate, ISNULL(residentzip,'') AS residentzip, "
	sSql = sSql & " ISNULL(legaldescription,'') AS legaldescription "
	sSql = sSql & " FROM egov_permitaddress WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("residentstreetnumber") <> "" Then
			response.write oRs("residentstreetnumber") & " "
		End If 
		If oRs("residentstreetprefix") <> "" Then
			response.write oRs("residentstreetprefix") & " "
		End If 
		response.write oRs("residentstreetname") & " "
		If oRs("streetsuffix") <> "" Then
			response.write oRs("streetsuffix") & " "
		End If 
		If oRs("streetdirection") <> "" Then
			response.write oRs("streetdirection") & " "
		End If 
		response.write "<br />"
		If oRs("residentunit") <> "" Then
			response.write oRs("residentunit") & "<br />"
		End If 
		If oRs("legaldescription") <> "" Then
			response.write oRs("legaldescription") & "<br />"
		End If 
		If oRs("residentcity") <> "" Then
			response.write oRs("residentcity") & ", "
		End If 
		If oRs("residentstate") <> "" Then
			response.write oRs("residentstate") & " "
		End If 
		If oRs("residentzip") <> "" Then
			response.write oRs("residentzip") 
		End If
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowInvoiceItems( iInvoiceId, bGroupByInvoiceCategories )
'--------------------------------------------------------------------------------------------------
Sub ShowInvoiceItems( iInvoiceId, bGroupByInvoiceCategories )
	Dim sSql, oRs, iPermitFeeCategoryTypeId, bIsInitial, cSubTotal, sCategoryName

	iPermitFeeCategoryTypeId = CLng(0) 
	bIsInitial = True 
	cSubTotal = CDbl(0.00)
	
	sSql = "SELECT I.invoicedamount, I.permitfeeprefix, I.permitfee, I.permitfeecategorytypeid, I.ispercentagetypefee, C.permitfeecategory "
	sSql = sSql & " FROM egov_permitinvoiceitems I, egov_permitfeecategorytypes C "
	sSql = sSql & " WHERE I.invoiceid = " & iInvoiceId
	sSql = sSql & " AND I.permitfeecategorytypeid = C.permitfeecategorytypeid "
	sSql = sSql & " ORDER BY C.displayorder, I.ispercentagetypefee, I.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not bGroupByInvoiceCategories Then 
		Do While Not oRs.EOF
			response.write "<tr>"
			response.write "<td align=""center"">" & oRs("permitfeeprefix") & "</td>"
			response.write "<td class=""feedesccell"">" & oRs("permitfee") & "</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td align=""center"">" & sStatus & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oRs("invoicedamount"),2,,,0) & " &nbsp;</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
	Else
		Do While Not oRs.EOF
			If iPermitFeeCategoryTypeId <> CLng(oRs("permitfeecategorytypeid")) Then
				If Not bIsInitial Then
					' Print out the category total here 
					response.write vbcrlf & "<tr>"
					response.write "<td colspan=""4"" align=""right""><strong>" & sCategoryName & " Total</strong></td>"
					response.write "<td align=""right"" class=""subtotal"">" & FormatNumber(cSubTotal,2,,,0) & " &nbsp;</td>"
					response.write "</tr>"
				Else
					bIsInitial = False 
				End If 
				iPermitFeeCategoryTypeId = CLng(oRs("permitfeecategorytypeid"))
				cSubTotal = CDbl(0.00) 
				sCategoryName = oRs("permitfeecategory")
				response.write vbcrlf & "<tr>"
				response.write "<td colspan=""5""><strong>" & sCategoryName & "</strong></td>"
				response.write "</tr>"
			End If 
			If oRs("ispercentagetypefee") Then
				' Print out the category sub total line
				response.write vbcrlf & "<tr>"
				response.write "<td colspan=""4"" align=""right""><strong>Subtotal</strong></td>"
				response.write "<td align=""right"" class=""subtotal"">" & FormatNumber(cSubTotal,2,,,0) & " &nbsp;</td>"
				response.write "</tr>"
			End If 
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center"">" & oRs("permitfeeprefix") & "</td>"
			response.write "<td class=""feedesccell"">" & oRs("permitfee") & "</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td align=""center"">" & sStatus & "</td>"
			cSubTotal = cSubTotal + CDbl(FormatNumber(oRs("invoicedamount"),2,,,0))
			response.write "<td align=""right"">" & FormatNumber(oRs("invoicedamount"),2,,,0) & " &nbsp;</td>"
			response.write "</tr>"

			oRs.MoveNext 
		Loop 
		' Print out the final category total here 
		response.write vbcrlf & "<tr>"
		response.write "<td colspan=""4"" align=""right""><strong>" & sCategoryName & " Total</strong></td>"
		response.write "<td align=""right"" class=""subtotal"">" & FormatNumber(cSubTotal,2,,,0) & " &nbsp;</td>"
		response.write "</tr>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowInvoiceItems_old( iPaymentId, sInvoiceList )
'--------------------------------------------------------------------------------------------------
Function ShowInvoiceItems_old( ByVal iPaymentId, ByRef sInvoiceList )
	Dim sSql, oRs, dTotal, iOldInvoiceId

	dTotal = CDbl(0.00) 
	iOldInvoiceId = CLng(0)

	sSql = "SELECT P.permitid, I.invoiceid, P.invoicedate, I.invoicedamount, ISNULL(F.permitfeeprefix,'&nbsp;') AS permitfeeprefix, "
	sSql = sSql & " F.permitfee, F.permitfeeid, S.invoicestatus, P.allfeeswaived, P.isvoided "
	sSql = sSql & " FROM egov_permitinvoiceitems I, egov_permitfees F, egov_permitinvoices P, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.permitfeeid = F.permitfeeid AND P.paymentid = " & iPaymentId 
	sSql = sSql & " AND P.invoicestatusid = S.invoicestatusid AND I.invoiceid = P.invoiceid " 
	sSql = sSql & " ORDER BY I.invoiceid, F.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write "<tr>"
		response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"
		If iOldInvoiceId <> CLng(oRs("invoiceid")) Then 
			If sInvoiceList <> "" Then 
				sInvoiceList = sInvoiceList & ", "
			End If 
			sInvoiceList = sInvoiceList & oRs("invoiceid")
			iOldInvoiceId = CLng(oRs("invoiceid"))
		End If 
		response.write "<td align=""center"">" & FormatDateTime(oRs("invoicedate"),2) & "</td>"
		response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber(oRs("permitid")) & "</td>"
		response.write "<td align=""center"">"
		If oRs("permitfeeprefix") <> "" Then
			response.write oRs("permitfeeprefix")
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td class=""feedesccell"">" & oRs("permitfee") & "</td>"
		response.write "<td align=""center"">" & oRs("invoicestatus") & "</td>"
		response.write "<td align=""right"">" & FormatNumber(oRs("invoicedamount"),2,,,0) & " &nbsp;</td>"
		If Not oRs("allfeeswaived") And Not oRs("isvoided") Then 
			dTotal = dTotal + CDbl(oRs("invoicedamount"))
		End If 
		response.write "</tr>"
		oRs.MoveNext 
	Loop 
	
	oRs.Close
	Set oRs = Nothing 
	ShowInvoiceItems = dTotal
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetInvoiceFeeStatus( iInvoiceId, iPermitFeeId )
'--------------------------------------------------------------------------------------------------
Function GetInvoiceFeeStatus( iInvoiceId, iPermitFeeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(invoiceitemid) AS hits FROM egov_permitinvoiceitems "
	sSql = sSql & " WHERE invoiceid = " & iInvoiceId & " AND permitfeeid = " & iPermitFeeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			GetInvoiceFeeStatus = "Paid"
		Else
			GetInvoiceFeeStatus = "Due"
		End If 
	Else
		GetInvoiceFeeStatus = "Due"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function ShowPayments( iInvoiceId )
'--------------------------------------------------------------------------------------------------
Function ShowPayments( iInvoiceId )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND L.invoiceid = " & iInvoiceId
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td align=""center"" valign=""top"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
		response.write "<td class=""feedesccell"">Payment #: " & oRs("paymentid") '& " &ndash; " & oRs("paymenttypename")
'		sCheckNo = GetCheckNo( oRs("paymentid") )
'		If sCheckNo <> "" Then 
'			response.write "&nbsp; &nbsp; Check #: " & sCheckNo
'		End If 

		' Show payment types and amount
		ShowInvoicePayments oRs("paymentid")

		response.write "</td>"
		response.write "<td>&nbsp;</td>"
		response.write "<td>&nbsp;</td>"
		response.write "<td align=""right"" valign=""top"">" & FormatNumber(oRs("paymenttotal"),2,,,0) & " &nbsp;</td>"
		response.write "</tr>"
		dTotal = dTotal + CDbl(oRs("paymenttotal"))
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	ShowPayments = FormatNumber(dTotal,2,,,0)
End Function 


'--------------------------------------------------------------------------------------------------
' Function ShowPayments_old( sInvoiceList )
'--------------------------------------------------------------------------------------------------
Function ShowPayments_old( sInvoiceList )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND L.invoiceid IN ( " & sInvoiceList & " ) "
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td align=""center"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
		response.write "<td>&nbsp;</td>"
		response.write "<td class=""feedesccell"" colspan=""2"">Payment #: " & oRs("paymentid") '& " &ndash; " & oRs("paymenttypename")
		sCheckNo = GetCheckNo( oRs("paymentid") )
		If sCheckNo <> "" Then 
			response.write "&nbsp; &nbsp; Check #: " & sCheckNo
		End If 
		response.write "</td>"
		response.write "<td>&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(oRs("paymenttotal"),2,,,0) & " &nbsp;</td>"
		response.write "</tr>"
		dTotal = dTotal + CDbl(oRs("paymenttotal"))
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	ShowPayments = FormatNumber(dTotal,2,,,0)
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowInvoicePayments( iPaymentId )
'--------------------------------------------------------------------------------------------------
Sub ShowInvoicePayments( iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(L.amount,0.00) AS amount, P.paymenttypename, P.requirescheckno "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_paymenttypes P "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND J.paymentid = " & iPaymentId
	sSql = sSql & " AND L.entrytype = 'debit' AND L.paymenttypeid = P.paymenttypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<br />"
		response.write " &nbsp; &nbsp; " & oRs("paymenttypename") 
		If oRs("requirescheckno") Then 
			response.write " #: " & GetCheckNo( iPaymentId )
		End If 
		response.write " for " & FormatCurrency(oRs("amount"),2)
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetPermitContactForPayment( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function GetPermitContactForPayment( iPaymentId )
Dim sSql, oRs

	' This will pull one or more contactid, they will all be the same contractor
	sSql = "SELECT permitcontactid FROM egov_permitinvoices WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitContactForPayment = CLng(oRs("permitcontactid"))
	Else
		GetPermitContactForPayment = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 



%>
