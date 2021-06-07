<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewinvoicesummary.asp
' AUTHOR: Steve Loar
' CREATED: 05/23/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the summary of invoices and payments for a permit 
'
' MODIFICATION HISTORY
' 1.0   05/23/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iInvoiceId, sInvoiceDate, sInvoiceTotal, iBillingId, iPermitId, sPaymentTotal, iPermitContactId
Dim sInvoiceList, bIsWaived, bIsVoided, bGroupByInvoiceCategories, sStatus, bIsDue

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "edit permits", sLevel	' In common.asp

iPermitId = CLng(request("permitid"))
iPermitContactId = CLng(request("permitcontactid"))

sInvoiceDate = ""
sInvoiceTotal = CDbl(0.00)
iBillingId = iPermitContactId
sPaymentTotal = CDbl(0.00 )
sInvoiceList = ""
bIsWaived = False 
bIsVoided = False 

sStatus = GetPermitInvoicesStatus( iPermitId, bIsDue )

bGroupByInvoiceCategories = GetIsGroupByInvoiceCategories( iPermitId )

%>

<html>
<head>
	<title>E-Gov Permit Invoice</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="Javascript">
	<!--

		function doClose()
		{
			//window.close();
			//window.opener.focus();
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

		function PayInvoices()
		{
			location.href = 'paypermitinvoices.asp?permitid=<%=iPermitId%>';
		}

	//-->
	</script>

</head>

<body>
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
<%	If bIsDue Then		%>
		<input type="button" class="button ui-button ui-widget ui-corner-all" value="Pay Invoices" onclick="PayInvoices();" />&nbsp;&nbsp;
<%	End If	%>
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
<%
	ShowInvoiceHeader iPermitId

	response.write "<hr /><br />"
	'response.write "<div id=""dateline"">&nbsp; Invoice #: " & iInvoiceId 
	'response.write " <span id=""invoicedate"">Invoice Date: " & sInvoiceDate & "</span></div>"
	response.write "<div id=""invoicestatusline"">" & UCase(sStatus) & "</div>"
	response.write "<hr />"
%>
	<div id="permitlocation">
		PERMIT NUMBER: &nbsp; <strong><% response.write GetPermitNumber( iPermitId )   ' in permitcommonfunctions.asp  %></strong>
		<hr />
<%		ShowPermitLocation iPermitId	%>
	</div>

	<div id="billingcontact">
		ACCOUNT
		<hr />
<%		ShowBillingContact iPermitContactId	%>
	</div>

	<table cellpadding="0" cellspacing="0" border="0" id="invoiceitems" class="tableadmin">
		<tr><th>Invoice #</th><th>Invoice Date</th><th>Fee Cat</th><th>Fee Description</th><th>Status</th><th>Amount</th></tr>
<%		sInvoiceTotal = ShowInvoiceItems( iPermitId, iPermitContactId, sInvoiceList, bGroupByInvoiceCategories )		%>
		<tr><td class="totalline" colspan="5" align="right"><strong>Invoice Total</strong></td><td class="totalline" align="right"><%=FormatNumber(sInvoiceTotal,2,,,0)%> &nbsp;</td></tr>
		<tr><th class="totalline" colspan="3">&nbsp;</th><th class="totalline"><strong>Payments</strong></th><th class="totalline">&nbsp;</th><th class="totalline">&nbsp;</th></tr>
<%		
		If sInvoiceList <> "" Then 
			sPaymentTotal = ShowPayments( sInvoiceList )		
		End If 
%>
		<tr><td class="totalline" colspan="5" align="right"><strong>Payment Total</strong></td><td class="totalline" align="right"><%=FormatNumber(sPaymentTotal,2,,,0)%> &nbsp;</td></tr>
		<tr><td class="totalline" colspan="5" align="right"><strong>Balance Due</strong></td><td class="totalline" align="right"><%= FormatNumber((CDbl(sInvoiceTotal) - CDbl(sPaymentTotal)),2,,,0)%> &nbsp;</td></tr>
	</table>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowInvoiceHeader_old( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowInvoiceHeader_old( iPermitId )

	response.write vbcrlf & "<div id=""invoiceheader"">"

	If OrgHasDisplay( Session("OrgID"), "invoice url" ) Then
		response.write "<img src=""" & GetOrgDisplay( Session("OrgID"), "invoice url" ) & """ border=""0"" />"
	End If 

	If OrgHasDisplay( Session("OrgID"), "invoice header" ) Then
		response.write "<div id=""invoiceheadertext"">" 
		response.write "<h3>Permit Invoice Summary</h3><p>"
		response.write GetOrgDisplay( Session("OrgID"), "invoice header" ) 
		response.write "</p><br /><br />"
		response.write "</div>"
	Else  
		response.write "<h3>" & Session("sOrgName") & " Permit Invoice Summary</h3><br /><br />"
	End If 

	response.write vbcrlf & "</div>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetInvoiceValues( iInvoiceId, sInvoiceDate, sInvoiceTotal, iBillingId, iPermitId )
'--------------------------------------------------------------------------------------------------
Sub GetInvoiceValues( ByVal iInvoiceId, ByRef sInvoiceDate, ByRef sInvoiceTotal, ByRef iBillingId, ByRef iPermitId )
	Dim oRs, sSql 

	sSql = "SELECT permitid, invoicedate, totalamount, permitcontactid FROM egov_permitinvoices WHERE orgid = " & session("orgid")
	sSql = sSql & " AND invoiceid = " & iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sInvoiceDate = DateValue(CDate(oRs("invoicedate")))
		sInvoiceTotal = FormatNumber(oRs("totalamount"),2,,,0)
		iBillingId = CLng(oRs("permitcontactid"))
		iPermitId = CLng(oRs("permitid"))
	Else
		sInvoiceDate = ""
		sInvoiceTotal = "0.00"
		iBillingId = 0
		iPermitId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
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
' Function ShowInvoiceItems( iPermitId, iPermitContactId, sInvoiceList )
'--------------------------------------------------------------------------------------------------
Function ShowInvoiceItems( ByVal iPermitId, ByVal iPermitContactId, ByRef sInvoiceList, ByVal bGroupByInvoiceCategories )
	Dim sSql, oRs, dTotal, iOldInvoiceId, iPermitFeeCategoryTypeId, bIsInitial, cSubTotal, sCategoryName

	dTotal = CDbl(0.00) 
	iOldInvoiceId = CLng(0)
	iPermitFeeCategoryTypeId = CLng(0) 
	bIsInitial = True 
	cSubTotal = CDbl(0.00) 

	sSql = "SELECT II.invoiceid, P.invoicedate, II.invoicedamount, ISNULL(II.permitfeeprefix,'&nbsp;') AS permitfeeprefix, "
	sSql = sSql & " II.permitfee, II.permitfeeid, S.invoicestatus, P.allfeeswaived, P.isvoided, II.permitfeecategorytypeid, "
	sSql = sSql & " II.ispercentagetypefee, C.permitfeecategory "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitinvoices P, egov_invoicestatuses S, egov_permitfeecategorytypes C, egov_permitinvoices I "
	sSql = sSql & " WHERE P.permitid = " & iPermitId 
	sSql = sSql & " AND P.invoicestatusid = S.invoicestatusid AND II.invoiceid = P.invoiceid AND P.permitcontactid = " & iPermitContactId 
	sSql = sSql & " AND II.permitfeecategorytypeid = C.permitfeecategorytypeid AND II.invoiceid = I.invoiceid AND I.isvoided = 0 "
	sSql = sSql & " ORDER BY C.displayorder, II.ispercentagetypefee, II.invoiceid, II.displayorder"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not bGroupByInvoiceCategories Then 
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
	Else
		Do While Not oRs.EOF
			If iPermitFeeCategoryTypeId <> CLng(oRs("permitfeecategorytypeid")) Then
				If Not bIsInitial Then
					' Print out the category total here 
					response.write vbcrlf & "<tr>"
					response.write "<td colspan=""5"" align=""right""><strong>" & sCategoryName & " Total</strong></td>"
					response.write "<td align=""right"" class=""subtotal"">" & FormatNumber(cSubTotal,2,,,0) & " &nbsp;</td>"
					response.write "</tr>"
				Else
					bIsInitial = False 
				End If 
				iPermitFeeCategoryTypeId = CLng(oRs("permitfeecategorytypeid"))
				cSubTotal = CDbl(0.00) 
				sCategoryName = oRs("permitfeecategory")
				response.write vbcrlf & "<tr>"
				response.write "<td colspan=""6""><strong>" & sCategoryName & "</strong></td>"
				response.write "</tr>"
			End If 
			If oRs("ispercentagetypefee") Then
				' Print out the category sub total line
				response.write vbcrlf & "<tr>"
				response.write "<td colspan=""5"" align=""right""><strong>Subtotal</strong></td>"
				response.write "<td align=""right"" class=""subtotal"">" & FormatNumber(cSubTotal,2,,,0) & " &nbsp;</td>"
				response.write "</tr>"
			End If 
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"
			If iOldInvoiceId <> CLng(oRs("invoiceid")) Then 
				If sInvoiceList <> "" Then 
					sInvoiceList = sInvoiceList & ", "
				End If 
				sInvoiceList = sInvoiceList & oRs("invoiceid")
				iOldInvoiceId = CLng(oRs("invoiceid"))
			End If 
			response.write "<td align=""center"">" & FormatDateTime(oRs("invoicedate"),2) & "</td>"
			response.write "<td align=""center"">" & oRs("permitfeeprefix") & "</td>"
			response.write "<td class=""feedesccell"">" & oRs("permitfee") & "</td>"
			'response.write "<td>&nbsp;</td>"
			response.write "<td align=""center"">" & oRs("invoicestatus") & "</td>"
			cSubTotal = cSubTotal + CDbl(FormatNumber(oRs("invoicedamount"),2,,,0))
			response.write "<td align=""right"">" & FormatNumber(oRs("invoicedamount"),2,,,0) & " &nbsp;</td>"
			If Not oRs("allfeeswaived") And Not oRs("isvoided") Then 
				dTotal = dTotal + CDbl(oRs("invoicedamount"))
			End If 
			response.write "</tr>"

			oRs.MoveNext 
		Loop 
		' Print out the final category total here 
		response.write vbcrlf & "<tr>"
		response.write "<td colspan=""5"" align=""right""><strong>" & sCategoryName & " Total</strong></td>"
		response.write "<td align=""right"" class=""subtotal"">" & FormatNumber(cSubTotal,2,,,0) & " &nbsp;</td>"
		response.write "</tr>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
	ShowInvoiceItems = dTotal
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetInvoiceFeeStatus( iInvoiceId, iPermitFeeId ) - NOt called from anywhere
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
' money ShowPayments( sInvoiceList )
'--------------------------------------------------------------------------------------------------
Function ShowPayments( ByVal sInvoiceList )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_permitinvoices I "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND I.invoiceid = L.invoiceid AND I.isvoided = 0 AND L.invoiceid IN ( " & sInvoiceList & " ) "
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"
	'response.write sSql
	session("showpaymentsSQL") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	session("showpaymentsSQL") = ""

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td align=""center"" valign=""top"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
		response.write "<td>&nbsp;</td>"
		response.write "<td class=""feedesccell"">Payment #: " & oRs("paymentid") '& " &ndash; " & oRs("paymenttypename")
'		sCheckNo = GetCheckNo( oRs("paymentid") )
'		If sCheckNo <> "" Then 
'			response.write "&nbsp; &nbsp; Check #: " & sCheckNo
'		End If 

		' Show payment types and amount
		ShowInvoicePayments oRs("paymentid")

		response.write "</td>"
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




%>
