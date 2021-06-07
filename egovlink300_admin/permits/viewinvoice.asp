<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewinvoice.asp
' AUTHOR: Steve Loar
' CREATED: 05/19/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays one invoice for a permit.
'
' MODIFICATION HISTORY
' 1.0   05/19/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iInvoiceId, sInvoiceDate, sInvoiceTotal, iBillingId, iPermitId, sPaymentTotal, sStatus, bIsWaived
Dim bIsVoided, bGroupByInvoiceCategories, bIsDue

sLevel = "../" ' Override of value from common.asp

'PageDisplayCheck "edit permits", sLevel	' In common.asp

iInvoiceId = CLng(request("invoiceid"))

sInvoiceDate = ""
sInvoiceTotal = ""
iBillingId = 0
iPermitId = 0
sPaymentTotal = 0.00
sStatus = ""
bIsWaived = False 
bIsVoided = False 

GetInvoiceValues iInvoiceId, sInvoiceDate, sInvoiceTotal, iBillingId, iPermitId, sStatus, bIsWaived, bIsVoided, bIsDue

bGroupByInvoiceCategories = GetIsGroupByInvoiceCategories( iPermitId )

%>

<html>
<head>
	<title>E-Gov Permit Invoice</title>

	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="Javascript">
	<!--
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.width = "70%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.height = "90%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.left = "15%";
		parent.document.getElementById("modal"+window.frameElement.getAttribute("data-close")).style.top = "5%";

		function doClose()
		{
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
	<br />
	<br />
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
<%					
	tooltipclass=""
	tooltip = ""
	disabled = ""
	If not bIsDue Then
		tooltipclass="tooltip"
		disabled = " disabled "
		tooltip = "<span class=""tooltiptext"">No Invoice is due.</span>"
	end if %>
	<button <%=disabled %> type="button" class="button ui-button ui-widget ui-corner-all <%=tooltipclass%>" onclick="PayInvoices();">Pay Invoices<%=tooltip%></button>&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
<%
	ShowInvoiceHeader iPermitId

	response.write "<hr />"
	response.write "<div id=""dateline"">&nbsp; Invoice #: " & iInvoiceId 
	If bIsVoided Then
		response.write " &ndash; VOID"
	End If 
	response.write " <span id=""invoicedate"">Invoice Date: " & sInvoiceDate & "</span></div>"
	response.write "<hr />"
	response.write "<div id=""invoicestatusline"">" & UCase(sStatus) & " INVOICE</div>"
	response.write "<hr />"
%>
	<div id="permitlocation">
		PERMIT NUMBER: &nbsp; <strong><% = GetPermitNumber( iPermitId ) %></strong>
		<hr />
<%		ShowPermitLocation iPermitId	%>
	</div>

	<div id="billingcontact">
		ACCOUNT
		<hr />
<%		ShowBillingContact iBillingId	%>
	</div>

	<table cellpadding="2" cellspacing="0" border="0" id="invoiceitems" class="tableadmin">
		<tr><th>Fee Cat</th><th>Fee Description</th><th>&nbsp;</th><th>Status</th><th>Amount</th></tr>
<%		ShowInvoiceItems iInvoiceId, bGroupByInvoiceCategories		%>
		<tr><td class="totalline" colspan="3">&nbsp;</td><td class="totalline" align="right"><strong>Invoice Total</strong></td><td class="totalline" align="right"><%=sInvoiceTotal%> &nbsp;</td></tr>
		<tr><th class="totalline">&nbsp;</th><th class="totalline"><strong>Payments</strong></th><th class="totalline" colspan="3">&nbsp;</th></tr>
<%		sPaymentTotal = ShowPayments( iInvoiceId )		%>
		<tr><td class="totalline" colspan="3">&nbsp;</td><td class="totalline" align="right"><strong>Payment Total</strong></td><td class="totalline" align="right"> <%=sPaymentTotal%> &nbsp;</td></tr>
		<tr><td class="totalline" colspan="3">&nbsp;</td><td class="totalline" align="right"><strong>Balance Due</strong></td><td class="totalline" align="right">
			<%
			If bIsVoided Then 
				response.write "0.00"
			Else 
				response.write FormatNumber((CDbl(sInvoiceTotal) - CDbl(sPaymentTotal)),2,,,0)
			End If 
			%>
			&nbsp;</td>
		</tr>
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
' Sub GetInvoiceValues( iInvoiceId, sInvoiceDate, sInvoiceTotal, iBillingId, iPermitId, sStatus, bIsWaived, bIsVoided, bIsDue )
'--------------------------------------------------------------------------------------------------
Sub GetInvoiceValues( ByVal iInvoiceId, ByRef sInvoiceDate, ByRef sInvoiceTotal, ByRef iBillingId, ByRef iPermitId, ByRef sStatus, ByRef bIsWaived, ByRef bIsVoided, ByRef bIsDue )
	Dim oRs, sSql 

	sSql = "SELECT I.permitid, I.invoicedate, I.totalamount, I.permitcontactid, S.invoicestatus, S.iswaived, S.isvoid, S.isdue "
	sSql = sSql & " FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.orgid = " & session("orgid")
	sSql = sSql & " AND I.invoiceid = " & iInvoiceId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sInvoiceDate = DateValue(CDate(oRs("invoicedate")))
		'sInvoiceTotal = FormatNumber(oRs("totalamount"),2,,,0)
		iBillingId = CLng(oRs("permitcontactid"))
		iPermitId = CLng(oRs("permitid"))
		sStatus = oRs("invoicestatus")
		If oRs("iswaived") Then
			bIsWaived = True 
			sInvoiceTotal = FormatNumber(0.00,2,,,0)
		Else
			bIsWaived = False 
			sInvoiceTotal = FormatNumber(oRs("totalamount"),2,,,0)
		End If 
		If oRs("isvoid") Then
			bIsVoided = True 
		Else
			bIsVoided = False 
		End If 
		If oRs("isdue") Then
			bIsDue = True 
		Else
			bIsDue = False 
		End If 
	Else
		sInvoiceDate = ""
		sInvoiceTotal = "0.00"
		iBillingId = 0
		iPermitId = 0
		sStatus = ""
		bIsWaived = False 
		bIsVoided = False
		bIsDue = False 
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
	sSql = sSql & " ORDER BY C.displayorder, I.ispercentagetypefee, I.displayorder "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not bGroupByInvoiceCategories Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
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
' Function ShowPayments( iInvoiceId )
'--------------------------------------------------------------------------------------------------
Function ShowPayments( iInvoiceId )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_permitinvoices I "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND I.invoiceid = L.invoiceid AND I.isvoided = 0 AND L.invoiceid = " & iInvoiceId
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td align=""center"" valign=""top"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
		response.write "<td class=""feedesccell"">Payment #: " & oRs("paymentid") '& " &ndash; " & oRs("paymenttypename")
		'sCheckNo = GetCheckNo( oRs("paymentid") )
		'If sCheckNo <> "" Then 
		'	response.write "&nbsp; &nbsp; Check #: " & sCheckNo
		'End If 
		
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
