<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
' 1.1	05/19/08 	Steve Loar - PageDisplayCheck added
' 1.2	01/08/09	David Boyer - Added "DisplayRosterPublic" fields for Craig,CO custom team registration
' 1.3	03/09/09	Steve Loar - Changes for Regatta Teams
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iPaymentId, iUid, sReceiptType, sJEType, iDisplayType, sNotes, iAdminuserid, iPriorPaymentId
Dim iMerchandiseOrderId, iuserid, nTotal, dPaymentDate, sSql, nRowTotal, iJournalEntryTypeId, sJournalEntryType
Dim dMerchandiseAmount, dShippingAndHandling, dSalesTax, dTotalAmount, dTotal

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "registration", sLevel	' In common.asp

iMerchandiseOrderId = CLng(request("orderid"))

%>
<html>
<head>
	<title>E-Gov Administration Console {Merchandise Order}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="merchandise.css" />
	<link rel="stylesheet" type="text/css" href="../classes/receiptprint.css" media="print" />

	<script language="javascript">
	<!--

		window.onload = function()
		{
			//factory.printing.header = "Printed on &d"
			factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
			factory.printing.portrait     = true;
			factory.printing.leftMargin   = 0.5;
			factory.printing.topMargin    = 0.5;
			factory.printing.rightMargin  = 0.5;
			factory.printing.bottomMargin = 0.5;

			// enable control buttons
			var templateSupported = factory.printing.IsTemplateSupported();
			var controls = idControls.all.tags("input");
			for ( i = 0; i < controls.length; i++ ) 
			{
				controls[i].disabled = false;
				if (templateSupported && controls[i].className == "ie55" )
					controls[i].style.display = "inline";
			}
		}

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	 <input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
	 <input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
</div>

<object id="factory" viewastext  style="display:none"
   classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
   codebase="../includes/smsx.cab#Version=6,3,434,12">
</object>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">

<%

	GetOrderInformation iMerchandiseOrderId, iPaymentId, dPaymentDate, iUserId, dMerchandiseAmount, dShippingAndHandling, dSalesTax

	dTotal = CDbl(dMerchandiseAmount) + CDbl(dShippingAndHandling) + CDbl(dSalesTax)
	dTotal = FormatNumber(dTotal,2)

	ShowReceiptHeader

	response.write vbcrlf & "<hr />"
	response.write " Date: " & DateValue(CDate(dPaymentDate)) & " &nbsp; &nbsp; &nbsp; &nbsp; "
	response.write " Receipt #: " & iPaymentId & " &nbsp; &nbsp; &nbsp; &nbsp; " 
	response.write " Order #: " & iMerchandiseOrderId
	response.write "<hr />" & vbcrlf

	' Show Ship TO Information
	response.write "<div id=""receipttopright"">" & vbcrlf
	response.write "<span class=""receipttitles"">Shipping Information</span><br />"
	ShowShippingLabel( iMerchandiseOrderId )
	response.write "</div>" & vbcrlf

	' Payee Information
	response.write ShowUserInfo( iUserId )
	'response.write "<hr />" & vbcrlf

	' List Merchandise items here
	response.write "<hr />" & vbcrlf
	response.write "<strong>Merchandise Items</strong>" & vbcrlf
	response.write "<hr />" & vbcrlf
	ShowMerchandiseItems( iMerchandiseOrderId )
	'response.write "<hr />" & vbcrlf

	' Show merchandise total
	response.write "<hr />" & vbcrlf
	response.write "<span id=""merchandisetotal"">" & dMerchandiseAmount & "</span>"
	response.write "<strong>Merchandise Total</strong>" & vbcrlf

	' Show Shipping and Handling
	response.write "<hr />" & vbcrlf
	response.write "<span id=""shippingtotal"">" & dShippingAndHandling & "</span>"
	response.write "<strong>Shipping And Handling</strong>" & vbcrlf

	' Show Sales Tax
	response.write "<hr />" & vbcrlf
	response.write "<span id=""salestax"">" & dSalesTax & "</span>"
	response.write "<strong>Sales Tax</strong>" & vbcrlf

	' Show Total
	response.write "<hr />" & vbcrlf
	response.write "<span id=""total"">" & dTotal & "</span>"
	response.write "<strong>Total</strong>" & vbcrlf
	response.write "<hr />" & vbcrlf

%>

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
' Sub GetOrderInformation( iMerchandiseOrderId, iPaymentId, dPaymentDate, iUserId, dMerchandiseAmount, dShippingAndHandling, dSalesTax )
'--------------------------------------------------------------------------------------------------
Sub GetOrderInformation( ByVal iMerchandiseOrderId, ByRef iPaymentId, ByRef dPaymentDate, ByRef iUserId, ByRef dMerchandiseAmount, ByRef dShippingAndHandling, ByRef dSalesTax )
	Dim sSql, oRs
	
	sSql = "SELECT userid, paymentid, orderdate, ISNULL(orderamount,0.00) AS orderamount, ISNULL(shippingfee,0.00) AS shippingfee, ISNULL(taxamount,0.00) AS taxamount "
	sSql = sSql & " FROM egov_merchandiseorders WHERE merchandiseorderid = " & iMerchandiseOrderId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iPaymentId = oRs("paymentid")
		dPaymentDate = FormatDateTime(oRs("orderdate"),2)
		iUserId = oRs("userid")
		dMerchandiseAmount = FormatNumber(oRs("orderamount"),2)
		dShippingAndHandling = FormatNumber(oRs("shippingfee"),2)
		dSalesTax = FormatNumber(oRs("taxamount"),2)
	Else 
		iPaymentId = 0
		dPaymentDate = ""
		iUserId = 0
		dMerchandiseAmount = 0.00
		dShippingAndHandling = 0.00
		dSalesTax = 0.00
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowReceiptHeader( )
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptHeader( )

	If OrgHasDisplay( Session("orgid"), "receipt header" ) Then
		response.write "<p class=""receiptheader"">" & GetOrgDisplay( Session("orgid"), "receipt header" ) 
		response.write "<br /><br />"
		response.write "Merchandise Purchase</p>"
	Else  
		response.write "<h3>" & Session("sOrgName") & " Merchandise Purchase</h3><br /><br />"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowUserInfo( iUserId )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( iUserId )
	Dim oCmd, oUser

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iuserid", 3, 1, 4, iUserId)
	    Set oUser = .Execute
	End With
	
	ShowUserInfo = ShowUserInfo & "<span class=""receipttitles"">"
	ShowUserInfo = ShowUserInfo & "Payee Information</span><br />"

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
	ShowUserInfo = ShowUserInfo & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhone(oUser("userhomephone")) & "</td></tr>"
	ShowUserInfo = ShowUserInfo & "</table>"

	oUser.Close
	Set oUser = Nothing
	Set oCmd = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowShippingLabel( iMerchandiseOrderId )
'--------------------------------------------------------------------------------------------------
Sub ShowShippingLabel( iMerchandiseOrderId )
	Dim sSql, oRs
	
	sSql = "SELECT shiptoname, shiptoaddress, shiptocity, shiptostate, shiptozip "
	sSql = sSql & " FROM egov_merchandiseorders WHERE merchandiseorderid = " & iMerchandiseOrderId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""shippinginfo"">"
		response.write "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & oRs("shiptoname") & "</strong><br />"
		response.write "<strong>" & oRs("shiptoaddress") & "<br />"
		response.write oRs("shiptocity") & ", " & oRs("shiptostate") & " " & oRs("shiptozip") & "</strong></td></tr>"
		response.write "</table>"
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowMerchandiseItems( iMerchandiseOrderId )
'--------------------------------------------------------------------------------------------------
Sub ShowMerchandiseItems( iMerchandiseOrderId )
	Dim sSql, oRs, dItemTotal
	
	sSql = "SELECT merchandise, merchandisecolor, merchandisesize, quantity, isnocolor, isnosize, itemprice "
	sSql = sSql & " FROM egov_merchandiseorderitems WHERE merchandiseorderid = " & iMerchandiseOrderId
	sSql = sSql & " AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY merchandise, merchandisecolor, displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" id=""viewmerchandiseitems"">"
		response.write vbcrlf & "<tr class=""receiptmerchandiseitemheader""><th align=""left"">Merchandise Item</th>"
		response.write "<th align=""center"">Color</th><th align=""center"">Size</th><th align=""center"">Price<br />Each</th>"
		response.write "<th align=""center"">Quantity</th><th align=""center"">Item<br />Total</th></tr>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
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
			response.write "<td align=""right"">" & FormatNumber(dItemTotal,2) & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Sub



%>