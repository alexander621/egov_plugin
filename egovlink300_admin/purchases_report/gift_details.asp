<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: gift_details.asp
' AUTHOR: John Stullenberger
' CREATED: 08/01/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Allows the viewing of gift purchase details
'				No Security check as this is called from multiple places
'
' MODIFICATION HISTORY
' 1.0   02/14/2006	JOHN STULLENBERGER - INITIAL VERSION
' 1.1	08/01/2006	Steve Loar - Made the original receipt into this view details script
' 2.0	07/30/2010	Steve Loar - Changed made for Point and Pay payments
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iGiftPaymentId 

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "citizen rec purchases" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iGiftPaymentId = CLng(request("igiftpaymentid"))

%>

<html>
<head>
	<title<%=Session("sOrgName")%> Commemorative Gift Purchase</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />
	<script language="javascript">
	<!--

		window.onload = function()
		{
			factory.printing.header = "Commemorative Gift Purchase - Printed on &d"
			factory.printing.footer = "&bCommemorative Gift Purchase - Printed on &d - Page:&p/&P"
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

<!--BEGIN: PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<div id="receiptlinks">
		<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)""><%=langBackToStart%></a><span id="printbutton"><input type="button" class="button" onclick="javascript:window.print();" value="Print" /></span>
	</div>

	<h3><%=Session("sOrgName")%> Commemorative Gift Purchase Details</h3>

	<% DisplayReciept iGiftPaymentId %>

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
' void DisplayReciept iGiftPaymentId
'--------------------------------------------------------------------------------------------------
Sub DisplayReciept( ByVal iGiftPaymentId )
	Dim sSql, oRs 

	' Get payment information 
	sSql = "SELECT giftamount, paymenttype, paymentlocation, giftname, firstname, lastname, address1, "
	sSql = sSql & " city, state, zip, phone, email, paymentdate, ISNULL(processingfee,0.00) AS processingfee, "
	sSql = sSql & " ISNULL(sva,'') AS sva, ISNULL(P.ordernumber,'') AS ordernumber "
	sSql = sSql & " FROM egov_gift_payment P, egov_gift G"
	sSql = sSql & " WHERE P.giftid = G.giftid AND P.giftpaymentid = " & iGiftPaymentId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		' Show the user info
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write vbcrlf & "<tr><th colspan=""2"" align=""left"">Purchaser Contact Information</th></tr>"
		response.write vbcrlf & "<tr><td width=""20%"" valign=""top"">Name:</td><td>" & oRs("firstname") & " " & oRs("lastname")
		response.write "</td></tr>"
		response.write vbcrlf & "<tr><td>Email:</td><td>" & oRs("email") & "</td></tr>"
		response.write vbcrlf & "<tr><td>Phone:</td><td>" & FormatPhone(oRs("phone")) & "</td></tr>"
		response.write vbcrlf & "<tr><td valign=""top"">Address:</td><td>" & oRs("address1") & "<br />" 
		response.write oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "</td></tr>"
		response.write vbcrlf & "</table></div>"


		' TRANSACTION RESULT DETAILS
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write "<tr><th colspan=""2"" align=""left""><b>Purchase Details</b></th></tr>"
		response.write "<tr><td width=""20%"">Purchase Date: </td><td> " & oRs("paymentdate") & "</td></tr>"
		
		response.write "<tr><td>Payment Method:</td><td> " & GetPaymentTypeName(oRs("paymenttype")) & " </td></tr>"
		response.write "<tr><td>Payment Location:</td><td> " & GetPaymentLocationName(oRs("paymentlocation")) & " </td></tr>"
		response.write "<tr><td>Amount: </td><td> " & FormatCurrency(oRs("giftamount"),2) & "</td></tr>"
		If oRs("sva") <> "" Then
			response.write "<tr><td>Processing Fee:</td><td> " & FormatCurrency(oRs("processingfee"),2) & "</td></tr>"
			response.write "<tr><td>Total Charges:</td><td> " & FormatCurrency( (CDbl(oRs("giftamount")) + CDbl(oRs("processingfee"))),2) & "</td></tr>"
			response.write "<tr><td>Order Number:</td><td> " & oRs("ordernumber") & "</td></tr>"
			response.write "<tr><td>SVA:</td><td> " & oRs("sva") & "</td></tr>"
		End If 
		response.write "</table></div>"
		
		' PRODUCT INFORMATION
		response.write vbcrlf & "<div class=""purchasereportshadow"">"
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" class=""purchasereport"">"
		response.write "<tr><th colspan=""2"" align=""left""><b>Gift Information</b></th></tr>"
		response.write "<tr><td width=""20%"">Order Number:</td><td> " & iGiftPaymentId & "F3000 </td></tr>"
		response.write "<tr><td>Product:</td><td>" & oRs("giftname") & "</td></tr>"
		response.write "<tr><td valign=top>Details: </td><td  valign=""top"">" & GetFieldValues( iGiftPaymentId ) &  "</td></tr>"
		response.write "</table></div>"
	Else
		response.write "<p>No details are available.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub

'------------------------------------------------------------------------------------------------------------
' string GetFieldValues( iGiftPaymentID )
'------------------------------------------------------------------------------------------------------------
Function GetFieldValues( ByVal iGiftPaymentID )
	Dim sSql, oRs, sReturnValue

	sReturnValue = ""
		
	sSql = "SELECT V.giftvalue, V.giftpaymentid, V.fieldid, F.fieldprompt "
	sSql = sSql & " FROM egov_gift_value V, egov_gift_fields F "
	sSql = sSql & " WHERE V.fieldid = F.fieldid AND V.giftpaymentid = " & iGiftPaymentID
	sSql = sSql & " ORDER BY V.giftpaymentid, V.fieldid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While NOT oRs.EOF 
		sReturnValue = sReturnValue & "<strong>" & oRs("fieldprompt") & "</strong> : <i>" & oRs("giftvalue") & "</i><br />" 			
		oRs.MoveNext
	Loop

	oRs.Close
	Set oRs = Nothing 

	GetFieldValues = sReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' string GetPaymentTypeName( iPaymentTypeId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentTypeName( ByVal iPaymentTypeId)
	Dim sSql, oRs

	sSql = "SELECT paymenttypename FROM egov_paymenttypes WHERE paymenttypeid = " & iPaymentTypeId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentTypeName = oRs("paymenttypename")
	Else
		GetPaymentTypeName = ""
	End If

	oRs.Close 
	Set oRs = Nothing
			
End Function


'--------------------------------------------------------------------------------------------------
' string GetPaymentLocationName( iPaymentLocationId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentLocationName( ByVal iPaymentLocationId )
	Dim sSql, oRs

	sSql = "SELECT paymentlocationname FROM egov_paymentlocations WHERE paymentlocationid = " & iPaymentLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentLocationName = oRs("paymentlocationname")
	Else
		GetPaymentLocationName = ""
	End If

	oRs.Close
	Set oRs = Nothing
	
End Function


'--------------------------------------------------------------------------------------------------
' string FormatPhone( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhone( ByVal Number )

	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If

End Function



%>


