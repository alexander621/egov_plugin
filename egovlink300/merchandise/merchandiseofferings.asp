<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="merchandisecommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiseofferings.asp
' AUTHOR: Steve Loar
' CREATED: 05/14/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Page that allows signup of Regatta Teams
'
' MODIFICATION HISTORY
' 1.0   05/14/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCartId, sSaveButton, iItemTypeId, sShipToName, sShipToAddress, sShipToCity, sShipToState
Dim sShipToZip, iUserid, iMerchandiseCount

Session("RedirectPage") = "merchandise/merchandiseofferings.asp" & iClassId
Session("RedirectLang") = "Return to Merchandise Offerings"
session("ManageURL") = ""
iMerchandiseCount = CLng(0)

'If they do not have a userid set, take them to the login page automatically
If request.cookies("userid") = "" Or request.cookies("userid") = "-1" Then 
	session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along to select your merchandise."
	response.redirect "../user_login.asp"
Else
	iUserid = request.cookies("userid")
End If 

If request("cartid") <> "" Then 
	iCartId = CLng(request("cartid"))
	sSaveButton = "Save Changes"
	sShipToName = GetCartValue( iCartId, "shiptoname" )
	sShipToAddress = GetCartValue( iCartId, "shiptoaddress" )
	sShipToCity = GetCartValue( iCartId, "shiptocity" )
	sShipToState = GetCartValue( iCartId, "shiptostate" )
	sShipToZip = GetCartValue( iCartId, "shiptozip" )
Else
	iCartId = CLng(0)
	sSaveButton = "Add To Cart"
	GetUserDefaultShippingValues iUserId, sShipToName, sShipToAddress, sShipToCity, sShipToState, sShipToZip
End If 

iItemTypeId = GetItemTypeId( "merchandise" )

%>

<html>
<head>

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If

	%>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="merchandise.css" />

	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function upperCase(x)
		{
			var y = document.getElementById(x).value;
			document.getElementById(x).value = y.toUpperCase();
		}

		function ValidateQty( oQty, iRow )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oQty.value = removeSpaces(oQty.value);
			//Remove commas that would cause problems in validation
			oQty.value = removeCommas(oQty.value);

			// Validate the format of the price
			if (oQty.value != "")
			{
				var rege = /^\d*$/
				var Ok = rege.exec(oQty.value);
				if ( Ok )
				{
					oQty.value = Number(oQty.value);
				}
				else 
				{
					oQty.value = 0;
					bValid = false;
				}
			}

			CalculateTotals( iRow );
			
			if ( bValid == false ) 
			{
				$("orderok").value = 'false';
				document.getElementById(oQty.id).focus();
				inlineMsg(oQty.id,'<strong>Invalid Value: </strong>Quantities should positive whole numbers.',10,oQty.id);
				return false;
			}
			return true;
		}

		function CalculateTotals( iRow )
		{
			// Calculate the new sub total
			$("subtotal"+iRow).value = format_number(Number($("quantity"+iRow).value) * Number($("price"+iRow).value), 2);

			//Calculate the total
			var iTotal = 0.00
			var iStop = Number($("maxmerchandiseitems").value) + 1;
			for (x = 1; x < iStop; x++ )
			{
				if ($("subtotal"+x))
				{
					iTotal += Number($("subtotal"+x).value);
				}
			}
			$("total").value = format_number(iTotal,2);
		}

		function validatePurchase()
		{
			if ($("orderok").value == 'true')
			{
				// Check that the total for the merchandise is not 0.00
				if (Number($F("total")) == Number('0.00'))
				{
					$("total").focus();
					inlineMsg($("total").id,'<strong>No Merchandise Total: </strong>Please select some items for purchase.',5,$("total").id);
					return;
				}

				// Make sure all fields are entered for the shipping info
				if ($F("shiptoname") == '')
				{
					$("shiptoname").focus();
					inlineMsg($("shiptoname").id,'<strong>Missing Shiping Information: </strong>Please enter a name.',5,$("shiptoname").id);
					return;
				}
				if ($F("shiptoaddress") == '')
				{
					$("shiptoaddress").focus();
					inlineMsg($("shiptoaddress").id,'<strong>Missing Shiping Information: </strong>Please enter an address.',5,$("shiptoaddress").id);
					return;
				}
				if ($F("shiptocity") == '')
				{
					$("shiptocity").focus();
					inlineMsg($("shiptocity").id,'<strong>Missing Shiping Information: </strong>Please enter a city.',5,$("shiptocity").id);
					return;
				}
				if ($F("shiptostate") == '')
				{
					$("shiptostate").focus();
					inlineMsg($("shiptostate").id,'<strong>Missing Shiping Information: </strong>Please enter a state.',5,$("shiptostate").id);
					return;
				}
				if ($F("shiptozip") == '')
				{
					$("shiptozip").focus();
					inlineMsg($("shiptozip").id,'<strong>Missing Shiping Information: </strong>Please enter a zip code.',5,$("shiptozip").id);
					return;
				}

				document.frmMerchandise.submit();

			}
			else
			{
				$("orderok").value = 'true';
			}
		}


	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<form name="frmMerchandise"" method="post" action="merchandisetocart.asp">
	<input type="hidden" name="cartid" value="<%=iCartId%>" />
	<input type="hidden" name="itemtypeid" value="<%=iItemTypeId%>" />
	<input type="hidden" name="egovuserid" value="<%=iUserid%>" />

	<p>
		Enter the quantities of the items you wish to purchase from the list below.<br />
<%		iMerchandiseCount = ShowMerchandise( iCartId )		%>						
	</p>

<%
	' Do not want any further displays if there is no merchandise to purchase
	If CLng(iMerchandiseCount) > CLng(0) Then 
%>
		<p>
			Enter the <strong>Shipping Information</strong> for this order. All fields are required.<br /><br />
				<table id="shippinginfoentry" cellpadding="0" cellspacing="0" border="0">
				<tr><td align="right">Name:&nbsp;</td><td><input type="text" id="shiptoname" name="shiptoname" value="<%=sShipToName%>" size="50" maxlength="50" /></td></tr>
				<tr><td align="right">Address:&nbsp;</td><td><input type="text" id="shiptoaddress" name="shiptoaddress" value="<%=sShipToAddress%>" size="50" maxlength="50" /></td></tr>
				<tr><td align="right">City:&nbsp;</td><td><input type="text" id="shiptocity" name="shiptocity" value="<%=sShipToCity%>" size="50" maxlength="50" /></td></tr>
				<tr><td align="right">State:&nbsp;</td><td><input type="text" id="shiptostate" name="shiptostate" value="<%=sShipToState%>" size="2" maxlength="2" onchange="upperCase(this.id);" /></td></tr>
				<tr><td align="right">Zip:&nbsp;</td><td><input type="text" id="shiptozip" name="shiptozip" value="<%=sShipToZip%>" size="10" maxlength="10" /></td></tr>
			</table>
		</p>

<%
		response.write vbcrlf & "<p>"
		response.write vbcrlf & "<input type=""button"" class=""button"" value=""" & sSaveButton & """ onclick=""validatePurchase()"" /> &nbsp;"
		response.write vbcrlf & "</p>"
	End If 
%>

</form>

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetUserDefaultShippingValues( iUserId, sShipToName, sShipToAddress, sShipToCity, sShipToState, sShipToZip )
'--------------------------------------------------------------------------------------------------
Sub GetUserDefaultShippingValues( ByVal iUserId, ByRef sShipToName, ByRef sShipToAddress, ByRef sShipToCity, ByRef sShipToState, ByRef sShipToZip )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(useraddress,'') AS useraddress, ISNULL(usercity,'') AS usercity, "
	sSql = sSql & " ISNULL(userstate,'') AS userstate, ISNULL(userzip,'') AS userzip "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sShipToName = Trim(oRs("userfname") & " " & oRs("userlname"))
		sShipToAddress = oRs("useraddress")
		sShipToCity = oRs("usercity")
		sShipToState = UCase(oRs("userstate"))
		sShipToZip = oRs("userzip")
	Else
		sShipToName = ""
		sShipToAddress = ""
		sShipToCity = ""
		sShipToState = ""
		sShipToZip = ""
	End If

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowMerchandise( iCartId )
'--------------------------------------------------------------------------------------------------
Function ShowMerchandise( iCartId )
	Dim sSql, oRs, sOldMerchandise, dSubTotal, dTotal, dPrice, iQty, iRowCount

	sOldMerchandise = "^^@^^"
	dTotal = CDbl(0.00)
	iRowCount = CLng(0)

	sSql = "SELECT C.merchandisecatalogid, M.merchandise, MC.merchandisecolor, MC.isnocolor, MS.merchandisesize, MS.isnosize, M.price, C.instock "
	sSql = sSql & " FROM egov_merchandisecatalog C, egov_merchandise M, egov_merchandisecolors MC, egov_merchandisesizes MS "
	sSql = sSql & " WHERE C.merchandiseid = M.merchandiseid "
	sSql = sSql & " AND C.merchandisecolorid = MC.merchandisecolorid "
	sSql = sSql & " AND C.merchandisesizeid = MS.merchandisesizeid "
	sSql = sSql & " AND C.orgid = " & iorgid & " AND C.showpublic = 1 AND M.showpublic = 1 "
	sSql = sSql & " ORDER BY M.merchandise, MC.merchandisecolor, MS.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		'DRAW TABLE WITH MERCHANDISE LISTED
		response.write vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""merchandiselist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		
		'HEADER ROW
		response.write vbcrlf & "<tr><th>Merchandise</th><th>Color</th><th>Size</th><th>Price<br />Each</th><th>Qty</th><th>Item<br />Total</th></tr>"

		' LOOP THRU AND DISPLAY The EVENTS
		Do While Not oRs.EOF
  			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
   			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
   			End If 
			response.write ">"

			response.write "<td"
			If sOldMerchandise <> oRs("merchandise") Then 
				sTdClass = "class=""newmerchandise"" "
				response.write " " & sTdClass & ">"
				response.write "<strong>" & oRs("merchandise") & "</strong>"
				sOldMerchandise = oRs("merchandise")
			Else
				sTdClass = ""
				response.write ">"
				response.write "&nbsp;"
			End If 
			response.write "<input type=""hidden"" name=""merchandisecatalogid" & iRowCount & """ value=""" & oRs("merchandisecatalogid") & """ /></td>"
			
			response.write "<td " & sTdClass & "align=""center"">"
			If oRs("isnocolor") Then
				response.write "&nbsp;"
			Else
				response.write oRs("merchandisecolor")
			End If 
			response.write "</td>"

			response.write "<td " & sTdClass & "align=""center"">"
			If oRs("isnosize") Then
				response.write "&nbsp;"
			Else
				response.write oRs("merchandisesize")
			End If 
			response.write "</td>"
			
			If oRs("instock") Then 
				response.write "<td " & sTdClass & "align=""center"">"
				dPrice = CDbl(oRs("price"))
				response.write FormatNumber(dPrice,2,,,0)
				response.write "<input type=""hidden"" id=""price" & iRowCount & """ name=""price" & iRowCount & """ value=""" & dPrice & """ />"
				response.write "</td>"
				
				response.write "<td " & sTdClass & "align=""center""><input type=""text"" size=""6"" maxlength=""6"" id=""quantity" & iRowCount & """ name=""quantity" & iRowCount & """ value="""
				If iCartId > CLng(0) Then
					' Get the quantity from the cart
					iQty = GetCartQuantity( iCartId, oRs("merchandisecatalogid") )
				Else
					iQty = 0
				End If 
				response.write iQty
				response.write """ onchange=""clearMsg('quantity" & iRowCount & "');return ValidateQty(this, " & iRowCount & ");"" /></td>"

				dSubTotal = FormatNumber(CDbl(dPrice * iQty),2,,,0)
				dTotal = CDbl(dTotal) + CDbl(dSubTotal)

				response.write "<td " & sTdClass & "align=""center""><input type=""text"" size=""10"" maxlength=""10"" readonly=""readonly"" id=""subtotal" & iRowCount & """ name=""subtotal" & iRowCount & """ value=""" & dSubTotal & """ tabindex=""-1"" /></td>"
			Else
				response.write "<td " & sTdClass & "align=""center"" colspan=""3"">Out of Stock</td>"
			End If 

			response.write "</tr>"
  			oRs.MoveNext
		Loop 
		' Total Row
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">Total</td><td align=""center""><input type=""text"" size=""10"" maxlength=""10"" readonly=""readonly"" id=""total"" name=""total"" value=""" & FormatNumber(dTotal,2,,,0) & """ tabindex=""-1"" /></td></tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 
		response.write vbcrlf & "<input type=""hidden"" id=""maxmerchandiseitems"" name=""maxmerchandiseitems"" value=""" & iRowCount & """ />"
		response.write vbcrlf & "<input type=""hidden"" id=""orderok"" name=""orderok"" value=""true"" />"

		

	Else
		response.write "<p><font color=""red""><b>We're sorry. No merchandise is available for purchase at this time.</b></font></p>"
	End If

	oRs.Close
	Set oRs = Nothing 

	ShowMerchandise = iRowCount

End Function 


'------------------------------------------------------------------------------
' Function GetCartQuantity( iCartId, iMerchandiseCatalogId )
'------------------------------------------------------------------------------
Function GetCartQuantity( iCartId, iMerchandiseCatalogId )
	Dim sSql, oRs

	sSql = "SELECT quantity FROM egov_class_cart_merchandiseitems WHERE cartid = " & iCartId
	sSql = sSql & " AND merchandisecatalogid = " & iMerchandiseCatalogId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCartQuantity = CLng(oRs("quantity"))
	Else
		GetCartQuantity = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 




%>