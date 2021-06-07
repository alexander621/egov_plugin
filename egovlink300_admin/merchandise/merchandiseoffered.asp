<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="merchandisecommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiseoffered.asp
' AUTHOR: Steve Loar
' CREATED: 04/29/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the merchandise that is offered for purchase
'
' MODIFICATION HISTORY
' 1.0	4/29/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCartId, sSaveButton, iUserId, iItemTypeId, sShipToName, sShipToAddress, sShipToCity, sShipToState
Dim sShipToZip

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "purchase merchandise", sLevel	' In common.asp

If request("cartid") <> "" Then
	iCartId = CLng(request("cartid"))
	iUserId = GetCartValue( iCartId, "userid" )
	sShipToName = GetCartValue( iCartId, "shiptoname" )
	sShipToAddress = GetCartValue( iCartId, "shiptoaddress" )
	sShipToCity = GetCartValue( iCartId, "shiptocity" )
	sShipToState = GetCartValue( iCartId, "shiptostate" )
	sShipToZip = GetCartValue( iCartId, "shiptozip" )
Else
	iCartId = 0
	If CLng(session("eGovUserId")) <> CLng(0) Then 
		iUserid = Session("eGovUserId")
	Else 
		iUserId = 0
	End If 
	sShipToName = ""
	sShipToAddress = ""
	sShipToCity = ""
	sShipToState = ""
	sShipToZip = ""
End If 

If CLng(iCartId) = CLng(0) Then
	sSaveButton = "Add To Cart"
Else
	sSaveButton = "Save Changes"
End If 

iItemTypeId = GetItemTypeId( "merchandise" )

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="merchandise.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<!--
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>


	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--

		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		function ViewCart()
		{
			location.href='../classes/class_cart.asp';
		}

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.frmMerchandise.searchname.value);
			var searchtext = document.frmMerchandise.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.frmMerchandise.egovuserid.length ; x++)
			{
				optiontext = document.frmMerchandise.egovuserid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.frmMerchandise.egovuserid.selectedIndex = x;
					document.frmMerchandise.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.frmMerchandise.searchstart.value = x;
					//document.frmMerchandise.submit();
					doAjax('getshiptoinfo.asp', 'userid=' + document.frmMerchandise.egovuserid.options[document.frmMerchandise.egovuserid.selectedIndex].value, 'UpdateShippingInfo', 'get', '0');

					return;
				}
			}
			document.frmMerchandise.results.value = 'End of List - No Match Found.';
			document.getElementById('searchresults').innerHTML = 'End of List - No Match Found.';
			document.frmMerchandise.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.frmMerchandise.searchstart.value = -1;
		}

		function UserPick()
		{
			document.frmMerchandise.searchname.value = '';
			document.frmMerchandise.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.frmMerchandise.searchstart.value = -1;
			//document.frmMerchandise.submit();

			doAjax('getshiptoinfo.asp', 'userid=' + document.frmMerchandise.egovuserid.options[document.frmMerchandise.egovuserid.selectedIndex].value, 'UpdateShippingInfo', 'get', '0');
		}

		function UpdateShippingInfo( sReturnJSON )
		{
			var json = sReturnJSON.evalJSON(true); 
			if (document.frmMerchandise.egovuserid.options[document.frmMerchandise.egovuserid.selectedIndex].value > 0)
			{
				if (json.flag == 'success')
				{
					$("shiptoname").value = json.shiptoname;
					$("shiptoaddress").value = json.shiptoaddress;
					$("shiptocity").value = json.shiptocity;
					$("shiptostate").value = json.shiptostate;
					$("shiptozip").value = json.shiptozip;
				}
			}
		}

		function EditUser()
		{
			var iUserId = document.frmMerchandise.egovuserid.options[document.frmMerchandise.egovuserid.selectedIndex].value;
			location.href='../dirs/update_citizen.asp?userid=' + iUserId;
		}

		function NewUser()
		{
			location.href='../dirs/register_citizen.asp';
		}

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

		function ValidatePrice( oPrice, iRow )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			// Validate the format of the price
			if (oPrice.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPrice.value);
				if ( Ok )
				{
					oPrice.value = format_number(Number(oPrice.value),2);
				}
				else 
				{
					oPrice.value = format_number(0,2);
					bValid = false;
				}
			}

			CalculateTotals( iRow );
			
			if ( bValid == false ) 
			{
				$("orderok").value = 'false';
				document.getElementById(oPrice.id).focus();
				inlineMsg(oPrice.id,'<strong>Invalid Value: </strong>Prices should numbers in currency format.',10,oPrice.id);
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
				// Make sure a registered user is picked
				if (document.frmMerchandise.egovuserid.options[document.frmMerchandise.egovuserid.selectedIndex].value == 0 )
				{
					tabView.set('activeIndex',0);
					$("egovuserid").focus();
					inlineMsg($("egovuserid").id,'<strong>Invalid Purchaser: </strong>Please select a registered user.',5,$("egovuserid").id);
					return;
				}

				// Check that the total for the merchandise is not 0.00
				if (Number($F("total")) == Number('0.00'))
				{
					tabView.set('activeIndex',1);
					$("total").focus();
					inlineMsg($("total").id,'<strong>No Merchandise Total: </strong>Please select some items for purchase.',5,$("total").id);
					return;
				}

				// Make sure all fields are entered for the shipping info
				if ($F("shiptoname") == '')
				{
					tabView.set('activeIndex',2);
					$("shiptoname").focus();
					inlineMsg($("shiptoname").id,'<strong>Missing Shiping Information: </strong>Please enter a name.',5,$("shiptoname").id);
					return;
				}
				if ($F("shiptoaddress") == '')
				{
					tabView.set('activeIndex',2);
					$("shiptoaddress").focus();
					inlineMsg($("shiptoaddress").id,'<strong>Missing Shiping Information: </strong>Please enter a address.',5,$("shiptoaddress").id);
					return;
				}
				if ($F("shiptocity") == '')
				{
					tabView.set('activeIndex',2);
					$("shiptocity").focus();
					inlineMsg($("shiptocity").id,'<strong>Missing Shiping Information: </strong>Please enter a city.',5,$("shiptocity").id);
					return;
				}
				if ($F("shiptostate") == '')
				{
					tabView.set('activeIndex',2);
					$("shiptostate").focus();
					inlineMsg($("shiptostate").id,'<strong>Missing Shiping Information: </strong>Please enter a state.',5,$("shiptostate").id);
					return;
				}
				if ($F("shiptozip") == '')
				{
					tabView.set('activeIndex',2);
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
<body class="yui-skin-sam">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<% If CartHasItems() Then %>
	<div id="topbuttons">
		<input type="button" name="viewcart" class="button" value="View Cart" onclick="ViewCart();" />
	</div>
<%	End If %>


	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Purchase Merchandise</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

		<form name="frmMerchandise" action="merchandisetocart.asp" method="post">
			<p>
				<input type="button" class="button" value="<%=sSaveButton%>" onclick="validatePurchase()" /> &nbsp;
			</p>
			<input type="hidden" name="cartid" value="<%=iCartId%>" />
			<input type="hidden" name="itemtypeid" value="<%=iItemTypeId%>" />

			<div id="demo" class="yui-navset">
				<ul class="yui-nav">
					<li><a href="#tab1"><em>Purchaser Information</em></a></li>
					<li><a href="#tab2"><em>Merchandise</em></a></li>
					<li><a href="#tab3"><em>Shipping Information</em></a></li>
				</ul>            
				<div class="yui-content">

					<div id="tab1"> <!-- Purchaser Information -->
						<p><br />
							Select the registered user who is making the purchase.<br /><br />
		<%					' Show pick of registered users and their detail info.
							ShowRegisteredUsers iUserId
		%>
						</p>
					</div>
					<div id="tab2"> <!-- Merchandise -->
						<p><br />
<%							ShowMerchandise iCartId		%>						
						</p>
					</div>
					<div id="tab3"> <!-- Shipping -->
						<p><br />
							Enter the Shipping Information for this order. All fields are required.<br /><br />
						</p>
						<table id="shippinginfoentry" cellpadding="0" cellspacing="0" border="0">
							<tr><td align="right">Name:&nbsp;</td><td><input type="text" id="shiptoname" name="shiptoname" value="<%=sShipToName%>" size="50" maxlength="50" /></td></tr>
							<tr><td align="right">Address:&nbsp;</td><td><input type="text" id="shiptoaddress" name="shiptoaddress" value="<%=sShipToAddress%>" size="50" maxlength="50" /></td></tr>
							<tr><td align="right">City:&nbsp;</td><td><input type="text" id="shiptocity" name="shiptocity" value="<%=sShipToCity%>" size="50" maxlength="50" /></td></tr>
							<tr><td align="right">State:&nbsp;</td><td><input type="text" id="shiptostate" name="shiptostate" value="<%=sShipToState%>" size="2" maxlength="2" onchange="upperCase(this.id);" /></td></tr>
							<tr><td align="right">Zip:&nbsp;</td><td><input type="text" id="shiptozip" name="shiptozip" value="<%=sShipToZip%>" size="10" maxlength="10" /></td></tr>
						</table>

					</div>
				</div>
			</div>

		</form>
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
' Sub ShowMerchandise( iCartId )
'--------------------------------------------------------------------------------------------------
Sub ShowMerchandise( iCartId )
	Dim sSql, oRs, sOldMerchandise, dSubTotal, dTotal, dPrice, iQty

	sOldMerchandise = "^^@^^"
	dTotal = CDbl(0.00)

	sSql = "SELECT C.merchandisecatalogid, M.merchandise, MC.merchandisecolor, MC.isnocolor, MS.merchandisesize, MS.isnosize, M.price, C.instock "
	sSql = sSql & " FROM egov_merchandisecatalog C, egov_merchandise M, egov_merchandisecolors MC, egov_merchandisesizes MS "
	sSql = sSql & " WHERE C.merchandiseid = M.merchandiseid "
	sSql = sSql & " AND C.merchandisecolorid = MC.merchandisecolorid "
	sSql = sSql & " AND C.merchandisesizeid = MS.merchandisesizeid "
	sSql = sSql & " AND C.orgid = " & session("orgid") & " AND C.showpublic = 1 AND M.showpublic = 1 "
	sSql = sSql & " ORDER BY M.merchandise, MC.merchandisecolor, MS.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		'DRAW TABLE WITH MERCHANDISE LISTED
		response.write vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""merchandiselist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		
		'HEADER ROW
		response.write vbcrlf & "<tr><th>Merchandise</th><th>Color</th><th>Size</th><th>Price</th><th>Qty</th><th>Item<br />Total</th></tr>"

		iRowCount = 0
		
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
				response.write "<td " & sTdClass & "align=""center""><input type=""text"" size=""6"" maxlength=""6"" id=""price" & iRowCount & """ name=""price" & iRowCount & """ value="""
				If iCartId > CLng(0) Then
					' Get the price from the cart
					dPrice = GetCartPrice( iCartId, oRs("merchandisecatalogid"), CDbl(oRs("price")) )
				Else
					dPrice = CDbl(oRs("price"))
				End If 
				response.write FormatNumber(dPrice,2,,,0)
				response.write """ onchange=""clearMsg('price" & iRowCount & "');return ValidatePrice(this, " & iRowCount & ");"" /></td>"
				
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
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">Total</td><td align=""center""><input type=""text"" size=""10"" maxlength=""10"" readonly=""readonly"" tabindex=""-1"" id=""total"" name=""total"" value=""" & FormatNumber(dTotal,2,,,0) & """ /></td></tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 
		response.write vbcrlf & "<input type=""hidden"" id=""maxmerchandiseitems"" name=""maxmerchandiseitems"" value=""" & iRowCount & """ />"
		response.write vbcrlf & "<input type=""hidden"" id=""orderok"" name=""orderok"" value=""true"" />"

	Else
		response.write "<p><font color=""red""><b>No merchandise is available for purchase.</b></font></p>"
	End If

	oRs.Close
	Set oRs = Nothing 


End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowRegisteredUsers( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegisteredUsers( iUserId )

	response.write vbcrlf & "<p>Name Search: <input type=""text"" name=""searchname"" value="""" size=""25"" maxlength=""50"" onchange=""javascript:ClearSearch();"" />"
	response.write vbcrlf & "<input type=""button"" class=""button"" value=""Search"" onclick=""javascript:SearchCitizens(document.frmMerchandise.searchstart.value);"" /> &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:NewUser();"" value=""New User"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""results"" value="""" />"
	response.write vbcrlf & "<input type=""hidden"" name=""searchstart"" value="""" />"
	response.write vbcrlf & "<span id=""searchresults""> </span>"
	response.write vbcrlf & "<br /><div id=""searchtip"">(last name, first name)</div>"
	response.write vbcrlf & "</p>"
	response.write vbcrlf & "<p>"
	response.write vbcrlf & "Select Name: <select id=""egovuserid"" name=""egovuserid"" onchange=""javascript:UserPick();"">"
	
	' Create the user pick dropdown
	ShowUserDropDown iUserId 
	
	response.write vbcrlf & "</select>"
	response.write vbcrlf & " &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:EditUser();"" value=""Edit User Profile"" />"
	response.write vbcrlf & "</p>" 
	response.write vbcrlf & "<div id=""userinfo""> </div>"
End Sub 


'------------------------------------------------------------------------------
' Sub ShowUserDropDown(iUserId)
'------------------------------------------------------------------------------
Sub ShowUserDropDown( iUserId )
	Dim oCmd, oRs

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oRs = .Execute
	End With

	response.write vbcrlf & "<option value=""0"">Select a Registered User...</option>"
	Do While Not oRs.EOF 
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"
		If CLng(iUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("userlname") & ", " & oRs("userfname") & " &ndash; " & oRs("useraddress") & "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing
End Sub 


'------------------------------------------------------------------------------
' Function GetCartPrice( iCartId, iMerchandiseCatalogId, dDefaultPrice )
'------------------------------------------------------------------------------
Function GetCartPrice( iCartId, iMerchandiseCatalogId, dDefaultPrice )
	Dim sSql, oRs

	sSql = "SELECT price FROM egov_class_cart_merchandiseitems WHERE cartid = " & iCartId
	sSql = sSql & " AND merchandisecatalogid = " & iMerchandiseCatalogId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCartPrice = CDbl(oRs("price"))
	Else
		GetCartPrice = dDefaultPrice
	End If 

	oRs.Close
	Set oRs = Nothing 
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
