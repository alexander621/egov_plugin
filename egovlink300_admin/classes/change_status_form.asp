<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DROP_REGISTRANT_FORM.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/26/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/17/06	Steve Loar - Security, Header and nav changed
' 2.0	04/27/07   Steve Loar  -  Overhauled for Menlo Park Project
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sClassName, sResidentTypeDesc, sHeadName, sName, iMemberCount, iFamilyMemberId, iMembershipId
Dim iOldPaymentId, sResidentType, cTotalPrice

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

iclassid = request("classid")
itimeid = request("timeid")
iclasslistid = request("classlistid")
sClassName = GetClassName( iclassid, itimeid )
cTotalPrice = 0.00

%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="classes.css" />

	<script src="tablesort.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/layers.js"></script>
	<script src="../scripts/removespaces.js"></script>

	<script>
	<!--

		// create the egov NameSpace
		var eGovLink = eGovLink || {}; 

		// create the class sub-NameSpace with the methods inside
		eGovLink.Class = (function () {

			var global_valfield;	// retain valfield for timer thread
			// --------------------------------------------
			//                  setfocus
			// Delayed focus setting to get around IE bug
			// --------------------------------------------

			var setFocusDelayed = function()
			{
			  global_valfield.focus();
			}

			var setfocus = function(valfield)
			{
			  // save valfield in global variable so value retained when routine exits
			  global_valfield = valfield;
			  setTimeout( 'setFocusDelayed()', 100 );
			}

			var UpdatePrice = function( oPrice )
			{
				// this is for the pricing amounts
				var bValid = true;
				var total = 0.00;
				var balance = 0.00;

				//alert( oPrice.value );

				// Remove any extra spaces
				oPrice.value = removeSpaces(oPrice.value);

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

				// Calculate a new total price
				if (document.frmStatus.pricetypeid.length)   // If there is more than one price checkbox
				{
					var checklength = document.frmStatus.pricetypeid.length;
					var i = checklength - 1;

					for (l = 0; l <= i; l++)
					{
						if (document.frmStatus.pricetypeid[l].checked)
						{ 
							//total += Number(document.frmStatus.pricetypeid[l].value);
							total += Number(eval('document.frmStatus.price' + document.frmStatus.pricetypeid[l].value + '.value'));
						}
					}
				}
				else   // There is only one price checkbox
				{
					if (document.frmStatus.pricetypeid.checked)
					{
						total += Number(eval('document.frmStatus.price' + document.frmStatus.pricetypeid.value + '.value'));
					}
				}

				total = format_number(total * Number(document.frmStatus.quantity.value),2);
				balance = total - Number(document.frmStatus.paymenttotal.value);
				document.frmStatus.totalprice.value = total;
				document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
				document.getElementById("balancedue").innerHTML = format_number(balance,2);

				if ( bValid == false )
				{
					alert('Prices should be numbers in currency format.\nPlease correct this price.');
					setfocus(oPrice);
					return false;
				}
				return true; 
			}

			var UpdatePriceTotal = function( iPrice, bChecked )
			{
				// this is for the pricing checkboxes
				var total = 0.00;
				var balance = 0.00;

				//alert( iPrice );

				if (iPrice != "")
				{
					total = Number(document.frmStatus.totalprice.value);
					if (bChecked)
					{
						total += Number(iPrice) * Number(document.frmStatus.quantity.value);
					}
					else
					{
						total -= Number(iPrice) * Number(document.frmStatus.quantity.value);
					}
					total = format_number(total, 2);
					balance = total - Number(document.frmStatus.paymenttotal.value);
					document.frmStatus.totalprice.value = total;
					document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
					document.getElementById("balancedue").innerHTML = format_number(balance,2);
				}
			}

			var addTotal = function()
			{
				// this is for the payments entered
				var total = 0.00;
				var balance = 0.00;
				var bValid = true;
				var oProblem;

				//Charge
				document.frmStatus.amount1.value = removeSpaces(document.frmStatus.amount1.value);
				if (document.frmStatus.amount1.value != "")
				{
					var rege1 = /^\d*\.?\d{0,2}$/
					var Ok1 = rege1.exec(document.frmStatus.amount1.value);
					if ( Ok1 )
					{
						total += Number(document.frmStatus.amount1.value);
						document.frmStatus.amount1.value = format_number(Number(document.frmStatus.amount1.value),2);
					}
					else
					{
						document.frmStatus.amount1.value = '';
						bValid = false;
						oProblem = document.frmStatus.amount1;
					}
				}

				// Check
				document.frmStatus.amount2.value = removeSpaces(document.frmStatus.amount2.value);
				if (document.frmStatus.amount2.value != "")
				{
					var rege2 = /^\d*\.?\d{0,2}$/
					var Ok2 = rege2.exec(document.frmStatus.amount2.value);
					if ( Ok2 )
					{
						total += Number(document.frmStatus.amount2.value);
						document.frmStatus.amount2.value = format_number(Number(document.frmStatus.amount2.value),2);
					}
					else
					{
						document.frmStatus.amount2.value = '';
						bValid = false;
						oProblem = document.frmStatus.amount2;
					}
				}

				// Cash
				document.frmStatus.amount3.value = removeSpaces(document.frmStatus.amount3.value);
				if (document.frmStatus.amount3.value != "")
				{
					var rege3 = /^\d*\.?\d{0,2}$/
					var Ok3 = rege3.exec(document.frmStatus.amount3.value);
					if ( Ok3 )
					{
						total += Number(document.frmStatus.amount3.value);
						document.frmStatus.amount3.value = format_number(Number(document.frmStatus.amount3.value),2);
					}
					else
					{
						document.frmStatus.amount3.value = '';
						bValid = false;
						oProblem = document.frmStatus.amount3;
					}
				}

				// Account transfer if there
				var bexists = eval(document.frmStatus["amount4"]);
				if (bexists)
				{
					document.frmStatus.amount4.value = removeSpaces(document.frmStatus.amount4.value);
					if (document.frmStatus.amount4.value != "")
					{
						var rege4 = /^\d*\.?\d{0,2}$/
						var Ok4 = rege4.exec(document.frmStatus.amount4.value);
						if ( Ok4 )
						{
							total += Number(document.frmStatus.amount4.value);
							document.frmStatus.amount4.value = format_number(Number(document.frmStatus.amount4.value),2);
						}
						else
						{
							document.frmStatus.amount4.value = '';
							bValid = false;
							oProblem = document.frmStatus.amount4;
						}
					}
				}

				// Other Payment
				document.frmStatus.amount8.value = removeSpaces(document.frmStatus.amount8.value);
				if (document.frmStatus.amount8.value != "")
				{
					var rege8 = /^\d*\.?\d{0,2}$/
					var Ok8 = rege8.exec(document.frmStatus.amount8.value);
					if ( Ok8 )
					{
						total += Number(document.frmStatus.amount8.value);
						document.frmStatus.amount8.value = format_number(Number(document.frmStatus.amount8.value),2);
					}
					else
					{
						document.frmStatus.amount8.value = '';
						bValid = false;
						oProblem = document.frmStatus.amount8;
					}
				}

				balance = Number(document.frmStatus.totalprice.value) - total;
				document.getElementById("displaypaymenttotal").innerHTML = format_number(total,2);
				document.getElementById("balancedue").innerHTML = format_number(balance,2);
				document.frmStatus.paymenttotal.value = total;
				//alert(document.frmStatus.amount.value);

				if ( bValid == false )
				{
					alert('Payment amounts should be currency or blank.\nPlease correct this payment amount.');
					setfocus(oProblem);
					return false;
				}
				return true; 
			}

			var validate = function()
			{
				if (Number(document.frmStatus.paymenttotal.value) == Number(document.frmStatus.totalprice.value))
				{
					document.frmStatus.submit();
				}
				else
				{
					alert('Cannot complete this activation. \nThe payment total does not equal the total price.');
					return;
				}
			}

			// This makes the functions publically accessible
			return {
				UpdatePrice: UpdatePrice,
				UpdatePriceTotal: UpdatePriceTotal,
				addTotal: addTotal,
				validate: validate
				};
		}());

	//-->
	</script>

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Recreation: Activate Person on Wait List</strong></font><br /><br />
			<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>
		</p>
		<!--END: PAGE TITLE-->

		<!--BEGIN: DROP FORM-->
		<%

		' GET INFORMATION FOR THIS REGISTRANT
		sSql = "SELECT * FROM egov_class_roster WHERE classid = " & iclassid & " AND classtimeid = " & itimeid & " AND classlistid = " & iclasslistid &" ORDER BY status, userlname"

		Set oRegistrant = Server.CreateObject("ADODB.Recordset")
		oRegistrant.Open sSql, Application("DSN"), 3, 1

		If Not oRegistrant.EOF Then
			sHeadName = oRegistrant("userfname") & " " & oRegistrant("userlname")
			sName = oRegistrant("firstname") & " " & oRegistrant("lastname")
			iUserId = oRegistrant("userid")
			sResidentTypeDesc = oRegistrant("description")
			sResidentType = oRegistrant("residenttype")
			iFamilyMemberId = oRegistrant("familymemberid")
			If IsNull(iFamilyMemberId) Then
				iFamilyMemberId = 0
			End If 
			iMemberCount = GetMemberCount( iFamilyMemberId, iUserId, iMembershipId )
			iOldPaymentId = oRegistrant("paymentid")
			iQuantity = oRegistrant("quantity")
		Else
			' Something is wrong
			sHeadName = ""
			sName = ""
			iUserId = 0
			sResidentTypeDesc = ""
			iFamilyMemberId = 0
			iMemberCount = 0
			iOldPaymentId = 0
			sResidentType = "N"
			iQuantity = 1
		End If

		oRegistrant.Close 
		Set oRegistrant = Nothing
		%>

		<form name="frmStatus" action="change_status.asp" method="post">

			<p><strong>Class: </strong><%=sClassName%> </p>

			<p><strong>Name: </strong><%=sName%> ( <%=sResidentTypeDesc%> )</p>

			<p><strong>Head of Household: </strong><%=sHeadName%></p>

			<p>
				<strong>Quantity: </strong><%=iQuantity%>
				<input type="hidden" name="quantity" value="<%=iQuantity%>" />
			</p>

			<fieldset><legend><strong> Pricing </strong></legend><br />
				<%
				' Price options
				'response.write "<p><strong>Price:</strong><br />"
				cTotalPrice = ShowPriceOptions( iclassid, Session("OrgID"), sResidentType, iMemberCount, iMembershipId, 0, iUserId, iQuantity )
				'response.write vbcrlf & "</p>"
			%>
			</fieldset>

			<p>

				<input type="hidden" name="iclasslistid" value="<%=iclasslistid%>" />
				<input type="hidden" name="classid" value="<%=iclassid%>" />
				<input type="hidden" name="timeid" value="<%=itimeid%>" />
				<input type="hidden" name="oldpaymentid" value="<%=iOldPaymentId%>" />
				<input type="hidden" name="paymentamount" value="0.00" />
				<input type="hidden" name="iUserId" value="<%=iUserId%>" />

				<fieldset><legend><strong> Payment </strong></legend><br />
					<input type="hidden" value="0.00" name="paymenttotal" />
<%					ShowPaymentChoices iUserId, cTotalPrice  %>
					<br /><br />
					<input class="button" type="button" onClick="eGovLink.Class.validate();" name="complete" value="Complete Purchase" />
				</fieldset>
			</p>
		</form>
		<!--END: DROP FORM-->

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>



<%
'--------------------------------------------------------------------------------------------------
' Function CheckResTypeExists(iClassid, iorgid, sResidentType)
'--------------------------------------------------------------------------------------------------
Function CheckResTypeExists( ByVal iClassid, ByVal iorgid, ByVal sResidentType )
	Dim sSql, oRs

	CheckResTypeExists = False 
	sSql = "Select count(T.pricetype) as hits "
	sSql = sSql & " from egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " where T.pricetypeid = P.pricetypeid "
	sSql = sSql & " and orgid = " & iorgid & " and P.classid = " & iClassid & " and T.pricetype = '" & sResidentType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If clng(oRs("hits")) > 0 Then 
		CheckResTypeExists = True 
	End If 

	oRs.close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetClassName( iClassId, iTimeId )
'--------------------------------------------------------------------------------------------------
Function GetClassName( ByVal iClassId, ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT classname, activityno FROM egov_class C, egov_class_time T "
	sSql = sSql & "WHERE C.classid = T.classid AND C.classid = " & iClassId & " AND T.timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetClassName = oRs("classname") & " &nbsp; ( " & oRs("activityno") & " )"
	Else
		GetClassName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPaymentChoices iUserID, sBalanceDue
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentChoices( ByVal iUserID, ByVal sBalanceDue )
	Dim sSql, oRs

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, requirescheckno, requirescitizenaccount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE O.paymenttypeid = P.paymenttypeid AND isadminmethod = 1 AND O.orgid = " & Session("OrgID")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		'response.write "HasPayableAccounts = " &  HasPayableAccounts( iUserId ) 
		response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""50%"">"

		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"">Citizen Location:</td><td>" 
		ShowPaymentLocations
		response.write "</td></tr>"

		Do While Not oRs.EOF 
			'If (Not oRs("requirescitizenaccount")) Or (oRs("requirescitizenaccount") And HasPayableAccounts( iUserId )) Then 
				response.write vbcrlf & "<tr>"
				response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">"
				response.write oRs("paymenttypename") & ": "
				response.write "</td><td>"
				response.write "<input type=""text"" value="""" name=""amount" & oRs("paymenttypeid") & """ size=""20"" maxlength=""20"" onchange=""eGovLink.Class.addTotal()"" />"
				If oRs("requirescheckno") Then
					response.write " &nbsp; <strong>Check #:</strong> <input type=""text"" value="""" name=""checkno"" size=""8"" maxlength=""8"" />"
				End If 
				If oRs("requirescitizenaccount") Then
					response.write "&nbsp; <strong>From:</strong>" 
					ShowFamilyAccounts iUserID 
				End If 
				response.write "</td></tr>"
			'End If 
			oRs.MoveNext
		Loop
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"">Payment Total:</td><td><span id=""displaypaymenttotal"">0.00</span></td></tr>"
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"">Balance Due:</td><td><span id=""balancedue"">" & FormatNumber(sBalanceDue,2,,,0) & "</span></td></tr>"
		response.write vbcrlf & "<tr><td class=""label"" align=""right"" nowrap=""nowrap"">Notes:</td><td><textarea name=""notes"" class=""purchasenotes""></textarea></td></tr>"
		response.write vbcrlf & "</table>"
	End If 
	
	oRs.Close
	Set oRs = Nothing

End Sub 



'--------------------------------------------------------------------------------------------------
' void ShowPaymentLocations
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oRs

	sSql = "SELECT paymentlocationid, paymentlocationname FROM egov_paymentlocations "
	sSql = sSql & "WHERE isadminmethod = 1 ORDER BY paymentlocationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""PaymentLocationId"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("paymentlocationid") & """>" & oRs("paymentlocationname") & "</option>"
		oRs.MoveNext
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' double ShowPriceOptions( iClassid, iorgid, sResidentType, iMemberCount, iMembershipId, iPriceDiscountId, iUserId, iQuantity )
'--------------------------------------------------------------------------------------------------
Function ShowPriceOptions( ByVal iClassid, ByVal iorgid, ByVal sResidentType, ByVal iMemberCount, ByVal iMembershipId, ByVal iPriceDiscountId, ByVal iUserId, ByVal iQuantity )
	Dim sSql, oRs, iCount, sDiscount, bMemberTypematch, sMemberType, iMinPricetype, iMaxPriceType
	Dim iFamilyMemberId, sMemberCode, cTotalPrice

	iCount = 0
	cTotalPrice = CDbl(0.00)

	sDiscount = GetDiscountPhrase( iPriceDiscountId )

	bResTypeMatch = CheckResTypeExists(iClassid, iorgid, sResidentType)

	' IF at least one person in the family is a member, then set up for member pricing match
	If iMemberCount > 0 Then 
		sMemberType = "M"
	Else
		sMemberType = "O"
	End If 

	sSql = "SELECT P.pricetypeid, T.pricetypename, T.ismember, P.amount, T.pricetype, P.accountid, "
	sSql = sSql & " T.isfee, T.isbaseprice, T.checkmembership, P.membershipid, T.isdropin "
	sSql = sSql & " FROM egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " WHERE T.pricetypeid = P.pricetypeid AND T.isdropin = 0 "
	sSql = sSql & " AND orgid = " & iorgid & " AND P.classid = " & iClassid & " ORDER BY P.pricetypeid"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	'response.write "<!--egov_class_pricetype_price.pricetypeid -->"

	If Not oRs.EOF Then 
		iMinPricetype = clng(oRs("pricetypeid"))
		iMaxPriceType = clng(oRs("pricetypeid"))
		response.write vbcrlf & "<table id=""pricetable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Do While Not oRs.EOF
			If clng(oRs("pricetypeid")) < iMinPricetype Then
				iMinPricetype = clng(oRs("pricetypeid"))
			End If 
			If clng(oRs("pricetypeid")) > iMaxPriceType Then
				iMaxPriceType = clng(oRs("pricetypeid"))
			End If 
			' Display new time pick
			response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap"" valign=""top"">"
			response.write "<input type=""checkbox"" "
			If oRs("isfee") Then 
				' Always check a fee
				response.write " checked=""checked"" "
				cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
			Else 
				If oRs("isbaseprice") Then 
					' always check a base price
					response.write " checked=""checked"" "
					cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
				Else
					If oRs("pricetype") = sResidentType Then 
						' if the resident type requirement matches
						response.write " checked=""checked"" "
						cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
					Else
						If oRs("checkmembership") Then
							If oRs("pricetype") = sMemberType Then 
								response.write " checked=""checked"" "
								cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
							End If 
						End If 
					End If 
				End If 
			End If 
			response.write "id=""pricetypeid" & oRs("pricetypeid") & """ name=""pricetypeid"" value=""" & oRs("pricetypeid") & """ onClick=""eGovLink.Class.UpdatePriceTotal(document.frmStatus.price" & oRs("pricetypeid") & ".value, this.checked);"" /> " 
			response.write " &nbsp; " & oRs("pricetypename") & "</td>"
			'response.write "<td class=""priceentrytd"" valign=""top""><input type=""text"" id=""amount" & oRs("pricetypeid") & """ name=""amount" & oRs("pricetypeid") & """ value=""" & FormatNumber(CDbl(oRs("amount")),2) & """ size=""6"" maxlength=""6"" onchange=""if (UpdatePrice(this) == false){alert('Prices should be currency or blank.\nPlease correct this price.');this.focus();this.select();}"" /></td>" 
			response.write "<td class=""priceentrytd"" valign=""top""><input type=""text"" id=""price" & oRs("pricetypeid") & """ name=""price" & oRs("pricetypeid") & """ value=""" & FormatNumber(CDbl(oRs("amount")),2) & """ size=""6"" maxlength=""6"" onchange=""eGovLink.Class.UpdatePrice(this);"" /></td>" 
			response.write "<td class=""priceentrytd"">" & FormatCurrency(oRs("amount")) & "</td>"
			
			response.write "<td class=""pricemember"">"
			If oRs("ismember") Then
				' Show the membership for the one that requires membership
				'ShowMembership iMembershipId
				ShowMembership oRs("membershipid")
			Else
				response.write " &nbsp; "
			End If 
			If oRs("isdropin") Then
				' Input for drop in date
				response.write "Date: <input type=""text"" class=""datefield"" id=""dropindate" & oRs("pricetypeid") & """ name=""dropindate" & oRs("pricetypeid") & """ value=""" & FormatDateTime(date(),2) & """ />&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('dropindate" & oRs("pricetypeid") & "');"" /></span>"
			End If 
			response.write "</td>"
			response.write "<td class=""pricemember"">"
			If sDiscount <> "" Then 
				response.write " (" & sDiscount & ")"
			Else 
				response.write " &nbsp; "
			End If 
			response.write "</td></tr>"
			iCount = iCount + 1
			oRs.movenext 
		Loop 
		cTotalPrice = cTotalPrice * iQuantity
		response.write "<tr><td>Total Price</td><td><span id=""displaytotalprice"">" & FormatNumber(cTotalPrice,2) & "</span></td></tr>"
		response.write "</table>"
		response.write vbcrlf & "<input type=""hidden"" id=""totalprice"" name=""totalprice"" value=""" & cTotalPrice & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMinPricetype & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMaxPriceType & """ />"
	End If 

	oRs.Close
	Set oRs = Nothing

	ShowPriceOptions = cTotalPrice

End Function  


'--------------------------------------------------------------------------------------------------
' boolean CheckResTypeExists( iClassid, iorgid, sResidentType ) 
'--------------------------------------------------------------------------------------------------
Function CheckResTypeExists( ByVal iClassid, ByVal iorgid, ByVal sResidentType )
	Dim sSql, oRs

	CheckResTypeExists = False 
	sSql = "SELECT COUNT(T.pricetype) AS hits "
	sSql = sSql & " FROM egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " WHERE T.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND orgid = " & iorgid & " AND P.classid = " & iClassid & " AND T.pricetype = '" & sResidentType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If clng(oRs("hits")) > 0 Then 
		CheckResTypeExists = True 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetMemberCount( iFamilyMemberId, iUserid, iMembershipId )
'--------------------------------------------------------------------------------------------------
Function GetMemberCount( ByVal iFamilyMemberId, ByVal iUserid, ByRef iMembershipId )
	Dim sSql, oRs, sMembershipstatus

	sMembershipstatus = "O"
	GetMemberCount = 0
	iMembershipId = ""

	sSql = "SELECT poolpassid FROM egov_poolpassmembers WHERE familymemberid = " & iFamilyMemberId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		sMembershipstatus = DetermineMembership( iFamilyMemberId, iUserid, oRs("poolpassid") )
		If sMembershipstatus = "M" Then 
			GetMemberCount = 1
			iMembershipId = oRs("poolpassid")
			Exit Do 
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>

