<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_cart.asp
' AUTHOR: Steve Loar
' CREATED: 03/24/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the shopping cart contents for a given user.
'
' MODIFICATION HISTORY
' 1.0 03/24/06  Steve Loar - Initial Version
' 1.1	10/17/06	 Steve Loar - Security, Header and nav changed
' 1.2 01/07/09  David Boyer - Added "DisplayRosterPublic" fields for Craig,CO custom team registration
' 1.3 11/19/09 David Boyer - Added "pants size" to team registration section
' 1.4 11/19/09 David Boyer - Now pull team/pants sizes from database
' 1.5	04/07/2010	Steve Loar - No more regatta team members, added team group size
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim iUserId, sUserName, sClassName, iClassId, iTimeId, sItemType, bIsRegattaTeam
 Dim sTeamName

 sLevel = "../" ' Override of value from common.asp

 If Not UserHasPermission( session("userid"), "registration" ) Then 
   	response.redirect sLevel & "permissiondenied.asp"
 End If 

 sClassName = ""
 iUserId    = 0
 sUserName  = ""

 If request("iuserid") = "" Then 
	'The id is not passed, so get the user in the cart
	iUserId = getCartuserid()
 Else 
	iUserId = request("iuserid")	' The person they are adding classes for
 End If 

 If request("iClassId") <> "" Then 
	iClassId   = request("iClassId")	' The last class added, so we can take them back there
	sClassName = getClassName( iClassId )
	If request("iTimeId") <> "" Then 
		iTimeId = request("iTimeId")
	Else
		iTimeId = 0
	End If 
 End If 

 If request("isregattateam") <> "" Then
	bIsRegattaTeam = True
Else
	bIsRegattaTeam = False 
 End If 

 response.Expires = 60
 response.Expiresabsolute = Now() - 1
 response.AddHeader "pragma","no-store"
 response.AddHeader "cache-control","private"
 response.CacheControl = "no-store" 'HTTP prevent back button

'Check for org features
 lcl_orghasfeature_residency_verification = orghasfeature("residency verification")

%>

<html>
<head>
	<meta charset="UTF-8">

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="javascript">
	<!--

		window.onbeforeunload = function () {   
			// stuff do do before the window is unloaded here.
		}
		truncateDecimals = function (number, digits) {
    			var multiplier = Math.pow(10, digits),
        			adjustedNum = number * multiplier,
        			truncatedNum = Math[adjustedNum < 0 ? 'ceil' : 'floor'](adjustedNum);
			
    			return truncatedNum / multiplier;
			};

		function RemoveItem( sClassName, iCartId, iTimeId, sBuyOrWait, iIsDropIn )
		{
			if (confirm('Remove ' + sClassName + ' from cart?'))
			{
				//alert(iIsDropIn);
				//alert('class_remove.asp?iCartId=' + iCartId + '&iTimeId=' + iTimeId + '&sBuyOrWait=' + sBuyOrWait + '&iIsDropIn=' + iIsDropIn);
				location.href='class_remove.asp?iCartId=' + iCartId + '&iTimeId=' + iTimeId + '&sBuyOrWait=' + sBuyOrWait + '&iIsDropIn=' + iIsDropIn;
			}
		}

		function RemoveRegattaTeamItem( sClassName, iCartId )
		{
			if (confirm('Remove ' + sClassName + ' from cart?'))
			{
				location.href='regattateamremove.asp?cartid=' + iCartId;
			}
		}

		function RemoveMerchandiseItem( iCartId )
		{
			if (confirm('Remove this merchandise purchase from the cart?'))
			{
				location.href='../merchandise/merchandiseremove.asp?cartid=' + iCartId;
			}
		}

		function UpdateCart()
		{
			// Change the action to the update page, then submit it
			document.cartForm.action = 'class_updatecart.asp';
			document.cartForm.submit();
		}

		function checkAmount()
		{
			//Credit Amounts
			if (document.cartForm.amount.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(document.cartForm.amount.value);
				if ( Ok )
				{
					document.cartForm.amount.value = Number(document.cartForm.amount.value);
					document.cartForm.amount.value = format_number(document.cartForm.amount.value,2);
				}
				else
				{
					alert('Values should be currency or blank.\nPlease correct this amount.');
					document.cartForm.amount.focus();
				}
			}
		}

		function validate()
		{
			if ($("#amount2").val() != "" && $("#checkno").val() == "")
			{
				if ( ! confirm('You have a check payment without a check number. \nDo you wish to continue?'))
				{
					$("#checkno").focus();
					return;
				}
			}
			if (convertToCents($("#amount").val()) == convertToCents($("#purchasetotal").val()))
			{
				//alert('OK');
				document.cartForm.submit();
			}
			else
			{
				alert('Cannot complete this purchase. \nThe payment total does not equal the cart total.');
				return;
			}
		}

///////////////////////////////////////////////////////////////////////////////////////////
		function addTotal()
		{
			var total = 0;
			var balance = 0;
			var rege;
			var Ok;
			var amount = "";

			rege = /^\d*\.?\d{0,2}$/

			//Charge
			if ($("#amount1").val().length > 0)
			{
				//alert(document.cartForm.amount1.value);
				amount = $("#amount1").val();
				// Remove any extra spaces
				amount = removeSpaces(amount);
				//Remove commas that would cause problems in validation
				amount = removeCommas(amount);

				if ( amount != "" )
				{
					//rege = /^\d*\.?\d{0,2}$/
					Ok = rege.exec(amount);
					if ( Ok )
					{
						total += convertToCents(amount);
						amount = Number(amount);
						$("#amount1").val( format_number(amount,2) );
					}
					else
					{
						//alert('Values should be currency or blank.\nPlease correct this amount.');
						$("#amount1").focus();
						inlineMsg('amount1','<strong>Invalid Value: </strong>Values should be currency or blank.',3,'amount1');
					}
				}
				else
					$("#amount1").val( "" );
			}

			// Check
			if ($("#amount2").val().length > 0)
			{
				amount = $("#amount2").val();
				// Remove any extra spaces
				amount = removeSpaces(amount);
				//Remove commas that would cause problems in validation
				amount = removeCommas(amount);

				if ( amount != "" )
				{
					//rege = /^\d*\.?\d{0,2}$/
					Ok = rege.exec(amount);
					if ( Ok )
					{
						total += convertToCents(amount);
						amount = Number(amount);
						$("#amount2").val( format_number(amount,2) );
					}
					else
					{
						//alert('Values should be currency or blank.\nPlease correct this amount.');
						$("#amount2").focus();
						inlineMsg('amount2','<strong>Invalid Value: </strong>Values should be currency or blank.',3,'amount2');
						//$("#amount2").focus();
					}
				}
				else
					$("#amount2").val( "" );
			}

			// Cash
			if ($("#amount3").val().length > 0)
			{
				amount = $("#amount3").val();
				// Remove any extra spaces
				amount = removeSpaces(amount);
				//Remove commas that would cause problems in validation
				amount = removeCommas(amount);

				if ( amount != "" )
				{
					//rege = /^\d*\.?\d{0,2}$/
					Ok = rege.exec(amount);
					if ( Ok )
					{
						total += convertToCents(amount);
						amount = Number(amount);
						$("#amount3").val( format_number(amount,2) );
					}
					else
					{
						//alert('Values should be currency or blank.\nPlease correct this amount.');
						$("#amount3").focus();
						inlineMsg('amount3','<strong>Invalid Value: </strong>Values should be currency or blank.',3,'amount3');
					}
				}
				else
					$("#amount3").val( "" );
			}

			// Account transfer 
			//var bexists = eval($("#amount4").val());
			if ($("#amount4").length > 0)
			{
				if ($("#amount4").val().length > 0)
				{
					amount = $("#amount4").val();
					// Remove any extra spaces
					amount = removeSpaces(amount);
					//Remove commas that would cause problems in validation
					amount = removeCommas(amount);

					if ( amount != "" )
					{
						//rege = /^\d*\.?\d{0,2}$/
						Ok = rege.exec(amount);
						if ( Ok )
						{
							total += convertToCents(amount);
							amount = Number(amount);
							$("#amount4").val( format_number(amount,2) );
						}
						else
						{
							//alert('Values should be currency or blank.\nPlease correct this amount.');
							$("#amount4").focus();
							inlineMsg('amount4','<strong>Invalid Value: </strong>Values should be currency or blank.',3,'amount4');
						}
					}
					else
						$("#amount4").val( "" );
				}
			}
			// Other  
			//var bexists = eval(document.cartForm["amount8"]);
			if ($("#amount8").length > 0)
			{
				if ($("#amount8").val().length > 0)
				{
					amount = $("#amount8").val();
					// Remove any extra spaces
					amount = removeSpaces(amount);
					//Remove commas that would cause problems in validation
					amount = removeCommas(amount);

					if ( amount != "" )
					{
						//rege = /^\d*\.?\d{0,2}$/
						Ok = rege.exec(amount);
						if ( Ok )
						{
							total += convertToCents(amount);
							amount = Number(amount);
							$("#amount8").val( format_number(amount,2) );
						}
						else
						{
							//alert('Values should be currency or blank.\nPlease correct this amount.');
							$("#amount8").focus();
							inlineMsg('amount8','<strong>Invalid Value: </strong>Values should be currency or blank.',3,'amount8');
						}
					}
					else
						$("#amount8").val( "" );
				}
			}
			//alert(total);
			//alert(balance);

			//total = convertToCents(total);

			balance = convertToCents($("#purchasetotal").val()) - total;


			$("#total").html( format_number(convertToDollars(total),2) );
			$("#balancedue").html( format_number(convertToDollars(balance),2) );
			$("#amount").val( convertToDollars(total) );

			//alert(Number($("#amount").val()));
			//alert(Number($("#purchasetotal").val()));
			
		}
/////////////////////////////////////////////////////////////////////////////////////////////////

		function convertToCents(val)
		{
			retVal = val + "";
		    	if (retVal.indexOf(".") >= 0)
			{
				//alert(retVal.length - retVal.indexOf("."));
				if (retVal.length - retVal.indexOf(".") == 1)
				{
					retVal = retVal + "00";
				}
				else if (retVal.length - retVal.indexOf(".") == 2)
				{
					retVal = retVal + "0";
				}
			}
			else
			{
				retVal = retVal + "00";
			}

			return parseInt(retVal.replace(".",""));
		}

		function convertToDollars(val)
		{
			retVal = val + "";
			if (retVal.indexOf("-") == 0)
			{
				if (retVal.length == 2)
				{
					retVal = "-00" + retVal.replace("-","");
				}
				else if (retVal.length == 3)
				{
					retVal = "-0" + retVal.replace("-","");
				}
			}
			else
			{
				if (retVal.length == 1)
				{
					retVal = "00" + retVal;
				}
				else if (retVal.length == 2)
				{
					retVal = "0" + retVal;
				}
			}

			return retVal.substr(0,retVal.length-2) + "." + retVal.substr(retVal.length-2);
		}

		function addTotalOld()
		{
			var total = 0.00;
			var balance = 0.00;
			var rege;
			var Ok;

			rege = /^\d*\.?\d{0,2}$/

			//Charge
			if (document.cartForm.amount1.value != "")
			{
				//alert(document.cartForm.amount1.value);
				// Remove any extra spaces
				document.cartForm.amount1.value = removeSpaces(document.cartForm.amount1.value);
				//Remove commas that would cause problems in validation
				document.cartForm.amount1.value = removeCommas(document.cartForm.amount1.value);
				//rege = /^\d*\.?\d{0,2}$/
				Ok = rege.exec(document.cartForm.amount1.value);
				if ( Ok )
				{
					total += Number(document.cartForm.amount1.value);
					document.cartForm.amount1.value = Number(document.cartForm.amount1.value);
					document.cartForm.amount1.value = format_number(document.cartForm.amount1.value,2);
				}
				else
				{
					alert('Values should be currency or blank.\nPlease correct this amount.');
					document.cartForm.amount1.focus();
				}
			}
			// Check
			if (document.cartForm.amount2.value != "")
			{
				// Remove any extra spaces
				document.cartForm.amount2.value = removeSpaces(document.cartForm.amount2.value);
				//Remove commas that would cause problems in validation
				document.cartForm.amount2.value = removeCommas(document.cartForm.amount2.value);
				//rege = /^\d*\.?\d{0,2}$/
				Ok = rege.exec(document.cartForm.amount2.value);
				if ( Ok )
				{
					total += Number(document.cartForm.amount2.value);
					document.cartForm.amount2.value = Number(document.cartForm.amount2.value);
					document.cartForm.amount2.value = format_number(document.cartForm.amount2.value,2);
				}
				else
				{
					alert('Values should be currency or blank.\nPlease correct this amount.');
					document.cartForm.amount2.focus();
				}
			}
			// Cash
			if (document.cartForm.amount3.value != "")
			{
				// Remove any extra spaces
				document.cartForm.amount3.value = removeSpaces(document.cartForm.amount3.value);
				//Remove commas that would cause problems in validation
				document.cartForm.amount3.value = removeCommas(document.cartForm.amount3.value);
				//rege = /^\d*\.?\d{0,2}$/
				Ok = rege.exec(document.cartForm.amount3.value);
				if ( Ok )
				{
					total += Number(document.cartForm.amount3.value);
					document.cartForm.amount3.value = Number(document.cartForm.amount3.value);
					document.cartForm.amount3.value = format_number(document.cartForm.amount3.value,2);
				}
				else
				{
					alert('Values should be currency or blank.\nPlease correct this amount.');
					document.cartForm.amount3.focus();
				}
			}
			// Account transfer 
			var bexists = eval(document.cartForm["amount4"]);
			if (bexists)
			{
				if (document.cartForm.amount4.value != "")
				{
					// Remove any extra spaces
					document.cartForm.amount4.value = removeSpaces(document.cartForm.amount4.value);
					//Remove commas that would cause problems in validation
					document.cartForm.amount4.value = removeCommas(document.cartForm.amount4.value);
					//rege = /^\d*\.?\d{0,2}$/
					Ok = rege.exec(document.cartForm.amount4.value);
					if ( Ok )
					{
						total += Number(document.cartForm.amount4.value);
						document.cartForm.amount4.value = Number(document.cartForm.amount4.value);
						document.cartForm.amount4.value = format_number(document.cartForm.amount4.value,2);
					}
					else
					{
						alert('Values should be currency or blank.\nPlease correct this amount.');
						document.cartForm.amount4.focus();
					}
				}
			}
			// Other  
			var bexists = eval(document.cartForm["amount8"]);
			if (bexists)
			{
				if (document.cartForm.amount8.value != "")
				{
					// Remove any extra spaces
					document.cartForm.amount8.value = removeSpaces(document.cartForm.amount8.value);
					//Remove commas that would cause problems in validation
					document.cartForm.amount8.value = removeCommas(document.cartForm.amount8.value);
					//var rege8 = /^\d*\.?\d{0,2}$/
					var Ok = rege.exec(document.cartForm.amount8.value);
					if ( Ok )
					{
						total += Number(document.cartForm.amount8.value);
						document.cartForm.amount8.value = Number(document.cartForm.amount8.value);
						document.cartForm.amount8.value = format_number(document.cartForm.amount8.value,2);
					}
					else
					{
						alert('Values should be currency or blank.\nPlease correct this amount.');
						document.cartForm.amount8.focus();
					}
				}
			}

			balance = Number(document.cartForm.purchasetotal.value) - total;
			document.getElementById("total").innerHTML = format_number(total,2);
			document.getElementById("balancedue").innerHTML = format_number(balance,2);
			document.cartForm.amount.value = total;
			//alert(document.cartForm.amount.value);
		}

		function EditTeam( iCartId, iClassId )
		{
			location.href='regattateamsignup.asp?classid=' + iClassId + '&cartid=' + iCartId;
		}

		function EditMembers( iCartId, iClassId )
		{
			location.href='regattamembersignup.asp?classid=' + iClassId + '&cartid=' + iCartId;
		}

		function EditMerchandise( iCartId )
		{
			location.href='../merchandise/merchandiseoffered.asp?cartid=' + iCartId;
		}

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
<%
response.write "<p>" & vbcrlf

If sClassName <> "" And Not bIsRegattaTeam Then 
	'response.write "<a href=""class_signup.asp?classid=" & iClassId & "&timeid=" & iTimeId & """><img src=""../images/arrow_2back.gif"" align=""absmiddle"" border=""0"" />&nbsp;Purchase Another " & sClassName & "</a><br />" & vbcrlf
	response.write "<input type=""button"" name=""purchaseAnotherClass"" id=""purchaseAnotherClass"" value=""Purchase Another " & sClassName & """ class=""button"" onclick=""location.href='class_signup.asp?classid=" & iClassId & "&timeid=" & iTimeId & "'"" />" & vbcrlf
End If 

'response.write "<a href=""roster_list.asp""><img src=""../images/arrow_2back.gif"" align=""absmiddle"" border=""0"" />&nbsp;Purchase another Class/Event</a><br /><br />" & vbcrlf
response.write "<input type=""button"" name=""purchaseAnotherClassEvent"" id=""purchaseAnotherClassEvent"" value=""Purchase Another Class/Event"" class=""button"" onclick=""location.href='roster_list.asp'"" />" & vbcrlf
response.write "</p>" & vbcrlf
response.write "<h3>Shopping Cart</h3>" & vbcrlf
response.write "<br /><br />" & vbcrlf

If CLng(iUserId) <> 0 And CartHasItems() Then
	'The cart Is Not empty and we have a userid
	 response.write ShowUserInfo( iUserId )
End If 

'Select the items from the cart and display them Session("orgid")
Dim sSql, iRowCount, nRowTotal, nTotal, sRowClassName, fAmount

sSql = "SELECT C.cartid, C.classid, C.userid, isnull(C.familymemberid,0) as familymemberid, C.quantity, C.amount, C.optionid, C.classtimeid, "
sSql = sSql & " C.pricetypeid, C.buyorwait, C.sessionid, C.orgid, C.isdropin, C.dropindate, C.itemtypeid, C.rostergrade, C.rostershirtsize, C.rosterpantssize, "
sSql = sSql & " C.rostercoachtype, C.rostervolunteercoachname, C.rostervolunteercoachdayphone, C.rostervolunteercoachcellphone, "
sSql = sSql & " C.rostervolunteercoachemail, C.isregatta, ISNULL(C.regattateamid,0) AS regattateamid, I.itemtype, I.isshippingfee, I.issalestax "
sSql = sSql & " FROM egov_class_cart C, egov_item_types I "
sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND sessionid = " & session.sessionid
sSql = sSql & " ORDER BY I.cartdisplayorder, C.cartid"

iRowCount = 1
nTotal = CDbl(0)
fAmount = CDbl(0)
nRowTotal = 0 

Set oCart = Server.CreateObject("ADODB.Recordset")
oCart.Open sSql, Application("DSN"), 0, 1

If Not oCart.eof Then 
	response.write "<form name=""cartForm"" method=""post"" action=""class_purchase.asp"">" & vbcrlf
	response.write "<input type=""hidden"" name=""iUserId"" value="""                        & iUserId  & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iClassId"" value="""                       & iClassId & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterGrade"" value="""                   & oCart("rostergrade")     & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterShirtSize"" value="""               & oCart("rostershirtsize") & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterPantsSize"" value="""               & oCart("rosterpantssize") & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterCoachType"" value="""               & oCart("rostercoachtype") & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterVolunteerCoachName"" value="""      & oCart("rostervolunteercoachname")      & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterVolunteerCoachDayPhone"" value="""  & oCart("rostervolunteercoachdayphone")  & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterVolunteerCoachCellPhone"" value=""" & oCart("rostervolunteercoachcellphone") & """ />" & vbcrlf
	response.write "<input type=""hidden"" name=""iRosterVolunteerCoachEmail"" value="""     & oCart("rostervolunteercoachemail")     & """ />" & vbcrlf

	'response.write vbcrlf & "<div class=""shadow"">"
	response.write vbcrlf & "<table border=""0"" cellpadding=""3"" cellspacing=""0"" id=""cart_table"">"
	response.write vbcrlf & "<tr>"
	response.write "<th colspan=""2"">Item</th>"
	response.write "<th>Participant</th>"
	response.write "<th>Age</th>"
	response.write "<th>Qty</th>"
	response.write "<th>Price</th>"
	response.write "</tr>"

	Do While Not oCart.EOF
		'sItemType = GetItemType( oCart("itemtypeid") )
		sItemType = oCart("itemtype")

		If iRowCount Mod 2 = 0 Then 
			lcl_class_rec = " class=""alt_row"""
		Else 
			lcl_class_rec = ""
		End If 

		Select Case sItemType
			Case "recreation activity"
				response.write vbcrlf & "<tr" & lcl_class_rec & ">"
				sRowClassName    = getClassName(oCart("classid")) & " (" & GetActivityNo(oCart("classtimeid")) & ")"
				iPriceDiscountId = getClassPriceDiscountId(oCart("classid"))
				sDiscount        = GetDiscountPhrase( iPriceDiscountId )

				If sDiscount <> "" Then 
					sDiscount = "<br /><span class=""discounttext"">(" & sDiscount & ")</span>"
				End If 

				If oCart("isdropin") Then 
					iIsDropIn = 1 
				Else 
					iIsDropIn = 0
				End If 

				response.write "<td class=""firstcartcell"">"
				response.write "<input type=""button"" class=""button"" name=""remove"" value=""Remove"" onclick=""RemoveItem('" & FormatForjavascript(sRowClassName) & "', " & oCart("cartid") & ", " & oCart("classtimeid") & ", '" & oCart("buyorwait") & "', " & iIsDropIn & " )"" /> &nbsp;"
				response.write "</td>"
				response.write "<td> "

				response.write sRowClassName
				response.write sDiscount

				If oCart("isdropin") Then 
					response.write "<br />Dropin on " & oCart("dropindate")
				End If 

				response.write "<input type=""hidden"" name=""isdropin." & iRowCount & """ value=""" & iIsDropIn & """ />"

				If oCart("buyorwait") = "W" Then 
					response.write "<br />Wait List"
				End If 

				response.write "</td>"

				showFamilyMemberInfo oCart("familymemberid")

				response.write "<td align=""center"">"
				response.write "<input type=""hidden"" name=""cartid." & iRowCount & """ value=""" & oCart("cartid") & """ />"

				If CLng(oCart("optionid")) = 2 Then 
					'Handle ticketed events
					response.write "<input type=""text"" name=""quantity." & iRowCount & """ size=""5"" maxlength=""5"" value=""" & oCart("quantity") & """ />"
				Else 
					'handle reqistrations
					response.write "<input type=""hidden"" name=""quantity." & iRowCount & """ value=""" & oCart("quantity") & """ />"
					response.write oCart("quantity")
				End If 

				response.write "</td>" 

				'If oCart("buyorwait") = "B" Then 
				'purchase

				If checkForDiscountOverride(oCart("cartid")) Then 
					fAmount = GetCartUnitPrice(oCart("cartid"))
				Else 
					fAmount = GetCartItemPrice(oCart("cartid"))
				End If 
				
				' Add to total as a type double BEFORE the format currency
				nTotal = nTotal + fAmount
				response.write "      <td align=""right"">" & FormatCurrency(fAmount) & "</td>" & vbcrlf
				'nRowTotal = CLng(oCart("quantity")) * CDbl(oCart("amount"))
				'nRowTotal = CDbl(oCart("amount"))
				'response.write "<td align=""right"">" & FormatCurrency(nRowTotal) & "</td>"
				'Else 
				'wait list
				'response.write "<td align=""center"" colspan=""2"">Wait List</td>"
				'response.write "<td align=""center"">Wait List</td>"
				'End If 

				response.write "</tr>" & vbcrlf

			Case "regatta team"
				' Add Teams to Regatta
				sClassName = getClassName(oCart("classid"))
				response.write vbcrlf & "<tr" & lcl_class_rec & ">"
				response.write "<td class=""firstcartcell"">"
				response.write "<input type=""button"" class=""button"" name=""remove"" value=""Remove"" onclick=""RemoveRegattaTeamItem('" & FormatForjavascript(sTeamName) & "', " & oCart("cartid") & " )"" /> &nbsp;"
				response.write "</td>"
				response.write "<td align=""center"" valign=""top"">" & sClassName & "<br />" & GetTeamName( oCart("cartid") ) &"<br />" & GetTeamGroup( oCart("cartid") )
				response.write "<br /><input type=""button"" class=""button"" value=""Edit/View"" onclick=""EditTeam(" & oCart("cartid") & ", " & oCart("classid") & ")"" />"
				response.write "</td>"
				sCaptainName = GetCaptainName( oCart("cartid") ) 
				'If sCaptainName <> "" Then
				'	sCaptainName = sCaptainName '& "<br />(Captain)"
				'End If 
				response.write "<td align=""center"">" & sCaptainName & "</td>"
				response.write "<td align=""center"">&nbsp;</td>" ' Age column
				response.write "<td align=""center"">"
				response.write "<input type=""hidden"" name=""cartid." & iRowCount & """ value=""" & oCart("cartid") & """ />"
				response.write oCart("quantity")
				response.write "</td>"
				
				
				fAmount = GetCartItemPrice( oCart("cartid") )
				' Add to total as a type double BEFORE the format to currency
				nTotal = nTotal + fAmount

				response.write "<td align=""right"">" & FormatCurrency(fAmount) & "</td>"
				response.write "</tr>"

			Case "regatta member"
				' Add Members to Regatta Team
				sClassName = getClassName(oCart("classid"))
				response.write vbcrlf & "<tr" & lcl_class_rec & ">"
				response.write "<td class=""firstcartcell"">"
				response.write "<input type=""button"" class=""button"" name=""remove"" value=""Remove"" onclick=""RemoveRegattaTeamItem('" & FormatForjavascript(sClassName) & "', " & oCart("cartid") & " )"" /> &nbsp;"
				response.write "</td>"
				response.write "<td>" & sClassName
				response.write "&nbsp; <input type=""button"" class=""button"" value=""Edit/View"" onclick=""EditMembers(" & oCart("cartid") & ", " & oCart("classid") & ")"" />"
				response.write "</td>"
				response.write "<td align=""center"">" & GetRegattaTeamName( oCart("regattateamid") ) & "</td>"
				response.write "<td align=""center"">&nbsp;</td>" ' Age column
				response.write "<td align=""center"">"
				response.write "<input type=""hidden"" name=""cartid." & iRowCount & """ value=""" & oCart("cartid") & """ />"
				response.write oCart("quantity")
				response.write "</td>"
				fAmount = GetCartItemPrice( oCart("cartid") )
				' Add to total as a type double BEFORE the format to currency
				nTotal = nTotal + fAmount
				response.write "<td align=""right"">" &  FormatCurrency(fAmount) & "</td>"
				response.write "</tr>"

			Case "merchandise"
				response.write vbcrlf & "<tr" & lcl_class_rec & ">"
				response.write "<td class=""firstcartcell"">"
				response.write "<input type=""button"" class=""button"" name=""remove"" value=""Remove"" onclick=""RemoveMerchandiseItem(" & oCart("cartid") & " )"" /> &nbsp;"
				response.write "</td>"
				response.write "<td><strong>Merchandise Items</strong>"
				response.write "&nbsp; <input type=""button"" class=""button"" value=""Edit/View"" onclick=""EditMerchandise(" & oCart("cartid") & ")"" />"
				response.write "</td>"
				response.write "<td align=""center"">&nbsp;</td>" ' Participant column
				response.write "<td align=""center"">&nbsp;</td>" ' Age column

				' Quantity
				response.write "<td align=""center"">"
				response.write "<input type=""hidden"" name=""cartid." & iRowCount & """ value=""" & oCart("cartid") & """ />"
				response.write oCart("quantity")
				response.write "</td>"

				' Total amount
				fAmount = GetCartItemPrice( oCart("cartid") )
				' Add to total as a type double BEFORE the format to currency
				nTotal = nTotal + fAmount
				response.write "<td align=""right"">" & FormatCurrency(fAmount) & "</td>"
				response.write "</tr>"

				' Loop through the merchandise and list them out here
				ShowMerchandiseSelection oCart("cartid"), lcl_class_rec
			
			Case "shipping and handling fees"
				response.write vbcrlf & "<tr" & lcl_class_rec & ">"
				response.write "<td class=""firstcartcell"">&nbsp;</td>"
				response.write "<td><strong>Shipping/Handling Fees</strong></td>"
				response.write "<td align=""center"">&nbsp;</td>" ' Participant column
				response.write "<td align=""center"">&nbsp;</td>" ' Age column
				response.write "<td align=""center"">&nbsp;</td>" ' Qty column
				' Total amount
				fAmount =  CDbl(oCart("amount"))
				' Add to total as a type double BEFORE the format to currency
				nTotal = nTotal + fAmount
				response.write "<td align=""right"">" & FormatCurrency(fAmount) & "</td>"
				response.write "</tr>"

			Case "sales tax"
				response.write vbcrlf & "<tr" & lcl_class_rec & ">"
				response.write "<td class=""firstcartcell"">&nbsp;</td>"
				response.write "<td><strong>Sales Tax</strong></td>"
				response.write "<td align=""center"">&nbsp;</td>" ' Participant column
				response.write "<td align=""center"">&nbsp;</td>" ' Age column
				response.write "<td align=""center"">&nbsp;</td>" ' Qty column
				' Total amount
				fAmount = CDbl(oCart("amount"))
				' Add to total as a type double BEFORE the format to currency
				nTotal = nTotal + fAmount
				response.write "<td align=""right"">" & FormatCurrency(fAmount) & "</td>"
				response.write "</tr>"
		End Select 

		iRowCount = iRowCount + 1
		oCart.MoveNext
	Loop 

	response.write "<tr id=""carttotal"">"
	response.write "<td colspan=""6"" align=""right"">"
	response.write "<input type=""hidden"" id=""purchasetotal"" name=""purchasetotal"" value=""" & nTotal & """ />Cart Total: " & FormatCurrency(nTotal)
	response.write "</td>" 
	response.write "</tr>" 
	response.write "</table>" 
	'response.write "</div>" ' end of shadow div

	'Update Buttons
	response.write "<div id=""updatebutton"">" 
	response.write "<input type=""hidden"" name=""totalitems"" value=""" & (iRowCount - 1) & """ />"
	response.write "<input type=""button"" class=""button"" name=""update"" value=""Update Cart Quantities"" onclick=""UpdateCart();"" /> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;" 
	response.write "</div>"
	response.write "<br />" 

	'Payment
	response.write "<fieldset><legend><strong>Payment&nbsp;</strong></legend><br />"
	response.write "<input type=""hidden"" value=""0.00"" id=""amount"" name=""amount"" />"

	ShowPaymentChoices iUserId, nTotal

	response.write "<br /><br />"
	response.write "<input type=""button"" class=""button"" name=""complete"" value=""Complete Purchase"" onClick=""validate()"" />" 
	response.write "</fieldset>" 
	response.write "</form>"
  Else 
     response.write "<p><strong>There are no items in the Cart.</strong></p>"
  End If 

oCart.Close
Set oCart = Nothing 


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
' Sub ShowPaymentChoices( ByVal iUserId, sBalanceDue )
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentChoices( ByVal iUserId, sBalanceDue )
	Dim sSql, oPayments

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, requirescheckno, requirescitizenaccount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE O.paymenttypeid = P.paymenttypeid "
	sSql = sSql & " AND isadminmethod = 1 "
	sSql = sSql & " AND O.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oPayments = Server.CreateObject("ADODB.Recordset")
	oPayments.Open sSql, Application("DSN"), 0, 1

	If Not oPayments.EOF Then 
		response.write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""50%"">"
		response.write "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Citizen Location:</td>"
		response.write "<td>"

		ShowPaymentLocations  'In class_global_functions.asp

		response.write "</td>"
		response.write "</tr>"

		Do While Not oPayments.EOF
			response.write "<tr>"
			response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">"
			response.write oPayments("paymenttypename") & ": "
			response.write "</td>"
			response.write "<td>"
			response.write "<input type=""text"" value="""" id=""amount" & oPayments("paymenttypeid") & """ name=""amount" & oPayments("paymenttypeid") & """ size=""10"" maxlength=""9"" onblur=""addTotal()"" />"

			If oPayments("requirescheckno") Then 
				response.write "&nbsp;<strong>Check #: </strong>"
				response.write "<input type=""text"" value="""" id=""checkno"" name=""checkno"" size=""8"" maxlength=""8"" />"
			End If 

			If oPayments("requirescitizenaccount") Then 
				response.write "&nbsp; <strong>From:</strong>" 
				ShowFamilyAccounts iUserId
			End If 

			response.write "</td>"
			response.write "</tr>"

			oPayments.MoveNext
		Loop

		response.write "<tr>" 
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Payment Total:</td>"
		response.write "<td><span id=""total"">0.00</span></td>"
		response.write "</tr>" 
		response.write "<tr>"
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Balance Due:</td>"
		response.write "<td><span id=""balancedue"">" & FormatNumber(sBalanceDue,2,,,0) & "</span></td>" 
		response.write "</tr>" 
		response.write "<tr>" 
		response.write "<td class=""label"" align=""right"" nowrap=""nowrap"">Notes:</td>"
		response.write "<td><textarea name=""notes"" class=""purchasenotes""></textarea></td>"
		response.write "</tr>"
		response.write "</table>"
	End If 

	oPayments.Close
	Set oPayments = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function HasPayableAccounts( ByVal iUserId )
'--------------------------------------------------------------------------------------------------
Function HasPayableAccounts( ByVal iUserId )
	Dim sSql, oAccounts

	sSql = "SELECT Count(userid) as hits "
	sSql = sSql & " FROM egov_users "
	sSql = sSql & " WHERE accountbalance > 0.00 "
	sSql = sSql & " AND familyid = " & GetFamilyId( iUserId )

	set oAccounts = Server.CreateObject("ADODB.Recordset")
	oAccounts.Open sSql, Application("DSN"), 0, 1

	If CLng(oAccounts("hits")) > CLng(0) Then 
		HasPayableAccounts = True
	Else 
		HasPayableAccounts = False
	End If 

	oAccounts.close
	Set oAccounts = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Function getClassName( iClassId )
'--------------------------------------------------------------------------------------------------
Function getClassName( iClassId )
	Dim sSql, oName

	sSql = "SELECT classname FROM egov_class WHERE classid = " & iClassId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 0, 1

	getClassName = oName("classname") 

	oName.Close
	Set oName = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub getClassPurchaserName( ByVal iUserId )
'--------------------------------------------------------------------------------------------------
Function getClassPurchaserName( iUserId )
	Dim sSql, oName

	sSql = "SELECT userfname, userlname FROM egov_users WHERE userid = " & iUserId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 1, 3

	getClassPurchaserName = oName("userfname") & " " & oName("userlname")

	oName.Close
	Set oName = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub showFamilyMemberInfo( iFamilyMemberId )
'--------------------------------------------------------------------------------------------------
Sub showFamilyMemberInfo( iFamilyMemberId )
	Dim sSql, oName
	
	If iFamilyMemberId <> 0 Then 
		sSql = "SELECT firstname, lastname, birthdate "
		sSql = sSql & " FROM egov_familymembers "
		sSql = sSql & " WHERE familymemberid = " & iFamilyMemberId

		Set oName = Server.CreateObject("ADODB.Recordset")
		oName.Open sSql, Application("DSN"), 1, 3

		response.write "      <td>" & oName("firstname") & " " & oName("lastname") & "</td>" & vbcrlf
		response.write "      <td align=""center"">" & vbcrlf

		If IsNull(oName("birthdate")) Then 
			response.write "&nbsp;"
		Else 
			'response.write DateDiff("yyyy", oName("birthdate"), Now())
			response.write GetChildAge(oName("birthdate"))
		End If 
		response.write "      </td>" & vbcrlf

		oName.Close
		Set oName = Nothing 

	Else 
		response.write "      <td>&nbsp;</td>" & vbcrlf
		response.write "      <td>&nbsp;</td>" & vbcrlf
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ShowUserInfo( iUserId )
'--------------------------------------------------------------------------------------------------
Function ShowUserInfo( iUserId )
	 Dim oCmd, sResidentDesc, sUserType, oUser

	lcl_return = ""

	sUserType = GetUserResidentType( iUserId )

	'If they are not one of these (R, N), we have to figure which they are
	If sUserType <> "R" AND sUserType <> "N" Then 
		'This leaves E and B - See if they are a resident, also
		sUserType = GetResidentTypeByAddress(iUserId, Session("orgid"))
	End If 

	sResidentDesc = GetResidentTypeDesc(sUserType)

	sSql = "SELECT userfname, userlname, useraddress, useraddress2, usercity, userstate, userzip, usercountry, useremail, userhomephone, "
	sSql = sSql & " userworkphone, userfax, userbusinessname, userpassword, userregistered, residenttype, residencyverified, "
	sSql = sSql & " registrationblocked, blockeddate, blockedadminid, blockedexternalnote, blockedinternalnote "
	sSql = sSql & " FROM egov_users "
	sSql = sSql & " WHERE userid = " & iUserId

	set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open sSql, Application("DSN"), 3, 1

	lcl_return = "<table border=""0"" cellpadding=""0"" cellspacing=""5"" id=""signupuserinfo"">" & vbcrlf
	lcl_return = lcl_return & "<tr>" & vbcrlf
	lcl_return = lcl_return & "<td align=""right"" valign=""top"">Name:</td>" & vbcrlf
	lcl_return = lcl_return & "<td>" & oUser("userfname") & " " & oUser("userlname") & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong>"

	If Not oUser("residencyverified") And oUser("residenttype") = "R" Then 
		If lcl_orghasfeature_residency_verification then
			lcl_return = lcl_return & " (not verified)"
		End If 
	End If 

	lcl_return = lcl_return & "</td>" & vbcrlf
	lcl_return = lcl_return & "</tr>" & vbcrlf
	lcl_return = lcl_return & "<tr>" & vbcrlf
	lcl_return = lcl_return & "<td align=""right"" valign=""top"">Email:</td>" & vbcrlf
	lcl_return = lcl_return & "<td>" & oUser("useremail") & "</td>" & vbcrlf
	lcl_return = lcl_return & "</tr>" & vbcrlf
	lcl_return = lcl_return & "<tr>" & vbcrlf
	lcl_return = lcl_return & "<td align=""right"" valign=""top"">Phone:</td>" & vbcrlf
	lcl_return = lcl_return & "<td>" & FormatPhone(oUser("userhomephone")) & "</td>" & vbcrlf
	lcl_return = lcl_return & "</tr>" & vbcrlf
	lcl_return = lcl_return & "<tr>" & vbcrlf
	lcl_return = lcl_return & "<td align=""right"" valign=""top"">Address:</td>" & vbcrlf
	lcl_return = lcl_return & "<td>" & oUser("useraddress") & "<br />" 

	If oUser("useraddress2") <> "" Then 
		lcl_return = lcl_return & oUser("useraddress2") & "<br />" & vbcrlf
	End If 

	If oUser("usercity") <> "" Or oUser("userstate") <> "" Or oUser("userzip") <> "" Then 
		lcl_return = lcl_return & oUser("usercity") & ", " & oUser("userstate") & " " & oUser("userzip") & vbcrlf
	End If 

	lcl_return = lcl_return & "      </td>" & vbcrlf
	lcl_return = lcl_return & "  </tr>" & vbcrlf
	'lcl_return = lcl_return & "<tr><td width=""85"" align=""right"" valign=""top"">Business:</td><td>" & oUser("userbusinessname") & "</td></tr>"
	lcl_return = lcl_return & "</table>" & vbcrlf

	oUser.Close
	Set oUser = Nothing 

	ShowUserInfo = lcl_return
	
End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowPaymentTypes()
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentTypes()
	Dim sSql, oTypes

	sSql = "SELECT P.paymenttypeid, P.paymenttypename "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE O.paymenttypeid = P.paymenttypeid "
	sSql = sSql & " AND isadminmethod = 1 "
	sSql = sSql & " AND O.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	'response.write sSql & "<br />" & vbcrlf

	Set oTypes = Server.CreateObject("ADODB.Recordset")
	oTypes.Open sSql, Application("DSN"), 3, 1

	Do While Not oTypes.EOF
		response.write "  <option value=""" & oTypes("paymenttypeid") & """>" & oTypes("paymenttypename") & "</option>" & vbcrlf
		oTypes.MoveNext
	Loop 

	oTypes.Close
	Set oTypes = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowMerchandiseSelection( iCartId, sRowClass )
'--------------------------------------------------------------------------------------------------
Sub ShowMerchandiseSelection( iCartId, sRowClass )
	Dim sSql, oRs

	sSql = "SELECT M.merchandise, MC.merchandisecolor, MC.isnocolor, MS.merchandisesize, MS.isnosize, I.quantity, I.price "
	sSql = sSql & " FROM egov_class_cart_merchandiseitems I, egov_merchandisecatalog C, egov_merchandise M, "
	sSql = sSql & " egov_merchandisecolors MC, egov_merchandisesizes MS "
	sSql = sSql & " WHERE I.merchandisecatalogid = C.merchandisecatalogid "
	sSql = sSql & " AND C.merchandiseid = M.merchandiseid AND C.merchandisecolorid = MC.merchandisecolorid "
	sSql = sSql & " AND C.merchandisesizeid = MS.merchandisesizeid AND I.cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr" & sRowClass & ">"
		response.write "<td class=""firstcartcell"">&nbsp;</td>"
		response.write "<td>"
		response.write "<table cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr><th>Item</th><th>Price</th><th>Qty</th><th>Total</th></tr>"
		Do While Not oRs.EOF
			response.write "<tr" & sRowClass & ">"
			response.write "<td>"
			response.write oRs("merchandise")
			If Not oRs("isnocolor") Then
				response.write ",&nbsp;" & oRs("merchandisecolor")
			End If 
			If Not oRs("isnosize") Then
				response.write ",&nbsp;" & oRs("merchandisesize")
			End If 
			response.write "</td>"
			response.write "<td align=""center"">"
			response.write FormatNumber(oRs("price"),2,,,0)
			response.write "</td>"
			response.write "<td align=""center"">"
			response.write oRs("quantity")
			response.write "</td>"
			response.write "<td align=""right"">"
			response.write FormatNumber((oRs("quantity") * oRs("price")),2,,,0)
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
		response.write "</table>"
		response.write "</td>"
		response.write "<td>&nbsp;</td>" ' Participant
		response.write "<td>&nbsp;</td>" ' Age
		response.write "<td>&nbsp;</td>" ' QTY
		response.write "<td>&nbsp;</td>" ' Price
		response.write "</tr>"
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' string GetTeamName( iCartId )
'--------------------------------------------------------------------------------------------------
Function GetTeamName( ByVal iCartId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(regattateam,'') AS regattateam FROM egov_class_cart_regattateams "
	sSql = sSql & "WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetTeamName = oRs("regattateam")
	Else
		GetTeamName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetTeamGroup( iCartId )
'--------------------------------------------------------------------------------------------------
Function GetTeamGroup( ByVal iCartId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(G.regattateamgroup,'') AS regattateamgroup "
	sSql = sSql & "FROM egov_class_cart_regattateams T, egov_regattateamgroups G "
	sSql = sSql & "WHERE T.regattateamgroupid = G.regattateamgroupid AND T.cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetTeamGroup = oRs("regattateamgroup")
	Else
		GetTeamGroup = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
