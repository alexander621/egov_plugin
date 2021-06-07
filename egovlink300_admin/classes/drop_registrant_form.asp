<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: DROP_REGISTRANT_FORM.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 04/26/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/26/2006   JOHN STULLENBERGER - INITIAL VERSION
' 1.2	03/18/2011	Steve Loar - Fixed bug in changing amount and immediately clicking Drop where total off
' 1.3	11/6/2013	Steve Loar - Adding drop reason picks
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>
<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="classes.css" />
	
	<script src="../scripts/jquery-1.7.2.min.js"></script>

	<script src="tablesort.js"></script>
	<script src="../scripts/layers.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/removespaces.js"></script>
	<script src="../scripts/removecommas.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
		<!--

		var global_valfield;	// retain valfield for timer thread
		// --------------------------------------------
		//                  setfocus
		// Delayed focus setting to get around IE bug
		// --------------------------------------------

		function setFocusDelayed()
		{
		  global_valfield.focus();
		}

		function setfocus(valfield)
		{
		  // save valfield in global variable so value retained when routine exits
		  global_valfield = valfield;
		  setTimeout( 'setFocusDelayed()', 100 );
		}

		function confirm_drop(iclasslistid, srostername)
		{
			if (confirm("Are you sure you want to drop (" + srostername + ")?"))
				{ 
					// DELETE HAS BEEN VERIFIED
					location.href='drop_registrant.asp?classid=<%=request("classid")%>&timeid=<%=request("timeid")%>&iclasslistid=' + iclasslistid;
				}
		}

		function ValidatePrice( oPrice )
		{
			var bValid = true;
			var total = 0.00;
			
			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			if (oPrice.value == "")
			{
				oPrice.value = 0.00;
			}

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
			if ( bValid == false )
			{
				alert('Amounts should be positive numbers in currency format.\nPlease correct this.');
				setfocus(oPrice);
				return false;
			}

			// Calculate a new total price
			if (document.frmDrop.pricetypeid.length)   // If there is more than one price checkbox
			{
				var checklength = document.frmDrop.pricetypeid.length;
				var i = checklength - 1;

				for (l = 0; l <= i; l++)
				{
					if (document.frmDrop.pricetypeid[l].checked)
					{ 
						total += Number(eval('document.frmDrop.amount' + document.frmDrop.pricetypeid[l].value + '.value'));
					}
				}
			}
			else   // There is only one price checkbox
			{
				if (document.frmDrop.pricetypeid.checked)
				{
					total += Number(eval('document.frmDrop.amount' + document.frmDrop.pricetypeid.value + '.value'));
				}
			}
			//total = Number(document.frmDrop.totalpaid.value);

			// Take away any refund fee
			//var exists = eval(document.frmDrop["refundfee"]);
			//if (exists)
			//{
			//	if (document.frmDrop.refundfeeid.checked)
			//	{ 
					total -= Number(document.frmDrop.refundfee.value);
			//	}
			//}

			document.frmDrop.totalrefund.value = total;
			document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
			
			return true;
		}

		function UpdatePriceTotal( iPrice, bChecked )
		{
			var total = 0.00;

			if (iPrice != "")
			{
				total = Number(document.frmDrop.totalprice.value);
				if (bChecked)
				{
					total += Number(iPrice);
				}
				else
				{
					total -= Number(iPrice);
				}
				document.frmDrop.totalprice.value = total;
				document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
			}
		}

		function validateForm()
		{
			// check the reason drop down pick here
			if ( parseInt($("#dropreasonid").val()) == parseInt(0) )
			{
				// show error msg and return
				inlineMsg($("#dropreasonid").attr('id'),'<strong>Drop Cannot Complete: </strong>Please select the reason for dropping.',8,"dropreasonid");
				return false;
			}
			
			var total = 0.00;

			// Calculate a new total price
			if (document.frmDrop.pricetypeid.length)   // If there is more than one price checkbox
			{
				var checklength = document.frmDrop.pricetypeid.length;
				var i = checklength - 1;

				for (l = 0; l <= i; l++)
				{
					if (document.frmDrop.pricetypeid[l].checked)
					{ 
						total += Number(eval('document.frmDrop.amount' + document.frmDrop.pricetypeid[l].value + '.value'));
					}
				}
			}
			else   // There is only one price checkbox
			{
				if (document.frmDrop.pricetypeid.checked)
				{
					total += Number(eval('document.frmDrop.amount' + document.frmDrop.pricetypeid.value + '.value'));
				}
			}

			total -= Number(document.frmDrop.refundfee.value);

			document.frmDrop.totalrefund.value = total;
			document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
			
			document.frmDrop.submit();
		}

		//-->
	</script>

</head>

<body>

 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<%
Dim iclassid, itimeid, iclasslistid, iqty, bHasRefundDebitAccount, cRefundFee, iRefundDebitId, sRefundName, oRegistrant

iclassid = CLng(request("classid"))
itimeid = CLng(request("timeid"))
iclasslistid = CLng(request("classlistid"))
iqty = clng(request("iqty"))
iRefundDebitId = 0
sRefundName = ""

bHasRefundDebitAccount = OrgHasRefundDebit( )

If bHasRefundDebitAccount Then 
	cRefundFee = GetRefundFee( iRefundDebitId, sRefundName )
	'iRefundDebitId = GetRefundId()
Else
	cRefundFee = CDbl(0.00)
End If 

%>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>Recreation: Drop Registrant</strong></font><br /><br />
	
	<input type="button" class="button" value="<< <%=langBackToStart%>" onclick="javascript:history.go(-1);" />
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: CLASS INFORMATION -->
<p> <% DisplayItem request("classid"), request("timeid") %> </p>
<!--END: CLASS INFORMATION -->


<!--BEGIN: DROP FORM-->
<%

' GET INFORMATION FOR THIS REGISTRANT
sSql = "SELECT lastname, firstname, userlname, userfname, amount, description, quantity, userid FROM egov_class_roster "
sSql = sSql & " WHERE classid = " & iclassid & " AND classtimeid = " & itimeid & " AND classlistid = " & iclasslistid 

Set oRegistrant = Server.CreateObject("ADODB.Recordset")
oRegistrant.Open sSql, Application("DSN"), 3, 1

If Not oRegistrant.EOF Then
	If ClassRequiresRegistration( iClassId ) Then
		sName = oRegistrant("firstname") & " " & oRegistrant("lastname")
	Else
		sName = oRegistrant("userfname") & " " & oRegistrant("userlname")
	End If 
	sHeadName = oRegistrant("userfname") & " " & oRegistrant("userlname")
	iHeadUserId = oRegistrant("userid")
	sResidentTypeDesc = oRegistrant("description")
	iQty = oRegistrant("quantity")
	iPaymentId = GetPaymentId( iclasslistid )
	curAmount = oRegistrant("amount") 
	curRefund = curAmount - 25
	If curRefund  < 0 Then
		curRefund  = 0
	End If
End If

oRegistrant.Close()
Set oRegistrant = Nothing

%>

<form name="frmDrop" action="drop_registrant.asp" method="post">
	<input type="hidden" name="iclasslistid" value="<%=iclasslistid%>" />
	<input type="hidden" name="classid" value="<%=iclassid%>" />
	<input type="hidden" name="timeid" value="<%=itimeid%>" />
	<input type="hidden" name="quantity" value="<%=iQty%>" />
	<input type="hidden" name="iUserId" value="<%=iHeadUserId%>" />
	<input type="hidden" name="oldpaymentid" value="<%=iPaymentId%>" />

	<p><strong>Name:</strong> <%=sName%> ( <%=sResidentTypeDesc%> )</p>

	<p><strong>Head of Household: </strong><%=sHeadName%></p>

	<p><strong>Quantity:</strong> <%=iQty%> </p>

	<fieldset><legend><strong> Payment Details </strong></legend>
		<% ShowPaymentLedgerDetails iclasslistid %>
	</fieldset>

	<fieldset>
		<legend><strong> Refund Details </strong></legend>
		<% ShowPurchaseLedgerDetails iclasslistid, cRefundFee, bHasRefundDebitAccount, iRefundDebitId, sRefundName  %>
	</fieldset>

	<p>
		<strong>Citizen Location:</strong> &nbsp; <% ShowPaymentLocations  ' In class_global_functions.asp' %>
	</p>

	<p>
		<strong>Apply the refund to:</strong> &nbsp; <% ShowRefundChoices iHeadUserId %>
	</p>

	<div class="dropinfoblock">
		<% ShowReasonPicks %>
	</div>

	<p>
		<table border="0" cellpadding="0" cellspacing="0" id="dropnotes">
		<tr><td valign="top" align="right" id="notestag"><strong>Notes:</strong> &nbsp; </td><td><textarea name="notes" class="purchasenotes"></textarea></td></tr>
		</table>
	</p>

	<p><input class="button" type="button" name="complete" value="Drop" onclick="validateForm();" /></p>

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
'  DISPLAYITEM CLASSID, TIMEID
'--------------------------------------------------------------------------------------------------
 Sub DisplayItem( ByVal iClassId, ByVal iTimeId )
	Dim sSql, oRs

	' GET SELECTED FACILITY INFORMATION
	sSql = "SELECT classname, activityno "
	sSql = sSql & "FROM egov_class "
	sSql = sSql & "LEFT JOIN egov_class_time ON egov_class.classid = egov_class_time.classid "
	sSql = sSql & "LEFT JOIN egov_class_instructor ON egov_class_time.instructorid = egov_class_instructor.instructorid "
	sSql = sSql & "WHERE egov_class.classid = " &  iClassId & " AND egov_class_time.timeid = " & iTimeId 
	sSql = sSql & " ORDER BY noenddate DESC, egov_class.startdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

    ' DISPLAY ITEM INFORMATION
    If Not oRs.EOF Then
		Response.Write("<h3>" &  oRs("classname") & " &nbsp; ( " & oRs("activityno") & " )</h3>" & vbCrLf)
	End If

    ' CLOSE OBJECTS
	oRs.Close 
    Set  oRs = Nothing 

 End Sub


'--------------------------------------------------------------------------------------------------
'  integer GetPaymentId( iClassListId )
'--------------------------------------------------------------------------------------------------
Function GetPaymentId( ByVal iClassListId )
	Dim sSql, oRs

	sSql = "SELECT paymentid FROM egov_class_list WHERE classlistid = " & iClassListId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentId = oRs("paymentid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPurchaseLedgerDetails iClassListId, cRefundFee, bHasRefundDebitAccount, iRefundDebitId, sRefundName 
'--------------------------------------------------------------------------------------------------
Sub ShowPurchaseLedgerDetails( ByVal iClassListId, ByVal cRefundFee, ByVal bHasRefundDebitAccount, ByVal iRefundDebitId, ByVal sRefundName )
	Dim sSql, oRs, cTotalRefund, cTotalPaid

	cTotalRefund = 0.00
	cTotalPaid = 0.00

	sSql = "SELECT C.paymentid, A.accountid, A.amount, itemid, pricetypename, itemtypeid, A.pricetypeid "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_list C, egov_price_types P "
	sSql = sSql & " WHERE  A.paymentid = C.paymentid AND A.itemid = C.classlistid "
	sSql = sSql & " AND P.pricetypeid = A.pricetypeid AND ispaymentaccount = 0 "  
	sSql = sSql & " AND itemtypeid = 1 AND A.itemid = " & iClassListId & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		iMinPricetype = clng(oRs("pricetypeid"))
		iMaxPriceType = clng(oRs("pricetypeid"))
		response.write "<table id=""droppricetable"" border=""0"" cellpadding=""2"" cellspacing=""0"">"
		Do While Not oRs.EOF
			If clng(oRs("pricetypeid")) < iMinPricetype Then
				iMinPricetype = clng(oRs("pricetypeid"))
			End If 
			If clng(oRs("pricetypeid")) > iMaxPriceType Then
				iMaxPriceType = clng(oRs("pricetypeid"))
			End If 
			
			response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap"" valign=""top"">"
			response.write "<span class=""pricecheck""><input type=""checkbox"" id=""pricetypeid" & oRs("pricetypeid") & """ name=""pricetypeid"" value=""" & oRs("pricetypeid") & """ checked=""checked"" /></span>"
			response.write oRs("pricetypename") & "</td><td class=""signcol"">+</td>"
			response.write "<td class=""priceentrytd"" valign=""top"">"
			response.write "<input type=""text"" id=""amount" & oRs("pricetypeid") & """ name=""amount" & oRs("pricetypeid") & """ value=""" & Replace(FormatNumber(CDbl(oRs("amount")),2),",","") & """ size=""10"" maxlength=""9"" onchange=""ValidatePrice(this);"" />"
			'response.write "<input type=""text"" id=""amount" & oRs("pricetypeid") & """ name=""amount" & oRs("pricetypeid") & """ value=""" & FormatNumber(CDbl(oRs("amount")),2) & """ size=""6"" maxlength=""6"" onchange=""ValidatePrice(this);"" />"

			cTotalRefund = cTotalRefund + CDbl(oRs("amount"))
			cTotalPaid = cTotalPaid + CDbl(oRs("amount"))
			
			response.write "</td>"
			response.write "<td>"
			response.write FormatCurrency(CDbl(oRs("amount"))) 
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
		If bHasRefundDebitAccount Then
			' Show the refund fee
			
			response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap"" valign=""top"">"
			'response.write "<input type=""checkbox"" checked=""checked"" id=""refundfeeid"" name=""refundfeeid"" value=""" & iRefundDebitId & """ onClick=""UpdatePriceTotal(-document.frmDrop.refundfee.value, this.checked);"" /> &nbsp; "
			response.write sRefundName & "</td><td class=""signcol"">&ndash;</td>"
			response.write "<td class=""priceentrytd"" valign=""top""><input type=""text"" id=""refundfee"" name=""refundfee"" value=""" & Replace(FormatNumber(CDbl(cRefundFee),2),",","") & """ size=""10"" maxlength=""9"" onchange=""ValidatePrice(this);"" />"
			response.write vbcrlf & "<input type=""hidden"" id=""refundfeeid"" name=""refundfeeid"" value=""" & iRefundDebitId & """ />"
			response.write "</td>"
			response.write "<td>- " & FormatCurrency(cRefundFee) & " (default fee amount)</td>"
			response.write "</tr>"
			cTotalRefund = cTotalRefund - CDbl(cRefundFee)
		End If 
		response.write vbcrlf & "<tr><td><strong>Total Refund</strong></td><td class=""signcol"">&nbsp;</td><td><span id=""displaytotalprice"">" & Replace(FormatNumber(cTotalRefund,2),",","") & "</span></td></tr>"
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "<input type=""hidden"" id=""totalrefund"" name=""totalrefund"" value=""" & cTotalRefund & """ />"
		response.write vbcrlf & "<input type=""hidden"" id=""totalpaid"" name=""totalpaid"" value=""" & CDbl(cTotalPaid) & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMinPricetype & """ />"
		response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMaxPriceType & """ />"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
'  ShowPaymentLedgerDetails iClassListId 
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLedgerDetails( ByVal iClassListId )
	Dim sSql, oRs, cTotalPaid, bCCUsed

	cTotalPaid = 0.00
	bCCUsed = False 

	sSql = "SELECT P.paymenttypename, A.amount, V.checkno, A.accountid, P.ispublicmethod, P.isadminmethod, P.requirescheckno, P.requirescitizenaccount, P.requirescreditcard "
	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_list C, egov_paymenttypes P, egov_verisign_payment_information V "
	sSql = sSql & " WHERE  A.paymentid = C.paymentid AND C.paymentid = V.paymentid AND A.paymenttypeid = P.paymenttypeid "
	sSql = sSql & " AND A.ledgerid = V.ledgerid AND A.paymenttypeid = V.paymenttypeid AND A.ispaymentaccount = 1 "
	sSql = sSql & " AND C.classlistid = " & iClassListId & " ORDER BY P.displayorder"
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""droppricename"" border=""0"" cellpadding=""2"" cellspacing=""0"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap"" valign=""top"">"
			response.write oRs("paymenttypename") & "</td>"
			response.write "<td>" & FormatCurrency(oRs("amount")) 
			' look up the account name if citizen account
			If oRs("requirescitizenaccount") Then
				response.write " &nbsp; From: " & GetCitizenName( oRs("accountid") )
			End If 
			If oRs("requirescreditcard") Then
				bCCUsed = True 
			End If 
			response.write "</td>"
			response.write "</tr>"
			cTotalPaid = cTotalPaid + CDbl(oRs("amount"))
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "<tr><td><strong>Total Paid</strong></td><td><span id=""displaytotalpaid"">" & FormatCurrency(CDbl(cTotalPaid),2) & "</span>"
		'response.write "<input type=""hidden"" name=""totalpaid"" value=""" & CDbl(cTotalPaid) & """ />"
		If bCCUsed Then
			response.write "<input type=""hidden"" name=""isccrefund"" value=""1"" />"
		Else 
			response.write "<input type=""hidden"" name=""isccrefund"" value=""0"" />"
		End If 
		response.write "</td></tr>"
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
'  Sub ShowRefundChoices( iHeadUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowRefundChoices( ByVal iHeadUserId )
	Dim sSql, oRs

	response.write vbcrlf & "<select name=""accountid"">"
	response.write vbcrlf & "<option value=""0"" selected=""selected"">" & GetRefundName() & "</option>"

	If OrgHasFeature( "citizen accounts" ) Then 
		sSql = "SELECT userfname, userlname, userid, ISNULL(accountbalance,0.00) AS accountbalance "
		sSql = sSql & " FROM egov_users WHERE isdeleted = 0  and familyid = " & GetFamilyId( iHeadUserId )
		sSql = sSql & " ORDER BY userlname, userfname"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("userid") & """>" & oRs("userfname") & " " & oRs("userlname") & " (" & FormatNumber(oRs("accountbalance"),2) & ") " & "</option>"
			oRs.MoveNext
		Loop 

		oRs.Close
		Set oRs = Nothing 
	End If 

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
'  boolean bHasRefundAccount = OrgHasRefundDebit()
'--------------------------------------------------------------------------------------------------
Function OrgHasRefundDebit( )
	Dim sSql, oRs

	sSql = "SELECT COUNT(P.paymenttypeid) AS hits From egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE P.paymenttypeid = O.paymenttypeid and P.isrefunddebit = 1 AND P.isforclasses = 1 "
	sSql = sSql & " AND O.orgid = " & Session("OrgID")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If clng(oRs("hits")) > clng(0) Then
		OrgHasRefundDebit = True 
	Else
		OrgHasRefundDebit = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
'  double dDefaultAmount = GetRefundFee( iRefundDebitId, sRefundName )
'--------------------------------------------------------------------------------------------------
Function GetRefundFee( ByRef iRefundDebitId, ByRef sRefundName )
	Dim sSql, oAccount

	sSql = "SELECT P.paymenttypeid, P.paymenttypename, ISNULL(O.defaultamount,0.00) AS defaultamount "
	sSql = sSql & " FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE P.paymenttypeid = O.paymenttypeid AND P.isrefunddebit = 1 AND P.isforclasses = 1 "
	sSql = sSql & " AND O.orgid = " & Session("OrgID")

	Set oAccount = Server.CreateObject("ADODB.Recordset")
	oAccount.Open sSql, Application("DSN"), 0, 1

	If Not oAccount.EOF Then
		GetRefundFee = CDbl(oAccount("defaultamount"))
		iRefundDebitId = oAccount("paymenttypeid")
		sRefundName = oAccount("paymenttypename")
	Else
		GetRefundFee = CDbl(0.00)
		iRefundDebitId = 0
		sRefundName = ""
	End If 

	oAccount.Close
	Set oAccount = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
'  Function GetRefundId( )
'--------------------------------------------------------------------------------------------------
Function GetRefundId( )
	Dim sSql, oRs

	sSql = "SELECT P.paymenttypeid FROM egov_paymenttypes P, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE P.paymenttypeid = O.paymenttypeid AND P.isrefunddebit = 1 AND P.isforclasses = 1 "
	sSql = sSql & " AND O.orgid = " & Session("OrgID")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundId = oRs("paymenttypeid") 
	Else
		GetRefundId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
'  ShowReasonPicks
'--------------------------------------------------------------------------------------------------
Sub ShowReasonPicks()
	Dim sSql, oRs

	sSql = "SELECT dropreasonid, dropreason FROM egov_class_dropreasons ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Response.write vbcrlf & "<label for=""dropreasonid"">Reason: </label><select id=""dropreasonid"" name=""dropreasonid"">"
		Response.write vbcrlf & "<option value=""0"">Select The Reason For Dropping</option>"
		Do While Not oRs.EOF 
			Response.write vbcrlf & "<option value=""" & oRs("dropreasonid") & """>" & oRs("dropreason") & "</option>"
			oRs.MoveNext
		Loop
		Response.write vbcrlf & "</select>"
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>

