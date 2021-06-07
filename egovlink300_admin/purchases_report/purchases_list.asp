<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: purchases_list.asp
' AUTHOR: Steve Loar
' CREATED: 07/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a list of purchases by citizen
'
' MODIFICATION HISTORY
' 1.0   07/31/2006	Steve Loar - INITIAL VERSION
' 1.1	10/06/2006	Steve Loar - Security, Header and nav changed
' 2.0	07/27/2010	Steve Loar - Modifications for PNP integration
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oList, orderBy, subTotals, showDetail, fromDate, toDate, today, sUserlname, iPurchaseCount, bgcolor
Dim iUserId, sResults, sSearchStart, sSearchName


sLevel     = "../"     'Override of value from common.asp

' check if the feature is available and they have right to this page.
PageDisplayCheck "citizen rec purchases", sLevel	' In common.asp

lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

orderBy    = request("orderBy")
subTotals  = request("subTotals")
showDetail = request("showDetail")
fromDate   = request("fromDate")
toDate     = request("toDate")
today      = Date()
sUserlname = request("userlname")

If request("userid") <> "" Then
   iUserId = request("userid")
Else
   If Session("eGovUserId") <> "" Then
      iUserid = Session("eGovUserId")
   Else
      iUserId = GetFirstUserId()
   End If
End If
Session("eGovUserId") = iUserId

'See if a search term was passed
If request("searchname") <> "" Then 
	  sSearchName = request("searchname")
Else
  	sSearchName = ""
End If 

If request("results") <> "" Then
	  sResults = request("results")
Else
  	sResults = ""
End If 

If request("searchstart") <> "" Then 
	  sSearchStart = request("searchstart")
Else 
  	sSearchStart = -1
End If 

If request("receipt_num") <> "" Then 
   iReceiptNum = CLng(Trim(request("receipt_num")))
Else 
   iReceiptNum = ""
End If 

If orderBy = "" Or IsNull(orderBy) Then
	orderBy  = "date" 
End If

If toDate = "" Or IsNull(toDate) Then
	toDate   = dateAdd("d",0,today) 
End If

If fromDate = "" Or IsNull(fromDate) Then 
	fromDate = DateSerial(Year(Now()),1,1) 
End If

toDate = dateAdd("d",1,toDate)

Set oList = Server.CreateObject("ADODB.Recordset")

oList.Fields.Append "purchaseday",   adInteger,    , adFldUpdatable
oList.Fields.Append "purchasemonth", adInteger,    , adFldUpdatable
oList.Fields.Append "purchaseyear",  adInteger,    , adFldUpdatable
oList.Fields.Append "purchasedate",  adVariant,  10, adFldUpdatable
'oList.Fields.Append "purchasedate", adVarChar, 10, adFldUpdatable
oList.Fields.Append "whatpurchased", adVarChar, 255, adFldUpdatable
oList.Fields.Append "purchaser",     adVarChar, 255, adFldUpdatable
oList.Fields.Append "amount",        adCurrency,   , adFldUpdatable
oList.Fields.Append "paymentid",     adInteger,    , adFldUpdatable
oList.Fields.Append "detailurl",     adVarChar, 255, adFldUpdatable
'oList.CursorType = adOpenDynamic
oList.CursorLocation = 3

oList.Open 


%>
<html>
<head>
	<title><%=langBSPayments%></title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="purchasesreport.css" />

	<script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script type="text/javascript" src="../scripts/selectAll.js"></script>
	<script type="text/javascript" src="../scripts/modules.js"></script>
	<script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	<script type="text/javascript" src="../scripts/isvaliddate.js"></script> 

	<script language="javascript">
	<!--
		function checkStat() 
		{
			if ( !(form1.statusInProgress.checked) &&  !(form1.statusPending.checked) && !(form1.statusRefund.checked) && !(form1.statusDenied.checked) &&  !(form1.statusCompleted.checked) && !(form1.statusProcessed.checked)) 
			{
				alert("You must select the status.");
				form1.statusPending.focus();
				return false;
			}
		}

		function CheckAllStatus() 
		{
			if (document.form1.CheckAllStat.checked) 
			{
				document.form1.statusPending.checked = true;
				document.form1.statusCompleted.checked = true;
				document.form1.statusDenied.checked = true;
			} 
			else 
			{
				document.form1.statusPending.checked = false;
				document.form1.statusCompleted.checked = false;
				document.form1.statusDenied.checked = false;
			}
		}

		function doCalendar_Old( ToFrom ) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doCalendar( sField ) 
		{
 			var w = (screen.width - 350)/2;
	 		var h = (screen.height - 350)/2;
		 	var sSelectedDate = $("#" + sField).val();

			 eval('window.open("calendarpicker2.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=BuyerForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

//		function  removePass( iPassId )
//		{
//			if (confirm("Delete Pass #" + iPassId + "?"))
//			{
//				location.href='poolpass_remove.asp?passid=' + iPassId;
//			}
//		}

		function SearchCitizens(  )
		{
			iSearchStart = $("#searchstart").val();
			if ( iSearchStart != "" )
			{
				var optiontext;
				var optionchanged;
				//alert(document.BuyerForm.searchname.value);
				var searchtext = document.BuyerForm.searchname.value;
				var searchchanged = searchtext.toLowerCase();

				iSearchStart = parseInt(iSearchStart) + 1;

				for (x=iSearchStart; x < document.BuyerForm.userid.length ; x++)
				{
					optiontext = document.BuyerForm.userid.options[x].text;
					optionchanged = optiontext.toLowerCase();
					if (optionchanged.indexOf(searchchanged) != -1)
					{
						document.BuyerForm.userid.selectedIndex = x;
						document.BuyerForm.results.value = ' Possible Match Found.';
						document.getElementById('searchresults').innerHTML = ' Possible Match Found.';
						document.BuyerForm.searchstart.value = x;
						document.BuyerForm.submit();
						return;
					}
				}
				document.BuyerForm.results.value = ' No Match Found.';
				document.getElementById('searchresults').innerHTML = ' No Match Found.';
				document.BuyerForm.searchstart.value = -1;
			}
			else
			{
				$("#searchstart").focus();
				inlineMsg('searchstart','<strong>Missing Value: </strong>Please enter a name to search for.',5,'searchstart');
			}
		}

		function ClearSearch()
		{
			document.BuyerForm.searchstart.value = -1;
		}

		function UserPick()
		{
			document.BuyerForm.searchname.value = '';
			document.BuyerForm.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.BuyerForm.searchstart.value = -1;
			document.BuyerForm.submit();
		}

		function checkForReceiptNum( p_field ) 
		{
			lcl_value = document.getElementById(p_field).value;

			if(p_field == "receipt_num") 
			{
				if(lcl_value != "") 
				{
					// blank out the other fields that are part of the other way to get results
					document.getElementById("searchname").value = "";
					document.getElementById("results").value = "";
					document.getElementById("searchstart").value = "";
					document.getElementById("userid").value = "0";
					document.getElementById("fromDate").value = "";
					document.getElementById("toDate").value = "";
				}
			}
			else
			{
				if(lcl_value != "") 
				{
					// blank out the receitp number that is not part of this search
					document.getElementById("receipt_num").value = "";
				}
			}
		}

		function validate()
		{
			// validate the fromDate is a valid date
			if ($("#fromDate").val() == "")
			{
				if ($("#receipt_num").val() == '')
				{
					//alert("Please enter a From Date");
					$("#fromDate").focus();
					inlineMsg('fromDate','<strong>Missing Value: </strong>Please enter a From date.',5,'fromDate');
					return;
				}
			}
			else
			{
				if (! isValidDate($("#fromDate").val()))
				{
					//alert("The From date should be a valid date in the format of MM/DD/YYYY. \nPlease enter it again.");
					$("#fromDate").focus();
					inlineMsg('fromDate','<strong>Invalid Value: </strong>The From date should be a valid date in the format of MM/DD/YYYY. \nPlease enter it again.',5,'fromDate');
					return;
				}
			} 

			// validate the toDate is a valid date
			if ($("#toDate").val() == "")
			{
				if ($("#receipt_num").val() == '')
				{
					//alert("Please enter a From Date");
					$("#toDate").focus();
					inlineMsg('toDate','<strong>Missing Value: </strong>Please enter a To date.',5,'toDate');
					return;
				}
			}
			else
			{
				if (! isValidDate($("#toDate").val()))
				{
					//alert("The To date should be a valid date in the format of MM/DD/YYYY. \nPlease enter it again.");
					$("#toDate").focus();
					inlineMsg('toDate','<strong>Invalid Value: </strong>The To date should be a valid date in the format of MM/DD/YYYY. \nPlease enter it again.',5,'toDate');
					return;
				}
			} 

			// validate the receipt_num is a number only
			if ($("#receipt_num").val() != '')
			{
				var rege = /^\d*$/
				var Ok = rege.exec($("#receipt_num").val());
				if ( ! Ok )
				{
					$("#receipt_num").focus();
					inlineMsg('receipt_num','<strong>Invalid Value: </strong>Receipt numbers must be numeric.',5,'receipt_num');
					return false;
				}
			} 
			else
			{
				// check that the 'from date' is before the 'to date'
				var fromDate = new Date($('#fromDate').val());
				var toDate = new Date($('#toDate').val());

				if (fromDate > toDate)
				{
					$("#toDate").focus();
					inlineMsg('toDate','<strong>Invalid Value: </strong>The To Date must be the same or after the From Date.',5,'toDate');
					return;
				}

			}

			document.BuyerForm.submit();
		}

	//-->
	</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
  <div id="centercontent">

<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
  <tr>
		<td><font size="+1"><b>Citizen Recreation Purchases</b></font></td>
  </tr>
  <tr>
      <td>
		 <!--BEGIN: SEARCH OPTIONS-->
<%			' Show pick of registered users 
			ShowRegisteredUsers iUserId, sSearchName, sResults, sSearchStart, FromDate, ToDate
%>
      </td>
  </tr>
 	<tr>
      <td colspan="3" valign="top">
	  <!--BEGIN: Purchases List -->
<% 
		If OrgHasFeature( "memberships" ) Then
			' Get PoolMemberships
			GetPoolpassPayments iUserId, FromDate, ToDate, iReceiptNum
		End If 

		If OrgHasFeature( "activities" ) Then
			' Get Classes and Event purchases
			GetClassPayments iUserId, FromDate, ToDate, iReceiptNum
		End If 
		
		If OrgHasFeature( "gifts" ) Then 
			' Get Commemorative Gift purchases
			GetGiftPurchases iUserId, FromDate, ToDate, iReceiptNum
		End If 

		If OrgHasFeature( "facilities" ) Then
			' Get facility rentals 
			GetFacilityRentals iUserId, FromDate, ToDate, iReceiptNum
		End If 

		If OrgHasFeature( "rentals" ) Then
			GetRentalPayments iUserId, FromDate, ToDate, iReceiptNum
		End If 

'		If OrgHasFeature( "payments" ) Then 
'			' Get Payments
'			GetPayments iUserId, FromDate, ToDate
'		End If 

		If Not oList.EOF Then 
			' Sort them by date
			oList.Sort = "purchaseyear desc,purchasemonth desc,purchaseday desc,whatpurchased desc,paymentid desc"
			oList.MoveFirst
			' Print out what has been collected
			'response.write "<div class=""purchasereportshadow"">" 
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""purchaseslist"">" 
			response.write "<tr class=""tablelist"">" 
			response.write "<th>&nbsp;</th>" 
			response.write "<th align=""left"">Date</th>" 
			response.write "<th align=""left"">Item</th>" 
			response.write "<th align=""left"">Payee</th>" 
			response.write "<th align=""right"">Amount</th>" 

			If OrgHasFeature( "activities" ) Then 
				response.write "<th align=""right"">Receipt #</th>" 
			End If 

			response.write "<th>&nbsp;</th>" 
			response.write "</tr>" 
			bgcolor = "#eeeeee"
			' LOOP AND DISPLAY THE RECORDS
			Do While Not oList.EOF 
				iPurchaseCount = iPurchaseCount + 1
				
				response.write vbcrlf & "<tr id=""" & iPurchaseCount & """"
				If iPurchaseCount Mod 2 = 0 Then
					response.write " class=""altrow"" "
				End If
				response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				'response.write "<tr bgcolor=""" &  bgcolor  & """ class=""tablelist"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='pointer';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">"
				response.write "<td class=""purchasecell"">&nbsp;</td>" 
				response.write "<td onClick=""location.href='" & oList("detailurl") & "';"">" & oList("purchasedate")  & "</td>" 
				response.write "<td onClick=""location.href='" & oList("detailurl") & "';"">" & oList("whatpurchased") & "</td>" 
				response.write "<td onClick=""location.href='" & oList("detailurl") & "';"">" & oList("purchaser")     & "</td>" 
				response.write "<td align=""right"" onClick=""location.href='" & oList("detailurl") & "';"">" & FormatCurrency(oList("amount"),2) & "</td>" 
				'cTotalAmount = cTotalAmount + CDbl(oRs("paymentamount"))

				If OrgHasFeature( "activities" ) Then 
					If oList("paymentid") = 0 Then 
						lcl_paymentid = ""
					Else 
						lcl_paymentid = oList("paymentid")
					End If 

					response.write "<td align=""right"" onClick=""location.href='" & oList("detailurl") & "';"">" & lcl_paymentid & "</td>" 
				End If 

				response.write "<td>&nbsp;</td>" 
				response.write "</tr>" 
				oList.MoveNext
				response.flush
			Loop 
			response.write vbcrlf & "</table>"
			'response.write "</div>" 
		Else
			response.write "<p>No payment information found.</p>"
		End If 

		oList.Close
		Set oList = Nothing 

		response.write vbcrlf & "</table>"
%>
	  <!-- END: Purchases LIST -->
      </td>
    </tr>
  </table>
</div>
</div>

<!--#include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void GetPoolpassPayments iUserId, sFromDate, sToDate, sReceiptNum
'--------------------------------------------------------------------------------------------------
Sub GetPoolpassPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate, ByVal sReceiptNum )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "

	If CLng(iUserId) > CLng(0) Then 
		varWhereClause = varWhereClause & " AND U.userid = " & iUserId & " "
	End If 

	If sReceiptNum <> "" And Not IsNull(sReceiptNum) Then 
		varWhereClause = varWhereClause & " AND 1 = 0 "
	End If 

	sSql = "SELECT P.poolpassid, U.userfname, U.userlname, P.paymentamount, P.paymenttype, P.paymentdate, m.membershipdesc, "
	sSql = sSql & " P.paymentlocation, R.description, T.description as residenttype, P.paymentresult, "
	sSql = sSql & " ISNULL(P.processingfee,0.00) AS processingfee "
	sSql = sSql & " FROM egov_poolpasspurchases P, egov_users U, egov_poolpassrates R, egov_poolpassresidenttypes T, egov_memberships m "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.paymentresult <> 'Pending' "
	sSql = sSql & " and p.membershipid = m.membershipid "
	sSql = sSql & " AND P.paymentresult <> 'Declined' AND U.userid = P.userid "
	sSql = sSql & " AND P.rateid = R.rateid AND R.residenttype = T.resident_type "
	sSql = sSql & " AND T.orgid = P.orgid " & varWhereClause
	sSql = sSql & " ORDER BY P.poolpassid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	 
	Do While Not oRs.EOF 
		oList.AddNew
		oList("purchaseday")   = Day(oRs("paymentdate"))
		oList("purchasemonth") = Month(oRs("paymentdate"))
		oList("purchaseyear")  = Year(oRs("paymentdate"))
		oList("purchasedate")  = DateValue(oRs("paymentdate"))
		'oList("whatpurchased") = "Membership &ndash; " & oRs("residenttype") & " &ndash; " & oRs("description")
		oList("whatpurchased") = oRs("membershipdesc") & " &ndash; " & oRs("residenttype") & " &ndash; " & oRs("description")
		oList("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oList("amount")        = CDbl(oRs("paymentamount")) + CDbl(oRs("processingfee"))
		oList("paymentid")     = 0
		oList("detailurl")     = "poolpass_details.asp?iPoolPassId=" & oRs("poolpassid")
		oList.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 
 
End Sub 


'--------------------------------------------------------------------------------------------------
' void GetClassPayments iUserId, sFromDate, sToDate, sReceiptNum
'--------------------------------------------------------------------------------------------------
Sub GetClassPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate, ByVal sReceiptNum )
	Dim varWhereClause, sSql, oRs

	If sReceiptNum = "" Then 
		varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' "
		varWhereClause = varWhereClause & " AND P.paymentdate < '" & sToDate & "') "

		if isnumeric(iUserId) then
			If iUserId > 0 Then 
				varWhereClause = varWhereClause & " AND P.userid = " & iUserId & " "
			End If 
		End If 
	Else 
		If sReceiptNum <> "" And Not IsNull(sReceiptNum) Then 
			varWhereClause = varWhereClause & " AND P.paymentid = " &  sReceiptNum 
		End If 
	End If 

	sSql = "SELECT P.paymentid, U.userfname, U.userlname, P.paymenttotal, P.paymentdate, J.journalentrytype "
	sSql = sSql & " FROM egov_class_payment P, egov_users U, egov_journal_entry_types J "
	sSql = sSql & " WHERE U.userid = P.userid AND P.isforrentals = 0 AND P.orgid = " & session("orgid")
	sSql = sSql & " AND P.journalentrytypeid < 3 AND J.journalentrytypeid = P.journalentrytypeid "
	sSql = sSql & varWhereClause
	sSql = sSql & " ORDER BY P.paymentid "
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	 
	Do While Not oRs.EOF 
		oList.AddNew
		oList("purchaseday")   = Day(oRs("paymentdate"))
		oList("purchasemonth") = Month(oRs("paymentdate"))
		oList("purchaseyear")  = Year(oRs("paymentdate"))
		oList("purchasedate")  = DateValue(oRs("paymentdate"))
		oList("whatpurchased") = "Classes & Events &ndash; " & UCase(Left(oRs("journalentrytype"),1)) & LCase(Mid(oRs("journalentrytype"),2))
		oList("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oList("amount")        = oRs("paymenttotal") + GetProcessingFee( oRs("paymentid") )
		oList("paymentid")     = oRs("paymentid")
		oList("detailurl")     = "../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") 
		oList.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void GetRentalPayments iUserId, sFromDate, sToDate, sReceiptNum
'------------------------------------------------------------------------------
Sub GetRentalPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate, ByVal sReceiptNum )
	Dim varWhereClause, sSql, oRs

	If sReceiptNum = "" Then 
		varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' "
		varWhereClause = varWhereClause & " AND P.paymentdate < '" & sToDate & "') "

		If iUserId > 0 Then 
			varWhereClause = varWhereClause & " AND P.userid = " & iUserId & " "
		End If 
	Else 
		If sReceiptNum <> "" And Not IsNull(sReceiptNum) Then 
			varWhereClause = varWhereClause & " AND P.paymentid = '" & sReceiptNum & "' "
		End If 
	End If 

	sSql = "SELECT P.paymentid, U.userfname, U.userlname, P.paymenttotal, P.paymentdate, P.reservationid, J.journalentrytype "
	sSql = sSql & " FROM egov_class_payment P, egov_users U, egov_journal_entry_types J "
	sSql = sSql & " WHERE  U.userid = P.userid AND P.isforrentals = 1 AND P.reservationid IS NOT NULL "
	sSql = sSql & " AND J.journalentrytypeid = P.journalentrytypeid " & varWhereClause
	sSql = sSql & " AND P.orgid = " & session("orgid")
	sSql = sSql & " ORDER BY P.paymentid "
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		oList.AddNew
		oList("purchaseday")   = Day(oRs("paymentdate"))
		oList("purchasemonth") = Month(oRs("paymentdate"))
		oList("purchaseyear")  = Year(oRs("paymentdate"))
		oList("purchasedate")  = DateValue(oRs("paymentdate"))
		sLabel = "Reservation #" & oRs("reservationid") & " &ndash; "  ' & GetFirsRentalName( oRs("reservationid") )
		If LCase(oRs("journalentrytype")) = "refund" Then
			sLabel = sLabel & "Refund"
		Else
			sLabel = sLabel & "Payment"
		End If 
		oList("whatpurchased") = sLabel
		oList("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oList("amount")        = oRs("paymenttotal") + GetProcessingFee( oRs("paymentid") )
		oList("paymentid")     = oRs("paymentid")
		oList("detailurl")     = "../rentals/viewpaymentreceipt.asp?paymentId=" & oRs("paymentid") 
		oList.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetGiftPurchases iUserId, sFromDate, sToDate, sReceiptNum
'--------------------------------------------------------------------------------------------------
Sub GetGiftPurchases( ByVal iUserId, ByVal sFromDate, ByVal sToDate, ByVal sReceiptNum )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "

	If CLng(iUserId) > CLng(0) Then 
		varWhereClause = varWhereClause & " AND U.userid = " & iUserId & " "
	End If 

	If sReceiptNum <> "" And Not IsNull(sReceiptNum) Then 
		varWhereClause = varWhereClause & " AND 1 = 0 "
	End If 

	sSql = "SELECT P.giftpaymentid, U.userfname, U.userlname, P.giftamount, P.paymentdate, G.giftname, "
	sSql = sSql & " ISNULL(P.processingfee,0.00) AS processingfee "
	sSql = sSql & " FROM egov_gift_payment P, egov_users U, egov_gift G "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.result = 'APPROVED' "
	sSql = sSql & " AND U.useremail = P.email AND G.giftid = P.giftid "
	sSql = sSql & varWhereClause
	sSql = sSql & " ORDER BY P.giftpaymentid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	 
	Do While Not oRs.EOF 
		oList.AddNew
		oList("purchaseday")   = Day(oRs("paymentdate"))
		oList("purchasemonth") = Month(oRs("paymentdate"))
		oList("purchaseyear")  = Year(oRs("paymentdate"))
		oList("purchasedate")  = DateValue(oRs("paymentdate"))
		oList("whatpurchased") = oRs("giftname")
		oList("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oList("amount")        = CDbl(oRs("giftamount")) + CDbl(oRs("processingfee"))
		oList("paymentid")     = 0
		oList("detailurl")     = "gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid")
		oList.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetFacilityRentals iUserId, sFromDate, sToDate, sReceiptNum
'--------------------------------------------------------------------------------------------------
Sub GetFacilityRentals( ByVal iUserId, ByVal sFromDate, ByVal sToDate, ByVal sReceiptNum )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.datecreated >= '" & sFromDate & "' AND P.datecreated < '" & sToDate & "') "

	If iUserId > 0 Then 
		varWhereClause = varWhereClause & " AND U.userid = " & iUserId & " "
	End If 

	If sReceiptNum <> "" And Not IsNull(sReceiptNum) Then 
		varWhereClause = varWhereClause & " AND 1 = 0 "
	End If 

	sSql = "SELECT P.facilityscheduleid, U.userfname, U.userlname, P.amount, P.datecreated, "
	sSql = sSql & " F.facilityname, ISNULL(P.processingfee,0.00) AS processingfee "
	sSql = sSql & " FROM egov_facilityschedule P, egov_users U, egov_facility F"
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.status = 'RESERVED' "
	sSql = sSql & " AND P.lesseeid = U.userid AND P.facilityid = F.facilityid "
	sSql = sSql & varWhereClause
	sSql = sSql & " ORDER BY P.facilityscheduleid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	 
	Do While Not oRs.EOF 
		oList.AddNew
		oList("purchaseday")   = Day(oRs("datecreated"))
		oList("purchasemonth") = Month(oRs("datecreated"))
		oList("purchaseyear")  = Year(oRs("datecreated"))
		oList("purchasedate")  = DateValue(oRs("datecreated"))
		oList("whatpurchased") = "Facility Rental &ndash; " & oRs("facilityname")
		oList("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oList("amount")        = CDbl(oRs("amount")) + CDbl(oRs("processingfee"))
		oList("paymentid")     = 0
		oList("detailurl")     = "facility_details.asp?iFacilityScheduleId=" & oRs("facilityscheduleid")
		oList.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void GetPayments iUserId, sFromDate, sToDate, sReceiptNum
'--------------------------------------------------------------------------------------------------
Sub GetPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate, ByVal sReceiptNum )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "

	If iUserId > 0 Then 
		varWhereClause = varWhereClause & " AND P.userid = " & iUserId & " "
	End If 

	If sReceiptNum <> "" And Not IsNull(sReceiptNum) Then 
		varWhereClause = varWhereClause & " AND 1 = 0 "
	End If 

	sSql = "SELECT P.paymentid, P.paymentdate, P.paymentamount, P.paymentservicename, P.userfname, isnull(P.userlname,'') as userlname "
	sSql = sSql & " FROM egov_payment_list P "
	sSql = sSql & " WHERE (P.paymentstatus = 'COMPLETED' OR P.paymentstatus = 'PROCESSED') "
	sSql = sSql & varWhereClause

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	 
	Do While Not oRs.EOF 
		oList.AddNew
		oList("purchaseday")   = Day(oRs("paymentdate"))
		oList("purchasemonth") = Month(oRs("paymentdate"))
		oList("purchaseyear")  = Year(oRs("paymentdate"))
		oList("purchasedate")  = DateValue(oRs("paymentdate"))
		oList("whatpurchased") = oRs("paymentservicename")
		oList("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oList("amount")        = oRs("paymentamount")
		oList("paymentid")     = oRs("paymentid")
		oList("detailurl")     = "payment_details.asp?iPaymentid=" & oRs("paymentid")
		oList.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRegisteredUsers iUserId, sSearchName, sResults, sSearchStart, sFromDate, sToDate 
'--------------------------------------------------------------------------------------------------
Sub ShowRegisteredUsers( ByVal iUserId, ByVal sSearchName, ByVal sResults, ByVal sSearchStart, ByVal sFromDate, ByVal sToDate )

	If OrgHasFeature( "activities" ) Then 
		lcl_onblur_searchname = "onblur=""checkForReceiptNum('searchname')"""
		lcl_onblur_userid     = "onblur=""checkForReceiptNum('userid')"""
		lcl_onblur_fromdate   = "onblur=""checkForReceiptNum('fromDate')"""
		lcl_onblur_todate     = "onblur=""checkForReceiptNum('toDate')"""
	Else 
		lcl_onblur_searchname = ""
		lcl_onblur_userid     = ""
		lcl_onblur_fromdate   = ""
		lcl_onblur_todate     = ""
	End If 

	response.write "<fieldset id=""purchasereport"">" 
	response.write "<legend><strong> Purchaser Information </strong></legend>" 
	response.write "<form name=""BuyerForm"" method=""post"" action=""purchases_list.asp"">" 
	response.write "<br />"
	response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""searchtable"">" 
	response.write "<tr>" 
	response.write "<td nowrap=""nowrap"">Name Search:</td> " 
	response.write "<td colspan=""3"">" 
	response.write "<input type=""text"" name=""searchname"" id=""searchname"" value=""" & sSearchName & """ size=""25"" maxlength=""50"" onchange=""javascript:ClearSearch();""" & lcl_onblur_searchname & " /> &nbsp; " 
	response.write "<input type=""button"" class=""button"" value=""Name Search"" onclick=""javascript:SearchCitizens();"" />" 
	response.write "<input type=""" & lcl_hidden & """ name=""results"" id=""results"" value="""" />" 
	response.write "<input type=""" & lcl_hidden & """ name=""searchstart"" id=""searchstart"" value=""" & sSearchStart & """ />" 
	response.write "<span id=""searchresults"">" & sResults & "</span>" 
	response.write "</td>" 
	response.write "</tr>" 
	response.write "<tr>" 
	response.write "<td>&nbsp;</td>" 
	response.write "<td colspan=""3""><div id=""searchtip"">(last name, first name)</div></td>" 
	response.write "</tr>" 
	response.write "<tr>" 
	response.write "<td nowrap=""nowrap"">Select Name:</td>" 
	response.write "<td colspan=""3"">" 
	response.write "<select name=""userid"" id=""userid""" & lcl_onblur_userid & ">" 
	ShowUserDropDown iUserId 
	response.write "</select>" 
	response.write "</td>" 
	response.write "</tr>" 
	response.write "<tr>" 
	response.write "<td><b>From:</b></td>" 
	response.write "<td>" 
	response.write "<input type=""text"" name=""fromDate"" id=""fromDate"" value=""" & sFromDate & """" & lcl_onblur_fromdate & " />" 
	response.write "<a href=""javascript:void doCalendar('fromDate');""><img src=""../images/calendar.gif"" border=""0""></a>" 
	response.write "</td>" 
	response.write "<td><b>To:</b></td>" 
	response.write "<td>" 
	response.write "<input type=""text"" name=""toDate"" id=""toDate"" value=""" & dateAdd("d",-1,sToDate) & """" & lcl_onblur_todate & " />" 
	response.write "<a href=""javascript:void doCalendar('toDate');""><img src=""../images/calendar.gif"" border=""0""></a>" 
	response.write "</td>" 
	response.write "</tr>" 

	If OrgHasFeature( "activities" ) Then 
		response.write "<tr>" 
		response.write "<td colspan=""4"">------------------------------------------------- <font color=""#ff0000"">OR</font> -------------------------------------------------</td>" 
		response.write "</tr>" 
		response.write "<tr>" 
		response.write "<td>Receipt#:</td>" 
		response.write "<td colspan=""3"">" 
		response.write "<input type=""text"" name=""receipt_num"" id=""receipt_num"" value=""" & iReceiptNum & """ size=""10"" maxlength=""10"" onblur=""checkForReceiptNum('receipt_num')"">" 
		response.write "</td>" 
		response.write "</tr>" 
	End If 

	response.write "</table>" 
	response.write "<input class=""button"" type=""button"" value=""Show Payments"" onclick=""validate()"" /><p>" 

	response.write "</form>" 
	response.write "</fieldset>" 
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowUserDropDown( iUserId )
'--------------------------------------------------------------------------------------------------
Sub ShowUserDropDown( ByVal iUserId )
	Dim sSql, oCmd, oRs

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oRs = .Execute
	End With

	If CLng(iUserId) = CLng(0) Then 
		lcl_selected_none = " selected=""selected"""
	Else 
		lcl_selected_none = ""
	End If 

	response.write "<option value=""0""" & lcl_selected_none & ">No Choice</option>" 

	Do While Not oRs.EOF
		If CLng(iUserId) = CLng(oRs("userid")) Then 
			lcl_selected = " selected=""selected"""
		Else 
			lcl_selected = ""
		End If 

		response.write "<option value=""" & oRs("userid") & """" & lcl_selected & ">" & oRs("userlname") & ", " & oRs("userfname") & " &ndash; " & oRs("useraddress") & "</option>" 

		oRs.MoveNext
		response.flush
	Loop 

	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing

End Sub  


'--------------------------------------------------------------------------------------------------
' integer GetFirstUserId()
'--------------------------------------------------------------------------------------------------
Function GetFirstUserId()
	Dim sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetFirstEgovUserByOrgid"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
		.Parameters.Append oCmd.CreateParameter("@iUserId", 3, 2, 4)
	    .Execute
	End With

	GetFirstUserId = oCmd.Parameters("@iUserId").Value

	Set oCmd = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetFirsRentalName( iReservationId )
'------------------------------------------------------------------------------
Function GetFirsRentalName( ByVal iReservationId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalname FROM egov_rentalreservationdates D, egov_rentals R "
	sSql = sSql & "WHERE D.rentalid = R.rentalid AND D.reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oRs.EOF Then
		GetFirsRentalName = oRs("rentalname")
	Else
		GetFirsRentalName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 





%>
