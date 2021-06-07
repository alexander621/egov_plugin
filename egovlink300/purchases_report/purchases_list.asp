<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: purchases_list.asp
' AUTHOR: Steve Loar
' CREATED: 07/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a list of purchases by citizen
'
' MODIFICATION HISTORY
' 1.0	07/31/2006	Steve Loar - INITIAL VERSION
' 2.0	08/04/2006	Steve Loar - Public version made from admin code
' 2.1	12/08/2009	David Boyer - Now check for the membership name to be used instead of hard-coded "pool"
' 2.2	02/10/2010	Steve Loar - Added rental reservations to the list.
' 2.3	07/20/2010	Steve Loar - Modified to include the PNP fee in the total amount
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim oDynamicRs, orderBy, subTotals, showDetail, fromDate, toDate, today
 dim sUserlname, iPurchaseCount, bgcolor, iUserId, sResults, sSearchStart
 dim sSearchName, oOrganization, nAccountCredit

 set oOrganization = New classOrganization

 orderBy      = Request("orderBy")
 subTotals    = Request("subTotals")
 showDetail   = Request("showDetail")
 fromDate     = Request("fromDate")
 toDate       = Request("toDate")
 sUserlname   = Request("userlname")
 today        = Date()
 sSearchName  = ""
	sResults     = ""
	sSearchStart = -1

'Session("RedirectPage") = "purchases_report/purchases_list.asp"
'Session("RedirectLang") = "Return to View Purchases"
'session("ManageURL") = ""

'If they do not have a userid set, take them to the login page automatically
 if request.cookies("userid") = "" Or request.cookies("userid") = "-1" then
	   session("LoginDisplayMsg") = "Please sign in first and then we'll send you right along."
   	response.redirect "../user_login.asp"
 end if

 iUserid = request.cookies("userid")
 session("eGovUserId") = iUserId

'See if a search term was passed
 if request("searchname") <> "" then
	   sSearchName = request("searchname")
 end if

 if request("results") <> "" then
   	sResults = request("results")
 end if

 if request("searchstart") <> "" then
	   sSearchStart = request("searchstart")
 end if

 if orderBy = "" OR IsNull(orderBy) then
	   orderBy = "date"
 end if

 if toDate = "" OR IsNull(toDate) then
	   toDate = dateAdd("d", 0, today)
 end if

'If fromDate = "" or IsNull(fromDate) Then fromDate = dateAdd("ww",-1,today) End If
 if fromDate = "" OR IsNull(fromDate) then
	   fromDate = DateSerial(Year(Now()), 1, 1)
 end if

 toDate = dateAdd("d", 1, toDate)

 set oDynamicRs = Server.CreateObject("ADODB.Recordset")

 oDynamicRs.Fields.Append "purchaseday", adInteger, , adFldUpdatable
 oDynamicRs.Fields.Append "purchasemonth", adInteger, , adFldUpdatable
 oDynamicRs.Fields.Append "purchaseyear", adInteger, , adFldUpdatable
 oDynamicRs.Fields.Append "purchaseid", adInteger, , adFldUpdatable
 oDynamicRs.Fields.Append "purchasedate", adVariant, 10, adFldUpdatable
 'oDynamicRs.Fields.Append "purchasedate", adVarChar, 10, adFldUpdatable
 oDynamicRs.Fields.Append "whatpurchased", adVarChar, 255, adFldUpdatable
 oDynamicRs.Fields.Append "purchaser", adVarChar, 255, adFldUpdatable
 oDynamicRs.Fields.Append "amount", adCurrency, , adFldUpdatable
 oDynamicRs.Fields.Append "detailurl", adVarChar, 255, adFldUpdatable
 'oDynamicRs.CursorType = adOpenDynamic
 oDynamicRs.CursorLocation = 3

 oDynamicRs.Open 

'Build the TITLE
 lcl_title = "E-Gov Services " & oOrganization.GetOrgName()

 if iorgid = 7 then
	   lcl_title = oOrganization.GetOrgName()
 end if

 nAccountCredit = GetCitizenAccountAmount( iUserid )  ' in common.asp
%>
<html>
<head>
	<title><%=lcl_title%></title>
	<meta http-equiv="Content-type" content="text/html;charset=UTF-8">
 <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no" />

	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />

	<style>
		table#purchaseslist
		{
			top: 0 !important;
			left: 0 !important;
			margin-top: 2em !important;
			margin-bottom: 3em !important;
			/* give a drop shadow */
			-moz-box-shadow: 3px 3px 7px #777;
			-webkit-box-shadow: 3px 3px 7px #777;
			box-shadow: 3px 3px 7px #777;
		}

		div#accounttotal
		{
			margin-top: 1em;
		}

  .fieldset
  {
     border: 1pt solid #c0c0c0;
     border-radius: 6px;
     background-color: #eeeeee;
  }

  .fieldset legend
  {
    border: 1pt solid #c0c0c0;
    border-radius: 5px;
    background-color: #ffffff;
    padding: 2px 5px;
    font-family: Verdana, sans-serif;
    font-size: 1.25em;  /* 20 / 16 */
    color: #800000;
  }

  #content table.purchasereport
  {
     width: 100%;
  }

  #buttonShowPayments,
  .searchDateImg
  {
     cursor: pointer;
  }

  /* ----------------------------------------------------------------------- */
  @media screen and (max-width: 800px) 
  {
     #centercontent
     {
        width: 100%;
        margin-left: 0px;
     }

     .fieldset
     {
        margin: 0px 5px;
     }

     #accounttotal
     {
        text-align: center;
     }

     #content table.purchasereport
     {
        border: none;
        box-shadow: none;
        border-bottom: 1px solid #666666;
     }
  }
	</style>

 <script src="../scripts/formvalidation_msgdisplay.js"></script>
	<script src="../scripts/jquery-1.9.1.min.js"></script>

	<script>
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

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			var sSelectedDate = '';

			if ($("#" + sField).val() != '')
			{
				// The value in the field
				sSelectedDate = $("#" + sField).val();
			}
			else
			{
				if (sField == 'toDate')
				{
					// Show the to date from where the from date is
					sSelectedDate = $("#fromDate").val();
				}

				if (sSelectedDate == '')
				{
					// This is today's date
					sSelectedDate = new Date();
					var month = sSelectedDate.getMonth() + 1;
					var day = sSelectedDate.getDate();
					var year = sSelectedDate.getFullYear();
					sSelectedDate = month + "/" + day + "/" + year;
				}
			}

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&updatefield=' + sField + '&updateform=BuyerForm", "_calendar", "width=350,height=250,toolbar=0,status=0,scrollbars=0,menubar=0,titlebar=0,location=0,dependent=yes,personalbar=no,left=' + w + ',top=' + h + '")');
		}

		function doCalendar_old(ToFrom) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function SearchCitizens( iSearchStart )
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
					document.BuyerForm.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.BuyerForm.searchstart.value = x;
					document.BuyerForm.submit();
					return;
				}
			}
			document.BuyerForm.results.value = 'No Match Found.';
			document.getElementById('searchresults').innerHTML = 'No Match Found.';
			document.BuyerForm.searchstart.value = -1;
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

  $(document).ready(function() {
     $('#fromDate').change(function() {
    				clearMsg("fromDateCalPop");
     });

     $('#toDate').change(function() {
    				clearMsg("toDateCalPop");
     });

     $('#fromDateCalPop').click(function() {
    				clearMsg("fromDateCalPop");

     			var daterege   = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
     			var dateFromOk = daterege.test($('#fromDate').val());

     			if (! dateFromOk ) {
        				$('#fromDate').focus();
        				inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
        				lcl_return_false = lcl_return_false + 1;
     			} else {
             doCalendar('fromDate');
     			}
     });

     $('#toDateCalPop').click(function() {
    				clearMsg("toDateCalPop");

     			var daterege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
	     		var dateToOk = daterege.test($('#toDate').val());

     			if (! dateToOk ) {
        				$('#toDate').focus();
        				inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
        				lcl_return_false = lcl_return_false + 1;
     			} else {
            doCalendar('toDate');
     			}
     });

     $('#buttonShowPayments').click(function() {
        var lcl_return_false = 0;
     			var daterege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;

     			var dateFromOk = daterege.test($('#fromDate').val());
	     		var dateToOk   = daterege.test($('#toDate').val());

     			if (! dateToOk ) {
        				$('#toDate').focus();
        				inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
        				lcl_return_false = lcl_return_false + 1;
     			} else {
        				clearMsg("toDateCalPop");
     			}

     			if (! dateFromOk ) {
        				$('#fromDate').focus();
        				inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
	        			lcl_return_false = lcl_return_false + 1;
     			} else {
        				clearMsg("fromDateCalPop");
     			}

        if(lcl_return_false > 0) {
       				return false;
     			} else {
       				$('#BuyerForm').submit();
       				return true;
     			}
     });
  });

	//-->
	</script>
</head>
<!--#Include file="../include_top.asp"-->
<%
  response.write "<p>" & vbcrlf
  response.write "<font class=""pagetitle"">" & oOrganization.GetOrgName() & " E-Gov Purchases</font><br />" & vbcrlf

  RegisteredUserDisplay( "../" )

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf

	'Show calendar picks
 	ShowDateChoices FromDate, ToDate

	'Get PoolMemberships
	 If oOrganization.OrgHasFeature( "memberships" ) Then 
		   GetPoolpassPayments iUserId, FromDate, ToDate
 	End If 

	'Get Classes and Event purchases
 	If oOrganization.OrgHasFeature( "activities" ) Then 
	   	GetClassPayments iUserId, FromDate, ToDate
 	end if

	'Get Commemorative Gift purchases
	 If oOrganization.OrgHasFeature( "gifts" ) Then 
		   GetGiftPurchases iUserId, FromDate, ToDate
 	End If 

	'Get facility rentals 
  If oOrganization.OrgHasFeature( "facilities" ) Then 
   		GetFacilityRentals iUserId, FromDate, ToDate
 	End If 

	'Get Rental Reservations
	 If oOrganization.OrgHasFeature( "rentals" ) Then
		   GetRentalPayments iUserId, FromDate, ToDate
 	End If 

	 If OrgHasFeature( iOrgId, "public account payments") Then
		   response.write "<div id=""accounttotal"">Balance on Your Account: " & FormatNumber(GetCitizenAccountAmount( request.cookies("userid") ),2) & "</div>"
 	End If 

  if not oDynamicRs.EOF then
		  'Sort them by date
   		oDynamicRs.Sort = "purchaseyear desc, purchasemonth desc, purchaseday desc, purchaseid desc, whatpurchased"
   		oDynamicRs.MoveFirst

  		'Print out what has been collected
		  'response.write vbcrlf & "<div class=""purchasereportshadow"">"
		   response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""purchasereport liquidtable"" id=""purchaseslist"">" & vbcrlf
   		response.write "    <thead><tr class=""tablelist"">" & vbcrlf
   		response.write "        <th>&nbsp;</th>" & vbcrlf
     response.write "        <th align=""left"">Date</th>" & vbcrlf
     response.write "        <th>Receipt</th>" & vbcrlf
     response.write "        <th align=""left"">Item</th>" & vbcrlf
 				response.write "        <th align=""left"">Purchaser</th>" & vbcrlf
     response.write "        <th align=""right"">Amount</th>" & vbcrlf
     response.write "        <th>&nbsp;</th>" & vbcrlf
 				response.write "    </tr></thead>"

   		bgcolor        = "#eeeeee"
     lcl_td_onclick = ""

  		'Loop and display the records
		   do while not oDynamicRs.eof
     			iPurchaseCount = iPurchaseCount + 1
     			bgcolor        = changeBGColor(bgcolor,"#eeeeee","#ffffff")
        lcl_td_onclick = " onclick=""location.href='" & oDynamicRs("detailurl") & "';"""

     			response.write "    <tr bgcolor=""" &  bgcolor  & """ class=""tablelist"" onMouseOver=""this.style.backgroundColor='#BDBABD';this.style.cursor='pointer';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">" & vbcrlf
     			response.write "        <td>&nbsp;</td>" & vbcrlf
     			response.write "        <td class=""repeatheaders"">Date</td>" & vbcrlf
     			response.write "        <td" & lcl_td_onclick & ">" & oDynamicRs("purchasedate")  & "</td>" & vbcrlf
     			response.write "        <td class=""repeatheaders"">Receipt</td>" & vbcrlf
			     response.write "        <td" & lcl_td_onclick & " align=""center"">" & oDynamicRs("purchaseid") & "</td>" & vbcrlf
     			response.write "        <td class=""repeatheaders"">Item</td>" & vbcrlf
			     response.write "        <td" & lcl_td_onclick & ">" & oDynamicRs("whatpurchased") & "</td>" & vbcrlf
     			response.write "        <td class=""repeatheaders"">Purchaser</td>" & vbcrlf
			     response.write "        <td" & lcl_td_onclick & ">" & oDynamicRs("purchaser")     & "</td>" & vbcrlf
     			response.write "        <td class=""repeatheaders"">Amount</td>" & vbcrlf
			     response.write "        <td" & lcl_td_onclick & " align=""right"">" & FormatCurrency(oDynamicRs("amount"),2) & "</td>" & vbcrlf
      		response.write "        <td>&nbsp;</td>" & vbcrlf
     			response.write "    </tr>" & vbcrlf

     			'cTotalAmount = cTotalAmount + CDbl(oDynamicRs("paymentamount"))
     			oDynamicRs.movenext
     loop

   		response.write "</table>" & vbcrlf
  else
   		response.write "<p>No payment information found.</p>" & vbcrlf
  end if

 	oDynamicRs.Close
 	set oDynamicRs    = nothing 
 	set oOrganization = nothing 

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
sub GetPoolpassPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate )
	Dim varWhereClause, sSql, oRs, sMembershipDesc

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "
	varWhereClause = varWhereClause & " AND U.userid = " & iUserId & " "

	sSql = "SELECT P.poolpassid, U.userfname, U.userlname, P.paymentamount, P.paymenttype, P.paymentdate, P.paymentlocation, "
	sSql = sSql & " R.description, T.description as residenttype, P.paymentresult, "
	sSql = sSql & " (select m.membershipdesc from egov_memberships m where m.membershipid = P.membershipid) AS membershipdesc "
	sSql = sSql & " FROM egov_poolpasspurchases P, egov_users U, egov_poolpassrates R, egov_poolpassresidenttypes T "
	sSql = sSql & " WHERE P.orgid = " & iOrgId
	sSql = sSql & " AND P.paymentresult <> 'Pending' "
	sSql = sSql & " AND P.paymentresult <> 'Declined' "
	sSql = sSql & " AND U.userid = P.userid "
	sSql = sSql & " AND P.rateid = R.rateid "
	sSql = sSql & " AND R.residenttype = T.resident_type "
	sSql = sSql & " AND T.orgid = P.orgid " & varWhereClause
	sSql = sSql & " ORDER BY P.poolpassid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	Do While Not oRs.EOF
		oDynamicRs.AddNew

		If oRs("membershipdesc") <> "" Then
			sMembershipDesc = oRs("membershipdesc")
		Else 
			sMembershipDesc = "Pool"
		End If 

		oDynamicRs("purchaseday")   = Day(oRs("paymentdate"))
		oDynamicRs("purchasemonth") = Month(oRs("paymentdate"))
		oDynamicRs("purchaseyear")  = Year(oRs("paymentdate"))
		oDynamicRs("purchasedate")  = DateValue(oRs("paymentdate"))
		oDynamicRs("purchaseid")    = oRs("poolpassid")
		oDynamicRs("whatpurchased") = sMembershipDesc & " Membership &ndash; " & oRs("residenttype") & " &ndash; " & oRs("description")
		oDynamicRs("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oDynamicRs("amount")        = oRs("paymentamount")
		oDynamicRs("detailurl")     = "poolpass_details.asp?iPoolPassId=" & oRs("poolpassid")
		oDynamicRs.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 
 
end sub

'------------------------------------------------------------------------------
sub GetClassPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "
	varWhereClause = varWhereClause & " AND P.userid = " & iUserId & " "

	sSql = "SELECT P.paymentid, U.userfname, U.userlname, P.paymenttotal, P.paymentdate, J.journalentrytype "
	sSql = sSql & " FROM egov_class_payment P, egov_users U, egov_journal_entry_types J "
	sSql = sSql & " WHERE  U.userid = P.userid AND P.journalentrytypeid < 3  AND P.isforrentals = 0 "
	sSql = sSql & " AND J.journalentrytypeid = P.journalentrytypeid " & varWhereClause
	sSql = sSql & " AND P.orgid = " & iOrgId
	sSql = sSql & " ORDER BY P.paymentid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	Do While Not oRs.EOF 
		oDynamicRs.AddNew
		oDynamicRs("purchaseday")   = Day(oRs("paymentdate"))
		oDynamicRs("purchasemonth") = Month(oRs("paymentdate"))
		oDynamicRs("purchaseyear")  = Year(oRs("paymentdate"))
		oDynamicRs("purchasedate")  = DateValue(oRs("paymentdate"))
		oDynamicRs("purchaseid")    = oRs("paymentid")
		oDynamicRs("whatpurchased") = "Recreation Activity &ndash; " & UCase(Left(oRs("journalentrytype"),1)) & LCase(Mid(oRs("journalentrytype"),2))
		oDynamicRs("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oDynamicRs("amount")        = CDbl(oRs("paymenttotal")) + GetProcessingFee( oRs("paymentid") )
		'oDynamicRs("detailurl")     = "../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") 
		oDynamicRs("detailurl")     = "../rd_classes/class_receipt.aspx?paymentid=" & oRs("paymentid") 
		oDynamicRs.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

end sub

'------------------------------------------------------------------------------
sub GetGiftPurchases( ByVal iUserId, ByVal sFromDate, ByVal sToDate )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "
	varWhereClause = varWhereClause & " and U.userid = " & iUserId & " "

	sSql = "SELECT P.giftpaymentid, U.userfname, U.userlname, P.giftamount, P.paymentdate, G.giftname "
	sSql = sSql & " FROM egov_gift_payment P, egov_users U, egov_gift G "
	sSql = sSql & " WHERE P.orgid = " & iOrgId & " AND P.result = 'APPROVED' "
	sSql = sSql & " AND U.useremail = P.email AND G.giftid = P.giftid " & varWhereClause
	sSql = sSql & " ORDER BY P.giftpaymentid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	Do While Not oRs.EOF 
		oDynamicRs.AddNew
		oDynamicRs("purchaseday")   = Day(oRs("paymentdate"))
		oDynamicRs("purchasemonth") = Month(oRs("paymentdate"))
		oDynamicRs("purchaseyear")  = Year(oRs("paymentdate"))
		oDynamicRs("purchasedate")  = DateValue(oRs("paymentdate"))
		oDynamicRs("purchaseid")    = oRs("giftpaymentid")
		oDynamicRs("whatpurchased") = oRs("giftname")
		oDynamicRs("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oDynamicRs("amount")        = oRs("giftamount")
		oDynamicRs("detailurl")     = "gift_details.asp?igiftpaymentid=" & oRs("giftpaymentid")
		oDynamicRs.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

end sub

'------------------------------------------------------------------------------
sub GetFacilityRentals( ByVal iUserId, ByVal sFromDate, ByVal sToDate )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " and (P.datecreated >= '" & sFromDate & "' AND P.datecreated < '" & sToDate & "') "
	varWhereClause = varWhereClause & " and U.userid = " & iUserId & " "

	sSql = "SELECT  P.facilityscheduleid, U.userfname, U.userlname, P.amount, P.datecreated, F.facilityname "
	sSql = sSql & " FROM egov_facilityschedule P, egov_users U, egov_facility F "
	sSql = sSql & " WHERE P.orgid = " & iOrgId & " AND P.status = 'RESERVED' "
	sSql = sSql & " AND P.lesseeid = U.userid AND P.facilityid = F.facilityid " & varWhereClause
	sSql = sSql & " ORDER BY P.facilityscheduleid "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	Do While Not oRs.EOF 
		oDynamicRs.AddNew
		oDynamicRs("purchaseday")   = Day(oRs("datecreated"))
		oDynamicRs("purchasemonth") = Month(oRs("datecreated"))
		oDynamicRs("purchaseyear")  = Year(oRs("datecreated"))
		oDynamicRs("purchasedate")  = DateValue(oRs("datecreated"))
		oDynamicRs("purchaseid")    = oRs("facilityscheduleid")
		oDynamicRs("whatpurchased") = "Facility Rental &ndash; " & oRs("facilityname")
		oDynamicRs("purchaser")     = oRs("userfname") & " " & oRs("userlname")
		oDynamicRs("amount")        = oRs("amount")
		oDynamicRs("detailurl")     = "facility_details.asp?iFacilityScheduleId=" & oRs("facilityscheduleid")
		oDynamicRs.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

end sub

'------------------------------------------------------------------------------
sub ShowDateChoices( ByVal sFromDate, ByVal sToDate )

 	response.write "<form name=""BuyerForm"" id=""BuyerForm"" method=""post"" action=""purchases_list.asp"">" & vbcrlf
 	response.write "<fieldset class=""fieldset"">" & vbcrlf
 	response.write "  <legend>Purchase Date Range</legend>" & vbcrlf
 	response.write "  <table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""searchtable"">" & vbcrlf
 	response.write "      <tr>" & vbcrlf
 	response.write "          <td valign=""top""><strong>From:</strong>" & vbcrlf
 	response.write "              <input type=""text"" id=""fromDate"" name=""fromDate"" value=""" & sFromDate & """ />&nbsp;" & vbcrlf
 	response.write "              <img src=""../images/calendar.gif"" name=""fromDateCalPop"" id=""fromDateCalPop"" border=""0"" class=""searchDateImg"" />" & vbcrlf
 	response.write "          </td>" & vbcrlf
 	response.write "          <td>&nbsp;</td>" & vbcrlf
 	response.write "          <td valign=""top""><strong>To:</strong>" & vbcrlf
 	response.write "              <input type=""text"" id=""toDate"" name=""toDate"" value=""" & DateAdd("d", -1, sToDate) & """ />&nbsp;" & vbcrlf
 	response.write "              <img src=""../images/calendar.gif"" name=""toDateCalPop"" id=""toDateCalPop"" border=""0"" class=""searchDateImg"" />" & vbcrlf
 	response.write "          </td>" & vbcrlf
 	response.write "      </tr>" & vbcrlf
 	response.write "  </table>" & vbcrlf
 	response.write "  <p><input type=""button"" name=""buttonShowPayments"" id=""buttonShowPayments"" value=""Show Payments"" /></p>" & vbcrlf
 	response.write "</fieldset>" & vbcrlf
 	response.write "</form>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub GetRentalPayments( ByVal iUserId, ByVal sFromDate, ByVal sToDate )
	Dim varWhereClause, sSql, oRs

	varWhereClause = " AND (P.paymentdate >= '" & sFromDate & "' AND P.paymentdate < '" & sToDate & "') "
	varWhereClause = varWhereClause & " AND P.userid = " & iUserId & " "

	sSql = "SELECT P.paymentid, U.userfname, U.userlname, P.paymenttotal, P.paymentdate, P.reservationid, J.journalentrytype "
	sSql = sSql & " FROM egov_class_payment P, egov_users U, egov_journal_entry_types J "
	sSql = sSql & " WHERE  U.userid = P.userid AND P.isforrentals = 1 AND P.reservationid IS NOT NULL "
	sSql = sSql & " AND J.journalentrytypeid = P.journalentrytypeid " & varWhereClause
	sSql = sSql & " AND P.orgid = " & iOrgId
	sSql = sSql & " ORDER BY P.paymentid "
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	Do While Not oRs.EOF 
		oDynamicRs.AddNew
		oDynamicRs("purchaseday") = Day(oRs("paymentdate"))
		oDynamicRs("purchasemonth") = Month(oRs("paymentdate"))
		oDynamicRs("purchaseyear") = Year(oRs("paymentdate"))
		oDynamicRs("purchasedate") = DateValue(oRs("paymentdate"))
		oDynamicRs("purchaseid") = oRs("paymentid")
		If oRs("journalentrytype") = "rentalpayment" Then 
			oDynamicRs("whatpurchased") = "Reservation &ndash; " & GetFirsRentalName( oRs("reservationid") )
		Else
			oDynamicRs("whatpurchased") = "Reservation &ndash; Refund"
		End If 
	
	oDynamicRs("purchaser") = oRs("userfname") & " " & oRs("userlname")
		oDynamicRs("amount") = oRs("paymenttotal") + GetProcessingFee( oRs("paymentid") )
  		oDynamicRs("detailurl") = "../rentals/view_receipt.asp?iPaymentId=" & oRs("paymentid") 
		'oDynamicRs("detailurl")     = "../rd_classes/class_receipt.aspx?paymentid=" & oRs("paymentid") 
		oDynamicRs.Update

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

end sub

'------------------------------------------------------------------------------
function GetFirsRentalName( ByVal iReservationId )
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

end function
%>
