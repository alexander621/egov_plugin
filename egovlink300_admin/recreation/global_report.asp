<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: global_report.asp
' AUTHOR: Steve Loar
' CREATED: 05/23/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the global financial report for recreation
'
' MODIFICATION HISTORY
' 1.0   05/23/2006   Steve Loar - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
' 1.1	06/26/07	Steve Loar - Menlo Park project changes
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "rec finance rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

Dim fromDate, toDate, tmpdate, curReportTotal

fromDate = Request("fromDate")
toDate = Request("toDate")

If toDate = "" or IsNull(toDate) Then 
	' set to today
	toDate = dateAdd("d",0,Date()) 
Else
	toDate = CDate(Request("toDate"))
End If

If fromDate = "" or IsNull(fromDate) Then 
	' set to 1/1 of this year
	fromDate = DateSerial(Year(Now()),1,1) 
Else 
	fromDate = CDate(Request("fromDate"))
End If

If toDate < fromDate Then
	tmpdate = toDate
	toDate = fromDate
	fromDate = tmpdate
End If 

sEndDate = dateAdd("d",1,toDate) 

curReportTotal = 0.00

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="global_report.css" />
	<link rel="stylesheet" type="text/css" media="print" href="receiptprint.css" />

<script language="JavaScript">
 <!--

	function doCalendar( ToFrom ) 
	{
		w = (screen.width - 350)/2;
		h = (screen.height - 350)/2;
		eval('window.open("gr_calendarpicker.asp?updateform=searchForm&updatefield=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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
		
		<h3><%=GetOrgName( Session("orgid") )%> Recreation Financial Report</h3>

		<div id="topbuttons">
			
		</div>
		<!--<p id="backbutton">
			<img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a>
		</p>-->

<!--BEGIN: SEARCH OPTIONS-->
	<script>
		function validate()
		{
			if (!isDate(document.getElementById("fromDate").value))
			{
				alert("Your From Date is not a valid date or is not in a valid format MM/DD/YYYY");
			}
			if (!isDate(document.getElementById("toDate").value))
			{
				alert("Your To Date is not a valid date or is not in a valid format MM/DD/YYYY");
			}

			if (isDate(document.getElementById("fromDate").value) && isDate(document.getElementById("toDate").value))
			{
				document.searchForm.submit();
			}
		}


		function isDate(txtDate)
		{
			var currVal = txtDate;
			if(currVal == '')
				return false;
  			
  			//Declare Regex  
			var rxDatePattern = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/; 
			var dtArray = currVal.match(rxDatePattern); // is format OK?
			
			if (dtArray == null)
				return false;
 			
			//Checks for mm/dd/yyyy format.
			dtMonth = dtArray[1];
			dtDay= dtArray[3];
			dtYear = dtArray[5];
			
			if (dtMonth < 1 || dtMonth > 12)
    			return false;
			else if (dtDay < 1 || dtDay> 31)
    			return false;
			else if ((dtMonth==4 || dtMonth==6 || dtMonth==9 || dtMonth==11) && dtDay ==31)
    			return false;
			else if (dtMonth == 2)
			{
   			var isleap = (dtYear % 4 == 0 && (dtYear % 100 != 0 || dtYear % 400 == 0));
   			if (dtDay> 29 || (dtDay ==29 && !isleap))
       			return false;
			}
			return true;
		}
	</script>
	<fieldset id="search">
		<legend><strong>Purchase Date Range</strong></legend>
		<form action="global_report.asp" method="post" name="searchForm">
			<strong>From:</strong>
			<input type="text" name="fromDate" id="fromDate" value="<%=fromDate%>" />
			<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>
			<span class="searchelement"><strong>To:</strong></span>
			<input type="text" name="toDate" id="toDate" value="<%=toDate%>" />
			<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
			<span class="searchelement"><input type="button" value="View Report" class="button" onClick="validate()" /></span>
			<input type="button" onclick="javascript:window.print();" class="button" id="globalreportprint" value="Print" />
		</form>
	</fieldset>

<!--END: SEARCH OPTIONS-->

	<% 
	If request.servervariables("REQUEST_METHOD") = "POST" Then 
		response.write "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"">"

		If OrgHasFeature( "facilities" ) Then %>

		<%
			' Facility RESERVATION TOTALS
			curReportTotal = curReportTotal + DisplayReservationTotals( fromDate, sEndDate )
		%>

	<% End If 

	If OrgHasFeature( "rentals" ) Then %>

		<%
			' Rentals Totals
			curReportTotal = curReportTotal + DisplayRentalReservations( fromDate, sEndDate )
		%>

<% 
	End If

	If OrgHasFeature( "gifts" ) Then %>

		<%
			' Commemorative GIFT TOTALS
			curReportTotal = curReportTotal + DisplayGiftTotals( fromDate, sEndDate )
		%>

	<% End If %>

	<% If OrgHasFeature( "memberships" ) Then %>

		<%
			' Pool Passes Totals
			curReportTotal = curReportTotal + DisplayPoolPassTotals( fromDate, sEndDate )
		%>

	<% End If %>

	<% If OrgHasFeature( "activities" ) Then %>

		<%
			' Class and Event Program Totals
			curReportTotal = curReportTotal + DisplayClassTotals( fromDate, sEndDate )
		%>

	<% End If %>

	<% If OrgHasFeature( "merchandise" ) Then %>

		<%
			' Merchandise Totals
			curReportTotal = curReportTotal + DisplayMerchandiseTotals( fromDate, sEndDate )
		%>

	<% End If %>



	<!-- <table class="instructortable" cellpadding="5" cellspacing="0" border="0"> -->
		<tr id="rectotalrow"><td><strong>Recreation Total:</strong></td><td class="amount" colspan="4"><strong><%=FormatCurrency(curReportTotal,2)%></strong></td></tr>
	</table>

<%	Else	
		response.write "<p><strong>To view the Recreation Financial Report, select from the purchase date range options above then click the &quot;View Report&quot; button.</strong></p>"
	End If	%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>

<%

'--------------------------------------------------------------------------------------------------
' double DisplayGiftTotals( sFromDate, sToDate )	
'--------------------------------------------------------------------------------------------------
Function DisplayGiftTotals( ByVal sFromDate, ByVal sToDate )
	Dim iGrandCount, curGrandTotal, sSql, oGift, curTotal, iWebCount, iOfficeCount, sGiftName, iWebGrandCount, iOfficeGrandCount

		iGrandCount = 0 
		iWebCount = 0
		iOfficeCount = 0
		iWebGrandCount = 0
		iOfficeGrandCount = 0
		curGrandTotal = 0.00
	
		sSql = "SELECT dbo.egov_gift.giftname, paymentlocation, SUM(dbo.egov_gift.amount) AS totalamount, "
		sSql = sSql & " COUNT(dbo.egov_gift_payment.giftpaymentid) AS giftcount, dbo.egov_gift.orgid "
		sSql = sSql & " FROM dbo.egov_gift_payment RIGHT OUTER JOIN "
		sSql = sSql & " dbo.egov_gift ON dbo.egov_gift_payment.giftid = dbo.egov_gift.giftid "
		sSql = sSql & " where (paymentdate Between '" & sFromDate & "' AND '" & sToDate & "') "
		sSql = sSql & " and result = 'APPROVED' and dbo.egov_gift.orgid = " & session("orgid")
		sSql = sSql & " GROUP BY dbo.egov_gift_payment.giftid, dbo.egov_gift.giftname, paymentlocation, dbo.egov_gift.orgid "
		sSql = sSql & " ORDER BY dbo.egov_gift.giftname, paymentlocation "

		Set oGift = Server.CreateObject("ADODB.Recordset")
		oGift.Open sSql, Application("DSN"), 0, 1

		response.write "<tr><th class=""nametitle"">Commemorative Gifts</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		
		If Not oGift.EOF Then
			'response.write "<h3>Commemorative Gifts</h3><br />"
			'response.write "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"">"
			'response.write "<tr><th class=""nametitle"">Gift Name</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
			'response.write "<tr><th class=""nametitle"">Commemorative Gifts</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
			
			sGiftName = oGift("giftname")
			Do While NOT oGift.EOF 
				If sGiftName <> oGift("giftname") Then
					response.write "<tr><td nowrap=""nowrap"">" & sGiftName & "</td><td class=""amount"">" & iWebCount & "</td><td class=""amount"">" & iOfficeCount & "</td><td class=""amount"">" & iWebCount + iOfficeCount & "</td><td class=""amount"">" & FormatCurrency(curTotal,2) & "</td></tr>"
					sGiftName = oGift("giftname")
					iWebCount = 0
					iOfficeCount = 0
					curTotal = 0.00
				End If 
				If clng(oGift("paymentlocation")) = 3 Then 
					iWebCount = iWebCount + clng(oGift("giftcount"))
					iWebGrandCount = iWebGrandCount + iWebCount
				Else 
					iOfficeCount = iOfficeCount + clng(oGift("giftcount"))
					iOfficeGrandCount = iOfficeGrandCount + iOfficeCount
				End If 
				curTotal = curTotal + CDbl(oGift("totalamount"))
				iGrandCount = iGrandCount + clng(oGift("giftcount"))
				curGrandTotal = curGrandTotal + CDbl(oGift("totalamount"))
				
				oGift.MoveNext
			Loop
			response.write "<tr><td nowrap=""nowrap"">" & sGiftName & "</td><td class=""amount"">" & iWebCount & "</td><td class=""amount"">" & iOfficeCount & "</td><td class=""amount"">" & iWebCount + iOfficeCount & "</td><td class=""amount"">" & FormatCurrency(curTotal,2) & "</td></tr>"
			'response.write "<tr class=""totalrow""><td><strong>Commemorative Gifts Total:</strong></td><td class=""amount""><strong>" & iWebGrandCount & "</strong></td><td class=""amount""><strong>" & iOfficeGrandCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"
			'response.write "</table>"
	
		End If

		response.write "<tr class=""totalrow""><td><strong>Commemorative Gifts Total:</strong></td><td class=""amount""><strong>" & iWebGrandCount & "</strong></td><td class=""amount""><strong>" & iOfficeGrandCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"

		oGift.close
		Set oGift = Nothing
		DisplayGiftTotals = curGrandTotal

End Function 


'--------------------------------------------------------------------------------------------------
' double DisplayRentalReservations( sFromDate, sToDate )
'--------------------------------------------------------------------------------------------------
Function DisplayRentalReservations( ByVal sFromDate, ByVal sToDate )
	Dim sSql, oRs, iGrandCount, curGrandTotal, sRentalname, iWebCount, iOfficeCount, curTotal
	Dim iWebGrandCount, iOfficeGrandCount

	iWebCount = 0
	iOfficeCount = 0
	curTotal = 0.00
	iGrandCount = 0 
	iWebGrandCount = 0
	iOfficeGrandCount = 0 
	curGrandTotal = 0.00

	sSql = "SELECT COUNT(R.reservationid) AS reservationcount, SUM(R.totalamount) AS reservationsamount, T.rentalname, "
	sSql = sSql & "CASE ISNULL(R.adminuserid,0) WHEN 0 THEN 'P' ELSE 'A' END AS paymentlocation "
	sSql = sSql & "FROM egov_rentalreservations R, egov_rentals T, egov_rentalreservationstatuses S "
	sSql = sSql & "WHERE R.originalrentalid = T.rentalid AND R.reservationstatusid = S.reservationstatusid AND S.isreserved = 1 "
	sSql = sSql & "AND (R.reserveddate between '" & sFromDate & "' AND '" & sToDate & "') AND R.orgid = " & session("orgid")
	sSql = sSql & " GROUP BY T.rentalname, CASE ISNULL(R.adminuserid,0) WHEN 0 THEN 'P' ELSE 'A' END "
	sSql = sSql & "ORDER BY T.rentalname, CASE ISNULL(R.adminuserid,0) WHEN 0 THEN 'P' ELSE 'A' END"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<tr><th class=""nametitle"">Rental Reservations</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"

	If Not oRs.EOF Then
		'response.write "<h3>Rental Reservations</h3><br />"
		'response.write "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"">"
		'response.write "<tr><th class=""nametitle"">Rental Name</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		'response.write "<tr><th class=""nametitle"">Rental Reservations</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		
		sRentalname = oRs("rentalname")
		Do While Not oRs.EOF
			' There may be as many as 2 rows per rental. One for public and one for admin side reservations, so group and then print
			If sRentalname <> oRs("rentalname") Then
				response.write "<tr><td nowrap=""nowrap"">" & sRentalname & "</td><td class=""amount"">" & iWebCount & "</td><td class=""amount"">" & iOfficeCount & "</td><td class=""amount"">" & (iWebCount + iOfficeCount) & "</td><td class=""amount"">" &   FormatCurrency(curTotal,2)  & "</td></tr>"
				sRentalname = oRs("rentalname")
				iWebCount = 0
				iOfficeCount = 0
				curTotal = 0.00
			End If 

			If oRs("paymentlocation") = "P" Then 
				iWebCount = iWebCount + clng(oRs("reservationcount"))
				iWebGrandCount = iWebGrandCount + clng(oRs("reservationcount"))
			Else 
				iOfficeCount = iOfficeCount + clng(oRs("reservationcount"))
				iOfficeGrandCount = iOfficeGrandCount + clng(oRs("reservationcount"))
			End If 
			curTotal = curTotal + CDbl(oRs("reservationsamount"))
			iGrandCount = iGrandCount + clng(oRs("reservationcount"))
			curGrandTotal = curGrandTotal + CDbl(oRs("reservationsamount"))

			oRs.MoveNext 
		Loop

		' Final rental row
		response.write "<tr><td nowrap=""nowrap"">" & sRentalname & "</td><td class=""amount"">" & iWebCount & "</td><td class=""amount"">" & iOfficeCount & "</td><td class=""amount"">" & iWebCount + iOfficeCount & "</td><td class=""amount"">" &   FormatCurrency(curTotal,2)  & "</td></tr>"

		' Totals Row
		'response.write "<tr class=""totalrow""><td nowrap=""nowrap""><strong>Rental Reservations Totals:</strong></td><td class=""amount""><strong>" & iWebGrandCount & "</strong></td><td class=""amount""><strong>" & iOfficeGrandCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"
		'response.write "</table>"
		
	End If 

	' Totals Row
	response.write "<tr class=""totalrow""><td nowrap=""nowrap""><strong>Rental Reservations Totals:</strong></td><td class=""amount""><strong>" & iWebGrandCount & "</strong></td><td class=""amount""><strong>" & iOfficeGrandCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' double DisplayReservationTotals( sFromDate, sToDate )
'--------------------------------------------------------------------------------------------------
Function DisplayReservationTotals( ByVal sFromDate, ByVal sToDate )
	Dim iGrandCount, curGrandTotal, sSql, oRs, sFacilityname, iWebCount, iOfficeCount, curTotal
	Dim iWebGrandCount, iOfficeGrandCount

	iWebCount = 0
	iOfficeCount = 0
	curTotal = 0.00
	iGrandCount = 0 
	iWebGrandCount = 0
	iOfficeGrandCount = 0 
	curGrandTotal = 0.00

'	sSql = "SELECT * FROM rpt_reservation_global_totals where orgid = '" & session("orgid") & "'"

	sSql = "SELECT COUNT(egov_facility.facilityid) AS ReservationCount, egov_facility.facilityname, paymentlocation, egov_facility.orgid, "
    sSql = sSql & " SUM(egov_facilityschedule.amount) AS ReservationsAmount, egov_facility.facilityid "
	sSql = sSql & " FROM egov_facilityschedule RIGHT OUTER JOIN "
    sSql = sSql & " egov_facility ON dbo.egov_facilityschedule.facilityid = egov_facility.facilityid "
	sSql = sSql & " where egov_facility.orgid = " & session("orgid") & " and status = 'RESERVED' and "
	sSql = sSql & " (datecreated Between '" & sFromDate & "' AND '" & sToDate & "') "
	sSql = sSql & " GROUP BY egov_facility.facilityid, egov_facility.facilityname, paymentlocation, egov_facility.orgid "
	sSql = sSql & " ORDER BY egov_facility.facilityname, paymentlocation "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<tr><th class=""nametitle"">Facility Reservations</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"

	If Not oRs.EOF Then
		'response.write "<h3>Facility Reservations</h3><br />"
		'response.write "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"">"
		'response.write "<tr><th class=""nametitle"">Facility Name</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		'response.write "<tr><th class=""nametitle"">Facility Reservations</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		
		sFacilityname = oRs("facilityname")
		Do While Not oRs.EOF 
			If sFacilityname <> oRs("facilityname") Then
				response.write "<tr><td nowrap=""nowrap"">" & sFacilityname & "</td><td class=""amount"">" & iWebCount & "</td><td class=""amount"">" & iOfficeCount & "</td><td class=""amount"">" & (iWebCount + iOfficeCount) & "</td><td class=""amount"">" &   FormatCurrency(curTotal,2)  & "</td></tr>"
				sFacilityname = oRs("facilityname")
				iWebCount = 0
				iOfficeCount = 0
				curTotal = 0.00
			End If 
			If clng(oRs("paymentlocation")) = 3 Then 
				iWebCount = iWebCount + clng(oRs("reservationcount"))
				iWebGrandCount = iWebGrandCount + clng(oRs("reservationcount"))
			Else 
				iOfficeCount = iOfficeCount + clng(oRs("reservationcount"))
				iOfficeGrandCount = iOfficeGrandCount + clng(oRs("reservationcount"))
			End If 
			curTotal = curTotal + CDbl(oRs("reservationsamount"))
			iGrandCount = iGrandCount + clng(oRs("reservationcount"))
			curGrandTotal = curGrandTotal + CDbl(oRs("reservationsamount"))
			
			oRs.MoveNext
		Loop
		response.write "<tr><td nowrap=""nowrap"">" & sFacilityname & "</td><td class=""amount"">" & iWebCount & "</td><td class=""amount"">" & iOfficeCount & "</td><td class=""amount"">" & iWebCount + iOfficeCount & "</td><td class=""amount"">" &   FormatCurrency(curTotal,2)  & "</td></tr>"
		'response.write "<tr class=""totalrow""><td nowrap=""nowrap""><strong>Facility Reservations Totals:</strong></td><td class=""amount""><strong>" & iWebGrandCount & "</strong></td><td class=""amount""><strong>" & iOfficeGrandCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"
		'response.write "</table>"

	End If

	response.write "<tr class=""totalrow""><td nowrap=""nowrap""><strong>Facility Reservations Totals:</strong></td><td class=""amount""><strong>" & iWebGrandCount & "</strong></td><td class=""amount""><strong>" & iOfficeGrandCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"

	oRs.close
	Set oRs = Nothing
	DisplayReservationTotals = curGrandTotal

End Function 


'--------------------------------------------------------------------------------------------------
' double DisplayPoolPassTotals( sFromDate, sToDate )
'--------------------------------------------------------------------------------------------------
Function DisplayPoolPassTotals( ByVal sFromDate, ByVal sToDate )
	Dim iGrandCount, curGrandTotal, sSql, oPool, iWebCount, iOfficeCount, iGrandWebCount, iGrandOfficeCount
	Dim sPassName, curTotal

		iGrandCount = 0 
		iGrandWebCount = 0
		iGrandOfficeCount = 0
		iWebCount = 0
		iOfficeCount = 0
		curGrandTotal = 0.00
		curTotal = 0.00
	
		sSql = "select  m.membershipdesc + ' - ' + T.description + ' ' + R.description as passname, paymentlocation, count(P.poolpassid) as totalcount, "
		sSql = sSql & " sum(paymentamount) as totalamount  "
		sSql = sSql & " from egov_poolpassresidenttypes T, egov_poolpassrates R, egov_poolpasspurchases P, egov_memberships m "
		sSql = sSql & " where T.orgid = " & session("orgid") & " and T.orgid = R.orgid  "
		sSql = sSql & " and T.orgid = P.orgid  and T.resident_type = R.residenttype "
		sSql = sSql & " and p.membershipid = m.membershipid "
		sSql = sSql & " and P.rateid = R.rateid and P.paymentresult <> 'Pending' and P.paymentresult <> 'Declined' "
		sSql = sSql & " and (P.paymentdate Between '" & sFromDate & "' AND '" & sToDate & "') "
		sSql = sSql & " group by m.membershipdesc, T.description, R.description, paymentlocation, T.displayorder, R.displayorder "
		sSql = sSql & " order by m.membershipdesc, T.displayorder, R.displayorder, paymentlocation"
		'response.write sSql

		Set oPool = Server.CreateObject("ADODB.Recordset")
		oPool.Open sSql, Application("DSN"), 0, 1

		response.write "<tr><th class=""nametitle"">Memberships</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		
		If Not oPool.EOF Then
			'response.write "<h3>Pool Passes</h3><br />"
			'response.write "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"">"
			'response.write "<tr><th class=""nametitle"">Pass Type</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
			'response.write "<tr><th class=""nametitle"">Pool Passes</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"

			sPassName = oPool("passname")
			Do While NOT oPool.EOF 
				If sPassName <> oPool("passname") Then 
					' New pass type so print out
					response.write "<tr><td nowrap=""nowrap"">" & sPassName & "</td><td class=""amount"">" &  iWebCount & "</td><td class=""amount"">" &  iOfficeCount & "</td><td class=""amount"">" &  iWebCount + iOfficeCount & "</td><td class=""amount"">" &   FormatCurrency(curTotal,2)  & "</td></tr>"
					' init for next type
					sPassName = oPool("passname")
					iWebCount = 0
					iOfficeCount = 0
					curTotal = 0.00
				End If 
				If UCase(oPool("paymentlocation")) = "ONLINE" Then
					iWebCount = iWebCount + clng(oPool("totalcount"))
					iGrandWebCount = iGrandWebCount + clng(oPool("totalcount"))
				Else
					iOfficeCount = iOfficeCount + clng(oPool("totalcount"))
					iGrandOfficeCount = iGrandOfficeCount + iOfficeCount
				End If 
				curTotal = curTotal + CDbl(oPool("totalamount"))
				iGrandCount = iGrandCount + clng(oPool("totalcount"))
				curGrandTotal = curGrandTotal + CDbl(oPool("totalamount"))
				
				oPool.MoveNext
			Loop
			response.write "<tr><td nowrap=""nowrap"">" & sPassName & "</td><td class=""amount"">" &  iWebCount & "</td><td class=""amount"">" &  iOfficeCount & "</td><td class=""amount"">" &  CStr(iWebCount + iOfficeCount) & "</td><td class=""amount"">" &   FormatCurrency(curTotal,2)  & "</td></tr>"
			'response.write "<tr class=""totalrow""><td><strong>Pool Passes Total:</strong></td><td class=""amount""><strong>" & iGrandWebCount & "</strong></td><td class=""amount""><strong>" & iGrandOfficeCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"
			'response.write "</table>"
	
		End If

		response.write "<tr class=""totalrow""><td><strong>Pool Passes Total:</strong></td><td class=""amount""><strong>" & iGrandWebCount & "</strong></td><td class=""amount""><strong>" & iGrandOfficeCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"

		oPool.close
		Set oPool = Nothing
		DisplayPoolPassTotals = curGrandTotal

End Function 


'--------------------------------------------------------------------------------------------------
' Function DisplayClassTotals( sFromDate, sToDate )
'--------------------------------------------------------------------------------------------------
Function DisplayClassTotals( ByVal sFromDate, ByVal sToDate )
	Dim iGrandCount, curGrandTotal, sSql, oCategory, iCatCount, curCatTotal, iCatWebCount, iCatOfficeCount
	Dim iGrandWebCount, iGrandOfficeCount, sCategoryTitle

	iGrandCount = 0 
	curGrandTotal = 0.00

	sSql = "SELECT categoryid, categorytitle FROM egov_class_categories WHERE orgid = " & session("orgid") & " AND isroot = 0 ORDER BY sequenceid"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSql, Application("DSN"), 0, 1

	response.write "<tr><th class=""nametitle"">Class and Event Programs</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
	
	If NOT oCategory.EOF Then
		'response.write "<h3>Class and Event Programs</h3><br />"

		'response.write "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"" >"
		'response.write "<tr><th class=""nametitle"">Program Name</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		'response.write "<tr><th class=""nametitle"">Class and Event Programs</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"

		Do While Not oCategory.EOF 
			iCatCount = 0
			iCatWebCount = 0
			iCatOfficeCount = 0
			curCatTotal = 0.00

			' Get the classes
			'response.write "<br /><br />" & oCategory("categoryid") ' this is a debugging line
			GetClasses oCategory("categoryid"), iCatCount, curCatTotal, sFromDate, sToDate, iCatWebCount, iCatOfficeCount
			iCatCount = iCatWebCount + iCatOfficeCount
			iGrandCount = iGrandCount + iCatCount
			iGrandWebCount = iGrandWebCount + iCatWebCount
			iGrandOfficeCount = iGrandOfficeCount + iCatOfficeCount
			curGrandTotal = curGrandTotal + curCatTotal
			sCategoryTitle = oCategory("categorytitle")
			response.write "<tr><td nowrap=""nowrap"">" & sCategoryTitle & "</td><td class=""amount"">" & iCatWebCount & "</td><td class=""amount"">" & iCatOfficeCount & "</td><td class=""amount"">" & iCatCount & "</td><td class=""amount"">" & FormatCurrency(curCatTotal,2) & "</td></tr>"
			'response.write "<tr><td nowrap=""nowrap"">" & sCategoryTitle & " (" & oCategory("categoryid") & ")</td><td class=""amount"">" & iCatWebCount & "</td><td class=""amount"">" & iCatOfficeCount & "</td><td class=""amount"">" & iCatCount & "</td><td class=""amount"">" & FormatCurrency(curCatTotal,2) & "</td></tr>"
			oCategory.MoveNext
		Loop
		'response.write "<tr class=""totalrow""><td><strong>Class and Event Programs Total:</strong></td><td class=""amount""><strong>" & iGrandWebCount & "</strong></td><td class=""amount""><strong>" & iGrandOfficeCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"
		'response.write "</table>"

	End If

	response.write "<tr class=""totalrow""><td><strong>Class and Event Programs Total:</strong></td><td class=""amount""><strong>" & iGrandWebCount & "</strong></td><td class=""amount""><strong>" & iGrandOfficeCount & "</strong></td><td class=""amount""><strong>" & iGrandCount & "</strong></td><td class=""amount""><strong>" & FormatCurrency(curGrandTotal,2) & "</strong></td></tr>"

	oCategory.close
	Set oCategory = Nothing

	DisplayClassTotals = curGrandTotal

End Function 


'--------------------------------------------------------------------------------------------------
' Function DisplayMerchandiseTotals( sFromDate, sEndDate )
'--------------------------------------------------------------------------------------------------
Function DisplayMerchandiseTotals( ByVal sFromDate, ByVal sEndDate )
	Dim sSql, oRs, iOrderCount, iWebCount, iOfficeCount, iCombinedCount, dAmount
	Dim iCombinedTotal, dAmountTotal, iWebTotal, iOfficeTotal, sOldMerchandise

	dAmountTotal = CDbl(0.00)
	iWebTotal = CLng(0)
	iOfficeTotal = CLng(0)
	iCombinedTotal = CLng(0)
	sOldMerchandise = "~~"

	sSql = "SELECT I.merchandise, L.ispublicmethod, "
	sSql = sSql & " SUM(I.quantity) AS quantity, SUM(I.quantity * I.itemprice) AS amount "
	sSql = sSql & " FROM egov_merchandiseorderitems I, egov_merchandiseorders O, egov_class_payment P, egov_paymentlocations L "
	sSql = sSql & " WHERE I.merchandiseorderid = O.merchandiseorderid AND O.orgid = " & session("orgid")
	sSql = sSql & " AND O.paymentid = P.paymentid AND P.paymentlocationid = L.paymentlocationid "
	sSql = sSql & " AND (O.orderdate BETWEEN '" & sFromDate & "' AND '" & sEndDate & "') "
	sSql = sSql & " GROUP BY I.merchandise, L.ispublicmethod "
	sSql = sSql & " ORDER BY I.merchandise"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<tr><th class=""nametitle"">Merchandise</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"

	If Not oRs.EOF Then
		'response.write vbcrlf & "<h3>Merchandise</h3><br />"
		'response.write vbcrlf & "<table class=""globalreporttable"" cellpadding=""5"" cellspacing=""0"" border=""0"" >"
		'response.write vbcrlf & "<tr><th class=""nametitle"">Item</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"
		'response.write vbcrlf & "<tr><th class=""nametitle"">Merchandise</th><th class=""numbers"">Online</th><th class=""numbers"">Office</th><th class=""numbers"">Total</th><th class=""numbers"">Amount</th></tr>"

		Do While Not oRs.EOF
			If sOldMerchandise <> oRs("merchandise") Then 
				If sOldMerchandise <> "~~" Then
					' Write the row out
					response.write vbcrlf & "<tr><td>" & sOldMerchandise & "</td>"
					response.write "<td class=""amount"">" & iWebCount & "</td>"
					response.write "<td class=""amount"">" & iOfficeCount & "</td>"
					response.write "<td class=""amount"">" & iCombinedCount & "</td>"
					response.write "<td class=""amount"">" & FormatCurrency(dAmount,2) & "</td>"
					response.write "</tr>"
					' Add to the Grand totals
					dAmountTotal = dAmountTotal + dAmount
					iWebTotal = iWebTotal + iWebCount
					iOfficeTotal = iOfficeTotal + iOfficeCount
					iCombinedCount = iCombinedCount + iWebCount + iOfficeCount
					iCombinedTotal = iCombinedTotal + iWebCount + iOfficeCount
				End If 
				' Initialize the counts
				iWebCount = CLng(0)
				iOfficeCount = CLng(0)
				iCombinedCount = CLng(0)
				dAmount = CDbl(0.00)
				sOldMerchandise = oRs("merchandise")
			End If 
			If oRs("ispublicmethod") Then
				iWebCount = iWebCount + CLng(oRs("quantity"))
			Else
				iOfficeCount = iOfficeCount + CLng(oRs("quantity"))
			End If 
			iCombinedCount = iCombinedCount + CLng(oRs("quantity"))
			dAmount = dAmount + CDbl(oRs("amount"))
			oRs.MoveNext
		Loop
		' write out the last row
		If sOldMerchandise <> "~~" Then
			response.write vbcrlf & "<tr><td>" & sOldMerchandise & "</td>"
			response.write "<td class=""amount"">" & iWebCount & "</td>"
			response.write "<td class=""amount"">" & iOfficeCount & "</td>"
			response.write "<td class=""amount"">" & iCombinedCount & "</td>"
			response.write "<td class=""amount"">" & FormatCurrency(dAmount,2) & "</td>"
			response.write "</tr>"
			' Add to the Grand totals
			dAmountTotal = dAmountTotal + dAmount
			iWebTotal = iWebTotal + iWebCount
			iOfficeTotal = iOfficeTotal + iOfficeCount
			iCombinedCount = iCombinedCount + iWebCount + iOfficeCount
			iCombinedTotal = iCombinedTotal + iWebCount + iOfficeCount
		End If 

		'response.write "<tr class=""totalrow""><td><strong>Merchandise Total:</strong></td><td class=""amount""><strong>" & iWebTotal & "</strong></td><td class=""amount""><strong>" & iOfficeTotal & "</strong></td><td class=""amount""><strong>" & iCombinedTotal & "</strong></td><td class=""amount""><strong>" & FormatCurrency(dAmountTotal,2) & "</strong></td></tr>"
		'response.write "</table>"
	End If 

	response.write "<tr class=""totalrow""><td><strong>Merchandise Total:</strong></td><td class=""amount""><strong>" & iWebTotal & "</strong></td><td class=""amount""><strong>" & iOfficeTotal & "</strong></td><td class=""amount""><strong>" & iCombinedTotal & "</strong></td><td class=""amount""><strong>" & FormatCurrency(dAmountTotal,2) & "</strong></td></tr>"

	oRs.CLose
	Set oRs = Nothing 

	DisplayMerchandiseTotals = dAmountTotal

End Function 


'--------------------------------------------------------------------------------------------------
' void GetClasses( iCategoryId, ByRef iCatCount, ByRef curCatTotal, sFromDate, sToDate, iCatWebCount, iCatOfficeCount )
'--------------------------------------------------------------------------------------------------
Sub GetClasses( ByVal iCategoryId, ByRef iCatCount, ByRef curCatTotal, ByVal sFromDate, ByVal sToDate, ByRef iCatWebCount, ByRef iCatOfficeCount )
	' Get the classes for this category that are not already counted
	Dim sSql, oClasses, iClassCount, curClassTotal, iClassWebCount, iClassOfficeCount

	' Get the classes that were not part of an earlier category
	sSql = "Select C.classid, C.classname, G.categoryid, C.isparent From egov_class C, egov_class_category_to_class G "
	sSql = sSql & " Where C.orgid = " & session("orgid") & " and C.statusid = 1 and C.classid = G.classid and G.categoryid = " & iCategoryId & " and "
	sSql = sSql & " G.classid not in (Select classid From egov_class_category_to_class where categoryid < " & iCategoryId & ") "
	sSql = sSql & " AND C.classid IN (SELECT itemnumber FROM egov_class_purchases WHERE orgid = " & session("orgid") & " and paymentdate Between '" & sFromDate & "' AND '" & sToDate & "')"
	response.write vbcrlf & "<!--" & sSql & "-->" & vbcrlf

	Set oClasses = Server.CreateObject("ADODB.Recordset")
	oClasses.Open sSql, Application("DSN"), 0, 1

	Do While Not oClasses.EOF
		iClassCount = 0
		iClassWebCount = 0
		iClassOfficeCount = 0
		curClassTotal = 0.00 

		' Get the class list counts
		'GetClassListCounts oClasses("classid"), iClassCount, curClassTotal, sFromDate, sToDate, iClassWebCount, iClassOfficeCount
		curClassTotal = GetClassRevenue( oClasses("classid"), sFromDate, sToDate )
		iClassWebCount = GetEnrolledCount( oClasses("classid"), sFromDate, sToDate, "ispublicmethod" )
		iClassOfficeCount = GetEnrolledCount( oClasses("classid"), sFromDate, sToDate, "isadminmethod" )
		'response.write "<br /> category=" & iCategoryId & " classid=" & oClasses("classid") & " total$=" & curClassTotal  & " web=" & iClassWebCount  & " admin=" & iClassOfficeCount & "<br />"

		If oClasses("isparent") Then
			' if this is a parent level, just add the money, the enrollees count on the child class/events
			curCatTotal = curCatTotal + curClassTotal
		Else
			'iCatCount = iCatCount + iCatWebCount + iCatOfficeCount
			iCatWebCount = iCatWebCount + iClassWebCount
			iCatOfficeCount = iCatOfficeCount + iClassOfficeCount
			curCatTotal = curCatTotal + curClassTotal
		End If 
		oClasses.MoveNext
	Loop

	oClasses.close
	Set oClasses = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' double GetClassRevenue( iClassid, sFromDate, sToDate )
'--------------------------------------------------------------------------------------------------
Function GetClassRevenue( ByVal iClassid, ByVal sFromDate, ByVal sToDate )
	Dim cPurchaseAmount, cRefundAmount

	cPurchaseAmount = GetPeriodAmount( iClassid, sFromDate, sToDate, "credit", 1 )
	'response.write cPurchaseAmount & "<br />"
	cRefundAmount = GetPeriodAmount( iClassid, sFromDate, sToDate, "debit", 2 )
	'response.write cRefundAmount & "<br />"

	GetClassRevenue = CDbl(cPurchaseAmount) - CDbl(cRefundAmount)

End Function 


'--------------------------------------------------------------------------------------------------
' double GetPeriodAmount( iClassid, sFromDate, sToDate, sType, iJournalEntryTypeID )
'--------------------------------------------------------------------------------------------------
Function GetPeriodAmount( ByVal iClassid, ByVal sFromDate, ByVal sToDate, ByVal sType, ByVal iJournalEntryTypeID )
	Dim sSql, oAmount

	sSql = "SELECT L.orgid, C.classid, ISNULL(sum(L.amount),0.00) AS amount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_class_list C "
	sSql = sSql & " WHERE L.orgid = " & session("orgid") & " AND L.itemtypeid = 1 AND L.ispaymentaccount = 0 AND L.entrytype = '" & sType & "' "
	sSql = sSql & " AND L.paymentid = J.paymentid and (J.paymentdate BETWEEN '" & sFromDate & "' AND '" & sToDate & "') AND J.notes NOT LIKE 'System generated for data migration' "
	sSql = sSql & " AND journalentrytypeid = " & iJournalEntryTypeID & " AND C.classlistid = L.itemid AND C.classid = " & iClassid
	sSql = sSql & " GROUP BY L.orgid, C.classid"

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSql, Application("DSN"), 0, 1

	If Not oAmount.EOF Then 
		GetPeriodAmount = CDbl(oAmount("amount"))
	Else
		GetPeriodAmount = CDbl(0.00)
	End If 

	oAmount.Close
	Set oAmount = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetEnrolledCount( iClassid, sFromDate, sToDate, sPurchaseMethod )
'--------------------------------------------------------------------------------------------------
Function GetEnrolledCount( ByVal iClassid, ByVal sFromDate, ByVal sToDate, ByVal sPurchaseMethod )
	Dim iEnrolled, iDropped

	iEnrolled = GetPeriodCount( iClassid, sFromDate, sToDate, "credit", 1, sPurchaseMethod )

	iDropped = GetPeriodCount( iClassid, sFromDate, sToDate, "debit", 2, sPurchaseMethod )


	GetEnrolledCount = CLng(iEnrolled) - CLng(iDropped)

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetPeriodCount( iClassid, sFromDate, sToDate, sType, iJournalEntryTypeID, sPurchaseMethod )
'--------------------------------------------------------------------------------------------------
Function GetPeriodCount( ByVal iClassid, ByVal sFromDate, ByVal sToDate, ByVal sType, ByVal iJournalEntryTypeID, ByVal sPurchaseMethod )
	Dim sSql, oCount, iTotalCount

	iTotalCount = CLng(0)

	sSql = "SELECT distinct L.orgid, C.classid, classlistid, isnull(quantity,0) as quantity "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_class_list C, egov_paymentlocations P "
	sSql = sSql & " WHERE L.orgid = " & session("orgid") & " and L.itemtypeid = 1 and L.ispaymentaccount = 0 and L.entrytype = '" & sType & "' and J.notes not like 'System generated for data migration' "
	sSql = sSql & " and L.paymentid = J.paymentid and (J.paymentdate between '" & sFromDate & "' and '" & sToDate & "') and J.paymentlocationid = P.paymentlocationid "
	sSql = sSql & " and journalentrytypeid = " & iJournalEntryTypeID & " and C.classlistid = L.itemid and C.classid = " & iClassid & " and " & sPurchaseMethod & " = 1 "
	sSql = sSql & " and C.status not in ('WAITLIST', 'WAITLIST REMOVED') "
'	If clng(iclassid) = clng(247) Then
'		response.write sSql & "<br />"
'	End If 

	Set oCount = Server.CreateObject("ADODB.Recordset")
	oCount.Open sSql, Application("DSN"), 0, 1

	Do While Not oCount.EOF 
		iTotalCount = iTotalCount + CLng(oCount("quantity"))
		oCount.MoveNext
	Loop 

	oCount.Close
	Set oCount = Nothing 
	GetPeriodCount = iTotalCount

End Function 


'--------------------------------------------------------------------------------------------------
' void GetClassListCounts( iClassid, ByRef iClassCount, ByRef curClassTotal, sFromDate, sToDate, iWebCount, iOfficeCount )
'--------------------------------------------------------------------------------------------------
Sub GetClassListCounts( ByVal iClassid, ByRef iClassCount, ByRef curClassTotal, ByVal sFromDate, ByVal sToDate, ByRef iWebCount, ByRef iOfficeCount )
	' Get the sum of attendees and tickets sold, plus the amount paid minus refunds
	Dim sSql, oClassList

	sSql = "select paymentlocationid, L.status, sum(isnull(quantity,0)) as totalcount, sum(isnull(amount,0) - isnull(refundamount,0)) as totalamount "
	sSql = sSql & " from egov_class_list L, egov_class_payment P where L.classid = " & iClassid
	sSql = sSql & " and (signupdate between '" & sFromDate & "' and '" & sToDate & "') and L.paymentid = P.paymentid"
	sSql = sSql & " group by paymentlocationid, L.status order by paymentlocationid, L.status"

'	response.write sSql & "<br />"
	Set oClassList = Server.CreateObject("ADODB.Recordset")
	oClassList.Open sSql, Application("DSN"), 0, 1

	Do While NOT oClassList.EOF 
		If Not IsNull(oClassList("totalcount")) Then 
			If UCase(oClassList("status")) = "ACTIVE" Then 
				' Only count the active students
				iClassCount = iClassCount + clng(oClassList("totalcount"))
				'response.write "<br /> --------------------- PaymentLocationId: " & oClassList("paymentlocationid")
				If clng(oClassList("paymentlocationid")) = 3 Then
					iWebCount = iWebCount + clng(oClassList("totalcount"))
				Else
					iOfficeCount = iOfficeCount + clng(oClassList("totalcount"))
				End If 
			End If 
		End If 
		' Add the dollars for active and dropped
		If Not IsNull(oClassList("totalamount")) Then 
			curClassTotal = curClassTotal + CDbl(oClassList("totalamount"))
		End If 
		'response.write "<br /> &nbsp;&nbsp;" & iClassid & " amount " & oClassList("totalamount") & "  Total: " & curClassTotal
		oClassList.movenext
	Loop 
	'response.write "<br /> &nbsp;&nbsp;" & iClassid & " Web: " & iWebCount & " Office: " & iOfficeCount
	

	oClassList.close
	Set oClassList = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' string GetOrgName( iOrgId )
'--------------------------------------------------------------------------------------------------
Function GetOrgName( ByVal iOrgId )
	Dim sSql, oName

	sSql = "Select orgname from organizations where orgid = " & iOrgId

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 0, 1

	If Not oName.EOF Then
		GetOrgName = oName("orgname")
	Else 
		GetOrgName = ""
	End If 

	oName.close
	Set oName = Nothing

End Function 



%>


