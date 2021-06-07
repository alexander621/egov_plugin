<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="merchandisecommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiseorderlist.asp
' AUTHOR: Steve Loar
' CREATED: 05/06/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	05/06/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sToPurchaseDate, sFromPurchaseDate, sSearch, sDisplayDateRange, iPaymentId, sBuyerName

sLevel = "../" ' Override of value from common.asp
sSearch = ""

' check if page is online and user has permissions in one call not two
PageDisplayCheck "merchandise orders", sLevel	' In common.asp

' Handle inspection date range. always want some dates to limit the search
If request("topurchasedate") <> "" And request("frompurchasedate") <> "" Then
	sFromPurchaseDate = request("frompurchasedate")
	sToPurchaseDate = request("topurchasedate")
	sSearch = sSearch & " AND (O.orderdate >= '" & request("frompurchasedate") & "' AND O.orderdate < '" & DateAdd("d",1,request("topurchasedate")) & "' ) "
	sDisplayDateRange = "From: " & request("frominspectiondate") & " &nbsp;To: " & request("topurchasedate")
Else
	' initially set these to yesterday
	sFromPurchaseDate = FormatDateTime(DateAdd("d",-1,Date),2)
	sToPurchaseDate = FormatDateTime(DateAdd("d",-1,Date),2)
	sDisplayDateRange = ""
End If 

If request("paymentid") <> "" Then
	iPaymentId = CLng(request("paymentid"))
	sSearch = sSearch & " AND O.paymentid = " & iPaymentId
Else
	iPaymentId = ""
End If 

If request("buyername") <> "" Then
	sBuyerName = request("buyername")
	sSearch = sSearch & " AND (UPPER(U.userfname) LIKE '%" & dbsafe(sBuyerName) & "%' OR UPPER(U.userlname) LIKE '%" & dbsafe(sBuyerName) & "%') "
Else
	sBuyerName = ""
End If 

'If sSearch <> "" Then 
'	session("sSearch") = sSearch 
'End If 

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="merchandise.css" />

	<script language="Javascript" src="tablesort.js"></script>
	<script language="Javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function Validate()
		{
			// check the frompurchasedate
			if (document.MerchandiseList.frompurchasedate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.MerchandiseList.frompurchasedate.value);
				if (! Ok)
				{
					alert("From purchase date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.MerchandiseList.frompurchasedate.focus();
					return;
				}
			}
			// check the topurchasedate
			if (document.MerchandiseList.topurchasedate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.MerchandiseList.topurchasedate.value);
				if (! Ok)
				{
					alert("To purchase date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.MerchandiseList.topurchasedate.focus();
					return;
				}
			}
			// Check the receipt number
			if (document.MerchandiseList.paymentid.value != "")
			{
				rege = /^\d*$/;
				Ok = rege.test(document.MerchandiseList.paymentid.value);
				if (! Ok)
				{
					alert("The Receipt # must be a number.  \nPlease enter it again.");
					document.MerchandiseList.paymentid.focus();
					return;
				}
			}
			document.MerchandiseList.submit();
		}

		function ExportOrders()
		{
			// build a string of selected teams
			var sOrderPicks = '';
			for (var t = 0; t <= parseInt($("ordercount").value); t++)
			{
				// See if a row exists for this one
				if ($("selectorder" + t))
				{
					// If it is marked for export, then add it
					if ($("selectorder" + t).checked == true)
					{
						if (sOrderPicks != '')
						{
							sOrderPicks += ',' ;
						}
						sOrderPicks += $("selectorder" + t).value;
					}
				}
			}
			if (sOrderPicks != '')
			{
				sOrderPicks = '(' + sOrderPicks + ')';
				location.href='merchandiseorderexport.asp?orderpicks=' + sOrderPicks;
			}
			else
			{
				alert("Please select some orders for the export, first.");
			}
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

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Merchandise Orders</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

	<!--BEGIN: FILTER SELECTION-->
	<div class="filterselection">
	 	<fieldset class="filterselection">
			<legend class="filterselection">Search Options</legend>
			<p>
				<form name="MerchandiseList" method="post" action="merchandiseorderlist.asp">
					<input type="hidden" id="isview" name="isview" value="1" />
					<table border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td>Order Date:</td>
							<td nowrap="nowrap">
								From:
								<input type="text" id="frompurchasedate" name="frompurchasedate" value="<%=sFromPurchaseDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('frompurchasedate');"><img src="../images/calendar.gif" border="0" /></a>
								&nbsp; To:
								<input type="text" id="topurchasedate" name="topurchasedate" value="<%=sToPurchaseDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('topurchasedate');"><img src="../images/calendar.gif" border="0" /></a>
								&nbsp;
								<%DrawDateChoices "purchasedate" %>
							</td>
						</tr>
						<tr>
							<td>Receipt #:</td>
							<td>
								<input type="text" maxlength="6" size="6" id="paymentid" name="paymentid" value="<%=iPaymentId%>" />
							</td>
						</tr>
						<tr>
							<td>Buyer Like:</td>
							<td>
								<input type="text" maxlength="50" size="50" id="buyername" name="buyername" value="<%=sBuyerName%>" />
							</td>
						</tr>
						<tr>
							<td>&nbsp;</td>
							<td colspan="2"><input class="button" type="button"  onclick="Validate()" value="Refresh Results" /> &nbsp;
<%							If request("isview") <> "" Then	
								response.write vbcrlf & "<input type=""button"" class=""button"" value=""Export Selected Orders to Excel"" onclick=""ExportOrders()"" />"
							End If 
%>
							</td>
						</tr>
					</table>
				</form>
			</p>
 		</fieldset>
	</div>
	<!--END: FILTER SELECTION-->

	<!--BEGIN: Merchandise Order LIST-->

	<% 
		If request("isview") <> "" Then		
			DisplayMerchandiseOrders sSearch
		Else 
			response.write "<p><strong>To view the merchandise orders, select from the filter options above then click the &quot;Refresh Results&quot; button.</strong></p>"
		End If 
	%>

	<!--END: Merchandise Order LIST-->
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
' Sub DisplayMerchandiseOrders( sSearch )
'--------------------------------------------------------------------------------------------------
Sub DisplayMerchandiseOrders( sSearch )
	Dim sSql, oRs

	sSql = "SELECT O.orderdate, O.paymentid, O.merchandiseorderid, ISNULL(O.taxamount,0.00) AS taxamount, "
	sSql = sSql & " O.orderamount, U.userfname, U.userlname, ISNULL(O.shippingfee,0.00) AS shippingfee "
	sSql = sSql & " FROM egov_merchandiseorders O, egov_users U "
	sSql = sSql & " WHERE O.userid = U.userid " & sSearch
	sSql = sSql & " AND O.orgid = " & session("orgid") & " ORDER BY orderdate DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		'DRAW TABLE WITH MERCHANDISE LISTED
		response.write vbcrlf & "<div id=""merchandiseorderlistshadow"">" 
		response.write vbcrlf & "<table class=""tableadmin sortable"" id=""merchandiseorderlist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		'HEADER ROW
		response.write vbcrlf & "<tr><th>&nbsp;</th><th>Order Date</th><th>Receipt #</th><th>Order #</th><th>Buyer</th>"
		response.write "<th>Merchandise<br />Total</th><th>Shipping</th><th>Sales<br />Tax</th><th>Total</th><th>View<br />Receipt</th></tr>"

		iRowCount = 0
		
		' LOOP THRU AND DISPLAY The EVENTS
		Do While Not oRs.EOF
  			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
  			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
  			End If 

			response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" 

			' Select the row
			response.write "<td align=""center""><input type=""checkbox"" id=""selectorder" & iRowCount & """ name=""selectorder" & iRowCount & """ value=""" & oRs("merchandiseorderid") & """ /></td>"

			' Order Date
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write FormatDateTime(oRs("orderdate"),2)
			response.write "</td>"

			' Receipt Number
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write oRs("paymentid")
			response.write "</td>"

			' Merchandise Order Number
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write oRs("merchandiseorderid")
			response.write "</td>"

			' Buyer Name
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write oRs("userfname") & " " & oRs("userlname")
			response.write "</td>"

			' Merchandise Total
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write FormatNumber(oRs("orderamount"),2)
			response.write "</td>"

			' Shipping
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write FormatNumber(oRs("shippingfee"),2)
			response.write "</td>"

			' Sales Tax
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			response.write FormatNumber(oRs("taxamount"),2)
			response.write "</td>"

			' Total
			response.write "<td align=""center"" onClick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"">"
			dTotal = CDbl(oRs("orderamount")) + CDbl(oRs("shippingfee")) + CDbl(oRs("taxamount"))
			response.write FormatNumber(dTotal,2)
			response.write "</td>"

			' View Receipt
			response.write "<td align=""center"">"
			response.write "<input type=""button"" class=""button"" value=""Receipt"" onclick=""location.href='../classes/view_receipt.asp?iPaymentId=" & oRs("paymentid") & "';"" />"
			'response.write "<input type=""button"" class=""button"" value=""Receipt"" onclick=""location.href='viewmerchandiseorder.asp?orderid=" & oRs("merchandiseorderid") & "';"" />"
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext
		Loop
		response.write "</table>"
		response.write "</div>"
		response.write "<input type=""hidden"" id=""ordercount"" name=""ordercount"" value=""" & iRowCount & """ />"
	Else
		response.write "<p><font color=""red""><b>No merchandise orders could be could be found that match your search criteria.</b></font></p>"
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Sub



%>


