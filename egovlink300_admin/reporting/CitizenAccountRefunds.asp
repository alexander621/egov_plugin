<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="reporting_functions.asp" //-->
<!-- #include file="citizen_reporting_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CitizenAccountRefunds.asp
' AUTHOR: SteveLoar
' CREATED: 01/25/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/22/2013	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim fromDate, toDate, today, sWhereClause, iLocationId, iAdminUserId, sNameLike
Dim from_time, to_time, where_time

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
PageDisplayCheck "citizen account refunds report", sLevel	' In common.asp

fromDate = Request("fromDate")
toDate = Request("toDate")
today = FormatDateTime(Date(),2)

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
	toDate = today 
End If

If fromDate = "" or IsNull(fromDate) Then 
	fromDate = today
End If

If Request("fromtime") <> "" Then 
	from_time = Request("fromtime")
Else 
	from_time = "none"
End If 
If Request("totime") <> "" Then 
	to_time = Request("totime")
Else
	to_time = "none"
End If 

If request("locationid") = "" Then
	iLocationId = 0
Else
	iLocationId = CLng(request("locationid"))
End If 

If request("adminuserid") = "" Then
	iAdminUserId = 0
Else
	iAdminUserId = CLng(request("adminuserid"))
End If 


' BUILD SQL WHERE CLAUSE
sWhereClause = " WHERE orgid = " & session("orgid") & " "

'sWhereClause = sWhereClause & " AND (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
If from_time = "none" Then 
	sWhereClause = sWhereClause & " AND paymentDate >= '" & fromDate & "' "
Else
	where_time = CDate( fromdate & " " & from_time )
	sWhereClause = sWhereClause & " AND paymentDate >= '" & where_time & "' "
End If 

If to_time = "none" Then 
	sWhereClause = sWhereClause & " AND paymentDate <= '" & DateAdd("d",1,toDate) & "' "
Else 
	where_time = CDate( todate & " " & to_time )
	sWhereClause = sWhereClause & " AND paymentDate <= '" & where_time & "' "
End If 

If iLocationId > 0 Then
	sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId & " "
End If 

If iAdminUserId > 0 Then
	sWhereClause = sWhereClause & " AND adminuserid = " & iAdminUserId & " "
End If 

If request("namelike") <> "" Then
	sNameLike = request("namelike")
	sWhereClause = sWhereClause & " AND ( userfname LIKE '%" & DBsafe( request("namelike") ) & "%' OR userlname LIKE '%" & DBsafe( request("namelike") ) & "%' )"
Else
	sNameLike = ""
End If 

%>


<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="reporting.css" />
	<link rel="stylesheet" href="pageprint.css" media="print" />
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">

	<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
  	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>

	<script src="../scripts/getdates.js"></script>
	<script src="../scripts/isvaliddate.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	<!--

		function validate( _Action )
		{
			var okToPost = true; 

			// check to date
			if ($("#toDate").val() != "")
			{
				if ( ! isValidDate( $("#toDate").val() ) )
				{
					inlineMsg("toDate",'<strong>Invalid Value: </strong>The withdrawl "To" date should be in the format of MM/DD/YYYY.',8,"toDate");
					okToPost = false;
				}
			}

			// check from date
			if ($("#fromDate").val() != "")
			{
				if ( ! isValidDate( $("#fromDate").val() ) )
				{
					inlineMsg("fromDate",'<strong>Invalid Value: </strong>The withdrawl "From" date should be in the format of MM/DD/YYYY.',8,"fromDate");
					okToPost = false;
				}
			}

			if (_Action === 'export')
			{
				document.reportFilter.action = "CitizenAccountRefundsExport.asp";
			}
			else
			{
				document.reportFilter.action = "CitizenAccountRefunds.asp";
			}

			if (okToPost)
			{
				document.reportFilter.submit();
			}
		}

		$(function() {
			$( "#toDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

		$(function() {
			$( "#fromDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

	//-->
	</script>
</head>
<body>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN: THIRD PARTY PRINT CONTROL-->
	<div id="idControls" class="noprint">
		<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
	</div>
	<!--END: THIRD PARTY PRINT CONTROL-->

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

		<form action="CitizenAccountRefunds.asp" method="post" name="reportFilter">

		<font size="+1"><b>Citizen Account Withdrawals</b></font><br /><br />

		<fieldset>
			<legend><strong>Select</strong></legend>
		
			<!--BEGIN: FILTERS-->
			<p>
			<table border="0" cellpadding="0" cellspacing="0" id="daterangepicks">
				<tr>
					<td><label for="fromDate">Withdrawl Date:</label>
						<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />&nbsp;
					</td>
					<td><% DrawTimeChoices "fromtime", from_time %></td>
					<td>
						<label for="toDate">To:</label>
						<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />&nbsp;
					</td>
					<td><% DrawTimeChoices "totime", to_time %></td>
					<td><%DrawDateChoices "Date" %></td>
				</tr>
			</table>
			</p>
			<p>
				<label for="locationid">Admin Location:</label><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
				<label for="adminuserid">Admin:</label><% ShowAdminUsers iAdminUserId %>
			</p>
			<p>
				<label for="namelike">Name Like:</label><input type="text" id="namelike" name="namelike" value="<%=sNameLike%>" placeholder="Partial Name of Citizen" maxlength="100" size="100" />
			</p>
			<p>
				<input class="button" type="button" value="View Report" onclick="validate('screen');" />
				&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="validate('export');" />
			</p>

		</fieldset>
		<!--END: FILTERS-->

		</form>

		<!-- REsults Here -->
<%
		' DISPLAY RESULTS
		'Display_Results sWhereClause
		Display_Citizen_Refund_Report sWhereClause

%>
		</div>
	</div>
	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%

'------------------------------------------------------------------------------------------------------------
' Display_Results sWhereClause 
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldPaymentId, dVoucherTotal,  dCardTotal, dMemoTotal, dGrandTotal, dSubTotal

	iOldPaymentId = CLng(0) 

	sSql = "SELECT paymentid, orgid, PaymentUserid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & "ISNULL(userhomephone,'') AS userhomephone, paymentdate, amount, isCCRefund, isRefundVoucher, isMemoTransfer "
	sSql = sSql & "FROM egov_Citizen_Account_Refunds " & sWhereClause
	sSql = sSql & " AND isExpiredCustomerCredit = 0 "
	sSql = sSql & "ORDER BY paymentid" 
	'response.write sSql & "<br /><br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	If oRequests.EOF then
		' Nothing found
		response.write "<p><strong>No refunds were found for your selection criteria.</strong></p>"
	Else
		' Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "paymentid", adInteger, , adFldUpdatable
		oDisplay.fields.append "paymentdate", adVariant, 10, adFldUpdatable
		oDisplay.fields.append "userid", adInteger, , adFldUpdatable
		oDisplay.fields.append "userfname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "userlname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "voucheramt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "cardamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "subtotal", adCurrency, , adFldUpdatable
		oDisplay.fields.append "memoamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "total", adCurrency, , adFldUpdatable

		oDisplay.CursorLocation = 3

		oDisplay.open 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("paymentid") = oRequests("paymentid")
				oDisplay("paymentdate") = DateValue(oRequests("paymentdate"))
				
				' this is who the money went to if a transfer
				oDisplay("userid") = oRequests("PaymentUserid")
				' this is who's account the money came from
				If Not IsNull(oRequests("userfname")) Then 
					oDisplay("userfname") = oRequests("userfname")
				End If 
				If Not IsNull(oRequests("userlname")) Then 
					oDisplay("userlname") = oRequests("userlname")
				End If 
				oDisplay("userhomephone") = oRequests("userhomephone")
				oDisplay("voucheramt") = 0.00
				oDisplay("cardamt") = 0.00
				oDisplay("subtotal") = 0.00
				oDisplay("memoamt") = 0.00
				oDisplay("total") = 0.00
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			If oRequests("isccrefund") Then
				' Credit Card Refund
				oDisplay("cardamt") = oRequests("amount")
				oDisplay("subtotal") = CDbl(oDisplay("subtotal")) + CDbl(oRequests("amount"))
			Else 
				If oRequests("isRefundVoucher") Then
					' Voucher Issued
					oDisplay("voucheramt") = oRequests("amount")
					oDisplay("subtotal") = CDbl(oDisplay("subtotal")) + CDbl(oRequests("amount"))
				Else
					' Refund To Memo account
					oDisplay("memoamt") = oRequests("amount")
				End If 
			End If 
			oDisplay("total") = CDbl(oDisplay("total")) + CDbl(oRequests("amount"))

			oDisplay.Update
			oRequests.MoveNext
		Loop

		' Show results
		oDisplay.MoveFirst
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"" id=""citizendeposits"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Account</th>"
		response.write "<th>Voucher<br />Amount</th><th>Card Amt</th><th>Total Card<br />&amp; Voucher</th><th>Memo Amt</th><th>Total<br />Amount</th></tr>"

		Do While Not oDisplay.EOF

			iRowCount = iRowCount + 1

			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"

			response.write "<td align=""center""><a href=""../purchases/viewjournal.asp?uid=" & oDisplay("userid") & "&pid=" & oDisplay("paymentid") & "&rt=c&it=ci&jet=d"">" & oDisplay("paymentid") & "</a></td>"
			response.write "<td align=""center"">" & oDisplay("paymentdate") & "</td>"
			response.write "<td align=""center"" valign=""top"">" & oDisplay("userfname") & " " & oDisplay("userlname") & "<br />" & FormatPhoneNumber(oDisplay("userhomephone")) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("voucheramt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cardamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("subtotal"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("memoamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("total"),2) & "</td>"
			response.write "</tr>"

			dVoucherTotal = dVoucherTotal + CDbl(oDisplay("voucheramt"))
			dCardTotal = dCardTotal + CDbl(oDisplay("cardamt"))
			dSubTotal = dSubTotal + CDbl(oDisplay("subtotal"))
			dMemoTotal = dMemoTotal + CDbl(oDisplay("memoamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("total"))

			oDisplay.MoveNext
		Loop 

		' Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">"
		response.write "<td colspan=""3"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dVoucherTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCardTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dMemoTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"

		oDisplay.Close
		Set oDisplay = Nothing 

	End If 

	oRequests.Close
	Set oRequests = Nothing 

End Sub 



%>
