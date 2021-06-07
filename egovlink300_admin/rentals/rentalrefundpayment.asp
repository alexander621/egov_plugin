<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="../reporting/reporting_functions.asp" //-->
<!-- #include file="../reporting/rental_reporting_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalrefundpayment.asp
' AUTHOR: SteveLoar
' CREATED: 12/17/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This report has rental refunds by payment type - Part of Menlo Park Project
'
' MODIFICATION HISTORY
' 1.0   12/17/2009	Steve Loar - INITIAL VERSION
' 1.1	05/11/2010	Steve Loar - Modified mechanics of date selections
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iLocationId, iAdminUserId, fromDate, toDate, today, sWhereClause
Dim from_time, to_time, where_time

' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "rental refund payment rpt", sLevel	' In common.asp

fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

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
sWhereClause = " WHERE orgid = " & session("orgid") 

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
	sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId
End If 

If iAdminUserId > 0 Then
	sWhereClause = sWhereClause & " AND adminuserid = " & iAdminUserId
End If 


%>

<html lang="en">
<head>
  	<meta charset="UTF-8">
  	
  	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="reporting.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalsstyles.css" />
	<link rel="stylesheet" href="receiptprint.css" media="print" />
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">

	<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
  	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>

	<script src="scripts/tablesort.js"></script>
	<script src="../scripts/getdates.js"></script>

	<script>
	  <!--
		
		function validate()
		{
			// check from date
			if (document.frmPFilter.fromDate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmPFilter.fromDate.value);
				if (! Ok)
				{
					alert("The refund from date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.fromDate.focus();
					return;
				}
			}
			// check to date
			if (document.frmPFilter.toDate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmPFilter.toDate.value);
				if (! Ok)
				{
					alert("The refund to date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.toDate.focus();
					return;
				}
			}

			document.frmPFilter.action = 'rentalrefundpayment.asp';
			document.frmPFilter.submit();
		}
		
		function exportReport()
		{
			// check from date
			if (document.frmPFilter.fromDate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmPFilter.fromDate.value);
				if (! Ok)
				{
					alert("The refund from date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.fromDate.focus();
					return;
				}
			}
			// check to date
			if (document.frmPFilter.toDate.value != "")
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmPFilter.toDate.value);
				if (! Ok)
				{
					alert("The refund to date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.toDate.focus();
					return;
				}
			}
			
			document.frmPFilter.action = 'rentalrefundpaymentexport.asp';
			document.frmPFilter.submit();
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

<form action="rentalrefundpayment.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>Refund Payment Report</b></font></td>
		</tr>
		<tr>
			<td>
				<fieldset>
					<legend><strong>Select</strong></legend>
					<!--BEGIN: FILTERS-->
					
						<table border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td><strong>Refund Date:</strong></td>
								<td>
									<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />&nbsp;
								</td>
								<td><% DrawTimeChoices "fromtime", from_time %></td>
								<td>
									<strong>To:</strong>
								</td>
								<td>
									<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />&nbsp;
								</td>
								<td><% DrawTimeChoices "totime", to_time %></td>
								<td><% DrawDateChoices "Date" %></td>
							</tr>
						</table>
					
					<p>
						<strong>Admin Location: </strong><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
						<strong>Admin: </strong><% ShowAdminUsers iAdminUserId %>&nbsp;&nbsp;
					</p>
					<p>
						<input class="button" type="button" value="View Report" onclick="validate();" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="exportReport()" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
<%
				' DISPLAY RESULTS
				'Display_Results sWhereClause
				Display_Rental_Refund_Report sWhereClause
%>
				<!-- END: DISPLAY RESULTS -->
      
			</td>
		 </tr>
	</table>
  </form>
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
	Dim sSql, oRs, oDisplay, iOldPaymentId, dVoucherTotal, dCardTotal, dMemoTotal, dGrandTotal, dSubTotal

	iOldPaymentId = CLng(0) 

	sSql = "SELECT paymentid, orgid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone, paymentdate, amount, isccrefund, priorbalance "
	sSql = sSql & " FROM egov_rentals_to_refund_method " & sWhereClause
	sSql = sSql & " ORDER BY paymentid" 
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If oRs.EOF then
		' EMPTY
		response.write "<p><strong>No refund payments found.</strong></p>"
	Else
		' Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "paymentid", adInteger, , adFldUpdatable
		oDisplay.fields.append "paymentdate", adVariant, 10, adFldUpdatable
		oDisplay.fields.append "userfname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "userlname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "voucheramt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "cardamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "subtotal", adCurrency, , adFldUpdatable
		oDisplay.fields.append "memoamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "total", adCurrency, , adFldUpdatable

		oDisplay.CursorLocation = 3
		'oDisplay.CursorType = 3

		oDisplay.open 

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("paymentid") = oRs("paymentid")
				oDisplay("paymentdate") = DateValue(oRs("paymentdate"))
				If Not IsNull(oRs("userfname")) Then 
					oDisplay("userfname") = oRs("userfname")
				End If 
				If Not IsNull(oRs("userlname")) Then 
					oDisplay("userlname") = oRs("userlname")
				End If 
				oDisplay("userhomephone") = oRs("userhomephone")
				oDisplay("voucheramt") = 0.00
				oDisplay("cardamt") = 0.00
				oDisplay("subtotal") = 0.00
				oDisplay("memoamt") = 0.00
				oDisplay("total") = 0.00
				iOldPaymentId = CLng(oRs("paymentid"))
			End If 
			If oRs("isccrefund") Then
				' Credit Card Refund
				oDisplay("cardamt") = oRs("amount")
				oDisplay("subtotal") = CDbl(oDisplay("subtotal")) + CDbl(oRs("amount"))
			Else 
				If IsNull(oRs("priorbalance")) Then
					' Voucher Issued
					oDisplay("voucheramt") = oRs("amount")
					oDisplay("subtotal") = CDbl(oDisplay("subtotal")) + CDbl(oRs("amount"))
				Else
					' Refund To Memo account
					oDisplay("memoamt") = oRs("amount")
				End If 
			End If 
			oDisplay("total") = CDbl(oDisplay("total")) + CDbl(oRs("amount"))

			oDisplay.Update
			oRs.MoveNext
		Loop

		' Show results
		oDisplay.MoveFirst
		
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" id=""rentalreceiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Payee</th>"
		response.write "<th>Voucher<br />Amount</th><th>Card Amt</th><th>Total Card<br />&amp; Voucher</th><th>Memo Amt</th><th>Total<br />Refund</th></tr>"
		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			response.write "<td align=""center""><a href=""viewpaymentreceipt.asp?paymentid=" & oDisplay("paymentid") & """>" & oDisplay("paymentid") & "</a></td>"
			response.write "<td align=""center"">" & oDisplay("paymentdate") & "</td>"
			response.write "<td align=""center"" valign=""top"">" & oDisplay("userfname") & " " & oDisplay("userlname") & "<br />" & FormatPhoneNumber(oDisplay("userhomephone")) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("voucheramt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cardamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("subtotal"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("memoamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("total"),2) & "</td>"
			dVoucherTotal = dVoucherTotal + CDbl(oDisplay("voucheramt"))
			dCardTotal = dCardTotal + CDbl(oDisplay("cardamt"))
			dSubTotal = dSubTotal + CDbl(oDisplay("subtotal"))
			dMemoTotal = dMemoTotal + CDbl(oDisplay("memoamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("total"))
			response.write "</tr>"
			oDisplay.MoveNext
		Loop 
		' Totals Row
		If bgcolor="#eeeeee" Then
			bgcolor="#ffffff" 
		Else
			bgcolor="#eeeeee"
		End If	
		response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""totalrow""><td colspan=""3"" align=""right"">Totals:</td>"
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

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
