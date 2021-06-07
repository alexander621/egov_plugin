<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="reporting_functions.asp" //-->
<!-- #include file="class_reporting_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: receipt_payment.asp
' AUTHOR: SteveLoar
' CREATED: 01/10/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This report has class refunds.
'
' MODIFICATION HISTORY
' 1.0   7/17/2007	Steve Loar - INITIAL VERSION
' 1.1	10/4/2007	Steve Loar - Adding payments to citizen accounts to the report
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iLocationId, iAdminUserId, iPaymentLocationId
Dim from_time, to_time, where_time

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "receipt payment rpt", sLevel	' In common.asp

' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
	toDate = today 
End If

If fromDate = "" or IsNull(fromDate) Then 
	'fromDate = cdate(Month(today)& "/1/" & Year(today)) 
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

If request("paymentlocationid") = "" Then
	iPaymentLocationId = 0
Else
	iPaymentLocationId = CLng(request("paymentlocationid"))
End If 

' BUILD SQL WHERE CLAUSE
varWhereClause = " WHERE orgid = " & session("orgid") 

'varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
If from_time = "none" Then 
	varWhereClause = varWhereClause & " AND paymentDate >= '" & fromDate & "' "
Else
	where_time = CDate( fromdate & " " & from_time )
	varWhereClause = varWhereClause & " AND paymentDate >= '" & where_time & "' "
End If 

If to_time = "none" Then 
	varWhereClause = varWhereClause & " AND paymentDate <= '" & DateAdd("d",1,toDate) & "' "
Else 
	where_time = CDate( todate & " " & to_time )
	varWhereClause = varWhereClause & " AND paymentDate <= '" & where_time & "' "
End If 

If iLocationId > 0 Then
	varWhereClause = varWhereClause & " AND adminlocationid = " & iLocationId
End If 
If iAdminUserId > 0 Then
	varWhereClause = varWhereClause & " AND adminuserid = " & iAdminUserId
End If 
If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		varWhereClause = varWhereClause & " AND paymentlocationid = 3 " 
	Else
		varWhereClause = varWhereClause & " AND paymentlocationid < 3 " 
	End If 
End If 


%>
<html lang="en">
<head>
	<meta charset="UTF-8">
  	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="reporting.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../classes/classes.css" />
	<link rel="stylesheet" href="pageprint.css" media="print" />
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
					alert("The payment from date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
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
					alert("The payment to date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.toDate.focus();
					return;
				}
			}

			document.frmPFilter.action = 'receipt_payment.asp';
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
					alert("The payment from date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
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
					alert("The payment to date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.frmPFilter.toDate.focus();
					return;
				}
			}
			
			document.frmPFilter.action = 'receipt_payment_export.asp';
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

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

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

	<form action="receipt_payment.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>Receipt Payment Report</b></font></td>
		</tr>
		<tr>
			<td>
				<fieldset>
					<legend><strong>Select</strong></legend>
				
					<p>
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td><strong>Payment Date: </strong></td>
							<td>
								<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />&nbsp;
								<!--<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>-->
							</td>
							<td><% DrawTimeChoices "fromtime", from_time %></td>
							<td>
								<strong>To:</strong>
							</td>
							<td>
								<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />&nbsp;
								<!--<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>-->
							</td>
							<td><% DrawTimeChoices "totime", to_time %></td>
							<td><% DrawDateChoices "Date" %></td>
						</tr>
					</table>
					</p>
					<p>
						<strong>Admin Location: </strong><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
						<strong>Admin: </strong><% ShowAdminUsers iAdminUserId %>&nbsp;&nbsp;
					</p>
					<p>
						<strong>Payment Location: </strong><% ShowPaymentLocations iPaymentLocationId %>
					</p>
					<!--END: DATE FILTERS-->
					<p>
						<input class="button" type="button" value="View Report" onclick="validate();" />
						<!--&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='receipt_payment_export.asp?fromDate=<%=fromDate%>&toDate=<%=toDate%>&locationid=<%=iLocationId%>&adminuserid=<%=iAdminUserId%>&paymentlocationid=<%=iPaymentLocationId%>'" />-->
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="exportReport()" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<%
				
				' DISPLAY RESULTS
				Display_Class_Payment_Report varWhereClause
				
				%>
      
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
' void Display_Results sWhereClause 
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sWhereClause )
	Dim sSql, oRs, oDisplay, iOldPaymentId, dCashTotal, dCheckTotal, dCardtotal, dOtherTotal, dMemoTotal
	Dim dGrandTotal, dCCCTotal, dCCCSubTotal, bHasData

	iOldPaymentId = CLng(0) 
	dCCCTotal = CDbl(0.0)
	bHasData = False 

	' make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "paymentid", adInteger, , adFldUpdatable
	oDisplay.fields.append "paymentdate", adVariant, 10, adFldUpdatable
	oDisplay.fields.append "item", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userid", adInteger, , adFldUpdatable
	oDisplay.fields.append "userfname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userlname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "checkamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "checkno", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "cashamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "cardamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "cccsubtotal", adCurrency, , adFldUpdatable
	oDisplay.fields.append "otheramt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "memoamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "paymenttotal", adCurrency, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3
	oDisplay.open 

	' Pull Class Purchases
	If AddToDisplay( "egov_class_to_payment_method", sWhereClause, oDisplay ) Then 
		bHasData = True
	End If 


	' Pull Citizen Account Deposits - now in CitizenAccountDeposits.asp
	'If AddToDisplay(  "egov_citizen_account_to_payment_method", sWhereClause, oDisplay  ) Then 
	'	bHasData = True
	'End If 

	If bHasData Then 
		' Sort the data by paymentid
		oDisplay.sort = "paymentid"
		' Show results
		oDisplay.MoveFirst

		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Payee</th>"
		response.write "<th>Check Amt<br />Check #</th><th>Cash Amt</th><th>Card Amt</th><th>Total Chck<br />Cash, Card</th><th>Other Amt</th>"
		response.write "<th>Memo Amt</th><th>Total<br />Paid</th></tr>"
		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			If oDisplay("item") = "Citizen Acct" Then
				response.write "<td align=""center""><a href=""../purchases/viewjournal.asp?uid=" & oDisplay("userid") & "&pid=" & oDisplay("paymentid") & "&rt=c&it=ci&jet=d"">" & oDisplay("paymentid") & "</a></td>"
			Else 
				response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("paymentid") & """>" & oDisplay("paymentid") & "</a></td>"
			End If 
			response.write "<td align=""center"">" & oDisplay("paymentdate") & "</td>"
			response.write "<td align=""center"" valign=""top"">" & oDisplay("userfname") & " " & oDisplay("userlname") & "<br />" & FormatPhoneNumber(oDisplay("userhomephone")) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("checkamt"), 2) & "<br />" & oDisplay("checkno") & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cashamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cardamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cccsubtotal"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("otheramt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("memoamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("paymenttotal"),2) & "</td>"
			dCheckTotal = dCheckTotal + CDbl(oDisplay("checkamt"))
			dCashTotal = dCashTotal + CDbl(oDisplay("cashamt"))
			dCardTotal = dCardTotal + CDbl(oDisplay("cardamt"))
			dOtherTotal = dOtherTotal + CDbl(oDisplay("otheramt"))
			dMemoTotal = dMemoTotal + CDbl(oDisplay("memoamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("paymenttotal"))
			dCCCTotal = dCCCTotal + CDbl(oDisplay("cccsubtotal"))
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
		response.write "<td align=""right"">" & FormatNumber(dCheckTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCashTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCardTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCCCTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dOtherTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dMemoTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"

	End If 

	oDisplay.Close
	Set oDisplay = Nothing 
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' boolean AddToDisplay( sFrom, sWhereClause, oDisplay )
'------------------------------------------------------------------------------------------------------------
Function AddToDisplay( ByVal sFrom, ByVal sWhereClause, ByRef oDisplay )
	Dim oRs, bHasData, sSql

	sSql = "SELECT paymentid, orgid, userid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone, paymenttotal, paymentdate, journalentrytype, amount, "
	sSql = sSql & " paymenttypename, checkno, isothermethod, requirescash, requirescreditcard, requirescitizenaccount, "
	sSql = sSql & " requirescheckno, paymentlocationname, adminlocationid, adminuserid, item, [Transaction ID] "
	sSql = sSql & " FROM " & sFrom & " " & sWhereClause
	sSql = sSql & " ORDER BY paymentid" 
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("paymentid")) <> CLng(iOldPaymentId) Then
				oDisplay.addnew 
				oDisplay("paymentid") = oRs("paymentid")
				oDisplay("paymentdate") = DateValue(oRs("paymentdate"))
				oDisplay("item") = oRs("item")
				oDisplay("userid") = oRs("userid")
				oDisplay("userfname") = oRs("userfname")
				oDisplay("userlname") = oRs("userlname")
				oDisplay("userhomephone") = oRs("userhomephone")
				oDisplay("paymenttotal") = oRs("paymenttotal")
				oDisplay("checkamt") = 0.00
				oDisplay("cashamt") = 0.00
				oDisplay("cardamt") = 0.00
				oDisplay("cccsubtotal") = 0.00
				oDisplay("otheramt") = 0.00
				oDisplay("memoamt") = 0.00
				dCCCSubTotal = 0.00
				iOldPaymentId = CLng(oRs("paymentid"))
			End If 
			If oRs("requirescheckno") Then
				oDisplay("checkamt") = oRs("amount")
				oDisplay("checkno") = oRs("checkno")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRs("amount"))
			End If 
			If oRs("requirescash") Then
				oDisplay("cashamt") = oRs("amount")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRs("amount"))
			End If 
			If oRs("requirescreditcard") Then
				oDisplay("cardamt") = oRs("amount")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRs("amount"))
			End If 
			If oRs("isothermethod") Then
				oDisplay("otheramt") = oRs("amount")
			End If 
			If oRs("requirescitizenaccount") Then
				oDisplay("memoamt") = oRs("amount")
			End If 
			oDisplay("cccsubtotal") = dCCCSubTotal

			oDisplay.Update
			oRs.MoveNext
		Loop
	Else
		bHasData = False
	End If 
	
	oRs.Close
	Set oRs = Nothing

	AddToDisplay = bHasData

End Function 


%>
