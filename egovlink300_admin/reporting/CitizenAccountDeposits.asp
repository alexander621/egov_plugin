<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="reporting_functions.asp" //-->
<!-- #include file="citizen_reporting_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CitizenAccountDeposits.asp
' AUTHOR: SteveLoar
' CREATED: 01/22/2013
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
PageDisplayCheck "citizen account deposits report", sLevel	' In common.asp

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
					inlineMsg("toDate",'<strong>Invalid Value: </strong>The deposit "To" date should be in the format of MM/DD/YYYY.',8,"toDate");
					okToPost = false;
				}
			}

			// check from date
			if ($("#fromDate").val() != "")
			{
				if ( ! isValidDate( $("#fromDate").val() ) )
				{
					inlineMsg("fromDate",'<strong>Invalid Value: </strong>The deposit "From" date should be in the format of MM/DD/YYYY.',8,"fromDate");
					okToPost = false;
				}
			}

			if (_Action === 'export')
			{
				document.reportFilter.action = "CitizenAccountDepositsExport.asp";
			}
			else
			{
				document.reportFilter.action = "CitizenAccountDeposits.asp";
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

		<form action="CitizenAccountDeposits.asp" method="post" name="reportFilter">

		<font size="+1"><b>Citizen Account Deposits</b></font><br /><br />

		<fieldset>
			<legend><strong>Select</strong></legend>
		
			<!--BEGIN: FILTERS-->
			<p>
			<table border="0" cellpadding="0" cellspacing="0" id="daterangepicks">
				<tr>
					<td><label for="fromDate">Deposit Date:</label>
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
		' This is the new call in reporting/reporting_functions.asp
		Display_Citizen_Payment_Report sWhereClause

%>
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

	' make a holding recordset. THis allows us to put multiple payment types into one row for display, so keep this.
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

	sSql = "SELECT paymentid, orgid, userid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone, paymenttotal, paymentdate, journalentrytype, amount, "
	sSql = sSql & " paymenttypename, checkno, isothermethod, requirescash, requirescreditcard, requirescitizenaccount, "
	sSql = sSql & " requirescheckno, paymentlocationname, adminlocationid, adminuserid, item, [Transaction ID] "
	sSql = sSql & " FROM egov_citizen_account_to_payment_method " & sWhereClause
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

			oRs.MoveNext
		Loop
	Else
		bHasData = False 
	End If 
	
	oRs.Close
	Set oRs = Nothing

	If bHasData Then
		' header row
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"" id=""citizendeposits"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Account</th>"
		response.write "<th>Check Amt<br />Check #</th><th>Cash Amt</th><th>Card Amt</th><th>Total Chck<br />Cash, Card</th><th>Other Amt</th>"
		response.write "<th>Memo Amt</th><th>Total<br />Deposit</th></tr>"

		oDisplay.MoveFirst

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
		
		response.write vbcrlf & "<tr class=""totalrow"">"
		response.write "<td colspan=""3"" align=""right"">Totals:</td>"
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

End Sub 


%>
