<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: account_distribution.asp
' AUTHOR: Steve Loar
' CREATED: 07/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   7/19/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iLocationId, iAdminUserId, iPaymentLocationId, iReportType, sRptTitle, sRptType

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "prior account distribution rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

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

If request("reporttype") = "" Then 
	iReportType = CLng(1)
Else
	iReportType = CLng(request("reporttype"))
End If 

If iReportType = CLng(1) Then
	sRptTitle = "Summary"
	sRptType = "Summary"
Else
	sRptTitle = "Detail"
	sRptType = "Detail"
End If 

If request("journalentrytypeid") = "" Then
	iJournalEntryTypeId = 0
Else
	iJournalEntryTypeId = CLng(request("journalentrytypeid"))
End If 

' BUILD SQL WHERE CLAUSE
varWhereClause = " AND (P.paymentDate >= '" & fromDate & "' AND P.paymentDate <= '" & DateAdd("d",1,toDate) & "') "
varWhereClause = varWhereClause & " AND A.orgid = " & session("orgid") 
If iLocationId > 0 Then
	varWhereClause = varWhereClause & " AND P.adminlocationid = " & iLocationId
End If 

If iAdminUserId > 0 Then
	varWhereClause = varWhereClause & " AND adminuserid = " & iAdminUserId
End If 

If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		varWhereClause = varWhereClause & " AND P.paymentlocationid = 3 " 
	Else
		varWhereClause = varWhereClause & " AND P.paymentlocationid < 3 " 
	End If 
End If 

If iJournalEntryTypeId > 0 Then 
	varWhereClause = varWhereClause & " AND P.journalentrytypeid = " & iJournalEntryTypeId
End If 

%>

<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		window.onload = function()
		{
		  //factory.printing.header = "Printed on &d"
		  factory.printing.footer = "&bPrinted on &d - Page:&p/&P";
		  factory.printing.portrait = true;
		  factory.printing.leftMargin = 0.5;
		  factory.printing.topMargin = 0.5;
		  factory.printing.rightMargin = 0.5;
		  factory.printing.bottomMargin = 0.5;
		 
		  // enable control buttons
		  var templateSupported = factory.printing.IsTemplateSupported();
		  var controls = idControls.all.tags("input");
		  for ( i = 0; i < controls.length; i++ ) 
		  {
			controls[i].disabled = false;
			if ( templateSupported && controls[i].className == "ie55" )
			  controls[i].style.display = "inline";
		  }
		}

	  //-->
	</script>

	<script language="Javascript" src="scripts/dates.js"></script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
	<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
</div>

<object id="factory" viewastext  style="display:none"
  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
   codebase="../includes/smsx.cab#Version=6,3,434,12">
</object>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<form action="account_distribution.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Account Distribution <%=sRptTitle%></strong></font></td>
		</tr>
		<tr>
			<td>
				<fieldset>
					<legend><strong>Select</strong></legend>
				
					<!--BEGIN: FILTERS-->
					<!--BEGIN: DATE FILTERS-->
					<p>
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td><strong>Payment Date: </strong></td>
							<td>
								<input type="text" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('From');"><img src="../images/calendar.gif" border="0" /></a>		 
							</td>
							<td>&nbsp;</td>
							<td>
								<strong>To:</strong>
							</td>
							<td>
								<input type="text" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('To');"><img src="../images/calendar.gif" border="0" /></a>
							</td>
							<td>&nbsp;</td>
							<td><%DrawDateChoices "Dates" %></td>
						</tr>
					</table>
					</p>
					<p>
						<strong>Admin Location: </strong><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
						<strong>Admin: </strong><% ShowAdminUsers iAdminUserId %>
					</p>
					<p>
						<strong>Payment Location: </strong><% ShowPaymentLocations iPaymentLocationId %>&nbsp;&nbsp;
						<strong>Report Type: </strong><% ShowReportTypes iReportType %>&nbsp;&nbsp;
						<strong>Entries: </strong><% ShowJournalEntryTypes iJournalEntryTypeId %>
					</p>
					<!--END: DATE FILTERS-->
					<p>
						<input class="button" type="submit" value="View Report" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='account_distribution_export.asp?fromDate=<%=fromDate%>&toDate=<%=toDate%>&locationid=<%=iLocationId%>&adminuserid=<%=iAdminUserId%>&paymentlocationid=<%=iPaymentLocationId%>&reporttype=<%=iReportType%>&journalentrytypeid=<%=iJournalEntryTypeId%>'" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
				<%
				
				If sRptType = "Detail" Then
					DisplayDetails varWhereClause
				Else
					DisplaySummary varWhereClause
				End If 
				
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
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' FUNCTION DRAWDATECHOICES(SNAME)
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( sName )

	response.write vbcrlf & "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub DisplaySummary( sWhereClause )
'------------------------------------------------------------------------------------------------------------
Sub DisplaySummary( sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal

	iOldAccountId = CLng(0) 
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)


	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, sum(L.amount) as amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P "
	sSql = sSql & " WHERE A.accountid = L.accountid and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	'sSql = sSql & " WHERE A.accountstatus = 'A' and A.accountid = L.accountid and ispaymentaccount = 0 "
	sSql = sSql & " and L.paymentid = P.paymentid and L.amount <> 0.00 " & sWhereClause 
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, A.accountid, L.entrytype ORDER BY A.accountid, L.entrytype"
'	response.write sSql
'	response.end

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	If oRequests.EOF then
		' EMPTY
		response.write "<p>No account activity found.</p>"
	Else
		' Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
		oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
		oDisplay.fields.append "creditamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "debitamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "totalamt", adCurrency, , adFldUpdatable

		oDisplay.CursorLocation = 3
		'oDisplay.CursorType = 3

		oDisplay.open 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId Then
				If oRequests("entrytype") = "credit" Then
					oDisplay("creditamt") = oRequests("amount")
					dTotal = dTotal + CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					oDisplay("totalamt") = dTotal 
				End If 
				If oRequests("entrytype") = "debit" Then
					oDisplay("debitamt") = -CDbl(oRequests("amount"))
					dTotal = dTotal - CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					oDisplay("totalamt") = dTotal 
				End If 
			Else
				oDisplay.addnew 
				oDisplay("accountid") = oRequests("accountid")
				oDisplay("accountname") = oRequests("accountname")
				oDisplay("accountnumber") = oRequests("accountnumber")
				If sRptType = "Detail" Then
					oDisplay("paymentid") = oRequests("paymentid")
				End If 
				oDisplay("creditamt") = 0.00
				oDisplay("debitamt") = 0.00
				oDisplay("totalamt") = 0.00
				If oRequests("entrytype") = "credit" Then
					oDisplay("creditamt") = CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					dTotal = CDbl(oRequests("amount"))
					oDisplay("totalamt") = CDbl(oRequests("amount"))
				End If 
				If oRequests("entrytype") = "debit" Then
					oDisplay("debitamt") = -CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					dTotal = -CDbl(oRequests("amount"))
					oDisplay("totalamt") = -CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
			End If 
			oDisplay.Update
			oRequests.MoveNext
		Loop
		'sort the Display recordset
		oDisplay.Sort = "accountname ASC, accountnumber ASC "

		' Show results
		oDisplay.MoveFirst
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">"
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Account Name</th><th>Account Number</th>"
		response.write "<th>Total Amt<br />Credited</th><th>Total Amt<br />Debited</th><th>Total Amt<br />Transfered</th></tr>"

		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			response.write "<td align=""left"">" & oDisplay("accountname") & "</td>"
			response.write "<td align=""center"">" & oDisplay("accountnumber") & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("creditamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("debitamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("totalamt"), 2) & "</td>"
			dGrandTotal = dGrandTotal + CDbl(oDisplay("totalamt"))
			response.write "</tr>"
			oDisplay.MoveNext
		Loop 
		' Totals Row
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""2"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalCredit, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalDebit, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

		oDisplay.Close
		Set oDisplay = Nothing 

	End If 

	oRequests.Close
	Set oRequests = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub DisplayDetails( sWhereClause )
'------------------------------------------------------------------------------------------------------------
Sub DisplayDetails( sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldAccountId, iOldPaymentId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal

	iOldAccountId = CLng(0) 
	iOldPaymentId = CLng(0)
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)


	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P "
	sSql = sSql & " WHERE A.accountid = L.accountid and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	sSql = sSql & " and L.paymentid = P.paymentid and L.amount <> 0.00 " & sWhereClause 
	sSql = sSql & " ORDER BY A.accountid, P.paymentid, L.entrytype"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	If oRequests.EOF then
		' EMPTY
		response.write "<p>No account activity found.</p>"
	Else
		' Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
		oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
		oDisplay.fields.append "receiptno", adInteger, , adFldUpdatable
		oDisplay.fields.append "paymentdate", adDBTimeStamp, , adFldUpdatable
		oDisplay.fields.append "creditamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "debitamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "totalamt", adCurrency, , adFldUpdatable

		oDisplay.CursorLocation = 3
		'oDisplay.CursorType = 3

		oDisplay.open 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId And CLng(oRequests("paymentid")) = iOldPaymentId Then
				If oRequests("entrytype") = "credit" Then
					oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRequests("amount"))
					dTotal = dTotal + CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					oDisplay("totalamt") = dTotal 
				End If 
				If oRequests("entrytype") = "debit" Then
					oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRequests("amount"))
					dTotal = dTotal - CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					oDisplay("totalamt") = dTotal 
				End If 
			Else
				oDisplay.addnew 
				oDisplay("accountid") = oRequests("accountid")
				oDisplay("accountname") = oRequests("accountname")
				oDisplay("accountnumber") = oRequests("accountnumber")
				oDisplay("receiptno") = oRequests("paymentid")
				oDisplay("paymentdate") = oRequests("paymentdate")
				oDisplay("creditamt") = 0.00
				oDisplay("debitamt") = 0.00
				oDisplay("totalamt") = 0.00
				If oRequests("entrytype") = "credit" Then
					oDisplay("creditamt") = CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					dTotal = CDbl(oRequests("amount"))
					oDisplay("totalamt") = CDbl(oRequests("amount"))
				End If 
				If oRequests("entrytype") = "debit" Then
					oDisplay("debitamt") = -CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					dTotal = -CDbl(oRequests("amount"))
					oDisplay("totalamt") = -CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			oDisplay.Update
			oRequests.MoveNext
		Loop
		'sort the Display recordset
		oDisplay.Sort = "accountname ASC, accountnumber ASC, receiptno ASC"

		' Show results
		oDisplay.MoveFirst
		response.write vbcrlf & "<div class=""receiptpaymentshadow"">"
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Account Name</th><th>Account Number</th><th>Receipt No.</th><th>Date</th>"
		response.write "<th>Total Amt<br />Credited</th><th>Total Amt<br />Debited</th><th>Total Amt<br />Transfered</th></tr>"

		bgcolor = "#eeeeee"
		iOldAccountId = CLng(0)
		dCreditSubTotal = CDbl(0.00)
		dDebitSubTotal = CDbl(0.00)
		dSubTotal = CDbl(0.00)
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			If iOldAccountId <> CLng(oDisplay("accountid")) Then 
				' Put out a sub total row
				If iOldAccountId <> CLng(0) Then
					response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"" align=""right"">Sub-Total:</td>"
					response.write "<td align=""right"">" & FormatNumber(dCreditSubTotal, 2) & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dDebitSubTotal, 2) & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dSubTotal,2) & "</td>"
					response.write "</tr>"
				End If 
				response.write "<td align=""left"">" & oDisplay("accountname") & "</td>"
				response.write "<td align=""center"">" & oDisplay("accountnumber") & "</td>"
				iOldAccountId = CLng(oDisplay("accountid"))
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal = CDbl(0.00)
				dSubTotal = CDbl(0.00)
			Else
				' Need place holders 
				response.write "<td>&nbsp;</td><td>&nbsp;</td>"
			End If 
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("receiptno") & """>" & oDisplay("receiptno") & "</a></td>"
			response.write "<td align=""right"">" & FormatDateTime(oDisplay("paymentdate"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("creditamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("debitamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("totalamt"), 2) & "</td>"
			dCreditSubTotal = dCreditSubTotal + CDbl(oDisplay("creditamt"))
			dDebitSubTotal = dDebitSubTotal + CDbl(oDisplay("debitamt"))
			dSubTotal = dSubTotal + CDbl(oDisplay("totalamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("totalamt"))
			response.write "</tr>"
			oDisplay.MoveNext
		Loop 
		' Put out a sub total row
		If iOldAccountId <> CLng(0) Then
			response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"" align=""right"">Sub-Total:</td>"
			response.write "<td align=""right"">" & FormatNumber(dCreditSubTotal, 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(dDebitSubTotal, 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(dSubTotal,2) & "</td>"
			response.write "</tr>"
		End If 
		' Totals Row
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalCredit, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalDebit, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"

		oDisplay.Close
		Set oDisplay = Nothing 

	End If 

	oRequests.Close
	Set oRequests = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowAdminLocations( iLocationId )
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminLocations( iLocationId )
	Dim sSql, oLocation
	
	sSql = "Select locationid, name from egov_class_location where orgid = " & session("orgid") & " order by name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open  sSQL, Application("DSN"), 3, 1

	If Not oLocation.EOF Then 
		response.write vbcrlf & "<select name=""locationid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iLocationId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All Locations</option>"
		Do While Not oLocation.EOF 
			response.write vbcrlf & "<option value=""" & oLocation("locationid") & """"
			If CLng(oLocation("locationid")) = CLng(iLocationId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oLocation("name") & "</option>"
			oLocation.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oLocation.close
	Set oLocation = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowPaymentLocations( iPaymentLocationId )
'------------------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations( iPaymentLocationId )

	response.write vbcrlf & "<select name=""paymentlocationid"">"
	response.write vbcrlf & "<option value=""0"" "
	If CLng(0) = CLng(iPaymentLocationId) Then ' none selected
		 response.write " selected=""selected"" "
	End If 
	response.write ">Web Site and Office</option>"

	response.write vbcrlf & "<option value=""1"""
	If CLng(1) = CLng(iPaymentLocationId) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Office Only</option>"

	response.write vbcrlf & "<option value=""2"""
	If CLng(2) = CLng(iPaymentLocationId) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Web Site Only</option>"

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowReportTypes( iReportType )
'------------------------------------------------------------------------------------------------------------
Sub ShowReportTypes( iReportType )
	
	response.write vbcrlf & "<select name=""reporttype"">"

	response.write vbcrlf & "<option value=""1"""
	If CLng(1) = CLng(iReportType) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Summary</option>"

	response.write vbcrlf & "<option value=""2"""
	If CLng(2) = CLng(iReportType) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Detail</option>"

	response.write vbcrlf & "</select>"
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowAdminUsers( iAdminUserId )
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminUsers( iAdminUserId )
	Dim sSql, oUser
	
	sSql = "Select userid, firstname, lastname from users where isrootadmin = 0 and orgid = " & session("orgid") & " order by lastname, firstname"

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open  sSQL, Application("DSN"), 3, 1

	If Not oUser.EOF Then 
		response.write vbcrlf & "<select name=""adminuserid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iAdminUserId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All</option>"
		Do While Not oUser.EOF 
			response.write vbcrlf & "<option value=""" & oUser("userid") & """"
			If CLng(oUser("userid")) = CLng(iAdminUserId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oUser("firstname") & " " & oUser("lastname") & "</option>"
			oUser.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oUser.close
	Set oUser = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowJournalEntryTypes( iJournalEntryTypeId )
'------------------------------------------------------------------------------------------------------------
Sub ShowJournalEntryTypes( iJournalEntryTypeId )
	Dim sSql, oTypes
	
	sSql = "SELECT journalentrytypeid, displayname + ' Only' as displayname FROM egov_journal_entry_types WHERE journalentrytype = 'refund' "
	sSql = sSql & " OR journalentrytype = 'purchase' ORDER BY displayorder"

	Set oTypes = Server.CreateObject("ADODB.Recordset")
	oTypes.Open  sSQL, Application("DSN"), 3, 1

	If Not oTypes.EOF Then 
		response.write vbcrlf & "<select name=""journalentrytypeid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(0) = CLng(iJournalEntryTypeId) Then ' none selected
			 response.write " selected=""selected"" "
		End If 
		response.write ">Show All</option>"
		Do While Not oTypes.EOF 
			response.write vbcrlf & "<option value=""" & oTypes("journalentrytypeid") & """"
			If CLng(oTypes("journalentrytypeid")) = CLng(iJournalEntryTypeId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oTypes("displayname") & "</option>"
			oTypes.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oTypes.close
	Set oTypes = Nothing 

End Sub 
%>