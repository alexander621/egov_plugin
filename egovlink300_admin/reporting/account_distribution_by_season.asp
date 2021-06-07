<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: account_distribution_by_season.asp
' AUTHOR: David Boyer
' CREATED: 07/08/2008
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   07/08/2008		David Boyer - INITIAL VERSION (copied and modified off of account_distribution_by_season.asp)
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iLocationId, iAdminUserId, iPaymentLocationId, iReportType, sRptTitle, sRptType, iAccountNo, bOrgHasAccounts

'INITIALIZE AND DECLARE VARIABLES
'SPECIFY FOLDER LEVEL
 sLevel = "../" ' Override of value from common.asp

'USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "account distribution by season rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

'PROCESS REPORT FILTER VALUES
'PROCESS DATE VALUES
 fromDate = Request("fromDate")
 toDate   = Request("toDate")
 today    = Date()

'IF EMPTY DEFAULT TO CURRENT TO DATE
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

If request("accountid") = "" Then
	iAccountNo = 0
Else
	iAccountNo = CLng(request("accountid"))
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

if request("classseasonid") <> "" then
   iClassSeasonID = request("classseasonid")
else
   iClassSeasonID = ""
end if

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

If CLng(iAccountNo) > CLng(0) Then
	varWhereClause = varWhereClause & " AND A.accountid = " & iAccountNo & " "
End If 

'Determine which season has been selected
 If iClassSeasonID <> "" Then 
    varWhereClause = varWhereClause & " AND C.classseasonid = " & iClassSeasonID
 End If 

' Not every org has general ledger accounts so we need to be able to hide/show accordingly.
bOrgHasAccounts = OrgHasFeature("gl accounts")

%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="JavaScript" src="../scripts/jquery-1.7.2.min.js"></script>
	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	<!--

		function doCalendar(ToFrom) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			//eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
			eval('window.open("calendarpicker.asp?updatefield=' + ToFrom + '&date=' + $("#" + ToFrom ).val() + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function viewReport()
		{
			document.frmPFilter.action = 'account_distribution_by_season.asp';
			document.frmPFilter.submit();
		}

		function exportReport()
		{
			document.frmPFilter.action = 'account_distribution_by_season_export.asp';
			document.frmPFilter.submit();
		}

	//-->
	</script>

	<script language="Javascript" src="scripts/dates.js"></script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<form action="account_distribution_by_season.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Account Distribution (By Season) <%=sRptTitle%></strong></font></td>
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
								<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>		 
							</td>
							<td>&nbsp;</td>
							<td>
								<strong>To:</strong>
							</td>
							<td>
								<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
							</td>
							<td>&nbsp;</td>
							<td><%DrawDateChoices "Dates" %></td>
						</tr>
					</table>
					</p>
					<!--END: DATE FILTERS-->
					<p>
						<strong>Admin Location: </strong><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
						<strong>Admin: </strong><% ShowAdminUsers iAdminUserId %>
					</p>
					<p>
						<strong>Payment Location: </strong><% ShowPaymentLocations iPaymentLocationId %>&nbsp;&nbsp;
						<strong>Report Type: </strong><% ShowReportTypes iReportType %>&nbsp;&nbsp;
						<strong>Entries: </strong><% ShowJournalEntryTypes iJournalEntryTypeId %>
					</p>
					<p>
						<strong>Season: </strong><% ShowClassSeasons iClassSeasonID %>
					</p>

<%					If bOrgHasAccounts Then		%>
						<p>
							<strong>GL Account: </strong>
<%
							ShowAccountPicks "accountid", iAccountNo, True 
%>
						</p>
<%					Else
						response.write "<input type=""hidden"" id=""accountid"" name=""accountid"" value=""0"" />"
					End If 
%>
					
					<p>
						<input class="button" type="button" value="View Report" onclick="viewReport()" />
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
'------------------------------------------------------------------------------------------------------------
' FUNCTION DRAWDATECHOICES(SNAME)
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=""calendarinput"" name=""" & sName & """>" & vbcrlf
	response.write "  <option value=""0"">Or Select Date Range from Dropdown...</option>" & vbcrlf
	response.write "  <option value=""11"">This Week</option>"   & vbcrlf
	response.write "  <option value=""12"">Last Week</option>"   & vbcrlf
	response.write "  <option value=""1"">This Month</option>"   & vbcrlf
	response.write "  <option value=""2"">Last Month</option>"   & vbcrlf
	response.write "  <option value=""3"">This Quarter</option>" & vbcrlf
	response.write "  <option value=""4"">Last Quarter</option>" & vbcrlf
	response.write "  <option value=""6"">Year to Date</option>" & vbcrlf
	response.write "  <option value=""5"">Last Year</option>"    & vbcrlf
	response.write "  <option value=""7"">All Dates to Date</option>" & vbcrlf
	response.write "</select>" & vbcrlf

End Sub

'------------------------------------------------------------------------------------------------------------
' Sub DisplaySummary( sWhereClause )
'------------------------------------------------------------------------------------------------------------
Sub DisplaySummary( ByVal sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal, bHasData

	iOldAccountId = CLng(0) 
	dTotal        = CDbl(0.00)
	dTotalCredit  = CDbl(0.00)
	dTotalDebit   = CDbl(0.00)
	dGrandTotal   = CDbl(0.00)
	bHasData      = False 

	' Got some data now make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "ClassSeasonID", adInteger, , adFldUpdatable
	oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
	oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "creditamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "debitamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "totalamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oDisplay.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3

	oDisplay.open 

	' Pull add account data except the citizen accounts
	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, 0 AS iscitizenaccount, sum(L.amount) as amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE A.accountid = L.accountid "
	sSql = sSql & " AND L.paymentid = P.paymentid "
	sSql = sSql & " AND L.amount <> 0.00 "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " GROUP BY C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount "
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, L.entrytype"
	'	response.write sSql & "<br />"
	'	response.end

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) <> iOldAccountId Then
				oDisplay.addnew 
				oDisplay("ClassSeasonID")    = oRequests("ClassSeasonID")
				oDisplay("accountid")        = oRequests("accountid")
				oDisplay("accountname")      = oRequests("accountname") 
				oDisplay("accountnumber")    = oRequests("accountnumber")
				oDisplay("ispaymentaccount") = oRequests("ispaymentaccount")
				oDisplay("iscitizenaccount") = oRequests("iscitizenaccount") 
				If sRptType = "Detail" Then
					oDisplay("paymentid") = oRequests("paymentid")
				End If 
				oDisplay("creditamt") = 0.00
				oDisplay("debitamt")  = 0.00
				oDisplay("totalamt")  = 0.00
				iOldAccountId = CLng(oRequests("accountid"))
			End If 
			If oRequests("entrytype") = "credit" Then
				oDisplay("creditamt") = CDbl(oRequests("amount"))
				'dTotal = CDbl(oRequests("amount"))
				oDisplay("totalamt") = CDbl(oDisplay("totalamt")) + CDbl(oRequests("amount"))
			End If 
			If oRequests("entrytype") = "debit" Then
				oDisplay("debitamt") = -CDbl(oRequests("amount"))
				'dTotal = -CDbl(oRequests("amount"))
				oDisplay("totalamt") = CDbl(oDisplay("totalamt")) - CDbl(oRequests("amount"))
			End If 
				
			oDisplay.Update
			oRequests.MoveNext
		Loop
	End If 

	oRequests.Close
	Set oRequests = Nothing 

	' Get the citizen accounts summary here
	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, "
	sSql = sSql & " 1 AS iscitizenaccount, sum(L.amount) as amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_organizations_to_paymenttypes OP, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE L.paymentid = P.paymentid "
	sSql = sSql & " AND L.paymenttypeid = 4 "
	sSql = sSql & " AND A.accountid = OP.accountid "
	sSql = sSql & " AND OP.paymenttypeid = L.paymenttypeid "
	sSql = sSql & " AND OP.orgid = P.orgid "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " GROUP BY C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount "
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, L.entrytype"
'	response.write sSql & "<br />"
'	response.end

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) <> iOldAccountId Then
				oDisplay.addnew 
				oDisplay("ClassSeasonID") = oRequests("ClassSeasonID")
				oDisplay("accountid")     = oRequests("accountid")
				oDisplay("accountname")   = oRequests("accountname") 
				oDisplay("accountnumber") = oRequests("accountnumber")
				'oDisplay("ispaymentaccount") = oRequests("ispaymentaccount")
				'oDisplay("iscitizenaccount") = oRequests("iscitizenaccount")
				oDisplay("ispaymentaccount") = True 
				oDisplay("iscitizenaccount") = True 
				If sRptType = "Detail" Then
  					oDisplay("paymentid") = oRequests("paymentid")
				End If 
				oDisplay("creditamt") = 0.00
				oDisplay("debitamt")  = 0.00
				oDisplay("totalamt")  = 0.00
				iOldAccountId = CLng(oRequests("accountid"))
			End If 
			If oRequests("entrytype") = "credit" Then
  				oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRequests("amount"))
		  		oDisplay("totalamt") = CDbl(oDisplay("totalamt")) + CDbl(oRequests("amount"))
			End If 
			If oRequests("entrytype") = "debit" Then
			  	oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRequests("amount"))
		  		oDisplay("totalamt") = CDbl(oDisplay("totalamt")) - CDbl(oRequests("amount"))
			End If

			oDisplay.Update
			oRequests.MoveNext
		Loop
	End If 

	oRequests.Close
	Set oRequests = Nothing 
		
	If bHasData Then 
		'sort the Display recordset
		oDisplay.Sort = "ClassSeasonID, ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC "

		' Show results
		oDisplay.MoveFirst
		'response.write "<div class=""receiptpaymentshadow"">" & vbcrlf
		response.write "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">" & vbcrlf
		response.write "  <tr class=""tablelist"">"              & vbcrlf
		response.write "      <th>Season</th>"                   & vbcrlf
		response.write "      <th>Account Name</th>"             & vbcrlf
		response.write "      <th>Account Number</th>"           & vbcrlf
		response.write "      <th>Total Amt<br />Credited</th>"  & vbcrlf
		response.write "      <th>Total Amt<br />Debited</th>"   & vbcrlf
		response.write "  </tr>" & vbcrlf

		bgcolor = "#eeeeee"
		lcl_prev_ClassSeasonID = ""

		Do While Not oDisplay.EOF
			bgcolor = changeBGcolor(bgcolor,"#eeeeee","#ffffff")

			response.write "  <tr bgcolor=""" &  bgcolor  & """>" & vbcrlf

			'Determine if we show the Season Name or not based on the previous row value.
			if lcl_prev_ClassSeasonID <> oDisplay("ClassSeasonID") then
			response.write "<td align=""left"">"   & getSeasonName(oDisplay("ClassSeasonID")) & "</td>" & vbcrlf
			else
			response.write "<td>&nbsp;</td>" & vbcrlf
			end if

			response.write "<td align=""left"">"   & oDisplay("accountname")                  & "</td>" & vbcrlf
			response.write "<td align=""center"">" & oDisplay("accountnumber")                & "</td>" & vbcrlf
			response.write "<td align=""right"">"  & FormatNumber(oDisplay("creditamt"), 2)   & "</td>" & vbcrlf
			response.write "<td align=""right"">"  & FormatNumber(oDisplay("debitamt"), 2)    & "</td>" & vbcrlf

			dTotalCredit = dTotalCredit + CDbl(oDisplay("creditamt"))
			dTotalDebit  = dTotalDebit + CDbl(oDisplay("debitamt"))
			dGrandTotal  = dGrandTotal + dTotalCredit - dTotalDebit

			response.write "  </tr>" & vbcrlf

			'Set the value of the current ClassSeasonID so that we can compare it to the next value to determine if it is displayed or not
			lcl_prev_ClassSeasonID = oDisplay("ClassSeasonID")

			oDisplay.MoveNext
		Loop
		' Totals Row
		response.write "  <tr class=""totalrow"">" & vbcrlf
		response.write "      <td colspan=""3"" align=""right"">Totals:</td>" & vbcrlf
		response.write "      <td align=""right"">" & FormatNumber(dTotalCredit, 2) & "</td>" & vbcrlf
		response.write "      <td align=""right"">" & FormatNumber(dTotalDebit, 2)  & "</td>" & vbcrlf
'		response.write "      <td align=""right"">" & FormatNumber((dTotalCredit + dTotalDebit),2) & "</td>" & vbcrlf
		response.write "  </tr>" & vbcrlf

		response.write "</table>" & vbcrlf
		'response.write "</div>"   & vbcrlf

	Else
		response.write "<p>No information could be found for the criteria selected.</p>"
	End If 

	oDisplay.Close
	Set oDisplay = Nothing 

End Sub 

'------------------------------------------------------------------------------------------------------------
' Sub DisplayDetails( sWhereClause )
'------------------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldAccountId, iOldPaymentId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal, bHasData

	iOldAccountId = CLng(0) 
	iOldPaymentId = CLng(0)
	dTotal        = CDbl(0.00)
	dTotalCredit  = CDbl(0.00)
	dTotalDebit   = CDbl(0.00)
	dGrandTotal   = CDbl(0.00)
	bHasData      = False 

	'Got some data now make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "ClassSeasonID", adInteger, , adFldUpdatable
	oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
	oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "receiptno", adInteger, , adFldUpdatable
	oDisplay.fields.append "paymentdate", adDBTimeStamp, , adFldUpdatable
	oDisplay.fields.append "paymenttypeid", adInteger, , adFldUpdatable
	oDisplay.fields.append "journalentrytypeid", adInteger, , adFldUpdatable
	oDisplay.fields.append "userid", adInteger, , adFldUpdatable
	oDisplay.fields.append "creditamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "debitamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "totalamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oDisplay.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3

	oDisplay.open 

	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, "
	sSql = sSql & " ISNULL(L.paymenttypeid,0) AS paymenttypeid, P.userid, P.journalentrytypeid, L.ispaymentaccount, 0 AS iscitizenaccount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE A.accountid = L.accountid "
	'sSql = sSql & " WHERE A.accountid = L.accountid and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	sSql = sSql & " AND L.paymentid = P.paymentid "
	sSql = sSql & " AND L.amount <> 0.00 "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, P.paymentid, L.entrytype"
'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) <> iOldAccountId Or CLng(oRequests("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("ClassSeasonID")    = oRequests("ClassSeasonID")
				oDisplay("accountid")        = oRequests("accountid")
				oDisplay("accountname")      = oRequests("accountname")
				oDisplay("accountnumber")    = oRequests("accountnumber")
				oDisplay("ispaymentaccount") = oRequests("ispaymentaccount")

				If oRequests("accountname") = "Citizen Accounts" Then
  					oDisplay("iscitizenaccount") = True 
				Else 
		  			oDisplay("iscitizenaccount") = False 
				End If

				oDisplay("receiptno")          = oRequests("paymentid")
				oDisplay("paymentdate")        = oRequests("paymentdate")
				oDisplay("paymenttypeid")      = oRequests("paymenttypeid")
				oDisplay("journalentrytypeid") = oRequests("journalentrytypeid")
				oDisplay("userid")             = oRequests("userid")
				oDisplay("creditamt")          = CDbl(0.00)
				oDisplay("debitamt")           = CDbl(0.00)
				oDisplay("totalamt")           = CDbl(0.00)
				iOldAccountId                  = CLng(oRequests("accountid"))
				iOldPaymentId                  = CLng(oRequests("paymentid"))
			End If 
			If oRequests("entrytype") = "credit" Then
				  oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRequests("amount"))
				  oDisplay("totalamt")  = oDisplay("totalamt") + CDbl(oRequests("amount"))
			End If 
			If oRequests("entrytype") = "debit" Then
  				oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRequests("amount"))
		  		oDisplay("totalamt") = oDisplay("totalamt") - CDbl(oRequests("amount"))
			End If 
			oDisplay.Update
			oRequests.MoveNext
		Loop
	End If 
	oRequests.Close 
	Set oRequests = Nothing 

	' Citizen Accounts
	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, "
	sSql = sSql & " ISNULL(L.paymenttypeid,0) AS paymenttypeid, P.userid, P.journalentrytypeid, L.ispaymentaccount, 1 AS iscitizenaccount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment P, egov_accounts A, egov_organizations_to_paymenttypes OP, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE L.paymentid = P.paymentid "
	sSql = sSql & " AND L.paymenttypeid = 4 "
	sSql = sSql & " AND A.accountid = OP.accountid "
	sSql = sSql & " AND OP.paymenttypeid = L.paymenttypeid "
	sSql = sSql & " AND OP.orgid = P.orgid "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, P.paymentid, L.entrytype"

'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) <> iOldAccountId Or CLng(oRequests("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("ClassSeasonID")      = oRequests("ClassSeasonID")
				oDisplay("accountid")          = oRequests("accountid")
				oDisplay("accountname")        = oRequests("accountname") 
				oDisplay("accountnumber")      = oRequests("accountnumber")
				oDisplay("ispaymentaccount")   = True 
				oDisplay("iscitizenaccount")   = True 
				oDisplay("receiptno")          = oRequests("paymentid")
				oDisplay("paymentdate")        = oRequests("paymentdate")
				oDisplay("paymenttypeid")      = oRequests("paymenttypeid")
				oDisplay("journalentrytypeid") = oRequests("journalentrytypeid")
				oDisplay("userid")             = oRequests("userid")
				oDisplay("creditamt")          = CDbl(0.00)
				oDisplay("debitamt")           = CDbl(0.00)
				oDisplay("totalamt")           = CDbl(0.00)
				iOldAccountId                  = CLng(oRequests("accountid"))
				iOldPaymentId                  = CLng(oRequests("paymentid"))
			End If 
			If oRequests("entrytype") = "credit" Then
				  oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRequests("amount"))
				  oDisplay("totalamt")  = oDisplay("totalamt") + CDbl(oRequests("amount"))
			End If 
			If oRequests("entrytype") = "debit" Then
				  oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRequests("amount"))
				  oDisplay("totalamt") = oDisplay("totalamt") - CDbl(oRequests("amount"))
			End If 
			oDisplay.Update
			oRequests.MoveNext
		Loop
		 
	End If 
	oRequests.Close
	Set oRequests = Nothing 

	If bHasData Then 
		'sort the Display recordset
		oDisplay.Sort = "ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC, receiptno ASC"

		' Show results
		oDisplay.MoveFirst
		'response.write "<div class=""receiptpaymentshadow"">" & vbcrlf
		response.Write "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">" & vbcrlf
		response.write "  <tr class=""tablelist"">"    & vbcrlf
		response.write "      <th>Season</th>"         & vbcrlf
		response.write "      <th>Account Name</th>"   & vbcrlf
		response.write "      <th>Account Number</th>" & vbcrlf
		response.write "      <th>Receipt No.</th>"    & vbcrlf
		response.write "      <th>Date</th>"           & vbcrlf
		response.write "      <th>Total Amt<br />Credited</th>" & vbcrlf
		response.write "      <th>Total Amt<br />Debited</th>"  & vbcrlf
		'  response.write "      <th>Total Amt<br />Transfered</th>" & vbcrlf
		response.write "  </tr>" & vbcrlf

		lcl_prev_ClassSeasonID = ""
		bgcolor         = "#eeeeee"
		iOldAccountId   = CLng(0)
		dCreditSubTotal = CDbl(0.00)
		dDebitSubTotal  = CDbl(0.00)
		dSubTotal       = CDbl(0.00)

		Do While Not oDisplay.EOF
			bgcolor = changeBGcolor(bgcolor,"#eeeeee","#ffffff")

			If iOldAccountId <> CLng(oDisplay("accountid")) Then 
				'Put out a sub total row
				If iOldAccountId <> CLng(0) Then 
					response.write "  <tr class=""totalrow"">" & vbcrlf
					response.write "      <td colspan=""5"" align=""right"">Sub-Total:</td>" & vbcrlf
					response.write "      <td align=""right"">" & FormatNumber(dCreditSubTotal, 2) & "</td>" & vbcrlf
					response.write "      <td align=""right"">" & FormatNumber(dDebitSubTotal, 2)  & "</td>" & vbcrlf
					'  				    	response.write "      <td align=""right"">" & FormatNumber(dSubTotal,2)        & "</td>" & vbcrlf
					response.write "  </tr>" & vbcrlf
					response.write "  <tr bgcolor=""" &  bgcolor  & """>" & vbcrlf
				Else 
					response.write "  <tr bgcolor=""" &  bgcolor  & """>" & vbcrlf
				End If 

				'Determine if we show the Season Name or not based on the previous row value.
				If lcl_prev_ClassSeasonID <> oDisplay("ClassSeasonID") Then 
					response.write "      <td align=""left"">"   & getSeasonName(oDisplay("ClassSeasonID")) & "</td>" & vbcrlf
				Else 
					response.write "      <td>&nbsp;</td>" & vbcrlf
				End If 

				response.write "      <td align=""left"">"   & oDisplay("accountname")   & "</td>" & vbcrlf
				response.write "      <td align=""center"">" & oDisplay("accountnumber") & "</td>" & vbcrlf

				iOldAccountId   = CLng(oDisplay("accountid"))
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal  = CDbl(0.00)
				dSubTotal       = CDbl(0.00)
			Else 
				'Need place holders 
				response.write "  <tr bgcolor=""" &  bgcolor  & """>" & vbcrlf
				response.write "      <td>&nbsp;</td>" & vbcrlf
				response.write "      <td>&nbsp;</td>" & vbcrlf
				response.write "      <td>&nbsp;</td>" & vbcrlf
			End If 

			If clng(oDisplay("journalentrytypeid")) > clng(2) Then 
				'citizen account activity
				response.write "      <td align=""center"">"
				response.write "          <a href=""../purchases/viewjournal.asp?uid=" & oDisplay("userid") & "&pid=" & oDisplay("receiptno") & "&rt=c&it=ci&jet=d"">" & oDisplay("receiptno") & "</a>"
				response.write "      </td>" & vbcrlf
			Else 
				'purchase
				response.write "       <td align=""center"">"
				response.write "           <a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("receiptno") & """>" & oDisplay("receiptno") & "</a>"
				response.write "       </td>" & vbcrlf
			End If 

			response.write "      <td align=""right"">" & FormatDateTime(oDisplay("paymentdate"), 2) & "</td>" & vbcrlf
			response.write "      <td align=""right"">" & FormatNumber(oDisplay("creditamt"), 2)     & "</td>" & vbcrlf
			response.write "      <td align=""right"">" & FormatNumber(oDisplay("debitamt"), 2)      & "</td>" & vbcrlf
			'  			response.write "      <td align=""right"">" & FormatNumber(oDisplay("totalamt"), 2)      & "</td>" & vbcrlf

			dCreditSubTotal = dCreditSubTotal + CDbl(oDisplay("creditamt"))
			dTotalCredit    = dTotalCredit + CDbl(oDisplay("creditamt"))
			dDebitSubTotal  = dDebitSubTotal + CDbl(oDisplay("debitamt"))
			dTotalDebit     = dTotalDebit + CDbl(oDisplay("debitamt"))
			dSubTotal       = dSubTotal + CDbl(oDisplay("totalamt"))
			dGrandTotal     = dGrandTotal + CDbl(oDisplay("totalamt"))

			response.write "  </tr>" & vbcrlf

			oDisplay.MoveNext
		Loop 

		'Put out a sub total row
		If iOldAccountId <> CLng(0) Then 
			response.write "  <tr class=""totalrow"">" & vbcrlf
			response.write "      <td colspan=""5"" align=""right"">Sub-Total:</td>" & vbcrlf
			response.write "      <td align=""right"">" & FormatNumber(dCreditSubTotal, 2) & "</td>" & vbcrlf
			response.write "      <td align=""right"">" & FormatNumber(dDebitSubTotal, 2)  & "</td>" & vbcrlf
			'  			response.write "      <td align=""right"">" & FormatNumber(dSubTotal,2)        & "</td>" & vbcrlf
			response.write "  </tr>" & vbcrlf
		End If 

		'Totals Row
		response.write "  <tr class=""totalrow"">" & vbcrlf
		response.write "      <td colspan=""5"" align=""right"">Totals:</td>" & vbcrlf
		response.write "      <td align=""right"">" & FormatNumber( dTotalCredit, 2 ) & "</td>" & vbcrlf
		response.write "      <td align=""right"">" & FormatNumber( dTotalDebit, 2 )  & "</td>" & vbcrlf
		'		response.write "      <td align=""right"">" & FormatNumber( dGrandTotal, 2 )  & "</td>" & vbcrlf
		response.write "  </tr>" & vbcrlf
		response.write "</table>" & vbcrlf
		'response.write "</div>"& vbcrlf

	Else
		response.write "<p>No information could be found for the criteria selected.</p>"
	End If

	oDisplay.Close
	Set oDisplay = Nothing

End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowAdminLocations( iLocationId )
'------------------------------------------------------------------------------------------------------------
Sub ShowAdminLocations( ByVal iLocationId )
	Dim sSql, oLocation
	
	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & session("orgid") & " ORDER BY name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open  sSql, Application("DSN"), 3, 1

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
Sub ShowPaymentLocations( ByVal iPaymentLocationId )

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
Sub ShowReportTypes( ByVal iReportType )
	
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
Sub ShowAdminUsers( ByVal iAdminUserId )
	Dim sSql, oUser
	
	sSql = "SELECT userid, firstname, lastname FROM users WHERE isrootadmin = 0 AND orgid = " & session("orgid") & " ORDER BY lastname, firstname"

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open  sSql, Application("DSN"), 3, 1

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
Sub ShowJournalEntryTypes( ByVal iJournalEntryTypeId )
	Dim sSql, oTypes
	
	sSql = "SELECT journalentrytypeid, displayname + ' Only' AS displayname FROM egov_journal_entry_types WHERE journalentrytype = 'refund' "
	sSql = sSql & " OR journalentrytype = 'purchase' ORDER BY displayorder"

	Set oTypes = Server.CreateObject("ADODB.Recordset")
	oTypes.Open  sSql, Application("DSN"), 3, 1

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

	oTypes.Close
	Set oTypes = Nothing 

End Sub 


'------------------------------------------------------------
Sub ShowClassSeasons( ByVal iClassSeasonID )
	Dim sSql, oClassSeason

	response.write "<select name=""ClassSeasonID"">" & vbcrlf
	response.write "<option value="""">All Seasons</option>" & vbcrlf

	sSql = "SELECT classseasonid, seasonname "
	sSql = sSql & " FROM egov_class_seasons "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY seasonyear, seasonname "

	Set oClassSeason = Server.CreateObject("ADODB.Recordset")
	oClassSeason.Open sSql, Application("DSN"), 3, 1

	Do While Not oClassSeason.EOF

		lcl_classseasonid = oClassSeason("classseasonid")

		If CStr(iClassSeasonID) = CStr(lcl_classseasonid) Then 
			lcl_selected = " selected=""selected"" "
		Else 
			lcl_selected = ""
		End If 

		response.write "<option value=""" & oClassSeason("classseasonid") & """" & lcl_selected & ">" & oClassSeason("seasonname") & "</option>" & vbcrlf

		oClassSeason.MoveNext
	Loop 

	response.write vbcrlf & "</select>"

	oClassSeason.Close
	Set oClassSeason = Nothing
	
End Sub


'---------------------------------------------------------
Function getSeasonName( ByVal p_classseasonid )
	Dim lcl_return, oSeasonName, sSql

	lcl_return = ""

	If p_classseasonid <> "" Then 
		sSql = "SELECT seasonname FROM egov_class_seasons WHERE classseasonid = " & p_classseasonid

		Set oSeasonName = Server.CreateObject("ADODB.Recordset")
		oSeasonName.Open sSql, Application("DSN"), 0, 1

		If Not oSeasonName.eof Then 
			lcl_return = oSeasonName("seasonname")
		End If 

		oSeasonName.Close
		Set oSeasonName = Nothing 

	End If 

	getSeasonName = lcl_return

End Function


'--------------------------------------------------------------------------------------------------
' void ShowAccountPicks sSelectName, iAccountNo, bShowAllPick
'--------------------------------------------------------------------------------------------------
Sub ShowAccountPicks( ByVal sSelectName, ByVal iAccountNo, ByVal bShowAllPick )
	Dim sSql, oRs

	sSql = "SELECT accountid, accountname FROM egov_accounts WHERE orgid = " & session("orgid")
	sSql = sSql & " AND accountstatus = 'A' ORDER BY accountname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	If bShowAllPick Then 
		response.write "<option value=""0"">Include All GL Accounts</option>"
	End If 
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("accountid") & """"

		If iAccountNo <> "" Then 
			If CLng(oRS("accountid")) = CLng(iAccountNo) Then
				response.write " selected=""selected"" "
			End If
		End If 

		response.write ">" & oRs("accountname") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 

End Sub

%>
