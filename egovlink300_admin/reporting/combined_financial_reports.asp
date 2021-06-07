<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="reporting_functions.asp" //-->
<!-- #include file="class_reporting_functions.asp" //-->
<!-- #include file="citizen_reporting_functions.asp" //-->
<!-- #include file="rental_reporting_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: combined_financial_reports.asp
' AUTHOR: SteveLoar
' CREATED: 03/18/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a combined report for Classes, Rentals, and Citizen Accounts. 
'		     Includes payment receipts, refunds, and account distribution reports.
'
' MODIFICATION HISTORY
' 1.0   3/18/2014	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iLocationId, iAdminUserId, iPaymentLocationId, iFinancialReportId, sWhereClause
Dim from_time, to_time, where_time, today, fromDate, toDate, exportReportScript
Dim iReservationTypeId, showReservationTypePicks, sNameLike, showNameLikeField
Dim showPaymentLocationPicks, sRptType, showGLAccountPicks, iAccountNo, showReportTypeAndEntries
Dim iReportType, iJournalEntryTypeId, sRptTitle, bOrgHasAccounts, bIncludeListOption
Dim sNameClause, sCitizenNameClause, transactionSource


' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "financial deposit report", sLevel	' In common.asp

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

showPaymentLocationPicks = true
If request("paymentlocationid") = "" Then
	iPaymentLocationId = 0
Else
	iPaymentLocationId = CLng(request("paymentlocationid"))
End If 

If request("accountid") = "" Then
	iAccountNo = 0
Else
	iAccountNo = CLng(request("accountid"))
End If 

If request("reporttype") = "" Then 
	iReportType = CLng(1)
Else
	iReportType = CLng(request("reporttype"))
End If 

If iReportType = CLng(1) Then
	sRptTitle = "Summary"
	sRptType = "Summary"
ElseIf iReportType = CLng(2) Then
	sRptTitle = "Detail"
	sRptType = "Detail"
Else 
	sRptTitle = "List"
	sRptType = "List"
End If 

If request("journalentrytypeid") = "" Then
	iJournalEntryTypeId = 0
Else
	iJournalEntryTypeId = CLng(request("journalentrytypeid"))
End If 

' Not every org has general ledger accounts so we need to be able to hide/show accordingly.
bOrgHasAccounts = OrgHasFeature("gl accounts")


' There should be 9 financial reports'

If request("financialreportid") = "" Then
	iFinancialReportId = 1
Else
	iFinancialReportId = CLng(request("financialreportid"))
End If 


showReservationTypePicks = false
iReservationTypeId = 0
If iFinancialReportId = 2 Or iFinancialReportId = 8 Then
	showReservationTypePicks = true
	If request("reservationtypeid") <> "" Then
		iReservationTypeId = CLng(request("reservationtypeid"))
	End If 
End If 


showNameLikeField = false 
sNameLike = ""
If iFinancialReportId = 3 Or iFinancialReportId = 6 Then
	showNameLikeField = true 
	showPaymentLocationPicks = false 
	iPaymentLocationId = 0
	
	If request("namelike") <> "" Then 
		sNameLike = request("namelike")
	End If 
End If 

If iFinancialReportId = 4 Or iFinancialReportId = 5 Or iFinancialReportId = 9 Then
	showPaymentLocationPicks = false 
	iPaymentLocationId = 0
End If 

' These are the account distribution reports (7,8,9) '
If iFinancialReportId > 6 Then 
	showGLAccountPicks = true
	showReportTypeAndEntries = true 
	If iFinancialReportId = 7 Then
		transactionSource = "classes"
	ElseIf iFinancialReportId = 8 Then 
		transactionSource = "rentals"
	Else
		'should be 9'
		transactionSource = "citizens"
	End If
Else
	showGLAccountPicks = false
	iAccountNo = 0
	showReportTypeAndEntries = false 
	transactionSource = ""
	iJournalEntryTypeId = 0
	'iReportType = CLng(0)
End If  

sNameClause = ""
sCitizenNameClause = ""
If iFinancialReportId = 8 Then 
	bIncludeListOption = true 
	showNameLikeField = true 
	If request("namelike") <> "" Then 
		sNameLike = request("namelike")
		sNameClause = " AND (renterlastname LIKE '%" & dbsafe( sNameLike ) & "%' OR renterfirstname LIKE '%" & dbsafe( sNameLike ) & "%') "
		sCitizenNameClause = " AND (renterlastname LIKE '%" & dbsafe( sNameLike ) & "%' OR renterfirstname LIKE '%" & dbsafe( sNameLike ) & "%')"
	Else
		sNameLike = ""
	End If 
Else 
	bIncludeListOption = false 
End If


' BUILD SQL WHERE CLAUSE
sWhereClause = " WHERE orgid = " & session("orgid") 

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

If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		sWhereClause = sWhereClause & " AND paymentlocationid = 3 " 
	Else
		sWhereClause = sWhereClause & " AND paymentlocationid < 3 " 
	End If 
End If 

If iReservationTypeId > 0 Then
	sWhereClause = sWhereClause & " AND reservationtypeid = " & iReservationTypeId
End If 

If sNameLike <> "" And iFinancialReportId <> 8 Then
	sWhereClause = sWhereClause & " AND ( userfname LIKE '%" & DBsafe( request("namelike") ) & "%' OR userlname LIKE '%" & DBsafe( request("namelike") ) & "%' )"
End If 

If CLng(iAccountNo) > CLng(0) Then
	sWhereClause = sWhereClause & " AND accountid = " & iAccountNo & " "
End If 

If iJournalEntryTypeId > 0 Then 
	sWhereClause = sWhereClause & " AND journalentrytypeid = " & iJournalEntryTypeId
End If 


' set the export for the current report'
Select Case iFinancialReportId 
	Case 1
		' class payments'
		exportReportScript = "receipt_payment_export.asp"
	Case 2
		' rentals payments'
		exportReportScript = "../rentals/rentalreceiptpaymentexport.asp"
	Case 3
		' citizen account deposits'
		exportReportScript = "CitizenAccountDepositsExport.asp"
	Case 4
		' class refunds'
		exportReportScript = "refund_payment_export.asp"
	Case 5
		' rentals refunds '
		exportReportScript = "../rentals/rentalrefundpaymentexport.asp"
	Case 6
		' citizen account withdrawals'
		exportReportScript = "CitizenAccountRefundsExport.asp"
	Case 7
		' class account distribution '
		exportReportScript = "account_distribution_export.asp"
	Case 8
		' rentals account distribution '
		exportReportScript = "../rentals/rentalaccountdistributionexport.asp"
	Case 9
		' citizen accounts, account distribution'
		exportReportScript = "citizen_account_distribution_export.asp"
	Case Else
		' if all else does not match for some reason, take them to a class payment report, for now.'
		exportReportScript = "receipt_payment_export.asp"
End Select 


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
  	
	<script src="../scripts/getdates.js"></script>
	<script src="../scripts/isvaliddate.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	<script>
	
		var validate = function( media ) {
			if ( datesAreValid() ) {
				if ( media === 'screen') {
					document.CombinedReport.action = 'combined_financial_reports.asp';
				}
				else {
					// this changes based on the report selected
					document.CombinedReport.action = '<% = exportReportScript %>';
				}
				document.CombinedReport.submit();
				return true;
			}
			else {
				return false;
			}
		};
		
		var datesAreValid = function() {
			var okToPost = true;
			// check from date
			if ($("#fromDate").val() != "") {
				if (! isValidDate($("#fromDate").val()) ) {
					inlineMsg("fromDate","<strong>Invalid Value: </strong>The transaction 'from date' should be a valid date in the format of MM/DD/YYYY.");
					okToPost = false;
				}
			}
			// check to date
			if ($("#toDate").val() != "") {
				if (! isValidDate($("#toDate").val()) ) {
					inlineMsg("toDate","<strong>Invalid Value: </strong>The transaction 'to date' should be a valid date in the format of MM/DD/YYYY.");
					okToPost = false;
				}
			}
			
			return okToPost;
			
		};
		
		// these function set up the date pickers
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

	<form action="combined_financial_reports.asp" method="post" name="CombinedReport">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><span id="combinedfinancialtitle">Financial Deposit Reports</td>
		</tr>
		<tr>
			<td>
				<!--BEGIN: FILTERS-->
				<fieldset class="fieldset">
					<legend>Select From The Following:</legend>
				
					<p>
						<% ShowReportChoices iFinancialReportId %>
					</p>
					
					<table border="0" cellpadding="0" cellspacing="0">
						<tr><td colspan="3"><strong>Transaction Date:</strong></td></tr>
						<tr>
							<td nowrap>
								<label for="fromDate">From:</label>
								<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />&nbsp;
								<% DrawTimeChoices "fromtime", from_time %>
							</td>
							<td nowrap>
								<label for="toDate">To:</label>
								<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />&nbsp;
								<% DrawTimeChoices "totime", to_time %>
							</td>
							<td><% DrawDateChoices "Date" %></td>
						</tr>
					</table>
					
					<% If showReservationTypePicks Then %>
						<p>
							<label for="reservationtypeid">Reservation Type:</label><% ShowRentalReservationTypeFilter iReservationTypeId, True  %>
						</p>
					<% End If %>
					
					<p>
						<label for="locationid">Admin Location:</label><% ShowAdminLocations iLocationId %>&nbsp;&nbsp;
						<label for="adminuserid">Admin:</label><% ShowAdminUsers iAdminUserId %>&nbsp;&nbsp;
					</p>
					
					<% If showNameLikeField Then %>
						<p>
							<label for="namelike">Name Like:</label><input type="text" id="namelike" name="namelike" value="<%=sNameLike%>" placeholder="Enter a Partial Name" maxlength="100" size="100" />
						</p>
					<% End If %>					
					
					<% If showPaymentLocationPicks Or showReportTypeAndEntries Then %>
						<p>
							<% If showPaymentLocationPicks Then %>
								<label for="paymentlocationid">Payment Location:</label><% ShowPaymentLocations iPaymentLocationId %>&nbsp;&nbsp;
							<% End If %>
							<% If showReportTypeAndEntries Then %>
								<label for="reporttype">Report Type:</label><% ShowReportTypes iReportType, bIncludeListOption %>&nbsp;&nbsp;
								<label for="journalentrytypeid">Entries:</label><% ShowJournalEntryTypes iJournalEntryTypeId, transactionSource %>
							<% End If %>
						</p>
					<% End If %>
					
					<% If showGLAccountPicks Then %>
						<% If bOrgHasAccounts Then %>
							<p>
								<label for="accountid">GL Account:</label><% ShowAccountPicks "accountid", iAccountNo, True %>
							</p>
						<% Else 
							response.write "<input type=""hidden"" id=""accountid"" name=""accountid"" value=""0"" />"
						   End If %>
					<% End If %>

					<p>
						<input class="button" type="button" value="View Report" onclick="validate('screen');" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="validate('export')" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
	</table>
	
	</form>
	
	<div id="resultsdisplay">
		<% 
			
			Select Case iFinancialReportId 
				Case 1
					Display_Class_Payment_Report sWhereClause 
				Case 2
					Display_Rental_Payment_Report sWhereClause
				Case 3
					Display_Citizen_Payment_Report sWhereClause
				Case 4
					Display_Class_Refund_Report sWhereClause
				Case 5
					Display_Rental_Refund_Report sWhereClause
				Case 6
					Display_Citizen_Refund_Report sWhereClause
				Case 7
					If sRptType = "Summary" Then
						Display_Class_Acct_Dist_Summary sWhereClause
					Else
						' Details'
						Display_Class_Acct_Dist_Details sWhereClause
					End If 
				Case 8
					If LCase(sRptType) = "summary" Then
						Display_Rental_Summary sWhereClause, sNameClause
					Else
						Display_Rental_Details sWhereClause, sNameClause, LCase(sRptType), sCitizenNameClause
					End If 
				Case 9
					If LCase(sRptType) = "summary" Then
						Display_Citizen_Acct_Dist_Summary sWhereClause
					Else 
						Display_Citizen_Acct_Dist_Details sWhereClause
					End If 
				Case Else
					response.write vbcrlf & "<div>This report does not exsist. Please try another.</div>"
			End Select 
		%>
	</div>
	
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


