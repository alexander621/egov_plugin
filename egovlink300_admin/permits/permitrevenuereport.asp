<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitrevenuereport.asp
' AUTHOR: Steve Loar
' CREATED: 01/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permit payments by payment media (cash, check, cc)
'
' MODIFICATION HISTORY
' 1.0   01/28/2009	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromPaymentDate, sToPaymentDate, sStreetNumber, sStreetName, sPermitNo
Dim sPayor, sInvoiceNo, sDisplayDateRange, sApplicant, iWaivedPick, iPermitCategoryId, sPermitLocation

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "permit revenue report", sLevel	' In common.asp

If request("waivedpick") <> "" Then
	iWaivedPick = clng(request("waivedpick"))
	sSearch = sSearch & " AND I.allfeeswaived = " & request("waivedpick")
Else
	iWaivedPick = 0
	sSearch = sSearch & " AND I.allfeeswaived = 0 "
End If 

' Handle payment date range. always want some dates to limit the search
If request("topaymentdate") <> "" And request("frompaymentdate") <> "" Then
	sFromPaymentDate = request("frompaymentdate")
	sToPaymentDate = request("topaymentdate")
	If iWaivedPick = 0 Then 
		sSearch = sSearch & " AND (J.paymentdate >= '" & request("frompaymentdate") & "' AND J.paymentdate < '" & DateAdd("d",1,request("topaymentdate")) & "' ) "
	Else
		sSearch = sSearch & " AND (I.invoicedate >= '" & request("frompaymentdate") & "' AND I.invoicedate < '" & DateAdd("d",1,request("topaymentdate")) & "' ) "
	End If 
	sDisplayDateRange = "From: " & request("frompaymentdate") & " &nbsp;To: " & request("topaymentdate")
Else
	' initially set these to today
	sFromPaymentDate = FormatDateTime(Date,2)
	sToPaymentDate = FormatDateTime(Date,2)
	sDisplayDateRange = ""
End If 

' handle address pick
If request("residentstreetnumber") <> "" Then 
	sStreetNumber = request("residentstreetnumber")
	sSearch = sSearch & "AND A.residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
End If 
If request("streetname") <> "" And request("streetname") <> "0000" Then 
	sStreetName = request("streetname")
	sSearch = sSearch & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"
End If 

' handle the permit number
If request("permitno") <> "" Then 
	sPermitNo = Trim(request("permitno"))
	sSearch = sSearch & BuildPermitNoSearch( sPermitNo )	' in permitcommonfunctions.asp
End If 

If request("invoiceno") <> "" Then 
	sInvoiceNo = CLng(request("invoiceno"))
	sSearch = sSearch & " AND I.invoiceid = " & sInvoiceNo
End If 

If request("applicant") <> "" Then 
	sApplicant = request("applicant")
	sSearch = sSearch & " AND ( C.company LIKE '%" & dbsafe(sApplicant) & "%' OR C.firstname LIKE '%" & dbsafe(sApplicant) & "%' OR C.lastname LIKE '%" & dbsafe(sApplicant) & "%' ) "
End If 

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = request("permitcategoryid")
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSearch = sSearch & " AND P.permitcategoryid = " & iPermitCategoryId
	End If 
Else 
	iPermitCategoryId = "0"
End If 

If request("permitlocation") <> "" Then
	sPermitLocation = request("permitlocation")
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 

'If sSearch <> "" Then 
'	session("sSql") = sSearch
'Else 
'	session("sSql") = ""
'End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--

		function validate( sReport )
		{
			// check the payment from date
			if ($("#frompaymentdate").val() != '')
			{
				if (! isValidDate($("#frompaymentdate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#frompaymentdate").focus();
					return;
				}
			}
			// check the inspection to date
			if ($("#topaymentdate").val() != '')
			{
				if (! isValidDate($("#topaymentdate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#topaymentdate").focus();
					return;
				}
			}

			if (sReport == 'view')
			{
				document.frmPermitSearch.action = 'permitrevenuereport.asp';
			}
			else
			{
				document.frmPermitSearch.action = 'permitrevenuereportexport.asp';
			}
			document.frmPermitSearch.submit();
		}

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<div id="idControls" class="noprint">
		<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	</div>

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<span id="printdaterange"><font size="+1"><strong><%=sDisplayDateRange%></strong></font></span>
				<font size="+1"><strong>Permit Revenue Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitrevenuereport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Permit Category:</td>
									<td><%	ShowPermitCategoryPicks iPermitCategoryId	' in permitcommonfunctions.asp	%></td>
								</tr>
								<tr>
									<td>Payment Date:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" id="frompaymentdate" name="frompaymentdate" value="<%=sFromPaymentDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp; To:
										<input type="text" id="topaymentdate" name="topaymentdate" value="<%=sToPaymentDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp;
										<%DrawPriorDateChoices "paymentdate" %>
									</td>
								</tr>
								<tr>
									<td>Address:</td><td><% DisplayLargeAddressList sStreetNumber, sStreetName %></td>
								</tr>
								<tr>
									<td>Location Like:</td><td><input type="text" name="permitlocation" size="100" maxlength="100" value="<%=sPermitLocation%>" /></td>
								</tr>
								<tr>
									<td>Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
								</tr>
								<tr>
									<td>Invoice #:</td><td><input type="text" id="invoiceno" name="invoiceno" size="20" maxlength="20" value="<%=sInvoiceNo%>" /></td>
								</tr>
								<tr>
									<td>Waived Invoices:</td>
									<td><% ShowWaivedChoices iWaivedPick %></td>
								</tr>
								<tr>
									<td>Applicant:</td><td><input type="text" id="applicant" name="applicant" size="50" maxlength="50" value="<%=sApplicant%>" /></td>
								</tr>
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate( 'view' );" />&nbsp;&nbsp;
<%										If request("isview") <> "" Then		%>
											<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="validate( 'excel' );" />
<%										End If		%>
									</td>
								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

			<!-- Begin: Report Display -->
<%			' if they choose to view the report, then display the payments
			If request("isview") <> "" Then	
				DisplayRevenueByFeeTypes sSearch, iWaivedPick
			Else 
				response.write "<strong>To view the permit revenue report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
			End If 
%>
			<!-- END: Report Display -->
		</div>
		</div>
		
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sStreetNumber, sStreetName 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If NOT oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = oRs("residentstreetprefix") & " " 
			End If 

			sCompareName = sCompareName & oRs("residentstreetname")

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetsuffix")
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetdirection")
			End If 

			response.write vbcrlf & "<option value=""" & sCompareName & """"

			If sStreetName = sCompareName Then 
				response.write " selected=""selected"" "
			End If 

			response.write " >"
			response.write sCompareName & "</option>" 
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayRevenueByFeeTypes sSearch, iWaivedPick 
'--------------------------------------------------------------------------------------------------
Sub DisplayRevenueByFeeTypes( ByVal sSearch, ByVal iWaivedPick )
	Dim sSql, oRs, iRowCount, dTotalAmount, dSubTotalAmount, iOldReportingFeeTypeId, sFeeType
	Dim sDate, sFrom, sWhere, sSum

	dTotalAmount = CDbl(0.00)
	dSubTotalAmount = CDbl(0.00)
	iOldReportingFeeTypeId = CLng(-1)
	iRowCount = 0

	If clng(iWaivedPick) = clng(0) Then
		sDate = "J.paymentdate"
		sFrom = "egov_class_payment J, egov_accounts_ledger AL,"
		sWhere = " AND II.permitfeeid = AL.permitfeeid AND II.invoiceid = AL.invoiceid AND J.paymentid = AL.paymentid"
		sSum = "AL.amount"
	Else
		sDate = "I.invoicedate"
		sFrom = ""
		sWhere = ""
		sSum = "I.totalamount"
	End If 

	' Pull the list of permits for the period and that match the selections
	sSql = "SELECT ISNULL(II.feereportingtypeid,0) AS feereportingtypeid, I.permitid, I.invoiceid, " & sDate & " AS paymentdate, "
	sSql = sSql & " ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, "
	sSql = sSql & " SUM(" & sSum & ") AS amount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitinvoices I, "
	sSql = sSql & sFrom & " egov_permitaddress A, egov_permits P, egov_permitcontacts C, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE II.invoiceid = I.invoiceid AND I.orgid = " & session("orgid") & " AND I.isvoided = 0 "
	sSql = sSql & " AND I.permitid = C.permitid AND C.isapplicant = 1 AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & sSearch
	sSql = sSql & " AND I.permitid = P.permitid AND P.isvoided = 0 AND A.permitid = I.permitid " & sWhere
	sSql = sSql & " GROUP BY II.feereportingtypeid, I.permitid, I.invoiceid, " & sDate & ", P.permitlocation, R.locationtype, dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) "
	sSql = sSql & " ORDER BY II.feereportingtypeid, I.permitid, I.invoiceid, " & sDate 
	'response.write "<!-- " & sSql & " --><br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""revenuereportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""revenuereport"">"
		response.write vbcrlf & "<tr><th>Fee Type</th><th>Permit #</th><th>Invoice #</th><th>Payment<br />Date</th><th>Address/Location</th><th>Amount</th></tr>"

		Do While Not oRs.EOF
			If iOldReportingFeeTypeId <> CLng(oRs("feereportingtypeid")) Then
				If iOldReportingFeeTypeId <> CLng(-1) Then
					response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">&nbsp;</td><td>" & sFeeType & " Total</td>"
					response.write "<td align=""right"">" & FormatNumber(dSubTotalAmount,2) & "</td></tr>"
					dSubTotalAmount = CDbl(0.00)
				End If 
				iRowCount = 0
				iOldReportingFeeTypeId = CLng(oRs("feereportingtypeid"))
				' Print permit fee type name
				sFeeType = GetFeeReportingType( iOldReportingFeeTypeId )
				response.write "<tr class=""totalrow""><td colspan=""6"">&nbsp;" & sFeeType & "</td></tr>"
			End If 

			dTotalAmount = dTotalAmount + CDbl(oRs("amount"))
			dSubTotalAmount = dSubTotalAmount + CDbl(oRs("amount"))
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			response.write "<td>&nbsp;</td>"
			response.write "<td align=""center"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
			response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"
			response.write "<td align=""center"" nowrap=""nowrap"">" & FormatDateTime(oRs("paymentdate"),2) & "</td>"

			'response.write "<td nowrap=""nowrap"">&nbsp;" & oRs("permitaddress") & "</td>"
			response.write "<td nowrap=""nowrap"" class=""addresscol"">"
			Select Case oRs("locationtype")
				Case "address"
					response.write oRs("permitaddress")

				Case "location"
					response.write Replace(oRs("permitlocation"),Chr(10),"<br />")

				Case Else
					response.write "&nbsp;"

			End Select  
			response.write "</td>"

			response.write "<td align=""right"">" & FormatNumber(oRs("amount"),2) & "&nbsp;</td>"
			oRs.MoveNext
		Loop 
		' last sub total Row
		If iOldReportingFeeTypeId <> CLng(-1) Then
			response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">&nbsp;</td><td>" & sFeeType & " Total</td>"
			response.write "<td align=""right"">" & FormatNumber(dSubTotalAmount,2) & "</td></tr>"
		End If 
		' Grand Total Row
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">&nbsp;</td><td>Grand Total</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalAmount,2) & "</td></tr>"
		response.write "</table></div>"
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowWaivedChoices iWaivedPick 
'--------------------------------------------------------------------------------------------------
Sub ShowWaivedChoices( ByVal iWaivedPick )

	response.write vbcrlf & "<select name=""waivedpick"">"
	response.write vbcrlf & "<option value=""0"""
	If clng(iWaivedPick) = clng(0) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Exclude</option>"
	response.write vbcrlf & "<option value=""1"""
	If clng(iWaivedPick) = clng(1) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Only Show</option>"
	response.write vbcrlf & "</select>"

End Sub 



%>
