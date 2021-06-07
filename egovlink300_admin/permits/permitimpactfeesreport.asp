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
' 1.1	09/21/2009	Steve Loar - Changed the no data found condition to display a $0 row instead of "No Data"
'								 Per the request of Loveland, OH
' 1.2	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromPaymentDate, sToPaymentDate, sStreetNumber, sStreetName, sPermitNo, sDisplayDateRange
Dim sYearStart, iStartYear, sYearEnd, iPermitCategoryId, sPermitLocation

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "permit impact fees report", sLevel	' In common.asp

' Handle payment date range. always want some dates to limit the search
If request("topaymentdate") <> "" And request("frompaymentdate") <> "" Then
	sFromPaymentDate = request("frompaymentdate")
	sToPaymentDate = request("topaymentdate")
	sSearch = sSearch & " AND (J.paymentdate >= '" & request("frompaymentdate") & "' AND J.paymentdate < '" & DateAdd("d",1,request("topaymentdate")) & "' ) "
	sDisplayDateRange = "From: " & request("frompaymentdate") & " &nbsp;To: " & request("topaymentdate")
Else
	' initially set these to today
	sFromPaymentDate = FormatDateTime(Date,2)
	sToPaymentDate = FormatDateTime(Date,2)
	sDisplayDateRange = ""
End If 

iStartYear = Year(CDate(sToPaymentDate))
sYearStart = "01/01/" & iStartYear
sYearEnd = DateAdd("d",1,CDate(sToPaymentDate))

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

		function validate()
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
				<font size="+1"><strong>Permit Impact Fees Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitimpactfeesreport.asp">
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
									<td>YTD Total Range:</td>
									<td nowrap="nowrap">
										The YTD Totals will be for the year of the second date and will end on that date.
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
<!--								<tr>
									<td>Payor:</td><td><input type="text" id="payor" name="payor" size="50" maxlength="50" value="<%=sPayor%>" /></td>
								</tr>

								<tr>
									<td>Invoice #:</td><td><input type="text" id="invoiceno" name="invoiceno" size="20" maxlength="20" value="<%=sInvoiceNo%>" /></td>
								</tr>
-->
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										
										'If request("isview") <> "" Then		%>
										<!--	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="location.href='permitissuedreportexport.asp'" /> -->
<%										'End If		%>
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
				DisplayImpactFeesReport sSearch, sYearStart, sYearEnd
			Else 
				response.write "<strong>To view the permit impact fees report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
	Dim sSql, oRs, sCompareName, sOldCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"
		sOldCompareName = "qwerty"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = UCase(oRs("residentstreetprefix")) & " " 
			End If 

			sCompareName = sCompareName & UCase(oRs("residentstreetname"))

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & UCase(oRs("streetsuffix"))
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & UCase(oRs("streetdirection"))
			End If 

			If sOldCompareName <> sCompareName Then 
				' only write out unique values
				sOldCompareName = sCompareName
				response.write vbcrlf & "<option value=""" & sCompareName & """"

				If sStreetName = sCompareName Then 
					response.write " selected=""selected"" "
				End If 

				response.write " >"
				response.write sCompareName & "</option>" 
			End If 

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayImpactFeesReport sSearch, sYearStart, sYearEnd 
'--------------------------------------------------------------------------------------------------
Sub DisplayImpactFeesReport( ByVal sSearch, ByVal sYearStart, ByVal sYearEnd )
	Dim sSql, oRs, iRowCount, dTotalAmount, dWaterMeter, dRoadImpact, dRecImpact, iUnits, sPermitNo
	Dim sAddress, dPaymentDate, iPermitId, dRowTotal, dWaterTotal, dRoadTotal, dRecTotal, dWaterImpact
	Dim dWaterImpactTotal, dWaterImpactYTD, dWaterYTD, dRoadYTD, dRecYTD, dTotalYTD

	dTotalAmount = CDbl(0.00)
	iRowCount = 0
	dWaterImpact = CDbl(0.00) 
	dWaterMeter = CDbl(0.00)
	dRoadImpact = CDbl(0.00)
	dRecImpact = CDbl(0.00)
	dRowTotal = CDbl(0.00)
	iPermitId = CLng(0)
	dWaterImpactTotal = CDbl(0.00)
	dWaterTotal = CDbl(0.00)
	dRoadTotal = CDbl(0.00)
	dRecTotal = CDbl(0.00)
	dWaterImpactYTD = CDbl(0.00)
	dWaterYTD = CDbl(0.00)
	dRoadYTD = CDbl(0.00)
	dRecYTD = CDbl(0.00)
	dTotalYTD = CDbl(0.00)

	sSql = "SELECT YEAR(J.paymentdate) AS paymentyear, MONTH(J.paymentdate) AS paymentmonth, DAY(J.paymentdate) AS paymentday, "
	sSql = sSql & " P.permitnumberyear, P.permitnumber, I.permitid, R.iswaterimpact, R.iswatermeter, R.isroadimpact, "
	sSql = sSql & " R.isrecreationimpact, ISNULL(P.permitlocation,'') AS permitlocation, LR.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, "
	sSql = sSql & " SUM(ISNULL(P.residentialunits,0)) AS residentialunits, SUM(AL.amount) AS amount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitinvoices I, egov_permitfeereportingtypes R, "
	sSql = sSql & " egov_accounts_ledger AL, egov_class_payment J, egov_permits P, egov_permitaddress A, egov_permitlocationrequirements LR "
	sSql = sSql & " WHERE  II.invoiceid = I.invoiceid AND I.orgid = " & session("orgid") & " AND I.isvoided = 0 AND "
	sSql = sSql & " J.paymentid = AL.paymentid  AND A.permitid = I.permitid AND P.permitlocationrequirementid = LR.permitlocationrequirementid "
	sSql = sSql & " AND (iswaterimpact = 1 OR iswatermeter = 1 OR isroadimpact = 1 OR isrecreationimpact = 1) AND P.isvoided = 0 AND P.permitid = I.permitid "
	sSql = sSql & " AND II.permitfeeid = AL.permitfeeid AND I.invoiceid = AL.invoiceid AND R.feereportingtypeid = II.feereportingtypeid " & sSearch
	sSql = sSql & " GROUP BY YEAR(J.paymentdate),MONTH(J.paymentdate),DAY(J.paymentdate), P.permitnumberyear, P.permitnumber, I.permitid, R.iswaterimpact, R.iswatermeter, R.isroadimpact, R.isrecreationimpact, P.permitlocation, LR.locationtype, dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) "
	sSql = sSql & " ORDER BY YEAR(J.paymentdate),MONTH(J.paymentdate),DAY(J.paymentdate), P.permitnumberyear, P.permitnumber, I.permitid, R.iswaterimpact DESC, R.iswatermeter DESC, R.isroadimpact DESC, R.isrecreationimpact DESC"

	'response.write "<!-- " & sSql & "<br /><br /> -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	response.write vbcrlf & "<div id=""impactfeereportshadow"">"
	response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""impactfeereport"">"
	response.write vbcrlf & "<tr><th>Payment<br />Date</th><th>Address/Location</th><th>Permit #</th><th># of Units</th><th>Water Impact</th><th>Water Meter</th><th>Road Impact</th><th>Rec. Impact</th><th>Total</th></tr>"

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If iPermitId <> CLng(oRs("permitid")) Then
				If iPermitId <> CLng(0) Then 
					' Write a row out
					iRowCount = iRowCount + 1
					response.write vbcrlf & "<tr"
					If iRowCount Mod 2 = 0 Then
						response.write " class=""altrow"""
					End If 
					response.write ">"
					response.write "<td align=""center"" nowrap=""nowrap"">" & dPaymentDate & "</td>"

					'response.write "<td nowrap=""nowrap"">&nbsp;" & sAddress & "</td>"
					response.write "<td nowrap=""nowrap"" class=""addresscol"">"
					response.write sAddress
					response.write "</td>"

					response.write "<td align=""center"">" & GetPermitNumber( iPermitId ) & "</td>"
					response.write "<td align=""center"">" & iUnits & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dWaterImpact,2) & "&nbsp;</td>"
					response.write "<td align=""right"">" & FormatNumber(dWaterMeter,2) & "&nbsp;</td>"
					response.write "<td align=""right"">" & FormatNumber(dRoadImpact,2) & "&nbsp;</td>"
					response.write "<td align=""right"">" & FormatNumber(dRecImpact,2) & "&nbsp;</td>"
					response.write "<td align=""right"">" & FormatNumber(dRowTotal,2) & "&nbsp;</td>"
					response.write "</tr>"
				End If
				
				' Set the new values
				'dPaymentDate = FormatDateTime(oRs("paymentdate"),2)
				dPaymentDate = FormatDateTime((oRs("paymentmonth") & "/" & oRs("paymentday") & "/" & oRs("paymentyear")),2)
				iPermitId = CLng(oRs("permitid"))
				'sAddress = oRs("permitaddress")
				Select Case oRs("locationtype")
					Case "address"
						sAddress = oRs("permitaddress")

					Case "location"
						sAddress = Replace(oRs("permitlocation"),Chr(10),"<br />")

					Case Else
						sAddress = "&nbsp;"
				End Select  
				iUnits = oRs("residentialunits")
				dWaterImpact = CDbl(0.00)
				dWaterMeter = CDbl(0.00)
				dRoadImpact = CDbl(0.00)
				dRecImpact = CDbl(0.00)
				dRowTotal = CDbl(0.00)

			End If 
			'Add to running amounts
			If oRs("iswaterimpact") Then 
				dWaterImpact = dWaterImpact + CDbl(oRs("amount"))
				dWaterImpactTotal = dWaterImpactTotal + CDbl(oRs("amount"))
			End If 
			If oRs("iswatermeter") Then 
				dWaterMeter = dWaterMeter + CDbl(oRs("amount"))
				dWaterTotal = dWaterTotal + CDbl(oRs("amount"))
			End If 
			If oRs("isroadimpact") Then 
				dRoadImpact = dRoadImpact + CDbl(oRs("amount"))
				dRoadTotal = dRoadTotal + CDbl(oRs("amount"))
			End If 
			If oRs("isrecreationimpact") Then 
				dRecImpact = dRecImpact + CDbl(oRs("amount"))
				dRecTotal = dRecTotal + CDbl(oRs("amount"))
			End If 
			dRowTotal = dRowTotal + CDbl(oRs("amount"))
			dTotalAmount = dTotalAmount + CDbl(oRs("amount"))

			oRs.MoveNext
		Loop 

		If iPermitId <> CLng(0) Then
			' Write the last row
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			response.write "<td align=""center"" nowrap=""nowrap"">" & dPaymentDate & "</td>"
			response.write "<td nowrap=""nowrap"">&nbsp;" & sAddress & "</td>"
			response.write "<td align=""center"">" & GetPermitNumber( iPermitId ) & "</td>"
			response.write "<td align=""center"">" & iUnits & "</td>"
			response.write "<td align=""right"">" & FormatNumber(dWaterImpact,2) & "&nbsp;</td>"
			response.write "<td align=""right"">" & FormatNumber(dWaterMeter,2) & "&nbsp;</td>"
			response.write "<td align=""right"">" & FormatNumber(dRoadImpact,2) & "&nbsp;</td>"
			response.write "<td align=""right"">" & FormatNumber(dRecImpact,2) & "&nbsp;</td>"
			response.write "<td align=""right"">" & FormatNumber(dRowTotal,2) & "&nbsp;</td>"
			response.write "</tr>"
		End If 

	Else
		'response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"

		' Write out a $0 row for the report at the request of Loveland OH.
		response.write vbcrlf & "<tr>"
		response.write "<td align=""center"" nowrap=""nowrap"">&nbsp;</td>"
		response.write "<td nowrap=""nowrap"">&nbsp;</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(0,2) & "&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(0,2) & "&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(0,2) & "&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(0,2) & "&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(0,2) & "&nbsp;</td>"
		response.write "</tr>"
	End If 

	' Grand Total Row
	response.write vbcrlf & "<tr class=""totalrow""><td>&nbsp;</td><td>Totals</td><td colspan=""2"">&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dWaterImpactTotal,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dWaterTotal,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dRoadTotal,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dRecTotal,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dTotalAmount,2) & "&nbsp;</td>"
	response.write "</tr>"

	' YTD Total Row
	GetYTDTotals sYearStart, sYearEnd, dWaterImpactYTD, dWaterYTD, dRoadYTD, dRecYTD 
	dTotalYTD = dWaterImpactYTD + dWaterYTD + dRoadYTD + dRecYTD
	response.write vbcrlf & "<tr class=""totalrow""><td>&nbsp;</td><td>Year to Date</td><td colspan=""2"">&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dWaterImpactYTD,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dWaterYTD,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dRoadYTD,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dRecYTD,2) & "&nbsp;</td>"
	response.write "<td align=""right"">" & FormatNumber(dTotalYTD,2) & "&nbsp;</td>"
	response.write "</tr>"

	response.write "</table></div>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetYTDTotals( sYearStart, sYearEnd, dWaterImpactYTD, dWaterYTD, dRoadYTD, dRecYTD )
'--------------------------------------------------------------------------------------------------
Sub GetYTDTotals( ByVal sYearStart, ByVal sYearEnd, ByRef dWaterImpactYTD, ByRef dWaterYTD, ByRef dRoadYTD, ByRef dRecYTD )
	Dim sSql, oRs

	sSql = "SELECT YEAR(J.paymentdate) AS paymentyear, "
	sSql = sSql & " R.iswaterimpact, R.iswatermeter, R.isroadimpact, R.isrecreationimpact, "
	sSql = sSql & " SUM(AL.amount) AS amount "
	sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitinvoices I, egov_permitfeereportingtypes R, "
	sSql = sSql & " egov_accounts_ledger AL, egov_class_payment J, egov_permits P, egov_permitaddress A "
	sSql = sSql & " WHERE II.invoiceid = I.invoiceid AND I.orgid = " & session("orgid") & " AND I.isvoided = 0 AND J.paymentid = AL.paymentid "
	sSql = sSql & " AND A.permitid = I.permitid AND (iswaterimpact = 1 OR iswatermeter = 1 OR isroadimpact = 1 OR isrecreationimpact = 1) "
	sSql = sSql & " AND P.isvoided = 0 AND P.permitid = I.permitid AND II.permitfeeid = AL.permitfeeid AND "
	sSql = sSql & " I.invoiceid = AL.invoiceid AND R.feereportingtypeid = II.feereportingtypeid AND "
	sSql = sSql & " (J.paymentdate >= '" & sYearStart & "' AND J.paymentdate < '" & sYearEnd & "' ) "
	sSql = sSql & " GROUP BY YEAR(J.paymentdate), R.iswaterimpact, R.iswatermeter, R.isroadimpact, R.isrecreationimpact"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If oRs("iswaterimpact") Then 
				dWaterImpactYTD = CDbl(oRs("amount"))
			End If
			If oRs("iswatermeter") Then 
				dWaterYTD = CDbl(oRs("amount"))
			End If 
			If oRs("isroadimpact") Then
				dRoadYTD = CDbl(oRs("amount"))
			End If 
			If oRs("isrecreationimpact") Then 
				dRecYTD = CDbl(oRs("amount"))
			End If 
			oRs.MoveNext
		Loop 
	Else
		dWaterImpactYTD = CDbl(0.00)
		dWaterYTD = CDbl(0.00)
		dRoadYTD = CDbl(0.00)
		dRecYTD = CDbl(0.00)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
