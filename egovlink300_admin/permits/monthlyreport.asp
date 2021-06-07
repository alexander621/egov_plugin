<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitmonthlyreport.asp
' AUTHOR: Steve Loar
' CREATED: 11/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Monthly Report of permits issued and additional fees
'
' MODIFICATION HISTORY
' 1.0   09/09/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStartMonth, iStartYear, iEndMonth, iEndYear, sEndDate, sStartDate, sYearStart, sYearEnd, iInclude
Dim sDisplayDateRange, sMonthName

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "permit monthly report", sLevel	' In common.asp

If request("selmonth") <> "" Then
	iStartMonth = clng(request("selmonth"))
Else
	iStartMonth = clng(Month(Date))   ' this month
End If 

sMonthName = MonthName(iStartMonth)
sDisplayDateRange = MonthName(iStartMonth)

If request("selyear") <> "" Then
	iStartYear = clng(request("selyear"))
Else
	iStartYear = clng(Year(Date)) ' This Year
End If 

sDisplayDateRange = sDisplayDateRange & " " & iStartYear
sStartDate = iStartMonth & "/01/" & iStartYear
sYearStart = "01/01/" & iStartYear

If iStartMonth < clng(12) Then
	iEndMonth = iStartMonth + 1
	iEndYear = iStartYear
Else
	iEndMonth = 1
	iEndYear = iStartYear + 1
End If 
sEndDate = iEndMonth & "/01/" & iEndYear
sYearEnd = sEndDate

If request("include") <> "" Then
	iInclude = request("include")
Else
	iInclude = 0
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="Javascript">
	<!--

		window.onload = function()
		{
		  //factory.printing.header = "Printed on &d"
		  //factory.printing.footer = "&bPrinted on &d - Page:&p/&P";
		  factory.printing.portrait = false;
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

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function validate()
		{
			document.frmPermitSearch.submit();
		}

	//-->
	</script>

</head>

<body>

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

			<!--BEGIN: PAGE TITLE-->
			<p id="pagetitle">
				<span id="printdaterange"><font size="+1"><strong><%=sDisplayDateRange%></strong></font></span>
				<font size="+1"><strong>Monthly Permits Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitmonthlyreport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="5" cellspacing="0" border="0">
								<tr>
									<td nowrap="nowrap">
										Issued Month: <% ShowReportMonth iStartMonth	%>
									</td>
								</tr>
								<tr>
									<td nowrap="nowrap">
										Issued Year: <% ShowReportYear iStartYear		%>
									</td>
								</tr>
								<tr>
									<td>Show: &nbsp;
										<select name="include">
											<option value="0"
<%												If clng(iInclude) = clng(0) Then response.write " selected=""selected"" "	%>
											>Exclude Voided Permits</option>
											<option value="1"
<%												If clng(iInclude) = clng(1) Then response.write " selected=""selected"" "	%>											
											>Only Voided Permits</option>
											<option value="2"
<%												If clng(iInclude) = clng(2) Then response.write " selected=""selected"" "	%>											
											>Include Voided Permits</option>
										</select>
									</td>	
								</tr>
								<tr>
									<td>
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										If request("isview") <> "" Then		%>
											<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="location.href='permitmonthlyreportexport.asp?selmonth=<%=iStartMonth%>&selyear=<%=iStartYear%>'" />
<%										End If		%>
									</td>
								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->
		<!--</div>-->

			<!-- Begin: Report Display -->
<%			' if they choose to view the report, then display the inspections
			If request("isview") <> "" Then	
				DisplayIssuedPermits sStartDate, sEndDate, iInclude, sMonthName
			Else 
				response.write "<strong>To view the monthly permits report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
			End If 
%>
			<!-- END: Report Display -->
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
' Sub ShowReportYear( iStartYear )
'--------------------------------------------------------------------------------------------------
Sub ShowReportYear( iStartYear )
	Dim iFirstYear, iEndYear, x

	' Start with 2008 and keep adding years until current year
	iFirstYear = 2008
	iEndYear = Year(Date)

	response.write vbcrlf & "<select name=""selyear"">"
	For x = iFirstYear To iEndYear
		response.write vbcrlf & "<option value=""" & x & """"
		If CLng(iStartYear) = CLng(x) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & x & "</option>"
	Next 
	response.write vbcrlf & "</select>"

End Sub


'--------------------------------------------------------------------------------------------------
' Sub ShowReportMonth( iStartMonth )
'--------------------------------------------------------------------------------------------------
Sub ShowReportMonth( iStartMonth )
	Dim x 

	response.write vbcrlf & "<select name=""selmonth"">"
	For x = 1 To 12
		response.write vbcrlf & "<option value=""" & x & """"
		If clng(iStartMonth) = clng(x) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & MonthName(x) & "</option>"
	Next 
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DisplayIssuedPermits( ByVal sSearch )
'--------------------------------------------------------------------------------------------------
Sub DisplayIssuedPermits( ByVal sStartDate, ByVal sEndDate, ByVal iInclude, ByVal sMonthName )
	Dim sSql, oRs, iRowCount, sReportGroup, dCostEstimateSubTotal, dCostEstimateTotal, sClass
	Dim iSubTotalUnits, iTotalUnits, dZoneSubTotal, dZoneTotal, dBBSSubTotal, dBBSTotal
	Dim dPenSubTotal, dPenTotal, dCertSubTotal, dCertTotal, dPermitSubTotal, dPermitTotal
	Dim iYTDUnits, iPreviousUnits, dYTDCostEstimate, dPreviousCostEstimate, dYTDPermitFees
	Dim dPreviousPermitFees, dYTDZone, dPreviousZone, dYTDBBS, dPreviousBBS, dYTDPen
	Dim dPreviousPen, dYTDCert, dPreviousCert, sIsVoided

	sReportGroup = "None"
	iRowCount = 0
	dCostEstimateSubTotal = CDbl(0.00)
	dCostEstimateTotal = CDbl(0.00)
	iSubTotalUnits = CLng(0)
	iTotalUnits = CLng(0)
	dZoneSubTotal = CDbl(0.00)
	dZoneTotal = CDbl(0.00)
	dBBSSubTotal = CDbl(0.00)
	dBBSTotal = CDbl(0.00)
	dPenSubTotal = CDbl(0.00)
	dPenTotal = CDbl(0.00)
	dCertSubTotal = CDbl(0.00)
	dCertTotal = CDbl(0.00)
	dPermitSubTotal = CDbl(0.00)
	dPermitTotal = CDbl(0.00)
	sClass = ""

	If clng(iInclude ) < clng(2) Then
		sIsVoided = " AND P.isvoided = " & iInclude
	Else
		sIsVoided = ""
	End If 

	sSql = "SELECT P.permitid, P.issueddate, P.isvoided, ISNULL(P.residentialunits,0) AS residentialunits, "
	sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 0 AS isold, U.reportgroup, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitusetypes U "
	sSql = sSql & " WHERE P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "' "
	sSql = sSql & " AND A.permitid = P.permitid AND P.usetypeid = U.usetypeid AND P.orgid = " & session("orgid")
	sSql = sSql & sIsVoided
	sSql = sSql & " UNION ALL "
	sSql = sSql & " SELECT DISTINCT P.permitid, P.issueddate, P.isvoided, 0 AS residentialunits, "
	sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 1 AS isold, U.reportgroup, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitinvoices I, egov_permitusetypes U "
	sSql = sSql & " WHERE A.permitid = P.permitid AND I.permitid = P.permitid AND I.invoicedate > P.issueddate AND P.issueddate < '" & sStartDate & "' "
	sSql = sSql & " AND I.invoicedate >= '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND P.usetypeid = U.usetypeid AND P.orgid = " & session("orgid")
	sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0 " & sIsVoided
	sSql = sSql & " ORDER BY U.reportgroup, P.issueddate, P.permitid, isold"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
		response.write vbcrlf & "<tr><th>Permit #</th><th>Issued<br />Date</th><th>Scope of Work</th><th>Address</th><th>New Residential<br />Units</th><th>Est. Cost</th>"
		response.write "<th>Permit</th><th>Zone Fee</th><th>BBS Fee</th></tr>"   '<th>Rev./Pen</th><th>C of O</th></tr>"
		Do While Not oRs.EOF
			If sReportGroup <> oRs("reportgroup") Then
				If sReportGroup <> "None" Then
					' Print out a subTotalLine
					response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">&nbsp;</td>"
					response.write "<td align=""center"">" & iSubTotalUnits & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dCostEstimateSubTotal,2) & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dPermitSubTotal,2) & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dZoneSubTotal,2) & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dBBSSubTotal,2) & "</td>"
					'response.write "<td align=""right"">" & FormatNumber(dPenSubTotal,2) & "</td>"
					'response.write "<td align=""right"">" & FormatNumber(dCertSubTotal,2) & "</td>"
					response.write "</tr>"
					dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
					iTotalUnits = iTotalUnits + iSubTotalUnits
					dZoneTotal = dZoneTotal + dZoneSubTotal
					dBBSTotal = dBBSTotal + dBBSSubTotal
					'dPenTotal = dPenTotal + dPenSubTotal
					'dCertTotal = dCertTotal + dCertSubTotal
					dPermitTotal = dPermitTotal + dPermitSubTotal
					sClass = " class=""reportgrouprow"" "
				End If 
				iRowCount = 1
				sReportGroup = oRs("reportgroup")
				response.write vbcrlf & "<tr" & sClass & "><td colspan=""11""><strong>" & sReportGroup & "</strong></td></tr>"
				iSubTotalUnits = CLng(0)
				dCostEstimateSubTotal = CDbl(0.00)
				dZoneSubTotal = CDbl(0.00)
				dBBSSubTotal = CDbl(0.00)
				'dPenSubTotal = CDbl(0.00)
				'dCertSubTotal = CDbl(0.00)
				dPermitSubTotal = CDbl(0.00)
			End If 

			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			
			response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber( oRs("permitid") )
			If oRs("isvoided") Then
				response.write "v"
			End If 
			response.write "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("issueddate"),2) & "</td>"
			response.write "<td>" & oRs("descriptionofwork") & "</td>"
			response.write "<td>" & oRs("permitaddress") & "</td>"
			response.write "<td align=""center"">" & oRs("residentialunits") & "</td>"
			iSubTotalUnits = iSubTotalUnits + CLng(oRs("residentialunits"))
			
			response.write "<td align=""right"" nowrap=""nowrap"">" & GetCostEstimate( oRs("permitid"), oRs("isold"), dCostEstimateSubTotal, sEndDate, sStartDate ) & "</td>"
			response.write "<td align=""right"" nowrap=""nowrap"">" & GetPermitFees( oRs("permitid"), oRs("isold"), dPermitSubTotal, sEndDate, sStartDate ) & "</td>"
			response.write "<td align=""right"" nowrap=""nowrap"">" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "iszone", dZoneSubTotal, sEndDate, sStartDate ) & "</td>"
			response.write "<td align=""right"" nowrap=""nowrap"">" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "isbbs", dBBSSubTotal, sEndDate, sStartDate ) & "</td>"
			'response.write "<td align=""right"" nowrap=""nowrap"">" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "isrevisionpenalty", dPenSubTotal ) & "</td>"
			'response.write "<td align=""right"" nowrap=""nowrap"">" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "iscertofoccupancy", dCertSubTotal ) & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
		' Print out the last subTotalLine
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">&nbsp;</td>"
		response.write "<td align=""center"">" & iSubTotalUnits & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCostEstimateSubTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dPermitSubTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dZoneSubTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dBBSSubTotal,2) & "</td>"
		'response.write "<td align=""right"">" & FormatNumber(dPenSubTotal,2) & "</td>"
		'response.write "<td align=""right"">" & FormatNumber(dCertSubTotal,2) & "</td>"
		response.write "</tr>"
		dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
		iTotalUnits = iTotalUnits + iSubTotalUnits
		dZoneTotal = dZoneTotal + dZoneSubTotal
		dBBSTotal = dBBSTotal + dBBSSubTotal
		'dPenTotal = dPenTotal + dPenSubTotal
		'dCertTotal = dCertTotal + dCertSubTotal
		dPermitTotal = dPermitTotal + dPermitSubTotal

		' Print out the TotalLine
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">" & sMonthName & " Totals</td>"
		response.write "<td align=""center"">" & iTotalUnits & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCostEstimateTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dPermitTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dZoneTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dBBSTotal,2) & "</td>"
		'response.write "<td align=""right"">" & FormatNumber(dPenTotal,2) & "</td>"
		'response.write "<td align=""right"">" & FormatNumber(dCertTotal,2) & "</td>"
		response.write "</tr>"

		' Get YTDs and calculate the Previous Totals
		iYTDUnits = GetYTDResidentialUnits( sYearStart, sYearEnd, iInclude )
		iPreviousUnits = CLng(iYTDUnits) - CLng(iTotalUnits)

		dYTDCostEstimate = GetYTDCostEstimate( sYearStart, sYearEnd, iInclude )
		dPreviousCostEstimate = FormatNumber(CDbl(dYTDCostEstimate) - CDbl(dCostEstimateTotal),2)

		dYTDPermitFees = GetYTDPermitFees( sYearStart, sYearEnd, iInclude )
		dPreviousPermitFees = FormatNumber(CDbl(FormatNumber(CDbl(dYTDPermitFees) - CDbl(dPermitTotal),2,,,0)),2)

		dYTDZone = GetYTDPermitReportingFees( sYearStart, sYearEnd, "iszone", iInclude )
		dPreviousZone = FormatNumber(CDbl(FormatNumber(CDbl(dYTDZone) - CDbl(dZoneTotal),2,,,0)),2)

		dYTDBBS = GetYTDPermitReportingFees( sYearStart, sYearEnd, "isbbs", iInclude )
		dPreviousBBS = FormatNumber(CDbl(FormatNumber(CDbl(dYTDBBS) - CDbl(dBBSTotal),2,,,0)),2)

		'dYTDPen = GetYTDPermitReportingFees( sYearStart, sYearEnd, "isrevisionpenalty", iInclude )
		'dPreviousPen = FormatNumber(CDbl(FormatNumber(CDbl(dYTDPen) - CDbl(dPenTotal),2,,,0)),2)

		'dYTDCert = GetYTDPermitReportingFees( sYearStart, sYearEnd, "iscertofoccupancy", iInclude )
		'dPreviousCert = FormatNumber(CDbl(dYTDCert) - CDbl(dPreviousCert),2)
		'dPreviousCert = FormatNumber(CDbl(FormatNumber(CDbl(dYTDCert) - CDbl(dPreviousCert),2,,,0)),2)

		' Print out the Previous Totals Line
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">Previous Totals</td>"
		response.write "<td align=""center"">" & iPreviousUnits & "</td>"
		response.write "<td align=""right"">" & dPreviousCostEstimate & "</td>"
		response.write "<td align=""right"">" & dPreviousPermitFees & "</td>"
		response.write "<td align=""right"">" & dPreviousZone & "</td>"
		response.write "<td align=""right"">" & dPreviousBBS & "</td>"
		'response.write "<td align=""right"">" & dPreviousPen & "</td>"
		'response.write "<td align=""right"">" & dPreviousCert & "</td>"
		response.write "</tr>"

		' Print out the YTD Line
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""4"">Year To Date</td>"
		response.write "<td align=""center"">" & iYTDUnits & "</td>"
		response.write "<td align=""right"">" & dYTDCostEstimate & "</td>"
		response.write "<td align=""right"">" & dYTDPermitFees & "</td>"
		response.write "<td align=""right"">" & dYTDZone & "</td>"
		response.write "<td align=""right"">" & dYTDBBS & "</td>"
		'response.write "<td align=""right"">" & dYTDPen & "</td>"
		'response.write "<td align=""right"">" & dYTDCert & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table></div>"
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


%>
