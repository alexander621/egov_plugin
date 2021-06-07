<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: monthlypermitreport.asp
' AUTHOR: Steve Loar
' CREATED: 07/09/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Monthly Report of permits issued - For clients other than Loveland, OH
'
' MODIFICATION HISTORY
' 1.0   07/09/2009	Steve Loar - INITIAL VERSION - Taken from permitmonthlyreport.asp
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStartMonth, iStartYear, iEndMonth, iEndYear, sEndDate, sStartDate, sYearStart, sYearEnd, iInclude
Dim sDisplayDateRange, sMonthName, iIncludeJobValue, bUsePermitJobValue, iPermitCategoryId, sSearch
Dim sPermitLocation

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "monthly permit report", sLevel	' In common.asp

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

If request("includejobvalue") <> "" Then
	iIncludeJobValue = request("includejobvalue")
	If clng(iIncludeJobValue) = clng(0) Then 
		bUsePermitJobValue = False 
	Else
		bUsePermitJobValue = True 
	End If 
Else
	iIncludeJobValue = 1
	bUsePermitJobValue = False 
End If

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = request("permitcategoryid")
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSearch = " AND P.permitcategoryid = " & iPermitCategoryId
	Else
		sSearch = ""
	End If 
Else 
	iPermitCategoryId = "0"
	sSearch = ""
End If 

'If request("permitlocation") <> "" Then
'	sPermitLocation = request("permitlocation")
'	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
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

		var doCalendar = function( sField ) {
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		};

		var validate = function() {
			document.frmPermitSearch.action = "monthlypermitreport.asp";
			document.frmPermitSearch.submit();
		};
		
		var exportMonthlyReport = function() {
			document.frmPermitSearch.action = "monthlypermitreportexport.asp";
			document.frmPermitSearch.submit();
		};

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
						<form name="frmPermitSearch" method="post" action="monthlypermitreport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="5" cellspacing="0" border="0">
								<tr>
									<td nowrap="nowrap">Permit Category: 
									<%	ShowPermitCategoryPicks iPermitCategoryId	' in permitcommonfunctions.asp	%></td>
								</tr>
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
								</tr>
									<td>
										Include: &nbsp; 
										<select name="includejobvalue">
											<option value="0"
<%												If clng(iIncludeJobValue) = clng(0) Then response.write " selected=""selected"" "	%>
											>Incremental Job Values as Estimated Cost</option>
											<option value="1"
<%											If clng(iIncludeJobValue) = clng(1) Then response.write " selected=""selected"" "	%>
											>Final Permit Job Values as Estimated Cost</option>
										</select>
									</td>
								</tr>
								<tr>
									<td>
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										If request("isview") <> "" Then		%>
											<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="exportMonthlyReport();" />
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
				DisplayIssuedPermits sStartDate, sEndDate, iInclude, sMonthName, bUsePermitJobValue, sSearch
			Else 
				response.write "<strong>To view the monthly permits report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
' void ShowReportYear iStartYear 
'--------------------------------------------------------------------------------------------------
Sub ShowReportYear( ByVal iStartYear )
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
' void ShowReportMonth iStartMonth 
'--------------------------------------------------------------------------------------------------
Sub ShowReportMonth( ByVal iStartMonth )
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
' void DisplayIssuedPermitsl sStartDate, sEndDate, iInclude, sMonthName, bUsePermitJobValue, sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayIssuedPermits( ByVal sStartDate, ByVal sEndDate, ByVal iInclude, ByVal sMonthName, ByVal bUsePermitJobValue, ByVal sSearch )
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

	sSql = "SELECT P.permitid, P.issueddate, P.jobvalue, P.isvoided, ISNULL(P.residentialunits,0) AS residentialunits, "
	sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 0 AS isold, U.reportgroup, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, O.occupancytype "
	sSql = sSql & " FROM egov_permits P  "
	sSql = sSql & " INNER JOIN egov_permitlocationrequirements R ON P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " INNER JOIN egov_permitaddress A ON A.permitid = P.permitid "
	sSql = sSql & " INNER JOIN egov_permitusetypes U ON P.usetypeid = U.usetypeid "
	sSql = sSql & " LEFT JOIN egov_occupancytypes O ON P.occupancytypeid = O.occupancytypeid "
	sSql = sSql & " WHERE P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "' "
	sSql = sSql & " AND P.orgid = " & session("orgid")
	sSql = sSql & sIsVoided & sSearch
	sSql = sSql & " UNION ALL "
	sSql = sSql & " SELECT DISTINCT P.permitid, P.issueddate, P.jobvalue, P.isvoided, 0 AS residentialunits, "
	sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 1 AS isold, U.reportgroup, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, O.occupancytype "
	sSql = sSql & " FROM egov_permits P  "
	sSql = sSql & " INNER JOIN egov_permitlocationrequirements R ON P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " INNER JOIN egov_permitusetypes U ON P.usetypeid = U.usetypeid "
	sSql = sSql & " INNER JOIN egov_permitaddress A ON A.permitid = P.permitid "
	sSql = sSql & " INNER JOIN egov_permitinvoices I ON I.permitid = P.permitid "
	sSql = sSql & " LEFT JOIN egov_occupancytypes O ON P.occupancytypeid = O.occupancytypeid "
	sSql = sSql & " WHERE I.invoicedate > P.issueddate AND P.issueddate < '" & sStartDate & "' "
	sSql = sSql & " AND I.invoicedate >= '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND P.orgid = " & session("orgid")
	sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0 " & sIsVoided & sSearch
	sSql = sSql & " ORDER BY U.reportgroup, P.issueddate, P.permitid, isold"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
		response.write vbcrlf & "<tr><th>Permit #</th><th>Issued<br />Date</th><th>Description of Work</th><th>Address/Location</th><th>Occupancy<br />Type</th><th>New Residential<br />Units</th><th>Est. Cost</th>"
		response.write "<th>Permit Fees</th></tr>"
		Do While Not oRs.EOF
			If sReportGroup <> oRs("reportgroup") Then
				If sReportGroup <> "None" Then
					' Print out a subTotalLine
					response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"">&nbsp;</td>"
					response.write "<td align=""center"">" & iSubTotalUnits & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dCostEstimateSubTotal,2) & "</td>"
					response.write "<td align=""right"">" & FormatNumber(dPermitSubTotal,2) & "</td>"
					response.write "</tr>"
					dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
					iTotalUnits = iTotalUnits + iSubTotalUnits
					dPermitTotal = dPermitTotal + dPermitSubTotal
					sClass = " class=""reportgrouprow"" "
				End If 
				iRowCount = 1
				sReportGroup = oRs("reportgroup")
				response.write vbcrlf & "<tr" & sClass & "><td colspan=""11""><strong>" & sReportGroup & "</strong></td></tr>"
				iSubTotalUnits = CLng(0)
				dCostEstimateSubTotal = CDbl(0.00)
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

			'response.write "<td>" & oRs("permitaddress") & "</td>"
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
			response.write "<td align=""center"">" & oRs("occupancytype") & "</td>"

			response.write "<td align=""center"">" & oRs("residentialunits") & "</td>"
			iSubTotalUnits = iSubTotalUnits + CLng(oRs("residentialunits"))
			
			response.write "<td align=""right"" nowrap=""nowrap"">"
			If bUsePermitJobValue Then
				response.write FormatNumber(oRs("jobvalue"),2)
				dCostEstimateSubTotal = dCostEstimateSubTotal + CDbl(oRs("jobvalue"))
			Else 
				response.write GetCostEstimate( oRs("permitid"), oRs("isold"), dCostEstimateSubTotal, sEndDate, sStartDate )
			End If 
			response.write "</td>"
			response.write "<td align=""right"" nowrap=""nowrap"">" & GetAllPermitFees( oRs("permitid"), oRs("isold"), dPermitSubTotal, sEndDate, sStartDate ) & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
		' Print out the last subTotalLine
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"">&nbsp;</td>"
		response.write "<td align=""center"">" & iSubTotalUnits & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCostEstimateSubTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dPermitSubTotal,2) & "</td>"
		response.write "</tr>"

		dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
		iTotalUnits = iTotalUnits + iSubTotalUnits
		dPermitTotal = dPermitTotal + dPermitSubTotal

		' Print out the TotalLine
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"">" & sMonthName & " Totals</td>"
		response.write "<td align=""center"">" & iTotalUnits & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCostEstimateTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dPermitTotal,2) & "</td>"
		response.write "</tr>"

		' Get YTDs and calculate the Previous Totals
		iYTDUnits = GetYTDResidentialUnits( sYearStart, sYearEnd, iInclude )
		iPreviousUnits = CLng(iYTDUnits) - CLng(iTotalUnits)

		If bUsePermitJobValue Then
			dYTDCostEstimate = GetYTDJobValues( sYearStart, sYearEnd, iInclude )
		Else 
			dYTDCostEstimate = GetYTDCostEstimate( sYearStart, sYearEnd, iInclude )
		End If 
		dPreviousCostEstimate = FormatNumber(CDbl(dYTDCostEstimate) - CDbl(dCostEstimateTotal),2)

		dYTDPermitFees = GetAllYTDPermitFees( sYearStart, sYearEnd, iInclude )
		dPreviousPermitFees = FormatNumber(CDbl(FormatNumber(CDbl(dYTDPermitFees) - CDbl(dPermitTotal),2,,,0)),2)

		' Print out the Previous Totals Line
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"">Previous Totals</td>"
		response.write "<td align=""center"">" & iPreviousUnits & "</td>"
		response.write "<td align=""right"">" & dPreviousCostEstimate & "</td>"
		response.write "<td align=""right"">" & dPreviousPermitFees & "</td>"
		response.write "</tr>"

		' Print out the YTD Line
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"">Year To Date</td>"
		response.write "<td align=""center"">" & iYTDUnits & "</td>"
		response.write "<td align=""right"">" & dYTDCostEstimate & "</td>"
		response.write "<td align=""right"">" & dYTDPermitFees & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table></div>"
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
