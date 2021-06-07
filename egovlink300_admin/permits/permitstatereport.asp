<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitstatereport.asp
' AUTHOR: Steve Loar
' CREATED: 12/01/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Monthly Report of permits issued and additional fees
'
' MODIFICATION HISTORY
' 1.0   12/01/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStartMonth, iStartYear, iEndMonth, iEndYear, sEndDate, sStartDate, sYearStart, sYearEnd
Dim iResidentialCount, dResidentialValue, iCommercialCount, dCommercialValue, iResidentialInspections
Dim iCommercialInspections, iSelectMonth, sReportTitle, sFeatureCheck, dResSqFt, dCommSqFt

If request("rpt") = "" Then
	sReportTitle = "State Report Information"
	sFeatureCheck = "permit state report"
Else
	sReportTitle = "Summary Report"
	sFeatureCheck = "permit summary report"
End If 

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK and feature availability check
PageDisplayCheck sFeatureCheck, sLevel	' In common.asp

If request("selmonth") <> "" Then
	If clng(request("selmonth")) > clng(0) Then 
		iStartMonth = clng(request("selmonth"))
		iSelectMonth = iStartMonth
	Else
		' whole year selected
		iStartMonth = clng(1) 
		iSelectMonth = 0
	End If 
Else
	iStartMonth = clng(Month(Date))   ' this month
	iSelectMonth = iStartMonth
End If 
If request("selyear") <> "" Then
	iStartYear = clng(request("selyear"))
Else
	iStartYear = clng(Year(Date)) ' This Year
End If 
sStartDate = iStartMonth & "/01/" & iStartYear
sYearStart = "01/01/" & iStartYear

If iStartMonth < clng(12) Then
	If clng(request("selmonth")) > clng(0) Then 
		iEndMonth = iStartMonth + 1
		iEndYear = iStartYear
	Else
		iEndMonth = 1		' if whole year selected set to start of next year
		iEndYear = iStartYear + 1
	End If 
Else
	iEndMonth = 1
	iEndYear = iStartYear + 1
End If 
sEndDate = iEndMonth & "/01/" & iEndYear
sYearEnd = sEndDate

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
				<font size="+1"><strong><%=sReportTitle%></strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="reportselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitstatereport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<input type="hidden" id="rpt" name="rpt" value="<%=request("rpt")%>" />
							<table cellpadding="5" cellspacing="0" border="0">
								<tr>
									<td nowrap="nowrap">
										Issued Month: <% ShowReportMonth iSelectMonth	%>
									</td>
								</tr>
								<tr>
									<td nowrap="nowrap">
										Issued Year: <% ShowReportYear iStartYear		%>
									</td>
								</tr>
								<tr>
									<td>
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										If request("isview") <> "" Then		%>
											<!-- <input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="location.href='permitstatereportexport.asp?selmonth=<%=iStartMonth%>&selyear=<%=iStartYear%>'" /> -->
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
<%			' if they choose to view the report, then display the data
			If request("isview") <> "" Then		
				GetPermitCountAndValue sStartDate, sEndDate, "isresidential", iResidentialCount, dResidentialValue , dResSqFt
				'on error resume next
				if session("orgid") = 8 then
					dResSqFt = FormatNumber(GetResidentialNewDwellingSQFT(sStartDate, sEndDate),0)
				end if
				'on error goto 0
				iResidentialInspections = GetPermitInspectionCount( sStartDate, sEndDate, "isresidential" )
				GetPermitCountAndValue sStartDate, sEndDate, "iscommercial", iCommercialCount, dCommercialValue , dCommSqFt
				iCommercialInspections = GetPermitInspectionCount( sStartDate, sEndDate, "iscommercial" )
			%>
				<div class="shadow">
					<table cellpadding="0" cellspacing="0" border="0" class="tableadmin" id="statereport">
						<tr><th>Permit Type</th><th>Permits<br />Issued</th><th>Inspections<br />Made</th><th>Total Sq Ft</th><th>Total<br />Valuation</th></tr>
						<tr>
							<td class="firstcol">Residential</td>
							<td align="center"><%=iResidentialCount%></td>
							<td align="center"><%=iResidentialInspections%></td>
							<td align="center"><%=dResSqFt%></td>
							<td align="center"><%=dResidentialValue%></td>
						</tr>
						<tr class="altrow">
							<td class="firstcol">Commercial</td>
							<td align="center"><%=iCommercialCount%></td>
							<td align="center"><%=iCommercialInspections%></td>
							<td align="center"><%=dCommSqFt%></td>
							<td align="center"><%=dCommercialValue%></td>
						</tr>
					</table>
				</div>
				
				<br />&nbsp;<br />

				<div id="stategroupscountshadow">
					<table cellpadding="0" cellspacing="0" border="0" id="stategroupscount">
						<tr><th>Use Group</th><th>Permits Issued</th></tr>
<%						ShowUseGroupCounts sStartDate, sEndDate		%>
					</table>
				</div>
				
<%			Else 
				response.write "<strong>To view the " & LCase(sReportTitle) & ", select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
	response.write vbcrlf & "<option value=""0"""
	If clng(iStartMonth) = clng(0) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Entire Year</option>"
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
' Sub GetPermitCountAndValue( sStartDate, sEndDate, sUseType, iPermitCount, dTotalValue )
'--------------------------------------------------------------------------------------------------
Sub GetPermitCountAndValue( ByVal sStartDate, ByVal sEndDate, ByVal sUseType, ByRef iPermitCount, ByRef dTotalValue, ByRef dTotalSqFt )
	Dim sSql, oRs
	
	' sUseType = isresidential, or iscommercial
	sSql = "SELECT COUNT(permitid) AS hits, ISNULL(SUM(P.jobvalue),0.00) AS totalvalue, ISNULL(SUM(totalsqft),0) as totalsqft "
	sSql = sSql & " FROM egov_permits P, egov_permitusetypes U "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.isvoided = 0 AND P.isonhold = 0 " 'AND P.issueddate IS NOT NULL "
	sSql = sSql & " AND P.usetypeid = U.usetypeid AND U." & sUseType & " = 1 "
	sSql = sSql & " AND P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "'"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iPermitCount = oRs("hits")
		dTotalValue = FormatNumber(oRs("totalvalue"),0)
		dTotalSqFt = FormatNumber(oRs("totalSqFt"),0)
	Else
		iPermitCount = 0
		dTotalValue = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' Function GetPermitInspectionCount(  sStartDate, sEndDate, sUseType )
'--------------------------------------------------------------------------------------------------
Function GetPermitInspectionCount(  ByVal sStartDate, ByVal sEndDate, ByVal sUseType )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitinspectionid) AS hits "
	sSql = sSql & " FROM egov_permitinspections I, egov_permits P, egov_permitusetypes U "
	sSql = sSql & " WHERE I.permitid = P.permitid AND P.orgid = " & session("orgid")
	sSql = sSql & " AND P.isvoided = 0 AND P.isonhold = 0 AND P.usetypeid = U.usetypeid AND U." & sUseType & " = 1 "
	sSql = sSql & " AND I.inspecteddate >= '" & sStartDate & "' AND I.inspecteddate < '" & sEndDate & "'"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitInspectionCount = oRs("hits")
	Else
		GetPermitInspectionCount = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

Function GetResidentialNewDwellingSQFT(  ByVal sStartDate, ByVal sEndDate )
	Dim sSql, oRs
	
	' sUseType = isresidential, or iscommercial
	sSql = "SELECT SUM(totalsqft) as ressqft "
	sSql = sSql & " from egov_permits P "
	sSql = sSql & " WHERE permittypeid = 17 AND P.isvoided = 0 AND P.isonhold = 0 "
	sSql = sSql & " AND P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "'"
	'response.write sSql & "<br /><br />"
	'response.flush

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		if isnull(oRs("ressqft")) then
			GetResidentialNewDwellingSQFT = 0
		else
			GetResidentialNewDwellingSQFT = oRs("ressqft")
		end if
	Else
		GetResidentialNewDwellingSQFT = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function


'--------------------------------------------------------------------------------------------------
' Sub ShowUseGroupCounts( sStartDate, sEndDate )
'--------------------------------------------------------------------------------------------------
Sub ShowUseGroupCounts( ByVal sStartDate, ByVal sEndDate )
	Dim sSql, oRs, iRowCount

	iRowCount = clng(0)

	sSql = "SELECT occupancytypeid, usegroupcode "
	sSql = sSql & " FROM egov_occupancytypes WHERE orgid = " & session("orgid") & " ORDER BY usegroupcode"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			response.write "<td align=""center"">"
			response.write oRs("usegroupcode")
			response.write "</td><td align=""center"">"
			response.write GetPermitCountByOccupancyTypeId( oRs("occupancytypeid"), sStartDate, sEndDate )
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetPermitCountByOccupancyTypeId( iOccupancyTypeId, sStartDate, sEndDate )
'--------------------------------------------------------------------------------------------------
Function GetPermitCountByOccupancyTypeId( ByVal iOccupancyTypeId, ByVal sStartDate, ByVal sEndDate )
	Dim sSql, oRs

	sSql = "SELECT COUNT(P.permitid) AS hits FROM egov_permits P, egov_permitusetypes U "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.isvoided = 0 AND P.isonhold = 0 "
	'sSql = sSql & " AND P.usetypeid = U.usetypeid AND U.iscommercial = 1 AND P.occupancytypeid = " & iOccupancyTypeId
	sSql = sSql & " AND P.usetypeid = U.usetypeid AND (U.iscommercial = 1 OR U.isresidential = 1) AND P.occupancytypeid = " & iOccupancyTypeId
	sSql = sSql & " AND P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitCountByOccupancyTypeId = FormatNumber(oRs("hits"),0)
	Else
		GetPermitCountByOccupancyTypeId = 0
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 



%>
