<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcensusreport.asp
' AUTHOR: Steve Loar
' CREATED: 12/02/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Monthly Report of permits issued and additional fees
'
' MODIFICATION HISTORY
' 1.0   12/02/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStartMonth, iStartYear, iEndMonth, iEndYear, sEndDate, sStartDate, sYearStart, sYearEnd
Dim iBuildings, iHousingUnits, iValuation, iTotalHousingUnits, iTotalValuation

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK and feature availability check
PageDisplayCheck "permit census report ", sLevel	' In common.asp

iTotalHousingUnits = CLng(0)
iTotalValuation = CLng(0)

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
				<font size="+1"><strong>Census Report Information</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="reportselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitcensusreport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
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
			%>
				<div class="shadow">
					<table cellpadding="0" cellspacing="0" border="0" class="tableadmin" id="statereport">
						<tr><th>Type of Structure</th><th>Item No.</th><th>Buildings</th><th>Housing Units</th><th>Valuation of Construction</th></tr>
<%						GetCensusData sStartDate, sEndDate, "= 1", iBuildings, iHousingUnits, iValuation, iTotalHousingUnits, iTotalValuation			%>
						<tr>
							<td class="firstcol">Single-family Houses</td>
							<td align="center">101</td>
							<td align="center">&mdash;</td>
							<td align="center"><%=iHousingUnits%></td>
							<td align="center"><%=iValuation%></td>
						</tr>
<%						GetCensusData sStartDate, sEndDate, "= 2", iBuildings, iHousingUnits, iValuation, iTotalHousingUnits, iTotalValuation			%>
						<tr class="altrow">
							<td class="firstcol">Two-unit Buildings</td>
							<td align="center">103</td>
							<td align="center"><%=iBuildings%></td>
							<td align="center"><%=iHousingUnits%></td>
							<td align="center"><%=iValuation%></td>
						</tr>
<%						GetCensusData sStartDate, sEndDate, "= 3 OR residentialunits = 4", iBuildings, iHousingUnits, iValuation, iTotalHousingUnits, iTotalValuation			%>
						<tr>
							<td class="firstcol">Three- and Four-unit Buildings</td>
							<td align="center">104</td>
							<td align="center"><%=iBuildings%></td>
							<td align="center"><%=iHousingUnits%></td>
							<td align="center"><%=iValuation%></td>
						</tr>
<%						GetCensusData sStartDate, sEndDate, "> 4", iBuildings, iHousingUnits, iValuation, iTotalHousingUnits, iTotalValuation			%>
						<tr class="altrow">
							<td class="firstcol">Five-or-more Unit Buildings</td>
							<td align="center">105</td>
							<td align="center"><%=iBuildings%></td>
							<td align="center"><%=iHousingUnits%></td>
							<td align="center"><%=iValuation%></td>
						</tr>
						<tr class="totalrow">
							<td class="firstcol">Total - Sum of 101-105</td>
							<td align="center">109</td>
							<td align="center">&mdash;</td>
							<td align="center"><%=iTotalHousingUnits%></td>
							<td align="center"><%=iTotalValuation%></td>
						</tr>
					</table>
				</div>
				
<%			Else 
				response.write "<strong>To view the census report information, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
' void GetCensusData sStartDate, sEndDate, sUnits, iHousingUnits, iValuation, iTotalHousingUnits, iTotalValuation 
'--------------------------------------------------------------------------------------------------
Sub GetCensusData( ByVal sStartDate, ByVal sEndDate, ByVal sUnits, ByRef iBuildings, ByRef iHousingUnits, ByRef iValuation, ByRef iTotalHousingUnits, ByRef iTotalValuation )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitid) AS permits, ISNULL(SUM(residentialunits),0) AS residentialunits, ISNULL(SUM(jobvalue),0.00) AS jobvalue "
	sSql = sSql & " FROM egov_permits P, egov_permitusetypes U "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.isvoided = 0 AND P.isonhold = 0 "
	sSql = sSql & " AND P.usetypeid = U.usetypeid AND (residentialunits " & sUnits & ") "
	sSql = sSql & " AND P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "'"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iBuildings = CLng(oRs("permits"))
		iHousingUnits = CLng(oRs("residentialunits"))
		iValuation = FormatNumber(oRs("jobvalue"),0)
	Else
		iBuildings = 0
		iHousingUnits = 0
		iValuation = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

	iTotalHousingUnits = FormatNumber(CLng(iTotalHousingUnits) + CLng(iHousingUnits),0)
	iTotalValuation = FormatNumber(CLng(iTotalValuation) + CLng(iValuation),0)

End Sub 

%>


