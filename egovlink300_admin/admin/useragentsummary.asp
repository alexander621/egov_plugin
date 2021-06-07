<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: useragentsummary.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the setup of mobile features for clients
'
' MODIFICATION HISTORY
' 1.0   04/14/2011   Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sAppSide, iLogOrgId, sWhere, sGroupBy, sSelect, sPlatform
Dim sStartDate, sEndDate, dEndDate

sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

sWhere = ""
sGroupBy = ", applicationside"
sSelect = ", applicationside"

If request("startdate") <> "" Then
	sStartDate = request("startdate")
Else
	' Find the start of the month
	sStartDate = Month(Date()) & "/01/" & Year(Date())
End If 
sWhere = " AND logdate >= '" & sStartDate + "' "

If request("enddate") <> "" Then
	sEndDate = request("enddate")
	dEndDate = CDate(sEndDate)
Else
	' Find the end of the month
	dEndDate = CDate( Month(Date()) & "/01/" & Year(Date()) )
	dEndDate = DateAdd( "m", 1, dEndDate )
	dEndDate = DateAdd( "d", -1, dEndDate )
	sEndDate = FormatDateTime( dEndDate, 2 )
End If 
' push the end date out a day so we can use the < operator and get everything wanted
dEndDate = DateAdd( "d", 1, dEndDate )
sWhere = sWhere & " AND logdate < '" & FormatDateTime( dEndDate ) & "' "

If request("appside") = "" Then
	sAppSide = "combined"
Else
	sAppSide = request("appside")
	If sAppSide <> "separate" And sAppSide <> "combined" Then 
		sWhere = sWhere & " AND applicationside = '" & sAppSide & "'"
	End If 
End If
If sAppSide = "combined" Then
	sGroupBy = ""
	sSelect = ", 'Combined' AS applicationside"
End If 

If request("orgid") = "" Then
	iLogOrgId = 0
Else
	iLogOrgId = request("orgid")
End If

If CLng(iLogOrgId) > CLng(0) Then
	sWhere = sWhere & " AND orgid = " & iLogOrgId
End If 

If request("platform") <> "" Then
	sPlatform = request("platform")
	Select Case sPlatform
		Case "mobile"
			sWhere = sWhere & " AND G.ismobile = 1 "
		Case "desktop"
			sWhere = sWhere & " AND G.ismobile = 0 "
	End Select 
End If 

%>

<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov User Agent Summary</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.5.2.min.js"></script>
	
	<script language="Javascript">
	<!--

		function RefreshResults()
		{
			document.frmLogSearch.submit();
		}

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			var sSelectedDate = $("#" + sField).val();

			// Set the end date to the start date
			if (sField == "enddate")
			{
				sSelectedDate = $("#startdate").val();
			}

			if (sSelectedDate == '')
			{
				// This is today's date
				sSelectedDate = new Date();
				var month = sSelectedDate.getMonth() + 1;
				var day = sSelectedDate.getDate();
				var year = sSelectedDate.getFullYear();
				sSelectedDate = month + "/" + day + "/" + year;
			}

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmLogSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}
	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>User Agent Summary</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmLogSearch" method="post" action="useragentsummary.asp">
							<table cellpadding="2" cellspacing="0" border="0" id="useragentsummarypicks">
								<tr>
									<td>Date Range:</td>
									<td nowrap="nowrap">From: <input type="text" id="startdate" name="startdate" value="<%=sStartDate%>" size="10" maxlength="10" />
										&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
									To: <input type="text" id="enddate" name="enddate" value="<%=sEndDate%>" size="10" maxlength="10" />
										&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
									</td>
								</tr>
								<tr>
									<td>Platform:</td><td><% DisplayPlatforms sPlatform %></td>
								</tr>
								<tr>
									<td>Application Side:</td><td><% DisplayAppSides sAppSide %></td>
								</tr>
								<tr>
									<td>Client:</td><td><% ShowOrgDropDown iLogOrgId %></td>
								</tr>
								<tr><td colspan="2">&nbsp;</td></tr>
								<tr>
			    					<td colspan="2"><input class="button" type="button" value="Refresh Results" onclick="RefreshResults();" /> 
									<!--&nbsp; &nbsp; <input type="button" class="button" value="Download to Excel" onClick="location.href='pagelogsummaryexport.asp'" /> -->
									</td>
  								</tr>
								</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

<%			ShowUserAgentSummary sWhere, sGroupBy, sSelect	%>

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
' void ShowUserAgentSummary sWhere, sGroupBy, sSelect
'--------------------------------------------------------------------------------------------------
Sub ShowUserAgentSummary( ByVal sWhere, ByVal sGroupBy, ByVal sSelect )
	Dim sSql, oRs, iRowCount, iTotalPageViews, iMaxLoadTime, iAvgLoadTime

	iTotalPageViews = CLng(0)
	iMaxLoadTime = CDbl(0.00)
	iAvgLoadTime = CDbl(0.00)
	sSql = "SELECT G.useragentdisplay, CASE G.ismobile WHEN 1 THEN 'Mobile' ELSE 'Desktop' END AS platform" & sSelect
	sSql = sSql & ", SUM(S.pageviews) AS pageviews, MAX(S.maxloadtime) AS maxloadtime, AVG(S.avgloadtime) AS avgloadtime "
	sSql = sSql & "FROM egov_pagelog_useragentgroup_summary S, UserAgent_Groups G "
	sSql = sSql & "WHERE S.useragentgroup = G.useragentgroup" & sWhere 
	sSql = sSql & " GROUP BY G.useragentdisplay, G.ismobile" & sGroupBy 
	sSql = sSql & " ORDER BY G.useragentdisplay, G.ismobile" & sGroupBy
	'response.write sSql & "<br /><br />"
	'response.End 
	session("sSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
%>
		
		<table id="useragentsummary" cellpadding="1" cellspacing="0" border="0" class="sortable">
			<tr><th>User Agent</th><th>Browser/Device<br />Platform</th><th>E-Gov Application<br />Side</th><th>Page<br />Views</th><th>Max Load<br />Time(secs)</th><th>Avg Load<br />Time(secs)</th></tr>
<%		
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write "<td align=""left"">&nbsp;" & oRs("useragentdisplay") & "</td>"
			response.write "<td align=""center"">" & oRs("platform") & "</td>"
			response.write "<td align=""center"">" & oRs("applicationside") & "</td>"
			response.write "<td align=""center"">" & FormatNumber(oRs("pageviews"),0) & "</td>"
			iTotalPageViews = iTotalPageViews + CLng(oRs("pageviews"))
			response.write "<td align=""center"">" & FormatNumber(oRs("maxloadtime"),3,,,0) & "</td>"
			If iMaxLoadTime < CDbl(oRs("maxloadtime")) Then
				iMaxLoadTime = CDbl(oRs("maxloadtime"))
			End If 
			response.write "<td align=""center"">" & FormatNumber(oRs("avgloadtime"),3,,,0) & "</td>"
			iAvgLoadTime = iAvgLoadTime + CDbl(FormatNumber(oRs("avgloadtime"),3,,,0))
			response.write "</tr>"
			oRs.MoveNext 
		Loop

		' Show a Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">"
		response.write "<td align=""center"" colspan=""3"">Totals</td>"
		response.write "<td align=""center"">" & FormatNumber(iTotalPageViews,0) & "</td>"
		response.write "<td align=""center"">" & iMaxLoadTime & "</td>"
		iAvgLoadTime = iAvgLoadTime / iRowCount
		response.write "<td align=""center"">" & FormatNumber(iAvgLoadTime,3,,,0) & "</td>"
		response.write "</tr>"
%>		
		</table>
<%
	Else
		response.write "<p><strong>No information could be found for the selected search options.</strong></p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DisplayAppSides sAppSide 
'--------------------------------------------------------------------------------------------------
Sub DisplayAppSides( ByVal sAppSide )
	response.write vbcrlf & "<select name=""appside"">"

	response.write vbcrlf & "<option value=""combined"""
	If sAppSide = "combined" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">All Sides Combined</option>"

	response.write vbcrlf & "<option value=""separate"""
	If sAppSide = "separate" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">All Sides Seperated</option>"

	response.write vbcrlf & "<option value=""admin"""
	If sAppSide = "admin" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Admin Only</option>"
	
	response.write vbcrlf & "<option value=""mobile"""
	If sAppSide = "mobile" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Mobile Only</option>"

	response.write vbcrlf & "<option value=""public"""
	If sAppSide = "public" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Public Only</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayPlatforms sPlatform
'--------------------------------------------------------------------------------------------------
Sub DisplayPlatforms( ByVal sPlatform )
	response.write vbcrlf & "<select name=""platform"">"

	response.write vbcrlf & "<option value=""all"""
	If sPlatform = "all" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">All Browser/Device Platforms</option>"

	response.write vbcrlf & "<option value=""mobile"""
	If sPlatform = "mobile" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Mobile Only</option>"

	response.write vbcrlf & "<option value=""desktop"""
	If sPlatform = "desktop" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Desktop Only</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOrgDropDown iOrgId 
'--------------------------------------------------------------------------------------------------
Sub  ShowOrgDropDown( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT orgname, orgcity, orgid, defaultstate FROM organizations ORDER BY orgcity, defaultstate "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""orgid"">"
		response.write vbcrlf & vbtab & "<option value=""0"">All Clients</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("orgid") & """ "
			If CLng(iOrgId) = CLng(oRs("orgid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("orgcity") & ", " & oRs("defaultstate") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 



%>
