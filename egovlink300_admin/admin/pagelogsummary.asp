<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: pagelogsummary.asp
' AUTHOR: Steve Loar
' CREATED: 10/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of page stats
'
' MODIFICATION HISTORY
' 1.0   10/15/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMonth, iYear, sAppSide, iLogOrgId, sWhere, sGroupBy, sSelect, sTodayWhere

sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

sWhere = ""
sTodayWhere = ""
sGroupBy = ", applicationside"
sSelect = ", applicationside"

If request("month") = "" Then
	iMonth = Month(Date)
Else
	iMonth = request("month")
End If 

sWhere = " WHERE MONTH(logdate) = " & iMonth

If request("year") = "" Then
	iYear = Year(Date)
Else
	iYear = request("year")
End If

sWhere = sWhere & " AND YEAR(logdate) = " & iYear


If request("appside") = "" Then
	sAppSide = "combined"
Else
	sAppSide = request("appside")
	If sAppSide <> "both" And sAppSide <> "combined" Then 
		sWhere = sWhere & " AND applicationside = '" & sAppSide & "'"
		sTodayWhere = sTodayWhere & " AND applicationside = '" & sAppSide & "'"
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
	sTodayWhere = sTodayWhere & " AND orgid = " & iLogOrgId
End If 

%>

<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Page Log Summary</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />
	
	<script language="Javascript">
	<!--

		function RefreshResults()
		{
			document.frmLogSearch.submit();
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
				<font size="+1"><strong>Page Log Summary</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmLogSearch" method="post" action="pagelogsummary.asp">
							<table cellpadding="2" cellspacing="0" border="0" id="pagelogpicks">
								<tr>
									<td>Month:</td><td><% DisplayMonths iMonth %></td>
								</tr>
								<tr>
									<td>Year:</td><td><% DisplayYears iYear %></td>
								</tr>
								<tr>
									<td>Application Side:</td><td><% DisplayAppSides sAppSide %></td>
								</tr>
								<tr>
									<td>Organization:</td><td><% ShowOrgDropDown iLogOrgId %></td>
								</tr>
								<tr><td colspan="2">&nbsp;</td></tr>
								<tr>
			    					<td colspan="2"><input class="button" type="button" value="Refresh Results" onclick="RefreshResults();" /> &nbsp; &nbsp; <input type="button" class="button" value="Download to Excel" onClick="location.href='pagelogsummaryexport.asp'" /></td>
  								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

<%			ShowPageLogSummary sWhere, sGroupBy, sSelect, iMonth, iYear, sTodayWhere	%>
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
' void ShowPageLogSummary sWhere, sGroupBy, sSelect, iMonth, iYear, sTodayWhere
'--------------------------------------------------------------------------------------------------
Sub ShowPageLogSummary( ByVal sWhere, ByVal sGroupBy, ByVal sSelect, ByVal iMonth, ByVal iYear, ByVal sTodayWhere )
	Dim sSql, oRs, iRowCount, iTotalPageViews, iMaxLoadTime, iAvgLoadTime

	iTotalPageViews = CLng(0)
	iMaxLoadTime = CDbl(0.00)
	iAvgLoadTime = CDbl(0.00)
	sSql = "SELECT logdate" & sSelect & ", SUM(pageviews) AS pageviews, MAX(maxloadtime) AS maxloadtime, AVG(avgloadtime) AS avgloadtime "
	sSql = sSql & " FROM egov_pagelog_summary " & sWhere 
	sSql = sSql & " GROUP BY logdate" & sGroupBy & " ORDER BY logdate" & sGroupBy
	'response.write sSql & "<br />"
	session("sSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
%>
		
		<table id="pagelogsummary" cellpadding="1" cellspacing="0" border="0" class="sortable">
			<tr><th>Log Date</th><th>Application<br />Side</th><th>Page<br />Views</th><th>Max Load<br />Time(secs)</th><th>Avg Load<br />Time(secs)</th></tr>
<%		
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"
			response.write "<td align=""center"">" & FormatDateTime(oRs("logdate"),2) & "</td>"
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

		' If this is for the current month and year, then get a row for today
		If clng(iMonth) = clng(Month(Date)) And clng(iYear) = clng(Year(Date)) Then
			sSql = "SELECT '" & FormatDateTime(Date,2) & "' AS logdate" & sSelect & ", COUNT(logdate) AS pageviews, MAX(loadtime) AS maxloadtime, AVG(loadtime) AS avgloadtime FROM egov_pagelog "
			sSql = sSql & " WHERE logdate > '" & FormatDateTime(Date,2) & " 00:00:00' and logdate < '" & FormatDateTime(DateAdd("d",1,Date),2) & " 00:00:00' " & sTodayWhere 
			If sGroupBy <> "" Then 
				sSql = sSql & " GROUP BY applicationside"
			End If 
			'response.write sSql & "<br />"
			session("sTodaySql") = sSql
			ShowTodaysLog sSql, iTotalPageViews, iMaxLoadTime, iAvgLoadTime, iRowCount
		Else
			session("sTodaySql") = ""
		End If 

		' Show a Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">"
		response.write "<td align=""center"" colspan=""2"">Totals</td>"
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
' Sub ShowTodaysLog( sSql, iTotalPageViews, iMaxLoadTime, iAvgLoadTime, iRowCount )
'--------------------------------------------------------------------------------------------------
Sub ShowTodaysLog( ByVal sSql, ByRef iTotalPageViews, ByRef iMaxLoadTime, ByRef iAvgLoadTime, ByRef iRowCount )
	Dim oRs

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write ">"
		response.write "<td align=""center"">" & FormatDateTime(oRs("logdate"),2) & "</td>"
		response.write "<td align=""center"">" & oRs("applicationside") & "</td>"
		response.write "<td align=""center"">" & oRs("pageviews") & "</td>"
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

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void DisplayMonths iMonth 
'--------------------------------------------------------------------------------------------------
Sub DisplayMonths( ByVal iMonth )

	response.write vbcrlf & "<select name=""month"">"
	For x = 1 To 12
		response.write vbcrlf & "<option value=""" & x & """"
		If clng(iMonth) = clng(x) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & MonthName(x) & "</option>"
	Next 
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayYears iYear 
'--------------------------------------------------------------------------------------------------
Sub DisplayYears( ByVal iYear )
	Dim iStartYear, iEndYear

	iStartYear = Year(DateAdd("yyyy", -3, Date))
	iEndYear = Year(Date)

	response.write vbcrlf & "<select name=""year"">"
	For x = iStartYear To iEndYear
		response.write vbcrlf & "<option value=""" & x & """"
		If clng(iYear) = clng(x) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & x & "</option>"
	Next 
	response.write vbcrlf & "</select>"

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

	response.write vbcrlf & "<option value=""both"""
	If sAppSide = "both" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">All Sides Seperated</option>"

	response.write vbcrlf & "<option value=""admin"""
	If sAppSide = "admin" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Admin Only</option>"

	response.write vbcrlf & "<option value=""public"""
	If sAppSide = "public" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Public Only</option>"

	response.write vbcrlf & "<option value=""mobile"""
	If sAppSide = "mobile" Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Mobile Only</option>"

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
		response.write vbcrlf & vbtab & "<option value=""0"">All Organizations</option>"
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
