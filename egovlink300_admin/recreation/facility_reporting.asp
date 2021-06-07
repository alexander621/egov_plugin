<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReportID

iReportID = clng(request("reportid"))


If iReportID = clng(1) Then  ' Totals Report
	If Not UserHasPermission( Session("UserId"), "facility totals" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If 
Else
	' Cancellation Report
	If Not UserHasPermission( Session("UserId"), "facility cancellations" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If 
End If 

if request.querystring("excel") = "true" then
sName = Replace(replace(Replace(replace(replace(Now(),":",""),"/",""),"AM",""),"PM","")," ","") ' NAME BASED ON DATETIME STRING
server.scripttimeout = 9000
response.clear
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=EGOV_EXPORT_" & sName & ".xls"
	DisplayReport iReportID 
response.flush
	response.end
end if 
%>
<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->

<%
sLevel = "../" ' Override of value from common.asp
%>
<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="reservation.css" />

	<script src="../scripts/jquery-1.9.1.min.js"></script>

	<script>
	function refreshreport( rptNo ) {
		var iSM = document.frmdate.sm.options[document.frmdate.sm.selectedIndex].value;
		var iSY = document.frmdate.sy.options[document.frmdate.sy.selectedIndex].value;
		var iEM = document.frmdate.em.options[document.frmdate.em.selectedIndex].value;
		var iEY = document.frmdate.ey.options[document.frmdate.ey.selectedIndex].value;

		if (parseInt(rptNo) == 2)
		{
			var facilityid = $("#facilityid").val();
			location.href='facility_reporting.asp?reportid=' + rptNo + '&sm=' +iSM + '&sy=' + iSY + '&em=' + iEM + '&ey=' + iEY + "&facilityid=" + facilityid;
		}
		else
			location.href='facility_reporting.asp?reportid=' + rptNo + '&sm=' +iSM + '&sy=' + iSY + '&em=' + iEM + '&ey=' + iEY;
	}

</script>

</head>
<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 


	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

	<%
			DisplayReport iReportID 
	%>
		</div>
	</div>
	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
' void DISPLAYREPORT(IREPORTID)
'--------------------------------------------------------------------------------------------------
Sub DisplayReport( ByVal iReportID )
	Dim iFacilityId

	Select Case iReportID

		Case 1
			GenerateFacilityTotalReport request("sm"),request("sy"),request("em"),request("ey")

		Case 2
			If request("facilityid") = "" Then
				iFacilityId = CLng(0)
			Else
				iFacilityId = request("facilityid")
			End If 

			GenerateFacilityCancellationReport request("sm"),request("sy"),request("em"),request("ey"), iFacilityId

		Case Else
			response.write "No report specified."

	End Select

End Sub


'--------------------------------------------------------------------------------------------------
' GENERATEFACILITYTOTALREPORT()
'--------------------------------------------------------------------------------------------------
Sub GenerateFacilityTotalReport( ByVal startmonth, ByVal startyear, ByVal endmonth, ByVal endyear )
	Dim oRs, sSql, sTempYear, sTempMonth
	
	' SET DEFAULT DATE IF NONE SUPPLIED
	If startmonth = "" or startyear = "" or endmonth = "" or endyear = "" Then
		startmonth = 1
		startyear = Year(Date())
		endmonth = Month(Date())
		endyear = Year(Date()) 
	End If


	' FLIP DATES IF REVERSED IF STARTYEAR GREATER THAN ENDYEAR
	If (startyear > endyear) Or (startyear = endyear And startmonth > endmonth) Then
		sTempYear = startyear
		sTempMonth	= startmonth
		startyear = endyear
		endyear = sTempYear
		startmonth = endmonth
		endmonth = startmonth
	End If


	' DISPAY MONTH/YEAR SELECTION
	' START MONTH SELECT
	response.write "<form name=""frmdate"" action=# method=""post"">"
	response.write "<p><select name=""sm"">"
	For m = 1 to 12
		sSelected = ""
		If clng(startmonth) = clng(m) Then
			sSelected = " selected=""selected"""
		End If
		response.write "<option" & sSelected & " value="& m &">" & MonthName(m) & "</option>"
	Next
	response.write "</select>"
	

	' START YEAR SELECT
	response.write "<select name=""sy"">"
	For y = Year(Date())-5 to Year(Date())+5
		sSelected = ""
		If clng(startyear) = clng(y) Then
			sSelected = " selected=""selected"""
		End If
		response.write "<option" & sSelected & " value="& y &">" & y & "</option>"
	Next
	response.write "</select> &mdash; "
	

	' END MONTH SELECT
	response.write "<select name=""em"">"
	For m = 1 to 12
		sSelected = ""
		If clng(endmonth) = clng(m) Then
			sSelected = " selected=""selected"""
		End If
		response.write "<option" & sSelected & " value="& m &">" & MonthName(m) & "</option>"
	Next
	response.write "</select>"
	

	' END YEAR SELECT
	response.write "<select name=""ey"">"
	For y = Year(Date())-5 to Year(Date())+5
		sSelected = ""
		If clng(endyear) = clng(y) Then
			sSelected = " selected=""selected"""
		End If
		response.write "<option" & sSelected & "  value="& y &">" & y & "</option>"
	Next
	response.write "</select>&nbsp;&nbsp;&nbsp;<input type=""button"" class=""facilitybutton"" value=""Display Report"" onClick=""refreshreport(1);"" /></form></p>"


	' GET INFORMATION FROM SQL
	' PROCESS SUPPLIED DATE PARAMS
	If startmonth = "" Then
		' NO WHERE CLAUSE
		sWhere = "WHERE orgid='" & session("orgid") & "'"
	Else
		' IF PARAMETERS SUPPLIED ADD WHERE CLAUSE
		sWhere = " WHERE ((Year = '" & clng(startyear) & "' AND Month >= '" & clng(startmonth) & "') OR (Year > '" & clng(startyear) & "' AND Year <= '" & clng(endyear) & "') OR  (year='" & clng(endyear) & "' AND Month <= '" & clng(endmonth) & "')) AND (orgid='" & session("orgid") & "')"
	End If

	
	sSql = "SELECT *, (SELECT SUM(TotalReservations) FROM rpt_facility_totals " 
	sSql = sSql & sWhere & ") AS GTR,(SELECT SUM(TotalAmount) FROM rpt_facility_totals  " 
	sSql = sSql & sWhere & ") AS GTA FROM rpt_facility_totals" & sWhere

	' ORDER BY CLAUSE
	sSql = sSql & " ORDER BY year, month, facilityname"
	'response.write sSql & "<br /><br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	' IF RECORDSET NOT EMPTY DISPLAY RESULTS
	If Not oRs.EOF Then
   
		' GET GRAND TOTALS
		iReserveTotal = oRs("GTR")
		curGrandTotal = FormatCurrency(oRs("GTA"),2)
		iCurrentMonth = clng( oRs("Month"))
		iPreviousMonth = ""
		blnFirst = True
		blnLast = False
		iMonthAmountTotal = 0
		iMonthReserveTotal = 0

		' DISPLAY REPORT HEADER
		response.write "<div style=""COLOR: #000000; font-family: verdana,sans-serif; font-size: 12px; font-weight:bold; margin-bottom: 2em;"">Total Lodge Counts from " & MonthName(startmonth) & " " & startYear & " to " & MonthName(endMonth) & " " & endYear & " </div>"

		response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""75%"">"
		response.write "<tr class=""tablelist""><th >&nbsp;</th><th># of Reservations</th><th >Total Dollars</th></tr>"

		' BEGIN REPORT DETAILS
		Do While Not oRs.EOF
				
				' UPDATE MONTH TOTAL INFORMATION
				iCurrentMonth = oRs("Month")
				
				If iCurrentMonth <> iPreviousMonth Then
					If Not blnFirst Then
						' PREVIOUS MONTH FOOTER ROW		
						response.write "<tr style=""font-weight:bold;background-color:#e0e0e0;"" align=center class=tablelist><td>Totals for " & sCurrentRange & "</td><td> " & iMonthReserveTotal & "</td><td> " & FormatCurrency(iMonthAmountTotal,2) & "</td></tr>"
					End If

					' SET ROW PLACEMENT
					blnFirst = False
					
					' MONTH HEADER ROW
					iPreviousMonth = iCurrentMonth 
					sCurrentRange = MonthName(oRs("Month")) & " " & oRs("Year")
					response.write "<tr align=""left"" class=""tablelist"" style=""font-weight: bold; border-top: 1px;""><td style=""font-weight:bold;border-top: 1px solid #000000;"">&nbsp;" & sCurrentRange & "</td><td style=""font-weight:bold;border-top: 1px solid #000000;"">&nbsp;</td><td style=""font-weight:bold;border-top: 1px solid #000000;"">&nbsp;</td></tr>"
					' CLEAR RUNNING TOTALS
					iMonthAmountTotal = 0
					iMonthReserveTotal = 0

				End If

				' DETAIL ROW
				response.write "<tr align=center class=""tablelist""><td>" & oRs("FACILITYNAME") & "</td><td> " & oRs("TOTALRESERVATIONS") & "</td><td> " & formatcurrency(oRs("totalamount"),2) & "</td></tr>"

				iMonthReserveTotal = iMonthReserveTotal + clng(oRs("TOTALRESERVATIONS"))
				iMonthAmountTotal = iMonthAmountTotal + clng(oRs("totalamount"))

				oRs.MoveNext
		Loop
		
		' PREVIOUS MONTH FOOTER ROW		
		response.write "<tr style=""font-weight:bold;background-color:#e0e0e0;"" align=center class=""tablelist""><td>Totals for " & sCurrentRange & "</td><td> " & iMonthReserveTotal & "</td><td> " & formatcurrency(iMonthAmountTotal,2) & "</td></tr>"

		' DISPLAY GRAND TOTALS
		response.write "<tr style=""font-weight:bold;"" align=center class=""tablelist""><td style=""font-weight:bold;border-top: 1px solid #000000;"">&nbsp;Grand Totals</td><td style=""font-weight:bold;border-top: 1px solid #000000;"">" & iReserveTotal & "</td><td style=""font-weight:bold;border-top: 1px solid #000000;"">" & curGrandTotal & "</td></tr>"
		response.write "</table>"

	Else

		' DISPLAY EMPTY REPORT
		response.write "<div style=""COLOR: #000000; font-family: verdana,sans-serif; font-size: 12px; font-weight:bold;margin-bottom: 2em;"">Total Lodge Counts from " & MonthName(startmonth) & " " & startYear & " to " & MonthName(endMonth) & " " & endYear & " </div>"
		response.write "<table cellspacing=0 cellpadding=2 class=""tablelist"" width=75% >"
		response.write "<tr class=tablelist><th >Contact Name</th><th>Lodge</th><th >Check In</th><th >Check Out</th><th >Cancel Date</th><th >Cancel Reason</th></tr>"
		response.write "<tr><td colspan=6>There Were No Reservations in the Specified Time Frame.</td></tr>"
		response.write "</table>"

	End If

	' DESTROY OBJECTS
	oRs.Close 
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' GENERATEFACILITYTOTALREPORT( startmonth, startyear, endmonth, endyear )
'--------------------------------------------------------------------------------------------------
Sub GenerateFacilityCancellationReport( ByVal startmonth, ByVal startyear, ByVal endmonth, ByVal endyear, ByVal iFacilityId )
	Dim sSql, oRs, sTempYear, sTempMonth, sWhere, dStartDate, dEndDate, bgcolor

	bgcolor = "#eeeeee"

	excel = false
	if request.querystring("excel") = "true" then excel = true

	if not excel then response.write "<div style=""COLOR: #000000; font-family: verdana,sans-serif; font-size: 12px; font-weight:bold;"">Lodge Cancellations</div>"

	' SET DEFAULT DATE IF NONE SUPPLIED
	If startmonth = "" or startyear = "" or endmonth = "" or endyear = "" Then
		startmonth = Month(Date())
		startyear = Year(Date())
		endmonth = Month(Date())
		endyear = Year(Date()) 
	End If


	' FLIP DATES IF REVERSED IF STARTYEAR GREATER THAN ENDYEAR
	If (startyear > endyear) Or (startyear = endyear And startmonth > endmonth) Then
		sTempYear = startyear
		sTempMonth	= startmonth
		startyear = endyear
		endyear = sTempYear
		startmonth = endmonth
		endmonth = startmonth
	End If

	if not excel then response.write vbcrlf & "<fieldset class=""filterselection"" id=""cancelreportpicks"">"
	if not excel then response.write vbcrlf & "<legend class=""filterselection"">Report Options</legend>"

	if not excel then response.write vbcrlf & "<form name=""frmdate"" action=# method=""post"">"
	if not excel then response.write vbcrlf & "<p><label for=""sm"">Check In Date Range:</label> <select id=""sm"" name=""sm"">"
	For m = 1 to 12
		sSelected = ""
		If clng(startmonth) = clng(m) Then
			sSelected = " selected=""selected"""
		End If
		if not excel then response.write "<option" & sSelected & " value="& m &">" & MonthName(m) & "</option>"
	Next
	if not excel then response.write "</select>"
	

	' START YEAR SELECT
	if not excel then response.write "<select name=""sy"">"
	For y = 2005 to Year(Date())+5
		sSelected = ""
		If clng(startyear) = clng(y) Then
			sSelected = " selected=""selected"""
		End If
		if not excel then response.write "<option" & sSelected & " value="& y &">" & y & "</option>"
	Next
	if not excel then response.write "</select> &mdash; "
	

	' END MONTH SELECT
	if not excel then response.write "<select name=""em"">"
	For m = 1 to 12
		sSelected = ""
		If clng(endmonth) = clng(m) Then
			sSelected = " selected=""selected"""
		End If
		if not excel then response.write "<option" & sSelected & " value="& m &">" & MonthName(m) & "</option>"
	Next
	if not excel then response.write "</select>"
	

	' END YEAR SELECT
	if not excel then response.write "<select name=""ey"">"
	For y = 2005 to Year(Date())+5
		sSelected = ""
		If clng(endyear) = clng(y) Then
			sSelected = " selected=""selected"""
		End If
		if not excel then response.write "<option" & sSelected & "  value="& y &">" & y & "</option>"
	Next
	if not excel then response.write "</select>"
	if not excel then response.write "</p>"

	' Show the facility picks
	if not excel then showFacilityPicks iFacilityId
	
	if not excel then response.write "<p><input type=""button"" class=""facilitybutton"" value=""Display Report"" onClick=""refreshreport(2);"" />&nbsp;&nbsp;&nbsp;<input type=""button""  class=""facilitybutton"" value=""Export to Excel"" onClick=""location.href=window.location+'&excel=true'"" /></p>"

	if not excel then response.write "</form></fieldset>"

	dStartDate = clng(startmonth) & "/1/" & clng(startyear)
	dStartDate = DateAdd( "d", -1, dStartDate )
	dEndDate = clng(endmonth) & "/1/" & clng(endyear)
	dEndDate = DateAdd( "m", 1, dEndDate )

	sWhere = " AND checkin > '" & dStartDate & "' AND checkin < '" & dEndDate & "'"

	If CLng(iFacilityId) > CLng(0) Then
		sWhere = sWhere & " AND facilityid = " & iFacilityId
	End If 
	'response.write sWhere & "<br ><br>"

	
	' GET INFORMATION FROM SQL
	sSql = "SELECT userlname, userfname, facilityname, checkindate, checkintime, checkoutdate, checkouttime, CONVERT( varchar, datecancelled, 101) AS datecancelled, canceldescription "
	sSql = sSql & "FROM rpt_facility_cancellations WHERE orgid = " & session("orgid") & sWhere
	sSql = sSql & " ORDER BY checkin"
	'response.write sSql & "<br ><br>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1


	' IF RECORDSET NOT EMPTY DISPLAY RESULTS
	If Not oRs.EOF Then

		' DISPLAY REPORT HEADER
		response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""75%"">"
		response.write "<tr class=""tablelist""><th>Contact Name</th><th>Lodge</th><th>Check In</th><th>Check Out</th><th>Cancel Date</th><th>Cancel Reason</th></tr>"

		' BEGIN REPORT DETAILS
		Do While Not oRs.EOF
			If bgcolor = "#eeeeee" Then 
				bgcolor = "#ffffff" 
			Else 
				bgcolor = "#eeeeee"
			End If 
			' DETAIL ROW
			response.write "<tr align=""center"" class=""tablelist"" bgcolor=""" &  bgcolor  & """>"
			response.write "<td>" & oRs("userlname") & ", " & oRs("userfname") &  "</td>"
			response.write "<td> " & oRs("facilityname") & "</td>"
			response.write "<td> " & oRs("checkindate") & " " & oRs("checkintime") &  "</td>"
			response.write "<td> " & oRs("checkoutdate") & " " & oRs("checkouttime") & "</td>"
			response.write "<td>" & oRs("datecancelled") & "</td>"
			response.write "<td> " & oRs("canceldescription") & "</td>"
			response.write "</tr>"

			oRs.MoveNext
		Loop

		response.write "</table>"

	Else

		' DISPLAY EMPTY REPORT
		response.write "<div><table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""75%"">"
		response.write "<tr class=""tablelist""><th>Contact Name</th><th>Lodge</th><th>Check In</th><th>Check Out</th><th>Cancel Date</th><th>Cancel Reason</th></tr>"
		response.write "<tr><td colspan=6>&nbsp;No Cancellations.</td></tr>"
		response.write "</table></div>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' showFacilityPicks iFacilityId
'--------------------------------------------------------------------------------------------------
Sub showFacilityPicks( ByVal iFacilityId )
	Dim sSql, oRs, sSelected

	sSql = "SELECT DISTINCT facilityid, facilityname FROM rpt_facility_cancellations WHERE orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<label for=""facilityid"">Lodge:</label> "
		response.write "<select id=""facilityid"" name=""facilityid"">"
		response.write vbcrlf & "<option" & sSelected & " value=""0"">All Lodges</option>"

		Do While Not oRs.EOF
			
			If CLng(iFacilityId) = CLng(oRs("facilityid")) Then
				sSelected = " selected=""selected"""
			Else 
				sSelected = ""
			End If

			response.write vbcrlf & "<option" & sSelected & " value="""& oRs("facilityid") &""">" &oRs("facilityname") & "</option>"

			oRs.MoveNext
		Loop 

		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 





%>


