<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: executive_summary.asp
' AUTHOR: SteveLoar
' CREATED: 10/17/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This pulls together the executive summary report. Part of the Menlo Park Project.
'
' MODIFICATION HISTORY
' 1.0   10/17/07		Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "executive summary rpt" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
Dim iClassSeasonId, sWhereClause, toDate, fromDate, iLocationId, iSupervisorId, iCategoryid, sFrom

sWhereClause = ""
sFrom = ""

If request("classseasonid") = "" Then 
	iClassSeasonId = GetRosterSeasonId()
Else
	iClassSeasonId = clng(request("classseasonid"))
End If 

If request("categoryid") = "" or CLng(request("categoryid")) = CLng(0) Then 
	iCategoryid = CLng(0)
Else
	iCategoryid = CLng(request("categoryid"))
End If 

If request("locationid") = "" Then
	iLocationId = CLng(0)
Else
	iLocationId = CLng(request("locationid"))
End If 

If request("supervisorid") = "" Then
	iSupervisorId = CLng(0)
Else
	iSupervisorId = CLng(request("supervisorid"))
End If 

fromDate = Request("fromDate")
toDate = Request("toDate")

' BUILD SQL WHERE CLAUSE
If iClassSeasonId > CLng(0) Then
	sWhereClause = sWhereClause & " AND classseasonid = " & iClassSeasonId
End If 

If fromDate <> "" And toDate <> "" Then 
	sWhereClause = sWhereClause & " AND (C.startdate >= '" & fromDate & " 00:00:00' AND C.startdate <= '" & toDate & " 00:00:00') "
End If 

If iCategoryid > CLng(0) Then
	sFrom = " , egov_class_category_to_class G "
	sWhereClause = sWhereClause & " AND C.classid = G.classid AND G.categoryid = " & iCategoryid
End If 

If iLocationId > CLng(0) Then
	sWhereClause = sWhereClause & " AND locationid = " & iLocationId
End If 

If iSupervisorId > CLng(0) Then
	sWhereClause = sWhereClause & " AND supervisorid = " & iSupervisorId
End If 

%>
<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="pageprint.css" media="print" />

	<script language="JavaScript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="Javascript">
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  //eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		  eval('window.open("calendarpicker.asp?updatefield=' + ToFrom + '&date=' + $("#" + ToFrom ).val() + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

	  //-->
	</script>

	<script language="Javascript" src="scripts/dates.js"></script>

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

<form action="executive_summary.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Executive Summary Report</strong></font></td>
		</tr>
		<tr>
			<td>
				<fieldset>
					<legend><strong>Select</strong></legend>
				
					<!--BEGIN: FILTERS-->
					<!--BEGIN: DATE FILTERS-->
					<p>
					<table border="0" cellpadding="2" cellspacing="0">
						<tr>
							<td><strong>Season:</strong></td>
							<td>
								<% ShowSeasonFilterPicks iClassSeasonId ' In class_global_functions.asp %>
							</td>
						</tr>
						<tr>
							<td><strong>Category:</strong></td>
							<td>
								<% DisplayCategorySelect iCategoryid  %>
							</td>
						</tr>
						<tr>
							<td><strong>Location:</strong></td>
							<td>
								<% ShowClassLocationPicks iLocationId %>
							</td>
						</tr>
						<tr>
							<td><strong>Supervisor:</strong></td>
							<td><% ShowActivitySupervisorPicks iSupervisorId %>
							</td>
						</tr>
						<tr>
							<td><strong>Start Date:</strong></td>
							<td>
								<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>		 
								&nbsp;
								<strong>To:</strong>
								<input type="text" id="toDate" name="toDate" value="<%=toDate%>" size="10" maxlength="10" />
								<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
								&nbsp; <%DrawDateChoices "Dates" %></td>
						</tr>
					</table>
					</p>
					<!--END: DATE FILTERS-->
					<p>
						<input class="button" type="submit" value="View Report" />
						<!-- &nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='activity_totals_export.asp?classseasonid=<%=iClassSeasonId%>&categoryid=<%=iCategoryid%>&locationid=<%=iLocationId%>&instructorid=<%=iInstructorId%>&supervisorid=<%=iSupervisorId%>&fromDate=<%=fromDate%>&toDate=<%=toDate%>&orderby=<%=iOrderBy%>&searchname=<%=sSearchName%>'" /> -->
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
<%	' DISPLAY RESULTS
	If request.servervariables("REQUEST_METHOD") = "POST" Then 
%>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
				
				<p class="executivesummary">
					<h3 class="executivesummary">Enrollment Totals</h3>
					Total Number of Residents..........................<%=GetEnrolledCount( "R", sWhereClause, sFrom ) %><br />
					Total Number of Non-Residents...................<%=GetEnrolledCount( "N", sWhereClause, sFrom ) %><br />
					Total Number of Enrollments.......................<%=GetEnrolledOrWait( "enrollmentsize", sWhereClause, sFrom ) %><br />
					Total Number of Wait List Holds...................<%=GetEnrolledOrWait( "waitlistsize", sWhereClause, sFrom ) %><br />
					Total Number from Web.............................<%=GetEnrolledWebOrOffice( "1", sWhereClause, sFrom ) %><br />
					Total Number from Office...........................<%=GetEnrolledWebOrOffice( "0", sWhereClause, sFrom ) %>
				</p><% response.flush %>

				<p class="executivesummary">
					<h3 class="executivesummary">Class Totals</h3>
					Total Number of Classes.............................<%=GetClassesCount( False, sWhereClause, sFrom ) %><br />
					Total Number of Cancelled Classes..............<%=GetClassesCount( True, sWhereClause, sFrom ) %><br />
					Total Number of Class Hours Offered...........<%=FormatNumber(GetNewTotalHours( False, sWhereClause, sFrom ),0,,,0)%><br />
					Total Number of Class Hours Held................<%=FormatNumber(GetNewTotalHours( True, sWhereClause, sFrom ),0,,,0)%>
				</p><% response.flush %>
				
				<p class="executivesummary">
					<h3 class="executivesummary">Revenue Totals</h3>
					<%
						dOfficeRevenue = GetClassRevenue( "office", sWhereClause, sFrom ) 
						dWebRevenue = GetClassRevenue( "web", sWhereClause, sFrom )
						dGrossRevenue = GetClassRevenue( "", sWhereClause, sFrom )
						dRefunds = GetClassRefunds( sWhereClause, sFrom )
						dNetRevenue = dGrossRevenue + dRefunds
						dInstructorNet = -GetInstructorNet( sWhereClause, sFrom )
						dNetIncome = dNetRevenue + dInstructorNet
					%>
					Office Revenue..........................................<%=FormatCurrency(dOfficeRevenue,2) %><br />
					Web Revenue............................................<%=FormatCurrency(dWebRevenue,2) %><br />
					Total Gross Revenue..................................<%=FormatCurrency(dGrossRevenue,2) %><br />
					Total Refunds............................................<%=FormatCurrency(dRefunds,2) %><br /><br />

					Net Revenue.............................................<%=FormatCurrency(dNetRevenue,2) %><br />
					Total Instructor Payments..........................<%=FormatCurrency(dInstructorNet,2) %><br /><br />
					<strong>
					Net Income............................................<%=FormatCurrency(dNetIncome,2) %>
					</strong>
				</p><% response.flush %>
				<!-- END: DISPLAY RESULTS -->
      
			</td>
		 </tr>
<%	Else %>
		<tr>
 			<td colspan="3" valign="top">
				<strong>To view the Executive Summary Report, select from the filter options above then click the &quot;View Report&quot; button.</strong>
			</td>
		 </tr>
<%	End If %>
	</table>
  </form>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'------------------------------------------------------------------------------------------------------------
' Sub DrawDateChoices( sName )
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(document.frmPFilter." & sName & ".value);"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassLocationPicks( iLocationId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassLocationPicks( ByVal iLocationId )
	Dim sSql, oRs

	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & SESSION("ORGID") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""locationid"" >" 
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iLocationId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">All Locations</option>"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("locationid") & """ "  
			If CLng(iLocationId) = CLng(oRs("locationid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If
	oRs.close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowActivitySupervisorPicks( iSupervisorId )
'--------------------------------------------------------------------------------------------------
Sub ShowActivitySupervisorPicks( ByVal iSupervisorId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname + ' ' + lastname AS name FROM users WHERE isclasssupervisor = 1 "
	sSql = sSql & "AND orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""supervisorid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iSupervisorId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">All Supervisors</option>"
		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("userid") & """ "  
			If CLng(iSupervisorId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"

	End If
	oRs.close
	Set oRs = Nothing
	
End Sub


'------------------------------------------------------------------------------------------------------------
' Function GetEnrolledCount( sResType, sWhereClause, sFrom )
'------------------------------------------------------------------------------------------------------------
Function GetEnrolledCount( ByVal sResType, ByVal sWhereClause, ByVal sFrom )
	Dim sSql, sResMatch, oRs

	If sResType = "R" Then
		sResMatch = " = 'R'"
	Else
		sResMatch = " != 'R'"
	End If 
	' select sum(enrolled) as enrolled from egov_enrollment_by_residency where orgid = 60 and residenttype = 'R' and classseasonid = 19
	sSql = "SELECT SUM(isnull(enrolled,0)) AS enrolled FROM egov_enrollment_by_residency C " & sFrom & " WHERE C.orgid = " & session("orgid")
	sSql = sSql & " AND C.residenttype" & sResMatch & sWhereClause

	'response.write "<!-- enrolled count: " & sSql & " -->"
	session("sSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	session("sSql") = ""
	
	If IsNull(oRs("enrolled")) Then
		GetEnrolledCount = CLng(0)
	Else 
		GetEnrolledCount = CLng(oRs("enrolled"))
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'------------------------------------------------------------------------------------------------------------
' Function GetEnrolledOrWait( sField, sWhereClause, sFrom )
'------------------------------------------------------------------------------------------------------------
Function GetEnrolledOrWait( ByVal sField, ByVal sWhereClause, ByVal sFrom )
	Dim sSql, oRs

	' select sum(enrollmentsize) as enrolled from egov_enrollment_waitlist_counts where orgid = 60
	sSql = "SELECT SUM(isnull(" & sField & ",0)) AS hits FROM egov_enrollment_waitlist_counts C " & sFrom & " WHERE orgid = " & session("orgid")
	sSql = sSql & sWhereClause

	'response.write sSql & "<br />"
	'response.write "<!-- enrolled or wait" & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If IsNull(oRs("hits")) Then
		GetEnrolledOrWait = CLng(0)
	Else 
		GetEnrolledOrWait = CLng(oRs("hits"))
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function


'------------------------------------------------------------------------------------------------------------
' Function GetEnrolledWebOrOffice( sWebMatch, sWhereClause, sFrom )
'------------------------------------------------------------------------------------------------------------
Function GetEnrolledWebOrOffice( ByVal sWebMatch, ByVal sWhereClause, ByVal sFrom )
	Dim sSql, oRs

	sSql = "SELECT SUM(isnull(enrolled,0)) AS enrolled FROM egov_enrollment_by_purchase_location C " & sFrom & " WHERE orgid = " & session("orgid")
	sSql = sSql & " AND ispublicmethod = " & sWebMatch & sWhereClause

	'response.write sSql & "<br />"
	'response.write "<!-- EnrolledWebOrOffice: " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If IsNull(oRs("enrolled")) Then
		GetEnrolledWebOrOffice = CLng(0)
	Else 
		GetEnrolledWebOrOffice = CLng(oRs("enrolled"))
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'------------------------------------------------------------------------------------------------------------
' Function GetClassesCount( bIsCancelled, sWhereClause, sFrom )
'------------------------------------------------------------------------------------------------------------
Function GetClassesCount( ByVal bIsCancelled, ByVal sWhereClause, ByVal sFrom )
	Dim sSql, oRs, sWhere
	
'	If sFrom = "" Then 
'		sFrom = " C "
'	End If 

	If bIsCancelled Then
		sWhere = " AND (upper(statusname) = 'CANCELLED' OR T.iscanceled = 1) "

	Else 
		sWhere = ""
	End if

	sSql = "SELECT count(C.classid) AS hits FROM egov_class_time T, egov_class_status S, egov_class C " & sFrom
	sSql = sSql & " WHERE C.statusid = S.statusid AND C.classid = T.classid AND C.orgid = " & session("orgid")
	sSql = sSql & sWhere & sWhereClause

	'response.write sSql & "<br />"
	'response.write "<!-- GetClassesCount: " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	GetClassesCount = CLng(oRs("hits"))

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'------------------------------------------------------------------------------------------------------------
' Function GetClassRevenue( sPurchasedFrom, sWhereClause, sFrom ) 
'------------------------------------------------------------------------------------------------------------
Function GetClassRevenue( ByVal sPurchasedFrom, ByVal sWhereClause, ByVal sFrom ) 
	Dim sSql, oRs, sWhere
	
	If sPurchasedFrom = "office" Then
		sWhere = " AND C.ispublicmethod = 0 "
	ElseIf sPurchasedFrom = "web" Then
		sWhere = " AND C.ispublicmethod = 1 "
	Else 
		sWhere = ""
	End if

	sSql = "SELECT SUM(amount) as amount FROM egov_activity_revenue_details C " & sFrom & " WHERE C.orgid = " & session("orgid")
	sSql = sSql & " AND C.amount > 0 " & sWhere & sWhereClause

	'response.write sSql & "<br />"
	'response.write "<!-- GetClassRevenue: " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		If IsNull(oRs("amount")) Then
			GetClassRevenue = CDbl(0.0)
		Else
			GetClassRevenue = CDbl(oRs("amount"))
		End If 
	Else
		GetClassRevenue = CDbl(0.0)
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'------------------------------------------------------------------------------------------------------------
' Function GetClassRefunds( sWhereClause, sFrom ) 
'------------------------------------------------------------------------------------------------------------
Function GetClassRefunds( ByVal sWhereClause, ByVal sFrom ) 
	Dim sSql, oRs
	
'	Changed this to try and prevent timeouts - Steve Loar 8/28/2009
	sSql = "SELECT SUM(amount) as amount FROM egov_activity_refund_details C " & sFrom & " WHERE C.orgid = " & session("orgid")
	sSql = sSql & " " & sWhereClause
	'dtb_debug(sSql)

	'response.write sSql & "<br />"
	'response.write "<!-- GetClassRefunds: " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		If IsNull(oRs("amount")) Then
			GetClassRefunds = CDbl(0.0)
		Else
			GetClassRefunds = CDbl(oRs("amount"))
		End If 
	Else
		GetClassRefunds = CDbl(0.0)
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'------------------------------------------------------------------------------------------------------------
' Function GetInstructorNet( sWhereClause, sFrom )
'------------------------------------------------------------------------------------------------------------
Function GetInstructorNet( ByVal sWhereClause, ByVal sFrom )
	Dim sSql, oRs
	
	sSql = "SELECT SUM(instructorpay) AS instructorpay FROM egov_activity_revenue_details C " & sFrom & " WHERE C.orgid = " & session("orgid")
	sSql = sSql & sWhereClause

	'response.write sSql & "<br />"
	'response.write "<!-- GetInstructorNet: " & sSql & " -->"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		If IsNull(oRs("instructorpay")) Then
			GetInstructorNet = CDbl(0.0)
		Else
			GetInstructorNet = CDbl(oRs("instructorpay"))
		End If 
	Else
		GetInstructorNet = CDbl(0.0)
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


'--------------------------------------------------------------------------------------------------
'  Function GetNewTotalHours( bIsNotCancelled, sWhereClause, sFrom )
'--------------------------------------------------------------------------------------------------
Function GetNewTotalHours( ByVal bIsNotCancelled, ByVal sWhereClause, ByVal sFrom )
	Dim dStartDate, dEndDate, dCurrDate, iMonth, iDay, iYear, iMeetingCount, dHours, sSql

	dHours = CDbl(0.0)

	If bIsNotCancelled Then
		sWhere = " AND upper(statusname) <> 'CANCELLED' AND T.iscanceled = 0 "
	Else 
		sWhere = ""
	End if

	sSql = "SELECT SUM(totalhours) AS totalhours FROM egov_class_time T, egov_class_status S, egov_class C " & sFrom
	sSql = sSql & " WHERE C.statusid = S.statusid AND C.classid = T.classid AND C.orgid = " & session("orgid")
	sSql = sSql & sWhere & sWhereClause

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If IsNull(oRs("totalhours")) Then
			dHours = CDbl(0.0)
		Else 
			dHours = CDbl(oRs("totalhours"))
		End If 
	Else
		dHours = CDbl(0.0)
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	GetNewTotalHours = dHours

End Function 


'--------------------------------------------------------------------------------------------------
'  Function GetTotalHours( bIsNotCancelled, sWhereClause, sFrom )
'--------------------------------------------------------------------------------------------------
Function GetTotalHours( ByVal bIsNotCancelled, ByVal sWhereClause, ByVal sFrom )
	Dim dStartDate, dEndDate, dCurrDate, iMonth, iDay, iYear, iMeetingCount, dHours, sSql

	dHours = CDbl(0.0)

	If bIsNotCancelled Then
		sWhere = " AND upper(statusname) <> 'CANCELLED' AND T.iscanceled = 0 "
	Else 
		sWhere = ""
	End if

	sSql = "SELECT C.classid, T.timeid FROM egov_class_time T, egov_class_status S, egov_class C " & sFrom
	sSql = sSql & " WHERE C.statusid = S.statusid AND C.classid = T.classid AND C.orgid = " & session("orgid")
	sSql = sSql & sWhere & sWhereClause

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF 
		GetActivityStartAndEndDates oRs("classid"), dStartDate, dEndDate

		If Not IsNull(dStartDate) And Not IsNull(dEndDate) Then
			If IsDate(dStartDate) And IsDate(dEndDate) Then 
				If Day(dStartDate) = Day(dEndDate) And Month(dStartDate) = Month(dEndDate) And Year(dStartDate) = Year(dEndDate) Then 
					' this is a one day event
					dHours = dHours + GetActivityHoursForDay( oRs("timeid"), WeekDayName(Weekday(dStartDate)) )
				Else
					' this class happens over several days, so walk the days from the start to the end
					dCurrDate = dStartDate
					Do While dCurrDate <= dEndDate
						If ClassMeetsThen( oRs("timeid"), WeekDayName(Weekday(dCurrDate)) ) Then
							dHours = dHours + GetActivityHoursForDay( oRs("timeid"), WeekDayName(Weekday(dCurrDate)) )
						End If
						dCurrDate = DateAdd("d", 1, dCurrDate )
					Loop 
				End If 
			End If 
		End If
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	GetTotalHours = dHours

End Function 


'--------------------------------------------------------------------------------------------------
'  Sub GetActivityStartAndEndDates( iClassid, dStartDate, dEndDate )
'--------------------------------------------------------------------------------------------------
Sub GetActivityStartAndEndDates( ByVal iClassid, ByRef dStartDate, ByRef dEndDate )
	Dim sSql, oRs

	sSql = "SELECT startdate, enddate FROM egov_class WHERE classid = " & iClassid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		dStartDate = oRs("startdate")
		dEndDate = oRs("enddate")
	Else
		dStartDate = "0/" 
		dEndDate = "0/" 
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
'  Function GetActivityHoursForDay( iTimeid, sDayOfWeek )
'--------------------------------------------------------------------------------------------------
Function GetActivityHoursForDay( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oRs, dHours, sAmOrPm, iColonPos, sHour, sMin, sStartDate, sEndDate

	dHours = CDbl(0.0)
	sSql = "SELECT starttime, endtime FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("starttime") <> "" And oRs("endtime") <> "" Then 
			' Get the start time in a format that can be used to compute the hours
			sAmOrPm = Right( oRs("starttime"), 2)
			iColonPos = InStr(oRs("starttime"), ":")
			sHour = Left( oRs("starttime"), ( iColonPos-1) )
			sMin = Mid( oRs("starttime"),(iColonPos + 1), ((Len(oRs("starttime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sStartDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			' Get the end time in a format that can be used to compute the hours
			sAmOrPm = Right( oRs("endtime"), 2)
			iColonPos = InStr(oRs("endtime"), ":")
			sHour = Left( oRs("endtime"), ( iColonPos-1) )
			sMin = Mid( oRs("endtime"),(iColonPos + 1), ((Len(oRs("endtime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sEndDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			dHours = dHours + abs(CDbl(DateDiff("s", sStartDate, sEndDate) / (60 * 60)))
		End If
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetActivityHoursForDay = CDbl(FormatNumber(dHours,2,,,0))

End Function 


'--------------------------------------------------------------------------------------------------
'  Function ClassMeetsThen( iTimeid, sDayOfWeek )
'--------------------------------------------------------------------------------------------------
Function ClassMeetsThen( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oRs

	sSql = "SELECT COUNT(timedayid) AS hits FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			ClassMeetsThen = True 
		Else
			ClassMeetsThen = False 
		End If 
	Else
		ClassMeetsThen = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Function 


%>
