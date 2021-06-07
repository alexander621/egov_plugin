<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: activity_total.asp
' AUTHOR: SteveLoar
' CREATED: 10/09/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This pulls together the activity totals report. Part of a Menlo Park Project.
'
' MODIFICATION HISTORY
' 1.0   10/09/2007	Steve Loar - INITIAL VERSION
' 1.1	09/30/2011	Steve Loar - Changine the OPEN column to Drop Ins for Menlo Park
' 1.2   10/15/2013	Steve Loar - Adding sort and filter for Class End Date
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

' INITIALIZE AND DECLARE VARIABLES
' SPECIFY FOLDER LEVEL
sLevel = "../" ' Override of value from common.asp


' USER SECURITY CHECK
If Not UserHasPermission( Session("UserId"), "activity totals rpt" ) Then
	response.redirect sLevel & "../permissiondenied.asp"
End If 


' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
Dim iLocationId, iSupervisorId, iPaymentLocationId, iClassSeasonId, sSearchName, sSearchActivity, iCategoryid
Dim iInstructorId, sWhereClause, sFrom, iOrderBy, sOrderBy, fromEndDate, toEndDate, toDateDisplay, toEndDateDisplay

sWhereClause = ""
sFrom = ""

If request("classseasonid") = "" Then 
	iClassSeasonId = GetRosterSeasonId()
Else
	iClassSeasonId = clng(request("classseasonid"))
End If 

If request("categoryid") = "" Or CLng(request("categoryid")) = CLng(0) Then 
	iCategoryid = CLng(0)
Else
	iCategoryid = CLng(request("categoryid"))
End If 

If request("locationid") = "" Then
	iLocationId = CLng(0)
Else
	iLocationId = CLng(request("locationid"))
End If 

If request("instructorid") = "" Then
	iInstructorId = CLng(0)
Else
	iInstructorId = CLng(request("instructorid"))
End If 

If request("supervisorid") = "" Then
	iSupervisorId = CLng(0)
Else
	iSupervisorId = CLng(request("supervisorid"))
End If 

If request("searchname") <> "" Then 
	sSearchName = dbsafe(request("searchname"))
End If 

fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

fromEndDate = Request("fromEndDate")
toEndDate = Request("toEndDate")

If request("orderby") = "" Then
	iOrderBy = clng(1)
Else
	iOrderBy = clng(request("orderby"))
End If 


' BUILD SQL WHERE CLAUSE
If iClassSeasonId > CLng(0) Then
	sWhereClause = sWhereClause & " AND C.classseasonid = " & iClassSeasonId
End If 

If iCategoryid > CLng(0) Then
	sFrom = ", egov_class_category_to_class G "
	sWhereClause = sWhereClause & " AND C.classid = G.classid AND G.categoryid = " & iCategoryid
End If 

If iLocationId > CLng(0) Then
	sWhereClause = sWhereClause & " AND C.locationid = " & iLocationId
End If 

If iInstructorId > CLng(0) Then
	sWhereClause = sWhereClause & " AND T.instructorid = " & iInstructorId
End If 

If iSupervisorId > CLng(0) Then
	sWhereClause = sWhereClause & " AND C.supervisorid = " & iSupervisorId
End If 

If sSearchName <> "" Then
	sWhereClause = sWhereClause & " AND C.classname LIKE '%" & sSearchName & "%' "
End If 

If fromDate <> "" Then 
	sWhereClause = sWhereClause & " AND C.startdate >= '" & fromDate & " 00:00:00' "
End If 

If toDate <> "" Then 
	toDateDisplay = toDate
	toDate = DateAdd( "d", 1, toDate )
	sWhereClause = sWhereClause & " AND C.startdate < '" & toDate & " 00:00:00' "
End If 

If fromEndDate <> "" Then 
	sWhereClause = sWhereClause & " AND C.enddate >= '" & fromEndDate & " 00:00:00' "
End If 

If toEndDate <> "" Then 
	toEndDateDisplay = toEndDate
	toEndDate = DateAdd( "d", 1, toEndDate )
	sWhereClause = sWhereClause & " AND C.enddate < '" & toEndDate & " 00:00:00' "
End If 

If iOrderby = clng(1) Then 
	sOrderBy = "classname, activityno"
ElseIf iOrderby = clng(2) Then 
	sOrderBy = "C.startdate, classname, activityno"
ElseIf iOrderby = clng(3) Then 
	sOrderBy = "C.enddate, classname, activityno"
End If 


%>
<html lang="en">
<head>
	<meta charset="UTF-8">
  	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="reporting.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../classes/classes.css" />
	<link rel="stylesheet" href="pageprint.css" media="print" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	<script src="scripts/tablesort.js"></script>
	<script src="scripts/dates.js"></script>

	<script>
	  <!--
		function doCalendar(ToFrom) {
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  //eval('window.open("../recreation/gr_calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		  eval('window.open("calendarpicker.asp?updatefield=' + ToFrom + '&date=' + $("#" + ToFrom ).val() + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

	  //-->
	</script>

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

<form action="activity_totals.asp" method="post" name="frmPFilter">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><strong>Activity Totals Report</strong></font></td>
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
							<td><strong>Instructor:</strong></td>
							<td><% ShowActivityInstructorPicks iInstructorId %>
							</td>
						</tr>
						<tr>
							<td><strong>Supervisor:</strong></td>
							<td><% ShowActivitySupervisorPicks iSupervisorId %>
							</td>
						</tr>
						<tr>
							<td><strong>Name Like:</strong></td>
							<td><input type="text" name="searchname" value="<%=sSearchName%>" size="75" maxlength="255" /></td>
						</tr>
						<tr>
							<td><strong>Start Date:</strong></td>
							<td>
								<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" size="10" maxlength="10" placeholder=">= date" />
								<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>		 
								&nbsp;
								<strong>To:</strong>
								<input type="text" id="toDate" name="toDate" value="<%=toDateDisplay%>" size="10" maxlength="10" placeholder="<= date" />
								<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
								&nbsp; <%DrawDateChoices "Dates" %></td>
						</tr>
						<tr>
							<td><strong>End Date:</strong></td>
							<td>
								<input type="text" id="fromEndDate" name="fromEndDate" value="<%=fromEndDate%>" size="10" maxlength="10" placeholder=">= date" />
								<a href="javascript:void doCalendar('fromEndDate');"><img src="../images/calendar.gif" border="0" /></a>		 
								&nbsp;
								<strong>To:</strong>
								<input type="text" id="toEndDate" name="toEndDate" value="<%=toEndDateDisplay%>" size="10" maxlength="10" placeholder="<= date" />
								<a href="javascript:void doCalendar('toEndDate');"><img src="../images/calendar.gif" border="0" /></a>
								&nbsp; <%DrawDateChoices "EndDates" %>
							</td>
						</tr>
						<tr>
							<td><strong>Order By:</strong></td>
							<td><% ShowOrderByPicks iOrderBy %>
							</td>
						</tr>
					</table>
					</p>
					<!--END: DATE FILTERS-->
					<p>
						<input class="button" type="submit" value="View Report" />
						&nbsp;&nbsp;<input type="button" class="button" value="Download to Excel" onClick="location.href='activity_totals_export.asp?classseasonid=<%=iClassSeasonId%>&categoryid=<%=iCategoryid%>&locationid=<%=iLocationId%>&instructorid=<%=iInstructorId%>&supervisorid=<%=iSupervisorId%>&fromDate=<%=fromDate%>&toDate=<%=toDateDisplay%>&fromEndDate=<%=fromEndDate%>&toEndDate=<%=toEndDateDisplay%>&orderby=<%=iOrderBy%>&searchname=<%=sSearchName%>'" />
					</p>

				</fieldset>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 
			<td colspan="3" valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->
				<%
				
				' DISPLAY RESULTS
				If request.servervariables("REQUEST_METHOD") = "POST" Then 
					Display_Results sWhereClause, sFrom, sOrderBy
				Else
					response.write "<strong>To view the Activity Totals Report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
				End If 
				
				%>
				<!-- END: DISPLAY RESULTS -->
      
			</td>
		 </tr>
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
' void Display_Results sWhereClause, sFrom, sOrderBy 
'------------------------------------------------------------------------------------------------------------
Sub Display_Results( ByVal sWhereClause, ByVal sFrom, ByVal sOrderBy )
	Dim sSql, oRs, iOpen, iMeetingCount, dHours, dRevenue, dPayment, dNetIncome, dTotalRevenue, iTotalDropIn
	Dim dTotalPayment, dTotalNetIncome, dTotalHrs, dTotalMeetings, dTotalMin, dTotalMax, dTotalRes, iDropInCount
	Dim dResCount, dNonResCount, dTotalNonRes, dTotalEnrollment, dTotalWait, iTotalOpen, iTotalAttendance

	dTotalRevenue = CDbl(0.0)
	dTotalPayment = CDbl(0.0)
	dTotalNetIncome = CDbl(0.0)
	dTotalHrs = CDbl(0.0)
	dTotalMeetings = CLng(0)
	dTotalMin = CLng(0)
	dTotalMax = CLng(0)
	dTotalRes = CLng(0)
	dTotalNonRes = CLng(0)
	dTotalEnrollment = CLng(0)
	dTotalWait = CLng(0)
	iTotalOpen = CLng(0)
	iTotalAttendance = CDbl(0.0)
	iTotalDropIn = CLng(0)

	sSql = "SELECT C.classname, C.classid, T.activityno, C.startdate, C.enddate, C.statusid, C.classseasonid, "
	sSql = sSql & " C.locationid, C.supervisorid, isnull(T.instructorid,0) as instructorid, T.min, T.max, T.timeid, "
	sSql = sSql & " T.enrollmentsize, T.waitlistsize, S.seasonname "
	sSql = sSql & " FROM egov_class C, egov_class_time T, egov_class_seasons S " & sFrom
	sSql = sSql & " WHERE C.classid = T.classid AND C.classseasonid = S.classseasonid "
	sSql = sSql & " AND C.orgid = " & Session("orgid") & sWhereClause
	sSql = sSql & " ORDER BY " & sOrderBy

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If oRs.EOF then
		' EMPTY
		response.write "<p><strong>No Activities found for your selection criteria.</strong></p>"
	Else
		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" id=""activitytotal"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>ClassName</th><th>Activity No.</th><th>Season</th><th>Start Date</th><th>End Date</th>"
		response.write "<th>#<br />Hrs</th><th>#<br />Sess</th><th>Min</th><th>Max</th><th>Res</th><th>Non<br />Res</th><th>Total<br />Enrld</th>"
		response.write "<th>Wait</th><th>Drop<br />In</th><th>Attnd</th><th>Total Revenue</th><th>Instr Payment</th><th>Net Income</th></tr>"

		bgcolor = "#eeeeee"

		Do While Not oRs.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			response.write "<td>" & oRs("classname") & "</td>"
			response.write "<td>" & oRs("activityno") & "</td>"
			response.write "<td>" & oRs("seasonname") & "</td>"
			response.write "<td>" & oRs("startdate") & "</td>"
			response.write "<td>" & oRs("enddate") & "</td>"
			dHours = CDbl(0.0)
			iMeetingCount = GetNewActivityMeetingCount( oRs("timeid"), dHours )
			dTotalHrs = dTotalHrs + CDbl(dHours)
			dTotalMeetings = dTotalMeetings + CLng(iMeetingCount)
			response.write "<td align=""center"">" & dHours & "</td>"
			response.write "<td align=""center"">" & iMeetingCount & "</td>"
			
			response.write "<td align=""center"">" & oRs("min") & "</td>"
			If IsNumeric(oRs("min")) Then
				dTotalMin = dTotalMin + CLng(oRs("min"))
			End If 
			response.write "<td align=""center"">" & oRs("max") & "</td>"
			If IsNumeric(oRs("max")) Then
				dTotalMax = dTotalMax + CLng(oRs("max"))
			End If 
			dResCount = GetResNonResClassCount( oRs("timeid"), "R" )
			dTotalRes = dTotalRes + CLng(dResCount)
			response.write "<td align=""center"">" & dResCount & "</td>"
			dNonResCount = GetResNonResClassCount( oRs("timeid"), "N" )
			dTotalNonRes = dTotalNonRes + dNonResCount
			response.write "<td align=""center"">" & dNonResCount & "</td>"
			response.write "<td align=""center"">" & oRs("enrollmentsize") & "</td>"
			dTotalEnrollment = dTotalEnrollment + CLng(oRs("enrollmentsize"))
			response.write "<td align=""center"">" & oRs("waitlistsize") & "</td>"
			dTotalWait = dTotalWait + CLng(oRs("waitlistsize"))
			If IsNull(oRs("max")) Then
				iOpen = "N/A"
			Else
				iOpen = CLng(oRs("max")) - CLng(oRs("enrollmentsize"))
				If iOpen < CLng(0) Then 
					iOpen = CLng(0)
				End If 
				iTotalOpen = iTotalOpen + iOpen
			End If 
			' Display Open Count
			'response.write "<td align=""center"">" & iOpen & "</td>"

			' Display Drop IN Count
			iDropInCount = GetDropInCount( oRs("timeid") )
			iTotalDropIn = iTotalDropIn + iDropInCount
			response.write "<td align=""center"">" & iDropInCount & "</td>"

			response.write "<td align=""right"">" & CLng(dHours * CDbl(oRs("enrollmentsize"))) & "</td>"
			iTotalAttendance = iTotalAttendance + CLng(dHours * CDbl(oRs("enrollmentsize")))
			getRevenueAndPay oRs("timeid"), dRevenue, dPayment
			dNetIncome = dRevenue - dPayment
			dTotalRevenue = dTotalRevenue + dRevenue
			dTotalPayment = dTotalPayment + dPayment
			dTotalNetIncome = dTotalNetIncome + dNetIncome
			response.write "<td align=""right"">" & FormatNumber(dRevenue,2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(dPayment,2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(dNetIncome,2) & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop 
		' Total for all Classes
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalHrs,2) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalMin,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalRes,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalEnrollment,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		'response.write "<td align=""center"">" & FormatNumber(iTotalOpen,0) & "</td>"
		response.write "<td align=""center"">" & FormatNumber(iTotalDropIn,0) & "</td>"
'		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"" colspan=""2"">" & FormatNumber(dTotalRevenue,2) & "</td>"
'		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"" colspan=""2"">" & FormatNumber(dTotalNetIncome,2) & "</td>"
		response.write "</tr>"
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"" align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalMeetings,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalMax,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalNonRes,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center"">" & FormatNumber(dTotalWait,0) & "</td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""right"">" & FormatNumber(iTotalAttendance,0) & "</td>"
'		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"" colspan=""2"">" & FormatNumber(dTotalPayment,2) & "</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "</tr>"
		response.write vbcrlf & "</table>"
	End If

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' void DrawDateChoices sName 
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(document.frmPFilter." & sName & ".value, '" + sName + "');"" id=""" & sName & """ class=""calendarinput"" name=""" & sName & """>"
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
' void ShowActivityInstructorPicks iInstructorId 
'--------------------------------------------------------------------------------------------------
Sub ShowActivityInstructorPicks( ByVal iInstructorId )
	Dim sSql, oRs

	sSql = "SELECT * FROM EGOV_CLASS_INSTRUCTOR WHERE ORGID = " & SESSION("ORGID") & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If not oRs.EOF Then
		response.write vbcrlf & "<select name=""instructorid"">"
		response.write vbcrlf & "<option value=""0"" "
		If CLng(iInstructorId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">All Instructors</option>"

		Do While NOT oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("instructorid") & """ "  
			If CLng(iInstructorId) = CLng(oRs("instructorid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " >" & oRs("lastname") & ", " & oRs("firstname")& "</option>"
			oRs.MoveNext
		Loop

		response.write vbcrlf & "</select>"

	End If

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' void ShowClassLocationPicks iLocationId 
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

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowActivitySupervisorPicks iSupervisorId 
'--------------------------------------------------------------------------------------------------
Sub ShowActivitySupervisorPicks( ByVal iSupervisorId )
	Dim sSql, oRs

	sSql = "Select userid, firstname + ' ' + lastname as name From users Where isclasssupervisor = 1 and orgid = " & SESSION("orgid") & " ORDER BY lastname, firstname"
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

	oRs.Close
	Set oRs = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' integer GetResNonResClassCount( iTimeid, sResType )
'--------------------------------------------------------------------------------------------------
Function GetResNonResClassCount( ByVal iTimeid, ByVal sResType )
	Dim sSql, sResMatch, oRs

	If sResType = "R" Then
		sResMatch = " = 'R'"
	Else
		sResMatch = " != 'R'"
	End If 
	
	sSql = "SELECT COUNT(attendeeuserid) AS hits FROM egov_class_list L, egov_users U WHERE L.attendeeuserid = U.userid "
	sSql = sSql & " AND L.status = 'ACTIVE' AND L.classtimeid = " & iTimeid & " AND U.residenttype " & sResMatch

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	GetResNonResClassCount = CLng(oRs("hits"))

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
'  integer GetNewActivityMeetingCount( iTimeid, dHours )
'--------------------------------------------------------------------------------------------------
Function GetNewActivityMeetingCount( ByVal iTimeid, ByRef dHours )
	Dim sSql, oRs, iMeetingCount

	sSql = "Select meetingcount, totalhours FROM egov_class_time WHERE timeid = " & iTimeid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		dHours = CDbl(oRs("totalhours"))
		iMeetingCount = CLng(oRs("meetingcount"))
	Else
		dHours = CDbl(0.0)
		iMeetingCount = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetNewActivityMeetingCount = iMeetingCount
End Function 


'--------------------------------------------------------------------------------------------------
'  void GetActivityStartAndEndDates iClassid, dStartDate, dEndDate 
'--------------------------------------------------------------------------------------------------
Sub GetActivityStartAndEndDates( ByVal iClassid, ByRef dStartDate, ByRef dEndDate )
	Dim sSql, oRs

	sSql = "Select startdate, enddate FROM egov_class WHERE classid = " & iClassid 

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
'  integer GetActivityMeetingCount( iClassid, iTimeid, dHours )
'--------------------------------------------------------------------------------------------------
Function GetActivityMeetingCount( ByVal iClassid, ByVal iTimeid, ByRef dHours )
	Dim dStartDate, dEndDate, dCurrDate, iMonth, iDay, iYear, iMeetingCount

	iMeetingCount = 0
	dHours = CDbl(0.0)

	GetActivityStartAndEndDates iClassid, dStartDate, dEndDate

	If IsNull(dStartDate) Or IsNull(dEndDate) Then
		' one or more dates is missing so cannot create a sheet
		iMeetingCount = 0
	ElseIf IsDate(dStartDate) And IsDate(dEndDate) Then 
		If Day(dStartDate) = Day(dEndDate) And Month(dStartDate) = Month(dEndDate) And Year(dStartDate) = Year(dEndDate) Then 
			' this is a one day event
			iMeetingCount = 1
			dHours = dHours + GetActivityHoursForDay( iTimeid, WeekDayName(Weekday(dStartDate)) )
		Else
			' this class happens over several days
			dCurrDate = dStartDate
			Do While dCurrDate <= dEndDate
				If ClassMeetsThen( iTimeid, WeekDayName(Weekday(dCurrDate)) ) Then
					iMeetingCount = iMeetingCount + 1
					dHours = dHours + GetActivityHoursForDay( iTimeid, WeekDayName(Weekday(dCurrDate)) )
				End If
				dCurrDate = DateAdd("d", 1, dCurrDate )
			Loop 
		End If 
	Else
		' one or more dates is not a date
		iMeetingCount = 0
	End If 

	GetActivityMeetingCount = iMeetingCount

End Function 


'--------------------------------------------------------------------------------------------------
'  boolean ClassMeetsThen( iTimeid, sDayOfWeek )
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


'--------------------------------------------------------------------------------------------------
'  double GetActivityHoursForDay( iTimeid, sDayOfWeek )
'--------------------------------------------------------------------------------------------------
Function GetActivityHoursForDay( ByVal iTimeid, ByVal sDayOfWeek )
	Dim sSql, oTime, dHours, sAmOrPm, iColonPos, sHour, sMin, sStartDate, sEndDate

	dHours = CDbl(0.0)
	sSql = "SELECT starttime, endtime FROM egov_class_time_days WHERE timeid = " & iTimeid & " AND " & sDayOfWeek & " = 1" 

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSql, Application("DSN"), 3, 1

	Do While Not oTime.EOF
		If oTime("starttime") <> "" And oTime("endtime") <> "" Then 
			' Get the start time in a format that can be used to compute the hours
			sAmOrPm = Right( oTime("starttime"), 2)
			iColonPos = InStr(oTime("starttime"), ":")
			sHour = Left( oTime("starttime"), ( iColonPos-1) )
			sMin = Mid( oTime("starttime"),(iColonPos + 1), ((Len(oTime("starttime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sStartDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			' Get the end time in a format that can be used to compute the hours
			sAmOrPm = Right( oTime("endtime"), 2)
			iColonPos = InStr(oTime("endtime"), ":")
			sHour = Left( oTime("endtime"), ( iColonPos-1) )
			sMin = Mid( oTime("endtime"),(iColonPos + 1), ((Len(oTime("endtime"))-2)-iColonPos))
			If UCase(sAmOrPm) = "PM" And clng(sHour) < clng(12) Then 
				sHour = clng(sHour) + clng(12)
			End If
			sEndDate = CDate(Month(now) &"/" & Day(now) & "/" & Year(now) & " " & sHour & ":" & sMin )
			dHours = dHours + abs(CDbl(DateDiff("s", sStartDate, sEndDate) / (60 * 60)))
		End If
		oTime.MoveNext 
	Loop 

	oTime.Close
	Set oTime = Nothing 

	GetActivityHoursForDay = CDbl(FormatNumber(dHours,2,,,0))

End Function 


'--------------------------------------------------------------------------------------------------
'  void getRevenueAndPay iClassTimeid, dRevenue, dPayment 
'--------------------------------------------------------------------------------------------------
Sub getRevenueAndPay( ByVal iClassTimeid, ByRef dRevenue, ByRef dPayment )
	Dim sSql, oRs
	
	dRevenue = CDbl(0.0)
	dPayment = CDbl(0.0)
	sSql = "SELECT classtimeid, SUM(amount) AS revenue, SUM(instructorpay) AS instructorpay FROM egov_activity_revenue_details "
	sSql = sSql & " WHERE classtimeid = " & iClassTimeid & " GROUP BY classtimeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		dRevenue = CDbl(oRs("revenue"))
		dPayment = CDbl(oRs("instructorpay"))
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
'  void ShowOrderByPicks iOrderBy 
'--------------------------------------------------------------------------------------------------
Sub ShowOrderByPicks( ByVal iOrderBy )

	response.write vbcrlf & "<select name=""orderby"">"

	response.write vbcrlf & "<option value=""1"" "
	If clng(iOrderBy) = clng(1) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Class Name</option>"

	response.write vbcrlf & "<option value=""2"" "
	If clng(iOrderBy) = clng(2) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Start Date</option>"

	response.write vbcrlf & "<option value=""3"" "
	If clng(iOrderBy) = clng(3) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">End Date</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' integer GetDropInCount( iTimeid )
'--------------------------------------------------------------------------------------------------
Function GetDropInCount( ByVal iTimeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(attendeeuserid) AS hits FROM egov_class_list L, egov_users U WHERE L.attendeeuserid = U.userid "
	sSql = sSql & " AND L.status = 'DROPIN' AND L.classtimeid = " & iTimeid 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetDropInCount = CLng(oRs("hits"))
	Else
		GetDropInCount = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 




%>
