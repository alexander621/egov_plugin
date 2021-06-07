<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: staff_attendance.asp
' AUTHOR: Terry Foster
' CREATED: 04/14/2016
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/14/2016	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp
sLocation = ""
%>
<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>
	

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link href="http://prevm.com/Content/bootstrap.min.css" rel="stylesheet" type="text/css" />
	<style>
		#content
		{
			padding-left:0;
			padding-right:0;
		}
		.boxcontent
		{
			padding-left:30px;
			padding-right:30px;
		}
		#keypad td
		{
			padding:5px;
		}
		#keypad td button
		{
			padding: 40px 46px;
		}
		a.btn-primary
		{
			color:white !important;
		}
	</style>
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
	<script src="http://prevm.com/Scripts/bootstrap.min.js" type="text/javascript"></script>
	
	
	


</head>
<body>



	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">

	<div class="boxcontent">
			<!--BEGIN: PAGE TITLE-->
				<font size="+1"><strong>Class Check In/Out</strong></font>
			<!--END: PAGE TITLE-->

			<% 
				LocationDatePicker()
				response.write "<br /><br />"
				if request.form("studentid") <> "" then
					ProcessCheckInOut()
				elseif request.querystring("locationid") <> "" then
					StudentSearch()
				end if 
			%>
	</div>

</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>

<%
Sub LocationDatePicker()
	%>
	<fieldset>
		<legend>Lookup</legend>
		<form method="GET" name="ClassListForm">
			<div class="form-group pull-left">
				<label for="locationid">Choose your location:</label>
				<%
				ShowLocationPicks(0)
				%>
			</div>
			<div class="form-group pull-left" style="margin-left:20px;">
				<label for="datesearch">Date:</label>
				<input type="text" name="dateSearch" id="dateSearch" size="10" value="<%if request.querystring("datesearch") = "" then%><%=date()%><%else%><%=request.querystring("datesearch")%><%end if%>" />
			    	<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('dateSearch');" /></span>
				<script>
					function doCalendar(sField) 
					{
						var w = (screen.width - 350)/2;
						var h = (screen.height - 350)/2;
						eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=ClassListForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
					}
				</script>
			</div>
			<div class="clearfix"></div>
			<input type="Submit" value="Start Attendance" class="btn" />
		</form>
	</fieldset>
	<%
End Sub
Sub ShowLocationPicks( ByVal iLocationId )
	Dim sSql, oRs

	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & Session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""locationid"">"

	iLocationId = request.querystring("locationid")
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("locationid") & """ "
			If CLng(oRs("locationid")) = CLng(iLocationId) Then 
				response.write " selected=""selected"" "
				sLocation = oRs("name")
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop 
	Else
		response.write vbcrlf & vbtab & "<option value=""0"">Unknown Location</option>"
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 
Sub StudentSearch()

	sDate = request.querystring("datesearch")
	if sDate = "" then sDate = date()

	sSQL = "SELECT c.classid,c.classname,l.attendeeuserid as studentid,c.locationid, " _
		& " ISNULL(att.userfname,'') + ' ' + ISNULL(att.userlname,'') as studentname " _
		& " FROM egov_class c " _
		& " INNER JOIN egov_class_list l ON l.classid = c.classid " _
		& " INNER JOIN egov_users att ON att.userid = l.attendeeuserid " _
		& " WHERE c.locationid = '" & dbsafe(request.querystring("locationid")) & "' " _
		& " and c.orgid = " & session("orgid") _
		& " and startdate <= '" & sDate & "' and enddate >= '" & sDate & "' " _
		& " AND (status = 'ACTIVE'  OR status = 'WAITLIST') " _
		& " ORDER BY studentname,studentid "

	set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1

	sSQL = "SELECT * FROM egov_checkinouttype"
	set oRsTypes = Server.CreateObject("ADODB.RecordSet")
	oRsTypes.Open sSQL, Application("DSN"), 3, 1

	%>
	<form method="POST">
		<h3>Location: <%=sLocation%>&nbsp;&nbsp;&nbsp; Date: <%=sDate%></h3>
		<input type="hidden" name="locationid" value="<%=request.querystring("locationid")%>" />
		<input type="hidden" name="locationname" value="<%=sLocation%>" />
		<table class="table table-bordered table-striped" style="width:500px;top:0;left:0;">
			<thead>
			<tr>
				<th>Check In/Out</th>
				<th>Class Name</th>
				<th>Name</th>
				<th>In/Out</th>
				<th>Type</th>
			</tr>
			</thead>
			<tbody>
			<%
			Do WHile Not oRs.EOF
				%>
				<tr>
					<input type="hidden" name="studentname_<%=oRs("StudentID")%>_<%=oRs("ClassID")%>" value="<%=oRs("studentname")%>" />
					<input type="hidden" name="classname_<%=oRs("StudentID")%>_<%=oRs("ClassID")%>" value="<%=oRs("classname")%>" />

					<td><input type="checkbox" name="studentid" value="<%=oRs("StudentID")%>_<%=oRs("ClassID")%>" class="form-control" /></td>
					<td style="vertical-align:middle;"><%=oRs("classname")%></td>
					<td style="vertical-align:middle;"><%=oRs("studentname")%></td>
					<%
						InOut = InOrOut(oRs("StudentID"), oRs("ClassID"))
						InOut = UCase(Left(InOut,1)) & LCase(Right(InOut, Len(InOut) - 1))
					%>
					<td style="vertical-align:middle;"><%=InOut%> </td>
					<td style="vertical-align:middle;">
						<select name="type_<%=oRs("StudentID")%>_<%=oRs("ClassID")%>">
							<%
							Do While Not oRsTypes.EOF
								%><option value="<%=oRsTypes("egov_checkinouttypeid")%>"><%=oRsTypes("Description")%></option><%
								oRsTypes.MoveNext
							loop
							oRsTypes.MoveFirst
							%>
						</select>
					</td>
				</tr>
				<%
				oRs.MoveNext
			loop
			%>
			</tbody>
		</table>
		<input type="Submit" value="Check In/Out" class="btn" />
	</form>
	<%
	oRsTypes.Close
	Set oRsTypes = Nothing
	oRs.Close
	Set oRs = Nothing
End Sub
Sub ProcessCheckInOut
	%>
	<h2>The following was checked in/out:</h2>
	<%
	'Insert Into  DB and show page
	arrStudents = split(dbsafe(request.form("studentid")),",")
	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.Open Application("DSN")
	for each x in arrStudents
		x = trim(x)

		arrGroup = split(x,"_")
		studentid = arrGroup(0)
		classid = arrGroup(1)
	
		checkinout = InOrOut(studentid, classid)
		studentname = dbsafe(request.form("studentname_" & x))
		classname = dbsafe(request.form("classname_" & x))
		checkinouttype = dbsafe(request.form("type_" & x))

		locationid = dbsafe(request.form("locationid"))
		locationname = dbsafe(request.form("locationname"))
		staffid = dbsafe(request.cookies("user")("userid"))
		staffname = dbsafe(request.cookies("user")("fullname"))
		datGMTDateTime = DateAdd("h",5,now())
		checkinoutdate = DateAdd("h",GetTimeOffset(session("orgid")),datGMTDateTime)

		sSQL = "INSERT INTO egov_classattendance (Checkinout,checkinouttype,classid,locationid,studentid,studentname,locationname,staffid,staffname,checkinoutdate) " _
			& " VALUES('" & checkinout & "','" & checkinouttype & "','" & classid & "','" & locationid & "','" & studentid & "','" & studentname & "','" & locationname & "','" & staffid & "','" & staffname & "','" & checkinoutdate & "')"
			'response.write sSQL & "<br />"
		oCmd.Execute(sSQL)
		Response.write "<h3>" & studentname & " - checked " & checkinout & " for " & classname & "</h3><br />"

	next
	oCmd.Close
	Set oCmd = Nothing

End Sub

Function GetStudentName(id)
	StudentName = ""
	sSQL = "SELECT ISNULL(userfname,'') + ' ' + ISNULL(userlname,'') as studentname FROM egov_users WHERE userid = '" & id & "'"
	'response.write sSQL & "<br />"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then
		StudentName = oRs("studentname")
	end if
	oRs.Close
	Set oRs = Nothing
	GetStudentName = StudentName
End Function
Function InOrOut(id,classid)
	inout = "in"
	sSQL = "SELECT TOP 1 checkinout " _
		& " FROM egov_classattendance " _
		& " WHERE studentid = '" & id & "' AND checkinoutdate > '" & date() & "' AND checkinoutdate < '" & dateadd("d",1,date()) & "' " _
		& " AND classid = '" & classid & "' " _
		& " ORDER BY checkinoutdate DESC "
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then
		if oRs("checkinout") = "in" then inout = "out"
	end if
	oRs.Close
	Set oRs = Nothing
	InOrOut = inout
End Function
%>
