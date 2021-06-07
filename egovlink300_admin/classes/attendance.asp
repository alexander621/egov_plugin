<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: attendance.asp
' AUTHOR: Terry Foster
' CREATED: 04/13/2016
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   04/13/2016	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp
%>
<html>
<head>
	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>
<meta content="width=device-width, minimum-scale=1, maximum-scale=1" name="viewport" />
	

	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link href="http://prevm.com/Content/bootstrap.min.css" rel="stylesheet" type="text/css" />
	<style>
		.hidetablet
		{
			display:none;
		}
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
				if request.querystring("locationid") = "" then
					LocationPicker()
				elseif request.form("output") <> "" then
					ProcessSignature()
				elseif request.form("pin") <> "" then
					StudentSearch()
				elseif session("parentid") = "" then
					CodeEntry()
				else
				end if 
			%>
	</div>

</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>

<%
Sub LocationPicker()
	%>
	<form method="GET">
	<div class="form-group">
	<label for="locationid">Choose your location:</label>
	<%
	ShowLocationPicks(0)
	%>
	</div>
	<br />
	<br />
		<input type="Submit" value="Start Taking Attendance" class="btn btn-primary btn-lg" />
	</form>
	<%
End Sub
Sub CodeEntry()
	%>
	&nbsp;&nbsp;&nbsp;Location: <%=GetLocationName(dbsafe(request.querystring("locationid")))%>
	<form method="POST">
		<center>
		<b>Enter your pin:</b><br />
		<input type="text" size="10" style="text-align:center;width:200px;height:50px;font-size:30px;" name="pin" id="pin" readonly="true" class="form-control" />
		<% if request.querystring("nostudent") = "true" then %><p class="bg-danger lead" style="margin-top:5px;margin-bottom:5px;">Sorry, we couldn't find any active students for your pin in this location.</p><%end if%>
		<table border="0" cellspacing="0" cellpadding="0" style="width:auto;top:0;left:0;" id="keypad">
			<tr>
				<td onclick="press(1);"><button type="button" class="btn btn-primary btn-lg">1</button></td>
				<td onclick="press(2);"><button type="button" class="btn btn-primary btn-lg">2</button></td>
				<td onclick="press(3);"><button type="button" class="btn btn-primary btn-lg">3</button></td>
			</tr>
			<tr>
				<td onclick="press(4);"><button type="button" class="btn btn-primary btn-lg">4</button></td>
				<td onclick="press(5);"><button type="button" class="btn btn-primary btn-lg">5</button></td>
				<td onclick="press(6);"><button type="button" class="btn btn-primary btn-lg">6</button></td>
			</tr>
			<tr>
				<td onclick="press(7);"><button type="button" class="btn btn-primary btn-lg">7</button></td>
				<td onclick="press(8);"><button type="button" class="btn btn-primary btn-lg">8</button></td>
				<td onclick="press(9);"><button type="button" class="btn btn-primary btn-lg">9</button></td>
			</tr>
			<tr>
				<td onclick="press('B');"><button type="button" class="btn btn-primary btn-lg">&lt;</button></td>
				<td onclick="press(0);"><button type="button" class="btn btn-primary btn-lg">0</button></td>
				<td onclick="press('X');"><button type="button" class="btn btn-primary btn-lg">X</button></td>
			</tr>
		</table>
		<br />
		<input type="Submit" value="Lookup" class="btn btn-primary btn-lg" />
		</center>
		<script>
			function press(x)
			{
				if (x == "B")
				{
					var curVal = $("#pin").val();
					$("#pin").val(curVal.substring(0,curVal.length-1));
				}
				else if (x == "X")
				{
					$("#pin").val("");
				}
				else
				{
					$("#pin").val($("#pin").val() + x);
				}
			}
		</script>
	</form>
	<%
End Sub
Sub ShowLocationPicks( ByVal iLocationId )
	Dim sSql, oRs

	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & Session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""locationid"" class=""form-control"">"
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("locationid") & """ "
			If CLng(oRs("locationid")) = CLng(iLocationId) Then 
				response.write " selected=""selected"" "
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

	badpin = false

	sSQL = "SELECT l.attendeeuserid as studentid,c.locationid,p.userid as parentid, " _
		& " ISNULL(p.userfname,'') + ' ' + ISNULL(p.userlname,'') as parentname, " _
		& " ISNULL(att.userfname,'') + ' ' + ISNULL(att.userlname,'') as studentname, " _
		& " loc.name as locationname " _
		& " FROM egov_class c " _
		& " INNER JOIN egov_class_list l ON l.classid = c.classid " _
		& " INNER JOIN egov_users att ON att.userid = l.attendeeuserid " _
		& " INNER JOIN egov_users p ON p.familyid = att.familyid " _
		& " INNER JOIN egov_class_location loc ON loc.locationid = c.locationid " _
		& " WHERE c.locationid = '" & dbsafe(request.querystring("locationid")) & "' " _
		& " and c.orgid = " & session("orgid") _
		& " and startdate <= GETDATE() and enddate >= GETDATE() " _
		& " AND (status = 'ACTIVE'  OR status = 'WAITLIST') " _
		& " and p.userid = '" & dbsafe(request.form("pin")) & "' " _
		& " GROUP BY l.attendeeuserid, c.locationid, p.userid, p.userfname,p.userlname, att.userfname,att.userlname,loc.name " 
	'response.write sSQL
	'response.end

	set oRs = Server.CreateObject("ADODB.RecordSet")
	on error resume next
	oRs.Open sSQL, Application("DSN"), 3, 1
	if err then
		badpin = true
	else
		if oRs.EOF then
			badpin = true
		end if
	end if
	on error goto 0
	if not badpin then
		%>
		  <form method="post" action="" class="sigPad" id="sigform" name="sigform" style="font-size:18px;">
		   <!--[if lt IE 9]><script src="flashcanvas.js"></script><![endif]-->
		<script src="jquery.signaturepad.min.js"></script>
		<link href="jquery.signaturepad.css" rel="stylesheet">
		<style>
			.sigPad
			{
				width:500px;
			}
			.sigWrapper
			{
				height:155px;
			}
			.sigPad input[type=checkbox]
			{
				width:34px;
			}
			.sigPad .checkbox label
			{
				font:inherit;
				line-height:34px;
			}
		</style>
		<input type="hidden" name="locationid" value="<%=oRs("locationid")%>" />
		<input type="hidden" name="parentid" value="<%=oRs("parentid")%>" />
		<input type="hidden" name="parentname" value="<%=oRs("parentname")%>" />
		<input type="hidden" name="locationname" value="<%=oRs("locationname")%>" />
		<input type="hidden" name="checkinouttype" value="1" />

		<b>Student<% if oRs.RecordCount > 1 then %>s<%end if%>:<br /></b>
		<%
		Do While Not oRs.EOF
			response.write "<div class=""checkbox""><label><input type=""checkbox"" id=""studentid"" name=""studentid"" value=""" & oRs("studentid") & """ class=""form-control studentid"">&nbsp;&nbsp;" & oRs("studentname") & "</label></div>"
			oRs.MoveNext
		loop
		oRs.MoveFirst
		%>
		<b>Location:</b> <%=oRs("locationname")%><br />
		<b>Parent Name:</b> <%=oRs("parentname")%><br />
		<br />
		<ul class="sigNav">
      		<li class="drawIt" style="font-size:inherit"><a href="#draw-it" >Signature</a></li>
      		<li class="clearButton"><a href="#clear">Clear</a></li>
    		</ul>
		<div class="sig sigWrapper">
      			<canvas class="pad" width="498" height="150"></canvas>
      			<input type="hidden" name="output" class="output">
    		</div>
		</form>
		<br />
		<input type="button" value="Check In/Out" class="btn btn-primary btn-lg" onclick="validate();" />
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="attendance.asp?locationid=<%=request.querystring("locationid")%>" class="btn btn-primary btn-lg" >Start Over</a>
		<script>
    			$(document).ready(function() {
      				$('.sigPad').signaturePad({drawOnly:true});
    				});

			function validate()
			{
				if (!$('input.studentid').is(':checked'))
				{
					alert("You must select one student");
				}
				else
				{
					if ($('.output').val() != '' && $('.output').val() != null) { 
						document.sigform.submit(); 
					}
					else
					{
						alert("You must sign the form");
					}
				}
			}
  		</script>
		<%
	else
		response.redirect "attendance.asp?locationid=" & request.querystring("locationid") & "&nostudent=true"
	end if
	oRs.Close
	Set oRs = Nothing
End Sub
Sub ProcessSignature
	%>
	<h2>The following was checked in/out:</h2>
	<%
	'Insert Into  DB and show page
	arrStudents = split(dbsafe(request.form("studentid")),",")
	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.Open Application("DSN")
	for each x in arrStudents
		'response.write x
		studentname = GetStudentName(x)

		checkinouttype = dbsafe(request.form("checkinouttype"))
		locationid = dbsafe(request.form("locationid"))
		studentid = x
		parentid = dbsafe(request.form("parentid"))
		parentname = dbsafe(request.form("parentname"))
		locationname = dbsafe(request.form("locationname"))
		parentsignature = dbsafe(request.form("output"))

		sSQL = "SELECT c.classid,classname " _
			 & " FROM egov_class c " _
			 & " INNER JOIN egov_class_list l ON l.classid = c.classid " _
			 & " WHERE l.attendeeuserid = '" & studentid & "' " _
			 & " AND (status = 'ACTIVE' OR status = 'WAITLIST') " _
			 & " and startdate <= GETDATE() and enddate >= GETDATE()  "
		set oClasses = Server.CreateObject("ADODB.RecordSet")
		oClasses.Open sSQL, Application("DSN"), 3, 1
		Do WHile Not oClasses.EOF

			classid = dbsafe(oClasses("classid"))
	
			checkinout = InOrOut(x, classid)
  			datGMTDateTime = DateAdd("h",5,now())
			checkinoutdate = DateAdd("h",GetTimeOffset(session("orgid")),datGMTDateTime)

			sSQL = "INSERT INTO egov_classattendance (Checkinout,checkinouttype,classid,locationid,studentid,parentid,studentname,parentname,locationname,parentsignature,checkinoutdate) " _
				& " VALUES('" & checkinout & "','" & checkinouttype & "','" & classid & "','" & locationid & "','" & studentid & "','" & parentid & "','" & studentname & "','" & parentname & "','" & locationname & "','" & parentsignature & "','" & checkinoutdate & "')"
			'response.write sSQL & "<br />"
			oCmd.Execute(sSQL)
			Response.write "<h3>" & studentname & " - checked " & checkinout & " for " & oClasses("classname") & "</h3><br />"

			oClasses.MoveNext
		loop
		oClasses.Close
		Set oClasses = Nothing

	next
	oCmd.Close
	Set oCmd = Nothing

	%>
	<a href="attendance.asp?locationid=<%=locationid%>" class="btn btn-lg btn-primary" >Done</a>
	<%
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
