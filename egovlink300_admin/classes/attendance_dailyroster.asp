<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: attendance_dailyroster.asp
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
if request.servervariables("REQUEST_METHOD") = "POST" then
	ProcessChanges()
end if
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
	<script src="jquery.signaturepad.min.js"></script>
	<link href="jquery.signaturepad.css" rel="stylesheet">
	
	
	


</head>
<body>



	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">

	<div class="boxcontent">
			<!--BEGIN: PAGE TITLE-->
				<font size="+1"><strong>Attendance Daily Roster</strong></font>
			<!--END: PAGE TITLE-->
		<form method="GET" name="ClassListForm">
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
				<input type="Submit" value="Load Rosters" class="btn" />
				</form>
			<%
	sDate = request.querystring("datesearch")
	if sDate = "" then sDate = date()
			sSQL = "SELECT egov_classattendanceid,checkinoutdate,Checkinout,locationname,studentname,parentname,staffname, " _
				& " c.description as checkinouttype,comment, class.classname,ParentSignature " _
				& " FROM egov_classattendance a " _
				& " INNER JOIN egov_class_location l ON l.locationid = a.locationid and l.orgid = '" & session("orgid") & "' " _
				& " INNER JOIN egov_checkinouttype c ON c.egov_checkinouttypeid = a.checkinouttype " _
				& " INNER JOIN egov_class class ON class.classid = a.classid " _
				& " WHERE checkinoutdate > '" & sDate & "' and '" & DateAdd("d",1,sDate) & "' > checkinoutdate " _
				& " ORDER BY locationname,studentname,checkinoutdate "
			set oRs = Server.CreateObject("ADODB.RecordSet")
			oRs.Open sSQL, Application("DSN"), 3, 1
			sLocation = ""
			%>
			<style>
				.col1 {width:300px;}
				.col2 {width:50px;}
				.col3 {width:200px;}
				.col4 {width:200px;}
				.col5 {width:300px;}
				.col6 {width:300px;}
				.col7 {width:50px;}
				.col8 {width:100px;}
			</style>
			<form method="POST">
			<%
			Do While Not oRs.EOF
				if sLocation <> oRs("LocationName") then
					sLocation = oRs("LocationName")
					response.write "<h3>" & sLocation & "</h3>"
					response.write "<table class=""table table-striped table-bordered"">"
					response.write "<thead><tr><th class=""col1"">Name</th><th class=""col8"">Class</th><th class=""col2"">In/Out</th><th class=""col3"">Type</th><th class=""col4"">Date/Time</th><th class=""col5"">Done By</th><!--th>Signature</th--><th class=""col6"">Comments</th><th class=""col7"">Delete?</th></tr></thead><tbody>"
				end if
				%><tr>
					<td><%=oRs("StudentName")%></td>
					<td><%=oRs("ClassName")%></td>
					<td><%=oRs("CheckInOut")%></td>
					<td><%=oRs("CheckInOutType")%></td>
					<td><%=oRs("CheckInOutDate")%></td>
					<td><%=oRs("ParentName")%><%=oRs("StaffName")%></td>
					<!--td>
					<div class="sigPad<%=oRs("egov_classattendanceid")%> signed">
						<div class="sigWrapper">
							<canvas class="pad" width="498" height="150"></canvas>
						</div>
					</div>
					<script>
						$(document).ready(function() {
      							$('.sigPad<%=oRs("egov_classattendanceid")%>').signaturePad({displayOnly:true}).regenerate('<%=oRs("ParentSignature")%>');
    						});
					</script>
					</td-->
					<td><textarea name="comment_<%=oRs("egov_classattendanceid")%>" style="width:100%;"><%=oRs("comment")%></textarea></td>
					<td><input type="checkbox" name="deletecheck" value="<%=oRs("egov_classattendanceid")%>" class="form-control" /></td>
				</tr><%
				oRs.MoveNext
				if not oRs.EOF then
					if sLocation <> oRs("LocationName") then response.write "</tbody></table>"
				else
					response.write "</tbody></table>"
				end if
			loop
			oRs.Close
			set oRs = Nothing
			%>
			<input type="Submit" value="Save Changes" class="btn" />
			</form>

	</div>

</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>
<%
Sub ProcessChanges()
	'Get Deletions
	arrDeletions = split(dbsafe(request.form("deletecheck")),",")
	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.Open Application("DSN")
	for each x in arrDeletions
		x = trim(x)
		sSQL = "DELETE FROM egov_classattendance WHERE egov_classattendanceid = '" & x & "'"
		oCmd.Execute(sSQL)
	next

	'Update Comments
	For each item in request.form
		if instr(item,"comment_") > 0 then
			'response.write item & " - " & request.form(item) & "<br />"
			arrItemID = split(item,"_")
			itemVal = "NULL"
			if request.form(item) <> "" then itemVal = "'" & request.form(item) & "'"
			sSQL = "UPDATE egov_classattendance set comment = " & itemVal & " WHERE egov_classattendanceid = '" & arrItemID(1) & "'"
			'response.write sSQL & "<br />"
			oCmd.Execute(sSQL)
		end if
	next
	oCmd.Close
	Set oCmd = Nothing
	'response.end
End Sub
%>
