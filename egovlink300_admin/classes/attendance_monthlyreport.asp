<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: attendance_monthlyreport.asp
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
	
	
	


</head>
<body>



	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">

	<div class="boxcontent">
			<!--BEGIN: PAGE TITLE-->
				<font size="+1"><strong>Attendance Monthly Report</strong></font>
			<!--END: PAGE TITLE-->
			<%
				sDate = request.querystring("startdate")
				eDate = request.querystring("enddate")
				if sDate = "" then sDate = month(date()) & "/1/" & year(date())
				if eDate = "" then eDate = DateAdd("d",-1,DateAdd("m",1,sDate))
			%>
		<form method="GET" name="ClassListForm">
				<label for="startdate">Start Date:</label>
				<input type="text" name="startdate" id="startdate" size="10" value="<%=sDate%>" />
			    	<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
				&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
				<label for="enddate">End Date:</label>
				<input type="text" name="enddate" id="enddate" size="10" value="<%=eDate%>" />
			    	<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
				<script>
					function doCalendar(sField) 
					{
						var w = (screen.width - 350)/2;
						var h = (screen.height - 350)/2;
						eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=ClassListForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
					}
				</script>
				<input type="Submit" value="Run Report" class="btn" />
				</form>
			<%
			if request.querystring("enddate") <> "" then eDate = DateAdd("d",1,eDate)
			sSQL = "SELECT Locationname, studentname, COUNT(checkday) as days " _
				& " FROM  " _
				& " ( " _
				& " SELECT a.locationname,studentname,DatePart(yyyy,checkinoutdate) + '-' + DatePart(m,checkinoutdate) + '-' + DatePart(d,checkinoutdate) as checkday " _ 
				& " FROM egov_classattendance a " _
				& " INNER JOIN egov_class_location l ON l.locationid = a.locationid and l.orgid = '" & session("orgid") & "' " _
				& " INNER JOIN egov_checkinouttype c ON c.egov_checkinouttypeid = a.checkinouttype " _
				& " WHERE checkinoutdate > '" & sDate & "' and '" & eDate & "' > checkinoutdate " _
				& " GROUP BY locationname, studentname, DatePart(yyyy,checkinoutdate),DatePart(m,checkinoutdate),DatePart(d,checkinoutdate) " _
				& " ) a " _
				& " GROUP BY locationname, studentname "
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
			<%
			Do While Not oRs.EOF
				if sLocation <> oRs("LocationName") then
					sLocation = oRs("LocationName")
					response.write "<h3>" & sLocation & "</h3>"
					response.write "<table class=""table table-striped table-bordered"" style=""width:50%"">"
					response.write "<thead><tr><th class=""col1"">Name</th><th class=""col7"">Days</th></tr></thead><tbody>"
				end if
				%><tr>
					<td><%=oRs("StudentName")%></td>
					<td><%=oRs("Days")%></td>
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

	</div>

</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>
