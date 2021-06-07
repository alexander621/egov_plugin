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
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
	<script src="http://prevm.com/Scripts/bootstrap.min.js" type="text/javascript"></script>
	
	
	


</head>
<body>



	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<div id="content">

	<div class="boxcontent">
			<style>
				table#report td, table#report th
				{
					padding:5px;
				}
			</style>
			<!--BEGIN: PAGE TITLE-->
				<font size="+1"><strong>Participant Hours By Session</strong></font>
				<br />
				<br />
			<!--END: PAGE TITLE-->
			<%
			sSQL = "SELECT s.seasonname, COUNT(p.userid) as participants, SUM(t.totalhours) as ParticipantHours " _
  				& " FROM egov_class_time t " _
  				& " INNER JOIN egov_class c ON c.classid = t.classid " _
  				& " INNER JOIN egov_class_seasons s ON s.classseasonid = c.classseasonid " _
  				& " INNER JOIN egov_class_list p ON p.classid = c.classid and p.classtimeid = t.timeid " _
  				& " WHERE c.orgid = '" & session("orgid") & "' and p.status = 'ACTIVE' " _
  				& " GROUP BY s.seasonname " _
  				& " ORDER BY s.seasonname "
			set oRs = Server.CreateObject("ADODB.RecordSet")
			oRs.Open sSQL, Application("DSN"), 3, 1
			%><table id="report" style="width:auto;" border="1" cellspacing="0" cellpadding="5"><%
			if not oRs.EOF then
				response.write "<tr><th>Season</th><th>Particpants</th><th>Participant Hours</th></tr>"
			end if
			Do WHile not oRs.EOF
				response.write "<tr><td><b>" & oRs("SeasonName") & "</b></td><td>" & oRs("Participants") & "</td><td>" & oRs("ParticipantHours") & "</td></tr>"
				oRs.MoveNext
			loop
			oRs.Close
			Set oRs = Nothing
			%>
			</table>

	</div>

</div>

<!--#include file="../admin_footer.asp"-->  
</body>
</html>

