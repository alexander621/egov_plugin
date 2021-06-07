<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewteamdetails.asp
' AUTHOR: Steve Loar
' CREATED: 03/10/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Displays team details in a popup called from regattamembersignup.asp
'
' MODIFICATION HISTORY
' 1.0	3/10/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRegattaTeamId, sTeamName, sCaptainname, sCaptainaddress, sCaptaincity, sCaptainstate, sCaptainzip
Dim sCaptainphone

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "regatta registration", sLevel	' In common.asp

iRegattaTeamId = CLng(request("regattateamid"))

GetTeamInformation iRegattaTeamId

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />


	<script language="Javascript" src="tablesort.js"></script>
	<script language="Javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		function doClose()
		{
			window.close();
			window.opener.focus();
		}

	//-->
	</script>
</head>
<body>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<p>
		<input type="button" class="button" value="Close" onclick="doClose();" /> 
	</p>

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>River Regatta Team Details</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

	<p>
		<font size="+1"><strong><%=sTeamName%></strong></font><br />
	</p>

	<table id="captaindata" cellpadding="3" cellspacing="0" border="0">
		<tr>
			<td valign="top" id="captainlabel"><strong>Captain:</strong><td>
			<td>
				<%=sCaptainname%><br />
				<%=sCaptainaddress%><br />
				<%=sCaptaincity%>, <%=sCaptainstate%>&nbsp;<%=sCaptainzip%><br />
				<%=sCaptainphone%>
			</td>
		</tr>
	</table>

	<p>
		<% ShowTeamMembers iRegattaTeamId	%>
	</p>

	</div>
</div>

<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub GetTeamInformation( iRegattaTeamId )
'--------------------------------------------------------------------------------------------------
Sub GetTeamInformation( iRegattaTeamId )
	Dim sSql, oRs

	sSql = "SELECT regattateam, captainname, captainaddress, captaincity, captainstate, captainzip, captainphone "
	sSql = sSql & " FROM egov_regattateams WHERE orgid = " & session("orgid") & " AND regattateamid = " & iRegattaTeamId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sTeamName = oRs("regattateam")
		sCaptainname = oRs("captainname")
		sCaptainaddress = oRs("captainaddress")
		sCaptaincity = oRs("captaincity")
		sCaptainstate = oRs("captainstate")
		sCaptainzip = oRs("captainzip")
		sCaptainphone = formatphonenumber(oRs("captainphone"))
	Else
		sTeamName = ""
		sCaptainname = ""
		sCaptainaddress = ""
		sCaptaincity = ""
		sCaptainstate = ""
		sCaptainzip = ""
		sCaptainphone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowTeamMembers( iRegattaTeamId )
'--------------------------------------------------------------------------------------------------
Sub ShowTeamMembers( iRegattaTeamId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT regattateammember, isteamcaptain FROM egov_regattateammembers "
	sSql = sSql & " WHERE regattateamid = " & iRegattaTeamId 
	sSql = sSql & " AND orgid = " & session("orgid") & " ORDER BY regattateammember"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write  vbcrlf & "<div class=""shadow"">" 
		response.write vbcrlf & "<table id=""regattateamlist"" cellpadding=""5"" cellspacing=""0"" border=""0"">" 
		response.write vbcrlf & "<tr><th>Team Members</th></tr>"
		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
		  	response.write vbcrlf & "<tr id=""" & iRowCount & """"
   			If iRowCount Mod 2 = 0 Then 
			    	response.write " class=""altrow"" "
   			End If 
			response.write "><td>" & oRs("regattateammember")
			If oRs("isteamcaptain") Then
				response.write " &nbsp; (Captain)"
			End If 
			response.write "</td></tr>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>" 
	Else
		response.write "<p>No members could be found for this team.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>