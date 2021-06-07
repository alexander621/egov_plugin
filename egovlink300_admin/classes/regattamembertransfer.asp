<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattamembertransfer.asp
' AUTHOR: Steve Loar
' CREATED: 08/03/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows the selection of a team to transfer a team member to.
'
' MODIFICATION HISTORY
' 1.0	8/03/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRegattaTeamMemberId, sTeamMemberName, iRegattaTeamId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "regatta registration", sLevel	' In common.asp

iRegattaTeamMemberId = CLng(request("regattateammemberid"))
iRegattaTeamId = CLng(request("regattateamid"))

sTeamMemberName = GetTeamMemberName( iRegattaTeamMemberId )

%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />


	<script language="Javascript" src="tablesort.js"></script>
	<script language="Javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<p>
		<input type="button" class="button" value="<< Back" onclick="javascript:location.href='regattateamlist.asp?regattateamid=<%=iRegattaTeamId%>'" />
	</p>

	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>River Regatta Team Member Transfer</strong></font><br />
	</p>

	<p>
		<strong>Transfering: <font size="+1"><%=sTeamMemberName%></font></strong>
	</p>
	<form name="frmTransfer" action="regattateammembertransferupdate.asp" method="post">
		<input type="hidden" name="regattateammemberid" value="<%=iRegattaTeamMemberId%>" />
		<input type="hidden" name="originalregattateamid" value="<%=iRegattaTeamId%>" />

		<p>
			<strong>Select the team you wish to transfer them to</strong><br />
<%			ShowRegattaTeams iRegattaTeamId		%>
		</p>

		<p>
			<input type="submit" class="button" value="Transfer Team Member" />	
		</p>

	</form>

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
' Function GetTeamMemberName( iRegattaTeamMemberId )
'--------------------------------------------------------------------------------------------------
Function GetTeamMemberName( iRegattaTeamMemberId )
	Dim sSql, oRs

	sSql = "SELECT regattateammember FROM egov_regattateammembers WHERE regattateammemberid = " & iRegattaTeamMemberId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetTeamMemberName = oRs("regattateammember")
	Else
		GetTeamMemberName = ""
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowRegattaTeams( iRegattaTeamId	)
'--------------------------------------------------------------------------------------------------
Sub ShowRegattaTeams( iRegattaTeamId )
	Dim sSql, oRs

	sSql = "SELECT regattateamid, regattateam FROM egov_regattateams WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY regattateam"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""regattateamid"" id=""regattateamid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("regattateamid") & """"
		If CLng(iRegattaTeamId) = CLng(oRs("regattateamid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("regattateam") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub



%>


