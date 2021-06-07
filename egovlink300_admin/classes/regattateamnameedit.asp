<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateamnameedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/05/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	8/05/2009	Steve Loar	-	Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRegattaTeamId, sTeamName

iRegattaTeamId = CLng(request("regattateamid"))

sTeamName = GetTeamName( iRegattaTeamId )

%>
<html>
	<head>
		<title>E-Gov Administration Console</title>

		<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
		<link rel="stylesheet" type="text/css" href="classes.css" />

		<script language="Javascript" src="../scripts/modules.js"></script>
		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="Javascript">
		<!--

			function doUpdate() 
			{
				// validate the edited name
				if ($("regattateam").value == '')
				{
					alert("The team name cannot be blank.\nPlease correct this and try saving again.");
					$("regattateam").focus();
					return;
				}

				//Do Ajax save call
				doAjax('regattateamnameupdate.asp', 'regattateamid=' + $("regattateamid").value + '&regattateam=' + $("regattateam").value, '', 'get', '0');
				
				// Update the parent window
				window.opener.document.getElementById("teamname").innerHTML = $("regattateam").value;

				// Close
				doClose();

				//standard post for testing
				//document.frmTeamName.submit();
			}

			function init()
			{
				$("regattateam").focus();
			}

			function doClose()
			{
				window.close();
				window.opener.focus();
			}

			window.onload = init; 

		//-->
		</script>
	</head>
	<body>
		<div id="content">
			<div id="centercontent">
				<font size="+1"><strong>Edit Regatta Team Name</strong></font><br /><br />
				<form name="frmTeamName" action="regattateamnameupdate.asp" method="post">
					<input type="hidden" id="regattateamid" name="regattateamid" value="<%=iRegattaTeamId%>" />
					<p> 
						<table cellpadding="5" cellspacing="0" border="0" id="regattateamnametable">
							<tr><td><input type="input" name="regattateam" id="regattateam" value="<%=sTeamName%>" size="100" maxlength="100" /></td></tr>
						</table>
						<br />
						<input type="button" class="button" id="savebutton" value="Save Changes" onclick="doUpdate();" />
							 &nbsp; &nbsp; 
						<input type="button" class="button" value="Cancel" onclick="doClose();" />
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%


'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function GetTeamName( iRegattaTeamId )
'--------------------------------------------------------------------------------------------------
Function GetTeamName( iRegattaTeamId )
	Dim sSql, oRs

	sSql = "SELECT regattateam FROM egov_regattateams WHERE orgid = " & session("orgid") & " AND regattateamid = " & iRegattaTeamId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetTeamName = oRs("regattateam")
	Else
		GetTeamName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>