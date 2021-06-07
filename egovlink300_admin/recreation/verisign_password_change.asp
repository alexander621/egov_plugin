<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: verisign_password_change.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "rec pay passwords" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 


' IF FORM POST PROCESS RESULTS
Dim sMsg
If request.servervariables("REQUEST_METHOD") = "POST" Then
	ProcessPasswordChange
	sMsg = "<p style=""Background-color:blue;color:white;"" >The password was updated at " & NOW() & "</p>"
End If
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script>
	<!--
		function confirmpass()
		{
			if (document.frmpassword.vpass.value == document.frmpassword.confirmvpass.value) 
			{
				document.frmpassword.submit();
			}
			else
			{
				alert('The passwords entered do not match. Please correct and try again!');
			}
		}
	//-->
	</script>
</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<!--BEGIN PAGE CONTENT-->

		<%=sMsg%>

		<p>
			<font size="+1">Please enter the new <strong>VERISIGN PAYMENT GATEWAY PASSWORD</strong> and confirm the password below:</font><br />
		</p>

		<p>
			<form name="frmpassword" action="verisign_password_change.asp" method="post">
				<div class="shadow">
					<table class="tableadmin" cellpadding="3" cellspacing="1" border="0">
						<tr>
							<td align="right"><b>Enter new Password: </b></td>
							<td><input name="vpass" type="password" style="" /></td>
						</tr>
						<tr>
							<td align="right"><b>Confirm new Password: </b></td>
							<td><input type="password" name="confirmvpass"  style="" /></td>
						</tr>
						<tr>
							<td align="center" colspan="2"><input type="button" class="button" value="Change" onClick="confirmpass();" ></td>
						</tr>
					</table>
				</div>
			</form>
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
' SUB PROCESSPASSWORDCHANGE()
'--------------------------------------------------------------------------------------------------
Sub ProcessPasswordChange()
	Dim sSQL
	
	' UPDATE VERISIGN PASSWORD 
	sSQL = "UPDATE egov_verisign_options SET password = '" & TRIM(request("vpass")) &"' WHERE orgid = " & session("orgid")
	
	RunSQL sSql 

End Sub


'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 

%>

