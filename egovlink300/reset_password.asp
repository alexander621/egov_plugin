<!DOCTYPE html>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_passencryption.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: forgot_password.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module sends a registered citizen their password.
'
' MODIFICATION HISTORY
' 1.0   ??
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim sError
%>
<html>
<head>
<title>E-Gov Services <%=sOrgName%></title>

<link type="text/css" rel="stylesheet" href="css/styles.css" />
<link type="text/css" rel="stylesheet" href="global.css" />
<link type="text/css" rel="stylesheet" href="css/style_<%=iorgid%>.css" />

<style type="text/css">
.fieldset,
.fieldset_doesnotexist
{
   margin: 10px;
   padding: 10px;
   border-radius: 6px;
}

.fieldset legend
{
   padding: 4px 8px;
   border: 1pt solid #808080;
   border-radius: 6px;
   color: #800000;
}

.fieldset_doesnotexist
{
   color: #800000;
   font-size: 1.25em;
}

#email
{
   width: 300px;
}

#passwordText
{
   margin: 5px 0px 10px 0px;
}

#buttonLookup,
#buttonLogin
{
   cursor: pointer;
}
</style>

</head>

<!--#Include file="include_top.asp"-->
<%
  if request.servervariables("request_method") = "POST" then
	resetPassword()
  else
    	displayResetForm()
  end if

sub resetPassword()

'store new password, verifying that the userid matches the key provided
sSQL = "SELECT userid,orgid FROM egov_users WHERE orgid = " & request("orgid") & " AND isdeleted = 0 AND pwresetdate >= '" & dateadd("h",-2,now()) & "' AND pwresetkey = '" & request("key") & "'"
set oReset = Server.CreateObject("ADODB.Recordset")
oReset.Open sSQL, Application("DSN"), 3, 1
if oReset.EOF then
	%><h2>Sorry, we couldn't reset your password.  Your link may have expired.</h2><%
else

	'Encrypt New Password
	newpassword = createHashedPassword(request("password"))



	'sSQL = "UPDATE egov_users SET userpassword = '" & request("password") & "', pwresetkey = NULL, pwresetdate = NULL WHERE userid = " & oReset("userid")
	sSQL = "UPDATE egov_users SET password = '" & newpassword & "', userpassword = NULL, pwresetkey = NULL, pwresetdate = NULL WHERE userid = " & oReset("userid")
     	RunSQLStatement(sSQL)

	%><h2>Your password has been reset</h2><%
end if
oReset.Close
set oReset = nothing
end sub

sub displayResetForm()
 
  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

'Look for existing key
sSQL = "SELECT userid,orgid FROM egov_users WHERE orgid = " & sOrgID & " AND isdeleted = 0 AND pwresetdate >= '" & dateadd("h",-2,now()) & "' AND pwresetkey = '" & request("key") & "'"
set oReset = Server.CreateObject("ADODB.Recordset")
oReset.Open sSQL, Application("DSN"), 3, 1

if oReset.EOF then
	response.write "<h2>Sorry, this password reset link is invalid.</h2>"
else%>
	<script>
		function validate()
		{
			var p1 = document.resetform.password.value;
			var p2 = document.resetform.verifypassword.value;
			if (p1 != "" && p1 == p2)
			{
				document.resetform.submit();
			}
			else
			{
				alert("Your passwords don't match or they are blank");
			}
		}
	</script>
	<h2>Password Reset Form</h2>
	<form method="POST" name="resetform">
		<input type="hidden" name="userid" value="<%=oReset("userid") %>" />
		<input type="hidden" name="orgid" value="<%=oReset("orgid") %>" />
		<input type="hidden" name="key" value="<%=request("key") %>" />
		New Password: <input name="password" type="password" size="20" />
		<br />
		<br />
		Confirm Password: <input name="verifypassword" type="password" size="20" />
		<br />
		<br />
	</form>
	<input type="button" value="Reset Password" onClick="validate()" />
	<br />
	<br />
	<br />
	<br />


<%end if
oReset.Close
set oReset = nothing
end sub 


%>
<!--#Include file="include_bottom.asp"-->    
