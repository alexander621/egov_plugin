<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../egovlink300_global/includes/inc_passencryption.asp" //-->
<% 
'response.end
	PageIsRequiredByLogin = True 
%>

<!-- #include file="includes/common.asp" //-->

<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: lookuppassword.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  page where admin user can lookup their password.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	???? - INITIAL VERSION
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

sLevel = "" ' Override of value from common.asp

'Dim iorgid,iPaymentGatewayID,blnOrgRegistration,blnQuerytool,blnFaq
SetOrganizationParameters()

%>

<html>
<head>
	<title><%=langBSHome%></title>

	<link rel="stylesheet" type="text/css" href="global.css" />

	<script language="JavaScript">
	<!--
	//-->
	</script>


</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs 0,0%>
  <% ShowHeader sLevel %>

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" width="100%"  class="start" >
    <tr>
      <td valign="top" width='151'> &nbsp;
		 <!--<center> <img src='../images/icon_directory.jpg'></center>-->
	 <br>
	      
      </td>
      <td colspan="2" valign="top">
<%

  if request.servervariables("request_method") = "POST" then
	resetPassword()
  else
    	displayResetForm()
  end if

%>

 </td>
  <td width="200">&nbsp;</td>
    </tr>
 </table>

 </div>
 </div>

<!--#Include file="admin_footer.asp"-->  
  

</body>
</html>
 



<%


'------------------------------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()

	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 0) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function
sub resetPassword()

'store new password, verifying that the userid matches the key provided
sSQL = "SELECT userid,orgid FROM users WHERE orgid = " & request("orgid") & " AND isdeleted = 0 AND pwresetdate >= '" & dateadd("h",-2,now()) & "' AND pwresetkey = '" & request("key") & "'"
set oReset = Server.CreateObject("ADODB.Recordset")
oReset.Open sSQL, Application("DSN"), 3, 1
if oReset.EOF then
	%><h2>Sorry, we couldn't reset your password.  Your link may have expired.</h2><%
else

	'Encrypt New Password
	newpassword = createHashedPassword(request("password"))



	'sSQL = "UPDATE egov_users SET userpassword = '" & request("password") & "', pwresetkey = NULL, pwresetdate = NULL WHERE userid = " & oReset("userid")
	sSQL = "UPDATE users SET epassword = '" & newpassword & "', password = NULL, pwresetkey = NULL, pwresetdate = NULL WHERE userid = " & oReset("userid")
     	RunSQLStatement(sSQL)

	%><h2>Your password has been reset</h2><p><a href="login.asp">Return to login</a><%
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
sSQL = "SELECT userid,orgid FROM users WHERE orgid = " & sOrgID & " AND isdeleted = 0 AND pwresetdate >= '" & dateadd("h",-2,now()) & "' AND pwresetkey = '" & request("key") & "'"
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
