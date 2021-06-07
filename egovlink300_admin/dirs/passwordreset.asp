<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
	PageIsRequiredByLogin = True 
%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp"-->
<!--#include file="../../egovlink300_global/includes/inc_passencryption.asp"-->

<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: passwordreset.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  page where admin user can lookup their password by using their username.
'
' MODIFICATION HISTORY
' 1.0	??/??/????	???? - INITIAL VERSION
' 2.0	10/14/2011	Steve Loar - Changed to block those flagged as deleted
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

sLevel = "../" ' Override of value from common.asp

SetOrganizationParameters()

%>

<html>
<head>

  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../global.css" />

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<% ShowHeader sLevel %>

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" width="100%" class="start">
    <tr>
      <td valign="top" width='151'> &nbsp;

	 <br />
       </td>
      <td colspan="2" valign="top">

<font size="+1"><b><%=langLookupPassword%></b></font>
	  <br><img src='../images/arrow_back.gif' align='absmiddle'> <a href='../login.asp'><%=langGoBack%></a>

<%
If Not IsEmpty(request.form("username")) Then
	strUserName = dbsafe(request.form("username"))
	sSQL = "SELECT userid,username,email FROM users WHERE orgid = '" & iorgid & "' and isdeleted = 0 AND (username = '" & strUserName & "' OR email = '" & strUserName & "')"
	'response.write sSQL
	'response.end
	Set oU = Server.CreateObject("ADODB.RecordSet")
	oU.Open sSQL, Application("DSN"), 3, 1
	if not oU.EOF then
		'Generate Key
     		key = bytesToHex(sha256hashBytes(stringToUTFBytes(GenerateRandomPassword())))

     		'save random key and datetime into user's record
     		sSQL = "UPDATE users SET pwresetkey = '" & key & "',pwresetdate = '" & now() & "' WHERE userid = " & oU("userid")
     		RunSQLStatement(sSQL)


		strBody = "Reset your password for """ & oU("username") & """. <a href=""" & session("egovclientwebsiteurl") & "/admin/resetpassword.asp?key=" & key & """>Click here to reset your password</a>. This link is valid for 2 hours."

		sendEmail "", oU("email"), "", "E-Gov Password Reset Request", strBody , strBody, "Y"
		response.write "<br /><br /><br />Password reset instructions were sent to the email address we have in our system.<br /><br /> If you do not recieve an email from us shortly, please contact your E-gov system administrator."
	else
		response.write "<br /><br /><br />Sorry, we were unable to find you credentials in our system. Please contact your E-gov system administrator."
	end if
	oU.Close
	Set oU = Nothing

else
	response.write "<br /><br /><br />Sorry, we were unable to find you credentials in our system. Please contact your E-gov system administrator."
End If
%>


 </td>
  <td width='200'>&nbsp;</td>
    </tr>
 </table>
</div>
 </div>

<!--#Include file="../admin_footer.asp"-->  
  
</body>
</html>
