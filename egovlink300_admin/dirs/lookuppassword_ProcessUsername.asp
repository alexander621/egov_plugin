<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
response.end
	PageIsRequiredByLogin = True 
%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp"-->

<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: lookuppassword_ProcessUsername.asp
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
Dim cmd, conn, email, username, password, strHost, objMail, Name, SenderEmail, Subject, Recipient, Body, oCdoMail, oCdoConf
If Not IsEmpty(request.form("emailstart")) Then

	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = Application("DSN")
	cmd.commandtext = "GetUserPasswordFromUsername"
	cmd.CommandType = 4
	cmd.parameters(1) = Left(request.form("username"),20)
	cmd.parameters(4) = iOrgID
	cmd.execute

	Select Case cmd.parameters(0)
		Case -2 
			response.write langNoUserName
		Case 0
			response.write langLookupPasswordByUsernameFailed
		Case 1
			password = cmd.parameters(2)
			username = cmd.parameters(1)
			email = cmd.parameters(3)

			' Get the recipients mailbox from a form.
			If Trim(email) <> "" And InStr(email,"@") Then 

				sendEmail "", email, "", "E-Gov Lost Password Request", "Username: " & username & "<br />" & vbcrlf & "Password: " & password, clearHTMLTags(sMsg), "Y"

				response.write "<br /><br /><br />Your username and password were sent to the email address we have in our system.<br /><br /> If you do not recieve an email from us shortly, please contact your E-gov system administrator."
			Else 
				response.write "<br /><br /><br />Sorry, we were unable to find you credentials in our system. Please contact your E-gov system administrator."
			End If 

	End Select 

	Set cmd = Nothing 
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
