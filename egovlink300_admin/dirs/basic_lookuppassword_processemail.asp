<meta name="viewport" content="width=device-width, initial-scale=1" />
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<% 
	PageIsRequiredByLogin = True 
%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="dir_constants.asp"-->

<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: lookuppassword_ProcessEmail.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2005 eclink, inc.
'			 All Rights Reserved.
'
' Description:  page where admin user can lookup their password by using their email address.
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


  <table border="0" cellpadding="10" cellspacing="0"  class="start"  style="margin-left:auto;margin-right:auto;">
    <tr>
      <td colspan="2" valign="top">

<font size="+1"><b><%=langLookupPassword%></b></font>
	  <br><img src='../images/arrow_back.gif' align='absmiddle'> <a href='../basic_login.asp?token=<%=request.querystring("token")%>'><%=langGoBack%></a>

<%
dim cmd,conn, email, username,password,strHost,objMail

If not isempty(request.form("emailstart")) Then
	Dim Name
	Dim SenderEmail
	Dim Subject
	Dim Recipient
	Dim Body, oCdoMail, oCdoConf, sMsg


	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open application("DSN")
	set cmd=Server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection=conn
				cmd.commandtext="GetUserPasswordFromemail"
				cmd.commandtype=&H0004
				cmd.parameters(1)=request.form("email")
				cmd.parameters(4)=iOrgID
				cmd.execute


	select case cmd.parameters(0)
		case -2 
			' NO EMAIL SPECIFIED
			response.write langNoEmail
		case 0
			' FAILED TO FIND SPECIFIED EMAIL
			response.write langLookupPasswordByEmailFailed
		case 1
			' EMAIL FOUND - RETURN RESULTS
			email = cmd.parameters(1)
			username = cmd.parameters(2)
			password = cmd.parameters(3)

			' Get the recipients mailbox from a form.
			If trim(email)<>"" and instr(email,"@") Then
				sMsg = "Username: " & username & vbcrlf & "<br />Password: " & password

				sendEmail "", email, "", "E-Gov Lost Password Request", sMsg, "", "N"
			

				response.write "<br /><br /><br />Your username and password were sent to the email address we have in our system.<br /><br /> If you do not recieve an email from us shortly, please contact your E-gov system administrator."
			else
				response.write "<br /><br /><br />Sorry, we were unable to find your email address. Please contact your E-gov system administrator."
			end if

			'ObjMail.close
			Set ObjMail= Nothing
	end select


	set cmd=nothing
	conn.close
	set conn=nothing
%>
<%End IF%>

 </td>
    </tr>
 </table>
</div>
 </div>

  

</body>
</html>
