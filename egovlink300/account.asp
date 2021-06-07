
	<!-- #include file="includes/common.asp" //-->
	<!-- #include file="includes/start_modules.asp" //-->
	<% Dim sError 
	good = ""
	bad = ""
	%>

<script>
function validateEmail()
{
	var msg="";
	if(!isValidEmail(document.chgemail.email.value))
	{
		msg+="The current email address you have entered is not valid.\n";
	}

	if((!isValidEmail(document.chgemail.newemail.value)) || (document.chgemail.newemail.value == ''))
	{
		msg+="The new email address you have entered is not valid.\n";
	}
	if(document.chgemail.newemail.value != document.chgemail.newemail2.value)
	{
		msg+="The new email addresses you have entered do not match.\n";
	}
	if(msg != "")
	{
		msg="Your form could not be submitted for the following reasons.\n\n" + msg;
		alert(msg);
	}
	else
	{
		document.chgemail.submit();
	}
}

function validatePass()
{
	var msg="";
	if(!isValidPassword(document.chgpass.password.value))
	{
		msg+="You must enter a current password.\n";
	}
	if((!isValidPassword(document.chgpass.newpassword.value)) && !(document.chgpass.newpassword.value == ''))
	{
		msg+="The password must be 6-10 alphanumeric characters only.\n";
	}
	if(document.chgpass.newpassword.value != document.chgpass.newpassword2.value)
	{
		msg+="The passwords you have entered do not match.\n";
	}
	if(msg != "")
	{
		msg="Your form could not be submitted for the following reasons.\n\n" + msg;
		alert(msg);
	}
	else
	{
		document.chgpass.submit();
	}
}

function isValidEmail(str)
{
	var exp=/[\w-]{1,}@{1}[\w-]{1,}\.{1}\w{1,}/;
	if(str.search(exp) == -1)
	{
		return false;
	}
	return true;
}

function isValidPassword(str)
{
	//var exp=/(\w{0,5}) | (\W{1,}) | (\w{11,})/;
	var exp=/\W{1,}/;
	if(str.search(exp) == -1 && str.length >= 6 && str.length <= 10)
	{
		return true;
	}
	return false;
}

</script>
<% if Request.ServerVariables("REQUEST_METHOD") = "POST" then 
	if request("password") <> "" Then 
		' from the above Javascript the max length is 10 although it can handle 50
		If Len(request("password")) > 10 Then
			bad = bad & "Your password is too long.<br />"
		Else 
			changepassword()
		End If 
	end if
	if request("email") <> "" Then 
		changeemail()
	end if
end if %>
<html>
<head>
<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>
<link rel="stylesheet" href="css/styles.css" type="text/css">

	<link href="global.css" rel="stylesheet" type="text/css">
	<script language="Javascript" src="scripts/modules.js"></script>
  
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->

<TR><TD VALIGN=TOP>
	<div class=title>ACCOUNT ADMINISTRATION</div>
		<% printPage() %>
   
<!--#Include file="include_bottom.asp"-->    

<% function changepassword()
	sSQL = "SELECT userid from egov_users WHERE userpassword LIKE '" & request("password") & "' AND userid='" & request.cookies("userid") & "'"
		'response.write sSQL
	Set login = Server.CreateObject("ADODB.Recordset")
	login.Open sSQL, Application("DSN"), 3, 1
	if login.recordcount <> 0 then
		'response.write "HERE"
		sSQL = "UPDATE egov_users SET userpassword = '" & request("newpassword") & "' WHERE userid='" & request.cookies("userid") & "' AND userpassword LIKE '" & request("password") & "'"
		set conn = server.createobject("adodb.connection")
		ConnectionString = Application("DSN")
		conn.open ConnectionString
		conn.Execute(sSQL)
		good = good & "Your Password has been changed<br>"
	else
		'response.write sSQL
		bad = bad & "Your current password is incorrect<br>"
	end if

end function %>

<% function changeemail()
	sSQL = "SELECT userid from egov_users WHERE useremail LIKE '" & request("email") & "' AND userid='" & request.cookies("userid") & "'"
		'response.write sSQL
	Set login = Server.CreateObject("ADODB.Recordset")
	login.Open sSQL, Application("DSN"), 3, 1
	iTotal = login.RecordCount
	if iTotal <> 0 then
		sSQL = "SELECT userid,useremail,userpassword from egov_users WHERE useremail LIKE '" & request("email") & "'"
		Set userexists = Server.CreateObject("ADODB.Recordset")
		userexists.Open sSQL, Application("DSN"), 3, 1
		iTotal = userexists.RecordCount
		if iTotal <>0 then
			sSQL = "UPDATE egov_users SET useremail = '" & request("newemail") & "' WHERE userid='" & request.cookies("userid") & "' AND useremail LIKE '" & request("email") & "'"
			set conn = server.createobject("adodb.connection")
			ConnectionString = Application("DSN")
			conn.open ConnectionString
			conn.Execute(sSQL)
			good = good & "Your Email Address has been changed<br>"
		else
			bad = bad & "That email address is already in use<br>"
		end if
	else
		bad = bad & "Your current Email Address is incorrect<br>"
	end if

end function %>


<% function printPage() %>
<font color=red><% =good %><% =bad %></font></b>
<br>
<table width=95% cellpadding=0 cellspacing=5 border=0>
	<form name=chgemail method=post action=account.asp>
	<tr>
		<td valign=top>
			<b>Change Email:</b>
			<div class=borderbox>
			<table>
				<tr><td align=right>Current Email:</td><td><input class=cart type=text name=email></td></tr>
				<tr><td align=right>New Email:</td><td><input class=cart type=text name=newemail></td></tr>
				<tr><td align=right>New Email:</td><td><input class=cart type=text name=newemail2></td></tr>
				<tr><td colspan=2 align=right><input class=cartbtn type=button value="CHANGE EMAIL" onclick="validateEmail();"></td></tr>
			</table>
			</div>

		</td>
	</form>
	<form name=chgpass method=post action=account.asp>
		<td valign=top>
			<b>Change Password:</b>
			<div class=borderbox>
			<table>
				<tr><td align=right>Current Password:</td><td><input class=cart type=password name=password></td></tr>
				<tr><td align=right>New Password:</td><td><input class=cart type=password name=newpassword></td></tr>
				<tr><td align=right>New Password:</td><td><input class=cart type=password name=newpassword2></td></tr>
				<tr><td colspan=2 align=right><input class=cartbtn type=button value="CHANGE PASSWORD" onclick="validatePass();"></td></tr>
			</table>
			</div>
		</td>
	</tr>
	</form>
</table>
<%end function%>
