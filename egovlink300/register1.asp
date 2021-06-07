<!-- #include file="includes/common.asp" //-->
	<!-- #include file="includes/start_modules.asp" //-->
	<% Dim sError %>

<% if Request.ServerVariables("REQUEST_METHOD") = "POST" then 
	'check to see if user exists
	sSQL = "SELECT userid FROM egov_users WHERE useremail='" & request("egov_users_useremail") & "' AND orgid='" & iorgid & "'"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN") , 3, 1

	if NOT oRs.EOF then
		errormsg = "<font style=""color:yellow;background-color:red;border: solid 1px #000000;""><B>&nbsp;<font color=black>Your registration could not be completed: </font> This email address already exists!&nbsp;</b></font>"
	else
		userid = ProcessRecords()
		response.cookies("userid") = userid
		response.redirect "default.asp"
	end if
end if
%>
<html>
<head>
<title>E-Gov Services <%=sOrgName%> - New Member Registration</title>
<link rel="stylesheet" href="css/styles.css" type="text/css">
<link href="global.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">
<script language="Javascript" src="scripts/modules.js"></script>
<script language="Javascript" src="scripts/easyform.js"></script>
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}
function validate() {
	var msg="";
	if(document.register.egov_users_userpassword.value != document.register.skip_userpassword2.value)
	{
		msg+="The passwords you have entered do not match.\n";
	}
	//if(document.register.egov_users_userstate.value == "0" && (document.register.countryName.value == "us" || document.register.countryName.value == "ca" || document.register.countryName.value == "mx" ))
	//{
		//msg+="You must choose a state or province\n";
	//}
	//if(document.register.prov.value == "" && (document.register.countryName.value != "us" && document.register.countryName.value != "ca" && document.register.countryName.value != "mx" )) 
	//{
		//msg+="You must enter a value for Province/Region or enter 'none' if this doesn't apply to you\n";
	//}
	if(msg != "")
	{
		msg="Your form could not be submitted for the following reasons.\n\n" + msg;
		alert(msg);
	}
	else {	
		if (validateForm('register')) { document.register.submit(); }
	}
}
</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->

<TR><TD VALIGN=TOP>
<p>
<font class=pagetitle>Welcome to the <%=sOrgName%> New Member Registration</font> <BR>
<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%>
</font>
</p>


<div style="margin-left:25px;">
<p>Registering to use <%=sOrgName%> E-Gov Services is FREE, quick and easy to establish!</p>   

<P>You can access your transaction history. For example, history of online payments using <%=sOrgName%> E-Gov Services or requests submitted via the <%=sOrgActionName%>. </p>   

<p>You can choose to have contact information (such as an address & telephone number) saved with your membership thereby eliminating the requirement to "re-type" this information into online forms.</p>   

<div style=""margin-left:20px; "" class=box_header4><%=sOrgName%> New Member Registration: </div>
<div class=groupsmall2>
	<form name="register" action=register.asp method=post>
	<input type=hidden name=columnnameid value="userid">
	<input type=hidden name="egov_users_userregistered" value="1">
	<input type=hidden name="egov_users_orgid" value="<%=iorgid%>">
	<input type=hidden name="ef:egov_users_useremail-text/req" value="Email Address">
	<input type=hidden name="ef:egov_users_userpassword-text/req" value="Password 1">
	<input type=hidden name="ef:skip_userpassword2-text/req" value="Password 2">
	<input type=hidden name="ef:egov_users_userhomephone-text/req/phone" value="Phone Number">
	<input type=hidden name="ef:egov_users_userfname-text/req" value="First name">
	<input type=hidden name="ef:egov_users_userlname-text/req" value="Last name">

	<!--input type=hidden name="ef:egov_users_userbusiness-text/req" value="Company name">
	<input type=hidden name="ef:egov_users_useraddress-text/req" value="Address">
	<input type=hidden name="ef:egov_users_usercity-text/req" value="City">
	<input type=hidden name="ef:egov_users_userzip-text/req/zip" value="Zip"-->
	<table>
		
		
		<%
		If errormsg <> "" Then
			response.write "<tr><td colspan=2 align=right>" & errormsg & "</td></tr>"
		End If
		%>
		

		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			Email:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_useremail")%>" name="egov_users_useremail" style="width:300;" maxlength="100"><font color=red></font>
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			Password:
		</td><td>
			<input type="password" value="<%=request.form("egov_users_userpassword")%>" name="egov_users_userpassword" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			Verify Password:
		</td><td>
			<input type="password" value="<%=request.form("skip_userpassword2")%>" name="skip_userpassword2" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			First Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required"> 
			<input type="text" value="<%=request.form("egov_users_userfname")%>" name="egov_users_userfname" style="width:300;" maxlength="100">
			</span>
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span>
			Last Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required">
			<input type="text" value="<%=request.form("egov_users_userlname")%>" name="egov_users_userlname" style="width:300;" maxlength="100">
			</span>
		</td></tr>
		<tr><td class=label align="right">
			Business Name:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userbusinessname")%>" name="egov_users_userbusinessname" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			<font color=red>*</font> Daytime Phone:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userhomephone")%>" name="egov_users_userhomephone" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			Fax:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userfax")%>" name="egov_users_userfax" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			Street:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_useraddress")%>" name="egov_users_useraddress" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			City:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_usercity")%>" name="egov_users_usercity" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			State / Province:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userstate")%>" name="egov_users_userstate" size="5" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			ZIP / Postal Code:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userzip")%>" name="egov_users_userzip" style="width:300;" maxlength="100">
		</td></tr>

		<tr><td colspan=2>
			<font color=red>*</font>
			Denotes required field, these fields must be filled in order to complete registration. 
			<br>Daytime phone must include area code please use either (555)555-5555 or 555-555-5555 format.
		</td></tr>

		<tr><td colspan=2 align=right><input class=actionbtn type=button value="Submit Registration Form" onClick="validate();"></td></tr>

	</table>
	

	</form>
	</div></div></div>

	<P>&nbsp;</p>
   
<!--#Include file="include_bottom.asp"-->    
<!--#Include file="includes\inc_dbfunction.asp"-->    
