<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%	Dim sError 
	Dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets	
%>

<% if Request.ServerVariables("REQUEST_METHOD") = "POST" then 
	'check to see if user exists
	sSQL = "SELECT userid FROM egov_users WHERE useremail='" & request("egov_users_useremail") & "' AND orgid='" & iorgid & "'"
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN") , 3, 1

	if NOT oRs.EOF then
		errormsg = "<font style=""color:yellow;background-color:red;border: solid 1px #000000;""><B>&nbsp;<font color=black>Your registration could not be completed: </font> This email address already exists!&nbsp;</b></font>"
	Else
		' Add them to the egov_users table
		userid = ProcessRecords()
		response.cookies("userid") = userid
		' Insert into the Family Members table
		AddFamilyMember userid, request.form("egov_users_userfname"), request.form("egov_users_userlname"), "Yourself", Date()
		' Take them back to where they came from
		If Session("RedirectPage") <> "" Then 
			sRedirect = Session("RedirectPage") 
			Session("RedirectPage") = ""
			response.redirect sRedirect
		Else
			response.redirect GetEGovDefaultPage(iorgid)
		End If 
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

	// set the work phone
	if (document.register.skip_work_areacode.value != "" || document.register.skip_work_exchange.value != "" || document.register.skip_work_line.value != "" || document.register.skip_work_ext.value != "")
	{
		var sPhone = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value;
		if (sPhone.length < 10)
		{
			msg += "Work Phone Number must be a valid phone number or blank";
		}
		else
		{
			document.register.egov_users_userworkphone.value = document.register.skip_work_areacode.value + document.register.skip_work_exchange.value + document.register.skip_work_line.value + document.register.skip_work_ext.value;
			var rege = /^\d+$/;
			var Ok = rege.exec(document.register.egov_users_userworkphone.value);
			if ( ! Ok )
			{
				msg += "Work Phone Number must be a valid phone number or blank";
			}
		}
	}

	// set the home phone number
	document.register.egov_users_userhomephone.value = document.register.skip_user_areacode.value + document.register.skip_user_exchange.value + document.register.skip_user_line.value;

	// Process the business address if one was chosen
	var bexists = eval(document.register["skip_Baddress"]);
	if(bexists)
	{
		//See if they picked from the business dropdown and put that in the address field 
		if (document.register.skip_Baddress.selectedIndex > -1)
		{
			var belement = document.register.skip_Baddress;
			var bselectedvalue = belement.options[belement.selectedIndex].value;

			//alert( bselectedvalue );
			//  0000 is the first pick that we do not want
			if (bselectedvalue != "0000")
			{
				document.register.egov_users_userbusinessaddress.value = bselectedvalue;
				document.register.egov_users_residenttype.value = "B";
			}
		}
	}

	// Process the resident address if one was chosen - this is second to set the local resident type
	var exists = eval(document.register["skip_Raddress"]);
	if(exists)
	{
		// See if they picked from the resident dropdown and put that in the address field 
		if (document.register.skip_Raddress.selectedIndex > -1)
		{
			var element = document.register.skip_Raddress;
			var selectedvalue = element.options[element.selectedIndex].value;

			//alert( selectedvalue );
			//  0000 is the first pick that we do not want
			if (selectedvalue != "0000")
			{
				document.register.egov_users_useraddress.value = selectedvalue;
				document.register.egov_users_residenttype.value = "R";
			}
		}
	}


	if(msg != "")
	{
		msg="Your form could not be submitted for the following reasons.\n\n" + msg;
		alert(msg);
	}
	else 
	{	
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
	<form name="register" action="register3.asp" method="post">
	<input type=hidden name="columnnameid" value="userid">
	<input type=hidden name="egov_users_userregistered" value="1">
	<input type=hidden name="egov_users_orgid" value="<%=iorgid%>">
	<input type=hidden name="ef:egov_users_useremail-text/req" value="Email Address">
	<input type=hidden name="ef:egov_users_userpassword-text/req" value="Password 1">
	<input type=hidden name="ef:skip_userpassword2-text/req" value="Password 2">
	<input type=hidden name="ef:egov_users_userhomephone-text/req/phone" value="Home Phone Number">
	<input type=hidden name="ef:egov_users_userfname-text/req" value="First name">
	<input type=hidden name="ef:egov_users_userlname-text/req" value="Last name">
	<input type=hidden name="egov_users_residenttype" value="N">

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
			<font color=red>*</font> Home Phone:
		</td><td>
			<input type="hidden" value="" name="egov_users_userhomephone">
			(<input type="text" value="" name="skip_user_areacode" size="3" maxlength="3">)&nbsp;
			<input type="text" value="" name="skip_user_exchange" size="3" maxlength="3">&ndash;
			<input type="text" value="" name="skip_user_line" size="4" maxlength="4">
		</td></tr>
		<tr><td class=label align="right">
			Fax:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userfax")%>" name="egov_users_userfax" style="width:300;" maxlength="100">
		</td></tr>
<%		bHasResidentStreets = HasResidentTypeStreets( iOrgid, "R" )
		bFound = False 
		If bHasResidentStreets  Then %>
			<tr><td class=label align="right">
					Resident Street: 
				</td><td>
					<% DisplayAddresses iorgid, "R"   %>
			</td></tr>
<%		End If %>
		<tr><td class=label align="right">
			<% If bHasResidentStreets Then %>
				Street (if not listed):
			<% Else %>
				Street:
			<% End If %>
		</td><td>
			<input type="text" value="<%If Not bfound then
											response.write sAddress
										End If %>" name="egov_users_useraddress" style="width:300;" maxlength="100">
		</td></tr>
		<!--<tr><td class=label align="right">
			Street:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_useraddress")%>" name="egov_users_useraddress" style="width:300;" maxlength="100">
		</td></tr> -->
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
		<tr><td class=label align="right">
			Business Name:
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userbusinessname")%>" name="egov_users_userbusinessname" style="width:300;" maxlength="100">
		</td></tr>
<%		bHasBusinessStreets = HasResidentTypeStreets( iOrgid, "B" )
		bFound = False 
		If bHasBusinessStreets  Then %>
			<tr><td class=label align="right">
					Business Street: 
				</td><td>
					<% DisplayAddresses iorgid, "B"   %>
			</td></tr>
<%		End If %>
		<tr><td class=label align="right">
			<% If bHasBusinessStreets Then %>
				Street (if not listed):
			<% Else %>
				Business Street:
			<% End If %>
		</td><td>
			<input type="text" value="<%=request.form("egov_users_userbusinessaddress")%>" name="egov_users_userbusinessaddress" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			Work Phone:
		</td><td>
			<input type="hidden" value="" name="egov_users_userworkphone">
			(<input type="text" value="" name="skip_work_areacode" size="3" maxlength="3">)&nbsp;
			<input type="text" value="" name="skip_work_exchange" size="3" maxlength="3">&ndash;
			<input type="text" value="" name="skip_work_line" size="4" maxlength="4">&nbsp;
			ext. <input type="text" value="" name="skip_work_ext" size="4" maxlength="4">
		</td></tr>


		<!--<tr><td colspan=2>
			<font color=red>*</font>
			Denotes required field, these fields must be filled in order to complete registration. 
			Daytime phone must include area code please use either format (555)555-5555 or 555-555-5555.
		</td></tr>-->

		<tr><td colspan="2" align="right"><input class="actionbtn" type="button" value="Submit Registration Form" onClick="validate();"></td></tr>

	</table>
	

	</form>
	</div></div></div>

	<P>&nbsp;</p>
   
<!--#Include file="include_bottom.asp"-->    
<!--#Include file="includes\inc_dbfunction.asp"-->    


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' FUNCTION DisplayAddresses( iorgid, sResidenttype )
'--------------------------------------------------------------------------------------------------
Sub DisplayAddresses( iorgid, sResidenttype )
	sSQL = "SELECT * FROM egov_residentaddresses where orgid=" & iorgid & " and residenttype='" & sResidenttype & "' order by residentstreetname, residentstreetnumber"
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN") , 3, 1

	response.write "<select name=""skip_" & sResidenttype & "address"">"	
	response.write vbcrlf &  "<option value=""0000"">Please select an address...</option>"
		
	Do While NOT oAddressList.EOF 
		response.write vbcrlf & "<option value=""" &  oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")  & """>" & oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname") & "</option>"
		oAddressList.MoveNext
	Loop

	response.write "</select>"

	oAddressList.close
	Set oAddressList = Nothing 

End Sub  

'--------------------------------------------------------------------------------------------------
' FUNCTION HasResidentTypeStreets( iOrgid, sResidenttype )
'--------------------------------------------------------------------------------------------------
Function HasResidentTypeStreets( iOrgid, sResidenttype )
	sSQL = "SELECT count(residentaddressid) as hits FROM egov_residentaddresses where orgid = " & iorgid & " and residenttype = '" & sResidenttype & "'"
	Set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN") , 3, 1

	If clng(oValues("hits")) > 0 Then
		HasResidentTypeStreets = True 
	Else
		HasResidentTypeStreets = False 
	End if
	
	oValues.close
	Set oValues = nothing
End Function 
%>
