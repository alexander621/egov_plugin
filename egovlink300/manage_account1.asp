<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError %>


<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" then 
		Call ProcessRecords()
		errormsg = "<font style=""color:blue;background-color:#e0e0e0;border: solid 1px #000000;""><B>&nbsp;Information Updated - " & nOW() & "&nbsp;</b></font>"
End If

' Set the session for the family update form to come back here
session("ManageURL") = "manage_account.asp"
Session("ManageLang") = "Return to Manage Account"

If Session("RedirectLang") <> "" Then
	sBackLang = Session("RedirectLang")
Else 
	sBackLang = "Back"
End If 

' USER VALUES
Dim sFirstName,sLastName,sAddress,sCity,sState,sZip,sPhone,sEmail,sFax,sBusinessName,sDayPhone,sPassword,iUserID
Dim bHasResidentStreets, bFound, sResidenttype, sBusinessAddress, bHasBusinessStreets, sWorkPhone

GetRegisteredUserValues()

%>

<html>
<head>
<title>E-Gov Services <%=sOrgName%> - Manage Account</title>

<link rel="stylesheet" href="css/styles.css" type="text/css">
<link href="global.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">

<script language="Javascript" src="scripts/modules.js"></script>
<script language="Javascript" src="scripts/easyform.js"></script>

<script language=javascript>
<!--

function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}

function UpdateFamily(iUserId)
	{
		location.href='family_members.asp?userid=' + iUserId;
	}

function Validate() {
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
		return;
	}
	else {	
		if (validateForm('register')) 
		{ 
			document.register.submit(); 
		}
	}
}

function GoBack(ReturnToURL)
{
	if (ReturnToURL != "")
	{
		location.href=ReturnToURL;
	}
	else
	{
		history.go(-1);
	}
}

//-->
</script>

</head>

<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->


<P>
<div align=left style="padding-bottom:20px;"> <% RegisteredUserDisplay() %> </div>
</p>

<div style="margin-left:25px;">

<div class=title>Manage Account</div>

<br /><br /><a href="javascript:GoBack('<%=Session("RedirectPage")%>')"><img src="images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=sBackLang%></a><br /><br />

<P><div style=""margin-left:20px; "" class=box_header4><%=sOrgName%> New Member Registration: </div>
<div class=groupsmall2>
	<form name="register" action="manage_account1.asp" method=post>
	<input type=hidden name="columnnameid" value="userid">
	<input type=hidden name="userid" value="<%=iuserid%>">
	<input type=hidden name="egov_users_orgid" value="<%=iorgid%>">
	<input type=hidden name="ef:egov_users_useremail-text/req" value="Email Address">
	<input type=hidden name="ef:egov_users_userpassword-text/req" value="Password 1">
	<input type=hidden name="ef:skip_userpassword2-text/req" value="Password 2">
	<input type=hidden name="ef:egov_users_userhomephone-text/req/phone" value="Phone Number">
	<input type=hidden name="ef:egov_users_userfname-text/req" value="First name">
	<input type=hidden name="ef:egov_users_userlname-text/req" value="Last name">
	<input type=hidden name="egov_users_residenttype" value="<%=sResidenttype%>">

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
			<input type="text" value="<%=sEmail%>" name="egov_users_useremail" style="width:300;" maxlength="100"><font color=red></font>
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			Password:
		</td><td>
			<input type="password" value="<%=sPassword%>" name="egov_users_userpassword" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			Verify Password:
		</td><td>
			<input type="password" value="<%=sPassword%>" name="skip_userpassword2" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span> 
			First Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required"> 
			<input type="text" value="<%=sFirstName%>" name="egov_users_userfname" style="width:300;" maxlength="100">
			</span>
		</td></tr>
		<tr><td class=label align="right">
			<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><font color=red>*</font></span>
			Last Name:
			</span>
		</td><td>
			<span class="cot-text-emphasized" title="This field is required">
			<input type="text" value="<%=sLastName%>" name="egov_users_userlname" style="width:300;" maxlength="100">
			</span>
		</td></tr>
<%		bHasResidentStreets = HasResidentTypeStreets( iOrgid, "R" )
		bFound = False 
		If bHasResidentStreets  Then %>
			<tr><td class=label align="right">
					Resident Street: 
				</td><td>
					<% DisplayAddresses iorgid, "R", sAddress, bFound %>
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
		<tr><td class=label align="right">
			City:
		</td><td>
			<input type="text" value="<%=sCity%>" name="egov_users_usercity" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			State / Province:
		</td><td>
			<input type="text" value="<%=sState%>" name="egov_users_userstate" size="5" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			ZIP / Postal Code:
		</td><td>
			<input type="text" value="<%=sZip%>" name="egov_users_userzip" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			<font color=red>*</font> Home Phone:
		</td><td>
			<!--<input type="text" value="<%=sDayPhone%>" name="egov_users_userhomephone" style="width:300;" maxlength="100">-->
			<input type="hidden" value="<%=sDayPhone%>" name="egov_users_userhomephone">
			(<input type="text" value="<%=Left(sDayPhone,3)%>" name="skip_user_areacode" size="3" maxlength="3">)&nbsp;
			<input type="text" value="<%=Mid(sDayPhone,4,3)%>" name="skip_user_exchange" size="3" maxlength="3">&ndash;
			<input type="text" value="<%=Right(sDayPhone,4)%>" name="skip_user_line" size="4" maxlength="4">
		</td></tr>
		<tr><td class=label align="right">
			Fax:
		</td><td>
			<input type="text" value="<%=sFax%>" name="egov_users_userfax" style="width:300;" maxlength="100">
		</td></tr>
		<!--<tr><td colspan="2" align="center"><strong>If you are not a resident, please include the following.</strong></td><tr>-->
		<tr><td class=label align="right">
			Business Name:
		</td><td>
			<input type="text" value="<%=sBusinessName%>" name="egov_users_userbusinessname" style="width:300;" maxlength="100">
		</td></tr>
<%		bHasBusinessStreets = HasResidentTypeStreets( iOrgid, "B" )
		bFound = False 
		If bHasBusinessStreets  Then %>
			<tr><td class=label align="right">
					Business Street: 
				</td><td>
					<% DisplayAddresses iorgid, "B", sBusinessAddress, bFound %>
			</td></tr>
<%		End If %>
		<tr><td class=label align="right">
			<% If bHasBusinessStreets Then %>
				Street (if not listed):
			<% Else %>
				Business Street:
			<% End If %>
		</td><td>
			<input type="text" value="<%If Not bfound then
											response.write sBusinessAddress
										End If %>" name="egov_users_userbusinessaddress" style="width:300;" maxlength="100">
		</td></tr>
		<tr><td class=label align="right">
			Work Phone:
		</td><td>
			<!--<input type="text" value="<%=sWorkPhone%>" name="egov_users_userworkphone" style="width:300;" maxlength="100">-->
			<input type="hidden" value="<%=sWorkPhone%>" name="egov_users_userworkphone">
			(<input type="text" value="<%=Left(sWorkPhone,3)%>" name="skip_work_areacode" size="3" maxlength="3">)&nbsp;
			<input type="text" value="<%=Mid(sWorkPhone,4,3)%>" name="skip_work_exchange" size="3" maxlength="3">&ndash;
			<input type="text" value="<%=Mid(sWorkPhone,7,4)%>" name="skip_work_line" size="4" maxlength="4">&nbsp;
			ext. <input type="text" value="<%=Mid(sWorkPhone,11,4)%>" name="skip_work_ext" size="4" maxlength="4">
		</td></tr>

		<tr><td colspan=2 align=right>
		<input class="actionbtn" type="button" value="Family Members" onClick="javascript:UpdateFamily(<%=iuserid%>);">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input class="actionbtn" type="button" value="Update Information" onClick="javascript:Validate();"></td></tr>

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
' FUNCTION DisplayResidentAddresses( iorgid, sAddress, bFound )
'--------------------------------------------------------------------------------------------------
Sub DisplayAddresses( iorgid, sResidenttype, sAddress, ByRef bFound )
	sSQL = "SELECT residentstreetnumber, residentstreetname FROM egov_residentaddresses_list where orgid=" & iorgid & " and residenttype='" & sResidenttype & "' order by residentstreetname, residentstreetnumber"
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN") , 3, 1

	response.write "<select name=""skip_" & sResidenttype & "address"">"	
	response.write "<option value=""0000"">Please select an address...</option>"
		
	Do While NOT oAddressList.EOF 
		response.write vbcrlf & "<option value=""" &  oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")  & """"
		If UCase(sAddress) = UCase(oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname")) Then
			bFound = True
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oAddressList("residentstreetnumber") & " " & oAddressList("residentstreetname") & "</option>"
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

'--------------------------------------------------------------------------------------------------
' FUNCTION REGISTEREDUSERDISPLAY()
'--------------------------------------------------------------------------------------------------
Function GetRegisteredUserValues()

	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
		
		sSQL = "SELECT * FROM egov_users WHERE userid = " & request.cookies("userid")
		Set oValues = Server.CreateObject("ADODB.Recordset")
		oValues.Open sSQL, Application("DSN") , 3, 1

		If NOT oValues.EOF Then
			sFirstName = oValues("userfname")
			sLastName = oValues("userlname")
			sAddress = oValues("useraddress")
			sState = oValues("userstate")
			sCity = oValues("usercity")
			sZip = oValues("userzip")
			sEmail = oValues("useremail")
			sFax = oValues("userfax")
			sBusinessName = oValues("userbusinessname")
			sPassword = oValues("userpassword")
			sDayPhone = oValues("userhomephone")
			sWorkPhone = oValues("userworkphone")
			iUserID = oValues("userid")
			If IsNull(oValues("residenttype")) Or oValues("residenttype") = "" Then
				sResidenttype = "N"
			Else 
				sResidenttype = oValues("residenttype")
			End If 
			sBusinessAddress = oValues("userbusinessaddress")
		End If

		oValues.close
		Set oValues = nothing
	Else
		' REDIRECT TO USER LOGIN
		response.redirect("user_login.asp")
	End If 

End Function
%>
