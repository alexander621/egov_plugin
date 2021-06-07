<!DOCTYPE HTML PUBLIC "-//W3C//DTD XHTML 1.1 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
	'COMMENTED OUT SINCE IT CAUSES A BAD REDIRECT
 '<!-- #include file="classes/class_global_functions.asp" //-->
%>
<%
 Dim sError

'Revert "evalform" requests to use "action.asp".
 if trim(request("actionid")) <> "" then
   	iActionID = request("actionid")

   	if IsNumeric(iActionId) then
     		iActionID = CLng(iActionID)

       response.redirect("action.asp?actionid=" & iActionID)
    else
       response.redirect("action.asp")
    end if
 else
    response.redirect("action.asp")
 end if

%>
<html>
<head>
<%
' CAPTURE CURRENT PATH
Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
Session("RedirectLang") = "Return to Action Line"
%>


<%If iorgid = 7 Then %>
	<title><%=sOrgName%></title>
<%Else%>
	<title>E-Gov Services <%=sOrgName%></title>
<%End If%>



<link rel="stylesheet" type="text/css" href="css/styles.css" />
<link rel="stylesheet" type="text/css" href="global.css" />
<link rel="stylesheet"type="text/css" href="css/style_<%=iorgid%>.css"  />

<script language="Javascript" src="scripts/modules.js"></script>
<script language="Javascript" src="scripts/easyform.js"></script>

<script language="Javascript">
<!--	
	function openWin2(url, name) {
	  popupWin = window.open(url, name,"resizable,width=500,height=450");
	}

	// Validate Tracking added by Steve Loar - 12/30/2005
	function ValidateTracking( form )
	{
		var rege = /^\d+$/;
		var Ok = rege.exec(form.REQUEST_ID.value);

		if (! Ok)
		{
			alert ("Tracking Numbers must be numeric. Please try your search again.");
			form.REQUEST_ID.focus();
			form.REQUEST_ID.select();
			return false;
		}
		return true;
	}

	function ValidateInput()
	{
		//alert(document.frmRequestAction.cot_txtDaytime_Phone.value);
		// Set the Phone number
		var Phone_exists = eval(document.frmRequestAction["cot_txtDaytime_Phone"]);
		if(Phone_exists)
		{
			document.frmRequestAction.cot_txtDaytime_Phone.value = document.frmRequestAction.skip_user_areacode.value + document.frmRequestAction.skip_user_exchange.value + document.frmRequestAction.skip_user_line.value;
		}
		// Set the Fax
		var Fexists = eval(document.frmRequestAction["cot_txtFax"]);
		if(Fexists)
		{
			document.frmRequestAction.cot_txtFax.value = document.frmRequestAction.skip_fax_areacode.value + document.frmRequestAction.skip_fax_exchange.value + document.frmRequestAction.skip_fax_line.value;
		}
		// alert(document.frmRequestAction.cot_txtDaytime_Phone.value);
		return  validateForm('frmRequestAction');
		//return true;
	}

	var isNN = (navigator.appName.indexOf("Netscape")!=-1);

	function autoTab(input,len, e) 
	{
		var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

		if(input.value.length >= len && !containsElement(filter,keyCode)) {
			input.value = input.value.slice(0, len);
		var addNdx = 1;

		while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
		{
			addNdx++;
			//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
		}

		input.form[(getIndex(input)+addNdx) % input.form.length].focus();
	}

	function containsElement(arr, ele) 
	{
		var found = false, index = 0;

		while(!found && index < arr.length)
			if(arr[index] == ele)
				found = true;
			else
				index++;
		return found;
	}

	function getIndex(input) 
	{
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
	}
		return true;
	}

//-->
</script>
</head>


<!--#Include file="include_top.asp"-->


<TR><TD VALIGN=TOP>


<!--BODY CONTENT-->
<p>
<font class=pagetitle>Welcome to the <%=sOrgName%> Action Line</font> <BR>

	<!--BEGIN:  USER REGISTRATION - USER MENU-->
	<% If sOrgRegistration Then %>
			<%  If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
					RegisteredUserDisplay("")
				Else %>
					<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
			<% End If %>
	<% Else %>
		<font class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
	<% End If%>
	<!--END:  USER REGISTRATION - USER MENU-->


</font>
</p>

<div style="margin-left:25px;">


<%
' ---------------------------------------------------------------------------------------
' BEGIN DISPLAY PAGE CONTENT
' ---------------------------------------------------------------------------------------
If trim(request("actionid")) <> "" Then 
	iActionId = request("actionid")
	If IsNumeric(iActionId) Then 
		iActionId = CLng(iActionId)
		Call subDisplayActionForm(iActionId,iorgid)
	Else
			response.redirect("action.asp")
		End If 
	Else
%>


<%
'--------------------------------------------------------------------------------------------------
' BEGIN: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
	iSectionID = 2
	sDocumentTitle = "MAIN"
	sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
	datDate = Date()	
	datDateTime = Now()
	sVisitorIP = request.servervariables("REMOTE_ADDR")
	'Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
'--------------------------------------------------------------------------------------------------
' END: VISITOR TRACKING
'--------------------------------------------------------------------------------------------------
%>


<table cellspacing="1" cellpadding="1" border="0">
	<tr valign="top">
		<td>
			<table cellspacing="1" cellpadding="1" border="0">
			     <tr valign="top">
				  <td>

					<!--BEGIN: TRACKING LOOKUP-->
				  	<p>
					<form name="frmActionLookup" action="action_request_lookup.asp" METHOD="POST" onsubmit="javascript: return ValidateTracking(this);">
					<div style=""margin-left:20px; "" class=box_header2>Check the Status of an Action Line Request	</div>
					<div class=groupsmall>
					  <table cellspacing="1" border="0"  width="300">
						<tr>
					 	  <td valign="top">
							<b>Tracking Number #: </b><input type="text" name="REQUEST_ID" size="10" style="width:80px;"> <input class=action type="submit" value="Search...">
						  </td>
						</tr>
					  </table></div></div>
					</form>
					<!--END: TRACKING LOOKUP-->



					<!--BEGIN: LIST FORMS-->
					<form action="action.asp?list=true" METHOD="POST">
					<div style=""margin-left:20px; "" class=box_header2>Create New Action Line Request	</div>
							<div class=groupsmall>
						<table cellspacing="1" border="0" width="300">
						<tr>
						  <td nowrap>
							 <% fnListForms()%>
						  </td>
						</tr>
						</table>
					</div></div>
					</form>
					<!--END: LIST FORMS-->


			  </td>
			  <td width=225 style="padding-left:15px;" valign="top">
			  	
				<!--BEGIN: FAQ BUTTON-->
				<% if iorgid = 15 then 
			  		response.write "<p><input type=""button"" onclick=""window.location='faq.asp'"" value=""Frequently Asked Questions"" class=""action_100""></p>"
			  	end if
			  	%>
				<!--END: FAQ BUTTON-->


				<!--BEGIN: REGISTER/LOGIN LINKS-->
				<%If sOrgRegistration AND (request.cookies("userid") = "" OR request.cookies("userid") = "-1") Then %>
				  <b>Personalized E-Gov Services</b>
				  <ul>
					<li><a href="user_login.asp">Click here to Login</a>
					<li><a href="register.asp">Click here to Register</a>
				  </ul>
				  <hr style="width: 90%; size: 1px; height: 1px;">
				<%End If%>
				<!--END: REGISTER/LOGIN LINKS-->


					<%=sActionDescription%></td>
			     </tr>
			</table>
		</td>
	</tr>
</table>
</div>

<% End If %>

</div>
</div>

<!--SPACING CODE-->
<p><bR>&nbsp;<bR>&nbsp;</p>
<!--SPACING CODE-->


<!--#Include file="include_bottom.asp"-->  



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' FUNCTION FNLISTFORMS()
'--------------------------------------------------------------------------------------------------
Function fnListForms()
	
	sLastCategory = "NONE_START"
	sSQL = "SELECT * FROM dbo.egov_form_list_200  WHERE ((orgid=" & iorgID & ")) AND (form_category_id <> 6) AND (action_form_internal <> 1) order by form_category_Sequence,action_form_name"

	Set oForms = Server.CreateObject("ADODB.Recordset")
	oForms.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForms.EOF Then
		
		Do while NOT oForms.EOF 

			sCurrentCategory = oForms("form_category_name")
			If sLastCategory = "NONE_START" Then
				response.write "<b><a name=""" & oForms("form_category_id")  & """>" & sCurrentCategory & "</a></b>"
			End If

			If (sCurrentCategory <> sLastCategory) AND (sLastCategory <> "NONE_START") Then
				response.write "<br><br><b><a name=""" & oForms("form_category_id")  & """>" & sCurrentCategory & "</a></b>"
			End If
			
			sTopic = Server.URLEncode(sCurrentCategory & " > " & oForms("action_form_name"))
		
			response.write "<li><a href=""action.asp?actionid=" & oForms("action_form_id") & """>" & oForms("action_form_name") &  "</a>"
			oForms.MoveNext

			sLastCategory = sCurrentCategory
		Loop

	Else

		response.write "<P style=""padding-top:10px;""><center><font  color=red><B><I>No action forms enabled.</I></B></font></P>"
	
	End If

	Set oForms = Nothing 

End Function


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYACTIONFORM(IFORMID,IORGID)
'--------------------------------------------------------------------------------------------------
Sub subDisplayActionForm(iFormID,iorgid)

	' GET FORM GENERAL INFORMATION
	Dim sTitle
	Dim sIntroText
	Dim sFooterText
	Dim sMask
	Dim blnEmergencyNote
	Dim sEmergencyText

	' GET FORM INFORMATION	
	sSQL = "SELECT * FROM egov_action_request_forms  WHERE action_form_id=" & iFormID

	Set oForm = Server.CreateObject("ADODB.Recordset")
	oForm.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForm.EOF Then
		' POPULATE DATA FROM RECORDSET
		sTitle = oForm("action_form_name")
		sIntroText = oForm("action_form_description")
		sFooterText = oForm("action_form_footer")
		sMask = oForm("action_form_contact_mask")
		blnEmergencyNote = oForm("action_form_emergency_note")
		sEmergencyText = oForm("action_form_emergency_text")
	
	Else
		response.redirect("action.asp")
	End If

	Set oForm = Nothing 
%>

<%
	'--------------------------------------------------------------------------------------------------
	' BEGIN: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
		iSectionID = 22
		sDocumentTitle = sTitle
		sURL = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
		datDate = Date()	
		datDateTime = Now()
		sVisitorIP = request.servervariables("REMOTE_ADDR")
		Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
	'--------------------------------------------------------------------------------------------------
	' END: VISITOR TRACKING
	'--------------------------------------------------------------------------------------------------
%>

<form name="frmRequestAction" action="action_eval_cgi.asp?list=true" METHOD="POST">
<input type="hidden" name="actionid" value="<%=iFormID%>">
<input type="hidden" name="actiontitle" value="<%=sTitle%>">

<div style="margin-top:20px; margin-left:20px;" >

<!--BEGIN: TITLE-->
<div style="margin-bottom: 5px"><font  class=formtitle  ><%=sTitle%>
<!--END: TITLE-->


<!--BEGIN: REGISTER/LOGIN LINKS-->
<%If sOrgRegistration AND (request.cookies("userid") = "" OR request.cookies("userid") = "-1") Then %>
	 - <a href="user_login.asp">Click here to Login</a> | 
	<a href="register.asp">Click here to Register</a> 
<%End If%>
<!--END: REGISTER/LOGIN LINKS-->
</font> </div>



<!--BEGIN: EMERGENCY NOTE-->
<%If blnEmergencyNote Then%>
<div class=warning><%=sEmergencyText%></div>
<%End If%>
<!--END: EMERGENCY NOTE-->


<div class=group>


<div class="orgadminboxf">

	<!--BEGIN: INTRO INFORMATION -->
		<P>
		<%If sIntroText <> "" Then
			response.write sIntroText
		Else
			response.write " - <i> Introduction text is currently blank </i> -"
		End If
		%></P>
	<!--END: INTRO INFORMATION -->

	
	<P>
	<!--BEGIN: CLASS INFORMATION-->
		<% DisplayItem request("classid") %>
	<!--END: CLASS INFORMATION-->
	</p>


	<!--BEGIN: CONTACT INFORMATION -->
		<P>
			<b><u>Contact Information:</u></b><br>
			<table>
			<%DrawContactTable(sMask)%>
			</table>
		</p>
	<!--END: CONTACT INFORMATION -->

	
	<!--BEGIN: FORM FIELD INFORMATION -->
		<P><% Call subDisplayQuestions(iFormID,sMask) %> </P>

		<p><font color=red>*</font> <B><i>Information is required.</i></b></P>
	<!--END: FORM FIELD INFORMATION -->
	
	
	<!--BEGIN: ENDING NOTES -->
		<P>
		<%If sFooterText <> "" Then
			response.write sFooterText
		Else
			response.write " - <i> Footer text is currently blank </i> -"
		End If
		%>
		</P>
	<!--END: ENDING NOTES -->

<%response.write "</form>"%>

<% End Sub %>

<%
'--------------------------------------------------------------------------------------------------
' FUNCTION DRAWCONTACTTABLE(sMASK)
'--------------------------------------------------------------------------------------------------
Function DrawContactTable(sMASK)

' BEGIN: GET USER PERSONNEL INFORMATION IF USER IS LOGGED INTO WEBSITE
If sOrgRegistration Then 
	If request("userid") <> "" and request("userid") <> "-1" Then
		
		iUserID = request("userid")
	
		sSQL = "SELECT * FROM egov_users WHERE userid=" & iUserID
		Set oInfo = Server.CreateObject("ADODB.Recordset")
		oInfo.Open sSQL, Application("DSN") , 3, 1

		If NOT oInfo.EOF Then
			' USER FOUND SET VALUES
			sFirstName = oInfo("userfname")
			sLastName = oInfo("userlname")
			sAddress = oInfo("useraddress")
			sCity = oInfo("usercity")
			sState = oInfo("userstate")
			sZip = oInfo("userzip")
			sEmail = oInfo("useremail")
			sHomePhone = oInfo("userhomephone")
			sWorkPhone = oInfo("userworkphone")
			sBusinessName = oInfo("userbusinessname")
			sFax = oInfo("userfax")

		Else
			' USER NOT FOUND SET VALUES TO EMPTY
			sFirstName = ""
			sLastName = ""
			sAddress = ""
			sCity = ""
			sState = ""
			sZip = ""
			sEmail = ""
			sHomePhone = ""
			sWorkPhone = ""
			sFax = ""
			sBusinessName = ""

		End If

		Set oInfo = Nothing

	End If
End If
' END: GET USER PERSONNEL INFORMATION IF USER IS LOGGED INTO WEBSITE
%>

	
	<% If IsDisplay(sMASK,1) Then %>
	
	<%
	If IsRequired(sMASK,1) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtFirst_Name-text/req"" value=""First Name"">"
	End If
	%>

	<tr><td align="right">
		<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><%=IsRequired(sMASK,1)%></span> 
		First Name:
		</span>
	</td><td>
		<span class="cot-text-emphasized" title="This field is required"> 
		<input type="text" value="<%=sFirstName%>" name="cot_txtFirst_Name" id="txtFirst_Name" style="width:300;" maxlength="100">
		</span>
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,2) Then %>

	<%
	If IsRequired(sMASK,2) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtLast_Name-text/req"" value=""Last Name"">"
	End If
	%>

	<tr><td align="right">
		<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><%=IsRequired(sMASK,2)%></span>
		Last Name:
		</span>
	</td><td>
		<span class="cot-text-emphasized" title="This field is required">
		<input type="text" value="<%=sLastName%>" name="cot_txtLast_Name" id="txtLast_Name" style="width:300;" maxlength="100">
		</span>
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,3) Then %>


	<%
	If IsRequired(sMASK,3) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtBusiness_Name-text/req"" value=""Business Name"">"
	End If
	%>

	<tr><td align="right"><%=IsRequired(sMASK,3)%>
		Business Name:
	</td><td>
		<input type="text" value="<%=sBusinessName%>" name="cot_txtBusiness_Name" id="txtBusiness_Name" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,4) Then %>

	<%
	If IsRequired(sMASK,4) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtEmail-text/req"" value=""Email Address"">"
	End If
	%>

	<tr><td align="right">
		<%=IsRequired(sMASK,4)%>Email:
	</td><td>
		<input type="text" value="<%=sEmail%>" name="cot_txtEmail" id="txtEmail" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,5) Then %>

	<%
	If IsRequired(sMASK,5) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtDaytime_Phone-text/req"" value=""Daytime Phone"">"
	End If
	%>

	<tr><td align="right">
		<%=IsRequired(sMASK,5)%>Daytime Phone:
	</td><td>
		<!--<input type="text" value="<%=sHomePhone%>" name="cot_txtDaytime_Phone" id="txtDaytime_Phone" style="width:300;" maxlength="100">-->
		<input type="hidden" value="<%=sHomePhone%>" name="cot_txtDaytime_Phone">
		(<input type="text" value="<%=Left(sHomePhone,3)%>" name="skip_user_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
		<input type="text" value="<%=Mid(sHomePhone,4,3)%>" name="skip_user_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
		<input type="text" value="<%=Mid(sHomePhone,7,4)%>" name="skip_user_line" onKeyUp="return autoTab(this, 4, event);" size="4" maxlength="4">
	</td></tr>
	<%End IF%>

	<% If IsDisplay(sMASK,6) Then %>
	<%
		If IsRequired(sMASK,6) <> "" Then
			response.write "<input type=hidden name=""ef:cot_txtFax/req"" value=""Fax"">"
		End If
	%>
		<tr><td align="right">
			<%=IsRequired(sMASK,6)%>Fax:
		</td><td>
				(<input type="text" value="<%=Left(sFax,3)%>" name="skip_fax_areacode" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">)&nbsp;
				<input type="text" value="<%=Mid(sFax,4,3)%>" name="skip_fax_exchange" onKeyUp="return autoTab(this, 3, event);" size="3" maxlength="3">&ndash;
				<input type="text" value="<%=Right(sFax,4)%>" name="skip_fax_line" onKeyUp="return autoTab(this, 4, event);"size="4" maxlength="4">
				<input type="hidden" value="<%=sFax%>" name="cot_txtFax">
		</td></tr>
	
	<%End IF%>

	<% If IsDisplay(sMASK,7) Then %>
		<tr><td align="right">
			<%=IsRequired(sMASK,7)%>Street:
		</td><td>
			<input type="text" value="<%=sAddress%>" name="cot_txtStreet" id="txtStreet" style="width:300;" maxlength="100">
		</td></tr>
	<%
		If IsRequired(sMASK,7) <> "" Then
			response.write "<input type=hidden name=""ef:cot_txtStreet/req"" value=""Street"">"
		End If
	%>
	<%End IF%>

	<% If IsDisplay(sMASK,8) Then %>

	<%
	If IsRequired(sMASK,8) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtCity/req"" value=""City"">"
	End If
	%>

	<tr><td align="right">
		<%=IsRequired(sMASK,8)%>City:
	</td><td>
		<input type="text" value="<%=sCity%>" name="cot_txtCity" id="txtCity" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,9) Then %>

	<%
	If IsRequired(sMASK,9) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtState_vSlash_Province/req"" value=""State or Province"">"
	End If
	%>

	<tr><td align="right">
		<%=IsRequired(sMASK,9)%>State / Province:
	</td><td>
		<input type="text" value="<%=sState%>" name="cot_txtState_vSlash_Province" id="txtState_vSlash_Province" size="5" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,10) Then %>

	<%
	If IsRequired(sMASK,10) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtZIP_vSlash_Postal_Code/req"" value=""Zipcode"">"
	End If
	%>

		<tr><td align="right">
		<%=IsRequired(sMASK,10)%>ZIP / Postal Code:
	</td><td>
		<input type="text" value="<%=sZip%>" name="cot_txtZIP_vSlash_Postal_Code" id="txtZIP_vSlash_Postal_Code" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<!--COMMENTED OUT JS 11/8/2004
	<tr><td align="right">
		Country:
	</td><td>
							
	<select name="cot_txtCountry" id="txtCountry">
		<option value="">Unspecified</option>
		<option value="United States" >United States</option>
		<option >Albania</option>
		<option >Algeria</option>
		<option >American Samoa</option>
		<option >Andorra</option>
		<option >Angola</option>
		<option >Anguilla</option>
		<option >Antarctica</option>
		<option >Antigua And Barbuda</option>
		<option >Argentina</option>
		<option >Armenia</option>
		<option >Aruba</option>
		<option >Australia</option>
		<option >Austria</option>
		<option >Azerbaijan</option>
		<option >Bahamas</option>
		<option >Bahrain</option>
		<option >Bangladesh</option>
		<option >Barbados</option>
		<option >Belarus</option>
		<option >Belgium</option>
		<option >Belize</option>
		<option >Benin</option>
		<option >Bermuda</option>
		<option >Bhutan</option>
		<option >Bolivia</option>
		<option >Bosnia and Herzegovina</option>
		<option >Botswana</option>
		<option >Bouvet Island</option>
		<option >Brazil</option>
		<option >British Indian Ocean Territory</option>
		<option >Brunei Darussalam</option>
		<option >Bulgaria</option>
		<option >Burkina Faso</option>
		<option >Burma</option>
		<option >Burundi</option>
		<option >Cambodia</option>
		<option >Cameroon</option>
		<option >Canada</option>
		<option >Cape Verde</option>
		<option >Cayman Islands</option>
		<option >Central African Republic</option>
		<option >Chad</option>
		<option >Chile</option>
		<option >China</option>
		<option >Christmas Island</option>
		<option >Cocos (Keeling) Islands</option>
		<option >Colombia</option>
		<option >Comoros</option>
		<option >Congo</option>
		<option >Congo, the Democratic Republic of the</option>
		<option >Cook Islands</option>
		<option >Costa Rica</option>
		<option >Cote d'Ivoire</option>
		<option >Croatia</option>
		<option >Cyprus</option>
		<option >Czech Republic</option>
		<option >Denmark</option>
		<option >Djibouti</option>
		<option >Dominica</option>
		<option >Dominican Republic</option>
		<option >East Timor</option>
		<option >Ecuador</option>
		<option >Egypt</option>
		<option >El Salvador</option>
		<option >England</option>
		<option >Equatorial Guinea</option>
		<option >Eritrea</option>
		<option >Espana</option>
		<option >Estonia</option>
		<option >Ethiopia</option>
		<option >Falkland Islands</option>
		<option >Faroe Islands</option>
		<option >Fiji</option>
		<option >Finland</option>
		<option >France</option>
		<option >French Guiana</option>
		<option >French Polynesia</option>
		<option >French Southern Territories</option>
		<option >Gabon</option>
		<option >Gambia</option>
		<option >Georgia</option>
		<option >Germany</option>
		<option >Ghana</option>
		<option >Gibraltar</option>
		<option >Great Britain</option>
		<option >Greece</option>
		<option >Greenland</option>
		<option >Grenada</option>
		<option >Guadeloupe</option>
		<option >Guam</option>
		<option >Guatemala</option>
		<option >Guinea</option>
		<option >Guinea-Bissau</option>
		<option >Guyana</option>
		<option >Haiti</option>
		<option >Heard and Mc Donald Islands</option>
		<option >Honduras</option>
		<option >Hong Kong</option>
		<option >Hungary</option>
		<option >Iceland</option>
		<option >India</option>
		<option >Indonesia</option>
		<option >Ireland</option>
		<option >Israel</option>
		<option >Italy</option>
		<option >Jamaica</option>
		<option >Japan</option>
		<option >Jordan</option>
		<option >Kazakhstan</option>
		<option >Kenya</option>
		<option >Kiribati</option>
		<option >Korea (North)</option>
		<option >Korea, Republic of</option>
		<option >Korea (South)</option>
		<option >Kuwait</option>
		<option >Kyrgyzstan</option>
		<option >Lao People's Democratic Republic</option>
		<option >Latvia</option>
		<option >Lebanon</option>
		<option >Lesotho</option>
		<option >Liberia</option>
		<option >Liechtenstein</option>
		<option >Lithuania</option>
		<option >Luxembourg</option>
		<option >Macau</option>
		<option >Macedonia</option>
		<option >Madagascar</option>
		<option >Malawi</option>
		<option >Malaysia</option>
		<option >Maldives</option>
		<option >Mali</option>
		<option >Malta</option>
		<option >Marshall Islands</option>
		<option >Martinique</option>
		<option >Mauritania</option>
		<option >Mauritius</option>
		<option >Mayotte</option>
		<option >Mexico</option>
		<option >Micronesia, Federated States of</option>
		<option >Moldova, Republic of</option>
		<option >Monaco</option>
		<option >Mongolia</option>
		<option >Montserrat</option>
		<option >Morocco</option>
		<option >Mozambique</option>
		<option >Myanmar</option>
		<option >Namibia</option>
		<option >Nauru</option>
		<option >Nepal</option>
		<option >Netherlands</option>
		<option >Netherlands Antilles</option>
		<option >New Caledonia</option>
		<option >New Zealand</option>
		<option >Nicaragua</option>
		<option >Niger</option>
		<option >Nigeria</option>
		<option >Niue</option>
		<option >Norfolk Island</option>
		<option >Northern Ireland</option>
		<option >Northern Mariana Islands</option>
		<option >Norway</option>
		<option >Oman</option>
		<option >Pakistan</option>
		<option >Palau</option>
		<option >Panama</option>
		<option >Papua New Guinea</option>
		<option >Paraguay</option>
		<option >Peru</option>
		<option >Philippines</option>
		<option >Pitcairn</option>
		<option >Poland</option>
		<option >Portugal</option>
		<option >Puerto Rico</option>
		<option >Qatar</option>
		<option >Reunion</option>
		<option >Romania</option>
		<option >Russia</option>
		<option >Russian Federation</option>
		<option >Rwanda</option>
		<option >Saint Kitts and Nevis</option>
		<option >Saint Lucia</option>
		<option >Saint Vincent and the Grenadines</option>
		<option >Samoa (Independent)</option>
		<option >San Marino</option>
		<option >Sao Tome and Principe</option>
		<option >Saudi Arabia</option>
		<option >Scotland</option>
		<option >Senegal</option>
		<option >Seychelles</option>
		<option >Sierra Leone</option>
		<option >Singapore</option>
		<option >Slovakia</option>
		<option >Slovenia</option>
		<option >Solomon Islands</option>
		<option >Somalia</option>
		<option >South Africa</option>
		<option >South Georgia and the South Sandwich Islands</option>
		<option >South Korea</option>
		<option >Spain</option>
		<option >Sri Lanka</option>
		<option >St. Helena</option>
		<option >St. Pierre and Miquelon</option>
		<option >Suriname</option>
		<option >Svalbard and Jan Mayen Islands</option>
		<option >Swaziland</option>
		<option >Sweden</option>
		<option >Switzerland</option>
		<option >Taiwan</option>
		<option >Tajikistan</option>
		<option >Tanzania</option>
		<option >Thailand</option>
		<option >Togo</option>
		<option >Tokelau</option>
		<option >Tonga</option>
		<option >Trinidad</option>
		<option >Trinidad and Tobago</option>
		<option >Tunisia</option>
		<option >Turkey</option>
		<option >Turkmenistan</option>
		<option >Turks and Caicos Islands</option>
		<option >Tuvalu</option>
		<option >Uganda</option>
		<option >Ukraine</option>
		<option >United Arab Emirates</option>
		<option >United Kingdom</option>
		<option SELECTED>United States</option>
		<option >United States Minor Outlying Islands</option>
		<option >Uruguay</option>
		<option >USA</option>
		<option >Uzbekistan</option>
		<option >Vanuatu</option>
		<option >Vatican City State (Holy See)</option>
		<option >Venezuela</option>
		<option >Viet Nam</option>
		<option >Virgin Islands (British)</option>
		<option >Virgin Islands (U.S.)</option>
		<option >Wales</option>
		<option >Wallis and Futuna Islands</option>
		<option >Western Sahara</option>
		<option >Yemen</option>
		<option >Zambia</option>
		<option >Zimbabwe</option>
	</select>
	
	</td></tr>-->
<%End Function%>


<%
'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	DBsafe = sNewString
End Function


'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYQUESTIONS(IFORMID)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayQuestions(iFormID,sMask)

	sSQL = "SELECT * FROM egov_action_form_questions WHERE formid=" & iFormID & " ORDER BY sequence"

	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oQuestions.EOF Then
	
		response.write "<table>"
	
	Do While NOT oQuestions.EOF 
		
		' ENUMERATE QUESTIONS
		iQuestionCount = iQuestionCount + 1
		
		' DETERMINE IF REQUIRED
		sIsrequired = oQuestions("isrequired")
		If sIsrequired = True Then
			sIsrequired = " <font color=red>*</font> "
		Else
			sIsrequired = ""
		End If

		Select Case oQuestions("fieldtype")

			Case "2"
			' BUILD RADIO QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input value=""" & arrAnswers(alist) & """ name=fmquestion" & iQuestionCount & " class=formradio type=radio>" & arrAnswers(alist) & "</td></tr>"
			Next

			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "4"
			' BUILD SELECT QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			response.write "<tr><td><select class=formselect name=fmquestion" & iQuestionCount & " >"
			For alist = 0 to ubound(arrAnswers)
				response.write "<option value=""" & arrAnswers(alist) & """>" & arrAnswers(alist) & "</option>" 
			Next
			response.write "</select></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "6"
			' BUILD CHECKBOX QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input value=""" & arrAnswers(alist) & """ name=fmquestion" & iQuestionCount & " class=formcheckbox type=checkbox>" & arrAnswers(alist) & "</td></tr>"
			Next

			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "8"
			' BUILD TEXT QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><input name=fmquestion" & iQuestionCount & " value="""" type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case "10"
			' BUILD TEXTAREA QUESTION
			If sIsrequired <> "" Then
				response.write "<input type=hidden name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" &  Left(oQuestions("prompt"),75) & "..."">"
			End If
			response.write "<input type=hidden name=""fmname" & iQuestionCount & """ value=""" &  oQuestions("prompt") & """>"
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><textarea name=fmquestion" & iQuestionCount & " class=""formtextarea"" ></textarea></td></tr>"
			response.write "<tr><TD>&nbsp;</td></tr>"

			Case Else

		End Select 

		oQuestions.MoveNext
	Loop

		response.write "</table>"

		' SEND EMAIL
		If (IsDisplay(sMASK,4)) Then
			response.write "<div style=""width:450px;text-align:right;""><input type=checkbox name=chkSendEmail checked value=""YES""> Check here to have email confirmation of this request sent.</div>"
		End If
		
		' SUBMIT BUTTON
		response.write "<div style=""width:450px;text-align:right;""><input class=actionbtn type=""button"" onclick=""if (ValidateInput()) {document.frmRequestAction.submit();}""  name=""btnSubmit"" value=""SEND REQUEST""></div>"


	End If

	Set oQuestions = Nothing 

End Sub

'------------------------------------------------------------------------------------------------------------
' FUNCTION ISREQUIRED(SMASK,IFIELD)
'------------------------------------------------------------------------------------------------------------
Function IsRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
		sReturnValue = " <font color=red>*</font> "
	Else
		sReturnValue = ""
	End If

	IsRequired = sReturnValue
End Function

'------------------------------------------------------------------------------------------------------------
' FUNCTION ISDISPLAY(SMASK,IFIELD)
'------------------------------------------------------------------------------------------------------------
Function IsDisplay(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "1" or sValue = "2" Then
		sReturnValue = True
	Else
		sReturnValue = False
	End If

	IsDisplay = sReturnValue
End Function


'--------------------------------------------------------------------------------------------------
' PUBLIC SUB DISPLAYITEM(CLASSID)
'--------------------------------------------------------------------------------------------------
 Public Sub DisplayItem(classid)


	' GET SELECTED FACILITY INFORMATION
	sSQL = "select *, IsNull(egov_class.startdate,0) as startdate, IsNull(egov_class.enddate,0) as enddate, IsNull(egov_class.imgurl,'EMPTY') as imgurl, (firstname + ' ' + lastname) as Instructor,IsNull(registrationstartdate,GetDate()) as registrationstartdate,IsNull(registrationenddate,GetDate()) as registrationenddate "
	sSql = sSql & " from egov_class left join egov_class_time ON egov_class.classid = egov_class_time.classid  left join egov_class_instructor ON egov_class_time.instructorid = egov_class_instructor.instructorid LEFT JOIN egov_class_pointofcontact ON egov_class.pocid = egov_class_pointofcontact.pocid LEFT JOIN egov_registration_option ON egov_class.optionid = egov_registration_option.optionid "
	sSql = sSql & " where egov_class.classid = '" &  classid & "' order by noenddate desc, startdate"


	Set oItem = Server.CreateObject("ADODB.Recordset")
	oItem.Open sSQL, Application("DSN"), 3, 1

    ' DISPLAY ITEM INFORMATION
    If NOT oItem.EOF Then
			
				' WRITE CLASS NAME
				Response.Write("<div class=""facilityname"">" &  oItem("classname") & "</div>" & vbCrLf)
				response.write "<input type=hidden	name=fmquestionclassname value=""" & oItem("classname") & """>"
				response.write "<input type=hidden name=fmnameclassname value=""ClassName"">"
				' CREATE CLASS IMAGE URL
				'sImgURL = ""
				'If  oItem("imgurl") = "EMPTY"  Then
					'sImgURL = "images/class_category_default.gif"
					'sImgAlt = "Default Class and Event Photo"
				'Else
					'sImgURL = oItem("imgurl")
					'sImgAlt = oItem("imgalttag")
				'End If
				

			'	Response.Write("<div class=classdesc>")
				
				' WRITE IMAGE
				'If  oItem("imgurl") <> "EMPTY"  Then
					'response.write("<img class=""categoryimage"" ALIGN=""left""  alt=""" & sImgAlt & """ src=""" & sImgURL & """>")
				'End If

				' WRITE DESCRIPTION
				'response.write oItem("classdescription") 

				'response.write("</div>")
		
				
				' DISPLAY ITEM DETAILS
				'response.write "<br><div><fieldset class=classdetails><legend><b>Date(s):</b></legend>
				response.write "<table>"
					
					' DISPLAY DETAILS VALUE PAIR

					' REGISTRATION DATES 
					If oItem("optionid") = 1 Then
						If NOT ISNULL(oItem("registrationstartdate")) Then
							'response.write "<tr><td class=classdetaillabel>Registration Start Date: </td><td class=classdetailvalue>" & 'FormatDateTime(oItem("registrationstartdate"),1) & "</td></tr>"
						End If 
						If NOT ISNULL(oItem("registrationenddate")) Then
							'response.write "<tr><td class=classdetaillabel>Registration End Date: </td><td class=classdetailvalue>" & 'FormatDateTime(oItem("registrationenddate"),1) & "</td></tr>"
						End If
					End If
				
					' START DATES
					If NOT ISNULL(oItem("startdate")) Then
						response.write "<tr><td class=classdetaillabel>Start Date: </td><td class=classdetailvalue>" & FormatDateTime(oItem("startdate"),1) & "</td></tr>"
					End If
					If NOT ISNULL(oItem("enddate")) Then
						response.write "<tr><td class=classdetaillabel>End Date: </td><td class=classdetailvalue>" & FormatDateTime(oItem("enddate"),1) & "</td></tr>"
					End If
					If NOT ISNULL(oItem("alternatedate")) Then
						'response.write "<tr><td class=classdetaillabel>Make Up Date: </td><td class=classdetailvalue>" & FormatDateTime(oItem("alternatedate"),1) & "</td></tr>"
					End If
					If NOT ISNULL(oItem("minage")) Then
						'response.write "<tr><td class=classdetaillabel>Minimum Age: </td><td class=classdetailvalue>" & oItem("minage") & "</td></tr>"
					End If
					If NOT ISNULL( oItem("maxage")) Then
						'response.write "<tr><td class=classdetaillabel>Maximum Age: </td><td class=classdetailvalue>" & oItem("maxage") & "</td></tr>"
					End If
					If NOT ISNULL(oItem("sexrestriction")) Then
						'response.write "<tr><td class=classdetaillabel>Sex: </td><td class=classdetailvalue>" & oItem("sexrestriction") & "</td></tr>"
					End If

					' DISPLAY INSTRUCTOR(S)
					subGetInstructors(classid)

					' DISPLAY POINT OF CONTACT INFORMATION
					'response.write "<tr><td class=classdetaillabel>Point of Contact Information </td><td class=classdetailvalue>&nbsp;</td></tr>"
					If NOT ISNULL(oItem("name")) Then
						'response.write "<tr><td class=classdetaillabel>Name: </td><td class=classdetailvalue>" & oItem("name") & "</td></tr>"
					End If
					If NOT ISNULL(oItem("phone")) Then
						'response.write "<tr><td class=classdetaillabel>Phone: </td><td class=classdetailvalue>" & FormatPhone(oItem("phone")) & "</td></tr>"
					End If
					If NOT ISNULL(oItem("email")) Then
						'response.write "<tr><td class=classdetaillabel>Email: </td><td class=classdetailvalue>" & oItem("email") & "</td></tr>"
					End If


					response.write "</table>"
					'response.write "</fieldset></div>"

					' DISPLAY TIMES
					DisplayClassTimes2 oItem("classid"), request("timeid")


		End If

        ' CLOSE OBJECTS
		oItem.close
        Set  oItem = Nothing 

 End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNGETTIMEDAYSOFWEEK(ICLASSID)
'--------------------------------------------------------------------------------------------------
Sub subGetInstructors(iclassid)
	
	sReturnValue = ""

	' GET INSTRUCTOR LIST FOR THIS CLASS
	sSQL = "SELECT * FROM egov_class_to_instructor inner join egov_class_instructor ON egov_class_to_instructor.instructorid=egov_class_instructor.instructorid where classid = '" & iClassId & "'"
	
	Set oInstructors = Server.CreateObject("ADODB.Recordset")
	oInstructors.Open sSQL, Application("DSN"), 3, 1
	
	' IF NOT EMPTY
	If NOT oInstructors.EOF Then

		' LOOP THRU INSTRUCTORS FOR THIS CLASS
		Do While NOT oInstructors.EOF 
			iCount = iCount + 1
			If iCount = 1 Then
				response.write "<tr><td class=classdetaillabel>Instructor(s): </td>"
			Else
				response.write "<tr><td class=classdetaillabel>&nbsp;</td>"
			End If
			response.write "<td><a href=""instructor_info.asp?iID=" & oInstructors("instructorid")& """ target=_NEW>" & oInstructors("firstname") & " " & oInstructors("lastname") &  "</a></td></tr>"
			oInstructors.MoveNext
		Loop

	Else
		' NO INSTRUCTORS FOUND
	End If

	' CLEAR OBJECTS
	Set oInstructors = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' SUB DISPLAYCLASSTIMES2(ICLASSID)
'--------------------------------------------------------------------------------------------------
Sub DisplayClassTimes2(iClassId,iTimeID)

	sSQL = "SELECT  egov_class_time.starttime, egov_class_time.endtime, egov_class_time.min, egov_class_time.max FROM egov_class_time where (egov_class_time.classid = '" & iClassId & "' and egov_class_time.timeid='" & iTimeID & "')"
	
	Set oClassTimes = Server.CreateObject("ADODB.Recordset")
	oClassTimes.Open sSQL, Application("DSN"), 3, 1
	
	' INSTRUCTOR INFORMATION
	If not oClassTimes.EOF Then

		' DISPLAY CLASS INFORMATION
		'response.write "<BR>"
		'response.write "<div><fieldset class=classdetails><legend><b>Time(s):</b></legend>"
		response.write "<table>"
		Do While NOT oClassTimes.EOF 
			response.write "<tr><td class=classdetaillabel>Time: </td><td> " & oClassTimes("starttime") & " - " & oClassTimes("endtime")  & " " & fnGetTimeDaysofWeek(iclassid) & "</td></tr>"
			response.write "<input type=hidden name=fmquestionclasstime value=""" &  oClassTimes("starttime") & " - " & oClassTimes("endtime")  & " " & fnGetTimeDaysofWeek(iclassid)  & """>"
			response.write "<input type=hidden name=fmnameclasstime value=""ClassTime"">"
			oClassTimes.MoveNext
		Loop

		response.write "</table>"
		'response.write "</fieldset></div>"

	Else
		' NO CLASSES FOUND
	End If

	Set oClassTimes = Nothing

End Sub


'--------------------------------------------------------------------------------------------------
' FUNCTION FNGETTIMEDAYSOFWEEK(ICLASSID)
'--------------------------------------------------------------------------------------------------
Function fnGetTimeDaysofWeek(iclassid)
	
	sReturnValue = ""

	' GET THE DAY OF THE WEEK VALUES FOR THE SPECIFIED
	sSQL = "SELECT dayofweek FROM egov_class_dayofweek where classid = '" & iClassId & "'"
	
	Set oClassDays = Server.CreateObject("ADODB.Recordset")
	oClassDays.Open sSQL, Application("DSN"), 3, 1
	
	' IF NOT EMPTY
	If not oClassDays.EOF Then

		' LOOP THRU AVAILABLE DAYS OF THE WEEK
		Do While NOT oClassDays.EOF 
			sReturnValue = sReturnValue &  weekdayname(oClassDays("dayofweek"),true) & " "
			oClassDays.MoveNext
		Loop

	Else
		' NO DAYS FOUND
	End If

	' CLEAR OBJECTS
	Set oClassDays = Nothing

	' RETURN DAYS OF THE WEEK
	fnGetTimeDaysofWeek = Trim(sReturnValue)

End Function


%>
