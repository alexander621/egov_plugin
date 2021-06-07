<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% Dim sError %>

<html>
<head>
<title>E-Gov Services <%=sOrgName%></title>
<link rel="stylesheet" href="css/styles.css" type="text/css">
<link href="global.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/style_<%=iorgid%>.css" type="text/css">
<script language="Javascript" src="scripts/modules.js"></script>
<script language="Javascript" src="scripts/easyform.js"></script>
<script language=javascript>
function openWin2(url, name) {
  popupWin = window.open(url, name,"resizable,width=500,height=450");
}

</script>
</head>


<!--#Include file="include_top.asp"-->


<TR><TD VALIGN=TOP>



<!--BODY CONTENT-->
<p>
<font class=pagetitle>Welcome to the <%=sOrgName%> Action Line</font> <BR>

<% If sOrgRegistration AND trim(request("actionid")) <> "" Then %>
		<%  If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
				RegisteredUserDisplay()
			Else %>
				<a href="user_login.asp">Click here to Login</a> |
				<a href="register.asp">Click here to Register</a>
		<% End If %>
<% Else %>
	<font class=datetagline>Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </font>
<% End If%>


</p>

<div style="margin-left:25px;">


<%
' ---------------------------------------------------------------------------------------
' BEGIN DISPLAY PAGE CONTENT
' ---------------------------------------------------------------------------------------
If trim(request("actionid")) <> "" Then 

	Call subDisplayActionForm(request("actionid"),iorgid)

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
	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
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
					<form name="frmActionLookup" action="action_request_lookup.asp" METHOD="POST">
					<div style=""margin-left:20px; "" class=box_header2>Check the Status of an Action Line Request</div>
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
					<div style=""margin-left:20px; "" class=box_header2>Create New Action Line Request</div>
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
			  

			  <%If sOrgRegistration Then %>
			  <b>Personalized E-Gov Services</b>
			  <ul>
				<li><a href="user_login.asp">Click here to Login</a>
				<li><a href="register.asp">Click here to Register</a>
			  </ul>
			  <hr style="width: 90%; size: 1px; height: 1px;">
			  <%End If%>
			  
			  
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
	sSQL = "SELECT * FROM dbo.egov_form_list_200  WHERE ((orgid=" & iorgID & ")) AND (form_category_id <> 6) order by form_category_Sequence,action_form_name"

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

<form name="frmRequestAction" action="action_cgi.asp?list=true" METHOD="POST">
<input type="hidden" name="actionid" value="<%=iFormID%>">
<input type="hidden" name="actiontitle" value="<%=sTitle%>">

<div style="margin-top:20px; margin-left:20px;" >

<!--BEGIN: TITLE-->
<font class=formtitle><%=sTitle%></font>
<!--END: TITLE-->


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

' GET USER PERSONNEL INFORMATION IF USER IS LOGGED INTO WEBSITE
iUserID = 715

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
	sFaxPhone = oInfo("userfax")

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

End If 
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
		<input type="text" value="<%=sHomePhone%>" name="cot_txtDaytime_Phone" id="txtDaytime_Phone" style="width:300;" maxlength="100">
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
		<input type="text" value="<%=sFaxPhone%>" name="cot_txtFax" id="txtFax" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,7) Then %>

	<%
	If IsRequired(sMASK,7) <> "" Then
		response.write "<input type=hidden name=""ef:cot_txtStreet/req"" value=""Street"">"
	End If
	%>

	<tr><td align="right">
		<%=IsRequired(sMASK,7)%>Street:
	</td><td>
		<input type="text" value="<%=sAddress%>" name="cot_txtStreet" id="txtStreet" style="width:300;" maxlength="100">
	</td></tr>
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
'------------------------------------------------------------------------------------------------------------
' FUNCTION DBSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
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
		response.write "<div style=""width:450px;text-align:right;""><input class=actionbtn type=""button"" onclick=""if (validateForm('frmRequestAction')) {document.frmRequestAction.submit();}""  name=""btnSubmit"" value=""SEND REQUEST""></div>"


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



%>
