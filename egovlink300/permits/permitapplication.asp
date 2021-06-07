<%
Response.AddHeader "Access-Control-Allow-Origin", "*"

session("RedirectPage") = "permits/permitapplication.asp"

%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<html>
<head>

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>

	<% if request("js") <> "true" then %>
	<link rel="stylesheet" type="text/css" href="//www.egovlink.com/permitcity/css/styles.css" />
	<link rel="stylesheet" type="text/css" href="//www.egovlink.com/permitcity/global.css" />
	<link rel="stylesheet" type="text/css" href="//www.egovlink.com/permitcity/css/style_<%=iorgid%>.css" />
	<% end if %>
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" type="text/css" href="//www.egovlink.com/permitcity/permits/permitapp.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
</head>

<!--#Include file="../include_top.asp"-->

<% if request("js") <> "true" then %>
<h2>Permit Application Form</h2>
<% end if %>
<div id="successmsg">
	<p>Your permit application was successfully received.  Your reference number is <span id="applicationid"></span>.</p>
	<p><a href="javascript:location.reload();">Return to Permit Application</a></p>
	
</div>

<%
sSQL = "SELECT q.*, ft.* " _
	& " FROM egov_permitapplication pa " _
	& " INNER JOIN egov_permitapplication_questions q ON q.permitapplicationid = pa.permitapplicationid " _
	& " INNER JOIN egov_permitapplication_fieldtypes ft ON ft.fieldtypeid = q.fieldtypeid " _
	& " WHERE pa.orgid = '" & iorgid & "' and parentquestionid = 0" _
	& " ORDER BY q.DisplayOrder " 

Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1


blnFoundForm = false
if not oRs.EOF then
	blnFoundForm = true%>
	<form id="permitAppFrm" name="permitAppFrm" method="POST" style="display:none;">
		<input type="hidden" name="orgid" value="<%=iorgid%>" />
		<input type="hidden" id="userid" name="userid" value="<%=request.cookies("userid")%>" />
<% end if

iCount = 0
Do While not oRs.EOF
	ShowQuestion oRs("questionname"), oRs("fieldtypebehavior"), oRs("options"), oRs("instructions"), oRs("permitapplicationquestionid"), oRs("parentquestionid"), oRs("isrequired"), oRs("defaultopened")

	oRs.MoveNext
loop

if blnFoundForm then%>
	<button id="submitBtn" type="button" class="button" onClick="submitPA();">Submit Application</button>
	</form>
<%end if%>



<%
oRs.Close
Set oRs = Nothing
%>


 <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="//www.egovlink.com/permitcity/scripts/easyform_enhance.js"></script>
<script type="text/javascript" src="//www.egovlink.com/permitcity/permits/permitapp.js"></script>


<!--#Include file="../include_bottom.asp"-->  

<%
sub ShowQuestion(strQuestionName, strFieldTypeBehavior, strOptions, strInstructions, intQuestionID, strParentQuestionID, blnIsRequired, blnDefaultOpened)
	strRequiredFlag = ""
	if blnIsRequired then strRequiredFlag = "<span class=""required"">*</span>"

	iCount = iCount + 1
	if strFieldTypeBehavior <> "grouping" and strFieldTypeBehavior <> "message" and strFieldTypeBehavior <> "address" and strFieldTypeBehavior <> "contact" then%>
		<b><%=strQuestionName & strRequiredFlag%><%
		if strInstructions <> "" then
			response.write "&nbsp;(" & strInstructions & ")"
		end if

		%>:</b><br />
	<%end if

	%><input type="hidden" name="customfield<%=iCount%>_questionid" value="<%=intQuestionID%>" /><%

	strReq = ""
	if blnIsRequired then strReq = "/req"

	Select Case strFieldTypeBehavior
		
		Case "radio"
			aChoices = Split(strOptions,Chr(10))
	
			for x = 0 to UBound(aChoices)
				sChoice = Replace(aChoices(x), Chr(13), "")
				response.write vbcrlf & "<input type=""radio"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sChoice & """"
				response.write " /> " & sChoice & "<br />" & vbcrlf
			Next
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-radio" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
			
		Case "select"
			response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>" & vbcrlf
			aChoices = Split(strOptions,Chr(10))
	
			for x = 0 to UBound(aChoices)
				sChoice = Replace(aChoices(x), Chr(13), "")
				response.write vbcrlf & "<option value=""" & sChoice & """"
				response.write " /> " & sChoice & "</option>" & vbcrlf
			Next
			response.write vbcrlf & "</select>" & vbcrlf
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
		
		Case "checkbox"
			aChoices = Split(strOptions,Chr(10))
	
			For x = 0 To UBound(aChoices)
				sChoice = Replace(aChoices(x), Chr(13), "")
				response.write vbcrlf & "<input type=""checkbox"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sChoice & """"
				response.write " /> " & sChoice & "<br />" & vbcrlf
			Next
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-checkbox" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf

		Case "textbox","email","phone"
			response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value="""" size=""100"" />" & vbcrlf
			strAddSet = ""
			if strFieldTypeBehavior = "email" then strAddSet = "/email"
			if strFieldTypeBehavior = "phone" then strAddSet = "/phone"
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-text" & strAddSet & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf

		Case "textarea"
			response.write "<textarea class=""customfields"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ rows=""25"">" & "</textarea>" & vbcrlf

		Case "date"
			response.write "<input type=""text"" class=""datepicker"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""""  size=""10"" maxlength=""10"" />" & vbcrlf
			' put a date picker
			'response.write "&nbsp;<a href=""javascript:void doCalendar('customfield" & iCount & "');""><img src=""../images/calendar.gif"" border=""0"" /></a>"
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-text/date" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
		
		Case "integer","decimal","money"
			if strFieldTypeBehavior = "money" then strFieldTypeBehavior = "currency"
			response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value="""" size=""50"" />" & vbcrlf
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-text/" & strFieldTypeBehavior & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
		case "permittype"
			ShowPermitTypePicks 0, strReq, strQuestionName
		case "usetype"
			ShowUseTypes 0, strReq, strQuestionName, iCount
		case "useclass"
			ShowUserClasses 0, strReq, strQuestionName, iCount
		case "workclass"
			ShowWorkClass 0, strReq, strQuestionName, iCount
		case "workscope"
			ShowWorkScopes 0, strReq, strQuestionName, iCount
		case "constructiontype"
			ShowConstructionTypes 0, strReq, strQuestionName, iCount
		case "occtype"
			ShowOccupancyTypes 0, strReq, strQuestionName, iCount
		case "message"
			response.write replace(strQuestionName, vbcrlf,"<br />")
		Case "grouping"
			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " active"
			%>
			<button type="button" class="collapsible <%=strDefaultOpened%>"><%=strQuestionName & strRequiredFlag%><%
			if strInstructions <> "" then
				response.write "&nbsp;(" & strInstructions & ")"
			end if

			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " style=""display:block"""

			%> <div class="expand-icon"><i class="fa fa-bars"></i></div></button>
			<div class="content" <%=strDefaultOpened%>>
				<br />
				<%GetChildQuestions intQuestionID%>
			</div>
			<%
		Case "contact"
			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " active"
			%>
			<input type="hidden" id="customfield<%=iCount%>_fieldtype" name="customfield<%=iCount%>_fieldtype" value="contact" />
			<button type="button" class="collapsible <%=strDefaultOpened%>"><%=strQuestionName & strRequiredFlag%><%
			if strInstructions <> "" then
				response.write "&nbsp;(" & strInstructions & ")"
			end if

			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " style=""display:block"""

			%> <div class="expand-icon"><i class="fa fa-bars"></i></div></button>
			<div class="content" <%=strDefaultOpened%>>
				<br />
				<table class="contactfields">
					<tr><td class="label">First Name<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_firstname" name="customfield<%=iCount%>_firstname" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">Last Name<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_lastname" name="customfield<%=iCount%>_lastname" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">Company<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_company" name="customfield<%=iCount%>_company" size="50" maxlength="100" /></td></tr>
					<tr><td class="label">Address<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_address" name="customfield<%=iCount%>_address" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">City<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_city" name="customfield<%=iCount%>_city" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">State<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_state" name="customfield<%=iCount%>_state" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">Zip<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_zip" name="customfield<%=iCount%>_zip" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">Email<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_email" name="customfield<%=iCount%>_email" size="50" maxlength="100" /></td></tr>
					<tr><td class="label">Phone<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_phone" name="customfield<%=iCount%>_phone" size="50" maxlength="10" /></td></tr>
					<tr><td class="label">Cell<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_cell" name="customfield<%=iCount%>_cell" size="50" maxlength="10" /></td></tr>
					<tr><td class="label">Fax<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_fax" name="customfield<%=iCount%>_fax" size="50" maxlength="10" /></td></tr>
					<input type="hidden" name="ef:customfield<%=iCount%>_firstname-text<%=strReq%>" value="<%=strQuestionName%> First Name" />
					<input type="hidden" name="ef:customfield<%=iCount%>_lastname-text<%=strReq%>" value="<%=strQuestionName%> Last Name" />
					<input type="hidden" name="ef:customfield<%=iCount%>_company-text<%=strReq%>" value="<%=strQuestionName%> Company" />
					<input type="hidden" name="ef:customfield<%=iCount%>_address-text<%=strReq%>" value="<%=strQuestionName%> Address" />
					<input type="hidden" name="ef:customfield<%=iCount%>_city-text<%=strReq%>" value="<%=strQuestionName%> City" />
					<input type="hidden" name="ef:customfield<%=iCount%>_state-text<%=strReq%>" value="<%=strQuestionName%> State" />
					<input type="hidden" name="ef:customfield<%=iCount%>_zip-text/zip<%=strReq%>" value="<%=strQuestionName%> Zip" />
					<input type="hidden" name="ef:customfield<%=iCount%>_county-text<%=strReq%>" value="<%=strQuestionName%> County" />
					<input type="hidden" name="ef:customfield<%=iCount%>_email-text/email<%=strReq%>" value="<%=strQuestionName%> Email" />
					<input type="hidden" name="ef:customfield<%=iCount%>_phone-text/phone<%=strReq%>" value="<%=strQuestionName%> Phone" />
					<input type="hidden" name="ef:customfield<%=iCount%>_cell-text/phone<%=strReq%>" value="<%=strQuestionName%> Cell" />
					<input type="hidden" name="ef:customfield<%=iCount%>_fax-text/phone<%=strReq%>" value="<%=strQuestionName%> Fax" />
				</table>
				<br />
			</div>
			<%
		case "address"
			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " active"
			%>
			<input type="hidden" id="customfield<%=iCount%>_fieldtype" name="customfield<%=iCount%>_fieldtype" value="address" />
			<button type="button" class="collapsible <%=strDefaultOpened%>"><%=strQuestionName & strRequiredFlag%><%
			if strInstructions <> "" then
				response.write "&nbsp;(" & strInstructions & ")"
			end if

			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " style=""display:block"""

			%> <div class="expand-icon"><i class="fa fa-bars"></i></div></button>
			<div class="content" <%=strDefaultOpened%>>
				<br />
				<table class="contactfields">
					<tr><td class="label">Street Number<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_streetnumber" name="customfield<%=iCount%>_streetnumber" size="50" maxlength="15" /></td></tr>
					<tr><td class="label">Street Name<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_streetname" name="customfield<%=iCount%>_streetname" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">Street Suffix<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_streetsuffix" name="customfield<%=iCount%>_streetsuffix" size="50" maxlength="15" /></td></tr>
					<tr><td class="label">City<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_city" name="customfield<%=iCount%>_city" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">State<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_state" name="customfield<%=iCount%>_state" size="50" maxlength="50" /></td></tr>
					<tr><td class="label">Zip<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_zip" name="customfield<%=iCount%>_zip" size="50" maxlength="10" /></td></tr>
					<tr><td class="label">County<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_county" name="customfield<%=iCount%>_county" size="50" maxlength="50" /></td></tr>
					<input type="hidden" name="ef:customfield<%=iCount%>_streetnumber-text/integer<%=strReq%>" value="<%=strQuestionName%> Street Number" />
					<input type="hidden" name="ef:customfield<%=iCount%>_streetname-text<%=strReq%>" value="<%=strQuestionName%> Street Name" />
					<input type="hidden" name="ef:customfield<%=iCount%>_streetsuffix-text<%=strReq%>" value="<%=strQuestionName%> Street Suffix" />
					<input type="hidden" name="ef:customfield<%=iCount%>_city-text<%=strReq%>" value="<%=strQuestionName%> City" />
					<input type="hidden" name="ef:customfield<%=iCount%>_state-text<%=strReq%>" value="<%=strQuestionName%> State" />
					<input type="hidden" name="ef:customfield<%=iCount%>_zip-text/zip<%=strReq%>" value="<%=strQuestionName%> Zip" />
					<input type="hidden" name="ef:customfield<%=iCount%>_county-text<%=strReq%>" value="<%=strQuestionName%> County" />
				</table>
				<br />
			</div>
			<%
	
	End Select 

	'response.flush
	if strFieldTypeBehavior <> "grouping" then GetChildQuestions intQuestionID

	response.write "<br />" & vbcrlf
	response.write "<br />" & vbcrlf
	'response.flush
end sub

Sub GetChildQuestions(intQuestionID)
	sSQL = "SELECT q.*, ft.* " _
		& " FROM egov_permitapplication pa " _
		& " INNER JOIN egov_permitapplication_questions q ON q.permitapplicationid = pa.permitapplicationid " _
		& " INNER JOIN egov_permitapplication_fieldtypes ft ON ft.fieldtypeid = q.fieldtypeid " _
		& " WHERE pa.orgid = '" & iorgid & "' and parentquestionid = '" & intQuestionID & "'"
	'response.write sSQL & "<br />" & vbcrlf
	'response.end

	Set oChild = Server.CreateObject("ADODB.RecordSet")
	oChild.Open sSQL, Application("DSN"), 3, 1
	Do While Not oChild.EOF
		ShowQuestion oChild("questionname"), oChild("fieldtypebehavior"), oChild("options"), oChild("instructions"), oChild("permitapplicationquestionid"), oChild("parentquestionid"), oChild("isRequired"), oChild("defaultopened")
		oChild.MoveNext
	loop
	oChild.Close
	Set oChild = Nothing

End Sub

'------------------------------------------------------------------------------
' void ShowPermitTypePicks( iPermitTypeId )
'------------------------------------------------------------------------------
Sub ShowPermitTypePicks( ByVal iPermitTypeId, strReq, strQuestionName )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, ISNULL(permittype,'') AS permittype, ISNULL(permittypedesc,'') AS permittypedesc "
	sSql = sSql & " FROM egov_permittypes "
	'sSql = sSql & " WHERE isbuildingpermittype = 1 AND orgid = "& session("orgid")
	sSql = sSql & " WHERE orgid = "& session("orgid")
	sSql = sSql & " ORDER BY permittype, permittypedesc, permittypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select id=""permittypeid"" name=""permittypeid"">"
		If CLng(iPermitTypeId) = CLng(0) Then
			response.write vbcrlf & "<option value="""">Please select a permit type...</option>"
		End If 
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value="""  & oRs("permittypeid") & """"
			If CLng(iPermitTypeId) = CLng(oRs("permittypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permittype") 
			If oRs("permittype") <> "" And oRs("permittypedesc") <> "" Then 
				response.write " &ndash; "
			End If 
			response.write oRs("permittypedesc")
			response.write "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
		response.write "<input type=""hidden"" name=""ef:permittypeid-select" & strReq & """ value=""Permit Type"" />" & vbcrlf
	Else
		response.write vbcrlf & "There are No Permit Types to select."
		response.write vbcrlf & "<input type=""hidden"" id=""permittypeid"" name=""permittypeid"" value=""0"" />"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

'--------------------------------------------------------------------------------------------------
' void ShowUseTypes iUseTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowUseTypes( intUseTypeId, strReq, strQuestionName, iCount )
	Dim sSql, oRs

	sSql = "SELECT usetypeid, usetype FROM egov_permitusetypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY usetype" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
		response.write vbcrlf & "<option value="""">Select a Use Type...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("usetypeid") & """"
			If CLng(intUseTypeId) = CLng(oRs("usetypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("usetype") & "</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
		response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuesetionName & """ />" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

'--------------------------------------------------------------------------------------------------
' void ShowUseClasses iUseClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowUseClasses( intUseClassId, strReq, strQuestionName, iCount )
	Dim sSql, oRs

	sSql = "SELECT useclassid, useclass FROM egov_permituseclasses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY useclass" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
	response.write vbcrlf & "<option value="""">Select a Use Class...</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("useclassid") & """"
		If CLng(intUseClassId) = CLng(oRs("useclassid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("useclass") & "</option>"
		oRs.MoveNext 
	Loop 
	response.write vbcrlf & "</select>"
	response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuesetionName & """ />" & vbcrlf

	oRs.Close
	Set oRs = Nothing 

End Sub 

'--------------------------------------------------------------------------------------------------
' void ShowWorkClass iWorkClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowWorkClass( intWorkClassId, strReq, strQuestionName, iCount )
	Dim sSql, oRs

	sSql = "SELECT workclassid, workclass FROM egov_permitworkclasses "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY workclass" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
		response.write vbcrlf & "<option value="""">Select...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("workclassid") & """"
			If CLng(intWorkClassId) = CLng(oRs("workclassid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("workclass") & "</option>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</select>"
		response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuesetionName & """ />" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub
'--------------------------------------------------------------------------------------------------
' void ShowWorkScopes iWorkScopeId 
'--------------------------------------------------------------------------------------------------
Sub ShowWorkScopes( intWorkScopeId, strReq, strQuestionName, iCount)
	Dim sSql, oRs

	sSql = "SELECT workscopeid, workscope FROM egov_permitworkscope "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY workscope" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
	response.write vbcrlf & "<option value="""">Select a Work Scope...</option>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("workscopeid") & """"
		If CLng(intWorkScopeId) = CLng(oRs("workscopeid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("workscope") & "</option>"
		oRs.MoveNext 
	Loop 
	response.write vbcrlf & "</select>"
	response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuesetionName & """ />" & vbcrlf

	oRs.Close
	Set oRs = Nothing 

End Sub 
'--------------------------------------------------------------------------------------------------
' void ShowConstructionTypes iConstructionTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowConstructionTypes( intConstructionTypeId, strReq, strQuestionName, iCount )
	Dim sSql, oRs

	sSql = "SELECT constructiontypeid, constructiontype FROM egov_constructiontypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY displayorder" 
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
	response.write vbcrlf & "<option value="""">Select...</option>"

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("constructiontypeid") & """"
			If CLng(intConstructionTypeId) = CLng(oRs("constructiontypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("constructiontype") & "</option>"
			oRs.MoveNext 
		Loop 
	End If 

	response.write vbcrlf & "</select>"
	response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuesetionName & """ />" & vbcrlf

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOccupancyTypes iOccupancyTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowOccupancyTypes( intOccupancyTypeId, strReq, strQuestionName, iCount )
	Dim sSql, oRs

	sSql = "SELECT occupancytypeid, ISNULL(usegroupcode,'') AS usegroupcode, ISNULL(occupancytype, '') AS occupancytype " 
	sSql = sSql & " FROM egov_occupancytypes WHERE orgid = " & session("orgid") & " ORDER BY usegroupcode, occupancytype" 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>"
	response.write vbcrlf & "<option value="""">Select...</option>"

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("occupancytypeid") & """"
			If CLng(intOccupancyTypeId) = CLng(oRs("occupancytypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" 
			If oRs("usegroupcode") <> "" Then 
				response.write oRs("usegroupcode") & " "
			End If 
			response.write oRs("occupancytype") & "</option>"
			oRs.MoveNext 
		Loop 
	End If 

	response.write vbcrlf & "</select>"
	response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuesetionName & """ />" & vbcrlf

	oRs.Close
	Set oRs = Nothing 

End Sub
%>
