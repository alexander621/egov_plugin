<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitapplication.asp
' AUTHOR: Terry Foster
' CREATED: 08/18/2020
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of permit applications
'
' MODIFICATION HISTORY
' 1.0   08/18/2020	Terry Foster - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem, iYearPick, sYearPick, sStreetNumber, sStreetName, sListedOwner, sPermitNo
Dim sContactname, iPermitStatusId, sFrom, sDistinct, iPermitTypeId, iStatusItemCount, sParcelIdNumber
Dim iPageSize, iCurrentPage, bInitialLoad, sFromActivityDate, sToActivityDate, iInvoiceNo, sLegalDescription
Dim iPermitCategoryId, sPermitLocation, sArchiveSearch, sPermitType, sPermitTypeDesc, bFoundArchiveStatus
Dim sTempArchiveSearch, iInvoiceNumber

ReDim aStatuses(0)

sLevel = "../" ' Override of value from common.asp
sSearch = ""
sFrom = ""
sArchiveSearch = ""
sDistinct = ""
bInitialLoad = False 

PageDisplayCheck "edit permits", sLevel	' In common.asp

if request.servervariables("REQUEST_METHOD") = "POST" then
	intPermitid = 0
	'Process Request
	if request.form("approve") = "yes" then 
		importPermit
	else
		denyPermit
	end if

	if intPermitID <> "0"  then response.redirect "permitedit.asp?permitid=" & intPermitID
	response.redirect "permitapplications.asp"
end if
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="//www.egovlink.com/permitcity/permits/permitapp.css" />
	<style>
		table.contactfields 
		{
			background: inherit !important;
		}

	</style>


	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script>
		var updateTarget = "";
		function searchStreet(x)
		{
			if ($("#customfield" + x + "_streetname").val() == "" && $("#customfield" + x + "_streetnumber").val() == "" )
			{
				//alert('Please enter either a number or street name before searching.');
				//$("#customfield" + x + "_streetname").focus();
			}
			else 
			{
				updateTarget = "customfield" + x + "_matchtarget";
				// Try to get a drop down of names
				doAjax('getaddresses.asp', 'fieldname=customfield' + x + '_matchid&searchstreet=' + $("#customfield" + x + "_streetname").val() + '&searchnumber=' + $("#customfield" + x + "_streetnumber").val() + '&searchowner=', 'updateAddress', 'get', '0');
			}
		
		}
		function searchContact(x)
		{
			if ($("#customfield" + x + "_firstname").val() != "" || $("#customfield" + x + "_lastname").val() != "" || $("#customfield" + x + "_company").val() != "" )
			{
				// Try to get a drop down of names
				updateTarget = "customfield" + x + "_matchtarget";
				doAjax('getpermitapplicants.asp', 'contactsonly=yes&fieldname=customfield' + x +'_matchid&searchfirstname=' + $("#customfield" + x + "_firstname").val() + '&searchlastname=' + $("#customfield" + x + "_lastname").val() + '&searchcompanyname=' + $("#customfield" + x + "_company").val(), 'updateContact', 'get', '0');
			}
			else
			{
				//alert('Please enter a name before searching.');
				//$("#searchname").focus();
			}
		}
		function updateAddress( sResult )
		{
			$("#" + updateTarget).html( sResult );
		}
		function updateContact( sResult )
		{
			$("#" + updateTarget).html( sResult );
		}
	</script>

	<script language="Javascript">
	<!--
	

	//-->
  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );


  function importPA()
  {
	  if (validateForm("permitAppFrm"))
	  {
		  document.permitAppFrm.submit();
	  }
  }
  function denyPA()
  {
	  document.getElementById("approve").value = "no";
	  document.permitAppFrm.submit();
  }
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="notcentercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<strong style="font-size:16px">Permit Application</strong>
			</p>
			<!--END: PAGE TITLE-->

				<input type="button" class="button ui-button ui-widget ui-corner-all" value="<< Back to Permit Application List" onclick="javascript:window.location='permitapplications.asp';" />
<%
sSQL = "SELECT q.*, ft.*, a.answer, pas.permitapplication_submittedid as applicationid, pas.permittypeid,pas.userid " _
	& " FROM egov_permitapplication_submitted pas " _
	& " INNER JOIN egov_permitapplication pa ON pas.permitapplicationid = pa.permitapplicationid " _
	& " INNER JOIN egov_permitapplication_questions q ON q.permitapplicationid = pa.permitapplicationid  " _
	& " INNER JOIN egov_permitapplication_fieldtypes ft ON ft.fieldtypeid = q.fieldtypeid " _
	& " LEFT JOIN egov_permitapplication_answers a ON a.permitapplication_submittedid = pas.permitapplication_submittedid and q.permitapplicationquestionid = a.permitapplicationquestionid " _
	& " WHERE pas.orgid = '" & session("orgid") & "' and pas.permitapplication_submittedid = '" & dbsafe(request("paid")) & "' and parentquestionid = 0  " _
	& " ORDER BY q.DisplayOrder "

Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1


blnFoundForm = false
if not oRs.EOF then
	intPermitTypeID = oRs("permittypeid")

	sSQL = "SELECT userfname,userlname " _
		& " FROM egov_users " _
		& " WHERE userid = " & oRs("userid")
	Set oU = Server.CreateObject("ADODB.RecordSet")
	oU.Open sSQL, Application("DSN"), 3, 1
	if not oU.EOF then
		response.write "<h2>Applicant: " & oU("userlname") & ", " & oU("userfname") & "</h2>"
	end if
	oU.Close
	Set oU = Nothing

	blnFoundForm = true%>
	<form id="permitAppFrm" name="permitAppFrm" action="permitapplication.asp" method="POST">
		<input type="hidden" name="orgid" value="<%=iorgid%>" />
		<input type="hidden" id="userid" name="userid" value="<%=oRs("userid")%>" />
		<input type="hidden" id="applicationid" name="applicationid" value="<%=oRs("applicationid")%>" />
		<input type="hidden" id="approve" name="approve" value="yes" />
<% end if

iCount = 0
Do While not oRs.EOF
	ShowQuestion oRs("questionname"), oRs("fieldtypebehavior"), oRs("options"), oRs("instructions"), oRs("permitapplicationquestionid"), oRs("parentquestionid"), oRs("isrequired"), oRs("defaultopened"), oRs("answer")

	oRs.MoveNext
loop

if blnFoundForm then%>
	<button id="submitBtn" type="button"  class="button ui-button ui-widget ui-corner-all" onClick="importPA();">Import Application</button>
	<br />
	<br />
	<br />
	<button id="submitBtn" type="button" class="button ui-button ui-widget ui-corner-all" onClick="denyPA();">Deny Application</button>
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




			</div>
		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  
	<!--#Include file="modal.asp"-->  

</body>

</html>


<%
sub ShowQuestion(strQuestionName, strFieldTypeBehavior, strOptions, strInstructions, intQuestionID, strParentQuestionID, blnIsRequired, blnDefaultOpened, strAnswer)
	strRequiredFlag = ""
	'if blnIsRequired then strRequiredFlag = "<span class=""required"">*</span>"

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
	'if blnIsRequired then strReq = "/req"

	Select Case strFieldTypeBehavior
		
		Case "radio"
			aChoices = Split(strOptions,Chr(10))
	
			for x = 0 to UBound(aChoices)
				sChoice = Replace(aChoices(x), Chr(13), "")
				response.write vbcrlf & "<input type=""radio"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sChoice & """"
				if instr(strAnswer,sChoice) then response.write " checked"
				response.write " /> " & sChoice & "<br />" & vbcrlf
			Next
			transferOptions iCount, intQuestionID
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-radio" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
			
		Case "select"
			response.write vbcrlf & "<select id=""customfield" & iCount & """ name=""customfield" & iCount & """>" & vbcrlf
			aChoices = Split(strOptions,Chr(10))
	
			for x = 0 to UBound(aChoices)
				sChoice = Replace(aChoices(x), Chr(13), "")
				response.write vbcrlf & "<option value=""" & sChoice & """"
				if strAnswer = sChoice then response.write " selected"
				response.write " /> " & sChoice & "</option>" & vbcrlf
			Next
			response.write vbcrlf & "</select>" & vbcrlf
			transferOptions iCount, intQuestionID
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-select" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
		
		Case "checkbox"
			aChoices = Split(strOptions,Chr(10))
	
			For x = 0 To UBound(aChoices)
				sChoice = Replace(aChoices(x), Chr(13), "")
				response.write vbcrlf & "<input type=""checkbox"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & sChoice & """"
				if instr(strAnswer,sChoice) then response.write " checked"
				response.write " /> " & sChoice & "<br />" & vbcrlf
			Next
			transferOptions iCount, intQuestionID
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-checkbox" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf

		Case "textbox","email","phone"
			response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & strAnswer & """ size=""100"" />" & vbcrlf
			transferOptions iCount, intQuestionID
			strAddSet = ""
			if strFieldTypeBehavior = "email" then strAddSet = "/email"
			if strFieldTypeBehavior = "phone" then strAddSet = "/phone"
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-text" & strAddSet & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf

		Case "textarea"
			response.write "<textarea class=""customfields"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ rows=""25"">" & strAnswer & "</textarea>" & vbcrlf
			transferOptions iCount, intQuestionID

		Case "date"
			response.write "<input type=""text"" class=""datepicker"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & strAnswer & """  size=""10"" maxlength=""10"" />" & vbcrlf
			transferOptions iCount, intQuestionID
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-text/date" & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
		
		Case "integer","decimal","money"
			if strFieldTypeBehavior = "money" then strFieldTypeBehavior = "currency"
			response.write "<input type=""text"" id=""customfield" & iCount & """ name=""customfield" & iCount & """ value=""" & strAnswer & """ size=""50"" />" & vbcrlf
			transferOptions iCount, intQuestionID
			response.write "<input type=""hidden"" name=""ef:customfield" & iCount & "-text/" & strFieldTypeBehavior & strReq & """ value=""" & strQuestionName & """ />" & vbcrlf
		case "permittype"
			ShowPermitTypePicks intPermitTypeID, strReq, strQuestionName
			response.write "(will transfer to Permit Type)"
		case "message"
			response.write strQuestionName
		case "permittype"
			ShowPermitTypePicks strAnswer, strReq, strQuestionName
		case "usetype"
			ShowUseTypes strAnswer, strReq, strQuestionName, iCount
		case "useclass"
			ShowUserClasses strAnswer, strReq, strQuestionName, iCount
		case "workclass"
			ShowWorkClass strAnswer, strReq, strQuestionName, iCount
		case "workscope"
			ShowWorkScopes strAnswer, strReq, strQuestionName, iCount
		case "constructiontype"
			ShowConstructionTypes strAnswer, strReq, strQuestionName, iCount
		case "occtype"
			ShowOccupancyTypes strAnswer, strReq, strQuestionName, iCount
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
			<button type="button" class="collapsible <%=strDefaultOpened%>"><%=strQuestionName & strRequiredFlag%><%transferOptions iCount, intQuestionID%><%
			if strInstructions <> "" then
				response.write "&nbsp;(" & strInstructions & ")"
			end if

			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " style=""display:block"""


			'Need to Pull Answer Data
			'Need to Pull Answer Data
			sSQL = "SELECT * FROM egov_permitapplication_contacts WHERE permitapplicationcontactid = '" & strAnswer & "'"
			Set oContact = Server.CreateObject("ADODB.RecordSet")
			oContact.Open sSQL, Application("DSN"), 3, 1
			if not oContact.EOF then
				strFirstName = oContact("firstname")
				strLastName = oContact("lastname")
				strCompany = oContact("company")
				strAddress = oContact("address")
				strCity = oContact("city")
				strState = oContact("state")
				strZip = oContact("zip")
				strEmail = oContact("email")
				strPhone = oContact("phone")
				strCell = oContact("cell")
				strFax = oContact("fax")
			else
				strFirstName = ""
				strLastName = ""
				strCompany = ""
				strAddress = ""
				strCity = ""
				strState = ""
				strZip = ""
				strEmail = ""
				strPhone = ""
				strCell = ""
				strFax = ""
			end if
			oContact.Close
			Set oContact = Nothing

			%> <div class="expand-icon"><i class="fa fa-bars"></i></div></button>
			<div class="content" <%=strDefaultOpened%>>
				<br />
				<table class="contactfields">
					<tr><td class="label">First Name<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_firstname" name="customfield<%=iCount%>_firstname" size="50" maxlength="50" value="<%=strFirstName%>" /></td></tr>
					<tr><td class="label">Last Name<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_lastname" name="customfield<%=iCount%>_lastname" size="50" maxlength="50" value="<%=strLastName%>" /></td></tr>
					<tr><td class="label">Company<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_company" name="customfield<%=iCount%>_company" size="50" maxlength="100" value="<%=strCompany%>" /></td></tr>
					<tr><td class="label">Address<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_address" name="customfield<%=iCount%>_address" size="50" maxlength="50" value="<%=strAddress%>" /></td></tr>
					<tr><td class="label">City<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_city" name="customfield<%=iCount%>_city" size="50" maxlength="50" value="<%=strCity%>" /></td></tr>
					<tr><td class="label">State<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_state" name="customfield<%=iCount%>_state" size="50" maxlength="50" value="<%=strState%>" /></td></tr>
					<tr><td class="label">Zip<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_zip" name="customfield<%=iCount%>_zip" size="50" maxlength="50" value="<%=strZip%>" /></td></tr>
					<tr><td class="label">Email<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_email" name="customfield<%=iCount%>_email" size="50" maxlength="100" value="<%=strEmail%>" /></td></tr>
					<tr><td class="label">Phone<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_phone" name="customfield<%=iCount%>_phone" size="50" maxlength="10" value="<%=strPhone%>" /></td></tr>
					<tr><td class="label">Cell<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_cell" name="customfield<%=iCount%>_cell" size="50" maxlength="10" value="<%=strCell%>" /></td></tr>
					<tr><td class="label">Fax<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_fax" name="customfield<%=iCount%>_fax" size="50" maxlength="10" value="<%=strFax%>" /></td></tr>
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
			<div style="background-color:#f1f1f1;padding:0 18px 10px 18px;">
				<span id="customfield<%=iCount%>_matchtarget">No Match Found</span>
				<script>searchContact(<%=iCount%>);</script>
				<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="searchContact(<%=iCount%>);" />
				<input type="checkbox" class="override" name="customfield<%=iCount%>_newcontact" />Create New Contact
			</div>
			<%
		case "address"
			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " active"
			%>
			<input type="hidden" id="customfield<%=iCount%>_fieldtype" name="customfield<%=iCount%>_fieldtype" value="address" />
			<button type="button" class="collapsible <%=strDefaultOpened%>"><%=strQuestionName & strRequiredFlag%><%transferOptions iCount, intQuestionID%><%
			if strInstructions <> "" then
				response.write "&nbsp;(" & strInstructions & ")"
			end if

			strDefaultOpened = ""
			if blnDefaultOpened or blnRequired then strDefaultOpened = " style=""display:block"""

			'Need to Pull Answer Data
			sSQL = "SELECT * FROM egov_permitapplication_address WHERE permitapplicationaddressid = '" & strAnswer & "'"
			Set oAdd = Server.CreateObject("ADODB.RecordSet")
			oAdd.Open sSQL, Application("DSN"), 3, 1
			if not oAdd.EOF then
				strStreetNumber = oAdd("residentstreetnumber")
				strStreetName = oAdd("residentstreetname")
				strStreetSuffix = oAdd("streetsuffix")
				strCity = oAdd("residentcity")
				strState = oAdd("residentstate")
				strZip = oAdd("residentzip")
				strCounty = oAdd("county")
			else
				strStreetNumber = ""
				strStreetName = ""
				strStreetSuffix = ""
				strCity = ""
				strState = ""
				strZip = ""
				strCounty = ""
			end if
			oAdd.Close
			Set oAdd = Nothing

			%> <div class="expand-icon"><i class="fa fa-bars"></i></div></button>
			<div class="content" <%=strDefaultOpened%>>
				<br />
				<table class="contactfields">
					<tr><td class="label">Street Number<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_streetnumber" name="customfield<%=iCount%>_streetnumber" size="50" maxlength="15" value="<%=strStreetNumber%>" /></td></tr>
					<tr><td class="label">Street Name<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_streetname" name="customfield<%=iCount%>_streetname" size="50" maxlength="50" value="<%=strStreetName%>" /></td></tr>
					<tr><td class="label">Street Suffix<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_streetsuffix" name="customfield<%=iCount%>_streetsuffix" size="50" maxlength="15" value="<%=strStreetSuffix%>" /></td></tr>
					<tr><td class="label">City<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_city" name="customfield<%=iCount%>_city" size="50" maxlength="50" value="<%=strCity%>" /></td></tr>
					<tr><td class="label">State<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_state" name="customfield<%=iCount%>_state" size="50" maxlength="50" value="<%=strState%>" /></td></tr>
					<tr><td class="label">Zip<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_zip" name="customfield<%=iCount%>_zip" size="50" maxlength="10" value="<%=strZip%>" /></td></tr>
					<tr><td class="label">County<%=strRequiredFlag%>:</td><td><input type="text" id="customfield<%=iCount%>_county" name="customfield<%=iCount%>_county" size="50" maxlength="50" value="<%=strCounty%>" /></td></tr>
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
			<div style="background-color:#f1f1f1;padding:0 18px 10px 18px;">
				<span id="customfield<%=iCount%>_matchtarget"></span>
				<script>searchStreet(<%=iCount%>);</script>
				<input type="button" class="button ui-button ui-widget ui-corner-all" value="Search" onclick="searchStreet(<%=iCount%>);" />
				<input type="checkbox" class="override" name="customfield<%=iCount%>_newaddress" />Create New Address
			</div>
			<%
	
	End Select 

	response.flush
	if strFieldTypeBehavior <> "grouping" then GetChildQuestions intQuestionID

	response.write "<br />" & vbcrlf
	response.write "<br />" & vbcrlf
	response.flush
end sub

sub transferOptions(iCount,intQuestionID)
	sSQL = "SELECT * FROM permitapplication_transfertypesforquestions WHERE orgid = '" & session("orgid") & "' AND (permittypeid = 0 or permittypeid = " & intPermitTypeID & ") and questionid = " & intQuestionID
	Set oTo = Server.CreateObject("ADODB.RecordSet")
	oTo.Open sSQL, Application("DSN"), 3, 1
	response.write "<select name=""customfield" & iCount & "_transfer"" class=""transferoptions"" onclick=""event.stopPropagation();"">"
	if oTo.EOF then 
		response.write "<option value=""0"">No Transfer Matches</option>"
	else
		response.write "<option value=""0"">Do Not Transfer</option>"
	end if
	Do While Not oTo.EOF
		strCF = ""
		if oTo("customfield") then strCF = "CF"
		strSelected = ""
		if oTo("defaulttransfertypeid") & "" = strCF & oTo("rowid") then strSelected = " selected"
		response.write "<option value=""" & strCF & oTo("rowid") & """" & strSelected & ">Transfer to " & oTo("transfertype") & "</option>" & vbcrlf
		oTo.MoveNext
	loop
	response.write "</select>"
	oTo.Close
	Set oTo = Nothing
end sub

Sub GetChildQuestions(intQuestionID)
sSQL = "SELECT q.*, ft.*, a.answer, pas.permittypeid " _
	& " FROM egov_permitapplication_submitted pas " _
	& " INNER JOIN egov_permitapplication pa ON pas.permitapplicationid = pa.permitapplicationid " _
	& " INNER JOIN egov_permitapplication_questions q ON q.permitapplicationid = pa.permitapplicationid  " _
	& " INNER JOIN egov_permitapplication_fieldtypes ft ON ft.fieldtypeid = q.fieldtypeid " _
	& " LEFT JOIN egov_permitapplication_answers a ON a.permitapplication_submittedid = pas.permitapplication_submittedid and q.permitapplicationquestionid = a.permitapplicationquestionid " _
	& " WHERE pas.orgid = '" & session("orgid") & "' and pas.permitapplication_submittedid = '" & dbsafe(request("paid")) & "' and parentquestionid = '" & intQuestionID & "'  " _
	& " ORDER BY q.DisplayOrder "
	'response.write sSQL & "<br />" & vbcrlf
	'response.end

	Set oChild = Server.CreateObject("ADODB.RecordSet")
	oChild.Open sSQL, Application("DSN"), 3, 1
	Do While Not oChild.EOF
		ShowQuestion oChild("questionname"), oChild("fieldtypebehavior"), oChild("options"), oChild("instructions"), oChild("permitapplicationquestionid"), oChild("parentquestionid"), oChild("isRequired"), oChild("defaultopened"), oChild("answer")
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
	sSql = sSql & " WHERE orgid = "& session("orgid")
	sSql = sSql & " ORDER BY permittype, permittypedesc, permittypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<input type=""hidden"" name=""permittypeid"" value=""" & iPermitTypeId & """ />"
		response.write vbcrlf & "<select id=""READONLYpermittypeid"" name=""READONLYpermittypeid"" disabled>"
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


Sub importPermit()

	intPermitTypeID = request.form("permittypeid")
	iAdminUserId = CLng(session("userid"))
	sPermitPrefix = GetPermitPrefix( intPermitTypeID ) 
	sPermitNoYear = "'" & CStr(Year(Date())) & "'"
	sExpirationDate = GetExpirationDate( intPermitTypeID )
	iUseTypeId = GetPermitTypeUseType( intPermitTypeID )
	iPermitLocationRequirementId = GetPermitTypeLocationRequirementId( intPermitTypeID )
	iPermitStatusId = GetInitialPermitStatusId()
	iPermitCategoryId = GetPermitTypeCategoryId( intPermitTypeID )

	strPermitTableFields = "orgid,permitnumberprefix,permitnumberyear,applicantuserid,adminuserid,permittypeid, permitstatusid, applieddate, expirationdate, " _
				& " lastactivitydate,isbuildingpermit, usetypeid, permitcategoryid, permitlocationrequirementid, "
	strPermitTableValues = session("orgid") & ", " & sPermitPrefix & ", " & sPermitNoYear & ", " & "'" & request.form("userid") & "', '" & iAdminUserId & "', " _
				& intPermitTypeID & ", " & iPermitStatusId & ",  dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), " & sExpirationDate & ", " _
				& " dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), 1, " & iUseTypeId & ", " & iPermitCategoryId & ", " _
				& iPermitLocationRequirementId & ", "
	strPermitAddresses = ""
	strPermitContacts = ""


	Set transferTypes = CreateObject("Scripting.Dictionary")
	sSQL = "SELECT * FROM egov_permitapplication_transfertypes"
	Set oTT = Server.CreateObject("ADODB.RecordSet")
	oTT.Open sSQL, Application("DSN"), 3, 1
	Do While Not oTT.EOF
		''response.write oTT("permitapplication_transfertypeid") & " - " & oTT("fieldname") & "<br />" & vbcrlf
		''response.flush
  		transferTypes.Add "" & oTT("permitapplication_transfertypeid"), oTT("fieldname") & ""
		''transferTypes.Item(oTT("permitapplication_transfertypeid")) = oTT("fieldname")
  		oTT.MoveNext
	Loop
	oTT.Close
	Set oTT = Nothing

	for x = 1 to 200
		intQuestionID = getField(x, "_questionid")
		strFieldType = getField(x, "_fieldtype")
		intTransferType = getField(x, "_transfer")

		if strFieldType = "address" then
			'Address Table
			'Populate Match Data OR NEW
			if getField(x, "_newaddress") = "on" or (getField(x, "_matchid") <> "0" and getField(x, "_matchid") <> "") then
				intID = ""
				if getField(x, "_newaddress") = "" and getField(x, "_matchid") <> "0" and getField(x, "_matchid") <> "" then
					intID = getField(x, "_matchid")
				elseif getField(x, "_newaddress") = "on" then
					'Enter new Contact, get rowid
					sSQL = "INSERT INTO egov_residentaddresses (orgid, residentstreetnumber, residentstreetname, streetsuffix, sortstreetname, residentcity, residentstate, residentzip, county) " _
						& " VALUES(" _
						& "'" & session("orgid") & "'," _
						& "'" & getField(x, "_streetnumber") & "'," _
						& "'" & getField(x, "_streetname") & "'," _
						& "'" & getField(x, "_streetsuffix") & "'," _
						& "'" & getField(x, "_streetname") & " " & getField(x, "_streetsuffix") & "'," _
						& "'" & getField(x, "_city") & "'," _
						& "'" & getField(x, "_state") & "'," _
						& "'" & getField(x, "_zip") & "'," _
						& "'" & getField(x, "_county") & "')"

					

					intID = RunIdentityInsert( sSQL )
				end if

				if intID <> "" then
					strPermitAddresses = strPermitAddresses _
						& "INSERT INTO egov_permitaddress (permitid, residentaddressid, residentstreetnumber, residentunit,  " _
							& " residentstreetprefix, residentstreetname, sortstreetname, residentcity, residentstate,  " _
							& " residentzip, county, orgid, residenttype, latitude, longitude, parcelidnumber, legaldescription,  " _
							& " listedowner, registereduserid, streetsuffix, streetdirection, propertytaxnumber, lotnumber,  " _
							& " lotwidth, lotlength, blocknumber, subdivision, section, township, range,  " _
							& " permanentrealestateindexnumber, collectorstaxbillvolumenumber ) " _
						& " SELECT ###PERMITID###, residentaddressid, residentstreetnumber, residentunit,  " _
								& " residentstreetprefix, residentstreetname, sortstreetname, residentcity, residentstate,  " _
								& " residentzip, county, orgid, residenttype, latitude, longitude, parcelidnumber, legaldescription,  " _
								& " listedowner, registereduserid, streetsuffix, streetdirection, propertytaxnumber, lotnumber,  " _
								& " lotwidth, lotlength, blocknumber, subdivision, section, township, range,  " _
								& " permanentrealestateindexnumber, collectorstaxbillvolumenumber " _
							& " FROM egov_residentaddresses " _
							& " WHERE residentaddressid = " & intID & "; "
				end if
			end if

		elseif strFieldType = "contact" then
			strContactType = transferTypes.Item(intTransferType)
			'Contact Table
			'Populate Match Data OR NEW
			intID = ""
			if strContactType <> "" and (getField(x, "_newcontact") = "on" or (getField(x, "_matchid") <> "0" and getField(x, "_matchid") <> "")) then
				if getField(x, "_newcontact") = "" and getField(x, "_matchid") <> "0" and getField(x, "_matchid") <> "" then
					intID = mid(getField(x, "_matchid"),2)
				elseif getField(x, "_newcontact") = "on" then
					'Enter new Contact, get rowid
					sSQL = "INSERT INTO egov_permitcontacttypes (orgid, company, firstname, lastname, address, city, state, zip, email, phone, cell, fax) " _
						& " VALUES(" _
						& "'" & session("orgid") & "'," _
						& "'" & getField(x, "_company") & "'," _
						& "'" & getField(x, "_firstname") & "'," _
						& "'" & getField(x, "_lastname") & "'," _
						& "'" & getField(x, "_address") & "'," _
						& "'" & getField(x, "_city") & "'," _
						& "'" & getField(x, "_state") & "'," _
						& "'" & getField(x, "_zip") & "'," _
						& "'" & getField(x, "_email") & "'," _
						& "'" & getField(x, "_phone") & "'," _
						& "'" & getField(x, "_cell") & "'," _
						& "'" & getField(x, "_fax") & "')"
					

					intID = RunIdentityInsert( sSQL )
				end if

				if intID <> "" then
					strPermitContacts = strPermitContacts _
						& "INSERT INTO egov_permitcontacts ( permitid, permitcontacttypeid, orgid, company, firstname,  " _
							& " lastname, address, city, state, zip, email, phone, cell, fax,  " _
 							& " is" & strContactType & ", " _
  							& " userid, userpassword, userworkphone, emergencycontact, emergencyphone,  " _
 							& " neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable,  " _
 							& " residencyverified, contractortypeid, businesstypeid, statelicense, employeecount, reference1,  " _
 							& " reference2, reference3, otherlicensedcity1, otherlicensedcity2, generalliabilityagent,  " _
 							& " generalliabilityphone, workerscompagent, workerscompphone, autoinsuranceagent,  " _
 							& " autoinsurancephone, bondagent, bondagentphone ) " _
 						& " SELECT ###PERMITID###, permitcontacttypeid, ct.orgid, company, firstname,  " _
							& " lastname, address, city, state, zip, email, phone, cell, fax, 1, ct.userid,  " _
 							& " userpassword, userworkphone, emergencycontact, emergencyphone,  " _
 							& " neighborhoodid, residenttype, userbusinessaddress, userunit, ISNULL(emailnotavailable,0),  " _
 							& " ISNULL(residencyverified,0),  " _
 							& " contractortypeid, businesstypeid, statelicense, employeecount, reference1,  " _
 							& " reference2, reference3, otherlicensedcity1, otherlicensedcity2, generalliabilityagent,  " _
 							& " generalliabilityphone, workerscompagent, workerscompphone, autoinsuranceagent,  " _
 							& " autoinsurancephone, bondagent, bondagentphone " _
 						& " FROM egov_permitcontacttypes ct " _
 						& " LEFT JOIN egov_users u ON u.userid = ct.userid " _
  						& " WHERE permitcontacttypeid = " & intID & "; "

					strPermitContacts = strPermitContacts _
						& "INSERT INTO egov_permitcontacts_licenses ( permitcontactid, permitid, licensetypeid, licensenumber, licenseenddate, licensee ) " _
						& "SELECT pc.permitcontactid, pc.permitid, ctl.licensetypeid, ctl.licensenumber, ctl.licenseenddate, ctl.licensee " _
						& "FROM egov_permitcontacttype_licenses ctl " _
						& "INNER JOIN egov_permitcontacts pc ON pc.permitcontacttypeid = ctl.permitcontacttypeid " _
						& "WHERE pc.permitid = ###PERMITID###; "
				end if

			end if

		elseif isnumeric(intTransferType) and intTransferType <> "" and intTransferType <> "0" then
			'Permit Table
			strPermitTableFields = strPermitTableFields & transferTypes.Item(intTransferType) & ", "
			strPermitTableValues = strPermitTableValues & "'" & getField(x,"") & "', "

		elseif instr(intTransferType,"CF") then
			'Custom Field
			'This will likely be something where I store these values and then I'll need to populate them in later after the CFs are copied
		end if
	next

	if right(strPermitTableFields,2) = ", " then strPermitTableFields = left(strPermitTableFields,len(strPermitTableFields)-2)
	if right(strPermitTableValues,2) = ", " then strPermitTableValues = left(strPermitTableValues,len(strPermitTableValues)-2)
	
	if strPermitTableFields <> "" then
		'Insert into Permit Table
		sSQL = "INSERT INTO egov_permits (" & strPermitTableFields & ") VALUES(" & strPermitTableValues & ")"
		'response.write sSQL & "<br />"
		intPermitID = RunIdentityInsert( sSQL )
	
		'Insert into Address Table
		if strPermitAddresses <> "" then RunSQL replace(strPermitAddresses,"###PERMITID###",intPermitID)

	
		'Insert into Contacts Table
		if strPermitContacts <> "" then RunSQL replace(strPermitContacts,"###PERMITID###",intPermitID)

		' Get the permit Type info and create an entry for that
		'CreatePermitPermitType iPermitId, iPermitTypeId
		sSQL = "INSERT INTO egov_permitpermittypes (permitid, permittypeid, orgid, permittype, permittypedesc, isbuildingpermittype, permitlocationrequirementid, permitcategoryid, expirationdays, isautoapproved, displayorder, permitnumberprefix, publicdescription,  " _
                         & " approvingofficial, permittitle, permitsubtitle, permitrighttitle, permittitlebottom, additionalfooterinfo, permitfooter, permitsubfooter, permitlogo, listfixtures, showfeetotal, showjobvalue, showfootages, showconstructiontype,  " _
                         & " showoccupancytype, showworkdesc, showproposeduse, showothercontacts, showelectricalcontractor, showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, groupbyinvoicecategories, invoicelogo,  " _
                         & " invoiceheader, showcounty, showparcelid, showplansby, showprimarycontact, hastempco, hasco, showapprovedasontco, showconsttypeontco, showocctypeontco, showoccupantsontco, showapprovedasonco,  " _
                         & " showconsttypeonco, showocctypeonco, showoccupantsonco, tempcologo, cologo, tempcotitle, tempcosubtitle, cotitle, cosubtitle, tempcoaddress, coaddress, tempcotoptext, cotoptext, tempcobottomtext, cobottomtext,  " _
                         & " tempcocoderef, cocoderef, tempcoapproval, coapproval, tempcofooter, cofooter, tempcosubfooter, cosubfooter, showtotalsqft, showapprovedas, showfeetypetotals, showoccupancyuse, showpayments, documentid,  " _
                         & " attachmentrevieweralert) " _
		& " SELECT " & intPermitID & ", permittypeid, orgid, permittype, permittypedesc, isbuildingpermittype, permitlocationrequirementid, permitcategoryid, expirationdays, isautoapproved, displayorder, permitnumberprefix, publicdescription,  " _
                         & " approvingofficial, permittitle, permitsubtitle, permitrighttitle, permittitlebottom, additionalfooterinfo, permitfooter, permitsubfooter, permitlogo, listfixtures, showfeetotal, showjobvalue, showfootages, showconstructiontype,  " _
                         & " showoccupancytype, showworkdesc, showproposeduse, showothercontacts, showelectricalcontractor, showmechanicalcontractor, showplumbingcontractor, showapplicantlicense, groupbyinvoicecategories, invoicelogo,  " _
                         & " invoiceheader, showcounty, showparcelid, showplansby, showprimarycontact, hastempco, hasco, showapprovedasontco, showconsttypeontco, showocctypeontco, showoccupantsontco, showapprovedasonco,  " _
                         & " showconsttypeonco, showocctypeonco, showoccupantsonco, tempcologo, cologo, tempcotitle, tempcosubtitle, cotitle, cosubtitle, tempcoaddress, coaddress, tempcotoptext, cotoptext, tempcobottomtext, cobottomtext,  " _
                         & " tempcocoderef, cocoderef, tempcoapproval, coapproval, tempcofooter, cofooter, tempcosubfooter, cosubfooter, showtotalsqft, showapprovedas, showfeetypetotals, showoccupancyuse, showpayments, documentid,  " _
                         & " attachmentrevieweralert " _
		& " FROM egov_permitTypes where permittypeid = " & intPermitTypeID
		RunSQL sSQL

		
		'Bring in Applicant Data
		sSQL = "INSERT INTO egov_permitcontacts ( permitid, orgid, company, firstname, " _
			& " lastname, address, city, state, zip, email, phone, cell, fax, contacttype, " _
			& " isapplicant, userid, userpassword, userworkphone, emergencycontact, emergencyphone, " _
			& " neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable, residencyverified ) " _
		& " SELECT " & intPermitID & ", orgid, userbusinessname, userfname, userlname, useraddress, usercity, userstate, userzip, " _
			& " useremail, userhomephone, usercell, userfax, 'U', 1, userid, userpassword, userworkphone, emergencycontact, " _
			& " emergencyphone, neighborhoodid, residenttype, userbusinessaddress, userunit, emailnotavailable," _
			& " residencyverified FROM egov_users WHERE userid = " & request.form("userid")
		RunSQL sSQL
				
		' Bring in the Fees and Fixtures
		sSQL = "INSERT INTO egov_permitfees ( permitid, permitfeetypeid, orgid, isfixturetypefee, isvaluationtypefee, isconstructiontypefee, permitfeeprefix, permitfee, " _
				& " permitfeecategorytypeid, permitfeemethodid, atleastqty, notmorethanqty, baseamount, quantity, unitqty, " _
				& " unitamount, minimumamount, isupfrontfee, isreinspectionfee, isbuildingpermitfee, accountid, isrequired, " _
				& " displayorder, feeamount, amountpaid, includefee, ispercentagetypefee, percentage, " _
				& " upfrontamount, feereportingtypeid, isresidentialunittypefee )" _
			& " SELECT " & intPermitID & ", F.permitfeetypeid, F.orgid, F.isfixturetypefee, F.isvaluationtypefee, F.isconstructiontypefee, " _
				& " ISNULL(F.permitfeeprefix, '') AS permitfeeprefix, ISNULL(F.permitfee, '') AS permitfee,F.permitfeecategorytypeid, " _
				& " F.permitfeemethodid, F.atleastqty, F.notmorethanqty,ISNULL(F.baseamount,0.00) AS baseamount, 0, F.unitqty, " _
				& " F.unitamount, F.minimumamount, F.isupfrontfee, F.isreinspectionfee, F.isbuildingpermitfee, F.accountid, T.isrequired, " _
				& " T.displayorder, CASE M.isFlatFee WHEN 1 THEN ISNULL(F.baseamount,0.00) ELSE 0.00 END, 0.00, 1, F.ispercentagetypefee, " _
				& " F.percentage, ISNULL(F.upfrontamount,0.00) AS upfrontamount, ISNULL(feereportingtypeid,0) AS feereportingtypeid," _
				& " isresidentialunittypefee" _
			& " FROM egov_permitfeetypes F, egov_permittypes_to_permitfeetypes T, egov_permitfeemethods M " _
			& " WHERE T.permitfeetypeid = F.permitfeetypeid AND F.permitfeemethodid = M.permitfeemethodid AND T.permittypeid = " & intPermitTypeID _
			& " ORDER BY T.displayorder, F.permitfeetypeid"
		RunSQL sSQL

		'Update Fee Total
		sSQL = "UPDATE egov_permits SET feetotal = (SELECT ISNULL(SUM(baseamount),0.00) FROM egov_permitfees WHERE permitid = " & intPermitID & " AND isflatfee = 1 and isrequired = 1) WHERE permitid = " & intPermitID & " "
		RunSQL sSql
		
		' Bring in any Reviews
		sSQL = "INSERT INTO egov_permitreviews (orgid, permitid, permittypeid, permitreviewtypeid, permitreviewtype,  " _
		 		& " reviewdescription, revieweruserid, revieworder, isrequired, isincluded, reviewstatusid, notifyonrelease ) " _
			& " SELECT F.orgid, " & intPermitID & ", T.permittypeid, F.permitreviewtypeid, F.permitreviewtype, " _
				& " F.reviewdescription, T.revieweruserid, T.revieworder, T.isrequired, T.isrequired, " _
				& " (SELECT reviewstatusid FROM egov_reviewstatuses WHERE isforpermits = 1 AND isinitialstatus = 1 AND orgid = F.orgid), " _
				& " T.notifyonrelease " _
			& " FROM egov_permitreviewtypes F, egov_permittypes_to_permitreviewtypes T " _
			& " WHERE F.permitreviewtypeid = T.permitreviewtypeid AND T.permittypeid = " & intPermitTypeID _
			& " ORDER BY T.revieworder, F.permitreviewtypeid"
		RunSQL sSql
		
		
		' Bring in any Inspections
		sSQL = "INSERT INTO egov_permitinspections (orgid, permitid, permittypeid, permitinspectiontypeid, permitinspectiontype,  " _
				& " inspectiondescription, inspectoruserid, inspectionorder, isrequired, isfinal, isincluded, routeorder, inspectionstatusid ) " _
			& " SELECT F.orgid, " & intPermitID & ", T.permittypeid, F.permitinspectiontypeid, F.permitinspectiontype, " _
				& " F.inspectiondescription, ISNULL(T.inspectoruserid,0) AS inspectoruserid, T.inspectionorder, T.isrequired, T.isfinal, 1,  999, " _
				& " (SELECT inspectionstatusid FROM egov_inspectionstatuses WHERE isforpermits = 1 AND orgid = F.orgid and isinitialstatus = 1) " _
			& " FROM egov_permitinspectiontypes F, egov_permittypes_to_permitinspectiontypes T " _
			& " WHERE F.permitinspectiontypeid = T.permitinspectiontypeid AND T.permittypeid = " & intPermitTypeID _
			& " ORDER BY T.inspectionorder, F.permitinspectiontypeid"
		RunSQL sSql
		
		' Bring in any Required Licenses
		sSQL = "INSERT INTO egov_permits_to_permitlicensetypes ( permitid, permittypeid, licensetypeid, orgid, licensetype, displayorder, isrequired ) " _
			& " SELECT " & intPermitID & ", p.permittypeid, l.licensetypeid, l.orgid, L.licensetype, L.displayorder, P.isrequired  " _
				& " FROM egov_permitlicensetypes L, egov_permittypes_to_permitlicensetypes P  " _
				& " WHERE L.licensetypeid = P.licensetypeid AND P.permittypeid = " & intPermitTypeID _
				& " ORDER BY L.displayorder "
		RunSQL sSql
		
		
		' Copy any custom fields that this permit type needs
		sSQL = "INSERT INTO egov_permitcustomfields ( permitid, orgid, fieldtypeid, fieldname, pdffieldname, " _
				& " prompt, valuelist, fieldsize, displayorder, customfieldtypeid ) " _
			& " SELECT " & intPermitID & ", F.orgid, F.fieldtypeid, F.fieldname, F.pdffieldname, F.prompt,  " _
				& " ISNULL(F.valuelist,'') AS valuelist, ISNULL(F.fieldsize,0) AS fieldsize, P.customfieldorder, F.customfieldtypeid " _
			& " FROM egov_permitcustomfieldtypes F, egov_permittypes_to_permitcustomfieldtypes P  " _
			& " WHERE F.customfieldtypeid = P.customfieldtypeid AND P.permittypeid = " & intPermitTypeID
		RunSQL sSql

		'Populate Custom Field Values
		
		' Log the permit creation
		MakeAPermitLogEntry intPermitID, "'New Permit Request Created'", "'New Permit Request Created'", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"
		

		'Update Submitted Table to record PermitID
		sSQL = "UPDATE egov_permitapplication_submitted SET workflowid = 3, permitid = " & intPermitID & " WHERE permitapplication_submittedid = '" & request("applicationid") & "'"
		'response.write sSQL & "<br />"
		RunSQL sSQL
	end if

	set transferTypes = Nothing


	'response.end
End Sub

Sub denyPermit()
	sSQL = "UPDATE egov_permitapplication_submitted SET workflowid = 2 WHERE permitapplication_submittedid = '" & request("applicationid") & "'"
	'response.write sSQL
	'response.end
	RunSQL sSQL
End Sub

function getField(intFieldID, strName)

	'response.write "customfield" & intFieldID & strName & " - " & Track_DBsafe(request.form("customfield" & intFieldID & strName)) & "<br />" & vbcrlf

	getField = Track_DBsafe(request.form("customfield" & intFieldID & strName))


end function
'-------------------------------------------------------------------------------------------------
' string GetPermitPrefix( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitPrefix( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(permitnumberprefix,'B') AS permitnumberprefix FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitPrefix = "'" & oRs("permitnumberprefix") & "'"
	Else
		GetPermitPrefix = "''"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 
'-------------------------------------------------------------------------------------------------
' date GetExpirationDate( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetExpirationDate( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(expirationdays,365) AS expirationdays FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetExpirationDate = "'" & DateAdd("d",CLng(oRs("expirationdays")), Date()) & "'"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function
'-------------------------------------------------------------------------------------------------
' integer GetPermitTypeLocationRequirementId( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitTypeLocationRequirementId( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitlocationrequirementid FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTypeLocationRequirementId = CLng(oRs("permitlocationrequirementid"))
	Else 
		GetPermitTypeLocationRequirementId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 
'-------------------------------------------------------------------------------------------------
' integer GetInitialPermitStatusId()
'-------------------------------------------------------------------------------------------------
Function GetInitialPermitStatusId()
	Dim sSql, oRs

	sSql = "SELECT permitstatusid FROM egov_permitstatuses WHERE isinitialstatus = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInitialPermitStatusId = oRs("permitstatusid")
	Else
		GetInitialPermitStatusId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 
'-------------------------------------------------------------------------------------------------
' integer GetPermitTypeCategoryId( iPermitTypeId )
'-------------------------------------------------------------------------------------------------
Function GetPermitTypeCategoryId( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitcategoryid FROM egov_permittypes WHERE permittypeid = " & iPermitTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetPermitTypeCategoryId = CLng(oRs("permitcategoryid"))
	Else 
		GetPermitTypeCategoryId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

%>
