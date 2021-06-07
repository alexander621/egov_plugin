<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

lcl_hidden = "hidden"   'SHOW/HIDE hidden fields.  HIDE = hidden, SHOW = text

If Not UserHasPermission( Session("UserId"), "form creator" ) Then
	  response.redirect sLevel & "permissiondenied.asp"
End If 

iFormID = request("iformid")
'iOrgID  = "10"
If request.servervariables("REQUEST_METHOD") = "POST" Then
	  Select Case request("TASK")

   		Case "new_question"
			    'ADD NEW QUESTION TO FORM
     			Call subAddQuestion(iorgid,iFormID)
   		Case Else
			    'DEFAULT ACTION
  	End Select
End If
%>
<html>
<head>
	<title> E-GovLink Forms Management </title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="JavaScript">
<!--
function confirm_delete(ifieldid,sname) {
  input_box = confirm("Are you sure you want to delete the question that begins (" + sname + ")? \nAll values for this question will be lost.");

  if (input_box==true) { 
      // DELETE HAS BEEN VERIFIED
      location.href='delete_field.asp?iformid=<%=iformid%>&ifieldid='+ ifieldid;
  }else{
      // CANCEL DELETE PROCESS
  }
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <%'DrawTabs tabActionline,1%>
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
    <!--<tr>
      <td><font size="+1"><b>Form Builder</b></font><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
    </tr>-->
  <tr>
      <td valign="top">
<%
	' GET FORM GENERAL INFORMATION
	Dim sTitle
	Dim sIntroText
	Dim sFooterText
	Dim sMask
	Dim blnEmergencyNote
	Dim sEmergencyText
	Dim blnIssueDisplay
	Dim blnFeeDisplay
	Dim sIssueMask
	Dim	iStreetNumberInputType 
	Dim	iStreetAddressInputType 
	Dim sIssueName
	Dim sIssueDesc
 Dim sResolvedStatus

	subGetFormInformation(iFormID)
%>

<!--<div style="margin-top:20px; margin-left:20px;" >-->

<% 'blnCanManageActionAlerts = HasPermission("CanManageActionAlerts") 
	If UserHasPermission( Session("UserId"), "alerts" ) Then
		  blnCanManageActionAlerts = True
	Else
		  blnCanManageActionAlerts = false
	End If 
%>
<p><font class="label">Forms - Edit Form </font> <small>[<a class="edit" href="copy_form.asp?task=copyme&iformid=<%=iFormID%>&iorgid=<%=session("orgid")%>">Copy This Form</a>]</small> <%if blnCanManageActionAlerts then %><small>[<a class="edit" href="../action_line/edit_form.asp?task=name&control=<%=iFormID%>&iorgid=<%=session("orgid")%>">Manage This Form</a>]</small><% end if %> <small>[<a class="edit" href="list_forms.asp">Return to Form List</a>]</small><hr size="1" width="600px;" style="text-align:left; color:#000000;">

<form name="frmResolvedStatus" action="edit_form.asp" method="POST">
<small>[<a class="edit" href="javascript:document.frmResolvedStatus.submit();">Automatically set status to "RESOLVED" upon submission. On/Off</a>] 
       (<font style="font-family:arial;font-size:10px;font-weight:bold;">
<%
if sResolvedStatus = "Y" then
	  response.write "ON"
else
	  response.write "OFF" 
end if
%>
        </font>)</small>
  <input type="<%=lcl_hidden%>" name="iformid" value="<%=iformid%>">
  <input type="<%=lcl_hidden%>" name="iorgid" value="<%=iorgid%>">
  <input type="<%=lcl_hidden%>" name="task" value="RESOLVEDSTATUS">
  <input type="<%=lcl_hidden%>" name="RESOLVEDSTATUS" value="<%=sResolvedStatus%>">
</form>

<p><small>[<a class="edit" href="edit_form.asp?task=name&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Name</a>]</small> - <font class="subtitle"><%=sTitle%></font></p>

<form name="frmEmergencyNote" action="edit_form.asp" method="POST" >
<small>[<a class="edit" href="edit_form.asp?task=enote&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Emergency Note</a>]</small> - <small>[<a class="edit" href="#" onClick="document.frmEmergencyNote.submit();">Toggle Emergency Notice On/Off </a>] (<font style="font-family:arial;font-size:10px;font-weight:bold;">
<%
If blnEmergencyNote Then
	  response.write "ON"
Else
	  response.write "OFF" 
End If
%>
</font>)</small>
<input type="<%=lcl_hidden%>" name="iformid" value="<%=iformid%>">
<input type="<%=lcl_hidden%>" name="iorgid" value="<%=iorgid%>">
<input type="<%=lcl_hidden%>" name="task" value="EMERGENCYNOTE">
<input type="<%=lcl_hidden%>" name="emergencynote" value="<%If blnEmergencyNote Then response.write "0" Else response.write "1" End If%>">
<% If blnEmergencyNote Then %>
<div class="warning"><%=sEmergencyText%></div>
<% End If %>
</form>
<p>
<div class="group">
  <div class="orgadminboxf">

	<!--BEGIN: INTRO INFORMATION -->
		<P><small>[<a class="edit" href="edit_form.asp?task=intro&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Intro</a>]</small></P>
		<P>
		<%If sIntroText <> "" Then
			response.write sIntroText
		Else
			response.write " - <i> Introduction text is currently blank </i> -"
		End If
		%></P>
	<!--END: INTRO INFORMATION -->

	<!--BEGIN: CONTACT INFORMATION -->
		<P><small>[<a class="edit" href="edit_form.asp?task=contact&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Contact</a>]</small></P>
		<P>
			<b><u>Contact Information:</u></b><br /><br />
			<table style="background-color:#e0e0e0;">
			<%DrawContactTable()%>
			</table>
		</p>
	<!--END: CONTACT INFORMATION -->

	<!--BEGIN: ISSUE LOCATION-->
	<% If OrgHasFeature("issue location") Then %>
		<!--BEGIN: ADD PROBLEM LOCATION-->
		<form name="frmIssueForm" action="edit_form.asp" method="POST" >
		<small>[<a class="edit" href="#" onClick="document.frmIssueForm.submit();">Toggle Issue Location On\Off </a>] (<font style="font-family:arial;font-size:10px;font-weight:bold;"><%
		If blnIssueDisplay Then
			  response.write "ON"
		Else
			  response.write "OFF" 
		End If
		%></font>)</small>
		<input type="<%=lcl_hidden%>" name="iformid" value="<%=iformid%>">
		<input type="<%=lcl_hidden%>" name="iorgid" value="<%=iorgid%>">
		<input type="<%=lcl_hidden%>" name="task" value="ENABLEISSUE">
		<input type="<%=lcl_hidden%>" name="ENABLEISSUE" value="<%If blnIssueDisplay Then response.write "0" Else response.write "1" End If%>">
		<%If blnIssueDisplay Then%>

			<!--<P>
			<small>[<a class="edit" href="edit_form.asp?task=issue&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Issue/Problem</a>]</small>
			-->
			
			<P>
				<small>[<a class="edit" href="edit_form.asp?task=issuename&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Name\Description</a>]</small> - <b><u><%=sIssueName%></u></b><br><br>
				<% subDisplayIssueLocation sIssueMask,sIssueQues %>
			</p>

		<%End If%>
		</form>
	<%End If%>
	<!--END:ISSUE LOCATION-->
	
	<!--BEGIN: FORM FIELD INFORMATION -->
		<p> 
		<small>[<a class="edit" href="add_field.asp?iformid=<%=iFormID%>&iorgid=<%=iorgid%>" >Add New Question</a>]</small> </P>
		
		<P><% Call subDisplayQuestions(iFormID,0) %> </P>
		
	<!--END: FORM FIELD INFORMATION -->
	

		<!--BEGIN: ADMINISTRATIVE QUESTIONS-->
	<% If UserHasPermission( Session("UserId"), "internalfields" ) Then %>
	<div style="background-color:#336699;font-weight:bold;color:#FFFFFF;padding:10px;" >Internal Use Only - Administrative Fields</div>
		<p> 
			<small>[<a class="edit" href="add_field.asp?isinternal=1&iformid=<%=iFormID%>&iorgid=<%=iorgid%>" >Add New Internal Question</a>]</small> 
		</P>
		
		<P>
			<% Call subDisplayQuestions(iFormID,1) %> 
		</P>
	<%End If%>
	<!--END: ADMINISTRATIVE QUESTIONS-->

	<!--BEGIN: FEE TOGGLE-->
	<% If UserHasPermission( Session("UserId"), "bzfees" ) Then %>
		<!--BEGIN: ADD PROBLEM LOCATION-->
		<form name=frmFeeForm action="edit_form.asp" method="POST" >
		<small>[<a class="edit" href=# onClick="document.frmFeeForm.submit();">Toggle Fees On\Off </a>] (<font style="font-family:arial;font-size:10px;font-weight:bold;"><%
		If blnFeeDisplay Then
  			response.write "ON"
		Else
		  	response.write "OFF" 
		End If
		%></font>)</small>
		<input type="<%=lcl_hidden%>" name="iformid" value="<%=iformid%>">
		<input type="<%=lcl_hidden%>" name="iorgid" value="<%=iorgid%>">
		<input type="<%=lcl_hidden%>" name="task" value="ENABLEFEE">
		<input type="<%=lcl_hidden%>" name="ENABLEFEE" value="<%If blnFeeDisplay Then response.write "0" Else response.write "1" End If%>">
		</form>
	<%End If%>
	<!--END: FEE TOGGLE-->

	<p><font color=red>*</font><B><i>Required Field</i></b></P>
	
	<!--BEGIN: ENDING NOTES -->
		<P><small>[<a class="edit" href="edit_form.asp?task=footer&iformid=<%=iFormID%>&iorgid=<%=iorgid%>">Edit Footer</a>]</small></P>  
		<P>
		<%If sFooterText <> "" Then
    			response.write sFooterText
		Else
    			response.write " - <i> Footer text is currently blank </i> -"
		End If
		%>
		</P>
	<!--END: ENDING NOTES -->

      </td>
    </tr>
  </table>
	</div>
</div>
<br>

<!--#Include file="../admin_footer.asp"-->

</body>
</html>

<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBGETFORMINFORMATION(IFORMID)
'--------------------------------------------------------------------------------------------------
Sub subGetFormInformation(iFormID)
	
	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_action_request_forms "
 sSQL = SSQL & " WHERE action_form_id=" & iFormID

	Set oForm = Server.CreateObject("ADODB.Recordset")
	oForm.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oForm.EOF Then
		 'POPULATE DATA FROM RECORDSET
		  sTitle                  = oForm("action_form_name")
		  sIntroText              = oForm("action_form_description")
		  sFooterText             = oForm("action_form_footer")
		  sMask                   = oForm("action_form_contact_mask")
		  blnEmergencyNote        = oForm("action_form_emergency_note")
		  sEmergencyText          = oForm("action_form_emergency_text")
		  blnIssueDisplay         = oForm("action_form_display_issue")
		  blnFeeDisplay           = oForm("action_form_display_fees")
		  sIssueMask              = oForm("action_form_issue_mask")
		  iStreetNumberInputType  = oForm("issuestreetnumberinputtype")
		  iStreetAddressInputType = oForm("issuestreetaddressinputtype")
		  sIssueName              = oForm("issuelocationname")
    sResolvedStatus         = oForm("action_form_resolved_status")
    sIssueQues              = oForm("issuequestion")

  		If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
		  	  sIssueName = "Issue/Problem Location:"
  		End If

  		sIssueDesc = oForm("issuelocationdesc")

  		If IsNull(sIssueDesc) Then
       'sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". Provide any additional information on problem location in the box below."
			    sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""Choose street from dropdown"". Provide any additional information on problem location in the box below."
		  End If

	End If

	Set oForm = Nothing 

End Sub

'--------------------------------------------------------------------------------------------------
' FUNCTION DRAWCONTACTTABLE()
'--------------------------------------------------------------------------------------------------
Function DrawContactTable()
%>
	
	<% If IsDisplay(sMASK,1) Then %>
	<tr><td align="right">
		<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><%=IsRequired(sMASK,1)%></span> 
		First Name:
		</span>
	</td><td>
		<span class="cot-text-emphasized" title="This field is required"> 
		<input type="text" value="" name="cot_txtFirst_Name" id="txtFirst_Name" style="width:300;" maxlength="100">
		</span>
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,2) Then %>
	<tr><td align="right">
		<span class="cot-text-emphasized" title="This field is required"><span class="cot-text-emphasized"><%=IsRequired(sMASK,2)%></span>
		Last Name:
		</span>
	</td><td>
		<span class="cot-text-emphasized" title="This field is required">
		<input type="text" value="" name="cot_txtLast_Name" id="txtLast_Name" style="width:300;" maxlength="100">
		</span>
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,3) Then %>
	<tr><td align="right"><%=IsRequired(sMASK,3)%>
		Business Name:
	</td><td>
		<input type="text" value="" name="cot_txtBusiness_Name" id="txtBusiness_Name" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,4) Then %>
	<tr><td align="right">
		<%=IsRequired(sMASK,4)%>Email:
	</td><td>
		<input type="text" value="" name="cot_txtEmail" id="txtEmail" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,5) Then %>
	<tr><td align="right">
		<%=IsRequired(sMASK,5)%>Daytime Phone:
	</td><td>
		<input type="text" value="" name="cot_txtDaytime_Phone" id="txtDaytime_Phone" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,6) Then %>
	<tr><td align="right">
		<%=IsRequired(sMASK,6)%>Fax:
	</td><td>
		<input type="text" value="" name="cot_txtFax" id="txtFax" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,7) Then %>
	<tr><td align="right">
		<%=IsRequired(sMASK,7)%>Street:
	</td><td>
		<input type="text" value="" name="cot_txtStreet" id="txtStreet" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,8) Then %>
	<tr><td align="right">
		<%=IsRequired(sMASK,8)%>City:
	</td><td>
		<input type="text" value="" name="cot_txtCity" id="txtCity" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,9) Then %>
	<tr><td align="right">
		<%=IsRequired(sMASK,9)%>State / Province:
	</td><td>
		<input type="text" value="" name="cot_txtState_vSlash_Province" id="txtState_vSlash_Province" size="5" maxlength="100">
	</td></tr>
	<%End IF%>
	<% If IsDisplay(sMASK,10) Then %>
		<tr><td align="right">
		<%=IsRequired(sMASK,10)%>ZIP / Postal Code:
	</td><td>
		<input type="text" value="" name="cot_txtZIP_vSlash_Postal_Code" id="txtZIP_vSlash_Postal_Code" style="width:300;" maxlength="100">
	</td></tr>
	<%End IF%>

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
' SUB SUBDISPLAYQUESTIONS(IFORMID,BLNTYPE)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayQuestions(iFormID,blnType)

	If blnType = "1" Then
		sSQL = "SELECT * FROM egov_action_form_questions WHERE formid=" & iFormID & " AND isinternalonly = 1 ORDER BY sequence"
	Else
		sSQL = "SELECT * FROM egov_action_form_questions WHERE formid=" & iFormID & " AND (isinternalonly <> 1 OR isinternalonly IS NULL)  ORDER BY sequence"
	End If

	Set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oQuestions.EOF Then
	
		response.write "<table style=""background-color:#e0e0e0;"">"
	
	Do While NOT oQuestions.EOF 
		
		' ENUMERATE QUESTIONS
		
		' DETERMINE IF REQUIRED
		sIsrequired = oQuestions("isrequired")
		If sIsrequired = True Then
			sIsrequired = " <font color=red>*</font> "
		Else
			sIsrequired = ""
		End If

		response.write "<tr><TD><small>[<a class=""edit"" href=""edit_field.asp?iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Edit</a>] [<a class=""edit"" href=# onclick=""confirm_delete('" & oQuestions("questionid") & "','" & JSsafe(Left(oQuestions("prompt"),25)) & "...');"">Delete</a>] [<a class=""edit"" href=""order_field.asp?direction=UP&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move Up</a>] [<a class=""edit"" href=""order_field.asp?direction=down&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move Down</a>] [<a class=""edit"" href=""order_field.asp?direction=top&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move to Top</a>] [<a class=""edit"" href=""order_field.asp?direction=bottom&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move to Bottom</a>]</small></td></tr>"

		Select Case oQuestions("fieldtype")

			Case "2"
			' BUILD RADIO QUESTION
			response.write "<tr><td class=question>" & sIsrequired & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
		   		response.write "<tr><td><input name=""question" & oQuestions("questionid") & """ class=formradio type=radio>" & arrAnswers(alist) & "</td></tr>"
			Next

			Case "4"
			' BUILD SELECT QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			response.write "<tr><td><select class=formselect>"
			For alist = 0 to ubound(arrAnswers)
				response.write "<option>" & arrAnswers(alist) & "</option>" 
			Next
			response.write "</select></td></tr>"

			Case "6"
			' BUILD CHECKBOX QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
			For alist = 0 to ubound(arrAnswers)
				response.write "<tr><td><input class=formcheckbox type=checkbox>" & arrAnswers(alist) & "</td></tr>"
			Next

			Case "8"
			' BUILD TEXT QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><input value="""" type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>"

			Case "10"
			' BUILD TEXTAREA QUESTION
			response.write "<tr><td class=question>" & sIsrequired  & oQuestions("prompt")& "</td></tr>"
			response.write "<tr><td><textarea class=formtextarea></textarea></td></tr>"

			Case Else

		End Select 

		response.write "<tr><TD>&nbsp;</td></tr>"

		oQuestions.MoveNext
	Loop

		response.write "</table>"
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

'------------------------------------------------------------------------------------------------------------
' FUNCTION JSSAFE( STRDB )
'------------------------------------------------------------------------------------------------------------
Function JSsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  strDB = Replace( strDB, "'", "\'" )
  strDB = Replace( strDB, chr(34), "\'" )
  strDB = Replace( strDB, ";", "\;" )
  strDB = Replace( strDB, "-", "\-" )
  strDB = Replace( strDB, chr(10), " " )
  strDB = Replace( strDB, chr(12), " " )
  strDB = Replace( strDB, chr(13), " " )
  JSsafe = strDB
End Function

'------------------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYISSUELOCATION(SISSUEMASK)
'------------------------------------------------------------------------------------------------------------
Sub subDisplayIssueLocation(sIssueMask,sIssueQues)
	
	If sIssueMask = "" or IsNull(sIssueMask) Then
		  sIssueMask = "121111"
	End If

	If trim(sIssueQues) = "" Then
 			sIssueQues = "Provide any additional information on problem location in the box below."
	End If
	
	response.write "<table border=""0"">" & vbcrlf
	response.write "  <tr>"
 response.write "      <td colspan=""2"">" & vbcrlf
	response.write            sIssueDesc & vbcrlf
	response.write "          <p>" & vbcrlf
	response.write "          <table border=""0"">" & vbcrlf
	response.write "            <tr>" & vbcrlf
 response.write "                <td align=""right"">" & vbcrlf
	response.write                      IsRequired("2",1)  & " Address: </td>" & vbcrlf
    if OrgHasFeature("large address list") then 
 response.write "                <td>"
 response.write "                    <input type=""text"" name=""residentstreetnumber"" value="""" size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
	response.write "                    <select name=""skip_address"">" & vbcrlf
	response.write "                      <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
 response.write "                    </select>&nbsp;" & vbcrlf
 response.write "                    <input type=""button"" value=""Validate Address"">" & vbcrlf
 response.write "                </td>" & vbcrlf
    else
	response.write "                <td>" & vbcrlf
 response.write "                    <select style=""width:300;"">" & vbcrlf
	response.write "                      <option>Choose street from dropdown</option>" & vbcrlf
	response.write "                    </select>" & vbcrlf
 response.write "                </td>" & vbcrlf
    end if
	response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
 response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td> - Or Other Not Listed - </td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
 response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td><input name=""ques_issue2"" type=""text"" size=""60""></td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
	response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "            <tr>" & vbcrlf
 response.write "                <td>&nbsp;</td>" & vbcrlf
 response.write "                <td>" & sIssueQues & "<br />" & vbcrlf
 response.write "                    <textarea name=""ques_issue6"" class=""formtextarea""></textarea>" & vbcrlf
 response.write "                </td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
	response.write "          </table>" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf

End Sub

'------------------------------------------------------------------------------------------------------------
' FNDRAWINPUTTYPE(IINPUTTYPE)
'------------------------------------------------------------------------------------------------------------
Function fnDrawInputType(iInputType)

				Select Case iInputType

					Case "1"
						' TEXT BOX ONLY
						sReturnValue = "<input type=text>"

					Case "2"
						' SELECT BOX ONLY
						sReturnValue = "<select><option>Please select street...</select>"

					Case "3"
						' TEXT OR SELECT BOX
						sReturnValue = "<select><option>Please select street...</select> <br> - Or Other Not Listed - <br> "
						sReturnValue = sReturnValue & "<input type=text>"
						
					Case Else
						' DEFAULT TO TEXT BOX ONLY
						sReturnValue = "<input type=text>"

				End Select

		fnDrawInputType = sReturnValue

End Function

%>
