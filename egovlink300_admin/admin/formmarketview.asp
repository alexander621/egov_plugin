<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
 dim sTitle, sIntroText, sFooterText, sMask, blnEmergencyNote, sEmergencyText, blnPublicSearchNote
 dim blnIssueDisplay, blnFeeDisplay, sIssueMask, iStreetNumberInputType
 dim blnMobileOptionsDisplayGeoLoc, blnMobileOptionsDisplayTakePic, blnShowMapInput
 dim iStreetAddressInputType, sIssueName, sIssueDesc, sIssueQues
 dim sResolvedStatus, sHideIssueLocAddInfo, sCustomFOILEmailEdit, appChecked, formName, helpText, lcl_redirectURL

'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 lcl_hidden = "hidden"   'SHOW/HIDE hidden fields.  HIDE = hidden, SHOW = text

 if not UserHasPermission(session("userid"),"form creator") then
 	  response.redirect sLevel & "permissiondenied.asp"
 end if

 iFormID = request("iformid")

'Check for org features
 lcl_orghasfeature_issue_location                       = orghasfeature("issue location")
 lcl_orghasfeature_large_address_list                   = orghasfeature("large address list")
 lcl_orghasfeature_requestmergeforms                    = orghasfeature("requestmergeforms")
 lcl_orghasfeature_actionline_usecustomfoilemailedits   = orghasfeature("actionline_usecustomfoilemailedits")
 lcl_orghasfeature_actionline_formcreator_mobileoptions = orghasfeature("actionline_formcreator_mobileoptions")

'Check for user permissions
 lcl_userhaspermission_alerts            = userhaspermission(session("userid"),"alerts")
 lcl_userhaspermission_internalfields    = userhaspermission(session("userid"),"internalfields")
 lcl_userhaspermission_bzfees            = userhaspermission(session("userid"),"bzfees")
 lcl_userhaspermission_form_letters      = userhaspermission(session("userid"),"form letters")
 lcl_userhaspermission_requestmergeforms = userhaspermission(session("userid"),"requestmergeforms")

'Build the query for the PDF Field Names
 session("CR_PDFFIELDNAMES") = "SELECT orgid FROM organizations WHERE orgid = " & session("orgid")
%>
<html>
<head>
	<title>E-GovLink Administration Consule {Forms Management}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<style type="text/css">
  .fieldset {
      border-radius: 5px;
  }

  .fieldset legend {
      background-color: #ffffff;
      border: 1pt solid #404040;
      border-radius: 5px;
      padding: 4px 10px;
      font-size: 1.125em;
      color: #800000;
  }

  #customFOILEditsNote1 {
      margin-bottom: 5px;
  }

  #customFOILEditsNote2 {
      margin-top: 20px;
      padding-top: 5px;
      border-top:1pt solid #c0c0c0;
  }

  #customFOILButton {
      cursor: pointer;
  } 

</style>

  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<script language="javascript">
<!--
function enableDisableFormOptions(iTaskID, iSectionID, iFormID)
{
    $('#task_' + iTaskID).val(iSectionID);
    $('#' + iFormID).submit();
}

function confirm_delete(ifieldid,sname) {
  input_box = confirm("Are you sure you want to delete the question that begins (" + sname + ")? \nAll values for this question will be lost.");

  if (input_box==true) { 
      // DELETE HAS BEEN VERIFIED
      location.href='delete_field.asp?iformid=<%=iformid%>&ifieldid='+ ifieldid;
  }else{
      // CANCEL DELETE PROCESS
  }
}

function openCustomReports(p_report) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("../customreports/customreports.asp?cr='+p_report+'&iFormID=<%=iFormID%>", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
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
      <td><font size="+1"><strong>Form Builder</strong></font><br /><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<a href="javascript:history.go(-1)"><%=langBackToStart%></a></td>
    </tr>-->
  <tr>
      <td valign="top">
<%
'Retrieve the form data
	subGetFormInformation(iFormID)

'Set up the status labels
 lcl_resolvedstatus_label              = "OFF"
 lcl_emergencynote_label               = "OFF"
 lcl_publicsearchnote_label               = "OFF"
 lcl_issuedisplay_label                = "OFF"
 lcl_mobileOptionsGeoLocDisplay_label  = "OFF"
 lcl_showmapinput_label  = "OFF"
 lcl_mobileOptionsTakePicDisplay_label = "OFF"
 lcl_feedisplay_label                  = "OFF"
 lcl_hideIssueLocAddInfo_label         = "OFF"

'Check Resolved status
 if sResolvedStatus = "Y" then
	   lcl_resolvedstatus_label = "ON"
 end if

'Check Emergency Note status
 if blnEmergencyNote then
    lcl_emergencynote_label = "ON"
 end if
'Check Public Search Note status
 if blnPublicSearchNote then
    lcl_publicsearchnote_label = "ON"
 end if

'Check Issue Display status
 if blnIssueDisplay then
    lcl_issuedisplay_label = "ON"
 end if

'Check Mobile Options Display status
 if blnMobileOptionsDisplayGeoLoc then
    lcl_mobileOptionsGeoLocDisplay_label = "ON"
 end if
 if blnShowMapInput then
    lcl_showmapinput_label = "ON"
 end if

 if blnMobileOptionsDisplayTakePic then
    lcl_mobileOptionsTakePicDisplay_label = "ON"
 end if

'Check Fee Display status
 if blnFeeDisplay then
    lcl_feedisplay_label = "ON"
 end if

'Check hideIssueLocAddInfo status
 if not sHideIssueLocAddInfo then
    lcl_hideIssueLocAddInfo_label = "ON"
 end if

'Check that the user has the "Alerts" feature assigned
 if lcl_userhaspermission_alerts then
		  blnCanManageActionAlerts = True
	else
		  blnCanManageActionAlerts = False
	end if

'Page Title, Buttons: Copy Form ---------
 response.write "<p>" & vbcrlf
 response.write "   <font class=""label"">VIEW MARKET FORM</font>" & vbcrlf
 response.write "   <input type=""button"" name=""copyForm"" id=""copyForm"" value=""Copy This Form"" class=""button"" onclick=""location.href='copy_form.asp?task=copyme&type=market&iformid=" & iFormID & "&iorgid=" & session("orgid") & "'"" />" & vbcrlf
 response.write "   <hr size=""1"" width=""600px;"" style=""text-align:left; color:#000000;"" />" & vbcrlf
 response.write "</p>" & vbcrlf


'Edit Name --------------------------------------------------------------------
 response.write "<p>" & vbcrlf
 'response.write "   <small>[<a class=""edit"" href=""edit_form.asp?task=name&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Name</a>]</small> - " & vbcrlf
 response.write "   <font class=""subtitle"">" & sTitle & "</font>" & vbcrlf
 response.write "</p>" & vbcrlf

'Emergency Note ---------------------------------------------------------------
 if blnEmergencyNote then
    response.write "<div class=""warning"">" & sEmergencyText & "</div>" & vbcrlf
 end if



 response.write "<p>" & vbcrlf
 response.write "<div class=""group"">" & vbcrlf
 response.write "  <div class=""orgadminboxf"">" & vbcrlf


'BEGIN: Intro Information -----------------------------------------------------
 response.write "<p>" & vbcrlf

 if sIntroText <> "" then
 			response.write "<p>" & sIntroText & "</p>" & vbcrlf
 else
 			response.write "<p>- <i> Introduction text is currently blank </i> -</p>" & vbcrlf
 end if

'BEGIN: Contact Information ---------------------------------------------------
 response.write "<p>" & vbcrlf
 response.write "   <strong><u>Contact Information:</u></strong><br /><br />" & vbcrlf
 response.write "   <table style=""background-color:#e0e0e0;"">" & vbcrlf

 DrawContactTable()

 response.write "   </table>" & vbcrlf
 response.write "</p>" & vbcrlf

'BEGIN: Issue Location --------------------------------------------------------
 if lcl_orghasfeature_issue_location then
    if blnIssueDisplay then
       lcl_enableissue = "0"
    else
       lcl_enableissue = "1"
    end if

    if not sHideIssueLocAddInfo then
       lcl_hideIssueLocAddInfo = "1"
    else
       lcl_hideIssueLocAddInfo = "0"
    end if

    response.write "Issue Location: " & lcl_issuedisplay_label & "<br />" & vbcrlf

 end if

'BEGIN: Mobile Fields ---------------------------------------------------------
 if lcl_orghasfeature_actionline_formcreator_mobileoptions then
    lcl_enableMobileOptionsGeoLoc  = "1"
    lcl_enableMobileOptionsTakePic = "1"
    lcl_enableShowMapInput = "1"

    if blnMobileOptionsDisplayGeoLoc then
       lcl_enableMobileOptionsGeoLoc = "0"
    end if

    if blnMobileOptionsDisplayTakePic then
       lcl_enableMobileOptionsTakePic = "0"
    end if

    if blnShowMapInput then
       lcl_enableShowMapInput = "0"
    end if

    response.write "File Uploads: " &  lcl_mobileOptionsTakePicDisplay_label & "<br />" & vbcrlf
    response.write "Map: " &  lcl_showmapinput_label & "<br />" & vbcrlf
 end if
'END: Mobile Fields -----------------------------------------------------------


'BEGIN: Form Field Information ------------------------------------------------
 response.write "<p>" & vbrlf

 subDisplayQuestions iFormID,0

if orghasfeature("actionline_formcreator_appoptions") then
 response.write "</p>" & vbcrlf
    response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend>App Options</legend>" & vbcrlf
    response.write "<input type=""checkbox"" " & appChecked & " name=""formobile"" value=""ON"" /> App Enabled<br />"
    response.write "Form Name for App: <input type=""text"" name=""mobilename"" value=""" & formName & """ style=""width:300px"" maxlength=""100"" /><br />"
    response.write "Help Text: <input type=""text"" name=""mobilehelptext"" value=""" & helpText & """ style=""width:500px"" maxlength=""500"" /><br /><br />"
    response.write "</fieldset><br /><br />" & vbcrlf
   end if

'if request.cookies("user")("userid") = "1710" then
    response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend>Custom Confirmation Page</legend>" & vbcrlf
    response.write "  <p>If you enter a URL below we'll direct people who complete this form to that page instead of the standard confirmation page.</p>" & vbcrlf
    response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""task"" value=""SAVECUSTOMURL"" />" & vbcrlf
    response.write "URL: <input type=""text"" name=""redirectURL"" value=""" & lcl_redirectURL & """ style=""width:500px"" maxlength=""500"" /><br /><br />"
    response.write "</fieldset><br /><br />" & vbcrlf
 'end if

'BEGIN: Administrative Questions ----------------------------------------------
 if lcl_userhaspermission_internalfields then
    response.write "<div style=""background-color:#336699;font-weight:bold;color:#ffffff;padding:10px;"">" & vbcrlf
    response.write "  Internal Use Only - Administrative Fields" & vbcrlf
    response.write "</div>" & vbcrlf
    response.write "<p>" & vbcrlf

    subDisplayQuestions iFormID,1

    response.write "</p>" & vbcrlf
 end if

'BEGIN: Fee -------------------------------------------------------------------
 if lcl_userhaspermission_bzfees then
    if blnFeeDisplay then
       lcl_feevalue = "0"
    else
       lcl_feevalue = "1"
    end if

    response.write "Fees: " & lcl_feedisplay_label & "<br />" & vbcrlf
 end if

'BEGIN: Ending Notes ----------------------------------------------------------
 response.write "<p><font color=""#ff0000"">*</font><strong><i>Required Field</i></strong></p>" & vbcrlf

 if sFooterText <> "" then
    response.write "<p>" & sFooterText & "</p>" & vbcrlf
 else
    response.write "<p>- <i> Footer text is currently blank </i> -</p>" & vbcrlf
 end If
 
 'BEGIN: Custom FOIL Email Edits for Rye --------------------------------------
  if lcl_orghasfeature_actionline_usecustomfoilemailedits then
     response.write "<fieldset class=""fieldset"">" & vbcrlf
     response.write "  <legend>Custom FOIL Email Edits</legend>" & vbcrlf
     response.write "  <div id=""customFOILEditsNote1"">" & sCustomFOILEmailEdit & "</div>" & vbcrlf
     response.write "  <div id=""customFOILEditsNote2"">" & vbcrlf
     response.write "    <strong>NOTE:</strong><br />" & vbcrlf
     response.write "    1. If there IS a value in this field, then the value will show up after the following sentence within " & vbcrlf
     response.write "    the Action Line - Create (public side) email from the web, mobile, and WordPress site(s).<br /><br />" & vbcrlf
     response.write "    <strong>Thank you for submitting your information to City of Rye on {{ date and time }}.</strong><br /><br />" & vbcrlf

     response.write "    2.  If there IS a value in this field, then the following TWO (2) lines will be hidden within the same email:<br /><br />" & vbcrlf
     response.write "    <strong>""Do not reply to this message.  Follow the instructions below...""</strong>" & vbcrlf
     response.write "  </div>" & vbcrlf
     response.write "</fieldset>" & vbcrlf
  end if
 'END: Custom FOIL Email Edits for Rye ----------------------------------------
%>
      </td>
    </tr>
  </table>
	</div>
</div>
<br />

<!--#Include file="../admin_footer.asp"-->

</body>
</html>

<%
'------------------------------------------------------------------------------
sub subGetFormInformation(iFormID)
	
  sSQL = "SELECT action_form_name, "
  sSQL = sSQL & " action_form_description, "
  sSQL = sSQL & " action_form_footer, "
  sSQL = sSQL & " action_form_contact_mask, "
  sSQL = sSQL & " action_form_emergency_note, "
  sSQL = sSQL & " action_form_emergency_text, "
  sSQL = sSQL & " action_form_display_issue, "
  sSQL = sSQL & " action_form_display_fees, "
  sSQL = sSQL & " action_form_issue_mask, "
  sSQL = sSQL & " issuestreetnumberinputtype, "
  sSQL = sSQL & " issuestreetaddressinputtype, "
  sSQL = sSQL & " issuelocationname, "
  sSQL = sSQL & " issuelocationdesc, "
  sSQL = sSQL & " action_form_resolved_status, "
  sSQL = sSQL & " issuequestion, "
  sSQL = sSQL & " hideIssueLocAddInfo, "
  sSQL = sSQL & " customFOILEmailEdits, "
  sSQL = sSQL & " display_mobileoptions_geoloc, "
  sSQL = sSQL & " display_mobileoptions_takepic, "
  sSQL = sSQL & " formobile, "
  sSQL = sSQL & " mobilename, "
  sSQL = sSQL & " mobilehelptext, "
  sSQL = sSQL & " publicsearchrequests, "
  sSQL = sSQL & " redirectURL, "
  sSQL = sSQL & " showmapinput "
  sSQL = sSQL & " FROM egov_action_request_forms "
  sSQL = SSQL & " WHERE action_form_id = " & iFormID

	set oForm = Server.CreateObject("ADODB.Recordset")
	oForm.Open sSQL, Application("DSN"), 3, 1
	
	if not oForm.eof then
		  sTitle                         = oForm("action_form_name")
		  sIntroText                     = oForm("action_form_description")
		  sFooterText                    = oForm("action_form_footer")
		  sMask                          = oForm("action_form_contact_mask")
		  blnEmergencyNote               = oForm("action_form_emergency_note")
		  blnPublicSearchNote            = oForm("publicsearchrequests")
		  sEmergencyText                 = oForm("action_form_emergency_text")
		  blnIssueDisplay                = oForm("action_form_display_issue")
		  blnFeeDisplay                  = oForm("action_form_display_fees")
		  sIssueMask                     = oForm("action_form_issue_mask")
		  iStreetNumberInputType         = oForm("issuestreetnumberinputtype")
		  iStreetAddressInputType        = oForm("issuestreetaddressinputtype")
		  sIssueName                     = oForm("issuelocationname")
    sResolvedStatus                = oForm("action_form_resolved_status")
    sIssueQues                     = oForm("issuequestion")
    sHideIssueLocAddInfo           = oForm("hideIssueLocAddInfo")
		  sCustomFOILEmailEdit           = oForm("customFOILEmailEdits")
    blnMobileOptionsDisplayTakePic = oForm("display_mobileoptions_takepic")
    blnMobileOptionsDisplayGeoLoc  = oForm("display_mobileoptions_geoloc")
    blnShowMapInput  = oForm("showmapinput")
		  lcl_redirectURL = oForm("redirectURL")


  		If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
		  	  sIssueName = "Issue/Problem Location:"
  		End If

  		sIssueDesc = oForm("issuelocationdesc")

  		If IsNull(sIssueDesc) Then
       'sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". "
       'sIssueDesc = sIssueDesc & "Provide any additional information on problem location in the box below."
			    sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""Choose street from dropdown"". "
       sIssueDesc = sIssueDesc & "Provide any additional information on problem location in the box below."
		  End If

		  appChecked = oForm("forMobile")
		  formName = oForm("mobileName")
		  helpText = oForm("mobileHelpText")

	end if

 oForm.close
	set oForm = nothing

end sub

'------------------------------------------------------------------------------
function DrawContactTable()

 'First Name
  if IsDisplay(sMASK,1) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & vbcrlf
     response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
     response.write "            <span class=""cot-text-emphasized"">" & IsRequired(sMASK,1) & "</span>" & vbcrlf
     response.write "            First Name:" & vbcrlf
     response.write "          </span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <span class=""cot-text-emphasized"" title=""This field is required""> " & vbcrlf
     response.write "          <input type=""text"" value="""" name=""cot_txtFirst_Name"" id=""txtFirst_Name"" style=""width:300;"" maxlength=""100"" />" & vbcrlf
     response.write "          </span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Last Name
  if IsDisplay(sMASK,2) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & vbcrlf
     response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
     response.write "            <span class=""cot-text-emphasized"">" & IsRequired(sMASK,2) & "</span>" & vbcrlf
     response.write "            Last Name:" & vbcrlf
     response.write "        		</span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
     response.write "         	<input type=""text"" value="""" name=""cot_txtLast_Name"" id=""txtLast_Name"" style=""width:300;"" maxlength=""100"" />" & vbcrlf
     response.write "         	</span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Business Name
  if IsDisplay(sMASK,3) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,3) & "Business Name:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtBusiness_Name"" id=""txtBusiness_Name"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Email
  if IsDisplay(sMASK,4) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,4) & "Email:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtEmail"" id=""txtEmail"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Daytime Phone
  if IsDisplay(sMASK,5) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,5) & "Daytime Phone:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtDaytime_Phone"" id=""txtDaytime_Phone"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Fax
  if IsDisplay(sMASK,6) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,6) & "Fax:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtFax"" id=""txtFax"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Street
  if IsDisplay(sMASK,7) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,7) & "Street:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtStreet"" id=""txtStreet"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'City
  if IsDisplay(sMASK,8) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,8) & "City:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtCity"" id=""txtCity"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'State
  if IsDisplay(sMASK,9) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,9) & "State / Province:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtState_vSlash_Province"" id=""txtState_vSlash_Province"" size=""5"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Zip
  if IsDisplay(sMASK,10) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">" & IsRequired(sMASK,10) & "ZIP / Postal Code:</td>" & vbcrlf
     response.write "      <td><input type=""text"" value="""" name=""cot_txtZIP_vSlash_Postal_Code"" id=""txtZIP_vSlash_Postal_Code"" style=""width:300;"" maxlength=""100"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

end function

'----------------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'----------------------------------------------------------------------------------------
sub subDisplayQuestions(iFormID,blnType)

	sSQL = "SELECT questionid, formid, orgid, prompt, fieldtype, isenabled, sequence, answerlist, isrequired, validationlist, "
 sSQL = sSQL & " isinternalonly, pdfformname, pushfieldid "
 sSQL = sSQL & " FROM egov_action_form_questions "
 sSQL = sSQL & " WHERE formid = " & iFormID

'Determine which form questions to retrieve
	if blnType = "1" then
  		sSQL = sSQL & " AND isinternalonly = 1 "
 else
		  sSQL = sSQL & " AND (isinternalonly <> 1 OR isinternalonly IS NULL) "
	end if

 sSQL = sSQL & " ORDER BY sequence"

	set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	if not oQuestions.eof then
	  	response.write "<table style=""background-color:#e0e0e0;"">" & vbcrlf

    while not oQuestions.eof
      'Determine if required
     		sIsrequired = oQuestions("isrequired")

       if sIsrequired = True then
       			sIsrequired = " <font color=""#ff0000"">*</font> "
     		else
       			sIsrequired = ""
       end if

    			response.write "  <tr><td class=""question"">" & sIsrequired & oQuestions("prompt") & "</td></tr>" & vbcrlf

      'FieldType values:
      '-----------------
      '  2  = Radio
      '  4  = Select
      '  6  = Checkbox
      '-----------------
      '  8  = Text
      '  10 = Text Area
      '-----------------
       response.write "  <tr>" & vbcrlf
       response.write "      <td>" & vbcrlf

       'Determine if the fieldtype is a radio, select, or checkbox.
       if oQuestions("fieldtype") = "2" OR oQuestions("fieldtype") = "4" OR oQuestions("fieldtype") = "6" then
          if oQuestions("answerlist") <> "" then
             arrAnswers = split(oQuestions("answerlist"),chr(10))

             if oQuestions("fieldtype") = "4" then
                response.write "<select class=""formselect"">" & vbcrlf
             end if

       						for alist = 0 to ubound(arrAnswers)
                if oQuestions("fieldtype") = "4" then
         			   				response.write "  <option>" & arrAnswers(alist) & "</option>" & vbcrlf
                else
                   if oQuestions("fieldtype") = "2" then
                      response.write "  <input type=""radio"" name=""question" & oQuestions("questionid") & """ id=""question" & oQuestions("questionid") & """ class=""formradio"" />" & arrAnswers(alist) & "<br />" & vbcrlf
                   else
                      response.write "  <input type=""checkbox"" class=""formcheckbox"" />" & arrAnswers(alist) & "<br />" & vbcrlf
                   end if
                end if
             next

             if oQuestions("fieldtype") = "4" then
                response.write "</select>" & vbcrlf
             end if
          end if
       else
          if oQuestions("fieldtype") = "8" then
             response.write "  <input value="""" type=""text"" style=""width:300px;"" maxlength=""100"" />" & vbcrlf
          else
             response.write "  <textarea class=""formtextarea""></textarea>" & vbcrlf
          end if

       end if

       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf

     		'select case oQuestions("fieldtype")

      			'case "2"
        			'Build radio question

      						'arrAnswers = split(oQuestions("answerlist"),chr(10))
			
      						'for alist = 0 to ubound(arrAnswers)
      					 '  		response.write "  <tr><td><input type=""radio"" name=""question" & oQuestions("questionid") & """ id=""question" & oQuestions("questionid") & """ class=""formradio"" />" & arrAnswers(alist) & "</td></tr>" & vbcrlf
      						'next

      			'case "4"
      					'Build select question

      						'arrAnswers = split(oQuestions("answerlist"),chr(10))
			
      						'response.write "  <tr><td><select class=""formselect"">" & vbcrlf

      						'for alist = 0 to ubound(arrAnswers)
      			   '				response.write "  <option>" & arrAnswers(alist) & "</option>" & vbcrlf
      						'next

      						'response.write "  </select></td></tr>" & vbcrlf

      			'case "6"
      					'Build checkbox question

      						'arrAnswers = split(oQuestions("answerlist"),chr(10))
			
      						'for alist = 0 to ubound(arrAnswers)
      			   '				response.write "  <tr><td><input type=""checkbox"" class=""formcheckbox"" />" & arrAnswers(alist) & "</td></tr>" & vbcrlf
      						'next

      			'case "8"
      					'Build text question
      						'response.write "  <tr><td><input value="""" type=""text"" style=""width:300px;"" maxlength=""100""></td></tr>" & vbcrlf

      			'case "10"
      					'Build textarea question
      						'response.write "  <tr><td><textarea class=""formtextarea""></textarea></td></tr>" & vbcrlf

      			'case else

      	'end select

     		response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

     		oQuestions.MoveNext
    wend

  		response.write "</table>" & vbcrlf

	end if

	set oQuestions = nothing 

end sub

'----------------------------------------------------------------------------------------
Function IsRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
		  sReturnValue = " <font color=red>*</font> "
	Else
		  sReturnValue = ""
	End If

	IsRequired = sReturnValue
End Function

'----------------------------------------------------------------------------------------
Function IsDisplay(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "1" or sValue = "2" Then
		  sReturnValue = True
	Else
		  sReturnValue = False
	End If

	IsDisplay = sReturnValue
End Function

'----------------------------------------------------------------------------------------
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

'----------------------------------------------------------------------------------------
Sub subDisplayIssueLocation(sIssueMask, sIssueQues, sHideIssueLocAddInfo)
	
	If sIssueMask = "" OR IsNull(sIssueMask) Then
		  sIssueMask = "121111"
	End If

	If sIssueQues = "" OR isnull(sIssueQues) Then
 			sIssueQues = "Provide any additional information on problem location in the box below."
	End If
	
	response.write "<table border=""0"" style=""border: 1pt solid #c0c0c0; border-radius: 4px; padding: 4px;"">" & vbcrlf
	response.write "  <tr>"
 response.write "      <td colspan=""2"">" & vbcrlf
	response.write            sIssueDesc & vbcrlf
	response.write "          <p>" & vbcrlf
	response.write "          <table border=""0"">" & vbcrlf
	response.write "            <tr>" & vbcrlf
 response.write "                <td align=""right"">" & vbcrlf
	response.write                      IsRequired("2",1)  & " Address: </td>" & vbcrlf
    if lcl_orghasfeature_large_address_list then 
 response.write "                <td>"
 response.write "                    <input type=""text"" name=""residentstreetnumber"" value="""" size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
	response.write "                    <select name=""skip_address"">" & vbcrlf
	response.write "                      <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
 response.write "                    </select>&nbsp;" & vbcrlf
 response.write "                    <input type=""button"" value=""Validate Address"" />" & vbcrlf
 response.write "                </td>" & vbcrlf
    else
	response.write "                <td>" & vbcrlf
 response.write "                    <select style=""width:300;"">" & vbcrlf
	response.write "                      <option>Choose street from dropdown</option>" & vbcrlf
	response.write "                    </select>" & vbcrlf
 response.write "                </td>" & vbcrlf
    end if
	response.write "            </tr>" & vbcrlf

'Or Other Not Listed
	response.write "            <tr>" & vbcrlf
 response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td> - Or Other Not Listed - </td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
	response.write "            <tr>" & vbcrlf
 response.write "                <td>&nbsp;</td>" & vbcrlf
	response.write "                <td><input name=""ques_issue2"" type=""text"" size=""60"" /></td>" & vbcrlf
 response.write "            </tr>" & vbcrlf

'Unit
 response.write "            <tr>" & vbcrlf
 response.write "                <td align=""right"">Unit:&nbsp;</td>" & vbcrlf
	response.write "                <td><input type=""text"" name=""streetunit"" size=""8"" maxlength=""10"" /></td>" & vbcrlf
 response.write "            </tr>" & vbcrlf

'Additional Info

 response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "            <tr>" & vbcrlf
	response.write "                <td colspan=""2"">" & vbcrlf
 response.write "                    <input type=""button"" name=""setHideIssueLocAddInfo"" id=""setHideIssueLocAddInfo"" value=""Toggle Additional Info On/Off"" class=""button"" onclick=""document.frmIssueForm.task.value='ISSUELOCADDINFO';document.frmIssueForm.submit();"" />" & vbcrlf
 response.write "                    <small>(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_hideIssueLocAddInfo_label & "</font>)</small>" & vbcrlf
 response.write "                </td>" & vbcrlf
 response.write "            </tr>" & vbcrlf

 if not sHideIssueLocAddInfo then
   	response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
    response.write "            <tr>" & vbcrlf
    response.write "                <td>&nbsp;</td>" & vbcrlf
    response.write "                <td>" & sIssueQues & "<br />" & vbcrlf
    response.write "                    <textarea name=""ques_issue6"" class=""formtextarea""></textarea>" & vbcrlf
    response.write "                </td>" & vbcrlf
    response.write "            </tr>" & vbcrlf
 end if

	response.write "          </table>" & vbcrlf
	response.write "      </td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf

End Sub

'----------------------------------------------------------------------------------------
Function fnDrawInputType(iInputType)

				Select Case iInputType

					Case "1"
						' TEXT BOX ONLY
						sReturnValue = "<input type=""text"" />" & vbcrlf

					Case "2"
						' SELECT BOX ONLY
						sReturnValue = "<select><option>Please select street...</option></select>" & vbcrlf

					Case "3"
						' TEXT OR SELECT BOX
						sReturnValue = "<select><option>Please select street...</option></select> <br /> - Or Other Not Listed - <br /> "
						sReturnValue = sReturnValue & "<input type=""text"" />" & vbcrlf
						
					Case Else
						' DEFAULT TO TEXT BOX ONLY
						sReturnValue = "<input type=""text"" />" & vbcrlf

				End Select

		fnDrawInputType = sReturnValue

end function
%>
