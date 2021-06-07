<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
iorgid = session("orgid")
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

 if request.servervariables("REQUEST_METHOD") = "POST" then
 	  select Case request("TASK")

    		Case "new_question"
			     'ADD NEW QUESTION TO FORM
       		Call subAddQuestion(iorgid,iFormID)
    		Case Else
	  	    'DEFAULT ACTION
  	 end select
 end if

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

'Page Title, Buttons: Copy Form, Manage Form, and Return to Form List ---------
 response.write "<p>" & vbcrlf
 response.write "   <font class=""label"">Forms - Edit Form </font>" & vbcrlf
 'response.write "   <small>[<a class=""edit"" href=""copy_form.asp?task=copyme&iformid=" & iFormID & "&iorgid=" & session("orgid") & """>Copy This Form</a>]</small>" & vbcrlf
 response.write "   <input type=""button"" name=""copyForm"" id=""copyForm"" value=""Copy This Form"" class=""button"" onclick=""location.href='copy_form.asp?task=copyme&iformid=" & iFormID & "&iorgid=" & session("orgid") & "'"" />" & vbcrlf

 if blnCanManageActionAlerts then
    'response.write "   <small>[<a class=""edit"" href=""../action_line/edit_form.asp?task=name&control=" & iFormID & "&iorgid=" & session("orgid") & """>Manage This Form</a>]</small>" & vbcrlf
    response.write "   <input type=""button"" name=""manageForm"" id=""manageForm"" value=""Manage This Form"" class=""button"" onclick=""location.href='../action_line/edit_form.asp?task=name&control=" & iFormID & "&iorgid=" & session("orgid") & "'"" />" & vbcrlf
 end if

 'response.write "   <small>[<a class=""edit"" href=""list_forms.asp"">Return to Form List</a>]</small>" & vbcrlf
 response.write "   <input type=""button"" name=""returnToFormList"" id=""returnToFormList"" value=""Return to Form List"" class=""button"" onclick=""location.href='list_forms.asp'"" />" & vbcrlf
 response.write "   <hr size=""1"" width=""600px;"" style=""text-align:left; color:#000000;"" />" & vbcrlf
 response.write "</p>" & vbcrlf

'Resolved Status --------------------------------------------------------------
 response.write "<form name=""frmResolvedStatus"" id=""frmResolvedStatus"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
 response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""task"" id=""task_resolvedstatus"" value=""RESOLVEDSTATUS"" />" & vbcrlf
 response.write "  <input type=""hidden"" name=""RESOLVEDSTATUS"" id=""RESOLVEDSTATUS"" value=""" & sResolvedStatus & """ />" & vbcrlf
 'response.write "<small>[<a class=""edit"" href=""javascript:document.frmResolvedStatus.submit();"">Automatically set status to ""RESOLVED"" upon submission. On/Off</a>] " & vbcrlf
 response.write "<small>Automatically set status to ""RESOLVED"" upon submission. " & vbcrlf
 response.write "(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_resolvedstatus_label & "</font>)</small>" & vbcrlf
 response.write "<input type=""button"" name=""setResolvedStatus"" id=""setResolvedStatus"" value=""Toggle Resolved Status On/Off"" class=""button"" onclick=""document.frmResolvedStatus.submit();"" />" & vbcrlf
 response.write "</form>" & vbcrlf

'Edit Name --------------------------------------------------------------------
 response.write "<p>" & vbcrlf
 'response.write "   <small>[<a class=""edit"" href=""edit_form.asp?task=name&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Name</a>]</small> - " & vbcrlf
 response.write "   <font class=""subtitle"">" & sTitle & "</font>" & vbcrlf
 response.write "   <input type=""button"" name=""editName"" id=""editName"" value=""Edit Name"" class=""button"" onclick=""location.href='edit_form.asp?task=name&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
 response.write "</p>" & vbcrlf

'Emergency Note ---------------------------------------------------------------
 if blnEmergencyNote then
    lcl_emergencynote = "0"
 else
    lcl_emergencynote = "1"
 end if

 response.write "<form name=""frmEmergencyNote"" id=""frmEmergencyNote"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
 response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""task"" id=""task_emergencynote"" value=""EMERGENCYNOTE"" />" & vbcrlf
 response.write "  <input type=""hidden"" name=""emergencynote"" id=""emergencynote"" value=""" & lcl_emergencynote & """ />" & vbcrlf
 'response.write "<small>[<a class=""edit"" href=""edit_form.asp?task=enote&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Emergency Note</a>]</small>" & vbcrlf
 'response.write " - " & vbcrlf
 'response.write "<small>[<a class=""edit"" href=""#"" onClick=""document.frmEmergencyNote.submit();"">Toggle Emergency Notice On/Off </a>] " & vbcrlf
 'response.write "(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_emergencynote_label & "</font>)</small>" & vbcrlf
 response.write "<input type=""button"" name=""editEmergencyNote"" id=""editEmergencyNote"" value=""Edit Emergency Note"" class=""button"" onclick=""location.href='edit_form.asp?task=enote&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
 response.write "<input type=""button"" name=""editEmergencyNoteToggle"" id=""editEmergencyNoteToggle"" value=""Toggle Emergency Notice On/Off"" class=""button"" onclick=""document.frmEmergencyNote.submit();"" /> " & vbcrlf
 response.write "<small>(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_emergencynote_label & "</font>)</small>" & vbcrlf

 if blnEmergencyNote then
    response.write "<div class=""warning"">" & sEmergencyText & "</div>" & vbcrlf
 end if

 response.write "</form>" & vbcrlf
 'PUBLICALLY SEARCHABLE--------------------------------------------------------
 if iorgid = "5" or iorgid = "125" then
 if blnPublicSearchNote then
    lcl_publicsearchnote = "0"
 else
    lcl_publicsearchnote = "1"
 end if
 response.write "<form name=""frmPublicallySearchable"" id=""frmPublicallySearchable"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
 response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
 response.write "  <input type=""hidden"" name=""task"" id=""task_publicsearch"" value=""PUBLICSEARCH"" />" & vbcrlf
 response.write "  <input type=""hidden"" name=""publicsearchnote"" id=""publicsearchnote"" value=""" & lcl_publicsearchnote & """ />" & vbcrlf
 response.write "<input type=""button"" name=""editPublicallySearchableToggle"" id=""editPublicallySearchableToggle"" value=""Toggle Making Requests Publically Searchable On/Off"" class=""button"" onclick=""document.frmPublicallySearchable.submit();"" /> " & vbcrlf
 response.write "<small>(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_publicsearchnote_label & "</font>)</small>" & vbcrlf
 response.write "</form>" & vbcrlf
 end if


 response.write "<p>" & vbcrlf
 response.write "<div class=""group"">" & vbcrlf
 response.write "  <div class=""orgadminboxf"">" & vbcrlf


'BEGIN: Intro Information -----------------------------------------------------
 response.write "<p>" & vbcrlf
 response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""background-color:#e0e0e0"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td>" & vbcrlf
 'response.write "          <small>[<a class=""edit"" href=""edit_form.asp?task=intro&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Intro</a>]</small>" & vbcrlf
 response.write "          <input type=""button"" name=""editIntro"" id=""editIntro"" value=""Edit Intro"" class=""button"" onclick=""location.href='edit_form.asp?task=intro&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
 response.write "      </td>" & vbcrlf

	'if lcl_orghasfeature_requestmergeforms OR lcl_userhaspermission_form_letters then
	if lcl_orghasfeature_requestmergeforms AND lcl_userhaspermission_requestmergeforms then
    response.write "      <td align=""right"">" & vbcrlf
    response.write "          <input type=""button"" name=""viewPDF_FieldNames"" id=""viewPDF_FieldNames"" value=""PDF Field Names"" class=""button"" onclick=""openCustomReports('PDFFIELDNAMES');"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
 end if

 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "</p>" & vbcrlf

 if sIntroText <> "" then
 			response.write "<p>" & sIntroText & "</p>" & vbcrlf
 else
 			response.write "<p>- <i> Introduction text is currently blank </i> -</p>" & vbcrlf
 end if

'BEGIN: Contact Information ---------------------------------------------------
 response.write "<p>" & vbcrlf
 'response.write "   <small>[<a class=""edit"" href=""edit_form.asp?task=contact&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Contact</a>]</small>" & vbcrlf
 response.write "   <input type=""button"" name=""editContact"" id=""editContact"" value=""Edit Contact"" class=""button"" onclick=""location.href='edit_form.asp?task=contact&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
 response.write "</p>" & vbcrlf
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

    response.write "<form name=""frmIssueForm"" id=""frmIssueForm"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
    response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""task"" id=""task_issue"" value=""ENABLEISSUE"" />" & vbcrlf
    response.write "  <input type=""hidden"" name=""ENABLEISSUE"" id=""ENABLEISSUE"" value=""" & lcl_enableissue & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""ISSUELOCADDINFO"" id=""ISSUELOCADDINFO"" value=""" & lcl_hideIssueLocAddInfo & """ />" & vbcrlf
    'response.write "<input type=""button"" name=""issueLocationToggle"" id=""issueLocationToggle"" value=""Toggle Issue Location On/Off"" class=""button"" onclick=""document.frmIssueForm.task.value='ENABLEISSUE';document.frmIssueForm.submit();"" />" & vbcrlf
    response.write "<input type=""button"" name=""issueLocationToggle"" id=""issueLocationToggle"" value=""Toggle Issue Location On/Off"" class=""button"" onclick=""enableDisableFormOptions('issue', 'ENABLEISSUE', 'frmIssueForm');"" />" & vbcrlf
    response.write "<small>(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_issuedisplay_label & "</font>)</small>" & vbcrlf

    if blnIssueDisplay then
       response.write "<p>" & vbcrlf
       'response.write "   <small>[<a class=""edit"" href=""edit_form.asp?task=issuename&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Name\Description</a>]</small> - " & vbcrlf
       response.write "   <strong><u>" & sIssueName & "</u></strong>" & vbcrlf
       response.write "   <input type=""button"" name=""editNameDesc"" id=""editNameDesc"" value=""Edit Name/Description"" class=""button"" onclick=""location.href='edit_form.asp?task=issuename&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
       response.write "   <br /><br />" & vbcrlf

       subDisplayIssueLocation sIssueMask, sIssueQues, sHideIssueLocAddInfo

       response.write "</p>" & vbcrlf
    end if

    response.write "</form>" & vbcrlf
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

    'response.write "<fieldset class=""fieldset"">" & vbcrlf
    'response.write "  <legend>Mobile Options</legend>" & vbcrlf
    response.write "<form name=""frmMobileOptionsForm"" id=""frmMobileOptionsForm"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
    response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""task"" id=""task_mobileoptions"" value=""ENABLEMOBILEOPTIONSTAKEPIC"" />" & vbcrlf
    response.write "  <input type=""hidden"" name=""ENABLEMOBILEOPTIONSTAKEPIC"" id=""ENABLEMOBILEOPTIONSTAKEPIC"" value=""" & lcl_enableMobileOptionsTakePic & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""ENABLEMOBILEOPTIONSGEOLOC"" id=""ENABLEMOBILEOPTIONSGEOLOC"" value=""" & lcl_enableMobileOptionsGeoLoc & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""ENABLESHOWMAPINPUT"" id=""ENABLESHOWMAPINPUT"" value=""" & lcl_enableShowMapInput & """ />" & vbcrlf
    'response.write "<div style=""text-align: center;"">" & vbcrlf
    response.write "<input type=""button"" name=""mobileOptionsToggle_takepic"" id=""mobileOptionsToggle_takepic"" value=""Toggle File Uploads - On/Off"" class=""button"" onclick=""enableDisableFormOptions('mobileoptions', 'ENABLEMOBILEOPTIONSTAKEPIC', 'frmMobileOptionsForm');"" />" & vbcrlf
    response.write "<small>(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" &  lcl_mobileOptionsTakePicDisplay_label & "</font>)</small>" & vbcrlf
    response.write "&nbsp;&nbsp;&nbsp;" & vbcrlf
    'if iorgid = 37 then
    response.write "<br /><br /><input type=""button"" name=""showMapInputToggle"" id=""showMapInputToggle"" value=""Toggle Map - On/Off"" class=""button"" onclick=""enableDisableFormOptions('mobileoptions', 'ENABLESHOWMAPINPUT', 'frmMobileOptionsForm');"" />" & vbcrlf
    response.write "<small>(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" &  lcl_showmapinput_label & "</font>)</small>" & vbcrlf
    'end if
    'response.write "</div>" & vbcrlf
    response.write "</form>" & vbcrlf
    'response.write "</fieldset>" & vbcrlf
 end if
'END: Mobile Fields -----------------------------------------------------------


'BEGIN: Form Field Information ------------------------------------------------
 response.write "<p>" & vbrlf
 'response.write "   <small>[<a class=""edit"" href=""add_field.asp?iformid=" & iFormID & "&iorgid=" & iorgid & """>Add New Question</a>]</small>" & vbcrlf
 response.write "   <input type=""button"" name=""addNewQuestion"" id=""addNewQuestion"" value=""Add New Question"" class=""button"" onclick=""location.href='add_field.asp?iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
 response.write "</p>" & vbcrlf
 response.write "<p>" & vbrlf

 subDisplayQuestions iFormID,0

if orghasfeature("actionline_formcreator_appoptions") then
'if request.cookies("user")("userid") = "1710" then
 response.write "</p>" & vbcrlf
    response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend>App Options</legend>" & vbcrlf
    response.write "<form name=""frmAppOptionsForm"" id=""frmAppOptionsForm"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
    response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""task"" value=""SAVEAPPSETTINGS"" />" & vbcrlf
    if appChecked then
	    appChecked = " checked"
    else
	    appChecked = ""
    end if
    response.write "<input type=""checkbox"" " & appChecked & " name=""formobile"" value=""ON"" /> App Enabled<br />"
    response.write "Form Name for App: <input type=""text"" name=""mobilename"" value=""" & formName & """ style=""width:300px"" maxlength=""100"" /><br />"
    response.write "Help Text: <input type=""text"" name=""mobilehelptext"" value=""" & helpText & """ style=""width:500px"" maxlength=""500"" /><br /><br />"
    response.write "<input type=""submit"" value=""Save App Settings"" />"
    response.write "</form>" & vbcrlf
    response.write "</fieldset><br /><br />" & vbcrlf
   end if
'if request.cookies("user")("userid") = "1710" then
    response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend>Custom Confirmation Page</legend>" & vbcrlf
    response.write "  <p>If you enter a URL below we'll direct people who complete this form to that page instead of the standard confirmation page.</p>" & vbcrlf
    response.write "<form name=""frmAppOptionsForm"" id=""frmAppOptionsForm"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
    response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""task"" value=""SAVECUSTOMURL"" />" & vbcrlf
    response.write "URL: <input type=""text"" name=""redirectURL"" value=""" & lcl_redirectURL & """ style=""width:500px"" maxlength=""500"" /><br /><br />"
    response.write "<input type=""submit"" value=""Save Custom URL"" />"
    response.write "</form>" & vbcrlf
    response.write "</fieldset><br /><br />" & vbcrlf
 'end if

'BEGIN: Administrative Questions ----------------------------------------------
 if lcl_userhaspermission_internalfields then
    response.write "<div style=""background-color:#336699;font-weight:bold;color:#ffffff;padding:10px;"">" & vbcrlf
    response.write "  Internal Use Only - Administrative Fields" & vbcrlf
    response.write "</div>" & vbcrlf
    response.write "<p>" & vbcrlf
    'response.write "   <small>[<a class=""edit"" href=""add_field.asp?isinternal=1&iformid=" & iFormID & "&iorgid=" & iorgid & """>Add New Internal Question</a>]</small>" & vbcrlf
    response.write "   <input type=""button"" name=""addNewInternalQuestion"" id=""addNewInternalQuestion"" value=""Add New Internal Question"" class=""button"" onclick=""location.href='add_field.asp?isinternal=1&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
    response.write "</p>" & vbcrlf
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

    response.write "<form name=""frmFeeForm"" id=""frmFeeForm"" action=""edit_form.asp"" method=""POST"">" & vbcrlf
    response.write "  <input type=""hidden"" name=""iformid"" id=""iformid"" value=""" & iformid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""iorgid"" id=""iorgid"" value=""" & iorgid & """ />" & vbcrlf
    response.write "  <input type=""hidden"" name=""task"" id=""task_fee"" value=""ENABLEFEE"" />" & vbcrlf
    response.write "  <input type=""hidden"" name=""ENABLEFEE"" id=""ENABLEFEE""  value=""" & lcl_feevalue & """ />" & vbcrlf
    'response.write "<small>[<a class=""edit"" href=""#"" onClick=""document.frmFeeForm.submit();"">Toggle Fees On\Off </a>] (<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_feedisplay_label & "</font>)</small>" & vbcrlf
    response.write "<input type=""button"" name=""feeToggle"" id=""feeToggle"" value=""Toggle Fees On/Off"" class=""button"" onclick=""document.frmFeeForm.submit();"" />" & vbcrlf
    response.write "(<font style=""font-family:arial;font-size:10px;font-weight:bold;"">" & lcl_feedisplay_label & "</font>)</small>" & vbcrlf
    response.write "</form>" & vbcrlf
 end if

'BEGIN: Ending Notes ----------------------------------------------------------
 response.write "<p><font color=""#ff0000"">*</font><strong><i>Required Field</i></strong></p>" & vbcrlf
 'response.write "<p><small>[<a class=""edit"" href=""edit_form.asp?task=footer&iformid=" & iFormID & "&iorgid=" & iorgid & """>Edit Footer</a>]</small></p>" & vbcrlf
 response.write "<p><input type=""button"" name=""editFooter"" id=""editFooter"" value=""Edit Footer"" class=""button"" onclick=""location.href='edit_form.asp?task=footer&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" /></p>" & vbcrlf

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
     response.write "  <input type=""button"" name=""customFOILButton"" id=""customFOILButton"" value=""Edit FOIL Email Edits"" onclick=""location.href='edit_form.asp?task=CUSTOMFOILEMAILEDITS&iformid=" & iFormID & "&iorgid=" & iorgid & "'"" />" & vbcrlf
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

     		response.write "  <tr>" & vbcrlf
       response.write "      <td>" & vbcrlf
       'response.write "          <small>" & vbcrlf
       'response.write "          [<a class=""edit"" href=""edit_field.asp?iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Edit</a>]" & vbcrlf
       'response.write "          [<a class=""edit"" href=# onclick=""confirm_delete('" & oQuestions("questionid") & "','" & JSsafe(Left(oQuestions("prompt"),25)) & "...');"">Delete</a>]" & vbcrlf
       'response.write "          [<a class=""edit"" href=""order_field.asp?direction=UP&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move Up</a>]" & vbcrlf
       'response.write "          [<a class=""edit"" href=""order_field.asp?direction=down&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move Down</a>]" & vbcrlf
       'response.write "          [<a class=""edit"" href=""order_field.asp?direction=top&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move to Top</a>]" & vbcrlf
       'response.write "          [<a class=""edit"" href=""order_field.asp?direction=bottom&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & """>Move to Bottom</a>]" & vbcrlf
       'response.write "          </small>" & vbcrlf
       response.write "          <input type=""button"" name=""editField"   & oQuestions("questionid") & """ id=""editField"   & oQuestions("questionid") & """ value=""Edit"" class=""button"" onclick=""location.href='edit_field.asp?iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & "'"" />" & vbcrlf
       response.write "          <input type=""button"" name=""deleteField" & oQuestions("questionid") & """ id=""deleteField" & oQuestions("questionid") & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oQuestions("questionid") & "','" & JSsafe(Left(oQuestions("prompt"),25)) & "...')"" />" & vbcrlf
       response.write "          <input type=""button"" name=""moveUp"      & oQuestions("questionid") & """ id=""moveUp"      & oQuestions("questionid") & """ value=""Move Up"" class=""button"" onclick=""location.href='order_field.asp?direction=UP&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & "'"" />" & vbcrlf
       response.write "          <input type=""button"" name=""moveDown"    & oQuestions("questionid") & """ id=""moveDown"    & oQuestions("questionid") & """ value=""Move Down"" class=""button"" onclick=""location.href='order_field.asp?direction=down&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & "'"" />" & vbcrlf
       response.write "          <input type=""button"" name=""moveTop"     & oQuestions("questionid") & """ id=""moveTop"     & oQuestions("questionid") & """ value=""Move to Top"" class=""button"" onclick=""location.href='order_field.asp?direction=top&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & "'"" />" & vbcrlf
       response.write "          <input type=""button"" name=""moveBottom"  & oQuestions("questionid") & """ id=""moveBottom"  & oQuestions("questionid") & """ value=""Move to Bottom"" class=""button"" onclick=""location.href='order_field.asp?direction=bottom&iformid=" & iFormID & "&iorgid=" & iorgid & "&ifieldid=" & oQuestions("questionid") & "'"" />" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
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
