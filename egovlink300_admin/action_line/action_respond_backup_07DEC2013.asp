<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: action_respond.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module is the New Action Line Requests.
'
' MODIFICATION HISTORY
' 2.0	10/16/2006	 Steve Loar  - Security, Header and Nav changed
' 2.1	04/16/2007	 Steve Loar  - Changes to email for To and From addresses
' 2.2  08/08/07  David Boyer - Added Sub-Status
' 2.3  08/29/07  David Boyer - Added Code-Sections
' 2.4  02/22/08  David Boyer - Added Export to PDF
' 2.5  07/09/08  David Boyer - Added Send Notification section
' 2.6  08/29/08  David Boyer - Added Department fields to update
' 2.7  10/09/08  David Boyer - Fixed permissions around "edit" links
' 2.8  11/12/08  David Boyer - Added "View PDF" button
' 2.9  11/19/08  David Boyer - Added Email Reminders
' 3.0  12/19/08  David Boyer - Converted "View PDF" to "PDFs" section
' 3.1  03/26/09  David Boyer - Fixed opening form letters with "&" in Additional Text
' 3.2  05/29/09  David Boyer - Added check to see if "Additional Information" textarea is displayed or not
' 3.3  06/17/09  David Boyer - Added "e=Y" to (action_respond.asp) urls in emails
' 3.4  07/15/09  David Boyer - Fixed the "unassigned department not available to user" removing the department from the request
' 3.5  07/15/09  David Boyer - Added "push content" to FAQs/Rumor Mill
' 3.6  08/03/09  David Boyer - Added "Delegate check" when sending emails.
' 3.7  08/10/09  David Boyer - Added "secure attachments"
' 3.8  08/10/09  David Boyer - Added "hiding the Activity Log"
' 3.9  01/29/10  David Boyer - Added "Add a Link" to "Note to Citizen" field.
' 4.0  02/01/10  David Boyer - Added log entry when sending an email to citizen.
' 4.1  08/18/10  David Boyer - Added "push content" to Community Calendar
' 4.2  09/27/11  David Boyer - Added "Linked Requests"
' 4.3	05/06/2013	Steve Loar - Pointed some PDF outputs to new viewXMLPDF.asp script
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 Dim datSubmitDate, sSubmitName, sUserEmail, strLocalTime, bFormHasIssueLocation
 strLocalTime = ""

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"requests") then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for org features
 lcl_orghasfeature_responsetimereporting                        = orghasfeature("responsetimereporting")
 lcl_orghasfeature_action_line_substatus                        = orghasfeature("action_line_substatus")
 lcl_orghasfeature_requestmergeforms                            = orghasfeature("requestmergeforms")
 lcl_orghasfeature_modify_actionline_department                 = orghasfeature("modify_actionline_department")
 lcl_orghasfeature_send_notification                            = orghasfeature("send_notification")
 lcl_orghasfeature_action_line_code_sections                    = orghasfeature("action_line_code_sections")
 lcl_orghasfeature_form_letters                                 = orghasfeature("form letters")
 lcl_orghasfeature_hide_email_actionline                        = orghasfeature("hide email actionline")
 lcl_orghasfeature_parcel_id                                    = orghasfeature("parcel_id")
 lcl_orghasfeature_actionline_emailreminders                    = orghasfeature("actionline_emailreminders")
 lcl_orghasfeature_activitylog_workorder                        = orghasfeature("activitylog_workorder")
 lcl_orghasfeature_faq                                          = orghasfeature("faq")
 lcl_orghasfeature_manage_faq                                   = orghasfeature("manage faq")
 lcl_orghasfeature_rumormill                                    = orghasfeature("rumormill")
 lcl_orghasfeature_rumormill_manage                             = orghasfeature("rumormill_manage")
 lcl_orghasfeature_calendar                                     = orghasfeature("calendar")
 lcl_orghasfeature_create_events                                = orghasfeature("create events")
 lcl_orghasfeature_pushcontent_faqs                             = orghasfeature("pushcontent_faqs")
 lcl_orghasfeature_pushcontent_rumormill                        = orghasfeature("pushcontent_rumormill")
 lcl_orghasfeature_pushcontent_communitycalendar                = orghasfeature("pushcontent_communitycalendar")
 lcl_orghasfeature_actionline_hide_requestlog                   = orghasfeature("actionline_hide_requestlog")
 lcl_orghasfeature_actionline_maintain_duedate                  = orghasfeature("actionline_maintain_duedate")
 lcl_orghasfeature_fileupload                                   = orghasfeature("fileupload")
 lcl_orghasfeature_actionline_secure_attachments                = orghasfeature("actionline_secure_attachments")
 lcl_orghasfeature_actionline_display_attachments_to_public     = orghasfeature("actionline_display_attachments_to_public")
 lcl_orghasfeature_actionline_linkedrequests                    = orghasfeature("actionline_linkedrequests")
 lcl_orghasfeature_actionline_reuse_completedate                = orghasfeature("actionline_reuse_completedate")

'Check for user permissions
 lcl_userhaspermission_action_line_substatus                    = userhaspermission(session("userid"),"action_line_substatus")
 lcl_userhaspermission_action_line_code_sections                = userhaspermission(session("userid"),"action_line_code_sections")
 lcl_userhaspermission_requestedit                              = userhaspermission(session("userid"),"requestedit")
 lcl_userhaspermission_internalfields                           = userhaspermission(session("userid"),"internalfields")
 lcl_userhaspermission_bzfees                                   = userhaspermission(session("userid"),"bzfees")
 lcl_userhaspermission_fileupload                               = userhaspermission(session("userid"),"fileupload")
 lcl_userhaspermission_can_close_requests                       = userhaspermission(session("userid"),"can close requests")
 lcl_userhaspermission_actionline_delete                        = userhaspermission(session("userid"),"actionline delete")
 lcl_userhaspermission_requestmergeforms                        = userhaspermission(session("userid"),"requestmergeforms")
 lcl_userhaspermission_actionline_emailreminders                = userhaspermission(session("userid"),"actionline_emailreminders")
 lcl_userhaspermission_faq                                      = userhaspermission(session("userid"),"faq")
 lcl_userhaspermission_manage_faq                               = userhaspermission(session("userid"),"manage faq")
 lcl_userhaspermission_rumormill                                = userhaspermission(session("userid"),"rumormill")
 lcl_userhaspermission_rumormill_manage                         = userhaspermission(session("userid"),"rumormill_manage")
 lcl_userhaspermission_calendar                                 = userhaspermission(session("userid"),"calendar")
 lcl_userhaspermission_create_events                            = userhaspermission(session("userid"),"create events")
 lcl_userhaspermission_pushcontent_faqs                         = userhaspermission(session("userid"),"pushcontent_faqs")
 lcl_userhaspermission_pushcontent_rumormill                    = userhaspermission(session("userid"),"pushcontent_rumormill")
 lcl_userhaspermission_pushcontent_communitycalendar            = userhaspermission(session("userid"),"pushcontent_communitycalendar")
 lcl_userhaspermission_actionline_secure_attachments            = userhaspermission(session("userid"),"actionline_secure_attachments")
 lcl_userhaspermission_actionline_hide_requestlog               = userhaspermission(session("userid"),"actionline_hide_requestlog")
 lcl_userhaspermission_actionline_maintain_duedate              = userhaspermission(session("userid"),"actionline_maintain_duedate")
 lcl_userhaspermission_actionline_display_attachments_to_public = userhaspermission(session("userid"),"actionline_display_attachments_to_public")

'Determine if we are hiding the Activity Log
 if lcl_orghasfeature_actionline_hide_requestlog AND lcl_userhaspermission_actionline_hide_requestlog then
    lcl_hide_activitylog = "Y"
 else
    lcl_hide_activitylog = "N"
 end if

'Set to use new permission levels
 blnCanEditAllActionItems  = False
 blnCanEditOwnActionItems  = False
 blnCanEditDeptActionItems = False
 blnCanEdit                = False   'This is the flag that will allow editing

 iPermissionLevelId = GetUserPermissionLevel(session("userid"),"requests")

 if clng(iPermissionLevelId) > 0 then
    sPermissionLevel = GetPermissionLevelName(iPermissionLevelId)
 else
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Override the flags for what they get
'Note that "View All" can edit nothing
 select Case sPermissionLevel
	  Case "View All - Edit All"
   		blnCanEditAllActionItems  = True 
   		blnCanEditDeptActionItems = True
   		blnCanEditOwnActionItems  = True
  	Case "View All - Edit Dept"
   		blnCanEditDeptActionItems = True
   		blnCanEditOwnActionItems  = True 
  	Case "View All - Edit Own"
   		blnCanEditOwnActionItems  = True 
  	Case "View Dept - Edit Dept"
   		blnCanEditDeptActionItems = True
   		blnCanEditOwnActionItems  = True 
  	Case "View Dept - Edit Own"
   		blnCanEditOwnActionItems  = True 
  	Case "View Own - Edit Own"
   		blnCanEditOwnActionItems  = True 
 end select

'If the use has clicked on the "UPDATE ACTION REQUEST" button
 if request.ServerVariables("REQUEST_METHOD") = "POST" then
    blnUpdate = False
    blnNotify = False

   'Determine which button the user clicked.
    if request("sAction") = "SEND NOTIFICATION" then
       Send_Notification(request("TrackID"))
       blnNotify = True
    else
       Update_Action(request("TrackID"))
       blnUpdate = True
    end if
   
    iTrackID = request("TrackID")
 end if

'Get Information for this request
 If iTrackID = "" Then
    iTrackID = request("control") 
 End If

'If the iTrackID STILL has no value then it tells us that the user 
'has attempted to open this screen without an "action_autoid"
'Since this is the case then simply send them back to the main Action Line => Requests screen.
 if iTrackID = "" then
    response.redirect("action_line_list.asp")
 end if

'---------------------------------------------------
'ORIGINAL QUERY
'sSQL = "SELECT *, (FirstName + ' ' + LastName) as EmployeeSubmitName, F.DeptID FROM egov_actionline_requests "
'sSQL = sSQL & " left outer join users on egov_actionline_requests.employeesubmitid=users.userid "
'sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egov_actionline_requests.userid = egov_users.userid "
'sSQL = sSQL & " LEFT OUTER JOIN egov_action_request_forms AS F ON egov_actionline_requests.category_id = F.action_form_id "
'sSQL = sSQL & " where action_autoid = " & iTrackID
'---------------------------------------------------

 sSQL = "SELECT *, (users.FirstName + ' ' + users.LastName) as EmployeeSubmitName, "
 sSQL = sSQL & " isnull(ear.groupid,F.DeptID) as DeptID, "
 sSQL = sSQL & " ears.action_status_id as sub_status_id, "
 sSQL = sSQL & " ears.status_name AS sub_status_name, "
 sSQL = sSQL & " isnull(ear.public_actionline_pdf,F.public_actionline_pdf) as public_actionline_pdf, "
 sSQL = sSQL & " f.hideIssueLocAddInfo "
 sSQL = sSQL & " FROM egov_actionline_requests AS ear "
 sSQL = sSQL &      " LEFT OUTER JOIN users on ear.employeesubmitid = users.userid "
 sSQL = sSQL &      " LEFT OUTER JOIN egov_users ON ear.userid = egov_users.userid "
 sSQL = sSQL &      " LEFT OUTER JOIN egov_action_request_forms AS F ON ear.category_id = F.action_form_id "
 sSQL = sSQL &      " LEFT OUTER JOIN egov_actionline_requests_statuses AS ears "
 sSQL = sSQL &                      "ON ear.sub_status_id = ears.action_status_id "
 'sSQL = sSQL &                      "AND upper(ear.status) = upper(ears.parent_status) "
 sSQL = sSQL & " WHERE ear.action_autoid = " & iTrackID
 sSQL = sSQL & " AND ear.orgid = " & session("orgid")

 set oRequest = Server.CreateObject("ADODB.Recordset")
 oRequest.Open sSQL, Application("DSN"), 3, 1

'CHECK FOR INFORMATION
if not oRequest.eof then

  'set the can edit flag based on edit permission and what the request is
   If blnCanEditAllActionItems Then 
      blnCanEdit = True
   Else
      If blnCanEditDeptActionItems Then 
        'isUserInDept is in common.asp
         If isUserInDept( Session("userid"), oRequest("DeptID") ) Then 
            blnCanEdit = True
         End If 
      End If 
      If (Not blnCanEdit) And blnCanEditOwnActionItems Then 
          if oRequest("assignedemployeeid") <> "" then
             if CLng(oRequest("assignedemployeeid")) = CLng(session("userid")) then
                blnCanEdit = True
             end if
          end if
      End If 
   End If

  'REQUEST FOUND GET INFORMATION	
   blnFound                  = True
   bFormHasIssueLocation     = oRequest("action_form_display_issue")
   sTitle                    = oRequest("category_title")
   iFormID                   = oRequest("category_id")
   sStatus                   = oRequest("status")
   sSubStatus                = oRequest("sub_status_name")
   sSubStatusID              = oRequest("sub_status_id")
   datSubmitDate             = oRequest("submit_date")
   sComment                  = oRequest("comment")
   sTheuserid                = oRequest("userid")
   iemployeeid               = oRequest("assignedemployeeid")
   blnFeeDisplay             = oRequest("action_form_display_fees")
   isubmitid                 = oRequest("employeesubmitid")
   icontactmethodid          = oRequest("contactmethodid")
   sDeptID                   = oRequest("deptid")
   sIssueName                = oRequest("issuelocationname")
   sPublicActionLinePDF      = oRequest("public_actionline_pdf")
   sHideIssueLocAddInfo      = oRequest("hideIssueLocAddInfo")
   sCompleteDate             = oRequest("complete_date")
   sSubmittedByRemoteAddress = oRequest("submittedby_remoteaddress")

   If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
      sIssueName = "Issue/Problem Location:"
   End If

   sIssueDesc = oRequest("issuelocationdesc")
   if IsNull(sIssueDesc) then
      sIssueDesc = "Please select the closest street number/streetname of problem location from list or select ""*not on list"". "
      sIssueDesc = sIssueDesc & "Provide any additional information on problem location in the box below."
   end if

  'Get Employee or Citizen that submitted the request
   if isubmitid < 0 OR IsNull(isubmitid) OR isubmitid = "" then
     'Display the Citizen Name as the Submitter
      sSubmitName = oRequest("userfname") & " " & oRequest("userlname") & " (Citizen)"
   else
     'Display the Employee Name as the Submitter
      sSubmitName =  oRequest("EmployeeSubmitName") & " (Admin Employee)"
   end if

   if datSubmitDate <> "" then
      lngTrackingNumber = iTrackID  & replace(FormatDateTime(cdate(datSubmitDate),4),":","")
   else
      lngTrackingNumber = "000000000"
   end if

   if oRequest("due_date") <> "" then
      sDueDate = FormatDateTime(oRequest("due_date"), vbshortdate)
   else
      sDueDate = ""
   end if

else
  'REQUEST NOT FOUND
   blnFound = False
end if

oRequest.close
set oRequest = nothing

'If the request is not found then redirect the user to the main action line screen.
 if not blnFound then
    response.redirect "action_line_list.asp" & lcl_return_str
 end if

'Record view of request if org has "Response Time Reporting" feature turned-on.
 if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    if lcl_orghasfeature_responsetimereporting then
       AddCommentTaskComment "Request viewed by " & session("FullName") & ".", null, Ucase(sStatus), iTrackID, Session("userid"), session("orgid"), sSubStatusID, "", ""
    end If
 end If

'The activity log tracking for the PDF and Work Order viewing/printing buttons needs to be done separately to work properly.
'*** This was previously done in viewPDF.asp but when done there double-records are inserted into the activity log. ***
'*** The reason is because the data is "streamed" back and the page is refreshed. ***
'*** Updating the activity log here gets around that issue. ***
 if request("actlog") <> "" then
    updateActivityLog_PDFs request("actlog"),request("pdfid"),sPublicActionLinePDF,sStatus,iTrackID,sSubStatusID
 end if

'Determine if the "push content" option is displayed.
 lcl_pushcontent          = false
 lcl_pushcontent_dropdown = "&nbsp;"
 lcl_pushcontent_options  = ""

 if (lcl_orghasfeature_faq              AND _
     lcl_userhaspermission_faq          AND _
     lcl_orghasfeature_manage_faq       AND _
     lcl_userhaspermission_manage_faq   AND _
     lcl_orghasfeature_pushcontent_faqs AND _
     lcl_userhaspermission_pushcontent_faqs) _

 OR (lcl_orghasfeature_rumormill             AND _
     lcl_userhaspermission_rumormill         AND _
     lcl_orghasfeature_pushcontent_rumormill AND _
     lcl_userhaspermission_pushcontent_rumormill) _

 OR (lcl_orghasfeature_calendar                      AND _
     lcl_orghasfeature_create_events                 AND _
     lcl_orghasfeature_pushcontent_communitycalendar AND _
     lcl_userhaspermission_calendar                  AND _
     lcl_userhaspermission_create_events             AND _
     lcl_userhaspermission_pushcontent_communitycalendar) _
 then
     lcl_pushcontent = true
 end if

 if lcl_pushcontent then
   'FAQs
    if lcl_orghasfeature_pushcontent_faqs AND lcl_userhaspermission_pushcontent_faqs then
       lcl_pushcontent_options = lcl_pushcontent_options & "  <option value=""FAQ"">FAQs</option>" & vbcrlf
    end if

   'RUMOR MILL
    if lcl_orghasfeature_pushcontent_rumormill AND lcl_userhaspermission_pushcontent_rumormill then
       lcl_pushcontent_options = lcl_pushcontent_options & "  <option value=""RUMORMILL"">Rumor Mill</option>" & vbcrlf
    end if

   'COMMUNITY CALENDAR
    if lcl_orghasfeature_calendar                          AND _
       lcl_orghasfeature_create_events                     AND _
       lcl_orghasfeature_pushcontent_communitycalendar     AND _
       lcl_userhaspermission_calendar                      AND _
       lcl_userhaspermission_create_events                 AND _
       lcl_userhaspermission_pushcontent_communitycalendar AND _
       checkIsPushForm(session("orgid"), iFormID) _
    then
       lcl_pushcontent_options = lcl_pushcontent_options & "  <option value=""COMMUNITYCALENDAR"">Community Calendar</option>" & vbcrlf
    end if

    if lcl_pushcontent_options <> "" then
       lcl_pushcontent_dropdown = lcl_pushcontent_dropdown & "<span class=""darkRedText"">Push request to </span>" & vbcrlf
       lcl_pushcontent_dropdown = lcl_pushcontent_dropdown & "  <select name=""pushFeature"" id=""pushFeature"">" & vbcrlf
       lcl_pushcontent_dropdown = lcl_pushcontent_dropdown & lcl_pushcontent_options
       lcl_pushcontent_dropdown = lcl_pushcontent_dropdown & "  </select>" & vbcrlf
       lcl_pushcontent_dropdown = lcl_pushcontent_dropdown & "  <input type=""button"" name=""pushContentButton"" id=""pushContentButton"" value=""Push >>"" class=""button"" onclick=""pushContent('" & iTrackID & "');"" />&nbsp;&nbsp;&nbsp;" & vbcrlf
    end if
  end if
%>
<html>
<head>
  <title><%=langBSActionLine%></title>
<!-- This metadata is for setting the priority and importance for CDO mail messages -->
<!--
METADATA
TYPE="typelib"
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
NAME="CDO for Windows 2000 Library"
-->
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <style type="text/css">
    #requestTitle {
       padding: 5px 5px 0px 10px;
    }

    #screenMsg {
       color:       #ff0000;
       font-size:   10pt;
       font-weight: bold;
    }

    #linkedRequests {
       /* border:  1pt solid #e0e0e0; */
    }

    #linkedRequests td,
    #linkedRequests th {
       text-align: left;
       padding:    2px 2px;
    }

    #linkedRequestAddButton {
       margin-bottom: 5px;
    }

    #row_linkedRequestsButtons {
       text-align:    right;
       padding-right: 5px;
    }

    #log {
       margin-top:       5px;
       border:           solid 1px #000000;
       background-color: #ffffff;
    }

    #logHeaderRow {
       border-bottom: solid 1px #000000;
    }

    #user_expand,
    .user_expand {
       cursor:          pointer;
       font-weight:     bold;
       text-decoration: underline;
    }

    .redText {
       color: #ff0000;
    }

    .darkRedText {
       color: #800000;
    }

    .divSection {
       padding:          5px;
       margin-top:       5px;
       border:           solid 1px #000000;
       background-color: #e0e0e0;
       border-radius:    5px 5px 5px 5px;
    }

    .formInformation_noComments {
       padding:          5px;
       margin-top:       5px;
       border-top:       solid 1px #000000;
       border-bottom:    solid 1px #000000;
       background-color: #E0E0E0;
    }
  </style>

  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/layers.js"></script>
 	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
  <script type="text/javascript" src="../scripts/isvaliddate.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

<script language="javascript">
<!--
$(document).ready(function(){
<% if lcl_orghasfeature_actionline_linkedrequests then %>
  $('input[name*="button_save_"]').css('display','none');
  $('textarea[name*="description_edit_"]').css('display','none');

  if($('#sTotalLinkedRequestRows').val() == '0') {
     $('#linkedRequests').css('visibility','hidden');
     $('#linkedRequestsRow_0').css('visibility','hidden');
  }
<% end if %>
  $('#contactInformation').click(function() {
     toggleDisplay('contact_user');
  });

  $('#issueLocation').click(function() {
     toggleDisplay('issue_location');
  });

  $('#formInformation').click(function() {
     toggleDisplay('comments');
  });

  $('#internalOnlyFields').click(function() {
     toggleDisplay('adminfields');
  });

  $('#feeBalance').click(function() {
     toggleDisplay('fees');
  });

  $('#attachments').click(function() {
     toggleDisplay('file_upload');
  });

  $('#requestActivityLog').click(function() {
     toggleDisplay('log');
  });

  $('#updateActionRequest').click(function() {
     toggleDisplay('update_form');
  });

  $('#sendEmailNotifications').click(function() {
     toggleDisplay('send_email');
  });

  $('#codeSections').click(function() {
     toggleDisplay('code_sections');
  });

  $('#formLetter').click(function() {
     toggleDisplay('letter_form');
  });

  $('#pdfForms').click(function() {
     toggleDisplay('pdf_form');
  });

});

<% if lcl_orghasfeature_actionline_linkedrequests then %>
function addLinkedRequestRow() {
  var LRtbl     = document.getElementById('linkedRequests');
  var totalrows = Number($('#sTotalLinkedRequestRows').val());

  //Increase the total rows by one.  This is index for the new row.
  totalrows = totalrows+1;

  if(totalrows == 1) {
     $('#linkedRequests').css('visibility','visible');
     $('#linkedRequestsRow_0').css('visibility','visible');
  }

  //Set up the new row.
  LRtbl    = document.getElementById('linkedRequests').insertRow(totalrows);
  LRtbl.id = 'linkedRequestsRow_' + totalrows;

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg   = "";
  var lcl_evenodd = totalrows/2;
      lcl_evenodd = lcl_evenodd.toString();

  if(lcl_evenodd.indexOf('.') > 0) {
     lcl_rowbg = "#eeeeee";
  }else{
     lcl_rowbg = "#ffffff";
  }

  LRtbl.style.background = lcl_rowbg;
  LRtbl.valign = "top";

  //Build the cells for the new row.
  var a = LRtbl.insertCell(0);  //Tracking Number
  var b = LRtbl.insertCell(1);  //Description
  var c = LRtbl.insertCell(2);  //Button Row

  //Build the cells in the new row.
  //Tracking Number
  var lcl_trackingnumber  = '<input type="hidden" name="linkedRequestID_'   + totalrows + '" id="linkedRequestID_'     + totalrows + '" value="0" size="10" maxlength="50" />';
      lcl_trackingnumber += '<input type="text" name="trackingnumber_edit_' + totalrows + '" id="trackingnumber_edit_' + totalrows + '" value="" size="15" maxlength="50" />';
      lcl_trackingnumber += '<span id="trackingnumber_text_' + totalrows + '"></span>';

  a.style.verticalAlign = 'top';
  a.innerHTML           = lcl_trackingnumber;

  //Description
  var lcl_description  = '<span id="description_text_'       + totalrows + '"></span>';
      lcl_description += '<textarea name="description_edit_' + totalrows + '" id="description_edit_' + totalrows + '" rows="3" cols="33"></textarea>';

  b.width               = "200"
  b.style.verticalAlign = 'top';
  b.innerHTML           = lcl_description

  //Button Row
  var lcl_button_row  = '<input type="button" name="button_edit_'   + totalrows + '" id="button_edit_'   + totalrows + '" value="Edit" class="button" onclick="editLink(\'' + totalrows + '\');" /> ';
      lcl_button_row += '<input type="button" name="button_save_'   + totalrows + '" id="button_save_'   + totalrows + '" value="Save Changes" class="button" onclick="saveLinkChanges(\'' + totalrows + '\');" /> ';
      lcl_button_row += '<input type="button" name="button_remove_' + totalrows + '" id="button_remove_' + totalrows + '" value="Remove Link" class="button" onclick="" />';

      lcl_button_row += '<input type="button" name="button_add_'    + totalrows + '" id="button_add_'    + totalrows + '" value="Add" class="button" onclick="addLink(\''       + totalrows + '\');" /> ';
      lcl_button_row += '<input type="button" name="button_cancel_' + totalrows + '" id="button_cancel_' + totalrows + '" value="Cancel" class="button" onclick="cancelLink(\'' + totalrows + '\');" />';

  c.style.verticalAlign = 'top';
  c.innerHTML           = lcl_button_row;

  //update the total row count.
  $('#sTotalLinkedRequestRows').val(totalrows)

  //Miscellaneous setup
  document.getElementById('button_edit_'   + totalrows).style.display = 'none';
  document.getElementById('button_save_'   + totalrows).style.display = 'none';
  document.getElementById('button_remove_' + totalrows).style.display = 'none';
  document.getElementById('button_linkrequest').disabled              = true;
}

function addLink(iRowCount) {
  var lcl_parent_trackingnumber = '<%=lngTrackingNumber%>';
  var lcl_parent_requestid      = '<%=iTrackID%>';
  var lcl_linked_trackingnumber;
  var lcl_description;

  if(iRowCount > 0) {
     lcl_linked_trackingnumber = $('#trackingnumber_edit_' + iRowCount).val();
     lcl_description           = $('#description_edit_'    + iRowCount).val();
  }


  var sParameter  = 'action=A';
      sParameter += '&orgid='                 + encodeURIComponent('<%=session("orgid")%>');
      sParameter += '&rowID='                 + encodeURIComponent(iRowCount);
      sParameter += '&parent_trackingnumber=' + encodeURIComponent(lcl_parent_trackingnumber);
      sParameter += '&parent_requestid='      + encodeURIComponent(lcl_parent_requestid);
      sParameter += '&linked_trackingnumber=' + encodeURIComponent(lcl_linked_trackingnumber);
      sParameter += '&description='           + encodeURIComponent(lcl_description);

  doAjax('linkedrequests_action.asp', sParameter, 'addLinkedRequests', 'post', '0');
}

function addLinkedRequests(iReturn) {
  var lcl_return_value = iReturn;

  if(lcl_return_value == 'DUPLICATE') {
     displayScreenMsg('Duplicate Request: The Tracking Number entered has already been "linked" to this Action Line Request...');
  } else {
     if(lcl_return_value.indexOf(lcl_return_value,",") >= 0) {
        var lcl_totalrows = Number(document.getElementById("sTotalLinkedRequestRows").value);
        var lcl_seperator = lcl_return_value.indexOf(",");
        var lcl_rowid;
        var lcl_lrid;
        var lcl_trackingnumber;
        var lcl_trackingnumber_len;
        var lcl_trackingnumber_text;
        var lcl_action_autoid;

        lcl_rowid               = lcl_return_value.substr(0,lcl_seperator);
        lcl_lrid                = lcl_return_value.substr(lcl_seperator+1);
        lcl_trackingnumber      = $('#trackingnumber_edit_' + lcl_rowid).val();
        lcl_trackingnumber_len  = lcl_trackingnumber.length;
        lcl_action_autoid       = lcl_trackingnumber.substr(0,lcl_trackingnumber_len - 4);
        lcl_trackingnumber_text = '<a href="action_respond.asp?init=Y&useSessions=1&control=' + lcl_action_autoid + '">' + lcl_trackingnumber + '</a>';

        $('#linkedRequestID_'     + lcl_rowid).val(lcl_lrid);

        $('#trackingnumber_edit_' + lcl_rowid).fadeOut('slow',function() {
          $('#trackingnumber_text_' + lcl_rowid).html(lcl_trackingnumber_text);
          document.getElementById('button_linkrequest').disabled = false;
        });

        $('#description_edit_' + lcl_rowid).fadeOut('slow',function() {
          $('#description_text_' + lcl_rowid).html($('#description_edit_' + lcl_rowid).val());
          $('#description_text_' + lcl_rowid).fadeIn('slow');
        });

        $('#button_add_' + lcl_rowid).fadeOut('slow',function() {
          $('#button_edit_' + lcl_rowid).fadeIn('slow');
          $('#button_remove_' + lcl_rowid).click(function(){
             removeLink(lcl_rowid,'<%=iTrackID%>',lcl_action_autoid);
          });
        });

        $('#button_cancel_' + lcl_rowid).fadeOut('slow',function() {
          $('#button_remove_' + lcl_rowid).fadeIn('slow');
        });
     }
  }
}

function editLink(iRowCount) {
  $('#button_edit_' + iRowCount).fadeOut('slow', function() {
    $('#button_save_' + iRowCount).fadeIn('slow');
  });

  $('#description_text_' + iRowCount).fadeOut('slow',function() {
    $('#description_edit_' + iRowCount).val($('#description_text_' + iRowCount).html());
    $('#description_edit_' + iRowCount).fadeIn('slow');
  });
}

function saveLinkChanges(iRowCount) {

  var lcl_linkedrequestid  = $('#linkedRequestID_'  + iRowCount).val();
  var lcl_description_edit = $('#description_edit_' + iRowCount).val();

  var sParameter  = 'action=S';
      sParameter += '&rowID='           + encodeURIComponent(iRowCount);
      sParameter += '&linkedrequestid=' + encodeURIComponent(lcl_linkedrequestid);
      sParameter += '&description='     + encodeURIComponent(lcl_description_edit);

  doAjax('linkedrequests_action.asp', sParameter, 'modifyLinkedRequests', 'post', '0');
}

function cancelLink(iRowCount) {
  var totalrows;

  clearScreenMsg();

  $('#linkedRequestsRow_' + iRowCount).hide('slow',function(){
    totalrows = Number($('#sTotalLinkedRequestRows').val());
    totalrows = totalrows - 1;

    if(totalrows < 0) {
       totalrows = 0;
    }

    $('#sTotalLinkedRequestRows').val(totalrows);
    document.getElementById('linkedRequests').deleteRow(iRowCount);
    document.getElementById('button_linkrequest').disabled = false;

    if(totalrows == 0) {
       $('#linkedRequests').css('visibility','hidden');
       $('#linkedRequestsRow_0').css('visibility','hidden');
    }
  });
}

function removeLink(iRowCount, iCurrentRequestID, iRequestToBeRemoved) {

  var iDisplayedTrackingNumber = $('#trackingnumber_edit_' + iRowCount).val();
  var r = confirm('Are you sure you want to delete the linked request: ' + iDisplayedTrackingNumber + '?');

  if(r == true) {
     //Hide the row that is to be deleted.
     $('#linkedRequestsRow_' + iRowCount).hide('slow');

     //Build the parameter string
     var sParameter  = 'action=D';
         sParameter += '&rowID='              + encodeURIComponent(iRowCount);
         sParameter += '&currentrequestid='   + encodeURIComponent(iCurrentRequestID);
         sParameter += '&requestToBeRemoved=' + encodeURIComponent(iRequestToBeRemoved);
         sParameter += '&isAjaxRoutine=Y';

     doAjax('linkedrequests_action.asp', sParameter, 'modifyLinkedRequests', 'post', '0');
  }
}

function modifyLinkedRequests(iReturn) {
  var totalrows;

  if(iReturn == 'D') {
//     totalrows = Number($('#sTotalLinkedRequestRows').val());
//     totalrows = totalrows - 1;

//     if(totalrows < 0) {
//        totalrows = 0;
//     }

//     $('#sTotalLinkedRequestRows').val(totalrows);

//     if(totalrows == 0) {
//        $('#linkedRequests').css('visibility','hidden');
//        $('#linkedRequestsRow_0').css('visibility','hidden');
//     }

//     displayScreenMsg('Successfully Deleted...');
    location.href = 'action_respond.asp?control=<%=iTrackID%>&success=SD';
  } else {
     if(iReturn.indexOf(iReturn,'S_') >= 0) {
        var lcl_rowid;

        lcl_rowid = iReturn.replace('S_','');

        $('#button_save_' + lcl_rowid).fadeOut('slow',function() {
          $('#button_edit_' + lcl_rowid).fadeIn('slow');
        });

        $('#description_edit_' + lcl_rowid).fadeOut('slow',function() {
          $('#description_text_' + lcl_rowid).html($('#description_edit_' + lcl_rowid).val());
          $('#description_text_' + lcl_rowid).fadeIn('slow');
        });

        displayScreenMsg('Changes Successfully Saved...');
     }
  }
}
<% end if %>
function toggleDisplay(iDivID) {
  //lcl_div = document.getElementById(iDivID);

  //if(lcl_div.style.display == "block" || lcl_div.style.display == "") {
  //   lcl_div.style.display = "none";
  //} else {
  //   lcl_div.style.display = "block";
  //}

  if($('#' + iDivID).css('display') == 'none') {
     $('#' + iDivID).slideDown('slow');
  } else {
     $('#' + iDivID).slideUp('slow');
  }
}

function deleteconfirm(ID) {
   if(confirm('Do you wish to permanently delete this request? \nThis will remove all data related to this request.')) {
      window.location="action_line_delete.asp?id=" + ID;
   }
}

//function update_user_display() {
//   if (document.getElementById('user_on').value == 'on') {
//      document.getElementById('user_expand').innerHTML = '<strong>+ <u>Contact Information:</u></strong>';
//      document.getElementById('user_on').value = 'off';
//   }else{
//      document.getElementById('user_expand').innerHTML = '<strong>- <u>Contact Information:</u></strong>';
//      document.getElementById('user_on').value = 'on';
//   }
//}

function openFormLetter(iAction) {
  //var FormData   = document.frmLetter;
  //var LtrData    = FormData.add_text.value;
  //var results    = LtrData.replace("\r", "<br />");
  //var lcl_results  = lcl_addText.replace('\r', '<br />');
  var lcl_letterid = document.getElementById('selLetterId').value;
  var lcl_addText  = document.getElementById('add_text').value;

  var lcl_results = lcl_addText;
  var lcl_reload   = "N";
  var lcl_display_toolbar = "no";
  var newFile;

  for (var i=0; i < lcl_results.length; i++) {
      //if(lcl_results.indexOf('\r') != '-1') {
      lcl_results = lcl_results.replace('&', '<<AMP>>');    //ampersands
      lcl_results = lcl_results.replace('\n', '<br />');    //new line
      lcl_results = lcl_results.replace('\r', '<br />');    //carriage return
      lcl_results = lcl_results.replace('\f', '<br />');    //form feed
      //}
  }

  //Convert ampersands so they do not mess up the URL
  //for (var i=0; i < lcl_results.length; i++) {
  //    if(lcl_results.indexOf("&") != "-1") {
  //       lcl_results = lcl_results.replace("&", "<<AMP>>");
  //    }
  //}

  lcl_letter_url  = "action_line_formletters.asp";
  lcl_letter_url += "?action="    + iAction;
  lcl_letter_url += "&add_text="  + lcl_results;
  //lcl_letter_url += "&iletterid=" + FormData.selLetterId.value;
  lcl_letter_url += "&iletterid=" + lcl_letterid;
  lcl_letter_url += "&iTrackID=<%=iTrackID%>";
  lcl_letter_url += "&iuserid=<%=sTheuserid%>";
  lcl_letter_url += "&status=<%=sStatus%>";
  lcl_letter_url += "&substatus=<%=sSubStatusID%>";

  if(iAction=="PREVIEW") {
     alert("This is a PREVIEW only, No Activity has been logged.");
  }else if((iAction=="PRINT") || (iAction=="EMAIL")) {
     lcl_reload = "Y";

     if(iAction=="PRINT") {
        lcl_display_toolbar = "yes";
     }
  }

  newFile = lcl_letter_url;

  if(lcl_reload=="Y") {
     parent.location.reload();
  }

  newWin = window.open(newFile,'popupName','width=600,height=500,toolbar=' + lcl_display_toolbar + ',left=50,top=50,scrollbars=yes,resizable=yes,status=yes');
  newWin.focus();
  return false;
}

//function PDFdocument() {
//   var FormData = document.frmLetter;
//   var LtrData  = FormData.add_text.value;
//   var results  = LtrData.replace("\r", "<br />");

//   for (var i=0; i < results.length; i++) {
//      if (results.indexOf("\r") != "-1") {
//         results = results.replace("\r", "<br />");
//      }
//   }

   //var newFile = "http://secure.eclink.com/egovlink/action_line/actionline_pdf.asp?sys=<%=Application("INSTANCE")%>&add_text=" + results + "&iletterid=" + FormData.selLetterId.value + "&action_autoid=<%=iTrackID%>&status=<%=sStatus%>&substatus=<%=sSubStatusID%>&orgid=<%=session("orgid")%>&userid=<%=session("userid")%>";
//   var newFile = "actionline_pdf.asp?sys=<%=Application("INSTANCE")%>&add_text=" + results + "&iletterid=" + FormData.selLetterId.value + "&action_autoid=<%=iTrackID%>&status=<%=sStatus%>&substatus=<%=sSubStatusID%>&orgid=<%=session("orgid")%>&userid=<%=session("userid")%>";
//   newWin = window.open(newFile);
//   newWin.focus();
//   return false;
//}

function fnDisplayPDF(){
   //var newFile = 'http://secure.eclink.com/egovlink/request_to_pdf_merge.asp?sys=<%=Application("INSTANCE")%>&irequestid=<%=iTrackID%>';
   var newFile = 'request_to_pdf_merge.asp?sys=<%=Application("INSTANCE")%>&irequestid=<%=iTrackID%>';
   newWin = window.open(newFile);
   newWin.focus();
   return false;
}

<%
 'Build the Attachment URL
  lcl_attachment_url = Application("common_url")
  lcl_attachment_url = lcl_attachment_url & "/public_documents300"
  lcl_attachment_url = lcl_attachment_url & "/" & session("virtualdirectory")
  lcl_attachment_url = lcl_attachment_url & "/attachments/"

  response.write "function viewAttachment(iFileName) { " & vbcrlf
  response.write "  lcl_width  = 800;" & vbcrlf
  response.write "  lcl_height = 700;" & vbcrlf
  response.write "  lcl_left   = (screen.availWidth/2)-(lcl_width/2);" & vbcrlf
  response.write "  lcl_top    = (screen.availHeight/2)-(lcl_height/2);" & vbcrlf

  response.write "  window.open('" & lcl_attachment_url & "' + iFileName, '_attachment', 'width=' + lcl_width + ',height=' + lcl_height + ',resizable=1,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + lcl_left + ',top=' + lcl_top);" & vbcrlf
  response.write "}"
%>
function confirm_delete(iattachmentid,irequestid) {
   if (confirm("Are you sure you want to delete this attachment?")) { 
      // DELETE HAS BEEN VERIFIED
      location.href="attachment_delete.asp?attachmentid=" + iattachmentid + "&irequestid=" + irequestid + "&status=<%=sStatus%>&substatusid=<%=sSubStatusID%>";
   }
}

<%
  if blnCanEdit AND lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then %>
function changeSubStatus(p_SubStatusID) {

  var list
  var list2
  var i
  var a
  var lcl_sub_status_id

  mainlist          = document.getElementById('selStatus');
  sub_list          = document.getElementById('selSubStatus');
  sub_list_row      = document.getElementById('sub_status_row');
  i = 0

  if((p_SubStatusID=="")||(p_SubStatusID==undefined)) {
      lcl_sub_status_id = "";
  }else{
      lcl_sub_status_id = p_SubStatusID;
  }
<%
  dim oMainStatus, oSubStatus, oSubStatus_Count, line_count, lcl_sub_line_count, lcl_total_count

 'Retrieve all of the MAIN statuses
  sSqlm = "SELECT action_status_id, status_name, orgid, parent_status, display_order, active_flag "
  sSqlm = sSqlm & " FROM egov_actionline_requests_statuses "
  sSqlm = sSqlm & " WHERE orgid = 0 "
  sSqlm = sSqlm & " AND parent_status = 'MAIN' "
  sSqlm = sSqlm & " AND active_flag = 'Y' "
  sSqlm = sSqlm & " ORDER BY display_order "

  set oMainStatus = Server.CreateObject("ADODB.Recordset")
  oMainStatus.Open sSqlm, Application("DSN"), 3, 1

  if not oMainStatus.eof then
     line_count = 0
	    do while NOT oMainStatus.EOF
        line_count = line_count + 1

	      	if line_count = 1 then
           response.write "if(mainlist.value==""" & oMainStatus("status_name") & """) {" & vbcrlf
        else
           response.write "}else if(mainlist.value==""" & oMainStatus("status_name") & """) {" & vbcrlf
        end if

   	   'Get the total count of SubStatuses
        sSQLc = "SELECT count(action_status_id) AS Total_SubStatus FROM egov_actionline_requests_statuses "
        sSQLc = sSQLc & " WHERE orgid = "         & clng(session("orgid"))
        sSQLc = sSQLc & " AND parent_status = '"  & oMainStatus("status_name") & "' "
        sSQLc = sSQLc & " AND active_flag = 'Y' "
        Set oSubStatus_Count = Server.CreateObject("ADODB.Recordset")
        oSubStatus_Count.Open sSQLc, Application("DSN"), 3, 1

        lcl_total_count = oSubStatus_Count("Total_SubStatus")

        if lcl_total_count > 0 then
		  
		        'Retrieve all of the Sub-Statuses for each MAIN status for the orgid and the form
           sSqls = "SELECT action_status_id, status_name "
           sSqls = sSqls & " FROM egov_actionline_requests_statuses "
           sSqls = sSqls & " WHERE orgid = "         & clng(session("orgid"))
           sSqls = sSqls & " AND parent_status = '"  & oMainStatus("status_name") & "' "
           sSqls = sSqls & " AND active_flag = 'Y' "
           sSqls = sSqls & " ORDER BY display_order "

           Set oSubStatus = Server.CreateObject("ADODB.Recordset")
           oSubStatus.Open sSqls, Application("DSN"), 3, 1

           If NOT oSubStatus.EOF Then
              response.write "//remove the current values" & vbcrlf
              response.write "for(var i=0; i < sub_list.length; i++) {" & vbcrlf
              response.write "    sub_list.remove(i);" & vbcrlf
              response.write "}" & vbcrlf
              response.write "//default with blank value" & vbcrlf
              response.write "document.forms[""frmUpdate""].selSubStatus.options[0] = new Option("""",""0"");" & vbcrlf

              response.write "sub_list.disabled = false;" & vbcrlf

             'Loop through the sub statuses
              lcl_sub_line_count = 1
              do while NOT oSubStatus.EOF

                 response.write "//build the new values" & vbcrlf
                 response.write "document.forms[""frmUpdate""].selSubStatus.options[" & lcl_sub_line_count & "] = new Option(""" & oSubStatus("status_name") & """,""" & oSubStatus("action_status_id") & """);" & vbcrlf

                 response.write "if(lcl_sub_status_id==" & oSubStatus("action_status_id") & ") {" & vbcrlf
                 response.write "   sub_list.selectedIndex=" & lcl_sub_line_count & ";" & vbcrlf
                 response.write "}" & vbcrlf

            				 lcl_sub_line_count = lcl_sub_line_count + 1
			            	 oSubStatus.movenext
              loop

         			  oSubStatus.Close
          		  oSubStatus_Count.Close

         			  set oSubStatus       = nothing
         			  set oSubStatus_Count = nothing 
		   
      		   else
              response.write "sub_list.disabled = true;" & vbcrlf
              response.write "sub_list.value = '';" & vbcrlf
           end if
        else
           response.write "sub_list.disabled = true;" & vbcrlf
           response.write "sub_list.value = '';" & vbcrlf
      		end if

        oMainStatus.movenext
     loop

     response.write "}" & vbcrlf

  end if

  oMainStatus.Close
  set oMainStatus = nothing 

  response.write "}" & vbcrlf
  end if
%>
function checkDepartmentInactive() {
<%
 'Build the error message
  lcl_error_msg = "<strong>Cannot send e-mail.</strong>  This department has been deleted and is no longer active in the system."
  lcl_error_msg = lcl_error_msg & " E-mails cannot be sent to this department."

 'If no department has been entered then do not send the email.
  response.write "lcl_dept = document.getElementById(""notifydeptid"").value;" & vbcrlf
  response.write "lcl_user = document.getElementById(""notifyuserid"").value;" & vbcrlf

 'Retrieve all of the INACTIVE groups
  sSQLd = "SELECT groupid FROM groups "
  sSQLd = sSQLd & " WHERE orgid = " & session("orgid")
  sSQLd = sSQLd & " AND isInactive = 1 "

  set oDept = Server.CreateObject("ADODB.Recordset")
  oDept.Open sSQLd, Application("DSN"), 3, 1

  lcl_inactive_depts = ""

  if NOT oDept.eof then
     while not oDept.eof
        if lcl_inactive_depts = "" then
           lcl_inactive_depts = "lcl_dept==" & oDept("groupid")
        else
           lcl_inactive_depts = lcl_inactive_depts & " || lcl_dept==" & oDept("groupid")
        end if
        oDept.movenext
     wend
  end if

  oDept.close
  set oDept = nothing

  response.write "//Make sure that atleast one dropdown has a value in it." & vbcrlf
  response.write "if((lcl_dept=="""") && (lcl_user=="""")) {" & vbcrlf
  response.write "   document.getElementById(""notifyuserid"").focus();" & vbcrlf
  response.write "   inlineMsg(document.getElementById(""notifyuserid"").id,'<strong>Required field missing: </strong>  A User and/or Department must be selected before sending a notification.',10,'notifyuserid');" & vbcrlf
  response.write "  	return false;" & vbcrlf

  if lcl_inactive_depts <> "" then
     response.write "}else{" & vbcrlf
     response.write "//If the group selected is INACTIVE then do not allow the email to be sent." & vbcrlf
     response.write "  if(" & lcl_inactive_depts & ") {" & vbcrlf
     response.write "     document.getElementById(""notifydeptid"").focus();" & vbcrlf
     response.write "     inlineMsg(document.getElementById('notifydeptid').id,'" & lcl_error_msg & "',10,'notifydeptid');" & vbcrlf
     response.write " 		 	return false;" & vbcrlf
     response.write "  }" & vbcrlf
  end if

  response.write "}" & vbcrlf
%>
}

//  function doPicker(iFieldID) {
//    lcl_width  = 600;
//    lcl_height = 400;
//    lcl_left   = (screen.availWidth/2)-(lcl_width/2);
//    lcl_top    = (screen.availHeight/2)-(lcl_height/2);

//    eval('window.open("linkpicker/linkpicker.asp?fid=' + iFieldID + '", "_picker", "width=' + lcl_width + ',height=' + lcl_height + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + lcl_left + ',top=' + lcl_top + '")');
//  }

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
  w = 600;
  h = 400;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";
  lcl_folderStart     = 0;

  //Determine which options will be displayed
  if((p_displayDocuments=="")||(p_displayDocuments==undefined)) {
      lcl_displayDocuments = "";
  }else{
      lcl_displayDocuments = "&displayDocuments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayActionLine=="")||(p_displayActionLine==undefined)) {
      lcl_displayActionLine = "";
  }else{
      lcl_displayActionLine = "&displayActionLine=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayPayments=="")||(p_displayPayments==undefined)) {
      lcl_displayPayments = "";
  }else{
      lcl_displayPayments = "&displayPayments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayURL=="")||(p_displayURL==undefined)) {
      lcl_displayURL = "";
  }else{
      lcl_displayURL = "&displayURL=Y";
  }

  if(lcl_folderStart > 0) {
     //lcl_showFolderStart = "&folderStart=unpublished_documents";
     lcl_showFolderStart = "&folderStart=CITY_ROOT";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
		    var caretPos = textEl.caretPos;
  			 caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? text + ' ' : text;
  } else {
   			textEl.value = textEl.value + text;
	 }
}

	function openPDFtoView( iRequestID, iDocumentID, iAction ) 
	{
		if(iRequestID == '') 
		{
			return false;
		}
		else
		{
			if(iDocumentID != '') 
			{
				// Preview_pdf and Print_pdf are handled here
				//window.open('viewPDF.asp?iRequestID='+iRequestID+'&docID='+iDocumentID+'&pdfaction='+iAction+'&hideActLog=<%=lcl_hide_activitylog%>');
				window.open('viewXMLPDF.asp?iRequestID=' + iRequestID + '&docID=' + iDocumentID + '&pdfaction=' + iAction + '&hideActLog=<%=lcl_hide_activitylog%>');
				// we are going to redirect to see what is being gathered in the viewing file. This is a short term thing, not prod.
				//location.href = 'viewXMLPDF.asp?iRequestID=' + iRequestID + '&docID=' + iDocumentID + '&pdfaction=' + iAction + '&hideActLog=<%=lcl_hide_activitylog%>';

				if(iAction == 'PRINT_PDF') 
				{
					location.href = 'action_respond.asp?control=<%=iTrackID%>&actlog='+iAction+'&pdfid='+document.getElementById("selectPDF").value;
				}

				return true;
			}
			else
			{
				if(iAction == 'PREVIEW_PDF' || iAction == 'PRINT_PDF')
				{
					document.getElementById("selectPDF").focus();
					inlineMsg(document.getElementById('selectPDF').id,'<strong>Required Field Missing: </strong>Please select a PDF.',10,'selectPDF');
					return false;
				} 
				else if(iAction == 'WORKORDER') 
				{
					//window.open('viewPDF.asp?iRequestID='+iRequestID+'&pdfaction=WORKORDER&hideActLog=<%=lcl_hide_activitylog%>');
					window.open('pdfview/work_order.aspx?iRequestID='+iRequestID+'&pdfaction=WORKORDER&hideActLog=<%=lcl_hide_activitylog%>');
					location.href = 'action_respond.asp?control=<%=iTrackID%>&actlog='+iAction;
					return true;
				} 
				else if(iAction == 'WORKORDER_CONDENSED') 
				{
					//window.open('viewPDF.asp?iRequestID='+iRequestID+'&pdfaction=WORKORDER_CONDENSED&hideActLog=<%=lcl_hide_activitylog%>');
					window.open('pdfview/work_order.aspx?iRequestID='+iRequestID+'&pdfaction=WORKORDER_CONDENSED&hideActLog=<%=lcl_hide_activitylog%>');
					location.href = 'action_respond.asp?control=<%=iTrackID%>&actlog='+iAction;
					return true;
				} 
				else 
				{
					var lcl_filename = document.getElementById("public_actionline_pdf").value;
					var lcl_file_ext = lcl_filename.substr(lcl_filename.length-4);

					if(lcl_file_ext.toUpperCase() == ".PDF") 
					{
						//window.open('viewPDF.asp?iRequestID='+iRequestID+'&pdf='+lcl_filename+'&pdfaction='+iAction+'&hideActLog=<%=lcl_hide_activitylog%>');
						window.open('viewXMLPDF.asp?iRequestID=' + iRequestID + '&pdf=' + lcl_filename + '&pdfaction=' + iAction + '&hideActLog=<%=lcl_hide_activitylog%>');
						// we are going to redirect to see what is being gathered in the viewing file. This is a short term thing, not prod.
						//location.href = 'viewXMLPDF.asp?iRequestID=' + iRequestID + '&pdf=' + lcl_filename + '&pdfaction=' + iAction + '&hideActLog=<%=lcl_hide_activitylog%>';

						location.href = 'action_respond.asp?control=<%=iTrackID%>&actlog='+iAction;
						return true;
					}
					else
					{
						document.getElementById("public_actionline_pdf").focus();
						inlineMsg(document.getElementById('viewPDF').id,'<strong>Invalid Value: </strong>The file is not a valid PDF file.',10,'viewPDF');
						return false;
					}
				}
			}
		}
	}

	function openTestPDF( iRequestID )
	{
		window.open('pdfview/work_order2.aspx?iRequestID=' + iRequestID );
	}


function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

<%
'------------------------------------------------------------------------------
 'Check to see if the org and user both have the "Action Line - Email Reminders" feature assigned.
 'If not then skip this function
  if lcl_orghasfeature_actionline_emailreminders AND lcl_userhaspermission_actionline_emailreminders then
'------------------------------------------------------------------------------
%>
function addReminderRow() {
  var mytbl     = document.getElementById('AddReminderTBL');
  var totalrows = Number(document.getElementById("sTotalReminderRows").value);

  //Increase the total rows by one.  This is index for the new row.
  totalrows = totalrows+1;

  //Set up the new row.
  mytbl = document.getElementById('AddReminderTBL').insertRow(totalrows);

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg   = "";
  var lcl_evenodd = totalrows/2;
      lcl_evenodd = lcl_evenodd.toString();

  if(lcl_evenodd.indexOf('.') > 0) {
     lcl_rowbg = "#ffffff";
  }else{
     lcl_rowbg = "#eeeeee";
  }

  mytbl.style.background = lcl_rowbg;

  //Build the cells for the new row.
  var a = mytbl.insertCell(0);  //Send Reminder To
  var b = mytbl.insertCell(1);  //Reminder Date
  var c = mytbl.insertCell(2);  //Additional Comments
  var d = mytbl.insertCell(3);  //Created Info
  var e = mytbl.insertCell(4);  //Remove Row (checkbox)

  //Build the cells in the new row.
  //Send Reminder To
  var lcl_sendto = '<input type="hidden" name="reminderid_'+totalrows+'" id="reminderid_'+totalrows+'" value="0" size="10" maxlength="10" />';
  lcl_sendto = lcl_sendto + '<select name="reminderSendTo_'+totalrows+'" id="reminderSendTo_'+totalrows+'" onchange="clearMsg(\'reminderSendTo_'+totalrows+'\')">';
  lcl_sendto = lcl_sendto + '<option value=""></option>';
  lcl_sendto = lcl_sendto + <% DrawAdminUsersNew_javascript "","Y" %>;
  lcl_sendto = lcl_sendto + '</select>';
  a.innerHTML=lcl_sendto;

  //Reminder Date
  var lcl_reminderdate_field = '<input type="text" name="reminderDate_'+totalrows+'" id="reminderDate_'+totalrows+'" size="10" maxlength="10" onchange="clearMsg(\'reminderDateLOV_'+totalrows+'\')" /> ';
      lcl_reminderdate_field = lcl_reminderdate_field + '<img src="../images/calendar.gif" id="reminderDateLOV_'+totalrows+'" border="0" style="cursor:hand" onclick="clearMsg(\'reminderDateLOV_'+totalrows+'\');doCalendar(\'reminderDate_'+totalrows+'\');" />';

  b.innerHTML=lcl_reminderdate_field

  //Additional Comments
  c.innerHTML='<textarea name="reminderComments_'+totalrows+'" id="reminderComments_'+totalrows+'" rows="2" cols="40"></textarea>';

  //Created Info
  d.innerHTML='&nbsp;';

  //Remove Row (checkbox)
  e.align="center";
  e.innerHTML='<input type="checkbox" name="reminderRemove_'+totalrows+'" id="reminderRemove_'+totalrows+'" value="Y" />';

  //update the total row count.
  document.getElementById("sTotalReminderRows").value = totalrows;
}
<%
'------------------------------------------------------------------------------
 end if 'End check for "Action Line - Email Reminders org/user feature assigned
'------------------------------------------------------------------------------
%>

function updateRequest() {
  lcl_form      = document.getElementById("frmUpdate");
  lcl_false_cnt = 0;

<%
'BEGIN: Due Date --------------------------------------------------------------
 if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
%>
  lcl_dueDate   = document.getElementById("due_date").value;
		var daterege  = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var dueDateOk = daterege.test(lcl_dueDate);

		if (lcl_dueDate != "" && ! dueDateOk ) {
      lcl_focus = document.getElementById("due_date");
      inlineMsg(document.getElementById("due_date_pop").id,'<strong>Invalid Value: </strong> The "Due Date" must be in date format.<br /><span class=""darkRedText"">(i.e. mm/dd/yyyy)</span>',10,'due_date_pop');
      lcl_false_cnt = lcl_false_cnt + 1
  }else{
      clearMsg("due_date_pop");
  } 

<%
 end if
'END: Due Date ----------------------------------------------------------------

'BEGIN: Email Reminders -------------------------------------------------------
'Check to see if the org and user both have the "Action Line - Email Reminders" feature assigned.
'If not then skip this validation.
 if lcl_orghasfeature_actionline_emailreminders AND lcl_userhaspermission_actionline_emailreminders then
'------------------------------------------------------------------------------
%>
  if(lcl_false_cnt == 0) {
     lcl_total_reminders = document.getElementById("sTotalReminderRows").value;
     lcl_i_start         = 1;

     //---------------------------------------------------------------------------
     //Check the Send Reminder To for ALL of the email reminders.
     //---------------------------------------------------------------------------
     //for (i=lcl_i_start; i<=lcl_total_reminders; ++ i) {
     for (i=lcl_total_reminders; lcl_i_start<=i; -- i) {
          if(document.getElementById("reminderSendTo_"+i).value == "") {
             //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
             if(document.getElementById("reminderRemove_"+i).checked!=true && (document.getElementById("reminderDate_"+i).value != "" || document.getElementById("reminderComments_"+i).value != "")) {
            			 inlineMsg(document.getElementById("reminderSendTo_"+i).id,'<strong>Required Field Missing: </strong>Send Reminder To',8,'reminderSendTo_'+i);
                lcl_false_cnt = lcl_false_cnt + 1;
             }else{
                clearMsg('reminderSendTo_'+i);
                document.getElementById("reminderRemove_"+i).checked=true;
             }

             if(lcl_false_cnt == 1) {
                lcl_focus = document.getElementById("reminderSendTo_"+i);
             }
          }else{
             clearMsg('reminderSendTo_'+i);
     	   	}
     }
  }

  if(lcl_false_cnt == 0) {
    //---------------------------------------------------------------------------
    //Check the Reminder Send Date for ALL of the email reminders.
    //---------------------------------------------------------------------------
    lcl_i_start = 1;
    //for (i=lcl_i_start; lcl_total_reminders>=i; -- i) {
    for (i=lcl_total_reminders; lcl_i_start<=i; -- i) {
         if(document.getElementById("reminderDate_"+i).value == "") {
            //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
            if(document.getElementById("reminderRemove_"+i).checked!=true && (document.getElementById("reminderSendTo_"+i).value != "" || document.getElementById("reminderComments_"+i).value != "")) {
           			 inlineMsg(document.getElementById("reminderDateLOV_"+i).id,'<strong>Required Field Missing: </strong>Reminder Send Date',8,'reminderDateLOV_'+i);
               lcl_false_cnt = lcl_false_cnt + 1;
            }else{
               document.getElementById("reminderRemove_"+i).checked=true;
            }

            if(lcl_false_cnt == 1) {
               lcl_focus = document.getElementById("reminderDateLOV_"+i);
            }
         }else{
            clearMsg('reminderDateLOV_'+i);

            if(document.getElementById("reminderRemove_"+i).checked!=true && document.getElementById("reminderDate_"+i).value!="" && (document.getElementById("reminderSendTo_"+i).value != "" || document.getElementById("reminderComments_"+i).value != "")) {
              	//Validate the format of the Renewal Start Date
              	var Ok = isValidDate(document.getElementById("reminderDate_"+i).value);
              	if(! Ok)	{
                  var lcl_message = "<strong>Invalid Value: </strong>The \"Reminder Send Date\" must be in a date format.<br /><span class=\"darkRedText\">(i.e. mm/dd/yyyy)</span>";
              			 inlineMsg(document.getElementById("reminderDateLOV_"+i).id,lcl_message,8,'reminderDateLOV_'+i);
                  lcl_false_cnt = lcl_false_cnt + 1;

                  if(lcl_false_cnt == 1) {
                     lcl_focus = document.getElementById("reminderDateLOV_"+i);
                  }
               }else{
                  clearMsg('reminderDateLOV_'+i);
               }
            }
       		}
    }
  }

  //If error messages exist then do not submit the form and return focus to the first field found in error.
//  if(lcl_false_cnt > 0) {
//     lcl_focus.focus();
//     return false;
//  }else{
//     lcl_false_cnt = 0;
//  }
<%
'------------------------------------------------------------------------------
 end if 'End check for "Action Line - Email Reminders org/user feature assigned
'END: Email Reminders ---------------------------------------------------------
%>

  if(lcl_false_cnt == 0) {

     //Check the "Note to Citizen".  If a value has been entered then check to see
     //  if the "send an email to citizen" has been checked.  If "no" then display
     //  an alert letting the user know.  If they they choose to continue then
     //  submit the form.  Otherwise, do not submit the form.

     if(document.getElementById("sendemail")) {
        if(document.getElementById("external_comment").value != "" && document.getElementById("sendemail").checked == false) {
           var r=confirm("A 'Note to Citizen' has been entered, but the 'Send email to Citizen' option has not been checked.  Would you like to continue?");
            if (r!=true) {
                lcl_focus     = document.getElementById("sendemail");
                lcl_false_cnt = 1;
            }
        }else if(document.getElementById("external_comment").value == "" && document.getElementById("sendemail").checked == true) {
           var r=confirm("The 'Send email to Citizen' option has been checked, but the 'Note to Citizen' has not been entered.  Would you like to continue?");
            if (r!=true) {
                lcl_focus     = document.getElementById("external_comment");
                lcl_false_cnt = 1;
            }
        }
     }
  }

  //Submit the form if there are no errors
  if(lcl_false_cnt > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("frmUpdate").submit();
  }
}

function enableDisableButton(iButtonID,iAction) {
  //Check for the button's HTML DOM ID.
  if(iButtonID != '') {
     lcl_button = document.getElementById(iButtonID);

     //Determine if we are disabling or enabling the button
     if(iAction == 'DISABLE') {
        //ONLY DISABLE the button IF the button is ENABLED
        if(lcl_button.disabled == false) {
           lcl_button.disabled = true;
        }
     } else {
        //ONLY ENABLE the button IF the button is DISABLED
        if(lcl_button.disabled == true) {
           lcl_button.disabled = false;
        }
     }

  }
}

function setupPDFButtons(iValue) {

  if(iValue != '') {
     enableDisableButton('previewPDF','ENABLE');
     enableDisableButton('printPDF','ENABLE')
  } else {
     enableDisableButton('previewPDF','DISABLE');
     enableDisableButton('printPDF','DISABLE')
  }
}

function validateAttachment() {
  var lcl_attachment = document.getElementById('filAttachment').value;

  //Always initially disable the button.  The code with determine if/when it is enabled.
  document.getElementById("saveAttachmentButton").disabled=true;

  if(lcl_attachment != "") {
     document.getElementById("saveAttachmentButton").disabled=true;

     //Make sure that an .EXE file is not being uploaded.
     lcl_length     = lcl_attachment.length;
     lcl_period_loc = lcl_attachment.indexOf(".");
     lcl_ext        = lcl_attachment.substr(lcl_period_loc+1);

     if(lcl_ext.toUpperCase()=='EXE') {
        alert("Invalid Value: File Type (" + lcl_ext + ").");
        return false;
     }else{
        document.getElementById("saveAttachmentButton").disabled=false;
     }
  }
}

function modifyAttachmentSecurity(iAttachmentID) {

  var lcl_isSecure = "off";

  if(document.getElementById("secureAttachment"+iAttachmentID).checked==true) {
     lcl_isSecure = "on";
  }

  //Build the parameter string
		var sParameter  = 'isAjaxRoutine=Y';
  sParameter     += '&isSecure='     + encodeURIComponent(lcl_isSecure);
  sParameter     += '&attachmentID=' + encodeURIComponent(iAttachmentID);

  clearScreenMsgAttachment();
  doAjax('updateAttachmentSecurity.asp', sParameter, 'displayScreenMsgAttachment', 'post', '0');
}

function modifyAttachmentsPublicDisplay(iAttachmentID) {

  var lcl_displayToPublic = "off";

  if(document.getElementById("displayToPublic"+iAttachmentID).checked==true) {
     lcl_displayToPublic = "on";
  }

  //Build the parameter string
		var sParameter  = 'isAjaxRoutine=Y';
  sParameter     += '&displayToPublic=' + encodeURIComponent(lcl_displayToPublic);
  sParameter     += '&attachmentID='    + encodeURIComponent(iAttachmentID);

  clearScreenMsgAttachment();
  doAjax('updateAttachmentDisplayToPublic.asp', sParameter, 'displayScreenMsgAttachment', 'post', '0');
}

<% if lcl_pushcontent then %>
function pushContent(iID) {
  iFeature = document.getElementById("pushFeature").value;

  if (iFeature == "FAQ" || iFeature == "RUMORMILL") {
      location.href = "../faq/manage_faq.asp?faqtype=" + iFeature + "&requestid=" + iID;

  } else if (iFeature == "COMMUNITYCALENDAR") {
      location.href = "../events/newevent.asp?requestid=" + iID;
  }
}
<% end if %>

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}

function displayScreenMsgAttachment(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsgAttachment").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsgAttachment()", (10 * 1000));
  }
}

function clearScreenMsgAttachment() {
  document.getElementById("screenMsgAttachment").innerHTML = "";
}

//-->
</script>
</head>
<%
  'Determine if there is a status message to display or not
   if blnUpdate then 
      iID        = request("FORM_ID")
      lcl_onload = "displayScreenMsg('Successfully Updated...');"
   elseif blnNotify then
      lcl_onload = "displayScreenMsg('Successfully Sent Notification...');"
   elseif request("success") = "NO_NOTIFY_SENDTO" then
      lcl_onload = "displayScreenMsg('Notification NOT SENT - A User and/or Department must be selected...');"
   elseif request("success") = "ATTACHMENT_ADDED" then
      lcl_onload = "displayScreenMsg('Attachment Successfully Uploaded...');"
   elseif request("success") = "SU" then
      lcl_onload = "displayScreenMsg('Successfully Updated...');"
   elseif request("success") = "SD" then
      lcl_onload = "displayScreenMsg('Successfully Deleted...');"
   else
      lcl_onload = ""
   end if

  'Check to see if the sub-status needs to be set.
   if blnCanEdit AND lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then
      if lcl_onload <> "" then
         lcl_onload = lcl_onload & "changeSubStatus(" & sSubStatusID & ");"
      else
         lcl_onload = "changeSubStatus(" & sSubStatusID & ");"
      end if
   end if

  'Enable/Disable the "SAVE" button for attachments.
  'Also make sure that the user has the (user) permission to be able to "edit".
   if lcl_orghasfeature_fileupload AND lcl_userhaspermission_fileupload AND lcl_userhaspermission_requestedit AND blnCanEdit then 
      if lcl_onload <> "" then
         lcl_onload = lcl_onload & "validateAttachment();"
      else
         lcl_onload = "validateAttachment();"
      end if
   end if

   response.write "<body onload=""" & lcl_onload & """>" & vbcrlf

   ShowHeader sLevel
%>
  <!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<table border=""0"" cellpadding=""2"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td id=""requestTitle"">" & vbcrlf
  response.write "          <font size=""+1""><strong>Review/Respond to Action Line Request (" & lngTrackingNumber & ")</strong></font><br />" & vbcrlf
                            displayButtons iTrackID, "WORKORDER"
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td align=""right"">" & lcl_pushcontent_dropdown & "</td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf

 'BEGIN: Action Line Request List ---------------------------------------------
 'Display the form name
  if sTitle <> "" then
     lcl_display_formtitle = sTitle
  else
     lcl_display_formtitle = "<font class=""redText"">UNKNOWN</font>"
  end if

  response.write "<blockquote>" & vbcrlf
  response.write "<table border=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>"  & vbcrlf
  response.write "</table>" & vbcrlf

 'BEGIN: Tracking Number, Date/Time Received, and Created By ------------------
 	sIsRootAdmin = UserIsRootAdmin(session("userid"))

  response.write "<table border=""0"" padding=""2"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <table border=""0"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td colspan=""2""><h3>" & lcl_display_formtitle & "</h3></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap""><strong>Tracking Number: </strong></td>" & vbcrlf
  response.write "                <td nowrap=""nowrap"" class=""darkRedText"">" & lngTrackingNumber & "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td nowrap=""nowrap""><strong>Date Time Received: </strong></td>" & vbcrlf
  response.write "                <td nowrap=""nowrap"" class=""darkRedText"">" & datSubmitDate & " ( " & getLocalTimeZone() & ")</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td nowrap=""nowrap""><strong>Created By:</strong></td>" & vbcrlf
  response.write "                <td nowrap=""nowrap"" class=""darkRedText"">" & sSubmitName & "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td nowrap=""nowrap""><strong>Completed Date:</strong></td>" & vbcrlf
  response.write "                <td nowrap=""nowrap"" class=""darkRedText"">" & sCompleteDate & "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  if sIsRootAdmin then
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td nowrap=""nowrap""><strong>Submitted By IP Address:</strong></td>" & vbcrlf
     response.write "                <td nowrap=""nowrap"" class=""darkRedText"">" & sSubmittedByRemoteAddress & "</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if

  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"">" & vbcrlf
                         if lcl_orghasfeature_actionline_linkedrequests then
                            displayLinkedRequests session("orgid"), iTrackID
                         end if
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 'END: Tracking Number, Date/Time Received, and Created By --------------------

 'BEGIN: Contact Information Correction ---------------------------------------
  response.write "<p>" & vbcrlf
  response.write "  + <span id=""contactInformation"" class=""user_expand"">Contact Information:</span>" & vbcrlf

  if lcl_userhaspermission_requestedit AND blnCanEdit then
     response.write "<input type=""button"" name=""editContactInfo"" id=""editContactInfo"" style=""cursor:pointer"" value=""Edit"" onclick=""location.href='corrections/correction_contact_info.asp?irequestid=" & iTrackID & "&status=" & sStatus & "&substatus=" & sSubStatusID & "';"" />" & vbcrlf
  end if

  response.write "<input type=""hidden"" id=""user_on"" name=""user_on"" value=""off"">" & vbcrlf

 'Get Contact Information
  fnDisplayUserInfo(sTheuserid)

  response.write "</p>" & vbcrlf
 'END: Contact Information Correction -----------------------------------------

 'BEGIN: Issue Location Information -------------------------------------------
  if bFormHasIssueLocation then
     SubDrawIssueLocationInformation iTrackID, blnCanEdit, sHideIssueLocAddInfo
  end if
 'END: Issue Location Information ---------------------------------------------

 'BEGIN: Form Correction ------------------------------------------------------
  response.write "<p>" & vbcrlf
  response.write "  + <span id=""formInformation"" class=""user_expand"">Form Information:</span>" & vbcrlf

 'Determine if the user can edit this section
  if lcl_userhaspermission_requestedit AND blnCanEdit then 
     response.write "<input type=""button"" name=""editFormInfo"" id=""editFormInfo"" style=""cursor:pointer"" value=""Edit"" onclick=""location.href='corrections/correction_request_form.asp?irequestid=" & iTrackID & "&ftype=PUB&status=" & sStatus & "&substatus=" & sSubStatusID & "';"" />" & vbcrlf
  end if

 'Display Action Item Information
  if sComment <> "" then
    'Format the comment
     lcl_display_comment = ""

     if sComment <> "" then
        lcl_display_comment = formatActivityLogComment(sComment)
     end if

     response.write "<div id=""comments"" class=""divSection"">" & lcl_display_comment & "</div>" & vbcrlf
  else
     response.write "<div id=""comments"" class=""formInformation_noComments"">&nbsp;&nbsp;&nbsp;<em><font class=""redText"">No comment/description provided!</em></font></div>" & vbcrlf
  end if

  response.write "</p>" & vbcrlf
 'END: Form Correction --------------------------------------------------------

 'BEGIN: Admin Fields Correction ----------------------------------------------
  if lcl_userhaspermission_internalfields then
     response.write "<p>" & vbcrlf
     response.write "  + <span id=""internalOnlyFields"" class=""user_expand"">Internal Only Use - Administrative Fields:</span>" & vbcrlf

    'Determine if the user can edit this section
     if lcl_userhaspermission_requestedit AND blnCanEdit then
        response.write "<input type=""button"" name=""editAdminFields"" id=""editAdminFields"" style=""cursor:pointer"" value=""Edit"" onclick=""location.href='corrections/correction_request_form.asp?irequestid=" & iTrackID & "&ftype=INT&status=" & sStatus & "&substatus=" & sSubStatusID & "';"" />" & vbcrlf
     end if

     response.write "<div id=""adminfields"" class=""divSection"">" & vbcrlf

     DisplayFormFieldsandAnswers iTrackID,1

     response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if
 'END: Admin Fields Correction ------------------------------------------------

 'BEGIN: Fee Balance ----------------------------------------------------------
  if lcl_userhaspermission_bzfees AND blnFeeDisplay then
     response.write "<p>" & vbcrlf
     response.write "  + <span id=""feeBalance"" class=""user_expand"">Fee Balance:</span>" & vbcrlf

    'Determine if the user can edit this section
     if lcl_userhaspermission_requestedit AND blnCanEdit then 
        response.write "<input type=""button"" name=""newFee"" id=""newFee"" style=""cursor:pointer"" value=""New Fee"" onclick=""location.href='fees/fees_new.asp?irequestid=" & iTrackID & "';"" />" & vbcrlf
        response.write "<input type=""button"" name=""newPayment"" id=""newPayment"" style=""cursor:pointer"" value=""New Payment"" onclick=""location.href='fees/payment_new.asp?irequestid=" & iTrackID & "';"" />" & vbcrlf
     end if

     response.write "<div id=""fees"" class=""divSection"">" & vbcrlf

     DisplayFeeBalance iTrackID

     response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if
 'END: Fee Balance ------------------------------------------------------------

 'BEGIN: Attachments ----------------------------------------------------------
  if lcl_orghasfeature_fileupload AND lcl_userhaspermission_fileupload then
     response.write "<p>" & vbcrlf
     response.write "  + <span id=""attachments"" class=""user_expand"">Attachments:</span>" & vbcrlf

     if lcl_userhaspermission_requestedit AND blnCanEdit then 
        subDisplayAttachments "E", iTrackID, sStatus, _
                              lcl_orghasfeature_actionline_secure_attachments, _
                              lcl_userhaspermission_actionline_secure_attachments, _
                              lcl_orghasfeature_actionline_display_attachments_to_public, _
                              lcl_userhaspermission_actionline_display_attachments_to_public
     else
        response.write "<div id=""file_upload"" style=""margin-top:5px;padding:5px;border:solid 1px #000000;background-color:#E0E0E0;"">" & vbcrlf

        subListAttachments iTrackID, _
                           lcl_orghasfeature_actionline_secure_attachments, _
                           lcl_userhaspermission_actionline_secure_attachments, _
                           "N", _
                           lcl_orghasfeature_actionline_display_attachments_to_public, _
                           lcl_userhaspermission_actionline_display_attachments_to_public

        response.write "</div>" & vbcrlf
     end if

     response.write "</p>" & vbcrlf
  end if
 'END: Attachments ------------------------------------------------------------

 'BEGIN: Request Activity Log -------------------------------------------------
  if NOT lcl_userhaspermission_actionline_hide_requestlog then
     lcl_display_substatus = ""

     if lcl_orghasfeature_action_line_substatus then
        lcl_display_substatus = " <em>(Sub-Status)</em>"
     end if

     response.write "<p>" & vbcrlf
     response.write "  + <span id=""requestActivityLog"" class=""user_expand"">Request Activity Log:</span>" & vbcrlf
     response.write "<div id=""log"">" & vbcrlf
     response.write "  <div id=""logHeaderRow"">" & vbcrlf
     response.write "    <table>" & vbcrlf
     response.write "      <tr><td><strong>User Name - Status" & lcl_display_substatus & " - Edit Date</strong></td></tr>" & vbcrlf
     response.write "    </table>" & vbcrlf
     response.write "  </div>" & vbcrlf
                       List_Comments(iTrackID)
     response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
  end if
 'END: Request Activity Log ---------------------------------------------------

 'BEGIN: Update Action Request ------------------------------------------------
  if blnCanEdit then
     response.write "<form name=""frmUpdate"" id=""frmUpdate"" action=""action_respond.asp"" method=""post"">" & vbcrlf
     response.write "  <input type=""hidden"" name=""TrackID"" id=""TrackID"" value="""                               & iTrackID     & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentStatus"" id=""currentStatus"" value="""                   & sStatus      & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentSubStatus"" id=""currentSubStatus"" value="""             & sSubStatusID & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""prevAssignedemployeeid"" id=""prevAssignedemployeeid"" value=""" & iemployeeid  & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentDepartmentID"" id=""currentDepartmentID"" value="""       & sDeptID      & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentDueDate"" id=""currentDueDate"" value="""                 & sDueDate     & """ />" & vbcrlf

     response.write "<p>" & vbcrlf
     response.write "  + <span id=""updateActionRequest"" class=""user_expand"">Update Action Request:</span>" & vbcrlf
     response.write "<div id=""update_form"" class=""divSection"">" & vbcrlf
     response.write "<table border=""0"" bordercolor=""#ff0000"">" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <table border=""0"">" & vbcrlf

     if session("orgid") = 15 then
        response.write "         <tr>" & vbcrlf
        response.write "             <td align=""right""><strong>Created by Employee:</strong></td>" & vbcrlf
        response.write "             <td>" & displaySubmitEmployee(isubmitid) & "</td>" & vbcrlf
        response.write "         </tr>" & vbcrlf
     end if

    'BEGIN: Reassignment Selection --------------------------------------------
     response.write "            <tr>" & vbcrlf
     response.write "                <td width=""30%"" align=""right""><strong>Assigned Employee:</strong></td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""assignedemployeeid"" id=""assignedemployeeid"">" & vbcrlf
     'response.write "                      <option value=""""></option>" & vbcrlf
                                           'DrawAdminUsers(iemployeeid)
                                           DrawAdminUsersAssignedHideDeleted(iemployeeid)
     response.write "                    </select>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
    'END: Reassignment Selection ----------------------------------------------

    'BEGIN: Status/Sub-Status -------------------------------------------------
    'Check to see if the sub-status needs to be set.
     lcl_onchange_status = ""

     if blnCanEdit AND lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then
        lcl_onchange_status = "changeSubStatus();"
     end if

     response.write "            <tr>" & vbcrlf
     response.write "                <td width=""20%"" align=""right""><strong>Status:</strong></td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""selStatus"" id=""selStatus"" onchange=""" & lcl_onchange_status & """>" & vbcrlf
     response.write "                      <option value=""SUBMITTED"""  & CheckSelected(sStatus,"SUBMITTED")  & ">SUBMITTED</option>"  & vbcrlf
     response.write "                      <option value=""INPROGRESS""" & CheckSelected(sStatus,"INPROGRESS") & ">INPROGRESS</option>" & vbcrlf
     response.write "                      <option value=""WAITING"""    & CheckSelected(sStatus,"WAITING")    & ">WAITING</option>"    & vbcrlf

     if lcl_userhaspermission_can_close_requests then
        response.write "                   <option value=""RESOLVED"""  & CheckSelected(sStatus,"RESOLVED")  & ">RESOLVED</option>" & vbcrlf
        response.write "                   <option value=""DISMISSED""" & CheckSelected(sStatus,"DISMISSED") & ">DISMISSED</option>" & vbcrlf
     end if

     response.write "                    </select>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr id=""sub_status_row"">" & vbcrlf

     if lcl_orghasfeature_action_line_substatus AND lcl_userhaspermission_action_line_substatus then
        response.write "                <td align=""right""><strong>Sub-Status:</strong></td>"  & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <select name=""selSubStatus"" id=""selSubStatus"">" & vbcrlf
        response.write "                      <option value=""0""></option>" & vbcrlf
        response.write "                    </select>" & vbcrlf
        response.write "                </td>" & vbcrlf
     else
        response.write "                <td align=""right"">&nbsp;</td>" & vbcrlf
        response.write "                <td><input type=""hidden"" name=""selSubStatus"" id=""selSubStatus"" value="""" /></td>" & vbcrlf
     end if

     response.write "            </tr>" & vbcrlf
    'END: Status/Sub-Status ---------------------------------------------------

    'BEGIN: Internal Note -----------------------------------------------------
     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""2"">" & vbcrlf
     response.write "                    <strong>Internal Note:</strong><br />" & vbcrlf
     response.write "                    <textarea name=""internal_comment"" id=""internal_comment"" rows=""5"" cols=""80""></textarea>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
    'END: Internal Note -------------------------------------------------------

    'BEGIN: Note to Citizen ---------------------------------------------------
     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""2"">" & vbcrlf
     response.write "                    <table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
     response.write "                      <tr>" & vbcrlf
     response.write "                          <td><strong>Note to Citizen:</strong></td>" & vbcrlf
     response.write "                          <td align=""right""><input type=""button"" value=""Add a Link"" class=""button"" onclick=""doPicker('frmUpdate.external_comment','Y','Y','Y','Y');"" /></td>" & vbcrlf
     response.write "                      </tr>" & vbcrlf
     response.write "                      <tr><td colspan=""2""><textarea name=""external_comment"" id=""external_comment"" rows=""5"" cols=""80""></textarea></td></tr>" & vbcrlf
     response.write "                    </table>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
    'END: Note to Citizen -----------------------------------------------------

    'BEGIN: Email Reminders ---------------------------------------------------
     if lcl_orghasfeature_actionline_emailreminders then
        response.write "         <tr>" & vbcrlf
        response.write "             <td colspan=""2"">" & vbcrlf
        response.write "                 <fieldset class=""fieldset"">" & vbcrlf
        response.write "                   <legend><strong>Email Reminders:&nbsp;&nbsp;</strong></legend>" & vbcrlf

        if lcl_userhaspermission_actionline_emailreminders then
           response.write "              <div align=""center"" class=""redText"">*** Email Reminder additions and changes are saved by clicking the ""UPDATE ACTION REQUEST"" button. ***</div><br />" & vbcrlf
           response.write "              <input type=""button"" name=""sAddReminder"" id=""sAddReminder"" value=""Add Reminder"" class=""button"" onclick=""addReminderRow();"" />" & vbcrlf
        end if

        response.write "                 <table id=""AddReminderTBL"" border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""margin-top:5px"" width=""100%"">" & vbcrlf
        response.write "                   <tr align=""left"" id=""addReminderRow_0"">" & vbcrlf
        response.write "                       <th>Send Reminder To</th>" & vbcrlf
        response.write "                       <th>Reminder<br />Send Date</th>" & vbcrlf
        response.write "                       <th>Additional Comments</th>" & vbcrlf
        response.write "                       <th align=""center"">Created By</th>" & vbcrlf

        if lcl_userhaspermission_actionline_emailreminders then
           response.write "                    <th align=""center"">Remove</th>" & vbcrlf
        end if

        response.write "                   </tr>" & vbcrlf

        sSQLr = "SELECT r.reminderid, "
        sSQLr = sSQLr & " r.orgid, "
        sSQLr = sSQLr & " r.action_autoid, "
        sSQLr = sSQLr & " r.send_date, "
        sSQLr = sSQLr & " r.sendto, "
        sSQLr = sSQLr & " r.comments, "
        sSQLr = sSQLr & " isnull(u.FirstName,'') AS SendToFirstName, "
        sSQLr = sSQLr & " isnull(u.LastName,'') AS SendToLastName, "
        sSQLr = sSQLr & " r.created_date, "
        sSQLr = sSQLr & " r.createdby, "
        sSQLr = sSQLr & " isnull(u2.FirstName,'') AS CreatedByFirstName, "
        sSQLr = sSQLr & " isnull(u2.LastName,'') AS CreatedByLastName, "
        sSQLr = sSQLr & " r.created_date "
        sSQLr = sSQLr & " FROM egov_action_reminders r "
        sSQLr = sSQLr &      " LEFT JOIN users u  ON r.sendto = u.userid "
        sSQLr = sSQLr &      " LEFT JOIN users u2 ON r.createdby = u2.userid "
        sSQLr = sSQLr & " WHERE r.orgid = " & session("orgid")
        sSQLr = sSQLr & " AND r.action_autoid = " & iTrackID
        sSQLr = sSQLr & " ORDER BY r.send_date"

        set oReminders = Server.CreateObject("ADODB.Recordset")
        oReminders.Open sSQLr, Application("DSN"), 1, 3

        iRowCount   = 0
        lcl_bgcolor = "#ffffff"

        if not oReminders.eof then
           do while not oReminders.eof
              iRowCount   = iRowCount + 1
              lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

              response.write "  <tr id=""addReminderRow_" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
   
             'ONLY users that have the feature/role assigned to their ID can modify email reminders.
             'if oReminders("createdby") <> session("userid") then
              if lcl_userhaspermission_actionline_emailreminders then
                 response.write "      <td>" & vbcrlf
                 response.write "          <input type=""hidden"" name=""reminderid_" & iRowCount & """ id=""reminderid_" & iRowCount & """ value=""" & oReminders("reminderid") & """ size=""10"" maxlength=""10"" />" & vbcrlf
                 response.write "          <select name=""reminderSendTo_" & iRowCount & """ id=""reminderSendTo_" & iRowCount & """ onchange=""clearMsg('reminderSendTo_" & iRowCount & "')"">" & vbcrlf
                 response.write "            <option value=""""></option>" & vbcrlf
                                             DrawAdminUsersNew oReminders("sendto"),"Y"
                 response.write "          </select>" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td>" & vbcrlf
                 response.write "          <input type=""text"" name=""reminderDate_" & iRowCount & """ id=""reminderDate_" & iRowCount & """ value=""" & oReminders("send_date") & """ size=""10"" maxlength=""10"" onchange=""clearMsg('reminderDateLOV_" & iRowCount & "')"" />" & vbcrlf
                 response.write "          <img src=""../images/calendar.gif"" id=""reminderDateLOV_" & iRowCount & """ border=""0"" style=""cursor:hand"" onclick=""clearMsg('reminderDateLOV_" & iRowCount & "');doCalendar('reminderDate_" & iRowCount & "');"" />" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td>" & vbcrlf
                 response.write "          <textarea name=""reminderComments_" & iRowCount & """ id=""reminderComments_" & iRowCount & """ rows=""2"" cols=""40""" & lcl_hide_textarea & ">" & oReminders("comments") & "</textarea>" & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td align=""center"">" & vbcrlf
                 response.write            oReminders("CreatedByFirstName") & " " & oReminders("CreatedByLastName") & "<br />" & vbcrlf
                 response.write            formatdatetime(oReminders("created_date"),vbshortdate) & vbcrlf
                 response.write "      </td>" & vbcrlf
                 response.write "      <td align=""center"">" & vbcrlf
                 response.write "          <input type=""checkbox"" name=""reminderRemove_" & iRowCount & """ id=""reminderRemove_" & iRowCount & """ value=""Y"" />" & vbcrlf
                 response.write "      </td>" & vbcrlf
              else
                 response.write "      <td>" & vbcrlf
                 response.write "          <input type=""hidden"" name=""reminderid_" & iRowCount & """ id=""reminderid_" & iRowCount & """ value=""" & oReminders("reminderid") & """ size=""10"" maxlength=""10"" />" & vbcrlf
                 response.write "          <input type=""hidden"" name=""reminderSendTo_" & iRowCount & """ id=""reminderSendTo_" & iRowCount & """ value=""" & oReminders("sendto") & """ size=""5"" maxlength=""10"" />" & vbcrlf
                 response.write            oReminders("SendToFirstName") & " " & oReminders("SendToLastName")
                 response.write "      </td>" & vbcrlf
                 response.write "      <td>" & vbcrlf
                 response.write "          <input type=""hidden"" name=""reminderDate_" & iRowCount & """ id=""reminderDate_" & iRowCount & """ value=""" & oReminders("send_date") & """ size=""10"" maxlength=""10"" />" & vbcrlf
                 response.write            oReminders("send_date")
                 response.write "      </td>" & vbcrlf
                 response.write "      <td>" & vbcrlf
                 response.write "          <input type=""hidden"" name=""reminderComments_" & iRowCount & """ id=""reminderComments_" & iRowCount & """ value=""" & oReminders("comments") & """ size=""10"" maxlength=""1000"" />" & vbcrlf
                 response.write            oReminders("comments")
                 response.write "      </td>" & vbcrlf
                 response.write "      <td align=""center"">" & vbcrlf
                 response.write            oReminders("CreatedByFirstName") & " " & oReminders("CreatedByLastName") & "<br />" & vbcrlf
                 response.write            formatdatetime(oReminders("created_date"),vbshortdate) & vbcrlf
                 response.write "      </td>" & vbcrlf
              end if

              response.write "  </tr>" & vbcrlf

              oReminders.movenext
           loop
        end if

        oReminders.close
        set oReminders = nothing

        response.write "                 </table>" & vbcrlf
        response.write "                 <input type=""hidden"" name=""sTotalReminderRows"" id=""sTotalReminderRows"" value=""" & iRowCount & """ />" & vbcrlf
        response.write "                 </fieldset>" & vbcrlf
        response.write "                 <br /><br />" & vbcrlf
        response.write "             </td>" & vbcrlf
        response.write "         </tr>" & vbcrlf
     end if
    'END: Email Reminders -----------------------------------------------------

     response.write "          </table>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"">" & vbcrlf
     'response.write "            <tr>" & vbcrlf
     'response.write "                <td>" & vbcrlf
     'response.write "                    <strong>Category:</strong><br />" & vbcrlf
     'response.write "                    <select name=""category_id"" id=""category_id"" onchange=""clearMsg('category_id');"">" & vbcrlf
     'response.write "                      getCategoryOptions iFormID" & vbcrlf
     'response.write "                    </select>" & vbcrlf
     'response.write "                    <br />" & vbcrlf
     'response.write "                </td>" & vbcrlf
     'response.write "            </tr>" & vbcrlf

    'BEGIN: Due Date ----------------------------------------------------------
     response.write "            <tr>" & vbcrlf
     response.write "                <td>" & vbcrlf

     if lcl_orghasfeature_actionline_maintain_duedate AND lcl_userhaspermission_actionline_maintain_duedate then
        response.write "<strong>Due Date:</strong><br />" & vbcrlf
        response.write "<input type=""text"" name=""due_date"" id=""due_date"" value=""" & sDueDate & """ size=""10"" maxlength=""10"" onchange=""clearMsg('due_date_pop')"" />" & vbcrlf
        response.write "<img src=""../images/calendar.gif"" id=""due_date_pop"" border=""0"" style=""cursor:hand"" onclick=""clearMsg('due_date_pop');doCalendar('due_date');"" />" & vbcrlf
     else
        response.write "<input type=""hidden"" name=""due_date"" id=""due_date"" value=""" & sDueDate & """ />" & vbcrlf
     end if

     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
    'END: Due Date ------------------------------------------------------------

    'BEGIN: Department --------------------------------------------------------
     response.write "            <tr>" & vbcrlf
     response.write "                <td>" & vbcrlf

     if lcl_orghasfeature_modify_actionline_department then
        response.write "<strong>Department:</strong><br />" & vbcrlf
        response.write "<select name=""deptid"" id=""deptid"" onchange=""clearMsg('deptid');"">" & vbcrlf
        'response.write "  <option value=""0""></option>" & vbcrlf
                          DrawDepartments sDeptID,"Y"
        response.write "</select>" & vbcrlf
     else
        response.write "<input type=""hidden"" name=""deptid"" id=""deptid"" value=""" & sDeptID & """ />" & vbcrlf
     end if

     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
    'END: Department ----------------------------------------------------------

     response.write "          </table>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf

    'BEGIN: Button Row --------------------------------------------------------
     response.write "  <tr>" & vbcrlf
     response.write "      <td colspan=""2"">" & vbcrlf
     response.write "          <input type=""button"" name=""sAction"" class=""button"" value=""UPDATE ACTION REQUEST"" onclick=""updateRequest()"" />" & vbcrlf

                              'if contact email exists, show option to send email
                               if sUserEmail <> "" AND NOT isnull(sUserEmail) then
                           			    response.write "<input type=""checkbox"" name=""sendemail"" id=""sendemail"" value=""yes"" /> Send email to Citizen?" & vbcrlf
                               end if

                              'Determine if user has permission to delete requests
                               if lcl_userhaspermission_actionline_delete then
                                  response.write "&nbsp; <input type=""button"" name=""sAction"" class=""button"" value=""DELETE ACTION REQUEST"" onclick=""deleteconfirm(" & iTrackID & ");"" />" & vbcrlf
                               end if

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
    'END: Button Row ----------------------------------------------------------

     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
     response.write "</form>" & vbcrlf
		end if
 'END: Update Action Request --------------------------------------------------

 'BEGIN: Send Notification ----------------------------------------------------
  if blnCanEdit then

    'Only display if the org has the "send_notification" feature turned on.
     if lcl_orghasfeature_send_notification then
        response.write "<form name=""frmNotify"" action=""action_respond.asp"" method=""post"">" & vbcrlf
        response.write "  <input type=""hidden"" name=""TrackID"" id=""TrackID"" value=""" & iTrackID & """ />" & vbcrlf
        response.write "  <input type=""hidden"" name=""prevnotifyuserid"" id=""prevnotifyuserid"" value=""" & inotifyuserid & """ />" & vbcrlf
        response.write "<p>" & vbcrlf
        response.write "  + <span id=""sendEmailNotifications"" class=""user_expand"">Send Email Notification:</span>" & vbcrlf
        response.write "<div id=""send_email"" class=""divSection"">" & vbcrlf
        response.write "<table border=""0"">" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td width=""20%"" align=""right""><strong>Notify User:</strong></td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <select name=""notifyuserid"" id=""notifyuserid"" onchange=""clearMsg('notifyuserid')"">" & vbcrlf
        response.write "            <option value=""""></option>" & vbcrlf
                                    DrawAdminUsersNew "","Y"
        response.write "          </select>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td width=""20%"" align=""right"" nowrap=""nowrap""><strong>Notify Department:</strong></td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <select name=""notifydeptid"" id=""notifydeptid"" onchange=""clearMsg('notifyuserid')"">" & vbcrlf
        response.write "            <option value=""""></option>" & vbcrlf
                                    DrawDepartments "","Y"
        response.write "          </select>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td colspan=""2"">" & vbcrlf
        response.write "          <strong>Additional Comments:</strong><br />" & vbcrlf
        response.write "          <textarea name=""notify_additional_comments"" id=""notify_additional_comments"" rows=""5"" cols=""80""></textarea>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbcrlf
        response.write "      <td colspan=""2"">" & vbcrlf
        response.write "          <input type=""submit"" name=""sAction"" id=""sAction"" class=""button"" value=""SEND NOTIFICATION"" onclick=""return checkDepartmentInactive();"" /> " & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "</table>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</p>" & vbcrlf
        response.write "</form>" & vbcrlf
     end if
  end if
 'END: Send Notification ------------------------------------------------------

 'Check to see if any form letters exist.  If so then the form letter buttons will function properly.
 'Otherwise, an error message is displayed.
  lcl_form_letters_exist = checkForFormLetters(iFormID)

 'BEGIN: Code Sections --------------------------------------------------------
  if lcl_orghasfeature_action_line_code_sections AND lcl_userhaspermission_action_line_code_sections then
     if lcl_form_letters_exist = "Y" then
        response.write "<p>" & vbcrlf
        response.write "  + <span id=""codeSections"" class=""user_expand""><a name=""#code_sections"">Code Sections:</a></span>&nbsp;" & vbcrlf

       'Determine if the user can edit the section
        if lcl_userhaspermission_requestedit AND blnCanEdit then 
           response.write "<input type=""button"" name=""editCodeSections"" id=""editCodeSections"" style=""cursor:pointer"" value=""Edit"" onclick=""location.href='corrections/correction_code_sections.asp?irequestid=" & iTrackID & "&status=" & sStatus & "&substatus=" & sSubStatusID & "';"" />" & vbcrlf
        end if

        response.write "  <div id=""code_sections"" class=""divSection"">" & vbcrlf

       'Check to see if any code_sections exist on the form(s) associated to this action line.
        sSQLce = "SELECT DISTINCT 'Y' FROM FormLetters "
        sSQLce = sSQLce & " LEFT OUTER JOIN egov_letter_to_form "
        sSQLce = sSQLce & " ON Formletters.FLid = egov_letter_to_form.letterid "
        sSQLce = sSQLce & " WHERE (orgid='" & session("orgid") & "') "
        sSQLce = sSQLce & " AND formid='"  & sDeptID          & "' "
        sSQLce = sSQLce & " OR FormLetters.blnAllMergeFields = 1 "
        sSQLce = sSQLce & " AND FLbody like ('%[*Code_Sections*]%') "

        set oCodeExist = Server.CreateObject("ADODB.Recordset")
        oCodeExist.Open sSQLce, Application("DSN") , 3, 1
	
        if not oCodeExist.EOF then
           lcl_code_sections_exist = "Y"
        else
           lcl_code_sections_exist = "N"
        end if 

        oCodeExist.Close
        set oCodeExist = nothing

        if lcl_code_sections_exist = "Y" then
           response.write "<p>" & vbcrlf
           response.write "<table border=""0"">" & vbcrlf

          'Retrieve all of the active code sections assigned to the action line request
           sSQLc = "SELECT cs.code_name "
           sSQLc = sSQLc & " FROM egov_actionline_code_sections cs, "
           sSQLc = sSQLc &      " egov_submitted_request_code_sections scs "
           sSQLc = sSQLc & " WHERE cs.action_code_id = scs.submitted_action_code_id "
           sSQLc = sSQLc & " AND cs.orgid = "                 & session("orgid")
           sSQLc = sSQLc & " AND scs.submitted_request_id = " & iTrackID
           'sSQLc = sSQLc & " AND cs.active_flag = 'Y' "
           'sSQLc = sSQLc & " AND cs.action_code_id in (" & REPLACE(REPLACE(sCodeSectionIDs,"(",""),")","") & ") "
           sSQLc = sSQLc & " ORDER BY UPPER(cs.code_name) "

           set oCodeSections = Server.CreateObject("ADODB.Recordset")
           oCodeSections.Open sSQLc, Application("DSN") , 3, 1
	
           if not oCodeSections.EOF then
              lcl_code_sections = ""
              do while not oCodeSections.eof
                 if lcl_code_sections = "" then
                    lcl_code_sections = oCodeSections("code_name")
                 else
                    lcl_code_sections = lcl_code_sections & "<br />" & oCodeSections("code_name")
                 end if
                 oCodeSections.movenext
              loop

              response.write "  <tr valign=""top"">" & vbcrlf
              response.write "      <td>" & lcl_code_sections & "</td>" & vbcrlf
              response.write "  </tr>" & vbcrlf
           else
              response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
           end if

           oCodeSections.close
           set oCodeSections = nothing

           response.write "</table>" & vbcrlf
           response.write "</p>" & vbcrlf

        end if  'end lcl_code_sections_exist

        response.write "</div>" & vbcrlf
        response.write "</p>" & vbcrlf
     end if  'end lcl_form_letters_exist
  end if  'end orgfeature
 'END: Code Sections ----------------------------------------------------------

 'BEGIN: Form Letters ---------------------------------------------------------
 'Determine if the user can edit the section
  if lcl_orghasfeature_form_letters AND lcl_userhaspermission_requestedit AND blnCanEdit then
     response.write "<form name=""frmLetter"" id=""frmLetter"" method=""post"">" & vbcrlf
     response.write "  <input type=""hidden"" name=""TrackID"" value=""" & iTrackID & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentStatus"" value=""" & sStatus & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentSubStatus"" value=""" & sSubStatusID & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentDepartmentID"" value=""" & sDeptID & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""currentDueDate"" value=""" & sDueDate & """ />" & vbcrlf
     'response.write "  <input type=""hidden"" name=""current_public_actionline_pdf"" value=""" & sPublicActionLinePDF & """ />" & vbcrlf
     '<input type="hidden" name="currentCategoryID" value=iFormID />

	    if lcl_form_letters_exist = "Y" then
        response.write "<p>" & vbcrlf
        response.write "  + <span id=""formLetter"" class=""user_expand"">Form Letter:</span>" & vbcrlf
        response.write "  <div id=""letter_form"" class=""divSection"">" & vbcrlf
        response.write "  <table>" & vbcrlf
        response.write "    <tr>" & vbcrlf
        response.write "        <td align=""right""><strong>Letter:</strong></td>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <select name=""selLetterId"" id=""selLetterId"">" & vbcrlf
                                      subListFLs iFormID
        response.write "            </select>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf
        response.write "  </table>" & vbcrlf
        response.write "  <table>" & vbcrlf
        response.write "    <tr>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <strong>Additional comments and signature:</strong><br />" & vbcrlf
        response.write "            <textarea name=""add_text"" id=""add_text"" rows=""5"" cols=""80""></textarea>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf
        response.write "    <tr>" & vbcrlf
        response.write "        <td>" & vbcrlf
        response.write "            <input type=""button"" class=""button"" value=""PREVIEW LETTER"" onclick=""return openFormLetter('PREVIEW');"" />" & vbcrlf
        response.write "            <input type=""submit"" class=""button"" value=""PRINT LETTER"" onclick=""return openFormLetter('PRINT');"" />" & vbcrlf

        if sUserEmail <> "" AND NOT isnull(sUserEmail) then
           response.write "            <input type=""submit"" class=""button"" value=""SEND EMAIL"" onclick=""return openFormLetter('EMAIL');"" />" & vbcrlf
        end if

        response.write "            <input type=""submit"" class=""button"" value=""Export to MS Word"" onclick=""return openFormLetter('WORDEXPORT');"" />" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf
        response.write "  </table>" & vbcrlf
        response.write "  </div>" & vbcrlf
        response.write "</p>" & vbcrlf
     else
        response.write "<input type=""hidden"" name=""selLetterId"" id=""selLetterId"" />" & vbcrlf
        response.write "<input type=""hidden"" name=""add_text"" id=""add_text"" />" & vbcrlf
     end if

     response.write "</form>" & vbcrlf

  end if
 'END: Form Letters -----------------------------------------------------------

 'BEGIN: PDF Forms ------------------------------------------------------------
  if lcl_orghasfeature_requestmergeforms AND lcl_userhaspermission_requestedit AND blnCanEdit then
    'Check to see if any PDFs exist.  PDFs are stored in the "PDFs" folder under the "unpublished documents" folder
     lcl_pdfs_exist = checkForPDFs(session("orgid"))

	    if lcl_pdfs_exist = "Y" then
        response.write "<p>" & vbcrlf
        response.write "  + <span id=""pdfForms"" class=""user_expand"">PDFs:</span>" & vbcrlf
        response.write "<div id=""pdf_form"" class=""divSection"">" & vbcrlf
        response.write "<table border=""0"">" & vbcrlf
        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <td>" & vbcrlf

       'Public-side PDF.  Only show if one has been associated to the request.
        if lcl_orghasfeature_requestmergeforms AND sPublicActionLinePDF <> "" then
           response.write "          <p>" & vbcrlf
           response.write "          <strong>Public-side PDF for New Action Line Requests:</strong>" & vbcrlf
           response.write "          <span class=""darkRedText"">" & sPublicActionLinePDF & "</span>&nbsp;" & vbcrlf
           response.write "          <input type=""button"" name=""viewPDF"" id=""viewPDF"" value=""View PDF"" class=""button"" onclick=""return openPDFtoView('" & iTrackID & "','','VIEW_PUBLIC_PDF')"" />" & vbcrlf
           response.write "          <input type=""hidden"" name=""public_actionline_pdf"" id=""public_actionline_pdf"" value=""" & sPublicActionLinePDF & """ size=""70"" maxlength=""1000"" onchange=""clearMsg('viewPDF')"" />&nbsp;" & vbcrlf
		   If  CLng(session("orgid")) = CLng(5) Then 
			   response.write "			 <input type=""button"" name=""viewtest"" value=""PDF Test"" class=""button"" onclick=""openTestPDF(" & iTrackID & ")"" />"
		   End If 
           response.write "          </p>" & vbcrlf
        end if

        response.write "          <p>" & vbcrlf
        response.write "          <strong>PDF:</strong>&nbsp;" & vbcrlf
        response.write "          <select name=""selectPDF"" id=""selectPDF"" onchange=""clearMsg('selectPDF');setupPDFButtons(this.value);"">" & vbcrlf
        response.write "            <option value=""""></option>" & vbcrlf
                                    displayPDFOptions(session("orgid"))
        response.write "          </select>" & vbcrlf
        response.write "          </p>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""button"" name=""previewPDF"" id=""previewPDF"" value=""PREVIEW PDF"" class=""button"" onclick=""clearMsg('selectPDF');return openPDFtoView('" & iTrackID & "',document.getElementById('selectPDF').value,'PREVIEW_PDF');"" />" & vbcrlf
        response.write "          <input type=""button"" name=""printPDF"" id=""printPDF"" value=""PRINT PDF"" class=""button"" onclick=""clearMsg('selectPDF');return openPDFtoView('" & iTrackID & "',document.getElementById('selectPDF').value,'PRINT_PDF');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "</table>" & vbcrlf
        response.write "</div>" & vbcrlf
        response.write "</p>" & vbcrlf

       'Disable the buttons until the user selects a PDF
        response.write "<script language=""javascript"">" & vbcrlf
        response.write "  setupPDFButtons(document.getElementById(""selectPDF"").value);" & vbcrlf
        response.write "</script>" & vbcrlf
     end if
  end if
 'END: PDFs -------------------------------------------------------------------

  response.write "</p>" & vbcrlf
'  response.write "</div>" & vbcrlf
  response.write "</blockquote>" & vbcrlf
 'END: Action Line Request List -----------------------------------------------

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
%>
<!-- #Include file="../admin_footer.asp" -->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
function Update_Action(iID)

  sEmail     = ""
  sCitizenID = ""

 'BEGIN: Update the employee assigned to the request. -------------------------
 'If the assigned employee has changed then notify the new assigned employee.
  if request("assignedemployeeid") <> "" then
     arrEmployee   = split(request("assignedemployeeid"),",")
     iEmployeeID   = arrEmployee(0)
     sEmployeeName = arrEmployee(1)

     sSQL = "SELECT * FROM egov_actionline_requests WHERE action_autoid=" & iID
     set oUpdate = Server.CreateObject("ADODB.Recordset")
     oUpdate.CursorLocation = 3
     oUpdate.Open sSQL, Application("DSN") , 1, 2

     if oUpdate("assignedemployeeid") <> iEmployeeID OR ISNULL(oUpdate("assignedemployeeid")) then
        oUpdate("assignedemployeeid") = iEmployeeID
        if iEmployeeID <> request("prevAssignedemployeeid") then
           sCommentLine = " This item has been re-assigned to " & sEmployeeName & "."
           '''AddCommentTaskComment sCommentLine,NULL,request("selStatus"), iID,Session("userid"),session("orgid")
           newCommentTask = 1

          'BODY CHANGES BASED ON orgid, 7 is the help desk
           If session("iorgid") <> "7" Then
              sEmailBody = "Action request --" & iID & replace(FormatDateTime(cdate(oUpdate("submit_date")),4),":","")
              sEmailBody = sEmailBody & "-- has been assigned to you on " & Now() & ".<br /><br />"
              sEmailBody = sEmailBody & "<strong>Click the following link to view this Action Line Request:</strong><br />"
              sEmailBody = sEmailBody & "<a href=""" & getEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iID & "&e=Y"">" & getEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iID & "&e=Y</a><br /><br />"
              sEmailBody = sEmailBody & "Please log into the admin web site and take the appropriate action.<br /><br />"
              sEmailBody = sEmailBody & "<strong>ACTION REQUEST DETAILS</strong><br />"
              sEmailBody = sEmailBody & fnPlainText(oUpdate("comment"))
           Else
              sEmailBody = "Action request --" & iID & replace(FormatDateTime(cdate(oUpdate("submit_date")),4),":","") & "-- has been assigned to you on " & Now() & ". Please log into the admin web site and take the appropriate action.<br /><br />"
              sEmailBody = sEmailBody & "<strong>HELPDESK TICKET DETAILS</strong>" & "<br /><br />"
              sEmailBody = sEmailBody & oUpdate("comment")
           End If

           blnSendEmail = True
        end if

        oUpdate.Update

       'SEND EMAIL NOTICE
        if blnSendEmail then
           sNoticeEmail = GetEmployeeEmail(iEmployeeID)

          'Check for a delegate
           getDelegateInfo iEmployeeID, lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

          'Send the email
           setupSendEmail "assign", iID, sEmailBody, sNoticeEmail, "", lcl_delegate_username, lcl_delegate_useremail
        end if

     end if

     oUpdate.close
     set oUpdate = nothing

     lcl_update_fields = 0

  end if
 'END: Update the employee assigned to the request. ---------------------------

 'BEGIN: Update ---------------------------------------------------------------
 '- Complete Date, Status, Sub-Status, Department, Category, Category Title, PDF, and Due Date
  sSQL = "UPDATE egov_actionline_requests SET "

	 if request("selStatus") = "RESOLVED" or request("selStatus") = "DISMISSED" then
     lcl_complete_date = now()

    'Setup the "complete_date" value
     if lcl_orghasfeature_actionline_reuse_completedate then

       'First we need to retrieve the last activity log entry status, sub-status, and log date
        getPreviousActionLineInfo iID, lcl_previous_status, lcl_previous_complete_date
        'getActivityLogInfo iRequestID, lcl_action_editdate, lcl_previous_status

       'Now we need to see if the status value matches the status value in the last log entry.
        if request("selStatus") = lcl_previous_status then
           if lcl_previous_complete_date <> "" then
              lcl_complete_date = lcl_previous_complete_date
           end if
        end if

     end if

     if lcl_complete_date <> "" then
        lcl_complete_date = "'" & lcl_complete_date & "'"
     end if

   		sSQL = sSQL & " complete_date = " & lcl_complete_date & ", "
  else
   		sSQL = sSQL & " complete_date = NULL, "
  end if

  if request("selSubStatus") <> "" then
     lcl_substatus = request("selSubStatus")
  else
     lcl_substatus = "NULL"
  end if

  if request("due_date") <> "" then
     sSQL = sSQL & " due_date = '" & request("due_date") & "', "
  else
     sSQL = sSQL & " due_date = NULL, "
  end if

		sSQL = sSQL & " status = '"       & request("selStatus") & "', "
  sSQL = sSQL & " sub_status_id = " & lcl_substatus        & ", "
  sSQL = sSQL & " groupid = "       & request("deptid")
  'sSQL = sSQL & " public_actionline_pdf = '" & request("public_actionline_pdf")         & "'"
  'sSQL = sSQL & " category_id = "            & request("category_id")                   & ", "
  'sSQL = sSQL & " category_title = '"        & getCategoryTitle(request("category_id")) & "' "
		sSQL = sSQL & " WHERE action_autoid = " & iID

  set oUpdateReq = Server.CreateObject("ADODB.Recordset")
  oUpdateReq.Open sSQL, Application("DSN"), 3, 1

  'oUpdateReq.close
  set oUpdateReq = nothing
 'END: Update -----------------------------------------------------------------

 'BEGIN: Send email to task originator (Contact) ------------------------------
	 if request("sendemail") = "yes" then
		   iTrackID = request("TrackID")

   		sSQLs = "SELECT * from egov_action_request_view where action_autoid=" & iID
   		Set oRS = Server.CreateObject("ADODB.Recordset")
   		oRS.Open sSQLs, Application("DSN"), 3, 1

   		lngTrackingNumber = iTrackID  & replace(FormatDateTime(cdate(oRS("submit_date")),4),":","")

   	'New HTML formatted response email
   		sMsgNew = sMsgNew & "<p>This automated message was sent by the " & getDefaultOrgValue( "orgname" ) & " " & GetFeatureName( "action line" ) & ". Do not reply to this message.  Please follow the instructions below "

   		if NOT lcl_orghasfeature_hide_email_actionline then
			     sMsgNew = sMsgNew & "or contact <strong>" & oRS("assigned_email") & "</strong>"
   		end if

   		sMsgNew = sMsgNew & " for inquiries regarding this email."
   		sMsgNew = sMsgNew & "</p>"
   		sMsgNew = sMsgNew & "<p>The status of your request has been updated, or new information has been added.</p>"
   		sMsgNew = sMsgNew & "<p><strong>TICKET STATUS:</strong> '" & UCASE(request("selStatus")) & "'</p>"

   		if Trim(request("external_comment")) <> "" then
			     'sMsgNew = sMsgNew & "<p><strong>LATEST ACTIVITY</strong>:<br /> " & replace(request("external_comment"),"'","`") & "</p>"
			     sMsgNew = sMsgNew & "<p><strong>LATEST ACTIVITY</strong>:<br /> " & request("external_comment") & "</p>" & vbcrlf
   		end if

   		sMsgNew = sMsgNew & "<p><strong>DETAILS:</strong></p>"
   		sMsgNew = sMsgNew & oRS("comment") & "<br /> " & vbcrlf
   		sMsgNew = sMsgNew & "<p><strong>FORM:</strong> " & oRS("action_formTitle") & "</p>"
   		sMsgNew = sMsgNew & "<p>"
   		sMsgNew = sMsgNew & "<strong>TRACKING NUMBER:</strong> " & lngTrackingNumber & "<br />"
   		sMsgNew = sMsgNew & "<strong>SUBMITTED</strong>: " & oRS("submit_date")

     if lcl_orghasfeature_actionline_maintain_duedate AND oRS("due_date") <> "" then
        sMsgNew = sMsgNew & "<br /><strong>DUE DATE</strong>: " & FormatDateTime(oRS("due_date"), vbshortdate)
     end if

   		sMsgNew = sMsgNew & "</p>"
   		sMsgNew = sMsgNew & "<p>To review the full ticket history, please follow the link below:</p>"
   		sMsgNew = sMsgNew & "<a href=""" & session("egovclientwebsiteurl") & "/action_request_lookup.asp?request_id=" & lngTrackingNumber & """>" & session("egovclientwebsiteurl") & "/action_request_lookup.asp?request_id=" & lngTrackingNumber & "</a>"
   		sMsgNew = sMsgNew & "<p>Make sure that the entire URL appears in your browser's address field.</p>"
   		sMsgNew = sMsgNew & "<p>Thank you for using our " & GetFeatureName( "action line" ) & " to better serve your needs.</p>"
   		'sMsgNew = sMsgNew & " " & vbcrlf 

   	'Get the to email address
     if NOT ISNULL(oRS("useremail")) then
  	   		sEmail     = Trim(oRS("useremail"))
        sCitizenID = oRS("userid")
   		else
			     sEmail     = ""
        sCitizenID = ""
   		end if

    'We have someone to send to (Citizen)
   		if sEmail <> "" then
        setupSendEmail "update", iID, sMsgNew, sEmail, "Y", "", ""
   		end if

     oRS.close
     set oRS = nothing
  end if
 'END: Send email to task originator (Contact) --------------------------------

 'BEGIN: Add Comments to Task (Activity Log) ----------------------------------
  if (request("selStatus") <> request("currentStatus")) _
  OR (request("selSubStatus") <> request("currentSubStatus")) _
  OR trim(request("internal_comment")) <> "" _
  OR trim(request("external_comment")) <> "" _
  OR request("deptid") <> request("currentDepartmentID") _
  OR request("due_date") <> request("currentDueDate") _
  OR newCommentTask = 1 then
  'OR trim(request("public_actionline_pdf")) <> trim(request("current_public_actionline_pdf")) _
  'OR (request("category_id") <> request("currentCategoryID")) _
     if request("internal_comment") <> "" and sCommentLine <> "" then
        intComment = request("internal_comment") & "<br />" & sCommentLine
     elseif request("internal_comment") <> "" then
        intComment = request("internal_comment")
     else
        intComment = sCommentLine
     end if

    'If the Department has been changed then log the change.
    'NOTE: The "deptid" selected and the "current deptid" cannot be NULL.
     if  (request("deptid") <> ""              AND NOT isnull(request("deptid"))) _
     AND (request("currentDepartmentID") <> "" AND NOT isnull(request("currentDepartmentID"))) then
        if CLng(request("deptid")) > 0 AND CLng(request("deptid")) <> CLng(request("currentDepartmentID")) then
           if intComment <> "" then
              intComment = intComment & "<br />" & vbcrlf
           end if

           intComment = intComment & "Department has been changed from """ & getDeptName(request("currentDepartmentID")) & """ " & vbcrlf
           intComment = intComment & "to """ & getDeptName(request("deptid")) & """" & vbcrlf
        end if
     end if

    'If the Public ActionLine PDf has been changed then log the change.
     'if request("public_actionline_pdf") <> request("current_public_actionline_pdf") then
         'if intComment <> "" then
             'intComment = intComment & "<br />" & vbcrlf
         'end if

         'intComment = intComment & "The Public-side PDF has been changed from """ & request("current_public_actionline_pdf") & """ " & vbcrlf
         'intComment = intComment & "to """ & request("public_actionline_pdf") & """" & vbcrlf
     'end if

    'If the Category has been changed then log the change.
     if request("category_id") <> "" AND NOT ISNULL(request("category_id")) then
        if CLng(request("category_id")) > 0 AND CLng(request("category_id")) <> CLng(request("currentCategoryId")) then
           if intComment <> "" then
              intComment = intComment & "<br />" & vbcrlf
           end if

           intComment = intComment & "Category has been changed from """ & getCategoryTitle(request("currentCategoryID")) & """ " & vbcrlf
           intComment = intComment & "to """ & getCategoryTitle(request("category_id")) & """" & vbcrlf
        end if
     end if

    'If the Due Date has been changed then log the change.
     if (request("due_date") <> ""       AND NOT isnull(request("due_date"))) _
     OR (request("currentDueDate") <> "" AND NOT isnull(request("currentDueDate"))) then
        if intComment <> "" then
           intComment = intComment & "<br />" & vbcrlf
        end if

        intComment = intComment & "The Due Date has been changed from """ & request("currentDueDate") & """ " & vbcrlf
        intComment = intComment & "to """ & request("due_date") & """" & vbcrlf
     end if

    'Check to see if the citizen's email is to be added to the external_comment.
     lcl_external_comment = request("external_comment")

     'if sEmail <> "" then
     '   lcl_admin_userid = session("userid")

       'Get the admin name and email
     '   lcl_admin_username  = getAdminName(lcl_admin_userid)
     '   lcl_admin_useremail = getUserEmail(lcl_admin_userid)

       'Get the submit date
     '   lcl_submit_date = ConvertDateTimetoTimeZone()

       'Set up the "Note to Citizen" email in Activity Log comment
     '   if lcl_external_comment <> "" then
     '      lcl_external_comment = lcl_external_comment & "<br />"
     '   end if

     '   lcl_external_comment = lcl_external_comment & lcl_admin_username & " sent a notification of this request to contact "
     '   lcl_external_comment = lcl_external_comment & sEmail
     '   lcl_external_comment = lcl_external_comment & " on " & lcl_submit_date

     'end if

  	  AddCommentTaskComment intComment, lcl_external_comment, request("selStatus"), iID, session("userid"), _
                           session("orgid"), request("selSubStatus"), sCitizenID, sEmail

    'BEGIN: Determine if any notifications are to be sent out -----------------
   		sSQL = "SELECT [tracking number] as tracking_number from egov_rpt_actionline where action_autoid=" & iID
   		set oGetTrackNum = Server.CreateObject("ADODB.Recordset")
   		oGetTrackNum.Open sSQL, Application("DSN"), 3, 1

   		lcl_tracking_number = oGetTrackNum("tracking_number")
     lcl_request_status  = ""
     lcl_form_id         = getActionLineFormID(session("orgid"), iID)

     if request("selStatus") <> "" then
        lcl_request_status = UCASE(request("selStatus"))
     end if

    'If the status has been set to "RESOLVED" or "DISMISSED" then send an email
     if (lcl_request_status = "RESOLVED" or lcl_request_status = "DISMISSED") AND lcl_request_status <> request("currentStatus") then
        setupAlertNotificationsEmail session("orgid"), iID, lcl_tracking_number, lcl_form_id, "request_closed"
     end if

    'Send an email to those simply wanting to know when a request has been updated (any update)
     setupAlertNotificationsEmail session("orgid"), iID, lcl_tracking_number, lcl_form_id, "request_updated"

     oGetTrackNum.close
     set oGetTrackNum = nothing

  end if
 'END: Add Comments to Task (Activity Log) ------------------------------------

 'BEGIN: Email Reminders ------------------------------------------------------
  if lcl_orghasfeature_actionline_emailreminders AND lcl_userhaspermission_actionline_emailreminders then
     lcl_total_reminders = request("sTotalReminderRows")
     i = 0

     if lcl_total_reminders > 0 then
        for i = 1 to lcl_total_reminders
          'Validate reminderid
           if dbready_number(request("reminderid_"&i)) then
              lcl_reminderid = request("reminderid_"&i)
           else
              lcl_reminderid = 0
           end if

          'Remove any reminders marked to be removed
           lcl_remove = request("reminderRemove_"&i)
           if lcl_remove = "Y" then
              sSQL = "DELETE FROM egov_action_reminders "
              sSQL = sSQL & " WHERE action_autoid = " & iID
              sSQL = sSQL & " AND orgid = " & session("orgid")
              sSQL = sSQL & " AND reminderid = " & lcl_reminderid
           else
             'Validate fields.
              if dbready_number(request("reminderSendTo_"&i)) then
                 lcl_sendto = request("reminderSendTo_"&i)
              else
                 lcl_sendto = "NULL"
              end if

              if dbready_date(request("reminderDate_"&i)) then
                 lcl_senddate = "'" & request("reminderDate_"&i) & "'"
              else
                 lcl_senddate = "NULL"
              end if

              if request("reminderComments_"&i) <> "" then
                 lcl_comments = "'" & dbsafe(request("reminderComments_"&i)) & "'"
              else
                 lcl_comments = "NULL"
              end if

             'Create reminder
              if lcl_reminderid = 0 then
                 sSQL = "INSERT INTO egov_action_reminders ("
                 sSQL = sSQL & "orgid, "
                 sSQL = sSQL & "action_autoid, "
                 sSQL = sSQL & "send_date, "
                 sSQL = sSQL & "sendto, "
                 sSQL = sSQL & "comments, "
                 sSQL = sSQL & "createdby, "
                 sSQL = sSQL & "created_date"
                 sSQL = sSQL & ") VALUES ("
                 sSQL = sSQL & session("orgid")  & ", "
                 sSQL = sSQL & iID               & ", "
                 sSQL = sSQL & lcl_senddate      & ", "
                 sSQL = sSQL & lcl_sendto        & ", "
                 sSQL = sSQL & lcl_comments      & ", "
                 sSQL = sSQL & session("userid") & ", "
                 sSQL = sSQL & "'" & Now()       & "'"
                 sSQL = sSQL & ")"
             'Update the reminder
              else
                 sSQL = "UPDATE egov_action_reminders SET "
                 sSQL = sSQL & " send_date = " & lcl_senddate & ", "
                 sSQL = sSQL & " sendto = "    & lcl_sendto   & ", "
                 sSQL = sSQL & " comments = "  & lcl_comments
                 sSQL = sSQL & " WHERE action_autoid = " & iID
                 sSQL = sSQL & " AND orgid = " & session("orgid")
                 sSQL = sSQL & " AND reminderid = " & lcl_reminderid
              end if
           end if

           if sSQL <> "" then
              set oReminderMaint = Server.CreateObject("ADODB.Recordset")
              oReminderMaint.Open sSQL, Application("DSN"), 3, 1
              'oReminderMaint.close
              set oReminderMaint = nothing
           end if
        next
     end if
  end if
 'END: Email Reminders --------------------------------------------------------

end function

'------------------------------------------------------------------------------
function List_Comments(iID)

'Check to see if the form has the option (egov_action_request_forms.action_form_resolved_status)
'to set the status to RESOLVED on creation.
 sSQLs = "SELECT f.action_form_resolved_status "
 sSQLs = sSQLs & " FROM egov_actionline_requests r, egov_action_request_forms f "
 sSQLs = sSQLs & " WHERE r.category_id = f.action_form_id "
 sSQLs = sSQLs & " AND r.action_autoid = " & iID

	set oResolve = Server.CreateObject("ADODB.Recordset")
	oResolve.Open sSQLs, Application("DSN"), 3, 1

 if not oResolve.eof then
    if oResolve("action_form_resolved_status") = "Y" then
       lcl_status = "RESOLVED"
    else
       lcl_status = "SUBMITTED"
    end if
 else
    lcl_status = "SUBMITTED"
 end if

 oResolve.close
 set oResolve = nothing

'Retrieve all of the entries for the activity log
	sSQL = "SELECT * "
	sSQL = sSQL & " FROM egov_action_responses egr "
	sSQL = sSQL & " LEFT OUTER JOIN egov_users ON egr.action_userid = egov_users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN users ON egr.action_userid = users.userid "
	sSQL = sSQL & " LEFT OUTER JOIN egov_actionline_requests_statuses AS es "
	sSQL = sSQL &               "ON egr.action_sub_status_id = es.action_status_id "
	sSQL = sSQL & " WHERE egr.action_autoid = " & iID
	sSQL = sSQL & " ORDER BY egr.action_editdate DESC"

	set oCommentList = Server.CreateObject("ADODB.Recordset")
	oCommentList.Open sSQL, Application("DSN"), 3, 1

 sBGColor = "#E0E0E0"
	
	if not oCommentList.EOF then
		  do while not oCommentList.eof
       lcl_substatus_name = oCommentList("status_name")

   	   if lcl_substatus_name <> "" then
			       lcl_substatus_name = " <em>(" & lcl_substatus_name & ")</em>"
  		   end if

		    	response.write "<div style=""background-color:" & sBGColor & """>" & vbcrlf
       response.write "<table>" & vbcrlf
    			response.write "  <tr>" & vbcrlf
       response.write "      <td>" & oCommentList("firstname") & " " & oCommentList("lastname") & " - " & UCASE(oCommentList("action_status")) & lcl_substatus_name & " - " & oCommentList("action_editdate") & "</td>" & vbcrlf
       response.write "  </tr>"

    			if oCommentList("action_externalcomment") <> "" then
          lcl_action_externalcomment = replace(oCommentList("action_externalcomment"),"default_novalue","")

         'Determine if an email was sent to the citizen.
          lcl_citizen_sentby_id      = "NULL"
          lcl_citizen_sentby_name    = "NULL"
          lcl_citizen_sentto_id      = "NULL"
          lcl_citizen_sentto_email   = "NULL"
          lcl_citizen_emailsent_date = "NULL"

          if oCommentList("citizen_sentto_email") <> "" then
             lcl_citizen_sentby_id      = oCommentList("citizen_sentby_id")
             lcl_citizen_sentby_name    = oCommentList("citizen_sentby_name")
             lcl_citizen_sentto_id      = oCommentList("citizen_sentto_id")
             lcl_citizen_sentto_email   = oCommentList("citizen_sentto_email")
             lcl_citizen_emailsent_date = oCommentList("citizen_emailsent_date")

            'Set up the "Note to Citizen" email in Activity Log comment
             if lcl_action_externalcomment <> "" then
                lcl_action_externalcomment = lcl_action_externalcomment & "<br />"
             end if

             lcl_action_externalcomment = lcl_action_externalcomment & lcl_citizen_sentby_name & " sent a notification of this request to contact "
             lcl_action_externalcomment = lcl_action_externalcomment & lcl_citizen_sentto_email
             lcl_action_externalcomment = lcl_action_externalcomment & " on " & lcl_citizen_emailsent_date

          end if

		      		response.write "  <tr>" & vbcrlf
          response.write "      <td>&nbsp;&nbsp;&nbsp;<strong>Note to Citizen: </strong><em>" & replace(lcl_action_externalcomment,chr(10),"<br />") & "</em></td>" & vbcrlf
          response.write "  </tr>"
    			end if

    			if oCommentList("action_citizen") <> "" then
          lcl_display_username = ""

          if trim(oCommentList("userfname")) <> "" then
             lcl_display_username = oCommentList("userfname")
          end if

          if trim(oCommentList("userlname")) <> "" then
             if lcl_display_username <> "" then
                lcl_display_username = lcl_display_username & " " & oCommentList("userlname")
             else
                lcl_display_username = oCommentList("userlname")
             end if
          end if

          if lcl_display_username <> "" then
             lcl_display_username = "<strong>" & lcl_display_username & ": </strong>"
          else
             lcl_display_username = "<strong>Citizen Comment: </strong>"
          end if

          lcl_display_comment = ""

          if oCommentList("action_citizen") <> "" then
             lcl_display_comment = formatActivityLogComment(oCommentList("action_citizen"))
          end if

				      response.write "  <tr>" & vbcrlf
          response.write "      <td>" & vbcrlf
          response.write "          &nbsp;&nbsp;&nbsp;" & lcl_display_username & vbcrlf
          'response.write "          <em>" & replace(oCommentList("action_citizen"),"default_novalue","") & "</em>" & vbcrlf
          response.write "          <em>" & lcl_display_comment & "</em>" & vbcrlf
          response.write "      </td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
    			end if

    			if oCommentList("action_internalcomment") <> "" then
  		    		response.write "  <tr>" & vbcrlf
          response.write "      <td>&nbsp;&nbsp;&nbsp;<strong>Internal Note: </strong><em>" & replace(replace(oCommentList("action_internalcomment"),"default_novalue",""),chr(10),"<br />") & "</em></td>" & vbcrlf
          response.write "  </tr>" & vbcrlf
    			end if

    			response.write "</table>" & vbcrlf
       response.write "</div>" & vbcrlf

    			oCommentList.MoveNext

       sBGColor = changeBGColor(sBGColor,"#E0E0E0","#FFFFFF")
    loop

			'DISPLAY SUBMIT DATE TIME AND USER
 			response.write "<div style=""background-color:" & sBGColor & ";"">" & vbcrlf
    response.write "<table>" & vbcrlf
  		response.write "  <tr><td>" & sSubmitName & " - " & UCASE(lcl_status) & " - " & datSubmitDate & "</td></tr>" & vbcrlf
		 	response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf
	else
 		'NO ACTIVITY FOR THIS REQUEST
  		response.write "<div style=""border-bottom:solid 1px #000000;background-color:#e0e0e0"">" & vbcrlf
    response.write "<table>" & vbcrlf
		  response.write "  <tr><td><font style=""color:red;font-size:12px;"">&nbsp;&nbsp;&nbsp;<em>No activity Reported.</em></td></tr>" & vbcrlf
  		response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf

		 'DISPLAY SUBMIT DATE TIME AND USER
  		response.write "<div style=""background-color:#FFFFFF;"">" & vbcrlf
    response.write "<table>" & vbcrlf
		  response.write "  <tr><td>" & sSubmitName & " - " & UCASE(lcl_status) & " - " & datSubmitDate & "</td></tr>" & vbcrlf
  		response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf
	End If

 oCommentList.close
 set oCommentList = nothing

End Function

'------------------------------------------------------------------------------
Function CheckSelected(sValue,sValue2)
	sReturnValue = ""

	If ucase(sValue) = sValue2 Then
		sReturnValue = " selected=""selected"""
	End If

	CheckSelected = sReturnValue
End Function

'------------------------------------------------------------------------------
function fnDisplayUserInfo(iID)

	if IsNull(iID) or iID="" then
	  	response.write "<div id=""contact_user"" class=""redText""><em>No information available for specified user.</em></div>" & vbcrlf
	else
 		'Get information for specified user
	 		sSQL = "SELECT * FROM egov_users WHERE userid = " & iID

 			set oUserInfo = Server.CreateObject("ADODB.Recordset")
	 		oUserInfo.Open sSQL, Application("DSN"), 3, 1
			
 			if not oUserInfo.eof then

   				sUserEmail = trim(oUserInfo("useremail"))

   				'response.write "<div id=""contact_user"" style=""margin-top:5px;border:solid 1px #000000;background-color:#E0E0E0;"">" & vbcrlf
   				response.write "<div id=""contact_user"" class=""divSection"">" & vbcrlf
       response.write "<table>" & vbcrlf
   				response.write "  <tr><td class=""label"" align=""right"">First Name:</td><td>"        & oUserInfo("userfname")                  & "</td></tr>" & vbcrlf
  	 			response.write "  <tr><td class=""label"" align=""right"">Last Name:</td><td>"         & oUserInfo("userlname")                  & "</td></tr>" & vbcrlf
  		 		response.write "  <tr><td class=""label"" align=""right"">Business Name:</td><td>"     & oUserInfo("userbusinessname")           & "</td></tr>" & vbcrlf
  			 	response.write "  <tr><td class=""label"" align=""right"">Email:</td><td>"             & oUserInfo("useremail")                  & "</td></tr>" & vbcrlf
  				 response.write "  <tr><td class=""label"" align=""right"">Daytime Phone:</td><td>"     & FormatPhone(oUserInfo("userhomephone")) & "</td></tr>" & vbcrlf
   				response.write "  <tr><td class=""label"" align=""right"">Fax:</td><td>"               & FormatPhone(oUserInfo("userfax"))       & "</td></tr>" & vbcrlf
   				response.write "  <tr><td class=""label"" align=""right"">Address:</td><td>"           & oUserInfo("useraddress")                & "</td></tr>" & vbcrlf
   				response.write "  <tr><td class=""label"" align=""right"">City:</td><td>"              & oUserInfo("usercity")                   & "</td></tr>" & vbcrlf
  	 			response.write "  <tr><td class=""label"" align=""right"">State / Province:</td><td>"  & oUserInfo("userstate")                  & "</td></tr>" & vbcrlf
  		 		response.write "  <tr><td class=""label"" align=""right"">Zip / Postal Code:</td><td>" & oUserInfo("userzip")                    & "</td></tr>" & vbcrlf
  			 	'response.write "  <tr><td class=""label"" align=""right"">Country:</td><td>" & oUser("usercountry") & "</td></tr>"
  				 response.write "  <tr><td class=""label"" align=""right"">Preferred Contact Method:</td><td>" & DisplayContactMethod(icontactmethodid) & "</td></tr>" & vbcrlf
   				response.write "</table></div>" & vbcrlf
			 else
	   			response.write "<div style=""display:none;"" id=""contact_user""><font class=""redText""><em>No information available for specified user.</em></font></div>" & vbcrlf
			 end if

    oUserInfo.close
    set oUserInfo = nothing

 end if

end function

'------------------------------------------------------------------------------
function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
end function

'------------------------------------------------------------------------------
Function fnPlaceEmailinQueue(sHost,sFromName,sFromEmail,sSendEmail,sSubject,iBodyFormat,sBodyMessage,iPriority,iErrorCode)
  
  Set oCmd = Server.CreateObject("ADODB.Command")
  
  With oCmd
    .ActiveConnection = Application("DSN")
    .CommandText = "AddEmailtoFailoverQueue"
    .CommandType = adCmdStoredProc
    .Parameters.Append oCmd.CreateParameter("Host", adVarChar , adParamInput, 50, sHost)
    .Parameters.Append oCmd.CreateParameter("FromName", adVarChar, adParamInput, 50, sFromName)
    .Parameters.Append oCmd.CreateParameter("FromEmail", adVarChar, adParamInput, 255, sFromEmail)
    .Parameters.Append oCmd.CreateParameter("SendEmail", adVarChar, adParamInput, 255, sSendEmail)
    .Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 1024, sSubject)
    .Parameters.Append oCmd.CreateParameter("BodyFormat", adInteger, adParamInput, 4, iBodyFormat)
    .Parameters.Append oCmd.CreateParameter("BodyMessage", adVarChar, adParamInput, 5000, sBodyMessage)
    .Parameters.Append oCmd.CreateParameter("Priority", adInteger, adParamInput, 4, iPriority)
    .Parameters.Append oCmd.CreateParameter("ErrorCode", adVarChar, adParamInput, 10, iErrorCode)
    .Execute
  End With
  
  Set oCmd = Nothing

End Function

'------------------------------------------------------------------------------
sub setupSendEmail(sEmailType, sActionAutoID, sEmailBody, sEmailToAddress, sendToCitizen, sDelegateUserName, sDelegateEmail)

 'Retrieve the data about the request
  sSQL = "SELECT * FROM egov_actionline_requests WHERE action_autoid=" & sActionAutoID
 	set oSendEmail = Server.CreateObject("ADODB.Recordset")
	 oSendEmail.Open sSQL, Application("DSN"), 3, 1

 'Set variables
  lcl_category_title      = ""
  lcl_category_subtitle   = ""
  lcl_contact_userid      = ""
  lcl_sendEmailToUser     = "N"
  lcl_sendEmailToDelegate = "N" 

  if not oSendEmail.eof then
     lcl_category_title = oSendEmail("category_title")
     lcl_contact_userid = oSendEmail("userid")
  end if

  oSendEmail.close
  set oSendEmail = nothing

 'Build the category sub-title.
 'If we are sending an email to the citizen then do NOT show the contact name in the subject line.
  if UCASE(sendToCitizen) <> "Y" then

    'Get the contact name if an ID exists on the request
     if lcl_contact_userid <> "" then
        sSQLe = "SELECT userfname, userlname FROM egov_users WHERE userid = " & lcl_contact_userid
        set oEmailContact = Server.CreateObject("ADODB.Recordset")
        oEmailContact.Open sSQLe, Application("DSN"), 3, 1

        if not oEmailContact.eof then
           lcl_contact_name = " (re: " & oEmailContact("userfname") & " " & oEmailContact("userlname") & ")"
        else
           lcl_contact_name = ""
        end if

        oEmailContact.close
        set oEmailContact = nothing

       'Set the category sub-title
        lcl_category_subtitle = lcl_contact_name

     end if
  else
     lcl_category_subtitle = " (re: Action Line Request)"
  end if

 'Build the email
  lcl_email_label   = ""
  lcl_email_subject = ""

 'Determine the type of email that is to be sent
 'example - Notification: Application Collection (re: Joe Smith)
  if UCASE(sEmailType) = "NOTIFY" then
     lcl_email_label   = "Notification"
  elseif UCASE(sEmailType) = "UPDATE" then
     lcl_email_label   = "Update"
  elseif UCASE(sEmailType) = "ASSIGN" then
     lcl_email_label   = "Assignment"
  end if

  if Clng(session("orgid")) <> Clng("7") then
    'From
     'lcl_email_from = session("sOrgName") & " (E-Gov Website) <webmaster@eclink.com>"
     lcl_email_from = session("sOrgName") & " (E-Gov Website) <noreply@eclink.com>"

    'Build the Email Subject
    '1. Check for the email type (label)
     if lcl_email_label <> "" then
        lcl_email_subject = lcl_email_label & ": "
     end if

    '2. Check for a category title (form name of the request.  i.e. Appliance Collection, Pothole, etc)
     lcl_email_subject = lcl_email_subject & lcl_category_title

    '3. Check for a category sub-title (i.e. Contact Name, "Action Line" text if sending to citizen, etc)
     lcl_email_subject = lcl_email_subject & lcl_category_subtitle

    'HTMLBody
     lcl_email_htmlbody = replace(sEmailBody,vbcrlf,"<br />")
     lcl_email_htmlbody = BuildHTMLMessage(lcl_email_htmlbody,"Y")

  else
     'lcl_email_from     = "EC Link HelpDesk <webmaster@eclink.com>"
     lcl_email_from     = "EC Link HelpDesk <noreply@eclink.com>"
     lcl_email_subject  = lcl_email_label & ": EC Link HelpDesk - HelpDesk Ticket"
     lcl_email_htmlbody = BuildHTMLMessage(sEmailBody,"Y")
  end if

 'Setup the SENDTO and check for a DELEGATE
  setupSendToAndDelegateEmails sEmailToAddress, sDelegateEmail, lcl_email_sendto, lcl_email_cc

 'Send the email
  sendEmail lcl_email_from, lcl_email_sendto, lcl_email_cc, lcl_email_subject, lcl_email_htmlbody, "", "Y"

end sub

'------------------------------------------------------------------------------
function DisplayContactMethod(iValue)

	sSQL = "SELECT * FROM egov_contactmethods WHERE rowid='" & iValue & "'"

	set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oMethods.EOF Then
  		iReturnValue = oMethods("contactdescription") 
	Else
		  iReturnValue = "NOT SPECIFIED"
	End If

 oMethods.close
	set oMethods = nothing
	
	DisplayContactMEthod = iReturnValue

end function

'------------------------------------------------------------------------------
function DisplaySubmitEmployee(iValue)

	sSQL = "SELECT * FROM Users WHERE userid='" & iValue & "'"

	Set oEmployee = Server.CreateObject("ADODB.Recordset")
	oEmployee.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oEmployee.EOF Then
  		iReturnValue = oEmployee("FirstName") & " " & oEmployee("LastName") 
	Else
		  iReturnValue = "NOT SPECIFIED"
	End If

 oEmployee.close
	set oEmployee = nothing
	
	DisplaySubmitEmployee = iReturnValue

end function

'------------------------------------------------------------------------------
Sub subListFLs(sDeptID)
	
'Retrieve all form letters that apply to this form
	sSQL = "SELECT * FROM FormLetters "
	sSQL = sSQL & " LEFT OUTER JOIN egov_letter_to_form "
	sSQL = sSQL & " ON Formletters.FLid = egov_letter_to_form.letterid "
	sSQL = sSQL & " WHERE (orgid='" & session("orgid") & "') "
	sSQL = sSQL & " AND (formid='"  & sDeptID          & "' "
	sSQL = sSQL & " OR FormLetters.blnAllMergeFields = 1) "
	sSQL = sSQL & " order by sequence "

	set oLetterList = Server.CreateObject("ADODB.Recordset")
	oLetterList.Open sSQL, Application("DSN") , 3, 1
	
	if not oLetterList.eof then
  		while not oLetterList.eof
  		  'Display only first 40 characters of title
    			if IsNull(oLetterList("FLtitle")) or oLetterList("FLtitle") = "" then
          iTitle = ""
    			elseif len(oLetterList("FLtitle")) > 50 then
          iTitle = left(oLetterList("FLtitle"),40) & "..."
     		else
          iTitle = oLetterList("FLtitle")
    			end if

    			response.write "<option value=" & oLetterList("FLid") & ">" & iTitle & "</option>"

    			oLetterList.MoveNext
    wend
	end if

 oLetterList.Close
	set oLetterList = nothing

end sub

'------------------------------------------------------------------------------
function GetEmployeeEmail(iValue)

	sSQL = "SELECT * FROM Users WHERE userid='" & iValue & "'"

	set oEmpEmail = Server.CreateObject("ADODB.Recordset")
	oEmpEmail.Open sSQL, Application("DSN") , 3, 1
	
	if not oEmpEmail.eof then
		  iReturnValue = oEmpEmail("Email")
	else
		  iReturnValue = "NOT SPECIFIED"
	end if

 oEmpEmail.close
	set oEmpEmail = nothing
	
	GetEmployeeEmail = iReturnValue
	
end function

'------------------------------------------------------------------------------
Function FormatPhone( Number )
  If Len(Number) = 10 Then
     FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
  Else
     FormatPhone = Number
  End If
End Function

'------------------------------------------------------------------------------
Sub SubDrawIssueLocationInformation( iRequestID, lcl_blnCanEdit, sHideIssueLocAddInfo )
	Dim sSql, oIssueLocation

	sSQL = "SELECT rowid, "
 sSQL = sSQL & " actionrequestresponseid, "
 sSQL = sSQL & " streetnumber, "
 sSQL = sSQL & " streetprefix, "
 sSQL = sSQL & " streetaddress, "
 sSQL = sSQL & " streetsuffix, "
 sSQL = sSQL & " streetdirection, "
 sSQL = sSQL & " streetunit, "
 sSQL = sSQL & " sortstreetname, "
 sSQL = sSQL & " city, "
 sSQL = sSQL & " state, "
 sSQL = sSQL & " zip, "
 sSQL = sSQL & " comments, "
 sSQL = sSQL & " latitude, "
 sSQL = sSQL & " longitude, "
 sSQL = sSQL & " validstreet, "
 sSQL = sSQL & " county, "
 sSQL = sSQL & " parcelidnumber, "
 sSQL = sSQL & " excludefromactionline, "
 sSQL = sSQL & " residenttype, "
 sSQL = sSQL & " listedowner, "
 sSQL = sSQL & " legaldescription "
 sSQL = sSQL & " FROM egov_action_response_issue_location "
 sSQL = sSQL & " WHERE actionrequestresponseid = " & iRequestID
 sSQL = sSQL & " AND excludefromactionline = 0 "

	set oIssueLocation = Server.CreateObject("ADODB.Recordset")
	oIssueLocation.Open sSQL, Application("DSN"), 3, 1
	
	If NOT oIssueLocation.EOF Then
		'sHasValue = trim(oIssueLocation("streetnumber") & oIssueLocation("streetaddress") & oIssueLocation("city") &  oIssueLocation("state") & oIssueLocation("zip") & oIssueLocation("comments"))

		'If sHasValue <>  "" Then
   response.write "<p>" & vbcrlf
			'response.write "  + <span id=""user_expand"" onclick=""toggleDisplay('issue_location');"">" & sIssueName & "</span>" & vbcrlf
			response.write "  + <span id=""issueLocation"" class=""user_expand"">" & sIssueName & "</span>" & vbcrlf

  'Determine if the user can edit this section
   if lcl_userhaspermission_requestedit AND lcl_blnCanEdit then
      response.write "<input type=""button"" name=""editIssueLocation"" id=""editIssueLocation"" style=""cursor:pointer"" value=""Edit"" onclick=""location.href='corrections/correction_issue_location.asp?requestid=" & iTrackID & "&status=" & sStatus & "&substatus=" & sSubStatusID & "';"" />" & vbcrlf
      'response.write "<input type=""button"" name=""searchissueduplicates"" value=""Search for Duplicates"" class=""button"" />"
  			 'response.write "- [<a href=""corrections/correction_issue_location.asp?requestid=" & iTrackID & "&status=" & sStatus & "&substatus=" & sSubStatusID & """>Edit</a>]" & vbcrlf
   end if 

			response.write "<div id=""issue_location"" class=""divSection"">" & vbcrlf

  	sSQLc = "SELECT f.action_form_display_issue "
   sSQLc = sSQLc & " FROM egov_action_request_forms f, egov_actionline_requests r "
   sSQLc = sSQLc & " WHERE r.category_id = f.action_form_id "
   sSQLc = sSQLc & " AND r.action_autoid = " & iRequestID

  	Set rs = Server.CreateObject("ADODB.Recordset")
  	rs.Open sSQLc, Application("DSN"), 3, 1

     if oIssueLocation("validstreet") <> "Y" AND rs("action_form_display_issue") then
        lcl_valid_street = "<font class=""redText"">&nbsp;*</em></small></font>" & vbcrlf
     else
        lcl_valid_street = ""
     end if

			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
   response.write "  <tr valign=""top"">" & vbcrlf
   response.write "      <td>" & vbcrlf
   response.write "          <table>" & vbcrlf
			response.write "            <tr>" & vbcrlf
   response.write "                <td class=""label"" align=""right"">Street:</td>" & vbcrlf

  'Build the street name
   lcl_street_name = buildStreetAddress(oIssueLocation("streetnumber"), oIssueLocation("streetprefix"), oIssueLocation("streetaddress"), oIssueLocation("streetsuffix"), oIssueLocation("streetdirection"))

   response.write "                <td>" & lcl_street_name & lcl_valid_street & "</td>" & vbcrlf
'   response.write "                <td>" & oIssueLocation("streetnumber") & " " & oIssueLocation("streetaddress") & lcl_valid_street & "</td>" & vbcrlf
   response.write "            </tr>" & vbcrlf
			response.write "            <tr>" & vbcrlf
   response.write "                <td class=""label"" align=""right"">City:</td>" & vbcrlf
   response.write "                <td>" & oIssueLocation("city") & "</td>" & vbcrlf
   response.write "            </tr>" & vbcrlf
			response.write "            <tr>" & vbcrlf
   response.write "                <td class=""label"" align=""right"">State:</td>" & vbcrlf
   response.write "                <td>" & oIssueLocation("state") & "</td>" & vbcrlf
   response.write "            </tr>" & vbcrlf
			response.write "            <tr>" & vbcrlf
   response.write "                <td class=""label"" align=""right"">Zip:</td>" & vbcrlf
   response.write "                <td>" & oIssueLocation("zip") & "</td>" & vbcrlf
   response.write "            </tr>" & vbcrlf

   if oIssueLocation("streetunit") <> "" then
   			response.write "            <tr>" & vbcrlf
      response.write "                <td class=""label"" align=""right"">Unit:</td>" & vbcrlf
      response.write "                <td>" & oIssueLocation("streetunit") & "</td>" & vbcrlf
      response.write "            </tr>" & vbcrlf
   end if

'   if GetOrgDisplayWithId(session("orgid"),26,1) = "" then
'      lcl_county_label = "County"
'   else
'      lcl_county_label = GetOrgDisplayWithId(session("orgid"),26,1)
'   end if

   lcl_display_id = 0

  'Build the custom label
   lcl_display_id   = GetDisplayId("address grouping field")
   lcl_county_label = GetOrgDisplayWithId(session("orgid"),lcl_display_id,true)

   if lcl_county_label = "" then
      lcl_county_label = GetDisplayName(lcl_display_id)
   end if

			response.write "            <tr>" & vbcrlf
   response.write "                <td class=""label"" align=""right"">" & lcl_county_label & ":</td>" & vbcrlf
   response.write "                <td>" & oIssueLocation("county") & "</td>" & vbcrlf
   response.write "            </tr>" & vbcrlf

   if lcl_orghasfeature_parcel_id then
   			response.write "            <tr>" & vbcrlf
      response.write "                <td class=""label"" align=""right"">Parcel ID:</td>" & vbcrlf
      response.write "                <td>" & oIssueLocation("parcelidnumber") & "</td>" & vbcrlf
      response.write "            </tr>" & vbcrlf
   			response.write "            <tr>" & vbcrlf
      response.write "                <td class=""label"" align=""right"">Listed Owner:</td>" & vbcrlf
      response.write "                <td>" & oIssueLocation("listedowner") & "</td>" & vbcrlf
      response.write "            </tr>" & vbcrlf
   			response.write "            <tr>" & vbcrlf
      response.write "                <td class=""label"" align=""right"">Legal Description:</td>" & vbcrlf
      response.write "                <td>" & oIssueLocation("legaldescription") & "</td>" & vbcrlf
      response.write "            </tr>" & vbcrlf
   end if

   if not sHideIssueLocAddInfo then
   			response.write "            <tr>" & vbcrlf
      response.write "                <td class=""label"" align=""right"">Comments:</td>" & vbcrlf
      response.write "                <td>" & oIssueLocation("comments") & "</td>" & vbcrlf
      response.write "            </tr>" & vbcrlf
   end if

   response.write "          </table>" & vbcrlf
   response.write "      </td>" & vbcrlf
   response.write "      <td align=""right"">" & vbcrlf

   if oIssueLocation("validstreet") <> "Y" then
      response.write "<font class=""redText"">* <small><em>= Non-Listed Street Address</em></small></font>" & vbcrlf
   else
      response.write "&nbsp;"
   end if

   response.write "      </td>" & vbcrlf
   response.write "  </tr>" & vbcrlf
			response.write "</table>" & vbcrlf
			
			response.write "</div>" & vbcrlf
			response.write "</p>" & vbcrlf
		'End If

	End If

	oIssueLocation.Close
	set oIssueLocation = nothing 

End Sub

'------------------------------------------------------------------------------
Function getEgovWebsiteURL()
	Dim sSql, oURL
	
	sSql = "SELECT OrgEgovWebsiteURL FROM organizations WHERE orgid = " & session("orgid")

	Set oURL = Server.CreateObject("ADODB.Recordset")
	oURL.Open sSql, Application("DSN"), 0, 1
	
	getEgovWebsiteURL = oURL("OrgEgovWebsiteURL")

	oURL.close
	Set oURL = Nothing

End Function 

'------------------------------------------------------------------------------
Sub DisplayFormFieldsandAnswers(irequestid,blnIsAdmin)

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM egov_submitted_request_fields "
	
	if blnIsAdmin then
  		sSQL = sSQL & " INNER JOIN egov_submitted_request_field_responses "
		  sSQL = sSQL &         " ON egov_submitted_request_fields.submitted_request_field_id = egov_submitted_request_field_responses.submitted_request_field_id "
  		sSQL = sSQL & " WHERE submitted_request_id='" & irequestid & "' "
		  sSQL = sSQL & " AND (submitted_request_field_isinternal = 1) "
  		sSQL = sSQL & " ORDER BY submitted_request_field_sequence, egov_submitted_request_fields.submitted_request_field_id"
	else
		  sSQL = sSQL & " WHERE submitted_request_id='" & irequestid & "' "
  		sSQL = sSQL & " AND (submitted_request_field_isinternal = 0 "
		  sSQL = sSQL & " OR submitted_request_field_isinternal IS NULL) "
  		sSQL = sSQL & " ORDER BY submitted_request_field_sequence"
	end if

	set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN"), 3, 1
	
	' IF QUESTIONS DISPLAY THEM
	If NOT oQuestions.EOF Then
		
		blnStart = True
		iCurrentID = 0

		' LOOP THRU QUESTIONS AND ANSWERS DISPLAYING THEM
		Do While NOT oQuestions.EOF 

			If CLng(iCurrentID) <> CLng(oQuestions("submitted_request_field_id")) or blnStart = True Then
  				response.write  "<p><strong>" & oQuestions("submitted_request_field_prompt") & "</strong><br />"
		  		blnStart = False
			Else
				  response.write ", "
			End If
			
  'If the input field is a radio/checkbox type field then check to see if no value was initially selected by
  'the user.  If that is the case then there is a hidden field with each radio/checkbox group containing a value of "DEFAULT_NOVALUE".
  'If the field has this value then simply display nothing on the screen.  Otherwise, display the value.
   if UCASE(Trim(oQuestions("submitted_request_field_response"))) = "DEFAULT_NOVALUE" then
      response.write "&nbsp;"
   else
   			response.write replace(trim(oQuestions("submitted_request_field_response")),chr(10),"<br />")
   end if

			iCurrentID = oQuestions("submitted_request_field_id")
			
			If CLng(iCurrentID) <> CLng(oQuestions("submitted_request_field_id")) or blnStart = True Then
			  	response.write  "</p>" & vbcrlf & vbcrlf
			End If
			
			oQuestions.MoveNext
		Loop

	Else

		response.write "<p>There are no values entered. Click <strong>Edit</strong> to update any available fields.</P>"

	End If

 oQuestions.close
	set oQuestions = nothing 

end sub

'------------------------------------------------------------------------------
sub SubDisplayPDFForms(iOrgID)

 dim sSQL, lcl_orgid

 if iOrgID <> "" then
    lcl_orgid = clng(iOrgID)
 end if

	'GET PDF FORMS FOR THIS ORGANIZATION
	sSQL = "SELECT pdfid, "
 sSQL = sSQL & " pdf_name, "
 sSQL = sSQL & " pdf_description, "
 sSQL = sSQL & " date_added, "
 sSQL = sSQL & " orgid, "
 sSQL = sSQL & " adminuserid "
 sSQL = sSQL & " FROM egov_action_request_pdfforms "
 sSQL = sSQL & " WHERE orgid = " & lcl_orgid
 sSQL = sSQL & " ORDER BY isdefault, pdf_name"

	set oPDFList = Server.CreateObject("ADODB.Recordset")
	oPDFList.Open sSQL,Application("DSN"),1,3

	response.write "<p>" & vbcrlf
 response.write "  + <span id=""user_expand"" onclick=""toggleDisplay('pdfforms');"">Request PDF Merge Forms</span> " & vbcrlf
	response.write "<div id=""pdfforms"" class=""divSection"">" & vbcrlf
	
	response.write "<strong>PDF Form:</strong> <select>" & vbcrlf
	response.write "<option selected >Select PDF Output Form...</option>" & vbcrlf

	' IF THERE ARE PDF FORMS DISPLAY THEM	
	If NOT oPDFList.EOF Then

		' LIST ALL PDF FORMS FOUND
		Do While NOT oPDFList.EOF

			sPDFDesc = oPDFList("pdf_description")

			response.write "<option >" & oPDFList("pdf_name") & "</option>" & vbcrlf
		
			' NEXT ROW
			oPDFList.MoveNext
		
		Loop

		' CLOSE AND DESTROY RECORDSET
		oPDFList.Close
		Set oPDFList = NOTHING
	
	End If

	response.write "</select>" & vbcrlf

	response.write "&nbsp;&nbsp;<input onclick=""fnDisplayPDF();"" type=button value=""Merge Request with PDF Form"">" & vbcrlf
	'response.write "<br />"
	'response.write "<span id=pdf_description>" & sPDFDesc & "</span>"

end sub

'----------------------------------------------------------------------
Function DisplayFeeBalance(iRequestID)

	'SELECT ALL FEES
	sSQL = "SELECT fees.* FROM egov_action_fees fees WHERE fees.action_autoid = " & iRequestID
	Set oFees = Server.CreateObject("ADODB.RecordSet")
	oFees.Open sSQL, Application("DSN"), 3, 1

	'SELECT ALL PAYMENTS
	sSQL = "SELECT paymentdate,paymentamount FROM egov_paymentinformation INNER JOIN egov_payments ON egov_payments.paymentinfoid = egov_paymentinformation.paymentinfoid WHERE orgid = " & session("orgid") & " AND payment_information LIKE '%custom_trackingnumber: " & iRequestID & "%' AND paymentstatus = 'COMPLETED'"
	Set oPayments = Server.CreateObject("ADODB.RecordSet")
	oPayments.Open sSQL, Application("DSN"), 3, 1

	'CREATE TEMP TABLE
	Set oFeeBalanceTemp = Server.CreateObject("ADODB.Connection")
	oFeeBalanceTemp.Open Application("DSN")
	sSQL = "CREATE TABLE #FeeBalance (dateadded datetime,userid int,Description varchar(50),Amount int)"
	oFeeBalanceTemp.Execute(sSQL)

	'INSERT FEES
	Do While Not oFees.EOF
		sSQL = "INSERT INTO #FeeBalance (dateadded,userid,description,amount) VALUES('" & oFees("DateAdded") & "'," & oFees("FeeCalculatedByID") & ",'" & oFees("FeeDescription") & "'," & oFees("FeeAmount") & ")"
		oFeeBalanceTemp.Execute(sSQL)
		oFees.MoveNext
	loop
	oFees.Close
	Set oFees = Nothing

	''INSERT PAYMENTS
	'FOR TESTING
	Do While Not oPayments.EOF
		sSQL = "INSERT INTO #FeeBalance (dateadded,userid,description,amount) VALUES('" & oPayments("paymentdate") & "',0,'Payment'," & oPayments("PaymentAmount") & ")"
		oFeeBalanceTemp.Execute(sSQL)
		oPayments.MoveNext
	loop
	oPayments.Close
	Set oPayments = Nothing

	'READ OUT OF THAT TEMP TABLE IN DATE ORDER
	sSQL = "SELECT fees.*,firstname,lastname FROM #FeeBalance fees LEFT JOIN users on users.userid=fees.userid ORDER BY dateadded"
	Set oFeeBalance = Server.CreateObject("ADODB.RecordSet")
	oFeeBalance.Open sSQL, oFeeBalanceTemp, 3, 1

	if not oFeeBalance.EOF then
		response.write "<div>" & vbcrlf
  response.write "<table cellpadding=""4"" cellspacing=""0"">" & vbcrlf
		response.write "  <tr style=""border-bottom:solid 1px #000000;"">" & vbcrlf
  response.write "      <th>Date Added</th>" & vbcrlf
  response.write "      <th>Added By</th>" & vbcrlf
		response.write "      <th>Type</th>" & vbcrlf
		response.write "      <th>Amount</th>" & vbcrlf
		response.write "  </tr>" & vbcrlf
	
		sBGColor = "#FFFFFF"
		while not oFeeBalance.eof
  		'Display Fee row
		   response.write "  <tr bgcolor=""" & sBGColor & """>" & vbcrlf
  			response.write "      <td>" & oFeeBalance("dateadded") & "</td>" & vbcrlf

  			if oFeeBalance("userid") = 0 then
		    		response.write "      <td>Citizen</td>" & vbcrlf
  			elseif oFeeBalance("userid") = -1 then
		    		response.write "      <td>System</td>" & vbcrlf
  			else
		    		response.write "      <td>" & oFeeBalance("firstname") & " " &  oFeeBalance("lastname") & "</td>" & vbcrlf
  			end if

  			response.write "      <td>" & oFeeBalance("Description") & "</td>" & vbcrlf
		  	response.write "      <td align=""right"">" & FormatCurrency(oFeeBalance("Amount")) & "</td>" & vbcrlf
  			response.write "  </tr>" & vbcrlf
			
     sBGColor = changeBGColor(sBGColor,"#FFFFFF","#E0E0E0")

  			oFeeBalance.movenext
  wend
		oFeeBalance.Close

	'Get Balance
		sSQL = "SELECT SUM(Amount) AS Balance FROM #FeeBalance"
		set oFeeBalance = Server.CreateObject("ADODB.RecordSet")
		oFeeBalance.Open sSQL, oFeeBalanceTemp, 3, 1
		if not oFeeBalance.EOF then Response.Write "<tr bgcolor=" & sBGColor & "><th colspan=""3"" align=""right"">Balance</th><td align=right>" & FormatCurrency(oFeeBalance("BALANCE")) & "</td></tr>"
	
		response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
	else
		response.write "<p>There are no fees entered. Click <strong>New Fee</strong> to add fees.</p>"
	end if

	oFeeBalance.Close
	set oFeeBalance = nothing

	'DESTROY THE TEMP TABLE
	sSQL = "DROP TABLE #FeeBalance"
	oFeeBalanceTemp.Execute(sSQL)
	oFeeBalanceTemp.Close
	set oFeeBalanceTemp = nothing

end function

'------------------------------------------------------------------------------
Function getLocalTimeZone()
	If strLocalTime = ""  then
		sSQL = "SELECT T.TZName	"
  sSQL = sSQL & " FROM organizations O, timezones T "
  sSQL = sSQL & " WHERE O.orgtimezoneid = T.timezoneid "
  sSQL = sSQL & " AND O.orgid = " & session("orgid")

		set oZone = Server.CreateObject("ADODB.RecordSet")
		oZone.Open sSQL, Application("DSN"), 3, 1

	 strLocalTime = oZone("TZName")

  oZone.close
		set oZone = nothing	

	End If

	getLocalTimeZone=strLocalTime 

End Function

'------------------------------------------------------------------------------
function checkForFormletters(p_formid)
  lcl_exists = "N"

  if p_formid <> "" then
    'This query checks to see if any form letters exist.
     sSQL = "SELECT 'Y' FROM FormLetters "
     sSQL = sSQL & " LEFT OUTER JOIN egov_letter_to_form ON Formletters.FLid = egov_letter_to_form.letterid "
     sSQL = sSQL & " WHERE (orgid='" &  session("orgid") & "') "
     sSQL = sSQL & " AND (formid='" & p_formid & "' "
     sSQL = sSQL & " OR FormLetters.blnAllMergeFields = 1) "
     sSQL = sSQL & " ORDER BY sequence "

     set oLetterExist = Server.CreateObject("ADODB.Recordset")
     oLetterExist.Open sSQL, Application("DSN") , 3, 1

     if not oLetterExist.eof then
        lcl_exists = "Y"
     else
        lcl_exists = "N"
     end if

     oLetterExist.close
     set oLetterExist = nothing
  end if

  checkForFormletters = lcl_exists

end function

'------------------------------------------------------------------------------
function checkForPDFs(iOrgID)
  dim lcl_exists, lcl_orgid

  lcl_exists = "N"
  lcl_orgid  = 0

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

 'This query checks to see if any PDFs exist.
  sSQL = "SELECT count(documentid) AS totalpdfs "
  sSQL = sSQL & " FROM documents "
  sSQL = sSQL & " WHERE orgid = " & lcl_orgid
  sSQL = sSQL & " AND upper(documenturl) like ('%UNPUBLISHED_DOCUMENTS%/PDFS/%')"
  sSQL = sSQL & " AND RIGHT(UPPER(documentURL),4) = '.PDF' "

  set oPDFExist = Server.CreateObject("ADODB.Recordset")
  oPDFExist.Open sSQL, Application("DSN") , 3, 1

  if oPDFExist("totalpdfs") > 0 then
     lcl_exists = "Y"
  end if

  oPDFExist.close
  set oPDFExist = nothing

  checkForPDFs = lcl_exists

end function

'------------------------------------------------------------------------------
sub displayPDFOptions(iOrgID)

  dim sSQL, lcl_orgid

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

 'Retrieve all of the PDFs
  sSQL = "SELECT distinct documenttitle, documentid, UPPER(documenttitle) "
  sSQL = sSQL & " FROM documents "
  sSQL = sSQL & " WHERE orgid = " & lcl_orgid
  sSQL = sSQL & " AND UPPER(documentURL) like ('%UNPUBLISHED_DOCUMENTS%/PDFS/%')"
  sSQL = sSQL & " AND RIGHT(UPPER(documentURL),4) = '.PDF' "
  sSQL = sSQL & " ORDER BY UPPER(documenttitle) "

  set oPDFs = Server.CreateObject("ADODB.Recordset")
  oPDFs.Open sSQL, Application("DSN") , 3, 1

  if not oPDFs.eof then
     while not oPDFs.eof
        response.write "<option value=""" & oPDFs("documentid") & """>" & oPDFs("documenttitle") & "</option>" & vbcrlf
        oPDFs.movenext
     wend
  end if

  oPDFs.close
  set oPDFs = nothing

end sub

'------------------------------------------------------------------------------
function send_notification(iID)

 lcl_orghasfeature_issue_location = orghasfeature("issue location")

'Determine who is to be notified
 lcl_notifyuserid = request("notifyuserid")
 lcl_notifydeptid = request("notifydeptid")
 lcl_admin_userid = session("userid")

'Atleast one of the "notify" fields must be populated.
 if lcl_notifyuserid <> "" OR lcl_notifydeptid <> "" then

   'Get the admin name and email
    'getUserInfo lcl_admin_userid, lcl_admin_username, lcl_admin_useremail
    lcl_admin_username  = getAdminName(lcl_admin_userid)
    lcl_admin_useremail = getUserEmail(lcl_admin_userid)

   'Build the body of the email
    lcl_request_url = getEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iID & "&e=Y"

    sEmailBody = "<p>" & lcl_admin_username & " has notified you about the following Action Line Request: "
    sEmailBody = sEmailBody & "<a href=""" & lcl_request_url & """>" & lcl_request_url & "</a></p>"

   'BEGIN: Issue Location -----------------------------------------------------
    lcl_address = ""
    sIssueName  = ""

'    if lcl_orghasfeature_issue_location then
       sSQLf = "SELECT f.issuelocationname, "
       sSQLf = sSQLf & " f.action_form_display_issue, "
       sSQLf = sSQLf & " f.hideIssueLocAddInfo "
       sSQLf = sSQLf & " FROM egov_action_request_forms f "
       sSQLf = sSQLf &      " INNER JOIN egov_actionline_requests r "
       sSQLf = sSQLf &              " ON f.action_form_id = r.category_id "
       sSQLf = sSQLf & " WHERE r.action_autoid = " & iID
       'sSQLf = sSQLf & " WHERE f.action_form_id = " & iID

       set oForm = Server.CreateObject("ADODB.Recordset")
       oForm.Open sSQLf, Application("DSN") , 3, 1

       if not oForm.eof then
          sIssueName           = UCASE(oForm("issuelocationname"))
          blnIssueDisplay      = oForm("action_form_display_issue")
          sHideIssueLocAddInfo = oForm("hideIssueLocAddInfo")

          If Trim(sIssueName) = "" OR IsNull(sIssueName) Then
             sIssueName = "ISSUE/PROBLEM LOCATION:"
          End If

         'Check to see if the "issue location" feature has been "turned on" for this form.
          'if blnIssueDisplay = True then

             sSQLi = "SELECT streetnumber, "
             sSQLi = sSQLi & " streetprefix, "
             sSQLi = sSQLi & " streetaddress, "
             sSQLi = sSQLi & " streetsuffix, "
             sSQLi = sSQLi & " streetdirection "
             sSQLi = sSQLi & " FROM egov_action_response_issue_location "
             sSQLi = sSQLi & " WHERE actionrequestresponseid = " & iID

             set oIssueLocation = Server.CreateObject("ADODB.Recordset")
             oIssueLocation.Open sSQLi, Application("DSN") , 3, 1

             if not oIssueLocation.eof then
                lcl_address = buildStreetAddress(oIssueLocation("streetnumber"), _
                                                 oIssueLocation("streetprefix"), _
                                                 oIssueLocation("streetaddress"), _
                                                 oIssueLocation("streetsuffix"), _
                                                 oIssueLocation("streetdirection"))

                'if not sHideIssueLocAddInfo then
                   sEmailBody = sEmailBody & "<p><strong>" & sIssueName & "</strong><br />"
                   sEmailBody = sEmailBody & lcl_address & "</p>" & vbcrlf
                'end if

             end if

             oIssueLocation.close
             set oIssueLocation = nothing

          'end if  'END blnIssueDisplay
       'end if  'END eof

       oForm.close
       set oForm = nothing

	   end if
   'END: Issue Location -------------------------------------------------------

   'BEGIN: Additional Comments ------------------------------------------------
    lcl_comments = ""

    if request("notify_additional_comments") <> "" then
       'lcl_comments = formatToFitEmailLineLength(request("notify_additional_comments"))
       lcl_comments = request("notify_additional_comments")

       sEmailBody = sEmailBody & "<p><strong>Additional Comments:</strong><br />"
       sEmailBody = sEmailBody & lcl_comments & "</p>" & vbcrlf
       'sEmailBody = sEmailBody & request("notify_additional_comments") & "</p>" & vbcrlf
    end if
  
   'First determine if the "Notify User" field has a value in it.
    if lcl_notifyuserid <> "" then
      'Get the user name and email
       lcl_notifyusername  = getAdminName(lcl_notifyuserid)
       lcl_notifyuseremail = getUserEmail(lcl_notifyuserid)

      'Check for a delegate
       getDelegateInfo lcl_notifyuserid, lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

      'Send the email
       setupSendEmail "notify", iID, sEmailBody, lcl_notifyuseremail, "", lcl_delegate_username, lcl_delegate_useremail
    end if
   'END: Additional Comments --------------------------------------------------

   'Second, check to see if the "Notify Department" field has a value in it.
    if lcl_notifydeptid <> "" then

       sSQLd = " SELECT distinct u.email, u.userid "
       sSQLd = sSQLd & " FROM usersgroups ug, users u "
       sSQLd = sSQLd & " WHERE u.userid = ug.userid "
       sSQLd = sSQLd & " AND (isrootadmin IS NULL OR isrootadmin = 0) "
       sSQLd = sSQLd & " AND u.email IS NOT NULL "
       sSQLd = sSQLd & " AND u.email <> '' "
       sSQLd = sSQLd & " AND ug.groupid = " & lcl_notifydeptid
       sSQLd = sSQLd & " AND u.orgid = "    & session("orgid")
       sSQLd = sSQLd & " ORDER BY u.email "

      	set oEmailDepts = Server.CreateObject("ADODB.Recordset")
      	oEmailDepts.Open sSQLd, Application("DSN"), 1, 3

       if NOT oEmailDepts.eof then
          while NOT oEmailDepts.eof

            'Check for a delegate
             getDelegateInfo oEmailDepts("userid"), lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

            'Send the email
             setupSendEmail "notify", iID, sEmailBody, oEmailDepts("email"), "", lcl_delegate_username, lcl_delegate_useremail

             oEmailDepts.movenext
          wend
       end if

      	oEmailDepts.close
      	set oEmailDepts = nothing 

    end if

   'Determine who the notification is being sent to.
   'Check for the user.
    if lcl_notifyusername <> "" then
       lcl_sentto = lcl_notifyusername
    end if

   'Check for the dept.
    if lcl_sentto <> "" then
       if lcl_notifydeptid <> "" then
          lcl_sentto = lcl_sentto & " (User) and " & getDeptName(lcl_notifydeptid) & " (Department)"
       end if
    else
       if lcl_notifydeptid <> "" then
          lcl_sentto = getDeptName(lcl_notifydeptid)
       end if
    end if       

   'Get the submit date
    lcl_submit_date = ConvertDateTimetoTimeZone()

   'Set up the Send Notification Activity Log comment
    sCommentLine = lcl_admin_username & " sent a notification of this request to "
    sCommentLine = sCommentLine & lcl_sentto
    sCommentLine = sCommentLine & " on " & lcl_submit_date

   'Determine if Additional Comments have been entered.
    if trim(request("notify_additional_comments")) <> "" then
       intComment = trim(request("notify_additional_comments"))
    end if

   'Build the Internal Comment
    if intComment <> "" then
       intComment = intComment & "<br />" & sCommentLine
    else
       intComment = sCommentLine
    end if

   'Create a record in the Activity Log
    AddCommentTaskComment trim(intComment), "", "", iID, session("userid"), session("orgid"), "", "", ""

 else
    response.redirect "action_respond.asp?control=" & iID & "&success=NO_NOTIFY_SENDTO"
 end if

end function

'------------------------------------------------------------------------------
'sub getUserInfo(ByVal p_userid, ByRef lcl_username, ByRef lcl_useremail)
' lcl_username  = ""
' lcl_useremail = ""

' if p_userid <> "" then
'   	sSQLu = "SELECT FirstName, LastName, email "
'    sSQLu = sSQLu & " FROM Users "
'    sSQLu = sSQLu & " WHERE userid = " & p_userid

'   	set oName = Server.CreateObject("ADODB.Recordset")
'   	oName.Open sSQLu, Application("DSN"), 1, 3

'    if NOT oName.eof then
'       lcl_username  = oName("FirstName") & " " & oName("LastName")
'       lcl_useremail = oName("email")
'    end if

'    oName.close
'    set oName = nothing

' end if

'end sub

'------------------------------------------------------------------------------
function getDeptName(p_value)
 lcl_return = ""

	sSQLd = "SELECT groupname "
 sSQLd = sSQLd & " FROM groups g "
 sSQLd = sSQLd & " WHERE g.orgid = " & session("orgid")
 sSQLd = sSQLd & " AND g.groupid = " & p_value

 set oDeptName = Server.CreateObject("ADODB.Recordset")
 oDeptName.Open sSQLd, Application("DSN"), 1, 3

 if NOT oDeptName.eof then
    lcl_return = oDeptName("groupname")
 end if

 oDeptName.close
 set oDeptName = nothing

 getDeptName = lcl_return

end function

'------------------------------------------------------------------------------
function getCategoryOptions(p_category_id)
	sSQL = "SELECT distinct form_category_id, form_category_name "
 sSQL = sSQL & " FROM dbo.egov_formlist "
 sSQL = sSQL & " WHERE form_category_id in (select distinct f2.form_category_id "
 sSQL = sSQL &                            " from dbo.egov_formlist f2 "
 sSQL = sSQL &                            " where f2.orgid = " & session("orgid") & ") "
 sSQL = sSQL & " ORDER BY form_category_name "

	set oCategoryOptions = Server.CreateObject("ADODB.Recordset")
	oCategoryOptions.Open sSQL, Application("DSN") , 3, 1

	if NOT oCategoryOptions.EOF then

  		do while NOT oCategoryOptions.eof

   				sCurrentCategory = oCategoryOptions("form_category_name")

   				if CStr(p_category_id) = CStr(oCategoryOptions("form_category_id")) then 
			     		selectA = "selected"
   				else
			     		selectA = ""
   				end if

   				response.write "  <option value=""" & oCategoryOptions("form_category_id") & """ " & selectA & ">" & oCategoryOptions("form_category_name") & "</option>" & vbcrlf
			
    			oCategoryOptions.movenext

    loop
	end if

 oCategoryOptions.close
	set oCategoryOptions = nothing

end function

'------------------------------------------------------------------------------
function getCategoryTitle(p_category_id)
  lcl_return = ""

  if p_category_id <> "" then
     sSQL = "SELECT action_form_name "
     sSQL = sSQL & " FROM egov_action_request_forms "
     sSQL = sSQL & " WHERE action_form_id = " & p_category_id

    	set oCatTitle = Server.CreateObject("ADODB.Recordset")
    	oCatTitle.Open sSQL, Application("DSN") , 3, 1

     if not oCatTitle.eof then
        lcl_return = trim(oCatTitle("action_form_name"))
     end if

     oCatTitle.close
     set oCatTitle = nothing

  end if

  getCategoryTitle = lcl_return

end function

'-- Copied from DrawAdminUsersNews (in includes/common.asp) -------------------
function DrawAdminUsersNew_javascript(suserid,isEmailRequired)
	dim sSql, oUsers, selected

	sSQL = "SELECT userid, FirstName, LastName "
 sSQL = sSQL & " FROM Users "
 sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " AND (IsRootAdmin IS NULL OR IsRootAdmin = 0) "

 if isEmailRequired = "Y" then
    sSQL = sSQL & " AND email IS NOT NULL "
    sSQL = sSQL & " AND email <> '' "
 end if

 sSQL = sSQL & " ORDER BY LastName, firstname "

	Set oAdminUsers = Server.CreateObject("ADODB.Recordset")
	oAdminUsers.Open sSQL, Application("DSN"), 1, 3

 i = 0
	while not oAdminUsers.eof
    i = i + 1
   	if suserid = oAdminUsers("userid") then
       selected = " selected=\""selected\"""
    else
       selected = ""
    end if

    if i > 1 then
       response.write "+"
    end if

  		response.write "'<option value=\""" & oAdminUsers("userid") & "\""" & selected & ">" & replace(oAdminUsers("FirstName"),"'","\'") & " " & replace(oAdminUsers("LastName"),"'","\'") & "</option>'" & vbcrlf
  		oAdminUsers.movenext
	wend

	oAdminUsers.close
	set oAdminUsers = nothing 

end function

'------------------------------------------------------------------------------
sub displayButtons(iRequestID,iButtons)

 'Build the parameter string for the return url.
  lcl_querystring_len  = ""
  lcl_first_amp_pos    = ""
  lcl_return_str_right = ""
  lcl_return_str       = ""
  lcl_control_pos      = ""

  if request.querystring <> "" then
     lcl_querystring_len = len(request.querystring)

    'Remove the [&control=999999] from the querystring
    'First check to see if it is the first parameter
     if ucase(left(request.querystring,8)) = "CONTROL=" then
        lcl_first_amp_pos = instr(request.querystring,"&")

       'Get the string to the RIGHT of the [control] parameter
        if lcl_first_amp_pos > 0 then
           lcl_return_str_right = mid(request.querystring,lcl_first_amp_pos+1,lcl_querystring_len)
        end if

       'Build the return string
        lcl_return_str = lcl_return_str_right
     else
        lcl_control_pos = instr(request.querystring,"&control=")
        lcl_return_str  = mid(request.querystring,1,lcl_control_pos-1)
     end if

     session("return_str") = lcl_return_str
  else
    'Check for the session variable.
     if session("return_str") <> "" then
        lcl_return_str = session("return_str")
     'else
     '   lcl_return_str = "useSessions=1"
     end if
  end if

  if lcl_return_str <> "" then
     lcl_return_str = "?" & lcl_return_str

     if instr(lcl_return_str,"useSessions") < 1 then
        lcl_return_str = lcl_return_str & "&useSessions=1"
     end if
  else
     lcl_return_str = "?useSessions=1"
  end if

 'Override the return string if the user has accessed this request from an email
  if session("isFromEmail") <> "" then
     lcl_isFromEmail = session("isFromEmail")
  else
     if request("e") <> "" then
        lcl_isFromEmail = request("e")
     else
        lcl_isFromEmail = "N"
     end if
  end if

  if lcl_isFromEmail = "Y" then
     session("isFromEmail") = "Y"
     lcl_return_str = "?init=Y"
  else
     session("isFromEmail") = ""
  end if

  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><input type=""button"" name=""sBack"" id=""sBack"" value=""Back"" class=""button"" onclick=""location.href='action_line_list.asp" & lcl_return_str & "'"" /></td>" & vbcrlf
  response.write "      <td align=""right"">" & vbcrlf

  if ucase(iButtons) = "WORKORDER" then
     response.write "          <input type=""button"" class=""button"" onclick=""return openPDFtoView('" & iTrackID & "','','WORKORDER');"" value=""Print Work Order"" />" & vbcrlf
     response.write "          <input type=""button"" class=""button"" onclick=""return openPDFtoView('" & iTrackID & "','','WORKORDER_CONDENSED');"" value=""Print Condensed Work Order"" />" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub updateActivityLog_PDFs(iActLog,iPDFID,iPDFName,iSelStatus,iID,iSelSubStatus)
 lcl_actlog            = UCASE(iActLog)
 lcl_pdfid             = iPDFID
 lcl_pdfname           = ""
 lcl_intComment        = ""
 lcl_updateActivityLog = "N"

'Get the name of the pdf
 if lcl_pdfid <> "" then
    if isnumeric(lcl_pdfid) then
      'Get the name of the pdf
       sSQL = "SELECT documenttitle "
       sSQL = sSQL & " FROM documents "
       sSQL = sSQL & " WHERE orgid = " & session("orgid")
       sSQL = sSQL & " AND documentid = " & lcl_pdfid

       set oPDFName = Server.CreateObject("ADODB.Recordset")
       oPDFName.Open sSQL, Application("DSN") , 3, 1

       if not oPDFName.eof then
          lcl_pdfname = oPDFName("documenttitle")
       end if

       oPDFName.close
       set oPDFName = nothing
    end if
 end if

'Build the Internal Message based on which button was pressed.
 if lcl_actlog = "PRINT_PDF" then

    lcl_intComment        = lcl_pdfname & " has been printed."
    lcl_updateActivityLog = "Y"

 elseif lcl_actlog = "WORKORDER" then

   'If the org has the "Activity Log Tracking - Action Line Work Orders"
   'then check to see if a record is to be inserted.
    if lcl_orghasfeature_activitylog_workorder then
       lcl_intComment        = "The Work Order has been printed."
       lcl_updateActivityLog = "Y"
    end if

 elseif lcl_actlog = "WORKORDER_CONDENSED" then

   'If the org has the "Activity Log Tracking - Action Line Work Orders"
   'then check to see if a record is to be inserted.
    if lcl_orghasfeature_activitylog_workorder then
       lcl_intComment = "The Work Order (Condensed) has been printed."
       lcl_updateActivityLog = "Y"
    end if

 elseif lcl_actlog = "VIEW_PUBLIC_PDF" then

    lcl_intComment = iPDFName & " (Public-side PDF) has been printed."
    lcl_updateActivityLog = "Y"

 end if

'Update the Activity Log
 if lcl_updateActivityLog = "Y" AND lcl_intComment <> "" then
    AddCommentTaskComment lcl_intComment,"",iSelStatus,iID,session("userid"),session("orgid"),iSelSubStatus, "", ""
 end if

end sub

'------------------------------------------------------------------------------
sub displayLinkedRequests(iOrgID, iRequestID)

  dim sSQL, lcl_orgid, lcl_requestid, lcl_bgcolor, iRowCount

  lcl_orgid     = 0
  lcl_requestid = 0
  iRowCount     = 0

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

  if iRequestID <> "" then
     lcl_requestid = clng(iRequestID)
  end if

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Linked Requests</legend>" & vbcrlf
  response.write "  <div id=""linkedRequestAddButton"" align=""left"">" & vbcrlf
  response.write "    <input type=""button"" id=""button_linkrequest"" value=""Link a Request"" class=""button"" onclick=""addLinkedRequestRow();"" />" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  <table id=""linkedRequests"" border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"" class=""tableadmin"">" & vbcrlf
  response.write "    <tr id=""linkedRequestsRow_0"">" & vbcrlf
  response.write "        <th>Tracking Number</th>" & vbcrlf
  response.write "        <th>Description</th>" & vbcrlf
  response.write "        <th>&nbsp;</th>" & vbcrlf
  response.write "    </tr>" & vbcrlf

  sSQL = " SELECT linkedrequestid, "
  sSQL = sSQL & " linked_trackingnumber as trackingnumber, "
  sSQL = sSQL & " linked_requestid as action_autoid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_actionline_linkedrequests "
  sSQL = sSQL & " WHERE orgid = " & lcl_orgid
  sSQL = sSQL & " AND parent_requestid = "  & lcl_requestid
  sSQL = sSQL & " AND linked_requestid <> " & lcl_requestid
  sSQL = sSQL & " UNION ALL "
  sSQL = sSQL & " SELECT linkedrequestid, "
  sSQL = sSQL & " parent_trackingnumber as trackingnumber, "
  sSQL = sSQL & " parent_requestid as action_autoid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_actionline_linkedrequests "
  sSQL = sSQL & " WHERE orgid = " & lcl_orgid
  sSQL = sSQL & " AND linked_requestid = "  & lcl_requestid
  sSQL = sSQL & " AND parent_requestid <> " & lcl_requestid

 	set oDisplayLinkedRequests = Server.CreateObject("ADODB.Recordset")
	 oDisplayLinkedRequests.Open sSQL, Application("DSN") , 3, 1

  if not oDisplayLinkedRequests.eof then
     lcl_bgcolor = "#eeeeee"

     do while not oDisplayLinkedRequests.eof
        iRowCount = iRowCount + 1

        response.write "    <tr id=""linkedRequestsRow_" & iRowCount & """ valign=""top"">" & vbcrlf
        response.write "        <td bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "            <input type=""hidden"" name=""linkedRequestID_" & iRowCount & """ id=""linkedRequestID_" & iRowCount & """ size=""5"" maxlength=""50"" value=""" & oDisplayLinkedRequests("linkedrequestid") & """ />" & vbcrlf
        response.write "            <input type=""hidden"" name=""trackingnumber_edit_" & iRowCount & """ id=""trackingnumber_edit_" & iRowCount & """ value=""" & oDisplayLinkedRequests("trackingnumber") & """ size=""10"" maxlength=""50"" />" & vbcrlf
        response.write "            <span id=""trackingnumber_text_" & iRowCount & """><a href=""action_respond.asp?init=Y&useSessions=1&control=" & oDisplayLinkedRequests("action_autoid") & """>"& oDisplayLinkedRequests("trackingnumber") & "</a></span>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "        <td width=""200"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "            <span id=""description_text_" & iRowCount & """>" & oDisplayLinkedRequests("description") & "</span>" & vbcrlf
        response.write "            <textarea name=""description_edit_" & iRowCount & """ id=""description_edit_" & iRowCount & """ rows=""3"" cols=""33"">" & oDisplayLinkedRequests("description") & "</textarea>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "        <td bgcolor=""" & lcl_bgcolor & """ id=""row_linkedRequestsButtons"" nowrap=""nowrap"">" & vbcrlf
        response.write "            <input type=""button"" name=""button_edit_"   & iRowCount & """ id=""button_edit_"   & iRowCount & """ value=""Edit"" class=""button"" onclick=""editLink('" & iRowCount & "');"" />" & vbcrlf
        response.write "            <input type=""button"" name=""button_save_"   & iRowCount & """ id=""button_save_"   & iRowCount & """ value=""Save Changes"" class=""button"" onclick=""saveLinkChanges('" & iRowCount & "');"" />" & vbcrlf
        response.write "            <input type=""button"" name=""button_remove_" & iRowCount & """ id=""button_remove_" & iRowCount & """ value=""Remove Link"" class=""button"" onclick=""removeLink(" & iRowCount & ", '" & lcl_requestid & "','" & oDisplayLinkedRequests("action_autoid") & "');"" />" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf

        lcl_bgcolor = changeBGColor(lcl_bgcolor, "#eeeeee", "#ffffff")

        oDisplayLinkedRequests.movenext
     loop
  end if

  oDisplayLinkedRequests.close
  set oDisplayLinkedRequests = nothing

  response.write "  </table>" & vbcrlf
  response.write "  <input type=""hidden"" name=""sTotalLinkedRequestRows"" id=""sTotalLinkedRequestRows"" value=""" & iRowCount & """ />" & vbcrlf
  response.write "</fieldset>" & vbcrlf

end sub
%>

