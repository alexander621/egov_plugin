<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="subscriptionslog_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: subscriptionslog_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a RSS Feed
'
' MODIFICATION HISTORY
' 1.0 06/29/09 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptionslog_maint") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if request("listtype") <> "" then
    lcl_list_type = ucase(request("listtype"))
 else
    lcl_list_type = ""
 end if

 if lcl_list_type = "BID" then
    lcl_pagetitle      = "Bid Postings"
    lcl_userpermission = "subscriptionslog_maint_bids"
 elseif lcl_list_type = "JOB" then
    lcl_pagetitle      = "Job Postings"
    lcl_userpermission = "subscriptionslog_maint_jobs"
 else
    lcl_pagetitle      = "Distribution Lists"
    lcl_userpermission = "subscriptionslog_maint"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 lcl_isRootAdmin = false

 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = true
 end if

'Retrieve the dl_logid of the subscription log record
'If no value exists then redirect them back to the main results screen
 if request("dl_logid") <> "" then
    lcl_dl_logid = request("dl_logid")
 else
    response.redirect "subscriptionslog_list.asp?listtype=" & lcl_list_type
 end if

'Set up local variables
 lcl_sentbyuserid    = 0
 lcl_sentbyusername  = ""
 lcl_sentdate        = ""
 lcl_completedate    = ""
 lcl_sendstatus      = ""
 lcl_email_fromname  = ""
 lcl_email_fromemail = ""
 lcl_email_subject   = ""
 lcl_email_body      = ""
 lcl_email_format    = ""
 lcl_containsHTML    = "No"

'Retrieve all of the data for the subscription log
 sSQL = "SELECT dl_logid, "
 sSQL = sSQL & " sentbyuserid, "
 sSQL = sSQL & " sentdate, "
 sSQL = sSQL & " completedate, "
 sSQL = sSQL & " sendstatus, "
 sSQL = sSQL & " email_fromname, "
 sSQL = sSQL & " email_fromemail, "
 sSQL = sSQL & " email_subject, "
 sSQL = sSQL & " email_body, "
 sSQL = sSQL & " email_format, "
 sSQL = sSQL & " dl_listids, "
 sSQL = sSQL & " containsHTML,scheduledDateTime, "
 sSQL = sSQL & " (select firstname + ' ' + lastname "
 sSQL = sSQL &  " from users "
 sSQL = sSQL &  " where userid = sentbyuserid) AS sentbyusername "
 sSQL = sSQL & " FROM egov_class_distributionlist_log "
 sSQL = sSQL & " WHERE orgid = "  & session("orgid")
 sSQL = sSQL & " AND dl_logid = " & lcl_dl_logid

	set oListLog = Server.CreateObject("ADODB.Recordset")
 oListLog.Open sSQL, Application("DSN"), 3, 1
	
	if not oListLog.eof then
    lcl_sentbyuserid      = oListLog("sentbyuserid")
    lcl_sentbyusername    = trim(oListLog("sentbyusername"))
    lcl_sentdate          = oListLog("sentdate")
    lcl_completedate      = oListLog("completedate")
    lcl_sendstatus        = oListLog("sendstatus")
    lcl_email_fromname    = oListLog("email_fromname")
    lcl_email_fromemail   = oListLog("email_fromemail")
    lcl_email_subject     = oListLog("email_subject")
    lcl_email_body        = oListLog("email_body")
    lcl_email_format      = getEmailFormatDesc(oListLog("email_format"))
    lcl_dl_listids        = oListLog("dl_listids")
    lcl_distributionlists = getDistributionListNames(session("orgid"), lcl_dl_listids)
    lcl_scheduledDateTime = oListLog("scheduledDateTime")

    if oListLog("containsHTML") then
       lcl_containsHTML = "Yes"
    end if

   'Set up the "sent by username"
    if lcl_sentbyusername = "" then
       lcl_sentbyusername = "<span style=""color:#800000;"">N/A</span>"
    end if

    if lcl_isRootAdmin then
       lcl_sentbyusername = "[<a href=""../dirs/update_user.asp?userid=" & lcl_sentbyuserid & "&currentpage=1"" target=""_blank"">" & lcl_sentbyuserid & "</a>] " & lcl_sentbyusername
    end if
 else
    response.redirect("subscrptionslog_list.asp?success=NE")
 end if

 oListLog.close
 set oListLog = nothing

'Check for a screen message
 lcl_onload  = "setMaxLength();"
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
 lcl_hidden = "HIDDEN"

'Set up required field icon
' lcl_required_field = "<span style=""color:#ff0000"">*</span>"
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_pagetitle%> - Send Log Details}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>

<script language="javascript">
var control_field = "";

function viewEmail() {
  window.open("subscriptionslog_viewemail.asp?dl_logid=<%=lcl_dl_logid%>&listtype=<%=lcl_list_type%>","_blank");
}

function confirmDelete() {
  var r = confirm('Are you sure you want to delete this log file?');
  if (r==true) {
      location.href="subscriptionslog_action.asp?user_action=DELETE&listtype=<%=lcl_list_type%>&dl_logid=<%=lcl_dl_logid%>";
  }
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>

<!-- #include file="../menu/menu.asp" -->

<div id="centercontent">
<table border="0" cellspacing="0" cellpadding="10" width="800" class="start">
  <form name="subscriptionslog_maint" id="subscriptionslog_maint" method="post" action="rssfeeds_action.asp">
    <input type="<%=lcl_hidden%>" name="dl_logid" value="<%=lcl_dl_logid%>" size="5" maxlength="5" />
    <input type="<%=lcl_hidden%>" name="listtype" id="listtype" value="<%=lcl_list_type%>" size="5" maxlength="100" />
    <input type="<%=lcl_hidden%>" name="user_action" value="" size="4" maxlength="4" />
    <input type="<%=lcl_hidden%>" name="orgid" value="<%=lcl_orgid%>" size="4" maxlength="10" />
  <tr>
      <td>
          <font size="+1"><strong><%=lcl_pagetitle%>: Send Log Details</strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Return to <%=lcl_pagetitle%> Send Log List" class="button" onclick="location.href='subscriptionslog_list.asp?listtype=<%=lcl_list_type%>'" />
      </td>
  </tr>
  <tr valign="top">
      <td>
          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <tr>
                <td align="left" style="font-size:10px;">
                    <% displayButtons "TOP", lcl_dl_logid, lcl_list_Type %>
                </td>
                <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
            </tr>
          </table>
          <table border="0" cellspacing="0" cellpadding="2" class="tableadmin">
            <tr>
                <th colspan="2" align="left"><%=lcl_pagetitle%></th>
            </tr>
            <tr>
                <td>Date Sent:</td>
                <td><%=lcl_sentdate%></td>
            </tr>
            <tr>
                <td>Sent By:</td>
                <td><%=lcl_sentbyusername%></td>
            </tr>
            <tr>
                <td>Send Status:</td>
                <td><%=lcl_sendstatus%></td>
            </tr>
	    <% if lcl_sendstatus = "SCHEDULED" then %>
            <tr>
                <td>Scheduled Date:</td>
                <td><%=lcl_scheduledDateTime%></td>
            </tr>
	    <% end if %>
            <tr>
                <td>Send Completed on:</td>
                <td><%=lcl_completedate%></td>
            </tr>
            <tr>
                <td>Email Format:</td>
                <td><%=lcl_email_format%></td>
            </tr>
            <tr>
                <td>Contains HTML:</td>
                <td><%=lcl_containsHTML%></td>
            </tr>
            <tr><td colspan="2">&nbsp;</td></tr>
            <tr>
                <td>From:</td>
                <td><%=lcl_email_fromname%> [<%=lcl_email_fromemail%>]</td>
            </tr>
            <tr>
                <td>Subject:</td>
                <td><%=lcl_email_subject%></td>
            </tr>
            <tr valign="top">
                <td>&nbsp;<br />Email Body:</td>
                <td>
                    <div align="right"><input type="button" name="viewEmailButton" id="viewEmailButton" value="View Email Sent" class="button" onclick="viewEmail()" /></div>
                    <textarea rows="10" cols="100"><%=lcl_email_body%></textarea>
                </td>
            </tr>
            <tr>
                <td nowrap="nowrap">Sent to Distribution List(s):</td>
                <td><%=lcl_distributionlists%></td>
            </tr>
            <tr><td colspan="2">&nbsp;</td></tr>
          </table>
          <% displayButtons "BOTTOM", lcl_dl_logid, lcl_list_type %>
      </td>
  </tr>
</table>
</div>

<!--#include file="../admin_footer.asp"-->

</body>
</html>
<%
'-----------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = p_value
     lcl_value = replace(lcl_value,"'","''")
     lcl_value = replace(lcl_value,"<","&lt;")
     lcl_value = replace(lcl_value,">","&gt;")
  else
     lcl_value = p_value
  end if

  dbsafe = lcl_value

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "RSS Feed does not exist..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iDL_LogID, iListType)

  if iTopBottom <> "" then
     iTopBottom = UCASE(iTopBottom)
  else
     iTopBottom = "TOP"
  end if

  if iTopBottom = "BOTTOM" then
     lcl_style_div = "padding-top: 5px;"
  else
     lcl_style_div = "padding-bottom: 5px;"
  end if

  'lcl_return_parameters = "?sc_org_name=" & session("sc_org_name") & "&sc_show_members=" & session("sc_show_members")
  lcl_return_parameters = ""

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='rssfeeds_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf
  response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
  response.write "<input type=""button"" name=""resendButton"" id=""resendButton"" value=""Copy to new email"" class=""button"" onclick=""location.href='../classes/dl_sendmail.asp?dl_logid=" & iDL_LogID & "&listtype=" & iListType & "';"" />" & vbcrlf
  response.write "<div>" & vbcrlf

end sub
%>
