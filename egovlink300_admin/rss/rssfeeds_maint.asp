<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rssfeeds_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a RSS Feed
'
' MODIFICATION HISTORY
' 1.0 04/08/09 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("rssfeeds") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../"  'Override of value from common.asp

'if request("screen_mode") = "ADD" then
'   if NOT UserHasPermission( session("userid"), "create_organizational_groups" ) then
'      if userhaspermission(session("userid"),"staff_directory") then
'         response.redirect "organizational_groups_list.asp?success=NO_ADD_ROLE"
'      else
'     	   response.redirect sLevel & "permissiondenied.asp"
'      end if
'   end if

'   lcl_page_title = "Add"

'else
'   if NOT UserHasPermission( session("userid"), "edit_organizational_groups" ) then
'      if userhaspermission(session("userid"),"staff_directory") then
'         response.redirect "organizational_groups_list.asp?success=NO_EDIT_ROLE"
'      else
'     	   response.redirect sLevel & "permissiondenied.asp"
'      end if
'   end if

'   lcl_page_title = "Edit"
'end if

'Retrieve the feedid of the organization group that is to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("feedid") <> "" then
    lcl_feedid = request("feedid")

    if isnumeric(lcl_feedid) then
       lcl_screen_mode = "EDIT"
    else
       response.redirect "rssfeeds_list.asp"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_feedid      = 0
 end if

'Set up local variables
 lcl_feedname      = ""
 lcl_isActive      = True
 lcl_rsstitle      = ""
 lcl_orgtitle      = ""
 lcl_description   = ""
 lcl_feedurl       = ""
 lcl_lastbuilddate = ""
 lcl_feature       = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the rss feed
    sSQL = "SELECT feedid, feedname, isActive, title, description, feedurl, lastbuilddate, feature "
    sSQL = sSQL & " FROM egov_rssfeeds "
    sSQL = sSQL & " WHERE feedid = "  & lcl_feedid

    set oRSSFeed = Server.CreateObject("ADODB.Recordset")
    oRSSFeed.Open sSQL, Application("DSN"), 3, 1

    if not oRSSFeed.eof then
       lcl_feedname      = oRSSFeed("feedname")
       lcl_isActive      = oRSSFeed("isActive")
       lcl_rsstitle      = oRSSFeed("title")
       lcl_orgtitle      = getOrgRSSTitle(lcl_feedid,session("orgid"))
       lcl_description   = oRSSFeed("description")
       lcl_feedurl       = oRSSFeed("feedurl")
       lcl_lastbuilddate = oRSSFeed("lastbuilddate")
       lcl_feature       = oRSSFeed("feature")
    else
       response.redirect("rssfeeds_list.asp?success=NE")
    end if

    oRSSFeed.close
    set oRSSFeed = nothing

 end if

'Determine if the the active checkbox is "checked".
 if lcl_isActive then
    lcl_checked_isActive = " checked=""checked"""
 else
    lcl_checked_isActive = ""
 end if

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
 lcl_hidden = "HIDDEN"

'Set the page title
 lcl_pagetitle = "RSS Feeds Maintenance"

'Set up required field icon
 lcl_required_field = "<span style=""color:#ff0000"">*</span>"

'Check to see if any RSS Items are associated to this RSS Feed.
 lcl_rssItemsExist = checkForRSSItems(lcl_feedid)
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_pagetitle%> - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<style>
.rssAlert
{
   border: 1pt solid #c0c0c0;
   border-radius: 6px;
   background-color: #efefef;
   margin: 10px 5px;
   padding: 10px;
   font-size: 1.5em;
   color: #ff0000;
}

.rssAlertName
{
   color: #000000;
   text-align: center;
}
</style>

  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
var control_field = "";

function doPicker(sFormField) {
  //w = (screen.width - 350)/2;
  //h = (screen.height - 350)/2;
  w = 600;
  h = 400;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += "&folderStart=published_documents";
  pickerURL += "&displayDocuments=Y";
  pickerURL += "&displayActionLine=Y";
  pickerURL += "&displayPayments=Y";
  pickerURL += "&displayURL=Y";

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function storeCaret (textEl) {
  if (textEl.createTextRange)
      textEl.caretPos = document.selection.createRange().duplicate();
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
      var caretPos = textEl.caretPos;
      caretPos.text =
      caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
      text + ' ' : text;
  }
   else
      textEl.value = textEl.value + text;
}

function confirmDelete() {
  if("<%=lcl_rssItemsExist%>"=="Y") {
     lcl_msg  = '"' + document.getElementById("rsstitle") + '" cannot be deleted as there are RSS Items associated to it.\n';
     lcl_msg += 'Set the RSS Feed to "inactive".';

     alert(lcl_msg);
  }else{
     var r = confirm('Are you sure you want to delete the "' + document.getElementById("rsstitle").value + '" RSS Feed?');
     if (r==true) {
         location.href="rssfeeds_action.asp?user_action=DELETE&feedid=<%=lcl_feedid%>";
     }
  }
}

function validateFields(p_action) {
  var lcl_false_count = 0;

  if(document.getElementById("feature").value=="") {
     document.getElementById("feature").focus();
     inlineMsg(document.getElementById("feature").id,'<strong>Required Field Missing: </strong> Feature',10,'feature');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("feature");
  }

  if(document.getElementById("feedurl").value=="") {
     document.getElementById("feedurl").focus();
     inlineMsg(document.getElementById("feedurl").id,'<strong>Required Field Missing: </strong> Feed URL',10,'feedurl');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("feedurl");
  }

  if(document.getElementById("rsstitle").value=="") {
     document.getElementById("rsstitle").focus();
     inlineMsg(document.getElementById("rsstitle").id,'<strong>Required Field Missing: </strong> Title',10,'rsstitle');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("rsstitle");
  }

  if(document.getElementById("feedname").value=="") {
     document.getElementById("feedname").focus();
     inlineMsg(document.getElementById("feedname").id,'<strong>Required Field Missing: </strong> Feed Name',10,'feedname');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("feedname");
  }

  if(lcl_false_count > 0) {
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("rssfeeds_maint").submit();
     return true;
  }
}

function displayFeature() {
  document.getElementById("displayfeature").innerHTML='[' + document.getElementById("feature").value + ']'
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
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="setMaxLength();displayFeature();<%=lcl_onload%>">
<% ShowHeader sLevel %>

<!-- #include file="../menu/menu.asp" -->

<div id="centercontent">
<table border="0" cellspacing="0" cellpadding="10" width="800" class="start">
  <form name="rssfeeds_maint" id="rssfeeds_maint" method="post" action="rssfeeds_action.asp">
    <input type="<%=lcl_hidden%>" name="feedid" id="feedid" value="<%=lcl_feedid%>" size="5" maxlength="5" />
    <input type="<%=lcl_hidden%>" name="screen_mode" id="screen_mode" value="<%=lcl_screen_mode%>" size="4" maxlength="4" />
    <input type="<%=lcl_hidden%>" name="user_action" id="user_action" value="" size="4" maxlength="4" />
    <input type="<%=lcl_hidden%>" name="orgid" id="orgid" value="<%=lcl_orgid%>" size="4" maxlength="10" />
  <tr>
      <td>
          <font size="+1"><strong><%=lcl_pagetitle%>: <%=lcl_screen_mode%></strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Return to RSS Feeds List" class="button" onclick="location.href='rssfeeds_list.asp'" />
      </td>
  </tr>
  <tr valign="top">
      <td>

          <div class="rssAlert">
            DO NOT CHANGE ANY INFORMATION ON THIS SCREEN EXCEPT THE "Org Title".  This screen was created SPECIFICALLY
            for me and allows me to set up and maintain the RSS Feeds and NOT the org specific data that may have been
            sent to the feed per organization.  DELETE OR MODIFY ONE OF THESE FEEDS AND YOU AFFECT EVERYONE.
            <div class="rssAlertName">David Boyer</div>
          </div>

          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <tr>
                <td align="left" style="font-size:10px;">
                    <% displayButtons "TOP", lcl_screen_mode %>
                </td>
                <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
            </tr>
          </table>
          <table border="0" cellspacing="0" cellpadding="2" class="tableadmin">
            <tr>
                <th colspan="2" align="left"><%=lcl_pagetitle%></th>
            </tr>
            <tr>
                <td><%=lcl_required_field%>Feed Name:</td>
                <td>
                    <input type="text" name="feedname" id="feedname" value="<%=lcl_feedname%>" size="50" maxlength="50" onchange="clearMsg('feedname');" />
                    <img src="../images/help.jpg" name="helpFeedName" id="helpFeedName" alt="Used in RSS Feed page.  MUST be unique!" />
                </td>
            </tr>
            <tr valign="top">
                <td><%=lcl_required_field%>Feature:</td>
                <td>
                    <select name="feature" id="feature" onchange="displayFeature();">
                      <% showFeatureOptions lcl_feature %>
                    </select>
                    <img src="../images/help.jpg" name="helpFeature" id="helpFeature" alt="Used in RSS Feed page to determine if org has the feature turned on." /><br />
                    <span id="displayfeature" style="color:#800000;"></span>
                </td>
            </tr>
            <tr><td colspan="2">&nbsp;</td></tr>
            <tr>
                <td><%=lcl_required_field%>Title:</td>
                <td>
                    <input type="text" name="rsstitle" id="rsstitle" value="<%=lcl_rsstitle%>" size="50" maxlength="500" onchange="clearMsg('rsstitle');" />
                    <img src="../images/help.jpg" name="helpFeedTitle" id="helpFeedTitle" alt="Default value used for title for this RSS Feed." />
                </td>
            </tr>
            <tr>
                <td>Org Title:</td>
                <td>
                    <input type="text" name="orgtitle" id="orgtitle" value="<%=lcl_orgtitle%>" size="50" maxlength="500" />
                    <img src="../images/help.jpg" name="helpFeedOrgTitle" id="helpFeedOrgTitle" alt="Value used by this org to override default title.  If left NULL then default title value is displayed." />
                </td>
            </tr>
            <tr>
                <td><%=lcl_required_field%>Feed URL:</td>
                <td><input type="text" name="feedurl" id="feedurl" value="<%=lcl_feedurl%>" size="50" maxlength="1000" onchange="clearMsg('feedurl');" /></td>
            </tr>
            <tr>
                <td>Is Active?</td>
                <td>
                    <input type="checkbox" name="isActive" value="on" id="isActive"<%=lcl_checked_isActive%>" />
                    <span style="font-size:10px; color:#800000">(A page MUST be created in the RSS folder on the public-side.)</span>
                </td>
            </tr>
          <%
            if lcl_screen_mode = "EDIT" then
               response.write "  <tr>" & vbcrlf
               response.write "      <td>Last Build Date:</td>" & vbcrlf
               response.write "      <td style=""color:#800000;"">" & lcl_lastbuilddate & "</td>" & vbcrlf
               response.write "  </tr>" & vbcrlf
            end if
          %>
            <tr><td colspan="2">&nbsp;</td></tr>
            <tr>
                <td colspan="2">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%">
                      <tr>
                          <td>Description:</td>
                          <td align="right"><input type="button" name="addLinkButton" id="addLinkButton" value="Add a Link" class="button" onclick="doPicker('rssfeeds_maint.article');" /></td>
                      </tr>
                    </table>
                    <textarea name="article" id="article" rows="20" cols="100" maxlength="1000"><%=lcl_description%></textarea>
                </td>
            </tr>
            <tr><td colspan="2">&nbsp;</td></tr>
          </table>
          <% displayButtons "BOTTOM", lcl_screen_mode %>
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
     lcl_value = REPLACE(p_value,"'","''")
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
sub displayButtons(iTopBottom, iScreenMode)

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

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" onclick=""return validateFields('ADDANOTHER');"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     'response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function checkForRSSItems(iFeedID)

  lcl_return = "N"

  if iFeedID <> "" then
     sSQL = "SELECT DISTINCT 'Y' AS lcl_exists "
     sSQL = sSQL & " FROM egov_rss "
     sSQL = sSQL & " WHERE orgid = " & session("orgid")
     sSQL = sSQL & " AND feedid = "  & iFeedID

     set oRSSItemsExists = Server.CreateObject("ADODB.Recordset")
     oRSSItemsExists.Open sSQL, Application("DSN"), 3, 1

     if not oRSSItemsExists.eof then
        lcl_return = oRSSItemsExists("lcl_exists")
     end if

     oRSSItemsExists.close
     set oRSSItemsExists = nothing

  end if

  checkForRSSItems = lcl_return

end function

'------------------------------------------------------------------------------
sub showFeatureOptions(iFeature)

  sSQL = "SELECT feature, featurename "
  sSQL = sSQL & " FROM egov_organization_features "
  sSQL = sSQL & " ORDER BY featurename "

  set oFeatureOptions = Server.CreateObject("ADODB.Recordset")
  oFeatureOptions.Open sSQL, Application("DSN"), 3, 1

  if not oFeatureOptions.eof then
     do while not oFeatureOptions.eof

        if UCASE(iFeature) = UCASE(oFeatureOptions("feature")) then
           lcl_selected_feature = " selected=""selected"""
        else
           lcl_selected_feature = ""
        end if

        response.write "  <option value=""" & oFeatureOptions("feature") & """" & lcl_selected_feature & ">" & oFeatureOptions("featurename") & "</option>" & vbcrlf

        oFeatureOptions.movenext
     loop
  end if

  oFeatureOptions.close
  set oFeatureOptions = nothing

end sub

'------------------------------------------------------------------------------
function getOrgRSSTitle(p_feedid,p_orgid)

  lcl_return = ""

  if p_feedid <> "" AND p_orgid <> "" then
     sSQL = "SELECT orgtitle "
     sSQL = sSQL & " FROM egov_rssfeeds_orgtitles "
     sSQL = sSQL & " WHERE orgid = " & p_orgid
     sSQL = sSQL & " AND feedid = "  & p_feedid

     set oFeedTitle = Server.CreateObject("ADODB.Recordset")
     oFeedTitle.Open sSQL, Application("DSN"), 3, 1

     if not oFeedTitle.eof then
        lcl_return = oFeedTitle("orgtitle")
     end if

     oFeedTitle.close
     set oFeedTitle = nothing

  end if

  getOrgRSSTitle = lcl_return

end function
%>