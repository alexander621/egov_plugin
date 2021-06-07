<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mayorsblog_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a Blog entry
'
' MODIFICATION HISTORY
' 1.0 04/03/09 David Boyer - Initial Version
' 1.1 06/09/09	David Boyer - Added checkbox for "send to" function.  (Send to features like RSS and eventually Twitter, etc.)
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("mayorsblog") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"mayorsblog_maint") then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the blogid of the organization group that is to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("blogid") <> "" then
    lcl_blogid = request("blogid")

    if isnumeric(lcl_blogid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "mayorsblog_list.asp"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_sendToLabel = "Create"
    lcl_blogid      = 0
 end if

'Set up local variables
 lcl_userid             = 0
 lcl_imagefilename      = ""
 lcl_title              = ""
 lcl_article            = ""
 lcl_createdbyid        = 0
 lcl_createdbydate      = ""
 lcl_createdbyname      = ""
 lcl_isInactive         = False
 lcl_lastmodifiedbyid   = 0
 lcl_lastmodifiedbydate = ""
 lcl_lastmodifiedbyname = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the blog
    sSQL = "SELECT b.blogid, b.orgid, b.userid, b.title, b.article, b.createdbyid, b.createdbydate, b.isInactive, "
    sSQL = sSQL & " b.lastmodifiedbyid, b.lastmodifiedbydate, u.imagefilename, u2.firstname + ' ' + u2.lastname AS createdbyname, "
    sSQL = sSQL & " u3.firstname + ' ' + u3.lastname AS lastmodifiedbyname "
    sSQL = sSQL & " FROM egov_mayorsblog b "
    sSQL = sSQL &      " LEFT OUTER JOIN users u  ON b.userid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON b.createdbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u3 ON b.lastmodifiedbyid = u3.userid AND u3.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE b.blogid = " & lcl_blogid

    set oBlog = Server.CreateObject("ADODB.Recordset")
    oBlog.Open sSQL, Application("DSN"), 3, 1

    if not oBlog.eof then
       lcl_orgid              = oBlog("orgid")
       lcl_userid             = oBlog("userid")
       lcl_imagefilename      = oBlog("imagefilename")
       lcl_title              = oBlog("title")
       lcl_article            = oBlog("article")
       lcl_createdbyid        = oBlog("createdbyid")
       lcl_createdbydate      = oBlog("createdbydate")
       lcl_createdbyname      = oBlog("createdbyname")
       lcl_isInactive         = oBlog("isInactive")
       lcl_lastmodifiedbyid   = oBlog("lastmodifiedbyid")
       lcl_lastmodifiedbydate = oBlog("lastmodifiedbydate")
       lcl_lastmodifiedbyname = oBlog("lastmodifiedbyname")
    else
       response.redirect("mayorsblog_list.asp?success=NE")
    end if

    oBlog.close
    set oBlog = nothing

 end if

'Check for org features
 lcl_orghasfeature_rssfeeds_mayorsblog = orghasfeature("rssfeeds_mayorsblog")

'Check for user permissions
 lcl_userhaspermission_rssfeeds_mayorsblog = userhaspermission(session("userid"),"rssfeeds_mayorsblog")

'BEGIN: Build Created By info -------------------------------------------
 lcl_displayCreatedByInfo = ""

 if lcl_createdbyname <> "" then
    if lcl_displayCreatedByInfo <> "" then
       lcl_displayCreatedByInfo = lcl_displayCreatedByInfo & lcl_createdbyname
    else
       lcl_displayCreatedByInfo = lcl_createdbyname
    end if
 end if

 if lcl_createdbydate <> "" then
    if lcl_displayCreatedByInfo <> "" then
       lcl_displayCreatedByInfo = lcl_displayCreatedByInfo & " on " & lcl_createdbydate
    else
       lcl_displayCreatedByInfo = lcl_createdbydate
    end if
 end if
'END: Build Created By info ---------------------------------------------

'BEGIN: Build Last Modified By info -------------------------------------
 lcl_displayLastModifiedByInfo = ""

 if lcl_createdbyname <> "" then
    if lcl_displayLastModifiedByInfo <> "" then
       lcl_displayLastModifiedByInfo = lcl_displayLastModifiedByInfo & lcl_lastmodifiedbyname
    else
       lcl_displayLastModifiedByInfo = lcl_lastmodifiedbyname
    end if
 end if

 if lcl_createdbydate <> "" then
    if lcl_displayLastModifiedByInfo <> "" then
       lcl_displayLastModifiedByInfo = lcl_displayLastModifiedByInfo & " on " & lcl_lastmodifiedbydate
    else
       lcl_displayLastModifiedByInfo = lcl_lastmodifiedbydate
    end if
 end if
'END: Build Last Modified By info ---------------------------------------

'Determine if the blog owner has a picture.
 if lcl_imagefilename <> "" then
    'lcl_blogimage_url = session("egovclientwebsiteurl")
    'lcl_blogimage_url = lcl_blogimage_url & "/admin/custom/pub/"
    'lcl_blogimage_url = lcl_blogimage_url & session("virtualdirectory")
    'lcl_blogimage_url = lcl_blogimage_url & "/unpublished_documents"
    'lcl_blogimage_url = lcl_blogimage_url & lcl_imagefilename

    if left(lcl_imagefilename,1) <> "/" then
       lcl_imagefilename = "/" & lcl_imagefilename
    end if

    lcl_blogimage_url = Application("CommunityLink_DocUrl")
    lcl_blogimage_url = lcl_blogimage_url & "/public_documents300/"
    lcl_blogimage_url = lcl_blogimage_url & session("virtualdirectory")
    lcl_blogimage_url = lcl_blogimage_url & "/unpublished_documents"
    lcl_blogimage_url = lcl_blogimage_url & lcl_imagefilename
 else
    lcl_blogimage_url = session("egovclientwebsiteurl")
    lcl_blogimage_url = lcl_blogimage_url & "/admin/images/notavailable_person.jpg"
 end if

'Determine if the the active checkbox is "checked".
'The displaying of this field is a "reverse-negative".
'If the value is TRUE then the field is UNCHECKED.
'If the value is FALSE then the field IS CHECKED.
 if lcl_isInactive then
    lcl_checked_isInactive = ""
 else
    lcl_checked_isInactive = " checked=""checked"""
 end if

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Determine if there is any additional processing needed from the past update
 if lcl_orghasfeature_rssfeeds_mayorsblog AND lcl_userhaspermission_rssfeeds_mayorsblog AND (lcl_success = "SU" OR lcl_success = "SA") then
    if request("sendTo_RSS") <> "" then
       lcl_onload = lcl_onload & "sendToRSS('" & request("sendTo_RSS") & "');"
    end if
 end if

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
 lcl_hidden = "HIDDEN"
%>
<html>
<head>
  <title>E-Gov Administration Console {Blog Maintenance - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
var control_field = "";

function confirmDelete() {
  //var r = confirm('Are you sure you want to delete the "' + document.getElementById("title").value + '" blog entry?  \r NOTE: Any/All comments will be deleted as well.');
  var r = confirm('Are you sure you want to delete the "' + document.getElementById("blogtitle").value + '" blog entry?');
  if (r==true) {
      location.href="mayorsblog_delete.asp?blogid=<%=lcl_blogid%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count = 0;

  if(document.getElementById("blogtitle").value=="") {
     document.getElementById("blogtitle").focus();
     inlineMsg(document.getElementById("blogtitle").id,'<strong>Required Field Missing: </strong> Title',10,'blogtitle');
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("blogtitle");
  }

  if(lcl_false_count > 0) {
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("mayorsblog_maint").submit();
     return true;
  }
}

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

<% if lcl_orghasfeature_rssfeeds_mayorsblog AND lcl_userhaspermission_rssfeeds_mayorsblog then %>
function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&isAjax=Y';

  doAjax('mayorsblog_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
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
</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>

<!-- #include file="../menu/menu.asp" -->

<div id="centercontent">
<table border="0" cellspacing="0" cellpadding="10" width="800" class="start">
  <form name="mayorsblog_maint" id="mayorsblog_maint" method="post" action="mayorsblog_action.asp">
    <input type="<%=lcl_hidden%>" name="blogid" value="<%=lcl_blogid%>" size="5" maxlength="5" />
    <input type="<%=lcl_hidden%>" name="screen_mode" value="<%=lcl_screen_mode%>" size="4" maxlength="4" />
    <input type="<%=lcl_hidden%>" name="user_action" id="user_action" value="" size="4" maxlength="20" />
    <input type="<%=lcl_hidden%>" name="orgid" value="<%=lcl_orgid%>" size="4" maxlength="10" />
  <tr>
      <td>
          <font size="+1"><strong>Blog Maintenance: <%=lcl_screen_mode%></strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Back to Blog List" class="button" onclick="javascript:location.href='mayorsblog_list.asp';" />
      </td>
  </tr>
  <tr valign="top">
      <td>
          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <tr>
                <td align="left" style="font-size:10px;">
                    <% displayButtons "TOP", lcl_screen_mode %>
                </td>
                <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
            </tr>
          </table>
          <table border="0" cellspacing="0" cellpadding="3" class="tableadmin">
            <tr>
                <td rowspan="5">
                    <img src="<%=lcl_blogimage_url%>" name="blogImage" id="blogImage" width="100" height="100" style="border:1pt solid #000000" />
                </td>
                <td nowrap="nowrap">Blog Owner:</td>
                <td width="80%">
                    <select name="userid" id="userid">
                      <% displayBlogOwners lcl_userid %>
                    </select>
                </td>
            </tr>
            <tr>
                <td nowrap="nowrap">Title:</td>
                <td><input type="text" name="blogtitle" id="blogtitle" value="<%=lcl_title%>" size="50" maxlength="50" onchange="clearMsg('blogtitle');" /></td>
            </tr>
            <tr>
                <td nowrap="nowrap">Is Active?</td>
                <td><input type="checkbox" name="isInactive" value="on" id="isInactive"<%=lcl_checked_isInactive%>" /></td>
            </tr>
          <%
            if lcl_screen_mode = "EDIT" then
               response.write "<tr>" & vbcrlf
               response.write "    <td nowrap=""nowrap"">Created By:</td>" & vbcrlf
               response.write "    <td style=""color:#800000"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
               response.write "</tr>" & vbcrlf
               response.write "<tr>" & vbcrlf
               response.write "    <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
               response.write "    <td style=""color:#800000"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
               response.write "</tr>" & vbcrlf
            else
               response.write "<tr><td colspan=""2""></td></tr>" & vbcrlf
               response.write "<tr><td colspan=""2""></td></tr>" & vbcrlf
            end if

            if lcl_orghasfeature_rssfeeds_mayorsblog AND lcl_userhaspermission_rssfeeds_mayorsblog then
               response.write "  <tr valign=""top"">" & vbcrlf
               response.write "      <td>&nbsp;</td>" & vbcrlf
               response.write "      <td nowrap=""nowrap"">On " & lcl_sendToLabel & " Send To:</td>" & vbcrlf
               response.write "      <td>" & vbcrlf
                                         displaySendToOption "RSS", lcl_screen_mode, "N", lcl_orghasfeature_rssfeeds_mayorsblog, lcl_userhaspermission_rssfeeds_mayorsblog
               response.write "      </td>" & vbcrlf
               response.write "  </tr>" & vbcrlf
            end if
          %>
            <tr valign="bottom">
                <td>Blog Article:</td>
                <td colspan="2" align="right">
          					     <input type="button" value="Find a Link" class="button" onClick="doPicker('mayorsblog_maint.article','Y','Y','Y','Y');" style="margin-right:50px;" />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <textarea name="article" id="article" rows="30" cols="120" maxlength="8000"><%=lcl_article%></textarea>
                </td>
            </tr>
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
sub displayBlogOwners(iUserID)

  sSQL = "SELECT userid, firstname, lastname "
  sSQL = sSQL & " FROM users "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY lastname, firstname "

  set oBlogOwners = Server.CreateObject("ADODB.Recordset")
  oBlogOwners.Open sSQL, Application("DSN"), 3, 1

  if not oBlogOwners.eof then
     do while not oBlogOwners.eof

        if iUserID = oBlogOwners("userid") then
           lcl_selected = " selected=""selected"""
        else
           lcl_selected = ""
        end if

        response.write "  <option value=""" & oBlogOwners("userid") & """" & lcl_selected & ">" & oBlogOwners("firstname") & " " & oBlogOwners("lastname") & "</option>" & vbcrlf

        oBlogOwners.movenext
     loop
  end if

  oBlogOwners.close
  set oBlogOwners = nothing

end sub

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
        lcl_return = "Blog does not exist..."
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
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='mayorsblog_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" onclick=""return validateFields('ADDANOTHER');"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub
%>