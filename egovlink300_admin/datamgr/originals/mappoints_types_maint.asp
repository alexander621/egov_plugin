<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints_types_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a Blog entry
'
' MODIFICATION HISTORY
' 1.0 03/05/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("egov_administration") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"egovadmin_mappointtypes_maint") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the mappoint_typeid to be maintain.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("mappoint_typeid") <> "" then
    lcl_mappoint_typeid = request("mappoint_typeid")

    if isnumeric(lcl_mappoint_typeid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "mappoints_types_list.asp"
    end if
 else
    lcl_screen_mode     = "ADD"
    lcl_sendToLabel     = "Create"
    lcl_mappoint_typeid = 0
 end if

'Set up local variables
 lcl_mappointtype       = ""
 lcl_description        = ""
 lcl_feature            = ""
 lcl_feature_maintain   = ""
 lcl_createdbyid        = 0
 lcl_createdbydate      = ""
 lcl_createdbyname      = ""
 lcl_lastmodifiedbyid   = 0
 lcl_lastmodifiedbydate = ""
 lcl_lastmodifiedbyname = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the blog
    sSQL = "SELECT t.mappoint_typeid, t.mappointtype, t.description, t.feature, t.feature_maintain, t.createdbyid, t.createdbydate, "
    sSQL = sSQL & " t.lastmodifiedbyid, t.lastmodifiedbydate, u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname "
    sSQL = sSQL & " FROM egov_mappoints_types t "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON t.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON t.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE t.mappoint_typeid = " & lcl_mappoint_typeid

    set oMPTypes = Server.CreateObject("ADODB.Recordset")
    oMPTypes.Open sSQL, Application("DSN"), 3, 1

    if not oMPTypes.eof then
       lcl_mappointtype       = oMPTypes("mappointtype")
       lcl_description        = oMPTypes("description")
       lcl_feature            = oMPTypes("feature")
       lcl_feature_maintain   = oMPTypes("feature_maintain")
       lcl_createdbyid        = oMPTypes("createdbyid")
       lcl_createdbydate      = oMPTypes("createdbydate")
       lcl_createdbyname      = oMPTypes("createdbyname")
       lcl_lastmodifiedbyid   = oMPTypes("lastmodifiedbyid")
       lcl_lastmodifiedbydate = oMPTypes("lastmodifiedbydate")
       lcl_lastmodifiedbyname = oMPTypes("lastmodifiedbyname")
    else
       response.redirect("mappoints_types_list.asp?success=NE")
    end if

    oMPTypes.close
    set oMPTypes = nothing

 end if

'Check for org features
' lcl_orghasfeature_rssfeeds_mayorsblog = orghasfeature("rssfeeds_mayorsblog")

'Check for user permissions
' lcl_userhaspermission_rssfeeds_mayorsblog = userhaspermission(session("userid"),"rssfeeds_mayorsblog")

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

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Map-Point Types Maintenance - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="javascript" src="../scripts/ajaxLit.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
var control_field = "";

function confirmDelete() {
  //var r = confirm('Are you sure you want to delete the "' + document.getElementById("title").value + '" blog entry?  \r NOTE: Any/All comments will be deleted as well.');
  var r = confirm('Are you sure you want to delete: "' + document.getElementById("description").value + '"');
  if (r==true) {
      location.href="mappoints_types_action.asp?user_action=DELETE&mappoint_typeid=<%=lcl_mappoint_typeid%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count = 0;

  if(document.getElementById("feature").value=="") {
     inlineMsg(document.getElementById("feature").id,'<strong>Required Field Missing: </strong> Feature',10,'feature');
     lcl_focus       = document.getElementById("feature");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("feature");
  }

  if(document.getElementById("mappointtype").value=="") {
     inlineMsg(document.getElementById("mappointtype").id,'<strong>Required Field Missing: </strong> MapPointType (code)',10,'mappointtype');
     lcl_focus       = document.getElementById("mappointtype");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("mappointtype");
  }

  if(document.getElementById("description").value=="") {
     inlineMsg(document.getElementById("description").id,'<strong>Required Field Missing: </strong> Description',10,'description');
     lcl_focus       = document.getElementById("description");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("description");
  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("mappoints_types_maint").submit();
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
  <form name="mappoints_types_maint" id="mappoints_types_maint" method="post" action="mappoints_types_action.asp">
    <input type="hidden" name="mappoint_typeid" value="<%=lcl_mappoint_typeid%>" size="5" maxlength="5" />
    <input type="hidden" name="screen_mode" value="<%=lcl_screen_mode%>" size="4" maxlength="4" />
    <input type="hidden" name="user_action" id="user_action" value="" size="4" maxlength="20" />
    <input type="hidden" name="orgid" value="<%=lcl_orgid%>" size="4" maxlength="10" />
  <tr>
      <td>
          <font size="+1"><strong>Map-Point Types Maintenance: <%=lcl_screen_mode%></strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Back to List" class="button" onclick="location.href='mappoints_types_list.asp';" />
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
                <th align="left" colspan="2">&nbsp;</th>
            </tr>
            <tr>
                <td nowrap="nowrap">Description:</td>
                <td><input type="text" name="description" id="description" value="<%=lcl_description%>" size="50" maxlength="500" onchange="clearMsg('description');" /></td>
            </tr>
            <tr>
                <td nowrap="nowrap">MapPointType (code):</td>
                <td><input type="text" name="mappointtype" id="mappointtype" value="<%=lcl_mappointtype%>" size="10" maxlength="50" onchange="clearMsg('mappointtype');" /></td>
            </tr>
            <tr>
                <td nowrap="nowrap">Feature:</td>
                <td><input type="text" name="feature" id="feature" value="<%=lcl_feature%>" size="50" maxlength="50" onchange="clearMsg('feature');" /></td>
            </tr>
            <tr>
                <td nowrap="nowrap">Feature (maintain):</td>
                <td><input type="text" name="feature_maintain" id="feature_maintain" value="<%=lcl_feature_maintain%>" size="50" maxlength="50" onchange="clearMsg('feature_maintain');" /></td>
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
          %>
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
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='mappoints_types_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf

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