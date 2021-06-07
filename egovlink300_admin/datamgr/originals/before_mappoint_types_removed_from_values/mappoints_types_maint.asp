<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
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
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"mappoints_types_maint") then
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
 lcl_description        = ""
 lcl_isInactive         = 0
 lcl_createdbyid        = 0
 lcl_createdbydate      = ""
 lcl_createdbyname      = ""
 lcl_lastmodifiedbyid   = 0
 lcl_lastmodifiedbydate = ""
 lcl_lastmodifiedbyname = ""
 lcl_checked_isActive   = " checked=""checked"""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the blog
    sSQL = "SELECT t.mappoint_typeid, t.description, t.isInactive, t.createdbyid, t.createdbydate, "
    sSQL = sSQL & " t.lastmodifiedbyid, t.lastmodifiedbydate, u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname "
    sSQL = sSQL & " FROM egov_mappoints_types t "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON t.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON t.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE t.mappoint_typeid = " & lcl_mappoint_typeid

    set oMPTypes = Server.CreateObject("ADODB.Recordset")
    oMPTypes.Open sSQL, Application("DSN"), 3, 1

    if not oMPTypes.eof then
       lcl_description        = oMPTypes("description")
       lcl_isInactive         = oMPTypes("isInactive")
       lcl_createdbyid        = oMPTypes("createdbyid")
       lcl_createdbydate      = oMPTypes("createdbydate")
       lcl_createdbyname      = oMPTypes("createdbyname")
       lcl_lastmodifiedbyid   = oMPTypes("lastmodifiedbyid")
       lcl_lastmodifiedbydate = oMPTypes("lastmodifiedbydate")
       lcl_lastmodifiedbyname = oMPTypes("lastmodifiedbyname")

      'If the Map-Point Type IS "inactive" then do NOT "check" the checkbox
      'If the Map-Point Type is NOT "inactive" then DO "check" the checkbox
       if lcl_isInactive then
          lcl_checked_isactive = ""
       end if
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

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

'Check for associated Map-Points to determine if this MapPointType can/cannot be deleted.
 lcl_canDelete = checkForMapPointsByMapPointTypeID(lcl_mappoint_typeid)

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {Map-Point Categories - <%=lcl_screen_mode%>}</title>

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

//  if(document.getElementById("feature").value=="") {
//     inlineMsg(document.getElementById("feature").id,'<strong>Required Field Missing: </strong> Feature',10,'feature');
//     lcl_focus       = document.getElementById("feature");
//     lcl_false_count = lcl_false_count + 1;
//  }else{
//     clearMsg("feature");
//  }

//  if(document.getElementById("mappointtype").value=="") {
//     inlineMsg(document.getElementById("mappointtype").id,'<strong>Required Field Missing: </strong> MapPointType (code)',10,'mappointtype');
//     lcl_focus       = document.getElementById("mappointtype");
//     lcl_false_count = lcl_false_count + 1;
//  }else{
//     clearMsg("mappointtype");
//  }

  if(document.getElementById("description").value=="") {
     inlineMsg(document.getElementById("description").id,'<strong>Required Field Missing: </strong> Description',10,'description');
     lcl_focus       = document.getElementById("description");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("description");
  }

  //---------------------------------------------------------------------------
  //Check the Map-Point Category Fields
  //---------------------------------------------------------------------------
  lcl_total_fields = document.getElementById("totalFields").value;
  lcl_i_start         = 1;

  for (i=lcl_total_fields; lcl_i_start<=i; -- i) {
       if(document.getElementById("fieldname"+i).value == "") {
          //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
          if(document.getElementById("deleteField"+i).checked!=true && (document.getElementById("fieldname"+i).value != "" || document.getElementById("displayInResults"+i).value != "" || document.getElementById("resultsOrder"+i).value != "")) {
         			 inlineMsg(document.getElementById("fieldname"+i).id,'<strong>Required Field Missing: </strong>Field Name',8,'fieldname'+i);
             lcl_false_count = lcl_false_count + 1;
          }else{
             clearMsg('fieldname'+i);
             document.getElementById("fieldname"+i).value = '';
          }

          if(lcl_false_count == 1) {
             lcl_focus = document.getElementById("fieldname"+i);
          }
       }else{
          clearMsg('fieldname'+i);
     		}
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

function addFieldRow() {
  var mytbl     = document.getElementById('addFieldTBL');
  var totalrows = Number(document.getElementById("totalFields").value);

  //Increase the total rows by one.  This is index for the new row.
  totalrows = totalrows+1;

  //Set up the new row.
  mytbl = document.getElementById('addFieldTBL').insertRow(totalrows);

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg   = "";
  var lcl_evenodd = totalrows/2;
      lcl_evenodd = lcl_evenodd.toString();

  if(lcl_evenodd.indexOf('.') > 0) {
     lcl_rowbg = "#eeeeee";
  }else{
     lcl_rowbg = "#ffffff";
  }

  mytbl.style.background = lcl_rowbg;

  //Build the cells for the new row.
  var a = mytbl.insertCell(0);  //Row Count
  var b = mytbl.insertCell(1);  //Field Name
  var c = mytbl.insertCell(2);  //Display In Results
  var d = mytbl.insertCell(3);  //Results Order
  var e = mytbl.insertCell(4);  //Delete Row and Additional Info

  //Build the cells in the new row.
  //Row Count
  a.innerHTML = totalrows + '. ';

  //Field Name
  b.innerHTML = '<input type="text" name="fieldname' + totalrows + '" id="fieldname' +  totalrows + '" value="" size="50" maxlength="100" onchange="clearMsg(\'fieldname' + totalrows + '\');" />';

  //Display In Results
  c.align     = "center";
  c.innerHTML = '<input type="checkbox" name="displayInResults' + totalrows + '" id="displayInResults' + totalrows + '" value="1" checked="checked" />';

  //Results Order
  d.align     = "center";
  d.innerHTML = '<input type="text" name="resultsOrder' + totalrows + '" id="resultsOrder' + totalrows + '" value="' + totalrows + '" size="3" maxlength="5" />';

  //Delete Row and Additional Info
  var lcl_delete_row  = '';
      lcl_delete_row += '<input type="hidden" name="mp_fieldid' + totalrows + '" id="mp_fieldid' + totalrows + '" value="" />';
      lcl_delete_row += '<input type="checkbox" name="deleteField' + totalrows + '" id="deleteField' + totalrows + '" value="Y" />';
      //lcl_delete_row += '<input type="hidden" name="deleteField' + totalrows + '" id="deleteField' + totalrows + '" value="N" size="1" maxlength="1" />';
      //lcl_delete_row += '<input type="button" name="deleteButton' + totalrows + '" id="deleteButton' + totalrows + '" value="Delete" class="button" onclick="deleteField(\'' + totalrows + '\');" />';

  e.align     = "center";
  e.innerHTML = lcl_delete_row;

  //update the total row count.
  document.getElementById("totalFields").value = totalrows;
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

//function deleteField(iRowID) {
//  document.getElementById("deleteField" + iRowID).value = "Y";
//  document.getElementById("addFieldRow" + iRowID).style.display = "none";
//}

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
    <input type="hidden" name="orgid" value="<%=session("orgid")%>" size="4" maxlength="10" />
  <tr>
      <td>
          <font size="+1"><strong>Map-Point Categories: <%=lcl_screen_mode%></strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Back to List" class="button" onclick="location.href='mappoints_types_list.asp';" />
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
  <tr valign="top">
      <td colspan="2">
          <p>
          <% displayButtons "TOP", lcl_screen_mode, lcl_canDelete %>
          <table border="0" cellspacing="0" cellpadding="3" class="tableadmin">
            <tr>
                <th align="left">Map-Point Category</th>
                <th align="right"><input type="checkbox" name="isActive" id="isActive" value="Y"<%=lcl_checked_isActive%> />Active</th>
            </tr>
            <tr>
                <td nowrap="nowrap">Description:</td>
                <td><input type="text" name="description" id="description" value="<%=lcl_description%>" size="50" maxlength="500" onchange="clearMsg('description');" /></td>
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
          </p>
          &nbsp;
          <p>
          <%
           'Retrieve any/all fields related to this Map-Point Type
            displayMPTypesFields session("orgid"), lcl_mappoint_typeid

           'Display the bottom row of buttons
            displayButtons "BOTTOM", lcl_screen_mode, lcl_canDelete
          %>
          </p>
      </td>
  </tr>
</table>
</div>

<%
 'Determine if there are any scripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if
%>

<!--#include file="../admin_footer.asp"-->

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displayMPTypesFields(iOrgID, iMapPointTypeID)

  iRowCount   = 0
  lcl_bgcolor = "ffffff"

  sSQL = "SELECT mp_fieldid, mappoint_typeid, orgid, fieldname, isnull(fieldtype,'') as fieldtype, displayInResults, resultsOrder "
  sSQL = sSQL & " FROM egov_mappoints_types_fields "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID
  sSQL = sSQL & " ORDER BY resultsOrder, mp_fieldid "
  'sSQL = "SELECT mp_fieldid, mappoint_typeid, orgid, fieldname, displayInResults, resultsOrder, 1 as displayListOrder "
  'sSQL = sSQL & " FROM egov_mappoints_types_fields "
  'sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID
  'sSQL = sSQL & " AND resultsOrder > 0 "
  'sSQL = sSQL & " UNION ALL "
  'sSQL = sSQL & " SELECT mp_fieldid, mappoint_typeid, orgid, fieldname, displayInResults, resultsOrder, 2 as displayListOrder "
  'sSQL = sSQL & " FROM egov_mappoints_types_fields "
  'sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID
  'sSQL = sSQL & " AND resultsOrder = 0 "
  'sSQL = sSQL & " ORDER BY 7, resultsOrder, mp_fieldid "

  set oMPTFields = Server.CreateObject("ADODB.Recordset")
  oMPTFields.Open sSQL, Application("DSN"), 3, 1

  if not oMPTFields.eof then
     response.write "<div style=""margin-bottom:5px;"">" & vbcrlf
     response.write "  <strong>Map-Point Category: Fields</strong><br />" & vbcrlf
     'response.write "  <input type=""button"" name=""reorderButton"" id=""reorderButton"" value=""Maintain Results Field Order"" class=""button"" onclick=""alert('coming soon');"" />" & vbcrlf
     response.write "  <input type=""button"" name=""addMPTField"" id=""addMPTField"" value=""Add Field"" class=""button"" onclick=""addFieldRow();"" />" & vbcrlf
     response.write "</div>" & vbcrlf
     response.write "<table id=""addFieldTBL"" border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
     response.write "  <tr id=""addFieldRow0"">" & vbcrlf
     response.write "      <th align=""left"" colspan=""2"">Field Name</th>" & vbcrlf
     response.write "      <th>Display<br />In Results</th>" & vbcrlf
     response.write "      <th>Results<br />Order</th>" & vbcrlf
     response.write "      <th>Remove</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oMPTFields.eof

        iRowCount   = iRowCount + 1
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"eeeeee","ffffff")

        if oMPTFields("displayInResults") then
           lcl_checked_displayInResults = " checked=""checked"""
           'lcl_resultsOrder             = oMPTFields("resultsOrder")
 '          lcl_resultsOrder             = iRowCount
        else
           lcl_checked_displayInResults = ""
'           lcl_resultsOrder             = oMPTFields("resultsOrder")
        end if

           lcl_resultsOrder             = oMPTFields("resultsOrder")

        response.write "  <tr id=""addFieldRow" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ align=""center"">" & vbcrlf
        response.write "      <td align=""left"">" & iRowCount & ".</td>" & vbcrlf
        response.write "      <td align=""left"">" & vbcrlf
        response.write "          <input type=""text"" name=""fieldname" & iRowCount & """ id=""fieldname" & iRowCount & """ value=""" & oMPTFields("fieldname") & """ size=""50"" maxlength=""100"" onchange=""clearMsg('fieldname" & iRowCount & "');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInResults" & iRowCount & """ id=""displayInResults" & iRowCount & """ value=""1""" & lcl_checked_displayInResults & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""text"" name=""resultsOrder" & iRowCount & """ id=""resultsOrder" & iRowCount & """ value=""" & lcl_resultsOrder & """ size=""3"" maxlength=""5"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""hidden"" name=""mp_fieldid" & iRowCount & """ id=""mp_fieldid" & iRowCount & """ value=""" & oMPTFields("mp_fieldid") & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""fieldtype" & iRowCount & """ id=""fieldtype" & iRowCount & """ value=""" & oMPTFields("fieldtype") & """ />" & vbcrlf

        if oMPTFields("fieldtype") = "" then
           response.write "          <input type=""checkbox"" name=""deleteField" & iRowCount & """ id=""deleteField" & iRowCount & """ value=""Y"" />" & vbcrlf
           'response.write "          <input type=""button"" name=""deleteButton" & iRowCount & """ id=""deleteButton" & iRowCount & """ value=""Delete"" class=""button"" onclick=""deleteField('" & iRowCount & "');"" />" & vbcrlf
        else
           response.write "          <input type=""hidden"" name=""deleteField" & iRowCount & """ id=""deleteField" & iRowCount & """ value=""N"" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oMPTFields.movenext
     loop

     response.write "</table>" & vbcrlf
  end if

  response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""" & iRowCount & """ size=""3"" maxlength=""100"" />" & vbcrlf

  oMPTFields.close
  set oMPTFields = nothing

end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode, iCanDelete)

  if iTopBottom <> "" then
     iTopBottom = UCASE(iTopBottom)
  else
     iTopBottom = "TOP"
  end if

  if iTopBottom = "BOTTOM" then
     lcl_style_div = "margin-top:5px;"
  else
     lcl_style_div = "margin-bottom:5px;"
  end if

  'lcl_return_parameters = "?sc_org_name=" & session("sc_org_name") & "&sc_show_members=" & session("sc_show_members")
  lcl_return_parameters = ""

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='mappoints_types_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" onclick=""return validateFields('ADDANOTHER');"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     if iCanDelete then
        response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     end if

     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

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
%>