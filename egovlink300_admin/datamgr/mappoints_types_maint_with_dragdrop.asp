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

 sLevel     = "../"  'Override of value from common.asp
 lcl_onload = ""

'Determine if there is a specific feature associated to the Map-Point Category
 lcl_mappoint_typeid = ""
 lcl_feature         = "mappoints_types_maint"
 lcl_isLimited       = False
 lcl_pagetitle       = "Map-Point Category"
 lcl_sectiontitle    = lcl_pagetitle & ": Fields"

 if request("f") <> "" AND request("f") <> "mappoints_types_maint" then
    lcl_feature      = request("f")
    lcl_isLimited    = True
    lcl_pagetitle    = getFeatureName(lcl_feature)
    lcl_sectiontitle = lcl_pagetitle

   'Retrieve the MapPoint_TypeID
    if request("mappoint_typeid") <> "" then
       lcl_mappoint_typeid = request("mappoint_typeid")
    else
       lcl_mappoint_typeid = getMapPointTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

       if lcl_mappoint_typeid = 0 then
         	response.redirect sLevel & "permissiondenied.asp"
       end if
    end if
 else
    if request("mappoint_typeid") <> "" then
       lcl_mappoint_typeid = request("mappoint_typeid")
    end if
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 else
    lcl_isRootAdmin = False
 end if

'Retrieve the mappoint_typeid to be maintain.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 'if request("mappoint_typeid") <> "" then
 if lcl_mappoint_typeid <> "" then
    'lcl_mappoint_typeid = request("mappoint_typeid")

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
 lcl_description             = ""
 lcl_isActive                = 1
 lcl_createdbyid             = 0
 lcl_createdbydate           = ""
 lcl_createdbyname           = ""
 lcl_lastmodifiedbyid        = 0
 lcl_lastmodifiedbydate      = ""
 lcl_lastmodifiedbyname      = ""
 lcl_mappointcolor           = "green"
 lcl_feature_public          = ""
 lcl_feature_maintain        = ""
 lcl_feature_maintain_fields = ""
 lcl_checked_isactive        = " checked=""checked"""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the blog
    sSQL = "SELECT t.mappoint_typeid, t.description, t.isActive, "
    sSQL = sSQL & " t.createdbyid, t.createdbydate, "
    sSQL = sSQL & " t.lastmodifiedbyid, t.lastmodifiedbydate, "
    sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
    sSQL = sSQL & " t.mappointcolor, t.feature_public, t.feature_maintain, t.feature_maintain_fields "
    sSQL = sSQL & " FROM egov_mappoints_types t "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON t.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON t.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE t.mappoint_typeid = " & lcl_mappoint_typeid

    set oMPTypes = Server.CreateObject("ADODB.Recordset")
    oMPTypes.Open sSQL, Application("DSN"), 3, 1

    if not oMPTypes.eof then
       lcl_description             = oMPTypes("description")
       lcl_isActive                = oMPTypes("isActive")
       lcl_createdbyid             = oMPTypes("createdbyid")
       lcl_createdbydate           = oMPTypes("createdbydate")
       lcl_createdbyname           = oMPTypes("createdbyname")
       lcl_lastmodifiedbyid        = oMPTypes("lastmodifiedbyid")
       lcl_lastmodifiedbydate      = oMPTypes("lastmodifiedbydate")
       lcl_lastmodifiedbyname      = oMPTypes("lastmodifiedbyname")
       lcl_mappointcolor           = oMPTypes("mappointcolor")
       lcl_feature_public          = oMPTypes("feature_public")
       lcl_feature_maintain        = oMPTypes("feature_maintain")
       lcl_feature_maintain_fields = oMPTypes("feature_maintain_fields")

      'If the Map-Point Type is NOT "active" then do NOT "check" the checkbox
       if not lcl_isActive then
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
 lcl_canDelete = False

 if lcl_isRootAdmin then
    lcl_canDelete = True
    'lcl_canDelete = checkForMapPointsByMapPointTypeID(lcl_mappoint_typeid)

    lcl_onload = lcl_onload & "displayFeature('feature_public','displayfeature_public');"
    lcl_onload = lcl_onload & "displayFeature('feature_maintain','displayfeature_maintain');"
    lcl_onload = lcl_onload & "displayFeature('feature_maintain_fields','displayfeature_maintain_fields');"
 end if

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

 lcl_onload = lcl_onload & "dragDropSetup();"

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_pagetitle%> - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

 	<script language="javascript" src="../scripts/ajaxLit.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script language="javascript" src="../scripts/mappoints_dragdrop.js"></script>

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
  var d = mytbl.insertCell(3);  //Display On Info Page
  var e = mytbl.insertCell(4);  //Results Order
  var f = mytbl.insertCell(5);  //Include "Add a Link"
  var g = mytbl.insertCell(6);  //Display as Multi-Line
  var h = mytbl.insertCell(7);  //Delete Row and Additional Info

  //Build the cells in the new row.
  //Row Count
  a.innerHTML = totalrows + '. ';

  //Field Name
  b.innerHTML = '<input type="text" name="fieldname' + totalrows + '" id="fieldname' +  totalrows + '" value="" size="50" maxlength="100" onchange="clearMsg(\'fieldname' + totalrows + '\');" />';

  //Display In Results
  c.align     = "center";
  c.innerHTML = '<input type="checkbox" name="displayInResults' + totalrows + '" id="displayInResults' + totalrows + '" value="1" checked="checked" />';

  //Display On Info Page
  d.align     = "center";
  d.innerHTML = '<input type="checkbox" name="displayInInfoPage' + totalrows + '" id="displayInInfoPage' + totalrows + '" value="1" checked="checked" />';

  //Results Order
  e.align     = "center";
  e.innerHTML = '<input type="text" name="resultsOrder' + totalrows + '" id="resultsOrder' + totalrows + '" value="' + totalrows + '" size="3" maxlength="5" />';

  //Include "Add a Link"
  f.align     = "center";
  f.innerHTML = '<input type="checkbox" name="hasAddLinkButton' + totalrows + '" id="hasAddLinkButton' + totalrows + '" value="1" />';

  //Include "Add a Link"
  g.align     = "center";
  g.innerHTML = '<input type="checkbox" name="isMultiLine' + totalrows + '" id="isMultiLine' + totalrows + '" value="1" />';

  //Delete Row and Additional Info
  var lcl_delete_row  = '';
      lcl_delete_row += '<input type="hidden" name="mp_fieldid' + totalrows + '" id="mp_fieldid' + totalrows + '" value="" />';
      lcl_delete_row += '<input type="checkbox" name="deleteField' + totalrows + '" id="deleteField' + totalrows + '" value="Y" />';
      //lcl_delete_row += '<input type="hidden" name="deleteField' + totalrows + '" id="deleteField' + totalrows + '" value="N" size="1" maxlength="1" />';
      //lcl_delete_row += '<input type="button" name="deleteButton' + totalrows + '" id="deleteButton' + totalrows + '" value="Delete" class="button" onclick="deleteField(\'' + totalrows + '\');" />';

  h.align     = "center";
  h.innerHTML = lcl_delete_row;

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

function displayFeature(p_fieldid, p_displayid) {
  document.getElementById(p_displayid).innerHTML='[' + document.getElementById(p_fieldid).value + ']'
}

function dragDropSetup() {
	// Create our helper object that will show the item while dragging
	dragHelper = document.createElement('DIV');
	dragHelper.style.cssText = 'position:absolute;display:none;';

	CreateDragContainer(
		document.getElementById('DragContainer1')
		//document.getElementById('DragContainer1'),
		//document.getElementById('DragContainer2'),
		//document.getElementById('DragContainer3')
	);

	document.body.appendChild(dragHelper);
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
  <form name="mappoints_types_maint" id="mappoints_types_maint" method="post" action="mappoints_types_action.asp">
    <input type="hidden" name="mappoint_typeid" value="<%=lcl_mappoint_typeid%>" size="5" maxlength="5" />
    <input type="hidden" name="screen_mode" value="<%=lcl_screen_mode%>" size="4" maxlength="4" />
    <input type="hidden" name="user_action" id="user_action" value="" size="4" maxlength="20" />
    <input type="hidden" name="orgid" value="<%=session("orgid")%>" size="4" maxlength="10" />
    <input type="hidden" name="f" value="<%=lcl_feature%>" size="10" maxlength="50" />
<%
  if not lcl_isRootAdmin then
     response.write "<input type=""hidden"" name=""feature_public"" id=""feature_public"" value=""" & lcl_feature_public & """ />" & vbcrlf
     response.write "<input type=""hidden"" name=""feature_maintain"" id=""feature_maintain"" value=""" & lcl_feature_maintain & """ />" & vbcrlf
     response.write "<input type=""hidden"" name=""feature_maintain_fields"" id=""feature_maintain_fields"" value=""" & lcl_feature_maintain_fields & """ />" & vbcrlf
  end if
%>
  <tr>
      <td>
          <font size="+1"><strong><%=lcl_pagetitle%>: <%=lcl_screen_mode%></strong></font><br />
          <%
            if not lcl_isLimited then
               response.write "<input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='mappoints_types_list.asp';"" />" & vbcrlf
            end if
          %>
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
  <tr valign="top">
      <td colspan="2">
          <p>
          <% displayButtons "TOP", lcl_screen_mode, lcl_canDelete, lcl_isLimited %>
          <table border="0" cellspacing="0" cellpadding="3" class="tableadmin">
            <tr>
                <th align="left"><%=lcl_pagetitle%></th>
                <th align="right">&nbsp;</th>
            </tr>
            <tr>
                <td nowrap="nowrap">Description:</td>
                <td><input type="text" name="description" id="description" value="<%=lcl_description%>" size="50" maxlength="500" onchange="clearMsg('description');" /></td>
            </tr>
<%
  response.write "  <tr>" & vbcrlf
  response.write "      <td nowrap=""nowrap"">Map-Point Color:</td>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <select name=""mappointcolor"" id=""mappointcolor"">" & vbcrlf
                              displayMapPointColors lcl_mappointcolor
  response.write "          </select>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td nowrap=""nowrap"">Active:</td>" & vbcrlf
  response.write "      <td><input type=""checkbox"" name=""isActive"" id=""isActive"" value=""Y""" & lcl_checked_isActive & "/></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  if lcl_screen_mode = "EDIT" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td nowrap=""nowrap"">Created By:</td>" & vbcrlf
     response.write "      <td style=""color:#800000"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
     response.write "      <td style=""color:#800000"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  else
     response.write "  <tr><td colspan=""2""></td></tr>" & vbcrlf
     response.write "  <tr><td colspan=""2""></td></tr>" & vbcrlf
  end if

  if lcl_isRootAdmin then
     response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td>Feature (public-side):</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <select name=""feature_public"" id=""feature_public"" onchange=""displayFeature('feature_public','displayfeature_public');"">" & vbcrlf
     response.write "            <option value=""""></option>" & vbcrlf
                                 showFeatureOptions lcl_feature_public
     response.write "          </select>" & vbcrlf
     response.write "          <img src=""../images/help.jpg"" name=""helpFeature_public"" id=""helpFeature_public"" alt=""Used to connect the Map-Points public displayed feature (home page option - i.e. Available Properties feature) to this Map-Point Type (category)."" /><br />" & vbcrlf
     response.write "          <span id=""displayfeature_public"" style=""color:#800000;""></span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td>Feature (maintenance):</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <select name=""feature_maintain"" id=""feature_maintain"" onchange=""displayFeature('feature_maintain','displayfeature_maintain');"">" & vbcrlf
     response.write "            <option value=""""></option>" & vbcrlf
                                 showFeatureOptions lcl_feature_maintain
     response.write "          </select>" & vbcrlf
     response.write "          <img src=""../images/help.jpg"" name=""helpFeature_maintain"" id=""helpFeature_maintain"" alt=""Used to determine which feature/permission to use when accessing the maintenance screen."" /><br />" & vbcrlf
     response.write "          <span id=""displayfeature_maintain"" style=""color:#800000;""></span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td>Feature (maintenance - fields):</td>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <select name=""feature_maintain_fields"" id=""feature_maintain_fields"" onchange=""displayFeature('feature_maintain_fields','displayfeature_maintain_fields');"">" & vbcrlf
     response.write "            <option value=""""></option>" & vbcrlf
                                 showFeatureOptions lcl_feature_maintain_fields
     response.write "          </select>" & vbcrlf
     response.write "          <img src=""../images/help.jpg"" name=""helpFeature_maintain_fields"" id=""helpFeature_maintain_fields"" alt=""Used to determine which feature/permission to use when accessing the Map-Point Types maintenance screen for admins.  It will limit the screen to this specific category from the navigation menu."" /><br />" & vbcrlf
     response.write "          <span id=""displayfeature_maintain_fields"" style=""color:#800000;""></span>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
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
            displayButtons "BOTTOM", lcl_screen_mode, lcl_canDelete, lcl_isLimited
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

  sSQL = "SELECT mp_fieldid, mappoint_typeid, orgid, fieldname, isnull(fieldtype,'') as fieldtype, "
  sSQL = sSQL & " hasAddLinkButton, isMultiLine, displayInResults, displayInInfoPage, resultsOrder "
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
     response.write "  <strong>" & lcl_sectiontitle & "</strong><br />" & vbcrlf
     'response.write "  <input type=""button"" name=""reorderButton"" id=""reorderButton"" value=""Maintain Results Field Order"" class=""button"" onclick=""alert('coming soon');"" />" & vbcrlf
     response.write "  <input type=""button"" name=""addMPTField"" id=""addMPTField"" value=""Add Field"" class=""button"" onclick=""addFieldRow();"" />" & vbcrlf
     response.write "</div>" & vbcrlf
     'response.write "<div class=""DragContainer"" id=""DragContainer1"">" & vbcrlf
     response.write "<table id=""addFieldTBL"" border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
     response.write "  <tr id=""addFieldRow0"">" & vbcrlf
     response.write "      <th align=""left"" colspan=""2"">Field Name</th>" & vbcrlf
     response.write "      <th>Display In<br />Results</th>" & vbcrlf
     response.write "      <th>Display On<br />Info Page</th>" & vbcrlf
     response.write "      <th>Display<br />Order</th>" & vbcrlf
     response.write "      <th>Include<br />""Add a Link""</th>" & vbcrlf
     response.write "      <th>Display as<br />Multi-Line</th>" & vbcrlf
     response.write "      <th>Remove</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>"
     5/3/2010response.write "<div class=""DragContainer"" id=""DragContainer1"">" & vbcrlf

     do while not oMPTFields.eof

        iRowCount                     = iRowCount + 1
        lcl_bgcolor                   = changeBGColor(lcl_bgcolor,"eeeeee","ffffff")
        lcl_checked_hasAddLinkButton  = isCheckboxChecked(oMPTFields("hasAddLinkButton"))
        lcl_checked_isMultiLine       = isCheckboxChecked(oMPTFields("isMultiLine"))
        lcl_checked_displayInResults  = isCheckboxChecked(oMPTFields("displayInResults"))
        lcl_checked_displayInInfoPage = isCheckboxChecked(oMPTFields("displayInInfoPage"))
        lcl_resultsOrder              = oMPTFields("resultsOrder")

        response.write "<div id=""Item" & iRowCount & """ class=""DragBox"" overClass=""OverDragBox"" dragClass=""DragDragBox"">" & vbcrlf
        response.write "[click here to drag]"
        response.write "<table>" & vbcrlf
        response.write "  <tr id=""addFieldRow" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ align=""center"">" & vbcrlf
        response.write "      <td align=""left"">" & iRowCount & ".</td>" & vbcrlf
        response.write "      <td align=""left"">" & vbcrlf
        response.write "          <input type=""text"" name=""fieldname" & iRowCount & """ id=""fieldname" & iRowCount & """ value=""" & oMPTFields("fieldname") & """ size=""50"" maxlength=""100"" onchange=""clearMsg('fieldname" & iRowCount & "');"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInResults" & iRowCount & """ id=""displayInResults" & iRowCount & """ value=""1""" & lcl_checked_displayInResults & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInInfoPage" & iRowCount & """ id=""displayInInfoPage" & iRowCount & """ value=""1""" & lcl_checked_displayInInfoPage & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""text"" name=""resultsOrder" & iRowCount & """ id=""resultsOrder" & iRowCount & """ value=""" & lcl_resultsOrder & """ size=""3"" maxlength=""5"" />" & vbcrlf
        response.write "      </td>" & vbcrlf

        if oMPTFields("fieldtype") = "" then
           response.write "      <td><input type=""checkbox"" name=""hasAddLinkButton" & iRowCount & """ id=""hasAddLinkButton" & iRowCount & """ value=""1""" & lcl_checked_hasAddLinkButton & " /></td>" & vbcrlf
           response.write "      <td><input type=""checkbox"" name=""isMultiLine" & iRowCount & """ id=""isMultiLine" & iRowCount & """ value=""1""" & lcl_checked_isMultiLine & " /></td>" & vbcrlf
        else
           response.write "      <td><input type=""hidden"" name=""hasAddLinkButton" & iRowCount & """ id=""hasAddLinkButton" & iRowCount & """ value=""0"" /></td>" & vbcrlf
           response.write "      <td><input type=""hidden"" name=""isMultiLine" & iRowCount & """ id=""isMultiLine" & iRowCount & """ value=""0"" /></td>" & vbcrlf
        end if

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
        response.write "</table>" & vbcrlf
        response.write "</div>"

        oMPTFields.movenext
     loop

     'response.write "</table>" & vbcrlf
     'response.write "</div>" & vbcrlf
  end if

  response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""" & iRowCount & """ size=""3"" maxlength=""100"" />" & vbcrlf

  oMPTFields.close
  set oMPTFields = nothing

end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode, iCanDelete, iIsLimited)

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

  if not iIsLimited then
     response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='mappoints_types_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf
  end if

  if lcl_screen_mode = "ADD" then
     'response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" onclick=""return validateFields('ADDANOTHER');"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     if iCanDelete AND not iIsLimited then
        response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     end if

     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub

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