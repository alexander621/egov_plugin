<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_layouts_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DataMgr Layout
'
' MODIFICATION HISTORY
' 1.0  01/28/2011  David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel          = "../"  'Override of value from common.asp
 lcl_isRootAdmin = False
 lcl_onload      = ""

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

'Retrieve the layoutid to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("layoutid") <> "" then
    lcl_layoutid = request("layoutid")

    if isnumeric(lcl_layoutid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "datamgr_layouts_list.asp"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_sendToLabel = "Create"
    lcl_layoutid  = 0
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature     = "datamgr_maintain_layouts"
 lcl_featurename = getFeatureName(lcl_feature)

 'if lcl_screen_mode = "ADD" then
 '   lcl_onload = lcl_onload & "validateFields('ADD');"
 'end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

'Set up local variables
 lcl_orgid                     = session("orgid")
 lcl_layoutname                = ""
 lcl_isActive                  = 1
 lcl_useLayoutSections         = 1
 lcl_totalcolumns              = "0"
 lcl_columnwidth_left          = "100"
 lcl_columnwidth_middle        = "0"
 lcl_columnwidth_right         = "0"
 lcl_createdbyid               = 0
 lcl_createdbydate             = ""
 lcl_createdbyname             = ""
 lcl_lastmodifiedbyid          = 0
 lcl_lastmodifiedbydate        = ""
 lcl_lastmodifiedbyname        = ""
 lcl_checked_isactive          = " checked=""checked"""
 lcl_checked_useLayoutSections = " checked=""checked"""
 lcl_selected_totalcolumns_1   = " selected=""selected"""
 lcl_selected_totalcolumns_2   = ""
 lcl_selected_totalcolumns_3   = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the layout
    sSQL = "SELECT dml.layoutid, "
    sSQL = sSQL & " dml.layoutname, "
    sSQL = sSQL & " dml.useLayoutSections, "
    sSQL = sSQL & " dml.isActive, "
    sSQL = sSQL & " dml.totalcolumns, "
    sSQL = sSQL & " dml.columnwidth_left, "
    sSQL = sSQL & " dml.columnwidth_middle, "
    sSQL = sSQL & " dml.columnwidth_right, "
    sSQL = sSQL & " dml.createdbyid, "
    sSQL = sSQL & " dml.createdbydate, "
    sSQL = sSQL & " dml.lastmodifiedbyid, "
    sSQL = sSQL & " dml.lastmodifiedbydate, "
    sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname "
    sSQL = sSQL & " FROM egov_dm_layouts dml "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON dml.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dml.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE dml.layoutid = " & lcl_layoutid

    set oDMLayouts = Server.CreateObject("ADODB.Recordset")
    oDMLayouts.Open sSQL, Application("DSN"), 3, 1

    if not oDMLayouts.eof then
       lcl_layoutid           = oDMLayouts("layoutid")
       lcl_layoutname         = oDMLayouts("layoutname")
       lcl_useLayoutSections  = oDMLayouts("useLayoutSections")
       lcl_isActive           = oDMLayouts("isActive")
       lcl_totalcolumns       = oDMLayouts("totalcolumns")
       lcl_columnwidth_left   = oDMLayouts("columnwidth_left")
       lcl_columnwidth_middle = oDMLayouts("columnwidth_middle")
       lcl_columnwidth_right  = oDMLayouts("columnwidth_right")
       lcl_createdbyid        = oDMLayouts("createdbyid")
       lcl_createdbydate      = oDMLayouts("createdbydate")
       lcl_createdbyname      = oDMLayouts("createdbyname")
       lcl_lastmodifiedbyid   = oDMLayouts("lastmodifiedbyid")
       lcl_lastmodifiedbydate = oDMLayouts("lastmodifiedbydate")
       lcl_lastmodifiedbyname = oDMLayouts("lastmodifiedbyname")

      'Determine if the checkbox(es) are checked or not
       if not oDMLayouts("isActive") then
          lcl_checked_isactive = ""
       end if

       if not oDMLayouts("useLayoutSections") then
          lcl_checked_useLayoutSections = ""
       end if

      'Determine which option is selected
       if lcl_totalcolumns = 2 then
          lcl_selected_totalcolumns_1 = ""
          lcl_selected_totalcolumns_2 = " selected=""selected"""
          lcl_selected_totalcolumns_3 = ""
       elseif lcl_totalcolumns = 3 then
          lcl_selected_totalcolumns_1 = ""
          lcl_selected_totalcolumns_2 = ""
          lcl_selected_totalcolumns_3 = " selected=""selected"""
       else
          lcl_selected_totalcolumns_1 = " selected=""selected"""
          lcl_selected_totalcolumns_2 = ""
          lcl_selected_totalcolumns_3 = ""
       end if
    else
       response.redirect("datamgr_layouts_list.asp?success=NE")
    end if

    oDMLayouts.close
    set oDMLayouts = nothing
 end if

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

 lcl_onload = lcl_onload & "enableDisableColumnWidthFields();"

'Check for a screen message
 lcl_success = request("success")
 'lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Show/Hide all "hidden" fields.  (HIDDEN = hide, TEXT = show)
 lcl_hidden = "hidden"

 dim lcl_scripts

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "layoutid", lcl_layoutid)
%>
<html>
<head>
  <title>E-Gov Administration Console {Layouts - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">

function enableDisableColumnWidthFields() {

  if(document.getElementById("totalcolumns")) {
     lcl_totalcolumns = document.getElementById("totalcolumns").value;
  }

  document.getElementById("columnwidth_left").disabled   = true;
  document.getElementById("columnwidth_middle").disabled = true;
  document.getElementById("columnwidth_right").disabled  = true;

  if(lcl_totalcolumns == 2) {
     document.getElementById("columnwidth_left").disabled  = false;
     document.getElementById("columnwidth_right").disabled = false;
  } else if(lcl_totalcolumns == 3) {
     document.getElementById("columnwidth_left").disabled   = false;
     document.getElementById("columnwidth_middle").disabled = false;
     document.getElementById("columnwidth_right").disabled  = false;
  } else {
     document.getElementById("columnwidth_left").disabled = false;
  }
}

function confirmDelete() {
  var r = confirm('Are you sure you want to delete this MapPoint Layout?');
  if (r==true) {

    <%
      lcl_delete_params = lcl_url_parameters
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "user_action", "DELETE")
    %>
      location.href="datamgr_layouts_action.asp<%=lcl_delete_params%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count = 0;

//  if(document.getElementById("description").value=="") {
//     inlineMsg(document.getElementById("description").id,'<strong>Required Field Missing: </strong> Description',10,'description');
//     lcl_focus       = document.getElementById("description");
//     lcl_false_count = lcl_false_count + 1;
//  }else{
//     clearMsg("description");
//  }

  //---------------------------------------------------------------------------
  //Check the Map-Point Category Fields
  //---------------------------------------------------------------------------
//  lcl_total_fields = document.getElementById("totalFields").value;
//  lcl_i_start         = 1;

//  for (i=lcl_total_fields; lcl_i_start<=i; -- i) {
//       if(document.getElementById("fieldname"+i).value == "") {
          //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
//          if(document.getElementById("deleteField"+i).checked!=true && (document.getElementById("fieldname"+i).value != "" || document.getElementById("displayInResults"+i).value != "" || document.getElementById("resultsOrder"+i).value != "")) {
//         			 inlineMsg(document.getElementById("fieldname"+i).id,'<strong>Required Field Missing: </strong>Field Name',8,'fieldname'+i);
//             lcl_false_count = lcl_false_count + 1;
//          }else{
//             clearMsg('fieldname'+i);
//             document.getElementById("fieldname"+i).value = '';
//          }

//          if(lcl_false_count == 1) {
//             lcl_focus = document.getElementById("fieldname"+i);
//          }
//       }else{
//          clearMsg('fieldname'+i);
//     		}
//  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("datamgr_layouts_maint").submit();
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

function openWin(p_url, p_width, p_height) {
  w = 600;
  h = 400;

  if((p_width!="")&&(p_width!=undefined)) {
      w = p_width;
  }

  if((p_height!="")&&(p_height!=undefined)) {
      h = p_height;
  }

  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_url = p_url;

  eval('window.open("' + lcl_url + '", "_picker", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
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
<%
  response.write "<form name=""datamgr_layouts_maint"" id=""datamgr_layouts_maint"" method=""post"" action=""datamgr_layouts_action.asp"">" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""layoutid"" id=""layoutid"" value=""" & lcl_layoutid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""screen_mode"" id=""screen_mode"" value=""" & lcl_screen_mode & """ size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""user_action"" id=""user_action"" value="""" size=""4"" maxlength=""20"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_featurename & ": " & lcl_screen_mode & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='datamgr_layouts_list.asp" & lcl_url_parameters & "';"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf
                            displayButtons "TOP", lcl_screen_mode, lcl_return_parameters
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""left"">DM Layout</th>" & vbcrlf
  response.write "                <th align=""right"" colspan=""3"">&nbsp;" & vbcrlf
  response.write "                </th>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Layout Name:</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <input type=""text"" name=""layoutname"" id=""layoutname"" size=""40"" maxlength=""100"" value=""" & lcl_layoutname & """ />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>&nbsp;</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""isActive"" id=""isActive"" value=""Y""" & lcl_checked_isactive & " /> Active" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""useLayoutSections"" id=""useLayoutSections"" value=""Y""" & lcl_checked_useLayoutSections & " /> Use Layout Sections" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Total Columns:</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <select name=""totalcolumns"" id=""totalcolumns"" onchange=""enableDisableColumnWidthFields()"">" & vbcrlf
  response.write "                      <option value=""1""" & lcl_selected_totalcolumns_1 & ">1</option>" & vbcrlf
  response.write "                      <option value=""2""" & lcl_selected_totalcolumns_2 & ">2</option>" & vbcrlf
  response.write "                      <option value=""3""" & lcl_selected_totalcolumns_3 & ">3</option>" & vbcrlf
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td>Column Width:<br />(in pixels)</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                      <tr align=""center"">" & vbcrlf
  response.write "                          <td>" & vbcrlf
  response.write "                              <input type=""text"" name=""columnwidth_left"" id=""columnwidth_left"" size=""5"" maxlength=""10"" value=""" & lcl_columnwidth_left & """ /><br />" & vbcrlf
  response.write "                              <span style=""font-size:10px"">LEFT</span>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                          <td>" & vbcrlf
  response.write "                              <input type=""text"" name=""columnwidth_middle"" id=""columnwidth_middle"" size=""5"" maxlength=""10"" value=""" & lcl_columnwidth_middle & """ /><br />" & vbcrlf
  response.write "                              <span style=""font-size:10px"">MIDDLE</span>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                          <td>" & vbcrlf
  response.write "                              <input type=""text"" name=""columnwidth_right"" id=""columnwidth_right"" size=""5"" maxlength=""10"" value=""" & lcl_columnwidth_right & """ /><br />" & vbcrlf
  response.write "                              <span style=""font-size:10px"">RIGHT</span>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  if lcl_screen_mode = "EDIT" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td nowrap=""nowrap"" style=""height:15px"">Created By:</td>" & vbcrlf
     response.write "      <td style=""color:#800000"" colspan=""3"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
     response.write "      <td style=""color:#800000"" colspan=""3"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  response.write "          </table>" & vbcrlf
                            displayButtons "BOTTOM", lcl_screen_mode, lcl_return_parameters
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

 'Determine if there are any scripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if
%>
<!--#include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode, iReturnParameters)

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

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='datamgr_layouts_list.asp" & iReturnParameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
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