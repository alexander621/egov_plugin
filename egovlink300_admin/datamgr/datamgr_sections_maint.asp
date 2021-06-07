<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_sections_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DataMgr Section
'
' MODIFICATION HISTORY
' 1.0  02/01/2011  David Boyer - Initial Version
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

'Retrieve the sectionid to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("sectionid") <> "" then
    lcl_sectionid = request("sectionid")

    if isnumeric(lcl_sectionid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "datamgr_sections_list.asp"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_sendToLabel = "Create"
    lcl_sectionid  = 0
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature     = "datamgr_maintain_sections"
 lcl_featurename = getFeatureName(lcl_feature)

 'if lcl_screen_mode = "ADD" then
 '   lcl_onload = lcl_onload & "validateFields('ADD');"
 'end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

'Set up local variables
 lcl_section_orgid              = 0
 lcl_sectionname                = ""
 lcl_sectiontype                = ""
 lcl_description                = ""
 lcl_isActive                   = 1
 lcl_isAccountInfoSection       = 0
 lcl_displaySectionName         = 1
 lcl_createdbyid                = 0
 lcl_createdbydate              = ""
 lcl_createdbyname              = ""
 lcl_lastmodifiedbyid           = 0
 lcl_lastmodifiedbydate         = ""
 lcl_lastmodifiedbyname         = ""
 lcl_checked_isactive           = " checked=""checked"""
 lcl_checked_isaccountinfo      = ""
 lcl_checked_displaysectionname = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the layout
    sSQL = "SELECT dms.sectionid, "
    sSQL = sSQL & " dms.sectionname, "
    sSQL = sSQL & " dms.sectiontype, "
    sSQL = sSQL & " dms.description, "
    sSQL = sSQL & " dms.isActive, "
    sSQL = sSQL & " dms.isAccountInfoSection, "
    sSQL = sSQL & " dms.displaySectionName, "
    sSQL = sSQL & " dms.section_orgid, "
    sSQL = sSQL & " dms.createdbyid, "
    sSQL = sSQL & " dms.createdbydate, "
    sSQL = sSQL & " dms.lastmodifiedbyid, "
    sSQL = sSQL & " dms.lastmodifiedbydate, "
    sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname "
    sSQL = sSQL & " FROM egov_dm_sections dms "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON dms.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dms.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE dms.sectionid = " & lcl_sectionid

    set oDMSections = Server.CreateObject("ADODB.Recordset")
    oDMSections.Open sSQL, Application("DSN"), 3, 1

    if not oDMSections.eof then
       lcl_sectionid            = oDMSections("sectionid")
       lcl_sectionname          = oDMSections("sectionname")
       lcl_sectiontype          = oDMSections("sectiontype")
       lcl_description          = oDMSections("description")
       lcl_isActive             = oDMSections("isActive")
       lcl_isAccountInfoSection = oDMSections("isAccountInfoSection")
       lcl_displaySectionName   = oDMSections("displaySectionName")
       lcl_section_orgid        = oDMSections("section_orgid")
       lcl_createdbyid          = oDMSections("createdbyid")
       lcl_createdbydate        = oDMSections("createdbydate")
       lcl_createdbyname        = oDMSections("createdbyname")
       lcl_lastmodifiedbyid     = oDMSections("lastmodifiedbyid")
       lcl_lastmodifiedbydate   = oDMSections("lastmodifiedbydate")
       lcl_lastmodifiedbyname   = oDMSections("lastmodifiedbyname")

      'Determine if the checkbox(es) are checked or not
       lcl_checked_isactive           = isCheckboxChecked(oDMSections("isActive"))
       lcl_checked_isaccountinfo      = isCheckboxChecked(oDMSections("isAccountInfoSection"))
       lcl_checked_displaysectionname = isCheckboxChecked(oDMSections("displaySectionName"))
    else
       response.redirect("datamgr_sections_list.asp?success=NE")
    end if

    oDMSections.close
    set oDMSections = nothing
 end if

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Show/Hide all "hidden" fields.  (HIDDEN = hide, TEXT = show)
 lcl_hidden = "hidden"

 dim lcl_scripts

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sectionid", lcl_sectionid)
%>
<html>
<head>
  <title>E-Gov Administration Console {Sections - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script language="javascript" src="../scripts/datamgr_fields_addrow.js"></script>

<script language="javascript">
function confirmDelete() {
  var r = confirm('Are you sure you want to delete this Section?');
  if (r==true) {

    <%
      lcl_delete_params = lcl_url_parameters
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "user_action", "DELETE")
    %>
      location.href="datamgr_sections_action.asp<%=lcl_delete_params%>";
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

function validateFields(p_action) {
  var lcl_false_count    = 0;
  var lcl_isFieldChecked = false;

  //---------------------------------------------------------------------------
  //Check the Map-Point Types Fields
  //---------------------------------------------------------------------------
  lcl_total_fields = document.getElementById("totalFields").value;
  lcl_i_start         = 1;

  for (i=lcl_total_fields; lcl_i_start<=i; -- i) {
       if(document.getElementById("fieldtype"+i).value == "") {
          //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
          if(document.getElementById("deleteField"+i).checked!=true && (document.getElementById("fieldtype"+i).value == "")) {
         			 inlineMsg(document.getElementById("fieldtype"+i).id,'<strong>Required Field Missing: </strong>Field Type',8,'fieldtype'+i);
             lcl_false_count = lcl_false_count + 1;
          }else{
             clearMsg('fieldtype'+i);
             document.getElementById("fieldtype"+i).value = '';
          }

          if(lcl_false_count == 1) {
             lcl_focus = document.getElementById("fieldtype"+i);
          }
       }else{
          clearMsg('fieldtype'+i);
     		}
  }

  //if(document.getElementById("sectionname").value=="") {
  //   inlineMsg(document.getElementById("sectionname").id,'<strong>Required Field Missing: </strong> Section Name',10,'sectionname');
  //   lcl_focus       = document.getElementById("sectionname");
  //   lcl_false_count = lcl_false_count + 1;
  //}else{
  //   clearMsg("sectionname");
  //}

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("datamgr_sections_maint").submit();
     return true;
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
<%
  response.write "<form name=""datamgr_sections_maint"" id=""datamgr_sections_maint"" method=""post"" action=""datamgr_sections_action.asp"">" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""sectionid"" id=""sectionid"" value=""" & lcl_sectionid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""screen_mode"" id=""screen_mode"" value=""" & lcl_screen_mode & """ size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""user_action"" id=""user_action"" value="""" size=""4"" maxlength=""20"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_featurename & ": " & lcl_screen_mode & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='datamgr_sections_list.asp" & lcl_url_parameters & "';"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf
                            displayButtons "TOP", lcl_screen_mode, lcl_return_parameters
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""left"">Section</th>" & vbcrlf
  response.write "                <th align=""right"" colspan=""3"">" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""isActive"" id=""isActive"" value=""Y""" & lcl_checked_isactive & " /> Active" & vbcrlf
  response.write "                </th>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Section Name:</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <input type=""text"" name=""sectionname"" id=""sectionname"" size=""40"" maxlength=""500"" value=""" & lcl_sectionname & """ onchange=""clearMsg('sectionname')"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Section Type:<br />(code use ONLY)</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <input type=""text"" name=""sectiontype"" id=""sectiontype"" size=""40"" maxlength=""100"" value=""" & lcl_sectiontype & """ onchange=""clearMsg('sectiontype')"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>&nbsp;</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""displaySectionName"" id=""displaySectionName"" value=""Y""" & lcl_checked_displaysectionname & " /> Display Section Name" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>&nbsp;</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""isAccountInfoSection"" id=""isAccountInfoSection"" value=""Y""" & lcl_checked_isaccountinfo & " /> Is an ""Account Info"" Section" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td>Description:<br />(internal use ONLY)</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <textarea name=""description"" id=""description"" maxlength=""4000"" style=""width:350px; height:100px;"">" & lcl_description & "</textarea>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Organization:<br />(internal use ONLY)</td>" & vbcrlf
  response.write "                <td colspan=""3"">" & vbcrlf
  response.write "                    <select name=""section_orgid"" id=""section_orgid"">" & vbcrlf
  response.write "                      <option value=""""></option>" & vbcrlf
                                        displayOrgOptions lcl_section_orgid
  response.write "                    </select>" & vbcrlf
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
  response.write "          </p>" & vbcrlf
  response.write "          <p>" & vbcrlf
                               lcl_isLimited     = False
                               lcl_isDisplayOnly = False
                               'displayMPTSectionFields "SECTION", 0, 0, lcl_sectionid, lcl_isRootAdmin, False, False
                               displaySectionFields session("orgid"), lcl_sectionid, lcl_isRootAdmin, lcl_isLimited, lcl_isDisplayOnly
                               displayButtons "BOTTOM", lcl_screen_mode, lcl_return_parameters
  response.write "          </p>" & vbcrlf
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
sub displaySectionFields(iOrgID, iSectionID, iIsRootAdmin, iIsLimited, iIsDisplayOnly)

  lcl_sectionid = iSectionID
  sSQL          = ""

  if not iIsDisplayOnly then
     response.write "<div style=""margin-top:20px; margin-bottom:5px;"">" & vbcrlf
     response.write "  <strong>" & lcl_sectiontitle & "</strong><br />" & vbcrlf
     response.write "  <input type=""button"" name=""addMPTField"" id=""addMPTField"" value=""Add Field"" class=""button"" onclick=""addFieldRow('" & iIsRootAdmin & "', '" & iIsLimited & "', 'addFieldTBL','totalFields','addFieldRow', '" & lcl_edit_type & "');"" />" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "<table id=""addFieldTBL"" border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "  <tr id=""addFieldRow0"">" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "      <th align=""left"">Label</th>" & vbcrlf
  response.write "      <th>Active</th>" & vbcrlf
  response.write "      <th>Display<br />Order</th>" & vbcrlf
  response.write "      <th>Display<br />Label</th>" & vbcrlf
  response.write "      <th>Display as<br />Multi-Line</th>" & vbcrlf
  response.write "      <th>Include<br />""Add a Link""</th>" & vbcrlf
  response.write "      <th>Remove</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  iRowCount   = 0
  lcl_bgcolor = "#ffffff"

  sSQL = "SELECT section_fieldid, "
  sSQL = sSQL & " sectionid, "
  sSQL = sSQL & " isActive, "
  sSQL = sSQL & " fieldname, "
  sSQL = sSQL & " fieldtype, "
  sSQL = sSQL & " displayFieldName, "
  sSQL = sSQL & " isMultiLine, "
  sSQL = sSQL & " hasAddLinkButton, "
  sSQL = sSQL & " displayOrder "
  sSQL = sSQL & " FROM egov_dm_sections_fields "
  sSQL = sSQL & " WHERE sectionid = " & lcl_sectionid
  sSQL = sSQL & " ORDER BY displayOrder "

  set oSectionFields = Server.CreateObject("ADODB.Recordset")
  oSectionFields.Open sSQL, Application("DSN"), 3, 1

  if not oSectionFields.eof then
     do while not oSectionFields.eof

        iRowCount                    = iRowCount + 1
        lcl_bgcolor                  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_checked_displayFieldName = isCheckboxChecked(oSectionFields("displayFieldName"))
        lcl_checked_isMultiLine      = isCheckboxChecked(oSectionFields("isMultiLine"))
        lcl_checked_hasAddLinkButton = isCheckboxChecked(oSectionFields("hasAddLinkButton"))
        lcl_checked_isActive         = isCheckboxChecked(oSectionFields("isActive"))
        lcl_displayOrder             = oSectionFields("displayOrder")

        response.write "  <tr id=""addFieldRow" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ align=""center"">" & vbcrlf
        response.write "      <td align=""left"">" & iRowCount & ".</td>" & vbcrlf
        response.write "      <td align=""right"">" & vbcrlf
        response.write "          <input type=""text"" name=""fieldname" & iRowCount & """ id=""fieldname" & iRowCount & """ value=""" & oSectionFields("fieldname") & """ size=""50"" maxlength=""100"" onchange=""clearMsg('fieldname" & iRowCount & "');"" />" & vbcrlf

        if iIsRootAdmin and not iIsLimited then
           response.write "<br /><strong>Field Type: </strong>(code use ONLY)&nbsp;" & vbcrlf
           response.write "<input type=""text"" name=""fieldtype" & iRowCount & """ id=""fieldtype" & iRowCount & """ value=""" & oSectionFields("fieldtype") & """ size=""15"" maxlength=""100"" onchange=""clearMsg('fieldtype" & iRowCount & "')"" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "      <td><input type=""checkbox"" name=""sectionfield_isActive" & iRowCount & """ id=""sectionfield_isActive" & iRowCount & """ value=""Y""" & lcl_checked_isActive & " /></td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""text"" name=""displayOrder" & iRowCount & """ id=""displayOrder" & iRowCount & """ value=""" & lcl_displayOrder & """ size=""3"" maxlength=""5"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td><input type=""checkbox"" name=""displayFieldName" & iRowCount & """ id=""displayFieldName" & iRowCount & """ value=""Y""" & lcl_checked_displayFieldName & " /></td>" & vbcrlf
        response.write "      <td><input type=""checkbox"" name=""isMultiLine" & iRowCount & """ id=""isMultiLine" & iRowCount & """ value=""1""" & lcl_checked_isMultiLine & " /></td>" & vbcrlf
        response.write "      <td><input type=""checkbox"" name=""hasAddLinkButton" & iRowCount & """ id=""hasAddLinkButton" & iRowCount & """ value=""1""" & lcl_checked_hasAddLinkButton & " /></td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""hidden"" name=""section_fieldid" & iRowCount & """ id=""section_fieldid" & iRowCount & """ value=""" & oSectionFields("section_fieldid") & """ />" & vbcrlf

        if not iIsRootAdmin or (iIsRootAdmin AND iIsLimited) then
           response.write "          <input type=""hidden"" name=""fieldtype" & iRowCount & """ id=""fieldtype" & iRowCount & """ value=""" & oSectionFields("fieldtype") & """ size=""20"" maxlength=""100"" />" & vbcrlf
        end if

        response.write "          <input type=""checkbox"" name=""deleteField" & iRowCount & """ id=""deleteField" & iRowCount & """ value=""Y"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oSectionFields.movenext
     loop
  end if

  oSectionFields.close
  set oSectionFields = nothing

  response.write "</table>" & vbcrlf
  response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""" & iRowCount & """ size=""3"" maxlength=""100"" />" & vbcrlf

end sub

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
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='datamgr_sections_list.asp" & iReturnParameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayOrgOptions(iOrgID)
 	dim sSQL, sOrgID

  sOrgID = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

	 sSQL = "SELECT orgid, "
  sSQL = sSQL & " orgcity "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE isdeactivated = 0 "
  sSQL = sSQL & " ORDER BY orgcity "

 	set oOrgOptions = Server.CreateObject("ADODB.Recordset")
 	oOrgOptions.Open sSQL, Application("DSN"), 3, 1

 	if not oOrgOptions.eof then
     do while not oOrgOptions.eof
        if sOrgID = oOrgOptions("orgid") then
           lcl_selected_orgid = " selected=""selected"""
        else
           lcl_selected_orgid = ""
        end if

        response.write "<option value=""" & oOrgOptions("orgid") & """" & lcl_selected_orgid & ">" & oOrgOptions("orgcity") & " [" & oOrgOptions("orgid") & "]</option>" & vbcrlf

        oOrgOptions.movenext
     loop
  end if

  oOrgOptions.close
  set oOrgOptions = nothing

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