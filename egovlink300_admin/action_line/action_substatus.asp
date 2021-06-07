<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_substatus.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the maintain action line sub-status
'
' MODIFICATION HISTORY
' 1.0 08/13/2007  David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"action_line_substatus") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Set up BODY onload
 lcl_onload  = lcl_onload & "$('#newSubStatus').focus();"
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Administration {Action Line: Sub-Status}</title>
	
	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	 <link rel="stylesheet" type="text/css" href="../global.css" />

<style>
  #addSubStatus td,
  #editSubStatus th,
  #editSubStatus td {
     white-space: nowrap;
  }

  #editSubStatusButtons {
     padding-bottom: 5px;
  }

  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
  }

  .newParentStatusRow {
     border-top: 1pt solid #000000;
  }
</style>

  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
<!--
$(document).ready(function() {
  $('#addButton').click(function() {
    if($('#newSubStatus').val() == '') {
       $('#newSubStatus').focus();
       inlineMsg(document.getElementById("newSubStatus").id,'<strong>Required Field Missing: </strong> Sub-Status Name.',10,'newSubStatus');
       return false;
    } else {
       $('#action').val('ADD');
       $('#subStatusMaint').submit();
    }
  });

  $('input[name*="saveChangesButton"]').click(function() {
     var lcl_total_substatuses = Number(0);
     var lcl_false_count       = Number(0);

     if($('#total_substatuses').val() != '') {
        lcl_total_substatuses = Number($('#total_substatuses').val());
     }

     for (var i = parseInt(lcl_total_substatuses); i >= 1 ; i--) {
        if($('#editSubStatus_' + i).val() == '') {
           $('#editSubStatus_' + i).focus();
           inlineMsg(document.getElementById('editSubStatus_' + i).id,'<strong>Required Field Missing: </strong> Sub-Status.',10,'editSubStatus_' + i);
           lcl_false_count = lcl_false_count + 1;
        }
     }

     if(lcl_false_count > 0) {
        return false;
     } else {
        $('#action').val('EDIT');
        $('#subStatusMaint').submit();
     }
  });
});

function moveStatus(iDirection, iLineNumber) {
  var lcl_direction   = '';
  var lcl_line_number = 0;

  if(iDirection != '' && iDirection != undefined) {
     lcl_direction = iDirection.toUpperCase();
  }

  if(iLineNumber != '' && iLineNumber != undefined) {
     lcl_line_number = Number(iLineNumber);
  }

  $('#action').val(lcl_direction);
  $('#action_linenumber').val(lcl_line_number)
  $('#subStatusMaint').submit();

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

function fnCheckCategory() {
  if (document.DeleteEventCategory.Category.value != '0') {
      return true;
  }else{
      return false;
  }
}

function fnCheckModify() {
  if (((document.ModifyEventCategory.Category.value != '0') || (document.ModifyEventCategory.CustomCategory.value != ''))) {
        return true;
  }else{
        return false;
  }
}

function confirm_delete(iLineNumber) {
  var lcl_line_number = 0;
  var lcl_status_name = '';

  if(iLineNumber != '' && iLineNumber != undefined) {
     lcl_line_number = iLineNumber;
  }

  lcl_status_name = $('#editSubStatus_' + lcl_line_number).val();

  input_box = confirm('Are you sure you want to delete "' + lcl_status_name + '"?');
  if(input_box==true) { 
     $('#action_linenumber').val(lcl_line_number);
     $('#action').val('DELETE');
     $('#subStatusMaint').submit();
  }
}
//-->
</script>
</head>
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""subStatusMaint"" id=""subStatusMaint"" method=""post"" action=""action_substatus_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""action"" id=""action"" value="""" size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""action_linenumber"" id=""action_linenumber"" size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Action Line: Sub-Status</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf

 'BEGIN: New Sub-Status -------------------------------------------------------
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "            <table id=""addSubStatus"" width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin"">" & vbcrlf
  response.write "              <tr><th align=""left"" colspan=""5"">Create a Sub-Status</th></tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td valign=""top"">Sub-Status:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""newSubStatus"" id=""newSubStatus"" size=""20"" maxlength=""50"" onchange=""clearMsg('newSubStatus')"" /></td>" & vbcrlf
  response.write "                  <td valign=""top"">Parent Status:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""newParentStatus"" id=""newParentStatus"">" & vbcrlf
                                          lcl_parent_status_new = ""

                                          displayParentStatusOptions lcl_parent_status_new
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td width=""70%"">" & vbcrlf
  response.write "                      <input type=""button"" name=""addButton"" id=""addButton"" value=""Add Sub-Status"" class=""button"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
 'END: New Sub-Status ---------------------------------------------------------

 'BEGIN: Edit Sub-Status ------------------------------------------------------
  response.write "          <p>" & vbcrlf
  response.write "          <div id=""editSubStatusButtons"">" & vbcrlf
  response.write "            <input type=""button"" name=""saveChangesButton"" id=""saveChangesButton"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "            <table id=""editSubStatus"" cellpadding=""5"" cellspacing=""0"" border=""0"" class=""tableadmin"">" & vbcrlf
  response.write "          		  <tr>" & vbcrlf
  response.write "          		      <th align=""left"">Parent Status</th>" & vbcrlf
  response.write "                  <th align=""left"">Sub-Status</th>" & vbcrlf
  response.write "             			  <th>Active</th>" & vbcrlf
  response.write "          			     <th>Change Parent Status to</th>" & vbcrlf
  response.write "                  <th>&nbsp;</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Retrieve all of the sub-statuses for the organization for each parent_status
  sSQLs = "SELECT s1.action_status_id, "
		sSQLs = sSQLs & " s1.status_name, "
		sSQLs = sSQLs & " s1.orgid, "
		sSQLs = sSQLs & " s1.parent_status, "
		sSQLs = sSQLs & " s1.display_order, "
		sSQLs = sSQLs & " s1.active_flag, "
		sSQLs = sSQLs & " s2.action_status_id AS parent_status_id, "
		sSQLs = sSQLs & " s2.display_order AS parent_display_order "
		sSQLs = sSQLs & " FROM egov_actionline_requests_statuses s1 "
		sSQLs = sSQLs &      " INNER JOIN egov_actionline_requests_statuses s2 ON s1.parent_status = s2.status_name "
		sSQLs = sSQLs & " WHERE s1.orgid = " & session("orgid")
  sSQLs = sSQLs & " AND s2.parent_status = 'MAIN' "
		sSQLs = sSQLs & " ORDER BY s2.display_order, s2.action_status_id, s1.display_order, s1.status_name "

		set oSubStatus = Server.CreateObject("ADODB.Recordset")
  oSubStatus.Open sSQLs, Application("DSN"), 3, 1

  i              = 0
		lcl_line_count = 0
  lcl_bgcolor    = "#eeeeee"

		if not oSubStatus.eof then
		   do while not oSubStatus.eof
        lcl_line_count = lcl_line_count + 1
						  i = i + 1

        lcl_display_parent_status    = "&nbsp;"
        lcl_new_parentrow_status     = ""
        lcl_new_parentrow_class      = ""
        lcl_selected_active_flag_yes = ""
        lcl_selected_active_flag_no  = " selected=""selected"""
        lcl_statusExistsOnRequest    = false

						  if ucase(lcl_parent_status) <> ucase(oSubStatus("parent_status")) then
           lcl_display_parent_status = "<strong>" & oSubStatus("parent_status") & "</strong>"

   							 if lcl_line_count > 1 then
			      					lcl_line_count = 1
              lcl_bgcolor    = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
           end if
						  end if

        if i > 1 AND lcl_line_count = 1 then
           lcl_new_parentrow_class = " class=""newParentStatusRow"""
        end if

        if oSubStatus("active_flag") = "Y" then
           lcl_selected_active_flag_yes = " selected=""selected"""
           lcl_selected_active_flag_no  = ""
        end if

        lcl_statusExistsOnRequest = checkSubStatusExistsOnRequest(session("orgid"), oSubStatus("action_status_id"))

        response.write "              <tr bgcolor=""" & lcl_bgcolor & """ align=""center"" valign=""top"">" & vbcrlf
        response.write "                  <td align=""left""" & lcl_new_parentrow_class & ">" & lcl_display_parent_status & "</td>" & vbcrlf
        response.write "                  <td" & lcl_new_parentrow_class & ">" & vbcrlf
        response.write "                      <input type=""hidden"" name=""editActionStatusID_" & i & """ id=""editActionStatusID_" & i & """ value=""" & oSubStatus("action_status_id") & """ size=""5"" maxlength=""10"" />" & vbcrlf
        response.write "                      <input type=""hidden"" name=""editParentOriginal_" & i & """ id=""editParentOriginal_" & i & """ value=""" & oSubStatus("parent_status") & """ size=""10"" maxlength=""20"" />" & vbcrlf
        response.write "                      <input type=""text"" name=""editSubStatus_" & i & """ id=""editSubStatus_" & i & """ value=""" & oSubStatus("status_name") & """ style=""width:133px;"" maxlength=""50"" />" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "                  <td" & lcl_new_parentrow_class & ">" & vbcrlf
        response.write "                      <select name=""editActive_" & i & """ id=""editActive_" & i & """>" & vbcrlf
        response.write "                        <option value=""Y""" & lcl_selected_active_flag_yes & ">Yes</option>" & vbcrlf
        response.write "                        <option value=""N""" & lcl_selected_active_flag_no  & ">No</option>" & vbcrlf
        response.write "                      </select>" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "                  <td" & lcl_new_parentrow_class & ">" & vbcrlf
        response.write "                      <select name=""editParentStatus_" & i & """ id=""editParentStatus_" & i & """>" & vbcrlf
                                                displayParentStatusOptions oSubStatus("parent_status")
        response.write "                      </select>" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "                  <td align=""left""" & lcl_new_parentrow_class & ">" & vbcrlf
        response.write "                      <input type=""button"" name=""moveUpButton"   & i & """ id=""moveUpButton"   & i & """ value=""Move Up"" class=""button"" onclick=""moveStatus('MOVEUP','" & i & "');"" />" & vbcrlf
        response.write "                      <input type=""button"" name=""moveDownButton" & i & """ id=""moveDownButton" & i & """ value=""Move Down"" class=""button"" onclick=""moveStatus('MOVEDOWN','" & i & "');"" />" & vbcrlf

        if not lcl_statusExistsOnRequest then
           response.write "                      <input type=""button"" name=""deleteButton" & i & """ id=""deleteButton" & i & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & i & "');"" />" & vbcrlf
        end if

        response.write "                  </td>" & vbcrlf
        response.write "              </tr>" & vbcrlf

        lcl_parent_status = ucase(oSubStatus("parent_status"))

						  oSubStatus.movenext
		   loop
		end if

  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <div id=""editSubStatusButtons"">" & vbcrlf
  response.write "            <input type=""button"" name=""saveChangesButton2"" id=""saveChangesButton2"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          </p>" & vbcrlf
  response.write "          <input type=""hidden"" name=""total_substatuses"" id=""total_substatuses"" value=""" & i & """ size=""5"" maxlength=""10"" />" & vbcrlf
 'END: Edit Sub-Status --------------------------------------------------------

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayParentStatusOptions(iParentStatus)

  dim sParentStatus, lcl_selected_parent_status

  if iParentStatus <> "" then
     if not containsApostrophe(iParentStatus) then
        sParentStatus = ucase(trim(iParentStatus))
     end if
  end if

  lcl_selected_parent_status = ""

  sSQLso = "SELECT action_status_id, "
  sSQLso = sSQLso & " status_name as parentstatus, "
  sSQLso = sSQLso & " orgid, "
  sSQLso = sSQLso & " parent_status, "
  sSQLso = sSQLso & " display_order, "
  sSQLso = sSQLso & " active_flag "
  sSQLso = sSQLso & " FROM egov_actionline_requests_statuses "
  sSQLso = sSQLso & " WHERE orgid = 0 "
  sSQLso = sSQLso & " AND parent_status = 'MAIN' "
  sSQLso = sSQLso & " AND active_flag = 'Y' "
  sSQLso = sSQLso & " ORDER BY display_order, status_name "

  set oMainStatus = Server.CreateObject("ADODB.Recordset")
  oMainStatus.Open sSQLso, Application("DSN"), 3, 1

		if not oMainStatus.eof then
				 do while not oMainStatus.eof
        lcl_parentStatusName       = ""
        lcl_selected_parent_status = ""

        if oMainStatus("parentstatus") <> "" then
           lcl_parentStatusName = ucase(trim(oMainStatus("parentstatus")))
        end if

        if lcl_parentStatusName = sParentStatus then
           lcl_selected_parent_status = " selected=""selected"""
        end if

        response.write "  <option value=""" & lcl_parentStatusName & """" & lcl_selected_parent_status & ">" & lcl_parentStatusName & "</option>" & vbcrlf

        oMainStatus.movenext
     loop
  else
     response.write "  <option value="">No Parent Statuses Available</option>" & vbcrlf
  end if

  oMainStatus.close
  set oMainStatus = nothing

end sub

'------------------------------------------------------------------------------
function checkSubStatusExistsOnRequest(iOrgID, iSubStatusID)
  dim lcl_return, sOrgID, sSubStatusID

  lcl_return   = false
  sOrgID       = 0
  sSubStatusID = 0

  if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iSubStatusID <> "" then
     if not containsApostrophe(iSubStatusID) then
        sSubStatusID = clng(iSubStatusID)
     end if
  end if

  if sOrgID > 0 AND sSubStatusID > 0 then
     sSQL = "SELECT count(action_autoid) as total_requests "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE orgid = " & sOrgID
     sSQL = sSQL & " AND sub_status_id = " & sSubStatusID

     set oSubStatusCheck = Server.CreateObject("ADODB.Recordset")
     oSubStatusCheck.Open sSQL, Application("DSN"), 3, 1

     if not oSubStatusCheck.eof then
        if oSubStatusCheck("total_requests") > 0 then
           lcl_return = true
        end if
     end if

     set oSubStatusCheck = nothing

  end if

  checkSubStatusExistsOnRequest = lcl_return

end function
%>