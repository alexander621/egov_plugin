<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_code_sections.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the maintain action line code sections
'
' MODIFICATION HISTORY
' 1.0 08/22/2007 David Boyer - INITIAL VERSION
' 1.1 08/01/2008 David Boyer - Added javascript field length edit.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "action_line_code_sections") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 dim oCmd, oRst, dDate, iDuration, sTimeZones, sLinks, bShown
 dim lcl_new_code_section, lcl_new_description

'Set the description character limit based off of the length of the column on the table (egov_actionline_code_sections.description)
 lcl_desc_char_length = 3000

'Set up BODY onload
 lcl_onload = "setMaxLength();"
 lcl_onload = lcl_onload & "document.getElementById('newCode').focus();"

'Check for success message
 lcl_success = ""
 lcl_msg     = ""

 if request("success") <> "" then
    if not containsApostrophe(request("success")) then
       lcl_success = ucase(request("success"))
    end if
 end if

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Administration Console {Action Line - Code Sections}</title>

 	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	 <link rel="stylesheet" type="text/css" href="../global.css" />

<style>
  #newCodeSection td,
  .editCodeSection td,
  .tableadmin th {
     white-space: nowrap;
  }

  #newCode,
  #newDescription,
  .editCode,
  .editDescription {
     width: 500px;
  }

  #editButtonDiv {
     padding-bottom: 5px;
  }

  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
  }
</style>

  <script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.1.min.js"></script>

<script type="text/javascript">
<!--
$(document).ready(function(){
   $('#addButton').click(function() {
      if($('#newCode').val() == '') {
         $('#addButton').focus();
         inlineMsg(document.getElementById('newCode').id,'<strong>Required Field Missing: </strong>Code',10,'newCode');
         return false;
      } else {
         $('#p_action').val('ADD');
         $('#maintainCodeSections').submit();
      }
   });

   $('input[name*="saveChangesButton"]').click(function() {
      var lcl_total_codesections = Number(0);
      var lcl_false_count        = Number(0);

      if($('#total_codesections').val() != '') {
         lcl_total_codesections = Number($('#total_codesections').val());
      }

      for (var i = parseInt(lcl_total_codesections); i >= 1 ; i--) {
         if($('#editCode_' + i).val() == '') {
            $('#editCode_' + i).focus();
            inlineMsg(document.getElementById('editCode_' + i).id,'<strong>Required Field Missing: </strong> Code.',10,'editCode_' + i);
            lcl_false_count = lcl_false_count + 1;
         }
      }

      if(lcl_false_count > 0) {
         return false;
      } else {
         $('#p_action').val('EDIT');
         $('#maintainCodeSections').submit();
      }
   });
});

function confirm_delete(iLineCount) {
  var lcl_codename;

  if($('#editCode_' + iLineCount).val() != '') {
     lcl_codename = $('#editCode_' + iLineCount).val();
  }

  input_box = confirm("Are you sure you want to delete '" + lcl_codename + "'?");

  if(input_box == true) { 
     $('#p_action').val('DELETE');
     $('#deleteActionCodeID').val($('#editActionCodeID_' + iLineCount).val());
     $('#maintainCodeSections').submit();
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
//-->
</script>
</head>

<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="document.new_substatus.p_new_substatus.focus()"> -->
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<form name=""maintainCodeSections"" id=""maintainCodeSections"" method=""post"" action=""action_code_sections_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""p_action"" id=""p_action"" value="""" size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""deleteActionCodeID"" id=""deleteActionCodeID"" value="""" size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Action Line: Code Sections</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf

 'BEGIN: Add Code Section -----------------------------------------------------
  response.write "          <p>" & vbcrlf
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "            <table id=""newCodeSection"" width=""100%"" border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <th align=""left"" colspan=""5"">Create a Code Section</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td valign=""top"">Code:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""newCode"" id=""newCode"" maxlength=""150"" onchange=""clearMsg('newCode');"" /></td>" & vbcrlf
  response.write "                  <td width=""70%"">&nbsp;</td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td valign=""top"">Description:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <textarea name=""newDescription"" id=""newDescription"" maxlength=""" & lcl_desc_char_length & """></textarea>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td width=""70%"" valign=""bottom"">" & vbcrlf
  response.write "                      <input type=""button"" name=""addButton"" id=""addButton"" value=""Add Code Section"" class=""button"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          </p>" & vbcrlf
 'END: Add Code Section -------------------------------------------------------

 'BEGIN: Edit Code Section ----------------------------------------------------
  response.write "          <p>" & vbcrlf
  response.write "          <div id=""editButtonDiv"">" & vbcrlf
  response.write "            <input type=""button"" name=""saveChangesButton"" id=""saveChangesButton"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "            <table border=""0"" cellpadding=""5"" cellspacing=""0"" class=""tableadmin"">" & vbcrlf
  response.write "          		  <tr>" & vbcrlf
  response.write "          		      <th align=""left"">Code Section</th>" & vbcrlf
  response.write "          			     <th>Active</th>" & vbcrlf
  response.write "                  <th>&nbsp;</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  lcl_bgcolor    = "#eeeeee"
  lcl_line_count = 0

  sSQL = "SELECT action_code_id, "
  sSQL = sSQL & " code_name, "
  sSQL = sSQL & " description, "
  sSQL = sSQL & " active_flag "
		sSQL = sSQL & " FROM egov_actionline_code_sections "
		sSQL = sSQL & " WHERE orgid = " & session("orgid")
		sSQL = sSQL & " ORDER BY UPPER(code_name), action_code_id "

		set oCodeSections = Server.CreateObject("ADODB.Recordset")
  oCodeSections.Open sSQL, Application("DSN"), 3, 1

		if not oCodeSections.eof then
		   do while not oCodeSections.eof
        lcl_line_count              = lcl_line_count + 1
        lcl_bgcolor                 = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_CSExistsOnRequest       = false
        lcl_selected_activeflag_yes = ""
        lcl_selected_activeflag_no  = " selected=""selected"""

        if oCodeSections("active_flag") = "Y" then
           lcl_selected_activeflag_yes = " selected=""selected"""
           lcl_selected_activeflag_no  = ""
        end if

        lcl_CSExistsOnRequest = checkCodeSectionUsedOnRequest(oCodeSections("action_code_id"))

        response.write "              <tr bgcolor=""" & lcl_bgcolor & """ align=""center"" valign=""top"">" & vbcrlf
        response.write "                  <td>" & vbcrlf
        response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""editCodeSection"">" & vbcrlf
        response.write "                        <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "                            <td align=""left"" valign=""top"">Code:</td>" & vbcrlf
        response.write "                            <td>" & vbcrlf
        response.write "                                <input type=""hidden"" name=""editActionCodeID_" & lcl_line_count & """ id=""editActionCodeID_" & lcl_line_count & """ value=""" & oCodeSections("action_code_id") & """ size=""5"" maxlength=""10"" />" & vbcrlf
        response.write "                                <input type=""text"" name=""editCode_" & lcl_line_count & """ id=""editCode_" & lcl_line_count & """ value=""" & oCodeSections("code_name") & """ class=""editCode"" maxlength=""150"" onchange=""clearMsg('editCode_" & lcl_line_count & "');"" />" & vbcrlf
        response.write "                            </td>" & vbcrlf
        response.write "                        </tr>" & vbcrlf
        response.write "                        <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "                            <td align=""left"" valign=""top"">Description:</td>" & vbcrlf
        response.write "                            <td>" & vbcrlf
        response.write "                                <textarea name=""editDescription_" & lcl_line_count & """ id=""editDescription_" & lcl_line_count & """ class=""editDescription"" maxlength=""" & lcl_desc_char_length & """>" & oCodeSections("description") & "</textarea>" & vbcrlf
        response.write "                            </td>" & vbcrlf
        response.write "                        </tr>" & vbcrlf
        response.write "                      </table>" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "                  <td>" & vbcrlf
        response.write "                      <select name=""editActive_" & lcl_line_count & """ id=""editActive_" & lcl_line_count & """>" & vbcrlf
        response.write "                        <option value=""Y""" & lcl_selected_activeflag_yes & ">Yes</option>" & vbcrlf
        response.write "                        <option value=""N""" & lcl_selected_activeflag_no  & ">No</option>" & vbcrlf
        response.write "                      </select>" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "                  <td>" & vbcrlf

        if not lcl_CSExistsOnRequest then
           response.write "                      <input type=""button"" name=""deleteCodeButton_" & oCodeSections("action_code_id") & """ id=""deleteCodeButton_" & oCodeSections("action_code_id") & """ value=""Delete"" class=""button"" onclick=""clearMsg('editCode_" & lcl_line_count & "');confirm_delete('" & lcl_line_count & "');"" />" & vbcrlf
        else
           response.write "&nbsp;" & vbcrlf
        end if

        response.write "                  </td>" & vbcrlf
        response.write "              </tr>" & vbcrlf

						  oCodeSections.movenext
					loop
		end if

  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "          <input type=""hidden"" name=""total_codesections"" id=""total_codesections"" value=""" & lcl_line_count & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "          <div id=""editButtonDiv"">" & vbcrlf
  response.write "            <input type=""button"" name=""saveChangesButton2"" id=""saveChangesButton2"" value=""Save Changes"" class=""button"" />" & vbcrlf
  response.write "          </div>" & vbcrlf
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
function checkCodeSectionUsedOnRequest(iActionCodeID)
  dim sActionCodeID, sSQLc, lcl_return

  sActionCodeID = ""
  lcl_return    = false

  if iActionCodeID <> "" then
     sActionCodeID = clng(iActionCodeID)
  end if

  sSQLc = "SELECT count(submitted_request_id) as total_requests "
  sSQLc = sSQLc & " FROM egov_submitted_request_code_sections "
  sSQLc = sSQLc & " WHERE submitted_action_code_id = " & sActionCodeID

		set oCheckForCS = Server.CreateObject("ADODB.Recordset")
  oCheckForCS.Open sSQLc, Application("DSN"), 3, 1

  if not oCheckForCS.eof then
     if oCheckForCS("total_requests") > 0 then
        lcl_return = true
     end if
  end if

  set oCheckForCS = nothing

  checkCodeSectionUsedOnRequest = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  dim lcl_return, lcl_orgid, lcl_dmid

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "Does not exist..."
     elseif iSuccess = "ERROR" then
        lcl_return = "ERROR"
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>
