<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="organizational_groups_global_functions.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("staff_directory") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

if request("screen_mode") = "ADD" then
   if NOT UserHasPermission( session("userid"), "create_organizational_groups" ) then
      if userhaspermission(session("userid"),"staff_directory") then
         response.redirect "organizational_groups_list.asp?success=NO_ADD_ROLE"
      else
     	   response.redirect sLevel & "permissiondenied.asp"
      end if
   end if

   lcl_page_title = "Add"

else
   if NOT UserHasPermission( session("userid"), "edit_organizational_groups" ) then
      if userhaspermission(session("userid"),"staff_directory") then
         response.redirect "organizational_groups_list.asp?success=NO_EDIT_ROLE"
      else
     	   response.redirect sLevel & "permissiondenied.asp"
      end if
   end if

   lcl_page_title = "Edit"
end if

'Perform the action the user has selected, if needed
'I=Insert, U=Update
' if request("user_action") <> "" then
'    if request("user_action") = "U" then
'       update_org()
'    elseif request("user_action") = "I" OR request("user_action") = "AA" then
'       insert_org(request("user_action"))
'    elseif request("user_action") = "D" then
'       delete_org request("org_group_id"),request("org_name")
'    end if
' end if

'Determine what mode the screen is in, ADD/EDIT
 if request("screen_mode") <> "" then
    lcl_screen_mode = request("screen_mode")
 else
    lcl_screen_mode = "EDIT"
 end if

'Retrieve the org_group_id of the organization group that is to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("org_group_id") <> "" then
    lcl_org_group_id = request("org_group_id")
 else
    if lcl_screen_mode <> "ADD" then
       response.redirect("organizational_groups_list.asp")
    end if
 end if

'Set up local variables
 lcl_org_name            = ""
 lcl_parent_org_group_id = ""
 lcl_org_level           = ""
 lcl_orgid               = ""
 lcl_address             = ""
 lcl_address2            = ""
 lcl_city                = ""
 lcl_state               = ""
 lcl_zip                 = ""
 lcl_phone_number        = ""
 lcl_phone_number_ext    = ""
 lcl_fax_number          = ""
 lcl_email               = ""
 lcl_active_flag         = ""

 'lcl_parent_org_group_id = 0
 lcl_limit_list          = "N"
 lcl_selected_yes        = " selected"
 lcl_selected_no         = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the organizational group
    sSQL = "SELECT org_name, "
    sSQL = sSQL & " parent_org_group_id, "
    sSQL = sSQL & " org_level, "
    sSQL = sSQL & " orgid, "
    sSQL = sSQL & " address, "
    sSQL = sSQL & " address2, "
    sSQL = sSQL & " city, "
    sSQL = sSQL & " state, "
    sSQL = sSQL & " zip, "
    sSQL = sSQL & " phone_number, "
    sSQL = sSQL & " phone_number_ext, "
    sSQL = sSQL & " fax_number, "
    sSQL = sSQL & " email, "
    sSQL = sSQL & " active_flag "
    sSQL = sSQL & " FROM egov_staff_directory_groups "
    sSQL = sSQL & " WHERE org_group_id = " & lcl_org_group_id

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sSQL, Application("DSN"), 3, 1

    if not rs.eof then
       lcl_org_name            = rs("org_name")

       if rs("parent_org_group_id") = 0 then
          lcl_parent_org_group_id = ""
       else
          lcl_parent_org_group_id = rs("parent_org_group_id")
       end if

       lcl_org_level        = rs("org_level")
       lcl_orgid            = rs("orgid")
       lcl_address          = rs("address")
       lcl_address2         = rs("address2")
       lcl_city             = rs("city")
       lcl_state            = rs("state")
       lcl_zip              = rs("zip")
       lcl_phone_number     = rs("phone_number")
       lcl_phone_number_ext = rs("phone_number_ext")
       lcl_fax_number       = rs("fax_number")
       lcl_email            = rs("email")
       lcl_active_flag      = rs("active_flag")
    else
       response.redirect("organizational_groups_list.asp?success=NE")
    end if
 else
    lcl_org_group_id = "" 
 end if

'First check to see if the current org group has any sub-org groups associated to it.
'If so then show ONLY the org groups that are currently on the same org_level or greater (less than) the 
'current org group because you can not assign a sub-group of an org_group to the org_group.
'If not then show the entire list.
 if lcl_parent_org_group_id <> "" then
    lcl_limit_list = check_for_sub_org_groups(lcl_parent_org_group_id)
 end if

 if lcl_active_flag = "N" then
    lcl_selected_yes = ""
    lcl_selected_no  = " selected"
 end if

'Check for screen message
 lcl_message = ""

 if request("success") <> "" then
    lcl_message = setupScreenMsg(request("success"))
    lcl_onload  = lcl_onload & "displayScreenMsg('" & lcl_message & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Link - Administration { Staff Directory }</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<style>
  #screenMsg {
     text-align:  right;
     color:       #ff0000;
     font-weight: bold;
  }
</style>

  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

<script type="text/javascript">
   var lcl_url_back  = 'organizational_groups_list.asp';
       lcl_url_back += '?sc_org_name=<%=session("sc_org_name")%>';
       lcl_url_back += '&sc_show_members=<%=session("sc_show_members")%>';

   var control_field = "";

$(document).ready(function() {

  $('#backButton').click(function() {
     location.href = lcl_url_back;
  });

  $('#cancelButton').click(function() {
     location.href = lcl_url_back;
  });

  $('#addButton').click(function() {
     validate('I');
  });

  $('#addAnotherButton').click(function() {
     validate('AA');
  });

  $('#saveButton').click(function() {
     clearMsg('org_name');
     validate('U');
  });

  $('#deleteButton').click(function() {
     validate('D');
  });
});

function validate(p_action) {
  var lcl_submit = 'Y';

  if(p_action=="D") {
     var r = confirm('Are you sure you want to delete the organizational group "' + $('#org_name').val() + '"?  \n NOTE: Any/All sub-organizational groups will also be deleted.');

     if (r==false) {
         lcl_submit = 'N';
     }
  }else{
     lcl_org_name = $('#org_name').val()

   		if(lcl_org_name == '') {
        $('#org_name').focus();
        inlineMsg(document.getElementById("org_name").id,'<strong>Required Field: </strong> "Group Name',10,'org_name');
        lcl_submit = 'N';
     }
  }

  $('#user_action').val(p_action);

  if(lcl_submit == 'Y') {
     $('#org_group_maint').submit();
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
<body onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!-- #include file="../menu/menu.asp" -->
<%
  response.write "<form name=""org_group_maint"" id=""org_group_maint"" method=""post"" action=""organizational_groups_action.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""org_group_id"" id=""org_group_id"" value=""" & lcl_org_group_id & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""screen_mode"" id=""screen_mode"" value="""   & lcl_screen_mode  & """ size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""orgid"" id=""orgid"" value="""               & lcl_orgid        & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""org_level"" id=""org_level"" value="""       & lcl_org_level    & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""user_action"" id=""user_action"" value="""" size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <p><font size=""+1""><strong>" & lcl_page_title & " Organizational Group: Staff Directory</strong></font></p>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td><span id=""screenMsg""></span></td>" & vbcrlf

  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""<< Back"" class=""button"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      displayButtonRow lcl_screen_mode
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "          <div class=""shadow"">" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"" class=""tableadmin"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <th colspan=""2"">&nbsp;</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""100"">&nbsp;Group Name:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""org_name"" id=""org_name"" value=""" & lcl_org_name & """ size=""50"" maxlength=""500"" onchange=""clearMsg('org_name');"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;Parent Group:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""parent_org_group_id"" id=""parent_org_group_id"">" & vbcrlf
  response.write "                        <option value=""""></option>" & vbcrlf
                                          display_organizational_groups_dropdown session("orgid"), _
                                                                                 lcl_org_group_id, _
                                                                                 lcl_parent_org_group_id, _
                                                                                 lcl_parent_org_group_id, _
                                                                                 lcl_org_level, _
                                                                                 lcl_limit_list, _
                                                                                 "Y", _
                                                                                 lcl_screen_mode
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;Address:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""address"" id=""address"" value=""" & lcl_address & """ size=""40"" maxlength=""100"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""address2"" id=""address2"" value=""" & lcl_address2 & """ size=""40"" maxlength=""100"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;City:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""city"" id=""city"" value=""" & lcl_city & """ size=""40"" maxlength=""100"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;State:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""state"" id=""state"" value=""" & lcl_state & """ size=""2"" maxlength=""2"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;Zip:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""zip"" id=""zip"" value=""" & lcl_zip & """ size=""5"" maxlength=""10"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;Phone:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""phone_number"" id=""phone_number"" value=""" & lcl_phone_number & """ size=""15"" maxlength=""15"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;Fax:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""fax_number"" id=""fax_number"" value=""" & lcl_fax_number & """ size=""15"" maxlength=""15"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;E-mail:</td>" & vbcrlf
  response.write "                  <td><input type=""text"" name=""email"" id=""email"" value=""" & lcl_email & """ size=""50"" maxlength=""500"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>&nbsp;Active:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""active_flag"" id=""active_flag"">" & vbcrlf
  response.write "                        <option value=""Y""" & lcl_selected_yes & ">Yes</option>" & vbcrlf
  response.write "                        <option value=""N""" & lcl_selected_no  & ">No</option>" & vbcrlf
  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

     if lcl_screen_mode <> "ADD" then
        display_org_group_users session("orgid"), _
                                lcl_org_group_id
     end if

  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!--#include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub display_organizational_groups_dropdown(iOrgID, p_org_group_id, p_current_parent_org_group_id, _
                                           p_parent_org_group_id, p_org_level, p_limit_list, _
                                           p_first_run, p_screen_mode)
'Retrieve all of the organizational groups
 sSQLg = "SELECT org_group_id, "
 sSQLg = sSQLg & " org_name, "
 sSQLg = sSQLg & " org_level "
 sSQLg = sSQLg & " FROM egov_staff_directory_groups "
 sSQLg = sSQLg & " WHERE orgid=" & iOrgID

 if p_first_run = "Y" then
    sSQLg = sSQLg & " AND (parent_org_group_id IS NULL OR parent_org_group_id = 0) "

   'Determine if the list should only display org groups equal to or higher (less than) than the current org_group org_level.
    if p_limit_list = "Y" then
       if p_org_level = "" then
          lcl_org_level = 1
       else
          lcl_org_level = p_org_level
       end if

       sSQLg = sSQLg & " AND org_level <= " & lcl_org_level
    end if

    lcl_first_run = "N"
 else
    sSQLg = sSQLg & " AND parent_org_group_id = " & p_parent_org_group_id
    lcl_first_run = p_first_run
 end if

 if p_screen_mode <> "ADD" then
    sSQLg = sSQLg & " AND org_group_id <> " & p_org_group_id
 end if

 sSQLg = sSQLg & " AND active_flag = 'Y' "
 sSQLg = sSQLg & " ORDER BY UPPER(org_name) "

 set rsg = Server.CreateObject("ADODB.Recordset")
 rsg.Open sSQLg, Application("DSN"), 3, 1

 if not rsg.eof then
    do while not rsg.eof
      'Determine how far to indent the org_name
       lcl_indent        = ((rsg("org_level")-1) * 5)
       lcl_indent_spaces = ""
       lcl_selected      = ""

       for x = 1 to lcl_indent
           lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
       next

       if p_current_parent_org_group_id <> "" then
          if rsg("org_group_id") = clng(p_current_parent_org_group_id) then
             lcl_selected = " selected"
          end if
       end if

       response.write "<option value=""" & rsg("org_group_id") & """" & lcl_selected & ">" & lcl_indent_spaces & rsg("org_name") & "</option>" & vbcrlf

      'Retrieve sub-organizational groups
       display_organizational_groups_dropdown iOrgID, _
                                              p_org_group_id, _
                                              p_current_parent_org_group_id, _
                                              rsg("org_group_id"), _
                                              lcl_org_level, _
                                              p_limit_list, _
                                              lcl_first_run, _
                                              p_screen_mode
       rsg.movenext
    loop
 end if

 set rsg = nothing

end sub

'-----------------------------------------------------------------------------
sub display_org_group_users(iOrgID, p_org_group_id)

   'Retreive all of the users for the orgid and those that are set to display on the Staff Directory page
    sSQLo = "SELECT u.userid, "
    sSQLo = sSQLo & " u.firstname, "
    sSQLo = sSQLo & " u.lastname, "
    sSQLo = sSQLo & " u.jobtitle, "
    sSQLo = sSQLo & " u.department, "
    sSQLo = sSQLo & " u.businessaddress, "
    sSQLo = sSQLo & " u.businessnumber, "
    sSQLo = sSQLo & " u.email, "
    sSQLo = sSQLo & " u.city, "
    sSQLo = sSQLo & " u.state, "
    sSQLo = sSQLo & " u.zipcode, "
    sSQLo = sSQLo & " u.staff_dir_display "
    sSQLo = sSQLo & " FROM users u, "
    sSQLo = sSQLo &      " egov_staff_directory_usergroups g "
    sSQLo = sSQLo & " WHERE u.userid = g.userid "
    sSQLo = sSQLo & " AND g.org_group_id = " & p_org_group_id
    sSQLo = sSQLo & " AND u.orgid = " & iOrgID
    sSQLo = sSQLo & " AND u.staff_dir_display = 'Y' "
    sSQLo = sSQLo & " ORDER BY UPPER(u.lastname), UPPER(u.firstname) "

    set rso = Server.CreateObject("ADODB.Recordset")
    rso.Open sSQLo, Application("DSN"), 3, 1

    if not rso.eof then
       lcl_bgcolor             = "#efefef"
       lcl_show_column_headers = "Y"

       do while not rso.eof
          if lcl_show_column_headers = "Y" then
             response.write "  <tr>" & vbcrlf
             response.write "      <td valign=""top"">" & vbcrlf
             response.write "          <div class=""shadow"">" & vbcrlf
             response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" class=""tableadmin"">" & vbcrlf
             response.write "            <tr>" & vbcrlf
             response.write "                <th align=""left"" colspan=""2"" width=""65%"">&nbsp;Staff Members</th>" & vbcrlf
             response.write "                <th align=""left"">Phone</th>" & vbcrlf
             response.write "            </tr>" & vbcrlf

             lcl_show_column_headers = "N"
          else
             lcl_show_column_headers = "N"
          end if

          lcl_phone     = "&nbsp;"
          lcl_job_title = ""
          lcl_bgcolor   = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

         'Build the phone number
          if rso("businessnumber") <> "" then
             lcl_phone = trim(rso("businessnumber"))
          end if

         'Set up the Job Title
          if rso("jobtitle") <> "" then
             lcl_job_title = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong><em>Job Title: </em></strong>" & rso("jobtitle")
          end if

          response.write "            <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf

          if lcl_job_title <> "" then
             response.write "                <td>" & rso("lastname") & ", " & rso("firstname") & "</td>" & vbcrlf
             response.write "                <td>" & lcl_job_title & "</td>" & vbcrlf
          else
             response.write "                <td colspan=""2"">" & rso("lastname") & ", " & rso("firstname") & "</td>" & vbcrlf
          end if

          response.write "                <td>" & lcl_phone & "</td>" & vbcrlf
          response.write "            </tr>" & vbcrlf

          rso.movenext
       loop

       response.write "          </table>" & vbcrlf
       response.write "          </div>" & vbcrlf
       response.write "      </td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
    end if

    rso.close
    set rso = nothing

end sub

'------------------------------------------------------------------------------
sub displayButtonRow(iScreenMode)

  dim sScreenMode

  if iScreenMode <> "" then
     if not containsApostrophe(iScreenMode) then
        sScreenMode = ucase(iScreenMode)
     end if
  end if

  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" />" & vbcrlf

  if sScreenMode = "ADD" then
     response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" />" & vbcrlf
  else
     response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" />" & vbcrlf
     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" />" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)
  dim lcl_return, lcl_success

  lcl_return  = ""
  lcl_success = ""

  if iSuccess <> "" then
     if not containsApostrophe(iSuccess) then
        lcl_success = ucase(iSuccess)

        if iSuccess = "SU" then
           lcl_return = "Successfully Updated..."
        elseif iSuccess = "SI" then
           lcl_return = "Successfully Created..."
        elseif iSuccess = "NE" then
           lcl_return = "Organizational Group does not exist..."
        end if
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = p_value
  end if

  dbsafe = lcl_value

end function
%>
