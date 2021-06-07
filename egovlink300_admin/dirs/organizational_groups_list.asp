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

if not userhaspermission(session("userid"), "staff_directory") then
  	response.redirect sLevel & "permissiondenied.asp"
end if

'Retrieve the search parameters
 lcl_sc_org_name     = request("sc_org_name")
 lcl_sc_show_members = "N"

 if request("sc_show_members") <> "" then
    lcl_sc_show_members = request("sc_show_members")
 end if

'Set up search criteria session variables
 session("sc_org_name")     = lcl_sc_org_name
 session("sc_show_members") = lcl_sc_show_members

'Determine if there is a screen message to display
 lcl_onload  = ""
 lcl_message = ""

 if request("success") <> "" then
    lcl_message = setupScreenMsg(request("success"))
    lcl_onload  = lcl_onload & "displayScreenMsg('" & lcl_message & "');"
 end if
%>
<html>
<head>
  <title>E-Gov Link Administration { Staff Directory }</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<style>
  #screenMsg {
     text-align:  right;
     color:       #ff0000;
     font-weight: bold;
  }

  #noRecordsFound {
     color:       #800000;
     font-weight: bold;
  }

  .redText {
     color: #ff0000;
  }
</style>

  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

<script type="text/javascript">
$(document).ready(function() {
  $('#addOrgGroupButton').click(function() {
     var lcl_url  = 'organizational_groups_maint.asp';
         lcl_url += '?screen_mode=ADD';

     location.href = lcl_url;
  });
});

function deleteOrg(iOrgGroupID) {
  var lcl_org_groupid = '';
  var lcl_org_name    = '';
  var lcl_delete_url  = '';

  //lcl_org_groupid = $('#orgGroupID' + iRowID).val();
  lcl_org_groupid = iOrgGroupID;
  lcl_org_name    = $('#orgName' + lcl_org_groupid).html();

  if(lcl_org_name != '') {
     lcl_org_name = lcl_org_name.replace("'","\'");
  }

  var r = confirm('Are you sure you want to delete the organizational group: "' + lcl_org_name + '"? \n NOTE: Any/All sub-organizational groups will also be deleted.');

  if(r == true) {
     lcl_delete_url  = 'organizational_groups_action.asp';
     lcl_delete_url += '?user_action=D';
     lcl_delete_url += '&orgid=<%=session("orgid")%>';
     lcl_delete_url += '&org_group_id=' + lcl_org_groupid;

     location.href = lcl_delete_url;
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
 'Determine which value is selected in the dropdown list
  lcl_selected_yes = ""
  lcl_selected_no  = " selected=""selected"""

  if lcl_sc_show_members = "Y" then
     lcl_selected_yes = " selected=""selected"""
     lcl_selected_no  = ""
  end if

  response.write "<form name=""staffDirectory"" id=""staffDirectory"" value=""organizational_groups_list.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""groupid"" id=""groupid"" value=""" & lcl_group_id & """ size=""10"" maxlength=""10"" />" & vbcrlf
  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""900"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>Manage Organizational Groups (Staff Directory)</strong></font></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "            <legend><strong>Search Options&nbsp;</strong></legend>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
  response.write "              <tr valign=""top"">" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td>Organizational Group:</td>" & vbcrlf
  response.write "                            <td><input type=""text"" name=""sc_org_name"" id=""sc_org_name"" value=""" & lcl_sc_org_name & """ size=""50"" maxlength=""500"" /></td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td colspan=""2"">" & vbcrlf
  response.write "                                Show Staff Members in results:&nbsp;" & vbcrlf
  response.write "                                <select name=""sc_show_members"" id=""sc_show_members"">" & vbcrlf
  response.write "                                  <option value=""Y""" & lcl_selected_yes & ">Yes</option>" & vbcrlf
  response.write "                                  <option value=""N""" & lcl_selected_no  & ">No</option>" & vbcrlf
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
  response.write "                      </table>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  response.write "              <tr><td colspan=""2""><input type=""submit"" name=""searchButton"" id=""searchButton"" class=""button"" value=""SEARCH"" /></td></tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "        <td valign=""top"">" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""button"" name=""addOrgGroupButton"" id=""addOrgGroupButton"" value=""Add Organizational Group"" class=""button"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td id=""screenMsg""></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "            <div class=""shadow"">" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" class=""tableadmin"">" & vbcrlf
  response.write "              <tr align=""left"">" & vbcrlf
  response.write "                  <th colspan=""2"" width=""65%"">&nbsp;Organizational Groups</th>" & vbcrlf
  response.write "                  <th>Phone</th>" & vbcrlf
  response.write "                  <th>E-mail</th>" & vbcrlf
  response.write "                  <th>Active</th>" & vbcrlf
  response.write "                  <th>&nbsp;</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf

                                session("bgcolor")     = "#ffffff"
                                lcl_org_group_id       = ""
                                lcl_totalrecords_found = 0

                                display_organizational_groups session("orgid"), _
                                                              lcl_org_group_id, _
                                                              lcl_sc_org_name, _
                                                              lcl_sc_show_members, _
                                                              lcl_totalrecords_found

  response.write "            </table>" & vbcrlf
  response.write "            </div>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "  </form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "  </body>" & vbcrlf
  response.write "  </html>" & vbcrlf

'------------------------------------------------------------------------------
'sub display_organizational_groups(p_org_group_id, p_sc_org_name, p_sc_phone_number, p_sc_phone_number_ext, p_sc_email, p_sc_show_members, p_total_records_found)
sub display_organizational_groups(iOrgID, p_org_group_id, p_sc_org_name, p_sc_show_members, p_total_records_found)

 lcl_search_criteria_entered = "Y"

 if p_total_records_found = "" then
    lcl_total_records_found = 0
 end if

 if p_total_records_found > 0 then
    lcl_total_records_found = clng(p_total_records_found)
 end if

'Determine if search criteria has been entered (not including the "show staff members" option
 if p_sc_org_name = "" then
    lcl_search_criteria_entered = "N"
 end if

'Retrieve all of the sub organizational groups
 sSQL = "SELECT org_group_id, "
 sSQL = sSQL & " org_name, "
 sSQL = sSQL & " org_level, "
 sSQL = sSQL & " phone_number, "
 sSQL = sSQL & " phone_number_ext, "
 sSQL = sSQL & " email, "
 sSQL = sSQL & " active_flag "
 sSQL = sSQL & " FROM egov_staff_directory_groups "
 sSQL = sSQL & " WHERE orgid = " & iOrgID

 if p_org_group_id <> "" then
    sSQL = sSQL & " AND parent_org_group_id = " & p_org_group_id
 else
   'If search criteria has been entered then it does not matter what "parent"
   'the org belongs to we are searching on the entire list.
    if  lcl_search_criteria_entered = "N" AND p_org_group_id = "" then
       sSQL = sSQL & " AND (parent_org_group_id IS NULL OR parent_org_group_id = 0) "
    end if
 end if

 'sSQL = sSQL & " AND active_flag = 'Y' "

 if p_sc_org_name <> "" then
    sSQL = sSQL & " AND UPPER(org_name) like ('%" & UCASE(trim(p_sc_org_name)) & "%') "
 end if

 if lcl_search_criteria_entered = "N" then
    sSQL = sSQL & " ORDER BY org_level, UPPER(org_name) "
 else
    sSQL = sSQL & " ORDER BY UPPER(org_name) "
 end if

 set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open sSQL, Application("DSN"), 3, 1

 if not rs.eof then
    do while not rs.eof
       lcl_total_records_found = lcl_total_records_found + 1
       lcl_bgcolor             = changeBGColor(session("bgcolor"), "#eeeeee","#ffffff")
       session("bgcolor")      = lcl_bgcolor
       lcl_phone               = ""
       lcl_indent_spaces       = ""
       lcl_org_path            = ""
       lcl_email               = "&nbsp;"
       lcl_active_flag         = "&nbsp;"

      'Build the phone number
       if rs("phone_number") <> "" then
          lcl_phone = rs("phone_number")
       end if

       if rs("phone_number_ext") <> "" then
          if lcl_phone <> "" then
             lcl_phone = lcl_phone & " x" & rs("phone_number_ext")
          else
             lcl_phone = "x" & rs("phone_number_ext")
          end if
       end if

       if lcl_phone = "" then
          lcl_phone = "&nbsp;"
       end if

       if rs("active_flag") = "Y" then
          lcl_active_flag = rs("active_flag")
       else
          lcl_active_flag = "<span class=""redText"">[N]</span>"
       end if

      'If any of the search criteria fields have been populated (besides the "show staff members" field)
      'then do NOT indent the results
       if lcl_search_criteria_entered = "N" then

         'Determine how far to indent the org_name
          lcl_indent = ((rs("org_level")-1) * 5)

          for x = 1 to lcl_indent
              lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
          next
       end if

      'If any of the search criteria fields have been populated (besides the "show staff members" field)
      'AND the org found has an ORG_LEVEL greater than 1 then show its root path in the results
       if lcl_search_criteria_entered = "Y" then
          if rs("org_level") > 1 then
             lcl_org_path = display_parent_group_path(rs("org_group_id"), "")
          else
             lcl_org_path = ""
          end if

          if lcl_org_path <> "" then
             lcl_org_path = "&nbsp;&nbsp;&nbsp;<strong><em>Belongs to: </em></strong>" & lcl_org_path
          end if
       end if

       if rs("email") <> "" then
          lcl_email = "<a href=""mailto:" & rs("email") & """>" & rs("email") & "</a>"
       end if

       lcl_onMouseOver = " onMouseOver=""tooltip.show('Click to edit');"""
       lcl_onMouseOut  = " onMouseOut=""tooltip.hide();"""

       response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
       response.write "    <td colspan=""2"">" & vbcrlf
       response.write          lcl_indent_spaces & "<a href=""organizational_groups_maint.asp?org_group_id=" & rs("org_group_id") & """ id=""orgName" & rs("org_group_id") & """" & lcl_onMouseOver & lcl_onMouseOut & ">" & rs("org_name") & "</a>" & vbcrlf
       response.write          lcl_org_path & vbcrlf
       response.write "    </td>" & vbcrlf
       response.write "    <td>"                  & lcl_phone       & "</td>" & vbcrlf
       response.write "    <td>"                  & lcl_email       & "</td>" & vbcrlf
       response.write "    <td align=""center"">" & lcl_active_flag & "</td>" & vbcrlf
       response.write "    <td align=""center"">" & vbcrlf
       response.write "        <input type=""button"" name=""deleteButton" & lcl_total_records_found & """ id=""deleteButton" & lcl_total_records_found & """ value=""Delete"" class=""button"" onclick=""deleteOrg('" & rs("org_group_id") & "');"" />" & vbcrlf
       'response.write "        <input type=""text"" name=""orgGroupID"     & rs("org_group_id")      & """ id=""orgGroupID"   & rs("org_group_id")      & """ value=""" & rs("org_group_id") & """ size=""3"" maxlength=""10"" />" & vbcrlf
       response.write "    </td>" & vbcrlf
       response.write "</tr>" & vbcrlf

      'If the user wishes to display the members in each organizational group then
      'check for any users assigned to current organizational groups
       if p_sc_show_members = "Y" then
          display_org_group_users session("orgid"), _
                                  rs("org_group_id"), _
                                  rs("org_level"), _
                                  lcl_search_criteria_entered, _
                                  lcl_bgcolor
       end if

      'Retrieve sub-organizational groups
       if lcl_search_criteria_entered = "N" then
          display_organizational_groups iOrgID, _
                                        rs("org_group_id"), _
                                        p_sc_org_name, _
                                        p_sc_show_members, _
                                        lcl_total_records_found
       end if

       rs.movenext
    loop
 else
    if lcl_total_records_found = 0 then
       response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
       response.write "    <td colspan=""4"" id=""noRecordsFound"">No Organizational Groups found</td>" & vbcrlf
       response.write "</tr>" & vbcrlf
    end if
 end if
end sub

'-----------------------------------------------------------------------------
function display_parent_group_path(p_org_group_id, p_parent_path)

if p_parent_path <> "" then
   lcl_parent_path = p_parent_path
else
   lcl_parent_path = ""
end if

'Retrieve the parent group organization name
 sSQL = "SELECT g.org_group_id, "
 sSQL = sSQL & " g.org_name "
 sSQL = sSQL & " FROM egov_staff_directory_groups g "
 sSQL = sSQL & " WHERE g.org_group_id = (select g2.parent_org_group_id "
 sSQL = sSQL &                         " from egov_staff_directory_groups g2 "
 sSQL = sSQL &                         " where g2.org_group_id = " & p_org_group_id
 sSQL = sSQL &                         " and g2.orgid = " & session("orgid")
 sSQL = sSQL &                         " and g2.active_flag = 'Y') "
 sSQL = sSQL & " AND g.orgid = " & session("orgid")
 sSQL = sSQL & " AND g.active_flag = 'Y' "

 Set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open sSQL, Application("DSN"), 3, 1

 if not rs.eof then
    if lcl_parent_path = "" then
       lcl_parent_path = rs("org_name")
    else
       lcl_parent_path = rs("org_name") & " --> " & lcl_parent_path
    end if

    lcl_parent_path = display_parent_group_path(rs("org_group_id"),lcl_parent_path)

 end if

 display_parent_group_path = lcl_parent_path

end function

'-----------------------------------------------------------------------------
sub display_org_group_users(iOrgID, p_org_group_id, p_org_level, p_search_criteria_entered, p_bgcolor)

  lcl_org_level     = p_org_level
  lcl_indent        = (lcl_org_level * 5)
  lcl_indent_spaces = ""

 'If search criteria has been entered then the indentation amount is set to always be the same
  if p_search_criteria_entered = "Y" then
     lcl_org_level = 1
  end if

 'Determine how far to indent the org_name
  for x = 1 to lcl_indent
     lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
  next

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
     do while not rso.eof
        lcl_job_title = ""
        lcl_phone     = "&nbsp;"
        lcl_email     = "&nbsp;"

        if rso("jobtitle") <> "" then
           lcl_job_title = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong><i>Job Title: </i></strong>" & rso("jobtitle")
        end if

       'Build the phone number
        if rso("businessnumber") <> "" then
           lcl_phone = rso("businessnumber")
        end if

        if rso("email") <> "" then
           lcl_email = rso("email")
        end if

        response.write "<tr bgcolor=""" & p_bgcolor & """>" & vbcrlf

        if lcl_job_title <> "" then
           response.write "    <td>" & lcl_indent_spaces & "- " & rso("lastname") & ", " & rso("firstname") & "</td>" & vbcrlf
           response.write "    <td>" & lcl_job_title & "</td>" & vbcrlf
        else
           response.write "    <td colspan=""2"">" & lcl_indent_spaces & "- " & rso("lastname") & ", " & rso("firstname") & "</td>" & vbcrlf
        end if

        response.write "    <td>" & lcl_phone & "</td>" & vbcrlf
        response.write "    <td><a href=""mailto:" & lcl_email & """>" & lcl_email & "</a></td>" & vbcrlf
        response.write "    <td colspan=""2"">&nbsp;</td>" & vbcrlf
        response.write "</tr>" & vbcrlf

        rso.movenext
     loop
  end if

  rso.close
  set rso = nothing
end sub

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)
  dim lcl_return, lcl_success

  lcl_return  = ""
  lcl_success = ""

  if iSuccess <> "" then
     if not containsApostrophe(iSuccess) then
        lcl_success = ucase(iSuccess)

        if lcl_success = "SI" then
           lcl_return = "Successfully Created..."
        elseif lcl_success = "D" then
           lcl_return = "Successfully Deleted..."
        elseif lcl_success = "NE" then
           lcl_return = "Organizational Group does not exist..."
        elseif lcl_success = "NO_ADD_ROLE" then
           lcl_return = "You do not have the permission to add Organizational Groups"
        elseif lcl_success = "NO_EDIT_ROLE" then
           lcl_return = "You do not have the permission to maintain Organizational Groups"
        end if
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function setup_bgcolor(p_bgcolor)
  if p_bgcolor = "#efefef" then
     lcl_bgcolor = "#ffffff"
  else
     lcl_bgcolor = "#efefef"
  end if

  setup_bgcolor = lcl_bgcolor

end function
%>