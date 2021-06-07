<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: staff_directory_info.asp
' AUTHOR:   David Boyer
' CREATED:  01/04/2008
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays detail information for Organizational Groups and Staff Members on Staff Directory
'
' MODIFICATION HISTORY
' 1.0  01/04/08	 David Boyer - Created
' 1.1  01/22/08  David Boyer - Added "isFeatureOffline" check
' 1.2	01/29/10	Steve Loar - Changed the mailto link and email displayed to be hidden from spam bots
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'To help prevent hacks.
 if NOT isnumeric(request("org_group_id")) then
    response.redirect "staff_directory.asp"
 end if

'Check to see if the feature is offline
if isFeatureOffline("staff_directory") = "Y" then
   response.redirect "outage_feature_offline.asp"
end if

lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'Retrieve the org_group_id of the organization group that is to be maintained.
'If no value exists then redirect them back to the main results screen
 if request("org_group_id") <> "" then
    lcl_org_group_id = 0
    on error resume next
    lcl_org_group_id = CLng(request("org_group_id"))
    on error goto 0
    if lcl_org_group_id = 0 then response.redirect("staff_directory.asp")
 else
    response.redirect("staff_directory.asp")
 end if

'Build the Return Button URL
 lcl_returnbutton_url = "staff_directory.asp"
 lcl_returnbutton_url = lcl_returnbutton_url & "?sc_org_group_id=" & session("sc_org_group_id")
 lcl_returnbutton_url = lcl_returnbutton_url & "&sc_first_name="   & session("sc_first_name")
 lcl_returnbutton_url = lcl_returnbutton_url & "&sc_last_name="    & session("sc_last_name")
 lcl_returnbutton_url = lcl_returnbutton_url & "&sc_show_members=" & session("sc_show_members")
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

 	<script language="javascript" src="scripts/modules.js"></script>
 	<script language="javascript" src="scripts/easyform.js"></script>
  <script language="javascript" src="scripts/ajaxLib.js"></script>
  <script language="javascript" src="scripts/removespaces.js"></script>
  <script language="javascript" src="scripts/setfocus.js"></script>

  <script type="text/javascript" src="https://code.jquery.com/jquery-1.5.2.min.js"></script>

<script>
  $(document).ready(function() {
     $('#returnButton').click(function(event) {
       location.href='<%=lcl_returnbutton_url%>';
     });
  });
</script>
</head>
<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"> -->
<!--#include file="include_top.asp"-->
<%
  response.write "  <form name=""org_group_maint"" method=""post"" action=""staff_directory_info.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""org_group_id"" id=""org_group_id"" value=""" & lcl_org_group_id & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & lcl_orgid & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""org_level"" id=""org_level"" value=""" & lcl_org_level & """ size=""4"" maxlength=""10"" />" & vbcrlf

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" style=""max-width:800px;"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  'response.write "      <td><a href=""staff_directory.asp?sc_org_group_id=" & session("sc_org_group_id") & "&sc_first_name=" & session("sc_first_name") & "&sc_last_name=" & session("sc_last_name") & "&sc_show_members=" & session("sc_show_members") & """>Return to Staff Directory</a></td>" & vbcrlf
  response.write "      <td><input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to Staff Directory"" class=""button"" /></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td>" & vbcrlf

 'Display all details about Organizational Group selected
  lcl_list_type = "PARENT"
  display_org_details lcl_org_group_id, lcl_list_type

 'Display all users assigned to this org_group_id if any exist
  display_org_group_users iorgid, lcl_org_group_id

 'Now check to see if this organizational group has any sub-level groups assigned to it.
 'We are only displaying the first level under this group and not every level.
 'Also display any staff members of any of these sub-level orgs.
  if check_for_sub_org_groups(lcl_org_group_id) = "Y" then
     sSQLs = " SELECT org_group_id, org_name "
     sSQLs = sSQLs & " FROM egov_staff_directory_groups "
     sSQLs = sSQLs & " WHERE parent_org_group_id = " & CLng(lcl_org_group_id)
     sSQLs = sSQLs & " AND active_flag = 'Y' "
     sSQLs = sSQLs & " ORDER BY UPPER(org_name) "

     set rss = Server.CreateObject("ADODB.Recordset")
     rss.Open sSQLs, Application("DSN"), 3, 1

     if not rss.eof then
        while not rss.eof
          'Display all details about Organizational Group selected
           lcl_list_type = "CHILD"
           display_org_details rss("org_group_id"), lcl_list_type

          'Display all users assigned to this org_group_id if any exist
           display_org_group_users iorgid, rss("org_group_id")

           rss.movenext
        wend
     end if
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!-- #include file="include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub display_organizational_groups_dropdown(p_orgid, p_org_group_id, p_current_parent_org_group_id, p_parent_org_group_id, _
                                           p_org_level, p_limit_list, p_first_run)
'Retrieve all of the organizational groups
 sSQLg = "SELECT org_group_id, "
 sSQLg = sSQLg & " org_name, "
 sSQLg = sSQLg & " org_level "
 sSQLg = sSQLg & " FROM egov_staff_directory_groups "
 sSQLg = sSQLg & " WHERE orgid=" & p_orgid

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

 sSQLg = sSQLg & " AND org_group_id <> " & p_org_group_id
 sSQLg = sSQLg & " AND active_flag = 'Y' "
 sSQLg = sSQLg & " ORDER BY UPPER(org_name) "

 set rsg = Server.CreateObject("ADODB.Recordset")
 rsg.Open sSQLg, Application("DSN"), 3, 1

 if not rsg.eof then
    while not rsg.eof
      'Determine how far to indent the org_name
       lcl_indent = ((rsg("org_level")-1) * 5)

       lcl_indent_spaces = ""
       for x = 1 to lcl_indent
           lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
       next

       if rsg("org_group_id") = clng(p_current_parent_org_group_id) then
          lcl_selected = " selected"
       else
          lcl_selected = ""
       end if

       response.write "<option value=""" & rsg("org_group_id") & """" & lcl_selected & ">" & lcl_indent_spaces & rsg("org_name") & "</option>" & vbcrlf

      'Retrieve sub-organizational groups
       display_organizational_groups_dropdown p_orgid, p_org_group_id, p_current_parent_org_group_id, _
                                              rsg("org_group_id"), lcl_org_level, p_limit_list, lcl_first_run

       rsg.movenext
    wend
 end if
end sub

'-----------------------------------------------------------------------------
sub display_org_details(p_org_group_id,p_list_type)
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
 sSQL = sSQL & " WHERE org_group_id = " & CLng(p_org_group_id)

 set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open sSQL, Application("DSN"), 3, 1

 if not rs.eof then
    lcl_org_name            = rs("org_name")

    if rs("parent_org_group_id") = 0 then
       lcl_parent_org_group_id = ""
    else
       lcl_parent_org_group_id = rs("parent_org_group_id")
    end if

    lcl_org_level           = rs("org_level")
    lcl_orgid               = rs("orgid")
    lcl_address             = rs("address")
    lcl_address2            = rs("address2")
    lcl_city                = rs("city")
    lcl_state               = rs("state")
    lcl_zip                 = rs("zip")
    lcl_phone_number        = rs("phone_number")
    lcl_phone_number_ext    = rs("phone_number_ext")
    lcl_fax_number          = rs("fax_number")
    lcl_email               = rs("email")
    lcl_active_flag         = rs("active_flag")
 else
    response.redirect("staff_directory.asp")
 end if

 response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
 response.write "  <tr class=""staff_org_name_title"">" & vbcrlf
 response.write "      <th align=""left"">"

 if check_for_sub_org_groups(p_org_group_id) = "Y" AND p_list_type = "CHILD" then
    response.write lcl_org_name & "&nbsp;&nbsp;&nbsp;"
    response.write "<a href=""staff_directory_info.asp?org_group_id=" & p_org_group_id & """>[View Sub-Organization Details]</a>"
 else
    response.write lcl_org_name
 end if

 response.write "      </th>" & vbcrlf
 response.write "  </tr>" & vbcrlf

 if lcl_address <> "" then
    response.write "  <tr><td>" & lcl_address & "</td></tr>" & vbcrlf
 end if

 if lcl_address2 <> "" then
    response.write "  <tr><td>" & lcl_address2 & "</td></tr>" & vbcrlf
 end if

'Build the City/State/Zip
 if lcl_city <> "" then
    lcl_city_state_zip = lcl_city
 end if

 if lcl_state <> "" then
    if lcl_city_state_zip <> "" then
       lcl_city_state_zip = lcl_city_state_zip & ", " & lcl_state
    else
       lcl_city_state_zip = lcl_state
    end if
 end if

 if lcl_zip <> "" then
    if lcl_city_state_zip <> "" then
       lcl_city_state_zip = lcl_city_state_zip & " " & lcl_zip
    else
       lcl_city_state_zip = lcl_zip
    end if
 end if

 if lcl_city_state_zip <> "" then
    response.write "  <tr><td>" & lcl_city_state_zip & "</td></tr>" & vbcrlf
 end if

 if lcl_phone_number <> "" then
    response.write "  <tr><td>Phone: " & lcl_phone_number & "</td></tr>" & vbcrlf
 end if

 if lcl_fax_number <> "" then
    response.write "  <tr><td>Fax: " & lcl_fax_number & "</td></tr>" & vbcrlf
 end if

 if lcl_email <> "" then
    'response.write "  <tr><td><a href=""mailto:" & lcl_email & """>" & lcl_email & "</a></td></tr>" & vbcrlf
	response.write "  <tr><td><a " & FormatMailToAsJavascript(lcl_email) & ">" &  FormatEmailAsDecimal(lcl_email)  & "</a></td></tr>" & vbcrlf
 end if

 response.write "</table>" & vbcrlf
end sub

'-----------------------------------------------------------------------------
sub display_org_group_users(p_orgid, p_org_group_id)

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
    sSQLo = sSQLo & " FROM users u, egov_staff_directory_usergroups g "
    sSQLo = sSQLo & " WHERE u.userid = g.userid "
    sSQLo = sSQLo & " AND g.org_group_id = " & CLng(p_org_group_id)
    sSQLo = sSQLo & " AND u.orgid = " & p_orgid
    sSQLo = sSQLo & " AND u.staff_dir_display = 'Y' "
    sSQLo = sSQLo & " ORDER BY UPPER(u.lastname), UPPER(u.firstname) "

    set rso = Server.CreateObject("ADODB.Recordset")
    rso.Open sSQLo, Application("DSN"), 3, 1

    if not rso.eof then
       lcl_bgcolor             = "#ffffff"
       lcl_show_column_headers = "Y"

       while not rso.eof
          if lcl_show_column_headers = "Y" then
             response.write "        <div class=""shadow"">" & vbcrlf
             response.write "        <table class=""staff_directory_table"">" & vbcrlf
             response.write "          <tr id=""columnHeaders"">" & vbcrlf
             response.write "              <th>&nbsp;Staff Members</th>" & vbcrlf
             response.write "              <th>Job Title</th>" & vbcrlf
             response.write "              <th>E-mail</th>" & vbcrlf
             response.write "              <th>Phone</th>" & vbcrlf
             response.write "          </tr>" & vbcrlf

             lcl_show_column_headers = "N"
          else
             lcl_show_column_headers = "N"
          end if

          if rso("jobtitle") <> "" then
             lcl_job_title = rso("jobtitle")
          else
             lcl_job_title = ""
          end if

          if rso("businessnumber") <> "" then
             lcl_phone = trim(rso("businessnumber"))
          else
             lcl_phone = "&nbsp;"
          end if

          if rso("email") <> "" then
             lcl_email = rso("email")
          else
             lcl_email = ""
          end if

          lcl_bgcolor            = changeBGColor(lcl_bgcolor, "#ffffff", "#eeeeee")
          lcl_display_email_href = ""
          lcl_display_email_text = ""
          lcl_display_email      = "&nbsp;"

          if lcl_email <> "" then
             lcl_display_email_href = FormatMailToAsJavascript(lcl_email)  'function in include_top_functions.asp
             lcl_display_email_text = FormatEmailAsDecimal(lcl_email)      'function in include_top_functions.asp
             lcl_display_email      = "<a " & lcl_display_email_href & ">" & lcl_display_email_text & "</a>" & vbcrlf
      		  end if

          response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
          response.write "      <td>" & rso("lastname") & ", " & rso("firstname") & "</td>"
          response.write "      <td>" & lcl_job_title & "</td>" & vbcrlf
       			response.write "      <td>" & lcl_display_email & "</td>" & vbcrlf
          response.write "      <td align=""right"">" & lcl_phone & "</td>" & vbcrlf
          response.write "  </tr>" & vbcrlf

          rso.movenext
       wend
       response.write "</table>" & vbcrlf
       response.write "</div>" & vbcrlf
    end if
    response.write "<p>" & vbcrlf
end sub

'-----------------------------------------------------------------------------
function check_for_sub_org_groups(p_parent_org_group_id)

  if p_parent_org_group_id <> 0 then
     sSQL = "SELECT distinct 'Y' AS lcl_exists " 
     sSQL = sSQL & " FROM egov_staff_directory_groups "
     sSQL = sSQL & " WHERE parent_org_group_id = " & CLng(p_parent_org_group_id)
     sSQL = sSQL & " AND active_flag = 'Y' "

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSQL, Application("DSN"), 3, 1

     if not rs.eof then
        lcl_exists = rs("lcl_exists")
     else
        lcl_exists = "N"
     end if
  else
     lcl_exists = "N"
  end if

  check_for_sub_org_groups = lcl_exists

end function

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
