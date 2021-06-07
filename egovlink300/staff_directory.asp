<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: staff_directory.asp
' AUTHOR:   David Boyer
' CREATED:  01/04/2008
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the Staff Directory
'
' MODIFICATION HISTORY
' 1.0  01/04/08	 David Boyer - Created
' 1.1  01/22/08  David Boyer - Added "isFeatureOffline" check.
' 1.2	01/29/10	Steve Loar - Changed the mailto link and email displayed to be hidden from spam bots
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("staff_directory") = "Y" then
   response.redirect "outage_feature_offline.asp"
end if

Dim oActionOrg

Set oActionOrg = New classOrganization

lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'Retrieve the search parameters
 'lcl_sc_org_group_id      = request("sc_org_group_id")
 lcl_sc_first_name        = request("sc_first_name")
 lcl_sc_last_name         = request("sc_last_name")
 'lcl_sc_phone_number     = request("sc_phone_number")
 'lcl_sc_phone_number_ext = request("sc_phone_number_ext")
 'lcl_sc_email            = request("sc_email")

 if request("sc_show_members") <> "" then
    lcl_sc_show_members  = request("sc_show_members")
 else
    lcl_sc_show_members  = "N"
 end if

'Set up search criteria session variables
 'session("sc_org_group_id") = lcl_sc_org_group_id
 session("sc_first_name")   = lcl_sc_first_name
 session("sc_last_name")    = lcl_sc_last_name

 'session("sc_phone_number")     = lcl_sc_phone_number
 'session("sc_phone_number_ext") = lcl_sc_phone_number_ext
 'session("sc_email")            = lcl_sc_email
 session("sc_show_members") = lcl_sc_show_members

 'Determine which value is selected in the dropdown list
  if lcl_sc_show_members = "Y" then
     lcl_selected_yes = " selected"
     lcl_selected_no  = ""
  else
     lcl_selected_yes = ""
     lcl_selected_no  = " selected"
  end if
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

</head>
<!--<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">-->
<!--#include file="include_top.asp"-->
<%
  response.write "<font class=""pagetitle"">" & vbcrlf

  if OrgHasDisplay( iOrgId, "staff directory page title" ) then
		   response.write GetOrgDisplay( iOrgId, "staff directory page title" )
	 else
     response.write "Welcome to the " & oActionOrg.GetOrgName() & ", " & oActionOrg.GetState() & ", " & oActionOrg.GetOrgFeatureName( "staff_directory" )
  end if

  response.write "</font>" & vbcrlf
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf

  response.write "<form name=""search_sort_form"" id=""search_sort_form"" value=""staff_directory.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""groupid"" id=""groupid"" value=""" & lcl_group_id & """ size=""10"" maxlength=""10"" />" & vbcrlf

  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" style=""max-width:900px;"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <fieldset>" & vbcrlf
  response.write "            <legend><strong>Search Options&nbsp;</strong></legend><br />" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td>First Name:</td>" & vbcrlf
  response.write "                          <td><input type=""text"" name=""sc_first_name"" id=""sc_first_name"" value=""" & lcl_sc_first_name & """ size=""25"" maxlength=""25"" /></td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td>Last Name:</td>" & vbcrlf
  response.write "                          <td><input type=""text"" name=""sc_last_name"" id=""sc_last_name"" value=""" & lcl_sc_last_name & """ size=""25"" maxlength=""25"" /></td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td colspan=""2"">" & vbcrlf
  response.write "                              Show Staff Members in results:&nbsp;" & vbcrlf
  response.write "                              <select name=""sc_show_members"" id=""sc_show_members"">" & vbcrlf
  response.write "                                <option value=""Y""" & lcl_selected_yes & ">Yes</option>" & vbcrlf
  response.write "                                <option value=""N""" & lcl_selected_no & ">No</option>" & vbcrlf
  response.write "                              </select>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  response.write "            <tr><td colspan=""2""><input type=""submit"" name=""searchButton"" id=""searchButton"" class=""button"" value=""SEARCH"" /></td></tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  
  If OrgHasDisplay( iorgid, "staff directory message" ) Then
    response.write vbcrlf & "<div id=""staffdirectorymessage"">" & GetOrgDisplay( iOrgId, "staff directory message" ) & "</div>"
  End If 


  response.write "<div class=""dirContainer"">" & vbcrlf
  response.write "<table class=""staff_directory_table"">" & vbcrlf
  response.write "  <tr id=""columnHeaders"">" & vbcrlf
  response.write "      <th colspan=""2"" width=""65%"">Organizational Groups</th>" & vbcrlf
  response.write "      <th>Phone</th>" & vbcrlf
  response.write "      <th>E-mail</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  session("bgcolor")      = "#ffffff"
  lcl_org_group_id        = ""
  lcl_sc_org_group_id     = ""
  lcl_total_records_found = 0

  display_organizational_groups lcl_org_group_id, lcl_sc_org_group_id, lcl_sc_first_name, lcl_sc_last_name, _
                                lcl_sc_show_members, lcl_total_records_found

  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf
%>
<!-- #include file="include_bottom.asp" -->
<%
'-----------------------------------------------------------------------------
sub display_organizational_groups(p_org_group_id, p_sc_org_group_id, p_sc_first_name, p_sc_last_name, _
                                  p_sc_show_members, p_total_records_found)

if p_total_records_found > 0 then
   lcl_total_records_found = p_total_records_found
else
   lcl_total_records_found = 0
end if

'Determine if search criteria has been entered (not including the "show staff members" option
 if  p_sc_first_name = "" and p_sc_last_name = "" then
     lcl_search_criteria_entered = "N"
 else
     lcl_search_criteria_entered = "Y"
 end if

'Retrieve all of the sub organizational groups
 sSQL = "SELECT org_group_id, "
 sSQL = sSQL & " org_name, "
 sSQL = sSQL & " org_level, "
 sSQL = sSQL & " phone_number, "
 sSQL = sSQL & " phone_number_ext, "
 sSQL = sSQL & " email "
 sSQL = sSQL & " FROM egov_staff_directory_groups "
 sSQL = sSQL & " WHERE orgid = " & iorgid

 if p_org_group_id <> "" then
    sSQL = sSQL & " AND parent_org_group_id = " & p_org_group_id
 else
   'If search criteria has been entered then it does not matter what "parent" the org belongs to we are searching on the entire list.
    if lcl_search_criteria_entered = "N" and p_org_group_id = "" then
       sSQL = sSQL & " AND (parent_org_group_id IS NULL OR parent_org_group_id = 0) "
    end if
 end if

 sSQL = sSQL & " AND active_flag = 'Y' "

'------------------------------------------------------------------------------
'Set up the search criteria statements
'Organization Name
' if p_sc_org_group_id <> "" then
'    sSQL = sSQL & " AND org_group_id = " & p_sc_org_group_id
' end if

'First Name
 if p_sc_first_name <> "" then
    sSQL = sSQL & " AND org_group_id in (select distinct ug.org_group_id "
    sSQL = sSQL &                      " from egov_staff_directory_usergroups ug, users u "
    sSQL = sSQL &                      " where ug.userid = u.userid "
    sSQL = sSQL &                      " and u.staff_dir_display = 'Y' "
    sSQL = sSQL &                      " and u.orgid = " & iorgid
    sSQL = sSQL &                      " and UPPER(u.firstname) like ('%" & dbsafe(UCASE(trim(p_sc_first_name))) & "%')) "
 end if

'Last Name
 if p_sc_last_name <> "" then
    sSQL = sSQL & " AND org_group_id in (select distinct ug.org_group_id "
    sSQL = sSQL &                      " from egov_staff_directory_usergroups ug, users u "
    sSQL = sSQL &                      " where ug.userid = u.userid "
    sSQL = sSQL &                      " and u.staff_dir_display = 'Y' "
    sSQL = sSQL &                      " and u.orgid = " & iorgid
    sSQL = sSQL &                      " and UPPER(u.lastname) like ('%" & dbsafe(UCASE(trim(p_sc_last_name))) & "%')) "
 end if

'Phone Number
' if p_sc_phone_number <> "" then
'    sSQL = sSQL & " AND phone_number like ('%" & (trim(p_sc_phone_number)) & "%') "
' end if

'Phone Number Extension
' if p_sc_phone_number_ext <> "" then
'    sSQL = sSQL & " AND phone_number_ext like ('%" & trim(p_sc_phone_number_ext) & "%') "
' end if

'Email
' if p_sc_email <> "" then
'    sSQL = sSQL & " AND UPPER(email) like ('%" & UCASE(trim(p_sc_email)) & "%') "
' end if
'------------------------------------------------------------------------------

'Set up the ORDER BY
'If criteria has been entered then order by the ORG_NAME
'Otherwise, order by the org_level
 if lcl_search_criteria_entered = "N" then
    sSQL = sSQL & " ORDER BY org_level, UPPER(org_name) "
 else
    sSQL = sSQL & " ORDER BY UPPER(org_name) "
 end if
'------------------------------------------------------------------------------

 Set rs = Server.CreateObject("ADODB.Recordset")
 rs.Open sSQL, Application("DSN"), 3, 1

 if not rs.eof then
    while not rs.eof
       lcl_total_records_found = lcl_total_records_found + 1
       lcl_bgcolor             = changeBGColor(session("bgcolor"), "#ffffff", "#eeeeee")
       session("bgcolor")      = lcl_bgcolor
       lcl_phone               = ""

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

      'If any of the search criteria fields have been populated (besides the "show staff members" field)
      'then do NOT indent the results
       if lcl_search_criteria_entered = "N" then

         'Determine how far to indent the org_name
          lcl_indent = ((rs("org_level")-1) * 5)

          lcl_indent_spaces = ""
          for x = 1 to lcl_indent
              lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
          next
       else
          lcl_indent_spaces = ""
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
       else
          lcl_org_path = ""
       end if

       if rs("email") <> "" then
          'lcl_email = "<a href=""mailto:" & rs("email") & """>" & rs("email") & "</a>"
          lcl_email = "<a " & FormatMailToAsJavascript(rs("email")) & ">" & FormatEmailAsDecimal(rs("email")) & "</a>"
       else
          lcl_email = "&nbsp;"
       end if

       response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
       response.write "    <td colspan=""2"">" & lcl_indent_spaces & "<a href=""staff_directory_info.asp?org_group_id=" & rs("org_group_id") & """><nobr>" & rs("org_name") & "</nobr></a>" & lcl_org_path & "</td>" & vbcrlf
       response.write "    <td nowrap=""nowrap""><nobr>" & lcl_phone & "</nobr></td>" & vbcrlf
       response.write "    <td>" & lcl_email & "</td>" & vbcrlf
       response.write "</tr>" & vbcrlf

      'If the user wishes to display the members in each organizational group then
      'check for any users assigned to current organizational groups
       if p_sc_first_name <> "" OR p_sc_last_name <> "" OR p_sc_show_members = "Y" then
          lcl_sc_show_members = "Y"
       else
          lcl_sc_show_members = "N"
       end if

       if lcl_sc_show_members = "Y" then
          display_org_group_users rs("org_group_id"), rs("org_level"), lcl_search_criteria_entered, _
                                  p_sc_first_name, p_sc_last_name, lcl_bgcolor
       end if

      'Retrieve sub-organizational groups
       if lcl_search_criteria_entered = "N" then
          display_organizational_groups rs("org_group_id"), "", p_sc_first_name, p_sc_last_name, _
                                        lcl_sc_show_members, lcl_total_records_found
       end if

       rs.movenext
    wend
 else
    if lcl_total_records_found = 0 then
       response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
       response.write "    <td colspan=""3"" class=""noRecordsFound"">No Organizational Groups found</strong></td>" & vbcrlf
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
 sSQL = "SELECT g.org_group_id, g.org_name "
 sSQL = sSQL & " FROM egov_staff_directory_groups g "
 sSQL = sSQL & " WHERE g.org_group_id = (select g2.parent_org_group_id "
 sSQL = sSQL &                         " from egov_staff_directory_groups g2 "
 sSQL = sSQL &                         " where g2.org_group_id = " & p_org_group_id
 sSQL = sSQL &                         " and g2.orgid = " & iorgid
 sSQL = sSQL &                         " and g2.active_flag = 'Y') "
 sSQL = sSQL & " AND g.orgid = " & iorgid
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
sub display_org_group_users(p_org_group_id, p_org_level, p_search_criteria_entered, p_sc_first_name, p_sc_last_name, p_bgcolor)

   'If search criteria has been entered then the indentation amount is set to always be the same
    if p_search_criteria_entered = "Y" then
       lcl_org_level = 1
    else
       lcl_org_level = p_org_level
    end if

   'Determine how far to indent the org_name
    lcl_indent = (lcl_org_level * 5)

    lcl_indent_spaces = ""
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
    sSQLo = sSQLo & " FROM users u, egov_staff_directory_usergroups g "
    sSQLo = sSQLo & " WHERE u.userid = g.userid "
    sSQLo = sSQLo & " AND g.org_group_id = " & p_org_group_id
    sSQLo = sSQLo & " AND u.orgid = " & iorgid
    sSQLo = sSQLo & " AND u.staff_dir_display = 'Y' "

    if p_sc_first_name <> "" then
       sSQLo = sSQLo & " AND UPPER(u.firstname) like ('%" & UCASE(trim(DBsafe(p_sc_first_name))) & "%') "
    end if

    if p_sc_last_name <> "" then
       sSQLo = sSQLo & " AND UPPER(u.lastname) like ('%" & UCASE(trim(DBsafe(p_sc_last_name))) & "%') "
    end if

    sSQLo = sSQLo & " ORDER BY UPPER(u.lastname), UPPER(u.firstname) "

    set rso = Server.CreateObject("ADODB.Recordset")
    rso.Open sSQLo, Application("DSN"), 3, 1

    if not rso.eof then
       while not rso.eof
          if rso("jobtitle") <> "" then
             lcl_job_title = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong><em>Job Title: </em></strong>" & rso("jobtitle")
          else
             lcl_job_title = ""
          end if

         'Build the phone number
          if rso("businessnumber") <> "" then
             lcl_phone = rso("businessnumber")
          else
             lcl_phone = "&nbsp;"
          end if

          if rso("email") <> "" then
             lcl_email = rso("email")
          else
             lcl_email = ""
          end if

          response.write "<tr bgcolor=""" & p_bgcolor & """>" & vbcrlf

          if lcl_job_title <> "" then
             response.write "    <td><nobr>" & lcl_indent_spaces & "- " & rso("lastname") & ", " & rso("firstname") & "</nobr></td>" & vbcrlf
             response.write "    <td><nobr>" & lcl_job_title & "</nobr></td>" & vbcrlf
          else
             response.write "    <td colspan=""2""><nobr>" & lcl_indent_spaces & "- " & rso("lastname") & ", " & rso("firstname") & lcl_job_title & "</nobr></td>" & vbcrlf
          end if

          response.write "    <td><nobr>" & lcl_phone & "</nobr></td>" & vbcrlf
          response.write "    <td>"

          If lcl_email <> "" Then 
             response.write "<a " & FormatMailToAsJavascript(lcl_email) & ">" &  FormatEmailAsDecimal(lcl_email)  & "</a></td>" & vbcrlf
          Else
             response.write "&nbsp;"
          End If

          response.write "</td>" & vbcrlf
          response.write "</tr>" & vbcrlf

          rso.movenext
       wend
    end if
end sub

'-----------------------------------------------------------------------------
sub display_organizational_groups_dropdown(p_parent_org_group_id, p_org_level, p_sc_org_group_id)

'Retrieve all of the sub organizational groups
 sSQLg = "SELECT org_group_id, "
 sSQLg = sSQLg & " org_name "
 sSQLg = sSQLg & " FROM egov_staff_directory_groups "
 sSQLg = sSQLg & " WHERE orgid=" & iorgid

 if p_parent_org_group_id = "" then
    sSQLg = sSQLg & " AND (parent_org_group_id IS NULL OR parent_org_group_id = 0) "
 else
    sSQLg = sSQLg & " AND parent_org_group_id = " & p_parent_org_group_id
 end if

 sSQLg = sSQLg & " AND active_flag = 'Y' "
 sSQLg = sSQLg & " ORDER BY UPPER(org_name) "

 set rsg = Server.CreateObject("ADODB.Recordset")
 rsg.Open sSQLg, Application("DSN"), 3, 1

 if not rsg.eof then
    while not rsg.eof
      'Determine how far to indent the org_name
       lcl_indent = (p_org_level * 5)

       lcl_indent_spaces = ""
       for x = 1 to lcl_indent
           lcl_indent_spaces = lcl_indent_spaces & "&nbsp;"
       next

      'Determine if the current record is selected
       if p_sc_org_group_id <> "" then
          if CStr(p_sc_org_group_id) = CStr(rsg("org_group_id")) then
             lcl_selected = " selected"
          else
             lcl_selected = ""
          end if
       else
          lcl_selected = ""
       end if

       response.write "<option value=""" & rsg("org_group_id") & """" & lcl_selected & ">" & lcl_indent_spaces & rsg("org_name") & "</option>" & vbcrlf

      'Retrieve sub-organizational groups
       display_organizational_groups_dropdown rsg("org_group_id"), p_org_level+1, p_sc_org_group_id

       rsg.movenext
    wend
 end if
end sub

'-----------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = REPLACE(p_value,"'","''")
  end if

  dbsafe = lcl_return

end function
%>
