<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: pool_attendance_list.asp
' AUTHOR:   David Boyer
' CREATED:  04/30/08
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  to allow users to maintain daily statistics about the pool
'
' MODIFICATION HISTORY
' 1.0  04/30/08  David Boyer - Created Code.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel     = "../" ' Override of value from common.asp
lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

if NOT UserHasPermission( session("userid"), "pool_attendance_view" ) then
   response.redirect sLevel & "permissiondenied.asp"
end if

'Retrieve the search criteria fields
 if UCASE(request("use_sessions")) = "Y" then
    lcl_sc_from_date = session("sc_from_date")
    lcl_sc_to_date   = session("sc_to_date")
    lcl_sc_order_by  = session("sc_order_by")
    lcl_init         = session("init")
 else
    lcl_sc_from_date = request("sc_from_date")
    lcl_sc_to_date   = request("sc_to_date")
    lcl_sc_order_by  = request("sc_order_by")
    lcl_init         = request("init")
 end if

'Check to see if this is the initial run of the screen
 if lcl_init <> "N" then
    lcl_init = "Y"
 end if

'If this is the initial run of the screen then default the From/To Date fields
 if lcl_init = "Y" then
    lcl_sc_from_date = month(date()) & "/1/" & year(date())
    lcl_sc_to_date   = date()
 end if

'Setup the session variables
 session("sc_from_date") = lcl_sc_from_date
 session("sc_to_date")   = lcl_sc_to_date
 session("sc_order_by")  = lcl_sc_order_by
 session("init")         = lcl_init

'-- Delete Pool Record -----------------------------------------------
 if request.querystring("CMD") = "delete_poolinfoid" then
    lcl_poolinfoid = request.querystring("PID")

   'Delete all of the weather records
    sSQLd1 = "DELETE FROM egov_pool_weather_log WHERE poolinfoid = " & lcl_poolinfoid
    set rsd1 = Server.CreateObject("ADODB.Recordset")
    rsd1.Open sSQLd1, Application("DSN"), 3, 1

   'Delete all of the incident records
    sSQLd2 = "DELETE FROM egov_pool_incidents_log WHERE poolinfoid = " & lcl_poolinfoid
    set rsd2 = Server.CreateObject("ADODB.Recordset")
    rsd2.Open sSQLd2, Application("DSN"), 3, 1

   'Delete all of the notes records
    sSQLd3 = "DELETE FROM egov_pool_info_notes WHERE poolinfoid = " & lcl_poolinfoid
    set rsd3 = Server.CreateObject("ADODB.Recordset")
    rsd3.Open sSQLd3, Application("DSN"), 3, 1

   'Now delete the main poolinfo record
    sSQLd = "DELETE FROM egov_pool_info WHERE poolinfoid = " & lcl_poolinfoid
    set rsd = Server.CreateObject("ADODB.Recordset")
    rsd.Open sSQLd, Application("DSN"), 3, 1

    lcl_success = "SD"

    set rsd  = nothing
    set rsd1 = nothing
    set rsd2 = nothing
    set rsd3 = nothing
    set rsd4 = nothing
 end if

'Determine if the user has Add, Edit, and/or Reporting roles.
'If so enable the proper ones.  Otherwise, diable them.
 lcl_enable_add     = "Y"
 lcl_enable_edit    = "Y"
 lcl_enable_reports = "Y"

 if NOT UserHasPermission( session("userid"), "pool_attendance_add" ) then
    lcl_enable_add = "N"
 end if

 if NOT UserHasPermission( session("userid"), "pool_attendance_edit" ) then
    lcl_enable_edit = "N"
 end if

 if NOT UserHasPermission( session("userid"), "pool_attendance_reporting" ) then
    lcl_enable_reports = "N"
 end if
%>
<html>
<head>
  <title>E-GovLink {Daily Pool Attendance}</title>
  
	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
 <script src="../scripts/selectAll.js"></script>
 <script language="javascript" src="../scripts/modules.js"></script>

<script language="javascript">
 function doCalendar(sField) {
   var w = (screen.width - 350)/2;
   var h = (screen.height - 350)/2;
   eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=pool_list", "_poollist", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
 }

function deleteconfirm(p_poolinfo_id) {
  lcl_pool_date = document.getElementById("pool_date_"+p_poolinfo_id).innerHTML;
  input_box=confirm("Are you sure you want to delete the \"" + lcl_pool_date + "\" attendance record?");
  if (input_box==true) {
      // DELETE HAS BEEN VERIFIED
      location.href='pool_attendance_list.asp?cmd=delete_poolinfoid&pid='+ p_poolinfo_id;
  }else{
      // CANCEL DELETE PROCESS
  }
}

checked=false;
function checkedAll () {
	 var x = document.getElementById('pool_list');
	 if (checked == false) {
      checked = true
  }else{
      checked = false
  }

 	for (var i=0; i < x.elements.length; i++) {
     	 x.elements[i].checked = checked;
 	}
}

function runReport(p_report) {
  //When running a report then switch the action of the form to point to the report screen
  lcl_form        = document.getElementById("pool_list");
  lcl_total_count = document.getElementById("p_total_records").value;
  lcl_form.action = "pool_reports.asp?report_type="+p_report+"&total_records="+lcl_total_count;
  document.getElementById("pool_list").submit();
}

function poolSearch() {
  //When running a report then switch the action of the form to point to the report screen
  lcl_form = document.getElementById("pool_list");
  lcl_form.action = "pool_attendance_list.asp"
  document.getElementById("pool_list").submit();
}

function clearMsg() {
  document.getElementById("status_message").innerHTML = "&nbsp;<p>";
}
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"> <!-- onLoad="document.searchform.username.focus()"> -->
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"-->

<div id="content">
    <div id="centercontent">

<table border="0" cellpadding="0" cellspacing="0" class="start" width="100%">
  <form action="pool_attendance_list.asp" method="post" name="pool_list" id="pool_list">
    <input type="<%=lcl_hidden%>" name="init" value="N" size="1" maxlength="1">
  <tr>
      <td><font size="+1"><strong><%=session("sOrgName")%>&nbsp;Daily Pool Attendance Maintenance</strong></font></td>
  </tr>
  <tr>
      <td>
          <fieldset>
            <legend><b>Search/Sort Criteria</b>&nbsp;</legend><p>
         			<table border="0" cellpadding="5" cellspacing="0" style="width: 600px">
       			    <tr>
                  <td nowrap="nowrap">From Date:</td>
                  <td>
                      <input type="text" name="sc_from_date" value="<%=lcl_sc_from_date%>" size="10" maxlength="10">&nbsp;
                      <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('sc_from_date');" /></span>
                  </td>
                  <td nowrap="nowrap">To Date:</td>
                  <td>
                      <input type="text" name="sc_to_date" value="<%=lcl_sc_to_date%>" size="10" maxlength="10">&nbsp;
                      <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('sc_to_date');" /></span>
                  </td>
                  <td>Sort:</td>
                  <td><% displaySortOptions lcl_sc_order_by %></td>
              </tr>
              <tr>
                  <td colspan="4"><input type="button" class="button" value="Search" onclick="poolSearch();"></td>
              </tr>
            </table>
          </fieldset>
          <p>
      </td>
  </tr>
  <tr>
      <td valign="top">
          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <%
              lcl_message = ""

              if request("success") = "SU" then
                 lcl_message = "<b style=""color:#FF0000"">*** Successfully Updated... ***</b>"
              elseif request("success") = "SA" then
                 lcl_message = "<b style=""color:#FF0000"">*** Successfully Created... ***</b>"
              elseif lcl_success = "SD" then
                 lcl_message = "<b style=""color:#FF0000"">*** Successfully Deleted... ***</b>"
              else
                 lcl_message = "&nbsp;"
              end if

             'Display the message if it exists
'              if lcl_message <> "&nbsp;" then
                 response.write "            <caption id=""status_message"" align=""right"">" & lcl_message & "<p></caption>" & vbcrlf
'              end if

             'Display Add link and/or Reporting buttons depending on user permissions
              if lcl_enable_add = "Y" OR lcl_enable_reports = "Y" then
                 response.write "            <tr>" & vbcrlf
                 response.write "                <td>" & vbcrlf
                 response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
                 response.write "                      <tr>" & vbcrlf
                 response.write "                          <td>" & vbcrlf

                'Show/Hide the Add link depending on user permissions
                 if lcl_enable_add = "Y" then
'                    response.write "                              <a href=""pool_attendance_maint.asp?pid=0"">" & vbcrlf
'                    response.write "                              <img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;New Daily Attendance</a>" & vbcrlf
                    response.write "                              <input type=""button"" id=""new_attendance_button"" value=""New Daily Attendance"" onclick=""clearMsg();location.href='pool_attendance_maint.asp?pid=0'"">" & vbcrlf
                 else
                    response.write "&nbsp;"
                 end if

                 response.write "                          </td>" & vbcrlf
                 response.write "                          <td>" & vbcrlf

                'Show/Hide the REPORTING buttons depending on user permissions
                 if lcl_enable_reports = "Y" then
                    response.write "                              <input type=""button"" name=""daily_attendance_report"" value=""Daily Attendance Report"" onclick=""clearMsg();runReport('DAILY');"">" & vbcrlf
                    response.write "                              <input type=""button"" name=""attendance_summary"" value=""Attendance Summary"" onclick=""clearMsg();runReport('SUMMARY');"">" & vbcrlf
                    response.write "                              <input type=""button"" name=""incident_summary"" value=""Injury/Incident Report"" onclick=""clearMsg();runReport('INCIDENT');"">" & vbcrlf
                 else
                    response.write "&nbsp;"
                 end if

                 response.write "                          </td>" & vbcrlf
                 response.write "                      </tr>" & vbcrlf
                 response.write "                    </table>" & vbcrlf
                 response.write "                </td>" & vbcrlf
                 response.write "            </tr>" & vbcrlf
              end if

              response.write "          </table>" & vbcrlf

 'Retrieve all of the daily pool attendance records
  sSQL = "SELECT poolinfoid, orgid, pool_date, total_members, total_punchcards, total_guests, total_groups_peoplecount "
  sSQL = sSQL & " FROM egov_pool_info_vw "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")

 'Evaluate the search criteria
  if lcl_sc_from_date <> "" AND lcl_sc_to_date <> "" then
     sSQL = sSQL & " AND pool_date BETWEEN '" & lcl_sc_from_date & "' AND '" & lcl_sc_to_date & "'"

  elseif lcl_sc_from_date <> "" AND lcl_sc_to_date = "" then
     sSQL = sSQL & " AND pool_date >= '" & lcl_sc_from_date & "'"

  elseif lcl_sc_from_date = "" AND lcl_sc_to_date <> "" then
     sSQL = sSQL & " AND pool_date <= '" & lcl_sc_to_date & "'"
  end if

 'Setup the ORDER BY
  if lcl_sc_order_by <> "" then
     lcl_order_by = REPLACE(lcl_sc_order_by,"_DESC"," DESC")

     if lcl_sc_order_by <> "POOL_DATE_DESC" AND lcl_sc_order_by <> "POOL_DATE" then
        lcl_order_by = lcl_order_by & ", pool_date desc"
     end if
  else
     lcl_order_by = "pool_date DESC "
  end if

  sSQL = sSQL & " ORDER BY " & lcl_order_by

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

  if rs.eof then
     response.write "<p><strong>No records found</strong>" & vbcrlf
  else
   		response.write "<div class=""shadow"">" & vbcrlf
   		response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" width=""100%"">" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <th><input type=""checkbox"" name=""selectDeselectAll"" id=""checkbox"" value=""Y"" onclick=""checkedAll()""></th>" & vbcrlf
     response.write "      <th align=""left"" width=""40%"">Date</th>" & vbcrlf
     response.write "      <th>Total<br>Members</th>" & vbcrlf
     response.write "      <th>Total<br>Punchcards</th>" & vbcrlf
     response.write "      <th>Total<br>Guests/Daily Rate</th>" & vbcrlf
     response.write "      <th>Total<br>Group Rate</th>" & vbcrlf
     response.write "      <th>Overall<br>Attendance</th>" & vbcrlf

    'If the user has the EDIT role then show the DELETE column header
     if lcl_enable_edit = "Y" then
        response.write "      <th>Del</th>" & vbcrlf
     else
        response.write "      <th>&nbsp;</th>" & vbcrlf
     end if

     response.write "  </tr>" & vbcrlf

     lcl_bgcolor = "#FFFFFF"
     iRowCount   = 0

     while not rs.eof
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")
        iRowCount   = iRowCount + 1

       'Get the totals for each attendance type
        lcl_total_members    = rs("total_members")
        lcl_total_punchcards = rs("total_punchcards")
        lcl_total_guests     = rs("total_guests")
        lcl_total_groups     = rs("total_groups_peoplecount")

'        getTotalMembers rs("poolinfoid"), lcl_total_members, lcl_total_punchcards, lcl_total_guests, lcl_total_groups

        if lcl_total_members = 0 then
           lcl_total_members = ""
        end if

        if lcl_total_punchcards = 0 then
           lcl_total_punchcards = ""
        end if

        if lcl_total_guests = 0 then
           lcl_total_guests = ""
        end if

        if lcl_total_groups = 0 then
           lcl_total_groups = ""
        end if

       'Show/Hide the EDIT "onclick" depending on user permissions
        if lcl_enable_edit = "Y" then
           lcl_onclick = " onclick=""location.href='pool_attendance_maint.asp?pid=" & rs("poolinfoid") & "'"""
        else
           lcl_onclick = ""
        end if

        response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """ id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
        response.write "      <td><input type=""checkbox"" name=""checkbox_" & iRowCount & """ id=""checkbox_" & iRowCount & """ value=""" & rs("poolinfoid") & """></td>" & vbcrlf
        response.write "      <td align=""left"" id=""pool_date_" & rs("poolinfoid") & """" & lcl_onclick & ">" & rs("pool_date") & "</td>" & vbcrlf
        response.write "      <td" & lcl_onclick & ">" & lcl_total_members    & "</td>" & vbcrlf
        response.write "      <td" & lcl_onclick & ">" & lcl_total_punchcards & "</td>" & vbcrlf
        response.write "      <td" & lcl_onclick & ">" & lcl_total_guests     & "</td>" & vbcrlf
        response.write "      <td" & lcl_onclick & ">" & lcl_total_groups     & "</td>" & vbcrlf
'        response.write "      <td>" & REPLACE(rs("total_attendance"),0,"") & "</td>" & vbcrlf

       'Calculate the Overall Attendance
        lcl_total_attendance = 0

        if lcl_total_members = "" OR isnull(lcl_total_members) then
           lcl_total_members = 0
        end if

        if lcl_total_punchcards = "" OR isnull(lcl_total_punchcards) then
           lcl_total_punchcards = 0
        end if

        if lcl_total_guests = "" OR isnull(lcl_total_guests) then
           lcl_total_guests = 0
        end if

        if lcl_total_groups = "" OR isnull(lcl_total_groups) then
           lcl_total_groups = 0
        end if

        lcl_total_attendance = lcl_total_members+lcl_total_punchcards+lcl_total_guests+lcl_total_groups

        if lcl_total_attendance = 0 then
           lcl_total_attendance = ""
        end if

        response.write "      <td" & lcl_onclick & ">" & lcl_total_attendance & "</td>" & vbcrlf
'response.write "<td>" & lcl_total_members & "+" & lcl_total_punchcards & "+" & lcl_total_guests & "+" & lcl_total_groups & "</td>"

       'If the user has the EDIT role then display the DELETE button.
        if lcl_enable_edit = "Y" then
           response.write "      <td><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""clearMsg();deleteconfirm(" & rs("poolinfoid") & ")""></td>" & vbcrlf
        else
           response.write "      <td>&nbsp;</td>" & vbcrlf
        end if
        response.write "  </tr>" & vbcrlf

        rs.movenext
     wend

     response.write "  <input type=""" & lcl_hidden & """ name=""p_total_records"" id=""p_total_records"" value=""" & iRowCount & """>" & vbcrlf
     response.write "</table>" & vbcrlf

  end if
%>
      </td>
  </tr>
  </form>
</table>
    </div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
sub displaySortOptions(p_value)

  select case UCASE(p_value)
     case "POOL_DATE_DESC"
           lcl_selected_date_desc       = " selected"
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "POOL_DATE"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = " selected"
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_MEMBERS_DESC"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = " selected"
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_MEMBERS"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = " selected"
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_PUNCHCARDS_DESC"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = " selected"
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_PUNCHCARDS"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = " selected"
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_GUESTS_DESC"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = " selected"
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_GUESTS"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = " selected"
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = ""
     case "TOTAL_GROUPS_DESC"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = " selected"
           lcl_selected_groups          = ""
     case "TOTAL_GROUPS"
           lcl_selected_date_desc       = ""
           lcl_selected_date            = ""
           lcl_selected_members_desc    = ""
           lcl_selected_members         = ""
           lcl_selected_punchcards_desc = ""
           lcl_selected_punchcards      = ""
           lcl_selected_guests_desc     = ""
           lcl_selected_guests          = ""
           lcl_selected_groups_desc     = ""
           lcl_selected_groups          = " selected"
  end select

  response.write "<select name=""sc_order_by"">" & vbcrlf
  response.write "  <option value=""POOL_DATE_DESC"""        & lcl_selected_date_desc       & ">DATE [Recent to Past]</option>" & vbcrlf
  response.write "  <option value=""POOL_DATE"""             & lcl_selected_date            & ">DATE [Past to Recent]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_MEMBERS_DESC"""    & lcl_selected_members_desc    & ">Total Members [High to Low]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_MEMBERS"""         & lcl_selected_members         & ">Total Members [Low to High]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_PUNCHCARDS_DESC""" & lcl_selected_punchcards_desc & ">Total Punchcards [High to Low]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_PUNCHCARDS"""      & lcl_selected_punchcards      & ">Total Punchcards [Low to High]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_GUESTS_DESC"""     & lcl_selected_guests_desc     & ">Total Guests/Daily Rate [High to Low]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_GUESTS"""          & lcl_selected_guests          & ">Total Guests/Daily Rate [Low to High]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_GROUPS_DESC"""     & lcl_selected_groups_desc     & ">Total Group Rate [High to Low]</option>" & vbcrlf
  response.write "  <option value=""TOTAL_GROUPS"""          & lcl_selected_groups          & ">Total Group Rate [Low to High]</option>" & vbcrlf
  response.write "</select>" & vbcrlf
end sub

'--------------------------------------------------------------------
 sub getTotalMembers(ByVal p_poolinfo_id, ByRef lcl_total_members, ByRef lcl_total_punchcards, ByRef lcl_total_guests, ByRef lcl_total_groups )
  sSQL = "SELECT total_members, total_punchcards, total_guests, total_groups "
  sSQL = sSQL & " FROM egov_pool_info_vw "
  sSQL = sSQL & " WHERE poolinfoid = " & p_poolinfo_id

  set oTotals = Server.CreateObject("ADODB.Recordset")
  oTotals.Open sSQL, Application("DSN") , 3, 1

  if not oTotals.eof then
     lcl_total_members    = oTotals("total_members")
     lcl_total_punchcards = oTotals("total_punchcards")
     lcl_total_guests     = oTotals("total_guests")
     lcl_total_groups     = oTotals("total_groups")
  else
     lcl_total_members    = 0
     lcl_total_punchcards = 0
     lcl_total_guests     = 0
     lcl_total_groups     = 0
  end if

  set oTotals = nothing

end sub
%>