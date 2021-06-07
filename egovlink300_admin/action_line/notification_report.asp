<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "notifications") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for user permissions
 lcl_userhaspermission_alerts = userhaspermission(session("userid"),"alerts")

'Check for search options
 lcl_sc_form_name   = ""
 lcl_sc_category    = ""
 lcl_sc_department  = ""
 lcl_sc_form_active = ""
 lcl_sc_orderby     = ""

 if request("sc_form_name") <> "" then
    lcl_sc_form_name = request("sc_form_name")
 end if

 if request("sc_category") <> "" then
    lcl_sc_category = request("sc_category")
 end if

 if request("sc_department") <> "" then
    lcl_sc_department = request("sc_department")
 end if

 if request("sc_form_active") <> "" then
    lcl_sc_form_active = request("sc_form_active")
 end if

 if request("sc_orderby") <> "" then
    lcl_sc_orderby = request("sc_orderby")
 end if
%>
<html>
<head>
  <title><%=langBSActionLine%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script src="../scripts/selectAll.js"></script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
 	<% ShowHeader sLevel %>
 	<!--#Include file="../menu/menu.asp"--> 
<%
  blnCanManageActionAlerts = True 

 'BEGIN: Page Content --------------------------------------------------------
  response.write "<div id=""content"">" & vbcrlf
  response.write "	 <div id=""centercontent"">" & vbcrlf
  response.write "<p><font size=""+1""><strong>Action Line Notification Report</strong></font></p>" & vbcrlf

 'BEGIN: Search Options -------------------------------------------------------

 'Determine which dropdown values are selected
  lcl_selected_form_active        = ""
  lcl_selected_form_active_on     = ""
  lcl_selected_form_active_off    = ""
  lcl_selected_orderby_formname   = ""
  lcl_selected_orderby_category   = ""
  lcl_selected_orderby_department = ""

  if lcl_sc_form_active = "ON" then
     lcl_selected_form_active_on  = " selected=""selected"""
  elseif lcl_sc_form_active = "OFF" then
     lcl_selected_form_active_off = " selected=""selected"""
  end if

  if lcl_sc_orderby = "category" then
     lcl_selected_orderby_category   = " selected=""selected"""
  elseif lcl_sc_orderby = "department" then
     lcl_selected_orderby_department = " selected=""selected"""
  else
     lcl_selected_orderby_formname  = " selected=""selected"""
  end if

  response.write "<p>" & vbcrlf
  response.write "<fieldset>" & vbcrlf
  response.write "  <legend>Search Options:&nbsp;</legend>" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:5px"">" & vbcrlf
  response.write "    <form name=""searchForm"" id=""searchForm"" method=""post"" action=""notification_report.asp"">" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td>Action Line Form Name:</td>" & vbcrlf
  response.write "        <td><input type=""text"" name=""sc_form_name"" id=""sc_form_name"" value=""" & lcl_sc_form_name & """ size=""40"" maxlength=""50"" /></td>" & vbcrlf
  response.write "        <td>Form Active:</td>" & vbcrlf
  response.write "        <td>" & vbcrlf
  response.write "            <select name=""sc_form_active"" id=""sc_form_active"">" & vbcrlf
  response.write "              <option value="""""    & lcl_selected_form_active     & "></option>" & vbcrlf
  response.write "              <option value=""ON"""  & lcl_selected_form_active_on  & ">ON</option>" & vbcrlf
  response.write "              <option value=""OFF""" & lcl_selected_form_active_off & ">OFF</option>" & vbcrlf
  response.write "            </select>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td>Category:</td>" & vbcrlf
  response.write "        <td><input type=""text"" name=""sc_category"" id=""sc_category"" value=""" & lcl_sc_category & """ size=""40"" maxlength=""50"" /></td>" & vbcrlf
  response.write "        <td>Order By:</td>" & vbcrlf
  response.write "        <td>" & vbcrlf
  response.write "            <select name=""sc_orderby"" id=""sc_orderby"">" & vbcrlf
  response.write "              <option value=""action_form_name""" & lcl_selected_orderby_formname   & ">Action Line Form Name</option>" & vbcrlf
  response.write "              <option value=""category"""         & lcl_selected_orderby_category   & ">Category</option>" & vbcrlf
  response.write "              <option value=""department"""       & lcl_selected_orderby_department & ">Department</option>" & vbcrlf
  response.write "            </select>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td>Department:</td>" & vbcrlf
  response.write "        <td><input type=""text"" name=""sc_department"" id=""sc_department"" value=""" & lcl_sc_department & """ size=""40"" maxlength=""50"" /></td>" & vbcrlf
  response.write "        <td colspan=""2"">&nbsp;</td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td colspan=""4"">" & vbcrlf
  response.write "            <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    </form>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
  response.write "</p>" & vbcrlf
 'END: Search Options ---------------------------------------------------------

  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf

 	if blnCanManageActionAlerts then
     session("RedirectPage") = "notification_report.asp"
     list_forms session("orgid"), lcl_sc_form_name, lcl_sc_category, lcl_sc_department, lcl_sc_form_active, lcl_sc_orderby
  else
     response.write "<div class=""orgadminboxf"">" & vbcrlf
     response.write "<p>" & vbcrlf
     response.write "  <strong>Security Alert!</strong><br />" & vbcrlf
     response.write "  You do not have permission to access the <strong>E-Gov Alert Manager</strong>.  Please contact your E-Gov Link administrator to inquire about gaining access to the <strong>E-Gov Alert Manager</strong>." & vbcrlf
     response.write "</p>" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub list_forms(p_orgid, p_sc_form_name, p_sc_category, p_sc_department, p_sc_form_active, p_sc_orderby)

 'Build WHERE clause
  sWhereClause = ""

  if p_sc_form_name <> "" then
     sWhereClause = sWhereClause & " AND UPPER(action_form_name) LIKE ('%" & UCASE(p_sc_form_name)  & "%') "
  end if

  if p_sc_category <> "" then
     sWhereClause = sWhereClause & " AND UPPER(form_category_name) LIKE ('%" & UCASE(p_sc_category)  & "%') "
  end if

  if p_sc_department <> "" then
     sWhereClause = sWhereClause & " AND UPPER(DeptName) LIKE ('%" & UCASE(p_sc_department)  & "%') "
  end if

  if p_sc_form_active <> "" then
     if p_sc_form_active = "OFF" then
        sWhereClause = sWhereClause & " AND action_form_enabled = 0 "
     else
        sWhereClause = sWhereClause & " AND action_form_enabled = 1 "
     end if
  end if

 'Set up the ORDER BY
  if p_sc_orderby <> "" then
     if p_sc_orderby = "action_form_name" then
        sOrderBy = "action_form_name, form_category_name"
     elseif p_sc_orderby = "category" then
        sOrderBy = "form_category_name, action_form_name"
     elseif p_sc_orderby = "department" then
        sOrderBy = "DeptName, action_form_name"
     end if
  else
     sOrderBy = "action_form_name, form_category_name"
  end if

		sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_forms_categories_view "
  sSQL = sSQL & " WHERE orgid = " & p_orgid
  sSQL = sSQL & sWhereClause
  sSQL = sSQL & " ORDER BY " & sOrderBy

	 set oRequests = Server.CreateObject("ADODB.Recordset")
 	oRequests.Open sSQL, Application("DSN"), 3, 1

  bgcolor       = "#eeeeee"
  lcl_linecount = 0
	
  if not oRequests.eof then
    	if request("useSessions") = 1 then
     		 if Len(session("pagenum")) <> 0 then
       				oRequests.AbsolutePage = clng(session("pagenum"))	
				       session("pageNum")     = clng(session("pagenum"))		
        else
       				oRequests.AbsolutePage = 1
				       session("pageNum")     = 1
        end if
   	 else
      	 if Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1 then
    		   		oRequests.AbsolutePage = 1
       				session("pageNum")     = 1
     		 else
       				if clng(Request("pagenum")) <= oRequests.PageCount then
     	   					oRequests.AbsolutePage = Request("pagenum")
     				   		session("pageNum")     = Request("pagenum")
       				else
        						oRequests.AbsolutePage = 1
					        	session("pageNum")     = 1
           end if
        end if
     end if

    'Display Record Statistics
     Dim abspage, pagecnt
     abspage = oRequests.AbsolutePage
     pagecnt = oRequests.PageCount

    'Replace PAGENUM field with random field for navigation purposes
     sQueryString  = replace(request.querystring,"pagenum","HFe301")

     response.write "<div class=""shadow"">" & vbcrlf
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" class=""tablelist"">" & vbcrlf
     response.write "  <tr class=""tablelist"">" & vbcrlf
     response.write "      <th>ID</th>" & vbcrlf
     response.write "      <th align=""left"">Action Line Form Name</th>" & vbcrlf
     response.write "      <th align=""left"">Category</th>" & vbcrlf
     response.write "      <th align=""left"">Department</th>" & vbcrlf
     response.write "      <th>Form Active</th>" & vbcrlf

    	do while not oRequests.eof
        bgcolor            = changeBGColor(bgcolor,"#eeeeee","#ffffff")
        lcl_linecount      = lcl_linecount + 1
        lcl_tr_onclick     = ""
        lcl_tr_onmouseover = ""
        lcl_tr_onmouseout  = ""

       'Determine if the user has permission assigned to maintain alerts
      		if lcl_userhaspermission_alerts then
        			lcl_tr_onclick     = " onclick=""location.href='edit_form.asp?control=" & oRequests("action_form_id") & "';"""
           lcl_tr_onmouseover = " onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"""
           lcl_tr_onmouseout  = " onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"""
        end if

     		'Determine if the form is available
        sEnabledColor = "red"
        sEnabledLabel = "OFF"
        blnEnabled    = 1

      		if oRequests("action_form_enabled") then
           sEnabledColor = "green"
           sEnabledLabel = "ON"
           blnEnabled    = 0
        end if

    				sEnabled = "<font style=""color:" & sEnabledColor & ";font-size:10px;"">" & sEnabledLabel & "</font>"

       'Format the columns
        lcl_display_action_form_id     = "&nbsp"
        lcl_display_action_form_name   = "&nbsp"
        lcl_display_form_category_name = "&nbsp"
        lcl_display_deptname           = "&nbsp"

        if oRequests("action_form_id") <> "" then
           lcl_display_action_form_id = oRequests("action_form_id")
        end if

        if trim(oRequests("action_form_name")) <> "" then
           lcl_display_action_form_name = trim(oRequests("action_form_name"))
        end if

        if trim(oRequests("form_category_name")) <> "" then
           lcl_display_form_category_name = trim(oRequests("form_category_name"))
        end if

        if trim(oRequests("deptname")) <> "" then
           lcl_display_deptname = trim(oRequests("deptname"))
        end if

      		response.write "  <tr bgcolor=""" & bgcolor & """" & lcl_tr_onclick & lcl_tr_onmouseover & lcl_tr_onmouseout & ">" & vbcrlf
        response.write "      <td align=""center"" width=""25""><strong>"      & lcl_display_action_form_id     & "</strong></td>" & vbcrlf
      		response.write "      <td nowrap=""nowrap""><strong>" & lcl_display_action_form_name   & "</strong></td>" & vbcrlf
      		response.write "      <td nowrap=""nowrap"">"         & lcl_display_form_category_name & "</td>" & vbcrlf
      		response.write "      <td nowrap=""nowrap"">"         & lcl_display_deptname           & "</td>" & vbcrlf
        response.write "      <td align=""center"">"          & sEnabled                       & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

      		ShowNotifyNames     oRequests("action_form_id"), bgcolor 
      		ShowEscalationNames oRequests("action_form_id"), bgcolor
        ShowNotifications   oRequests("action_form_id"), bgcolor

      		oRequests.MoveNext 
     loop

   	 response.write "</table>" & vbcrlf
   	 response.write "</div>" & vbcrlf
     response.write "<div align=""right""><strong>Total: </strong>[" & lcl_linecount & "]</div>" & vbcrlf

  end if

end sub

'------------------------------------------------------------------------------
sub ShowNotifyNames(iAction_form_id, sBgcolor)
	 Dim sSql

 	sSQL = "SELECT F.action_form_id, U.firstname, U.lastname, isnull(O.firstname,'') as firstname2, "
 	sSQL = sSQL & "isnull(O.lastname,'') as lastname2, isnull(P.firstname,'') as firstname3, isnull(P.lastname,'') as lastname3 "
 	sSQL = sSQL & "FROM egov_action_request_forms F "
  sSQL = sSQL &     " LEFT OUTER JOIN users U ON U.userid = F.assigned_userid "
  sSQL = sSQL &     " LEFT OUTER JOIN users O ON O.userid = F.assigned_userid2 "
  sSQL = sSQL &     " LEFT OUTER JOIN users P ON P.userid = F.assigned_userid3 "
 	sSQL = sSQL & "WHERE action_form_id = " & iAction_form_id

 	set oNotify = Server.CreateObject("ADODB.Recordset")
	 oNotify.Open sSQL, Application("DSN"), 3, 1

 	if not oNotify.EOF then
    'Build Notify list
     lcl_notify_names = ""

     if oNotify("firstname") <> "" OR oNotify("lastname") <> "" then
        if lcl_notify_names <> "" then
           lcl_notify_names = lcl_notify_names & ", " & trim(oNotify("firstname") & " " & oNotify("lastname"))
        else
           lcl_notify_names = trim(oNotify("firstname") & " " & oNotify("lastname"))
        end if
     end if

     if oNotify("firstname2") <> "" OR oNotify("lastname2") <> "" then
        if lcl_notify_names <> "" then
           lcl_notify_names = lcl_notify_names & ", " & trim(oNotify("firstname2") & " " & oNotify("lastname2"))
        else
           lcl_notify_names = trim(oNotify("firstname2") & " " & oNotify("lastname2"))
        end if
     end if

     if oNotify("firstname3") <> "" OR oNotify("lastname3") <> "" then
        if lcl_notify_names <> "" then
           lcl_notify_names = lcl_notify_names & ", " & trim(oNotify("firstname3") & " " & oNotify("lastname3"))
        else
           lcl_notify_names = trim(oNotify("firstname3") & " " & oNotify("lastname3"))
        end if
     end if

  	 	response.write "  <tr bgcolor=""" & sBgcolor & """>" & vbcrlf
     response.write "      <td>&nbsp;</td>" & vbcrlf
     response.write "      <td colspan=""4"">" & vbcrlf
     response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""padding-left:5px; padding-top:5px; background-color:" & sBgcolor & ";"">" & vbcrlf
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td width=""100"" style=""color:#800000"">Alert/Notify:</td>" & vbcrlf
     response.write "                <td>" & lcl_notify_names & "</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "          </table>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 	oNotify.close 
	 set oNotify = nothing 

end sub

'------------------------------------------------------------------------------
sub ShowEscalationNames(iAction_form_id, sBgcolor)
 	Dim sSql, iRow

 	iRow = 0

 	response.write "  <tr bgcolor=""" & sBgcolor & """>" & vbcrlf
  response.write "      <td>&nbsp;</td>" & vbcrlf
  response.write "      <td colspan=""4"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""padding-left:5px; padding-top:5px; background-color:" & sBgcolor & ";"">" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td width=""100"" style=""color:#800000"">Escalations:</td>" & vbcrlf
  response.write "                <td colspan=""4"">" & vbcrlf

 	sSQL = "SELECT escnotify, esccriteria, esctime, firstname, lastname "
 	sSQL = sSQL & " FROM egov_action_escalations"
	sSQL = sSQL & " LEFT JOIN  users ON escnotify = userid "
 	sSQL = sSQL & " WHERE action_form_id = " & iAction_form_id
  sSQL = sSQL & " ORDER BY esccriteria, esctime"

 	set oEscNames = Server.CreateObject("ADODB.Recordset")
 	oEscNames.Open sSQL, Application("DSN"), 3, 1

  if not oEscNames.eof then
    	do while not oEscNames.eof
      		iRow = iRow + 1

	if oEscNames("ESCNotify") = -1 then
		response.write "Employee Assigned to Request"
	else
        	response.write oEscNames("firstname") & " " & oEscNames("lastname") 
	end if
	response.write " &ndash; " & oEscNames("esccriteria") & " " & oEscNames("esctime") & " day(s)<br />" & vbcrlf

      		oEscNames.movenext
    	loop
  else
     response.write "None set" & vbcrlf
 	end if

 	response.write "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 	oEscNames.close
	 set oEscNames = nothing

end sub

'------------------------------------------------------------------------------
sub ShowNotifications(iActionFormID, sBGColor)
 	Dim sSql, iRow

 	iRow = 0

 	response.write "  <tr bgcolor=""" & sBgcolor & """>" & vbcrlf
  response.write "      <td>&nbsp;</td>" & vbcrlf
  response.write "      <td colspan=""4"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""padding-left:5px; padding-top:5px; background-color:" & sBgcolor & ";"">" & vbcrlf
  response.write "            <tr valign=""top"">" & vbcrlf
  response.write "                <td width=""100"" style=""color:#800000"">Notifications:</td>" & vbcrlf
  response.write "                <td colspan=""4"">" & vbcrlf

 	sSQL = "SELECT n.sendto, u.firstname, u.lastname, n.email_action "
 	sSQL = sSQL & " FROM egov_action_notifications n, users u "
 	sSQL = sSQL & " WHERE n.sendto = u.userid "
 	sSQL = sSQL & " AND n.action_form_id = " & iActionFormID
  sSQL = sSQL & " ORDER BY u.lastname, u.firstname, n.notificationid "

 	set oNotifications = Server.CreateObject("ADODB.Recordset")
 	oNotifications.Open sSQL, Application("DSN"), 3, 1

  if not oNotifications.eof then
    	do while not oNotifications.eof
      		iRow = iRow + 1

       'Format the "email_action"
        lcl_display_email_action = ""

        if oNotifications("email_action") <> "" then
           if oNotifications("email_action") = "request_closed" then
              lcl_display_email_action = "Requests set to closed status (RESOLVED or DISMISSED)"
           else
              lcl_display_email_action = "Requests are updated"
           end if
        end if

        response.write oNotifications("firstname") & " " & oNotifications("lastname") & " &ndash; " & lcl_display_email_action & "<br />" & vbcrlf

      		oNotifications.movenext
    	loop
  else
     response.write "None set" & vbcrlf
 	end if

 	response.write "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 	oNotifications.close 
	 set oNotifications = nothing

end sub
%>


