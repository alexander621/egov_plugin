<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"), "alerts" ) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 blnCanManageActionAlerts = True 

'Retrieve the search options
 lcl_sc_formname   = ""
 lcl_sc_formactive = ""
 lcl_sc_orderby    = "action_form_name"

 if request("sc_formname") <> "" then
    lcl_sc_formname = request("sc_formname")
 end if

 if request("sc_formactive") <> "" then
    lcl_sc_formactive = request("sc_formactive")
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
 'BEGIN: Page Content ---------------------------------------------------------
 'Determine the Order By "selected" value
  lcl_selected_orderby_action_form_name = " selected=""selected"""
  lcl_selected_orderby_category         = ""
  lcl_selected_orderby_department       = ""
  lcl_selected_orderby_assignedto       = ""

  if lcl_sc_orderby = "category" then
     lcl_selected_orderby_action_form_name = ""
     lcl_selected_orderby_category         = " selected=""selected"""
     lcl_selected_orderby_department       = ""
     lcl_selected_orderby_assignedto       = ""
  elseif lcl_sc_orderby = "department" then
     lcl_selected_orderby_action_form_name = ""
     lcl_selected_orderby_category         = ""
     lcl_selected_orderby_department       = " selected=""selected"""
     lcl_selected_orderby_assignedto       = ""
  elseif lcl_sc_orderby = "assignedto" then
     lcl_selected_orderby_action_form_name = ""
     lcl_selected_orderby_category         = ""
     lcl_selected_orderby_department       = ""
     lcl_selected_orderby_assignedto       = " selected=""selected"""
  end if

 'Determine the form active "selected" value
  lcl_selected_active_on  = ""
  lcl_selected_active_off = ""

  if lcl_sc_formactive = "ON" then
     lcl_selected_active_on  = " selected=""selected"""
     lcl_selected_active_off = ""
  elseif lcl_sc_formactive = "OFF" then
     lcl_selected_active_on  = " selected=""selected"""
     lcl_selected_active_off = ""
  end if

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>(E-Gov Alert Manager) - Manage Action Line Request Forms</strong></font></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <fieldset class=""fieldset"">" & vbcrlf
  response.write "            <legend>Search Options&nbsp;</legend>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin-top:10pt"">" & vbcrlf
  response.write "              <form name=""searchForm"" id=""searchForm"" method=""post"" action=""manage_action_forms.asp"">" & vbcrlf
  response.write "              <tr valign=""top"">" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      Form Name: " & vbvcrlf
  response.write "                      <input type=""text"" name=""sc_formname"" id=""sc_formname"" value=""" & lcl_sc_formname & """ size=""50"" maxlength=""50"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td>Form Active:</td>" & vbcrlf
  response.write "                            <td>" & vbcrlf
  response.write "                                <select name=""sc_formactive"" id=""sc_formactive"">" & vbcrlf
  response.write "                                  <option value=""""></option>" & vbcrlf
  response.write "                                  <option value=""ON"""  & lcl_selected_active_on  & ">ON</option>" & vbcrlf
  response.write "                                  <option value=""OFF""" & lcl_selected_active_off & ">OFF</option>" & vbcrlf
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td>Order By:</td>" & vbcrlf
  response.write "                            <td>" & vbcrlf
  response.write "                                <select name=""sc_orderBy"" id=""sc_orderBy"">" & vbcrlf
  response.write "                                  <option value=""action_form_name""" & lcl_selected_orderby_action_form_name & ">Action Line Form Name</option>" & vbcrlf
  response.write "                                  <option value=""category"""         & lcl_selected_orderby_category         & ">Category</option>" & vbcrlf
  response.write "                                  <option value=""department"""       & lcl_selected_orderby_department       & ">Department</option>" & vbcrlf
  response.write "                                  <option value=""assignedto"""       & lcl_selected_orderby_assignedto       & ">Assigned To</option>" & vbcrlf
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
  response.write "                      </table>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <input type=""submit"" name=""searchButton"" id=""searchButton"" class=""button"" value=""Search"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "              </form>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf

  if blnCanManageActionAlerts then
  			session("RedirectPage") = "manage_action_forms.asp"
  			List_Forms session("orgid"), lcl_sc_formname, lcl_sc_formactive, lcl_sc_orderby
  else
     response.write "<div class=""orgadminboxf"">" & vbcrlf
     response.write "  <p>" & vbcrlf
     response.write "     <strong>Security Alert!</strong><br />You do not have permission to access the <strong>E-Gov Alert Manager</strong>.  " & vbcrlf
     response.write "     Please contact your E-Govlink administrator to inquire about gaining access to the <strong>E-Gov Alert Manager</strong>." & vbcrlf
     response.write "  </p>" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>

<!--#Include file="../admin_footer.asp"-->

<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
function List_Forms(iOrgID, iSC_FormName, iSC_FormActive, iSC_OrderBy)

  sWhereClause = ""

 'Build WHERE clause
  if iSC_FormName <> "" then
     sWhereClause = sWhereClause & " AND UPPER(action_form_name) LIKE ('%" & UCASE(replace(iSC_FormName,"'","''")) & "%') "
  end if

  if iSC_FormActive <> "" then
     if iSC_FormActive = "ON" then
        sWhereClause = sWhereClause & " AND action_form_enabled = 1 "
     else
        sWhereClause = sWhereClause & " AND action_form_enabled = 0 "
     end if
  end if

 'Build the order by
  if iSC_OrderBy = "category" then
     sSC_OrderBy = "form_category_name, action_form_name"
  elseif iSC_OrderBy = "department" then
     sSC_OrderBy = "deptname, action_form_name"
  elseif iSC_OrderBy = "assignedto" then
     sSC_OrderBy = "sAssignedName"
  else
     sSC_OrderBy = "action_form_name, form_category_name"
  end if

 'List Action Requests
  sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_forms_categories_view "
  sSQL = sSQL & " WHERE orgid=" & iOrgID

  if sWhereClause <> "" then
     sSQL = sSQL & sWhereClause
  end if

  sSQL = sSQL & " ORDER BY " & sSC_OrderBy

  set oRequests = Server.CreateObject("ADODB.Recordset")
  oRequests.Open sSQL, Application("DSN"), 3, 1

  if not oRequests.eof then
     if Len(Request("pagenum")) = 0 OR clng(Request("pagenum")) < 1 then
			   		'oRequests.AbsolutePage = 1
			   		lcl_AbsolutePage = 1
   					session("pageNum")     = 1
   		else
			   		'if clng(Request("pagenum")) <= oRequests.PageCount then
 					   		'oRequests.AbsolutePage = Request("pagenum")
			   		if clng(Request("pagenum")) <= lcl_PageCount then
 					   		lcl_AbsolutePage = Request("pagenum")
   	 						session("pageNum")     = Request("pagenum")
   					else
 		   					'oRequests.AbsolutePage = 1
 		   					lcl_AbsolutePage = 1
	 				   		session("pageNum")     = 1
        end if
    	end if
  else
   		'oRequests.AbsolutePage = 1
   		lcl_AbsolutePage = 1
					session("pageNum")     = 1
  end if
 
 'Display Record Statistics
  Dim abspage, pagecnt
  'abspage      = oRequests.AbsolutePage
  'pagecnt      = oRequests.PageCount
  abspage      = lcl_AbsolutePage
  pagecnt      = lcl_PageCount
	 sQueryString = replace(request.querystring,"pagenum","HFe301")  'Replace PAGENUM field with random field for navigation purposes

 'Display forward and backward navigation top
  response.write "<div class=""shadow"">" & vbcrlf
  response.write "<table cellspacing=""0"" cellpadding=""5"" class=""tablelist"" width=""100%"">" & vbcrlf
  response.write "  <tr class=""tablelist"">" & vbcrlf
  response.write "      <th>ID</th>" & vbcrlf
  response.write "      <th align=""left"">Action Line Form Name</th>" & vbcrlf
  response.write "      <th align=""left"">Category</th>" & vbcrlf
  response.write "      <th align=""left"">Department</th>" & vbcrlf
  response.write "      <th align=""left"" nowrap=""nowrap"">Assigned To</th>" & vbcrlf
  response.write "      <th>Form Active</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  bgcolor = "#eeeeee"
	
  do while not oRequests.eof
     bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")

  		'Determine if the form is available
     sEnabledColor = "red"
     sEnabledLabel = "OFF"
     blnEnabled    = 1

   		if oRequests("action_form_enabled") then
        sEnabledColor = "green"
        sEnabledLabel = "ON"
        blnEnabled    = 0
     end if
     sUserColor = "black"
     sUserEnabled = ""
     if oRequests("userdeleted") then
	     sUserColor = "red"
	     sUserEnabled = " - INACTIVE"
     end if
 				sEnabled = "<font style=""color:" & sEnabledColor & ";font-size:10px;"">" & sEnabledLabel & "</font>"

    	response.write "  <tr bgcolor=" & bgcolor & " onclick=""location.href='edit_form.asp?control=" & oRequests("action_form_id") & "';"" onMouseOver=""this.style.backgroundColor='#93bee1';this.style.cursor='hand';"" onMouseOut=""this.style.backgroundColor='';this.style.cursor='';"">" & vbcrlf
     response.write "      <td width=""25""><strong>" & oRequests("action_form_id") & "</strong></td>" & vbcrlf
     response.write "      <td><strong>" & oRequests("action_form_name") & "</strong></td>" & vbcrlf
     response.write "      <td>" & oRequests("form_category_name") & " </td>" & vbcrlf
     response.write "      <td>" & oRequests("DeptName") & " </td>" & vbcrlf
     response.write "      <td><strong style=""color:" & sUserColor & ";"">" & lcase(oRequests("sAssignedName")) & sUserEnabled & "</strong></td>" & vbcrlf
     response.write "      <td align=""center"">" & sEnabled & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
		
   		oRequests.MoveNext 
	 loop

  oRequests.close
  set oRequests = nothing

	 response.write "</table>" & vbcrlf
	 response.write "</div>" & vbcrlf

end function
%>
