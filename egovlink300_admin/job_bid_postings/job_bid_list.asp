<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="job_bid_global_functions.asp" //-->
<!-- #include file="../customreports/customreports_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: job_bid_list.asp
' AUTHOR:   David Boyer
' CREATED:  01/30/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/30/08  David Boyer - INITIAL VERSION
' 1.1  10/20/08  David Boyer - Added "# Bids Uploaded" to results list
' 1.2  05/21/09  David Boyer - Added "Click Counter" custom report.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("job_postings,bid_postings") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=hide

'Check the type of list and then check for the permission
 if UCASE(request("sc_list_type")) = "JOB" then
    if not UserHasPermission( Session("userid"), "job_postings" ) then
  	    response.redirect sLevel & "permissiondenied.asp"
    end if
 elseif UCASE(request("sc_list_type")) = "BID" then
    if not UserHasPermission( Session("userid"), "bid_postings" ) then
  	    response.redirect sLevel & "permissiondenied.asp"
    end if
 else
    if not UserHasPermission( Session("UserId"), "distribution lists" ) then
  	    response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

'Retrieve the search parameters
 lcl_sc_jobbid_id   = request("sc_jobbid_id")
 lcl_sc_title       = request("sc_title")
 lcl_sc_status_id   = request("sc_status_id")
 lcl_sc_active_flag = request("sc_active_flag")
 lcl_sc_list_type   = request("sc_list_type")
 lcl_sc_orderby     = request("sc_orderby")

'Set up the ORDER BY
 if lcl_sc_orderby <> "" then
    if lcl_sc_orderby = "title" then
       lcl_orderby = "title"
    elseif lcl_sc_orderby = "status_id" then
       lcl_orderby = "6"
    elseif lcl_sc_orderby = "active_flag" then
       lcl_orderby = "active_flag"
    elseif lcl_sc_orderby = "jobbid_id" then
       lcl_orderby = "jobbid_id"
    'elseif lcl_sc_orderby = "list_type" then
    'lcl_orderby = "UPPER(distributionlisttype)"
    end if
 else
    lcl_sc_orderby = "title"
    lcl_orderby    = "title"
 end if

'Set up the link parameters for the return url for the search criteria options
 lcl_return_url_parameters = ""
 if lcl_sc_title <> "" then
    lcl_return_url_parameters = "sc_title=" & lcl_sc_title
 end if

 if lcl_sc_jobbid_id <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_jobbid_id=" & lcl_sc_jobbid_id
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_jobbid_id=" & lcl_sc_jobbid_id
    end if
 end if

 if lcl_sc_status_id <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_status_id=" & lcl_sc_status_id
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_status_id=" & lcl_sc_status_id
    end if
 end if

 if lcl_sc_active_flag <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_active_flag=" & lcl_sc_active_flag
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_active_flag=" & lcl_sc_active_flag
    end if
 end if

 if lcl_sc_list_type <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_list_type=" & lcl_sc_list_type
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_list_type=" & lcl_sc_list_type
    end if
 end if

 if lcl_sc_orderby <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_orderby=" & lcl_sc_orderby
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_orderby=" & lcl_sc_orderby
    end if
 end if

 if lcl_return_url_parameters <> "" then
    lcl_return_url_parameters = "&" & REPLACE(lcl_return_url_parameters,"%","<<PER>>")
 end if

'Convert the % in the search criteria
 lcl_sc_title = REPLACE(request("sc_title"),"<<PER>>","%")

'Determine the list type and what label should be displayed on the page
 if UCASE(lcl_sc_list_type) = "JOB" then
    lcl_list_label        = "Job Postings"
    lcl_new_list_label    = "Category"
    lcl_new_posting_label = "New Job Posting"
 elseif UCASE(lcl_sc_list_type) = "BID" then
    lcl_list_label        = "Bid Postings"
    lcl_new_list_label    = "Category/Sub-Category"
    lcl_new_posting_label = "New Bid Posting"
 end if

'Check for org features
 'lcl_orghasfeature_customreports             = orghasfeature("customreports")
 lcl_orghasfeature_customreports_clickcounts = orghasfeature("customreports_clickcounts")
 lcl_orghasfeature_clickcounter_postings     = orghasfeature("clickcounter_postings")

'Check for user permissions
 lcl_userhaspermission_create_job_postings       = userhaspermission(session("userid"),"create_job_postings")
 lcl_userhaspermission_create_bid_postings       = userhaspermission(session("userid"),"create_bid_postings")
 lcl_userhaspermission_edit_job_postings         = userhaspermission(session("userid"),"edit_job_postings")
 lcl_userhaspermission_edit_bid_postings         = userhaspermission(session("userid"),"edit_bid_postings")
 'lcl_userhaspermission_customreports             = userhaspermission(session("userid"),"customreports")
 lcl_userhaspermission_customreports_clickcounts = userhaspermission(session("userid"),"customreports_clickcounts")

'Check for permission to create a new job/bid posting
 lcl_display_add = "Y"

 if UCASE(lcl_sc_list_type) = "JOB" then
    if not lcl_userhaspermission_create_job_postings then
  	    lcl_display_add = "N"
    end if
 elseif UCASE(lcl_sc_list_type) = "BID" then
    if not lcl_userhaspermission_create_bid_postings then
  	    lcl_display_add = "N"
    end if
 end if

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Build BODY onload
 lcl_onload = lcl_onload & "document.search_sort_form.sc_title.focus();"
%>
<html>
<head>
	<title>E-Gov Administration Console {<%=lcl_list_label%>}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	
<script language="javascript" src="tablesort.js"></script>
<script language="javascript" src="../scripts/modules.js"></script>
<script language="javascript">
<!--
//function deleteconfirm(ID, sName) {
function deleteconfirm(ID) {
  lcl_name = document.getElementById("posting_id_"+ID).innerHTML;

// 	if(confirm('Do you wish to delete \'' + sName + '\'?')) {
 	if(confirm('Do you wish to delete \'' + lcl_name + '\'?')) {
     lcl_redirect_url = "&sc_title=<%=lcl_sc_title%>&sc_status_id=<%=lcl_sc_status_id%>&sc_list_type=<%=lcl_sc_list_type%>&sc_orderby=<%=lcl_sc_orderby%>";

 				window.location="job_bid_action.asp?posting_id="+ID+"&cmd=D"+lcl_redirect_url;
		}
}

function openWin2(url, name) {
		popupWin = window.open(url, name,"resizable,width=800,height=400");
}

function openCustomReports(p_report,p_posting_id) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("../customreports/customreports.asp?cr='+p_report+'&posting_id='+p_posting_id+'", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
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
<body onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	 <div id="centercontent">

  <p><font size="+1"><strong><%=Session("sOrgName")%>&nbsp;<%=lcl_list_label%></strong></font></p>

<table border="0" cellspacing="0" cellpadding="5">
  <tr>
      <td colspan="2">
          <fieldset>
            <legend><strong>Search/Sorting Option(s)&nbsp;</strong></legend><br />
          <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <form name="search_sort_form" value="job_bid_list.asp">
              <input type="<%=lcl_hidden%>" name="sc_list_type" value="<%=lcl_sc_list_type%>" size="15" maxlength="100" />
            <tr valign="top">
                <td width="60%">
                    <table border="0" cellspacing="0" cellpadding="1">
                      <tr>
                          <td><%=lcl_sc_list_type%> ID:</td>
                          <td width="75%"><input type="text" name="sc_jobbid_id" value="<%=lcl_sc_jobbid_id%>" size="30" maxlength="500" /></td>
                      </tr>
                      <tr>
                          <td>Title:</td>
                          <td width="75%"><input type="text" name="sc_title" value="<%=lcl_sc_title%>" size="30" maxlength="512" /></td>
                      </tr>
                      <tr>
                          <td>Status:</td>
                          <td width="75%">
                              <select name="sc_status_id">
                                <option value=""></option>
                              <%
                                'Retrieve all statuses on existing job/bid postings
                                 sSQLs = "SELECT distinct jb.status_id, s.status_name, s.status_order "
                                 sSQLs = sSQLs & " FROM egov_jobs_bids jb, egov_statuses s "
                                 sSQLs = sSQLs & " WHERE jb.status_id = s.status_id "
                                 sSQLs = sSQLs & " AND jb.orgid = " & session("orgid")
                                 sSQLs = sSQLs & " AND jb.posting_type = '" & lcl_sc_list_type & "'"
                                 sSQLs = sSQLs & " AND (s.status_name <> '' OR s.status_name IS NOT NULL) "
                                 sSQLs = sSQLs & " ORDER BY s.status_order "

                                 set oJBStatus = Server.CreateObject("ADODB.Recordset")
                                	oJBStatus.Open sSQLs, Application("DSN"), 0, 1

                                 if not oJBStatus.eof then
                                    while not oJBStatus.eof
                                       if lcl_sc_status_id <> "" then
                                          if clng(lcl_sc_status_id) = clng(oJBStatus("status_id")) then
                                             lcl_selected = " selected"
                                          else
                                             lcl_selected = ""
                                          end if
                                       else
                                          lcl_selected = ""
                                       end if

                                       response.write "  <option value=""" & oJBStatus("status_id") & """" & lcl_selected & ">" & oJBStatus("status_name") & "</option>" & vbcrlf

                                       oJBStatus.movenext
                                    wend
                                 end if

                                 oJBStatus.close
                                 set oJBStatus = nothing
                              %>
                              </select>
                          </td>
                      </tr>
                    </table>
                </td>
                <td>
                    <table border="0" cellspacing="0" cellpadding="2">
                      <tr>
                          <td>Active:</td>
                          <td>
                              <select name="sc_active_flag">
                              <%
                                 if lcl_sc_active_flag = "Y" then
                                    lcl_selected_yes = "selected"
                                    lcl_selected_no  = ""
                                 elseif lcl_sc_active_flag = "N" then
                                    lcl_selected_yes = ""
                                    lcl_selected_no  = "selected"
                                 else
                                    lcl_selected_yes = ""
                                    lcl_selected_no  = ""
                                 end if
                              %>
                                <option value=""></option>
                                <option value="Y" <%=lcl_selected_yes%>>Yes</option>
                                <option value="N" <%=lcl_selected_no%>>No</option>
                              </select>
                          </td>
                      </tr>
                      <tr>
                          <td>Sort:</td>
                          <td>
                              <select name="sc_orderby">
                              <%
                                'select order by value
                                 if lcl_sc_orderby = "Title" then
                                    lcl_title_selected              = " selected"
                                    lcl_jobbid_id_selected          = ""
                                    lcl_status_id_selected          = ""
                                    lcl_active_flag_selected        = ""
                                 elseif lcl_sc_orderby = "jobbid_id" then
                                    lcl_title_selected              = ""
                                    lcl_jobbid_id_selected          = " selected"
                                    lcl_status_id_selected          = ""
                                    lcl_active_flag_selected        = ""
                                 elseif lcl_sc_orderby = "status_id" then
                                    lcl_title_selected              = ""
                                    lcl_jobbid_id_selected          = ""
                                    lcl_status_id_selected          = " selected"
                                    lcl_active_flag_selected        = ""
                                 elseif lcl_sc_orderby = "active_flag" then
                                    lcl_title_selected              = ""
                                    lcl_jobbid_id_selected          = ""
                                    lcl_status_id_selected          = ""
                                    lcl_active_flag_selected        = " selected"
                                 else
                                    lcl_title_selected              = " selected"
                                    lcl_jobbid_id_selected          = ""
                                    lcl_status_id_selected          = ""
                                    lcl_active_flag_selected        = ""
                                 end if
                              %>
                                <option value="title"<%=lcl_title_selected%>>Title</option>
                                <option value="jobbid_id"<%=lcl_jobbid_id_selected%>><%=lcl_sc_list_type%> ID</option>
                                <option value="status_id"<%=lcl_status_id_selected%>>Status</option>
                                <option value="active_flag"<%=lcl_active_flag_selected%>>Active</option>
                              </select>
                          </td>
                      </tr>
                    </table>
                </td>
            </tr>
            <tr><td colspan="2"><input type="submit" value="SEARCH" method="post" action="job_bid_list.asp" class="button" /></td></tr>
            </form>
          </table>
          </fieldset>
          <p>
      </td>
  </tr>
  <tr>
      <td>
          <% displayButtons lcl_display_add, lcl_new_posting_label, lcl_return_url_parameters %>
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
  <tr>
      <td colspan="2">
<%
'Retrieve all of the job/bids for this sub-category
 sSQL = "SELECT jb.posting_id, jb.jobbid_id, jb.title, jb.active_flag, "
 sSQL = sSQL & " (select s.status_name from egov_statuses s where s.status_id = jb.status_id) AS status_name, "
 sSQL = sSQL & " (select count(ub.userbidid) from egov_jobs_bids_userbids ub where ub.posting_id = jb.posting_id) AS total_uploaded, "
 sSQL = sSQL & " (select count(cc.postings_clickid) from egov_clickcounter_postings cc where cc.posting_id = jb.posting_id) AS times_clicked "
 sSQL = sSQL & " FROM egov_jobs_bids jb "
 sSQL = sSQL & " WHERE jb.posting_type = '" & lcl_sc_list_type & "'"
 sSQL = sSQL & " AND jb.orgid = " & session("orgid")

 if lcl_sc_title <> "" then
    sSQL = sSQL & " AND UPPER(jb.title) LIKE ('%" & UCASE(lcl_sc_title) & "%') "
 end if

 if lcl_sc_jobbid_id <> "" then
    sSQL = sSQL & " AND UPPER(jb.jobbid_id) LIKE ('%" & UCASE(lcl_sc_jobbid_id) & "%') "
 end if

 if lcl_sc_status_id <> "" then
    sSQL = sSQL & " AND jb.status_id = " & lcl_sc_status_id
 end if

 if lcl_sc_active_flag <> "" then
    sSQL = sSQL & " AND jb.active_flag = '" & lcl_sc_active_flag & "'"
 end if

 sSQL = sSQL & " ORDER BY " & lcl_orderby

	set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	if not oList.eof then

    response.write "<div class=""shadow"">" & vbcrlf
    response.write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" class=""tableadmin"">" & vbcrlf
    response.write "  <tr align=""left"" valign=""bottom"">" & vbcrlf
    response.write "      <th>" & lcl_sc_list_type & " ID</th>" & vbcrlf
    response.write "      <th>Title</th>" & vbcrlf
    response.write "      <th>Status</th>" & vbcrlf

    if lcl_sc_list_type = "BID" then
       response.write "      <th align=""center""># Bids<br />Uploaded</th>" & vbcrlf
    end if

    response.write "      <th align=""center"">Active</th>" & vbcrlf

   'Show the "Click Counts" report column if the org and user have the proper permissions
   'Also, since we are checking the permissions, if they are valid for the user and org set the session variable for the report type
    'and lcl_orghasfeature_customreports
    'and lcl_userhaspermission_customreports
    if  lcl_orghasfeature_clickcounter_postings _
    and lcl_orghasfeature_customreports_clickcounts _
    and lcl_userhaspermission_customreports_clickcounts then
        session("CR_CLICKCOUNTS") = "POSTINGS"

        response.write "      <th align=""center"">Click Counts<br />Report</th>" & vbcrlf
    end if

    response.write "      <th>&nbsp;</th>" & vbcrlf
    response.write "  </tr>" & vbcrlf

    lcl_bgcolor = "#ffffff"
    iRowCount   = 0

    do while not oList.eof
       lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")
       iRowCount   = iRowCount + 1

       if oList("active_flag") = "Y" then
          lcl_active_flag = oList("active_flag")
       else
          lcl_active_flag = ""
       end if

      'Check for permission to edit job/bid posting
       lcl_display_edit = "Y"

       if UCASE(lcl_sc_list_type) = "JOB" then
          if not lcl_userhaspermission_edit_job_postings then
  	          lcl_display_edit = "N"
          end if
       elseif UCASE(lcl_sc_list_type) = "BID" then
          if not lcl_userhaspermission_edit_bid_postings then
  	          lcl_display_edit = "N"
          end if
       end if

       if lcl_display_edit = "Y" then
          lcl_edit_url = " onclick=""location.href='job_bid_maint.asp?posting_id=" & oList("posting_id") & lcl_return_url_parameters & "'"""
       else
          lcl_edit_url = ""
       end if

       response.write "<tr bgcolor=""" & lcl_bgcolor & """ id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
       response.write "    <td" & lcl_edit_url & "><span id=""posting_id_" & oList("posting_id") & """>" & oList("jobbid_id") & "</span></td>" & vbcrlf
       response.write "    <td" & lcl_edit_url & ">" & trim(oList("title"))  & "</td>" & vbcrlf
       response.write "    <td" & lcl_edit_url & ">" & oList("status_name")  & "</td>" & vbcrlf

       if lcl_sc_list_type = "BID" then
          response.write "    <td" & lcl_edit_url & " align=""center"">" & oList("total_uploaded") & "</td>" & vbcrlf
       end if

       response.write "    <td" & lcl_edit_url & " align=""center"">" & lcl_active_flag & "</td>" & vbcrlf

      'Show the "Click Counts" report column if the org and user have the proper permissions
       'and lcl_orghasfeature_customreports
       'and lcl_userhaspermission_customreports
       if  lcl_orghasfeature_clickcounter_postings _
       and lcl_orghasfeature_customreports_clickcounts _
       and lcl_userhaspermission_customreports_clickcounts then

           response.write "    <td align=""center"">" & vbcrlf

          'Only show the buttons if the posting has been clicked on.
           if oList("times_clicked") > 0 then
              response.write "<input type=""button"" name=""clickCountReportButton"" id=""clickCountReportButton"" value=""View"" class=""button"" onclick=""openCustomReports('CLICKCOUNTS','" & oList("posting_id") & "')"" />" & vbcrlf
           else
              response.write "&nbsp;" & vbcrlf
           end if

           response.write "</td>" & vbcrlf
       end if

       response.write "    <td align=""center""><input type=""button"" name=""delete"" id=""delete"" value=""Delete"" class=""button"" onclick=""deleteconfirm(" & oList("posting_id") & ")"" /></td>" & vbcrlf
       response.write "</tr>" & vbcrlf

    			oList.movenext
    loop

    response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf

		  oList.close
  		Set oList = Nothing 
 else
  		response.write "<font color=""#ff0000""><strong>No " & lcl_list_label & " currenty exist.</strong></font>" & vbcrlf
 end if
%>
      </td>
  </tr>
  <tr>
      <td colspan="2">
          <% displayButtons lcl_display_add, lcl_new_posting_label, lcl_return_url_parameters %>
      </td>
  </tr>
</table>
 	</div>
</div>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displayButtons(iDisplayAdd, iNewPostingLabel, iReturnParams)

  if iDisplayAdd = "Y" then
     response.write "<input type=""button"" name=""newposting"" id=""newposting"" value=""" & iNewPostingLabel & """ class=""button"" onclick=""location.href='job_bid_maint.asp?posting_id=0" & iReturnParams & "';"" />" & vbcrlf
  else
     response.write "&nbsp;"
  end if

end sub
%>
