<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: job_bid_userbids_list.asp
' AUTHOR:   David Boyer
' CREATED:  10/15/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Shows all of the "user bids" uploaded by citizens submitting their bids on a (job/bid) posting(s)
'
' MODIFICATION HISTORY
' 1.0  10/15/08  David Boyer - Initial Version
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
' if UCASE(request("sc_list_type")) = "JOB" then
'    if not UserHasPermission( Session("userid"), "job_postings" ) then
'  	    response.redirect sLevel & "permissiondenied.asp"
'    end if
 if UCASE(request("sc_list_type")) = "BID" then
    if not UserHasPermission( session("userid"), "view_uploaded_userbids" ) then
  	    response.redirect sLevel & "permissiondenied.asp"
    end if
 else
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the search parameters
 lcl_sc_list_type        = request("sc_list_type")
 lcl_sc_uploadid         = request("sc_uploadid")
 lcl_sc_userlabel        = request("sc_userlabel")
 lcl_sc_userid           = request("sc_userid")
 lcl_sc_jobbid_id        = request("sc_jobbid_id")
 lcl_sc_title            = request("sc_title")
 lcl_sc_userbusinessname = request("sc_userbusinessname")
 lcl_sc_enddate_to       = request("sc_enddate_to")
 lcl_sc_enddate_from     = request("sc_enddate_from")
 lcl_sc_orderby          = request("sc_orderby")

'Set up the ORDER BY
 if lcl_sc_orderby <> "" then
    if lcl_sc_orderby = "jobbid_id" then
       lcl_orderby = "jb.jobbid_id"
    elseif lcl_sc_orderby = "title" then
       lcl_orderby = "jb.title"
    elseif lcl_sc_orderby = "uploadid" then
       lcl_orderby = "ub.uploadid"
    elseif lcl_sc_orderby = "enddate" then
       lcl_orderby = "17, jb.end_date"
    end if
 else
    lcl_sc_orderby = "jobbid_id"
    lcl_orderby    = "jb.jobbid_id"
 end if

'Set up the link parameters for the return url for the search criteria options
 lcl_return_url_parameters = ""

 if lcl_sc_list_type <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_list_type=" & lcl_sc_list_type
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_list_type=" & lcl_sc_list_type
    end if
 end if

 if lcl_sc_uploadid <> "" then
     if lcl_return_url_parameters = "" then
        lcl_return_url_parameters = "sc_uploadid=" & lcl_sc_uploadid
     else
        lcl_return_url_parameters = lcl_return_url_parameters & "&sc_uploadid=" & lcl_sc_uploadid
     end if
 end if

 if lcl_sc_userlabel <> "" then
     if lcl_return_url_parameters = "" then
        lcl_return_url_parameters = "sc_userlabel=" & lcl_sc_userlabel
     else
        lcl_return_url_parameters = lcl_return_url_parameters & "&sc_userlabel=" & lcl_sc_userlabel
     end if
 end if

 if lcl_sc_userid <> "" then
     if lcl_return_url_parameters = "" then
        lcl_return_url_parameters = "sc_userid=" & lcl_sc_userid
     else
        lcl_return_url_parameters = lcl_return_url_parameters & "&sc_userid=" & lcl_sc_userid
     end if
 end if

 if lcl_sc_title <> "" then
     if lcl_return_url_parameters = "" then
        lcl_return_url_parameters = "sc_title=" & lcl_sc_title
     else
        lcl_return_url_parameters = lcl_return_url_parameters & "&sc_title=" & lcl_sc_title
     end if
 end if

 if lcl_sc_jobbid_id <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_jobbid_id=" & lcl_sc_jobbid_id
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_jobbid_id=" & lcl_sc_jobbid_id
    end if
 end if

 if lcl_sc_userbusinessname <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_userbusinessname=" & lcl_sc_userbusinessname
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_userbusinessname=" & lcl_sc_userbusinessname
    end if
 end if

 if lcl_sc_enddate_to <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_enddate_to=" & lcl_sc_enddate_to
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_enddate_to=" & lcl_sc_enddate_to
    end if
 end if

 if lcl_sc_enddate_from <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_enddate_from=" & lcl_sc_enddate_from
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_enddate_from=" & lcl_sc_enddate_from
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

'Determine the list type and what label should be displayed on the page
 if UCASE(lcl_sc_list_type) = "JOB" then
    lcl_list_label        = "Job Postings: User Bids"
    lcl_new_list_label    = "Category"
    lcl_new_posting_label = "New Job Posting: User Bid"
 elseif UCASE(lcl_sc_list_type) = "BID" then
    lcl_list_label        = "Bid Postings: User Bids"
    lcl_new_list_label    = "Category/Sub-Category"
    lcl_new_posting_label = "New Bid Posting: User Bid"
 end if

'Set the width for all "container" tables
 lcl_table_width = "1000px"

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone()

 'lcl_current_date = date() & " " & time()

'Set the "FirstViewedBy" Info
 if request("viewbid") = "Y" AND request("userbidid") <> "" then
    setFirstViewedByInfo request("userbidid")
 end if

'Get the local date/time for the org
 lcl_local_datetime = ConvertDateTimetoTimeZone
%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	
<script language="javascript" src="tablesort.js"></script>
<script language="javascript" src="../scripts/modules.js"></script>
<script language="javascript">
<!--
function openWin2(url, name) {
		popupWin = window.open(url, name,"resizable,width=800,height=400");
}

function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1&updateform=search_sort_form&updatefield=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function viewBid(iRowID,iUserBidID) {
  //lcl_viewbid_url = '<% 'session("egovclientwebsiteurl")%>/admin'+document.getElementById("filepath"+iRowID).value;
  //lcl_viewbid_url = '<%=Application("CommunityLink_DocUrl")%>/public_documents300/<%=session("sitename")%>'+document.getElementById("filepath"+iRowID).value;
  lcl_viewbid_url = document.getElementById("filepath"+iRowID).value;

  w = (screen.width - 800)/2;
  h = (screen.height - 600)/2;
  popupWin = window.open(lcl_viewbid_url, "_viewbid","resizable,width=800,height=600,left=" + w + ",top=" + h);

<%
  if lcl_return_url_parameters <> "" then
     if left(lcl_return_url_parameters,1) = "&" then
        lcl_return_url_parameters = "?" & mid(lcl_return_url_parameters,2)
     end if
  end if
%>
  location.href = 'job_bid_userbids_list.asp<%=lcl_return_url_parameters%>&userbidid='+iUserBidID+'&viewbid=Y';
}
//-->
</script>
</head>
<body onload="javascript:document.search_sort_form.sc_title.focus()">
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
          <table border="0" cellspacing="0" cellpadding="0" width="<%=lcl_table_width%>">
            <form name="search_sort_form" value="job_bid_userbids_list.asp">
              <input type="<%=lcl_hidden%>" name="sc_list_type" value="<%=lcl_sc_list_type%>" size="15" maxlength="100" />
            <tr valign="top">
                <td width="0%">
                    <table border="0" cellspacing="0" cellpadding="1">
                      <tr>
                          <td>Bid Upload ID:</td>
                          <td width="75%"><input type="text" name="sc_uploadid" value="<%=lcl_sc_uploadid%>" size="40" maxlength="500" /></td>
                      </tr>
                      <tr>
                          <td>Label:</td>
                          <td width="75%"><input type="text" name="sc_userlabel" value="<%=lcl_sc_userlabel%>" size="40" maxlength="500" /></td>
                      </tr>
                      <tr>
                          <td>Title:</td>
                          <td width="75%"><input type="text" name="sc_title" value="<%=lcl_sc_title%>" size="40" maxlength="512" /></td>
                      </tr>
                      <tr>
                          <td colspan="2">
                              <table border="0" cellspacing="0" cellpadding="2" width="100%" style="margin-top:2px">
                                <tr>
                                    <td>&nbsp;End Date&nbsp;</td>
                                    <td>
                                        To: <input type="text" name="sc_enddate_to" value="<%=lcl_sc_enddate_to%>" size="10" maxlength="10" />
                                            <img src="../images/calendar.gif" border="0" style="cursor:hand" onclick="void doCalendar('sc_enddate_to');" />
                                    </td>
                                    <td>
                                        From: <input type="text" name="sc_enddate_from" value="<%=lcl_sc_enddate_from%>" size="10" maxlength="10" />
                                              <img src="../images/calendar.gif" border="0" style="cursor:hand" onclick="void doCalendar('sc_enddate_from');" />
                                    </td>
                                </tr>
                              </table>
                          </td>
                      </tr>
                    </table>
                </td>
                <td width="50%">
                    <table border="0" cellspacing="0" cellpadding="1">
                      <tr>
                          <td><%=lcl_sc_list_type%> ID:</td>
                          <td width="75%"><input type="text" name="sc_jobbid_id" value="<%=lcl_sc_jobbid_id%>" size="30" maxlength="500" /></td>
                      </tr>
                      <tr>
                          <td>Submitted By:</td>
                          <td width="75%">
                              <select name="sc_userid" id="sc_userid">
                                <option value=""></option>
                              <%
                               'Retrieve all of the users who have submitted bids
                                sSQLu = "SELECT userid, userlname, userfname "
                                sSQLu = sSQLu & " FROM egov_users u "
                                sSQLu = sSQLu & " WHERE userid IN (select distinct ub.userid "
                                sSQLu = sSQLu &                " from egov_jobs_bids_userbids ub "
                                sSQLu = sSQLu &                " where orgid = " & session("orgid")
                                sSQLu = sSQLu &                " and UPPER(ub.posting_type) = '" & UCASE(lcl_sc_list_type) & "') "
                                sSQLu = sSQLu & " ORDER BY userlname, userfname, userid "

                               	set oSCUser = Server.CreateObject("ADODB.Recordset")
                               	oSCUser.Open sSQLu, Application("DSN"), 3, 1

                                if not oSCUser.eof then
                                   while not oSCUser.eof
                                      if cstr(lcl_sc_userid) = cstr(oSCUser("userid")) then
                                         lcl_selected = " selected=""selected"""
                                      else
                                         lcl_selected = ""
                                      end if

                                      response.write "<option value=""" & oSCUser("userid") & """" & lcl_selected & ">" & oSCUser("userfname") & " " & oSCUser("userlname") & "</option>" & vbcrlf
                                      oSCUser.movenext
                                   wend
                                end if

                                oSCUser.close
                                set oSCUser = nothing
                              %>
                              </select>
                          </td>
                       </tr>
                      <tr>
                          <td>Business:</td>
                          <td width="75%"><input type="text" name="sc_userbusinessname" value="<%=lcl_sc_userbusinessname%>" size="30" maxlength="255" /></td>
                      </tr>
                      <tr>
                          <td>Order By:</td>
                          <td width="75%">
                              <select name="sc_orderby">
                              <%
                                 lcl_jobbid_id_selected  = ""
                                 lcl_title_selected      = ""
                                 lcl_uploadid_selected   = ""
                                 lcl_enddate_selected    = ""

                                'select order by value
                                 if lcl_sc_orderby = "jobbid_id" then
                                    lcl_jobbid_id_selected = "selected"
                                 elseif lcl_sc_orderby = "title" then
                                    lcl_title_selected     = "selected"
                                 elseif lcl_sc_orderby = "uploadid" then
                                    lcl_uploadid_selected  = "selected"
                                 elseif lcl_sc_orderby = "enddate" then
                                    lcl_enddate_selected   = "selected"
                                 else
                                    lcl_jobbid_id_selected = "selected"
                                 end if
                              %>
                                <option value="jobbid_id" <%=lcl_jobbid_id_selected%>><%=lcl_sc_list_type%> ID</option>
                                <option value="title" <%=lcl_title_selected%>>Title</option>
                                <option value="uploadid" <%=lcl_uploadid_selected%>>Bid Upload ID</option>
                                <option value="enddate" <%=lcl_enddate_selected%>>End Date</option>
                              </select>
                          </td>
                      </tr>
                    </table>
                </td>
            </tr>
            <tr><td colspan="2"><input type="submit" value="SEARCH" class="button" method="post" action="job_bid_userbids_list.asp" /></td></tr>
            </form>
          </table>
          </fieldset>
          <p>
      </td>
  </tr>
  <tr>
      <td colspan="2">
<%
'Set up the search criteria field limitations
 lcl_where_clause = ""

 if lcl_sc_uploadid <> "" then
    lcl_where_clause = lcl_where_clause & " AND UPPER(ub.uploadid) LIKE ('%" & UCASE(lcl_sc_uploadid) & "%') "
 end if

 if lcl_sc_userlabel <> "" then
    lcl_where_clause = lcl_where_clause & " AND UPPER(ub.userlabel) LIKE ('%" & UCASE(lcl_sc_userlabel) & "%') "
 end if

 if lcl_sc_jobbid_id <> "" then
    lcl_where_clause = lcl_where_clause & " AND UPPER(jb.jobbid_id) LIKE ('%" & UCASE(lcl_sc_jobbid_id) & "%') "
 end if

 if lcl_sc_title <> "" then
    lcl_where_clause = lcl_where_clause & " AND UPPER(jb.title) LIKE ('%" & UCASE(lcl_sc_title) & "%') "
 end if

 if lcl_sc_userid <> "" then
    lcl_where_clause = lcl_where_clause & " AND ub.userid = " & lcl_sc_userid
 end if

 if lcl_sc_userbusinessname <> "" then
    lcl_where_clause = lcl_where_clause & " AND UPPER(u.userbusinessname) LIKE ('%" & UCASE(lcl_sc_userbusinessname) & "%') "
 end if

 if lcl_sc_enddate_to <> "" then
    lcl_where_clause = lcl_where_clause & " AND (cast(CONVERT(varchar(10), jb.end_date, 101) AS datetime) >= '" & lcl_sc_enddate_to & "' OR jb.end_date = '1/1/1900') "
 end if

 if lcl_sc_enddate_from <> "" then
    lcl_where_clause = lcl_where_clause & " AND (cast(CONVERT(varchar(10), jb.end_date, 101) AS datetime) <= '" & lcl_sc_enddate_from & "' OR jb.end_date = '1/1/1900') "
 end if

'Retrieve all of the "user bids"
 sSQL = "SELECT ub.userbidid, ub.posting_id, ub.posting_type, ub.userid, ub.orgid, ub.submitdate, ub.uploadid, ub.filelocation, "
 sSQL = sSQL & " ub.filename, jb.jobbid_id, jb.title, u.userfname + ' ' + u.userlname AS user_fullname, "
 sSQL = sSQL & " u.userbusinessname, jb.end_date AS end_date, '1' AS orderby_enddate, "
 sSQL = sSQL & " (select status_name from egov_statuses where status_id = jb.status_id) AS posting_status, "
 sSQL = sSQL & " isnull(ub.userbid_firstviewedby,0) AS firstviewedby, "
 sSQL = sSQL & " (select FirstName + ' ' + LastName from users where userid = isnull(ub.userbid_firstviewedby,0)) AS firstviewedby_name, "
 sSQL = sSQL & " isnull(ub.userbid_firstvieweddate,'') AS firstvieweddate, ub.userlabel "
 sSQL = sSQL & " FROM egov_jobs_bids_userbids ub, egov_jobs_bids jb, egov_users u "
 sSQL = sSQL & " WHERE ub.posting_id = jb.posting_id "
 sSQL = sSQL & " AND ub.userid = u.userid "
 sSQL = sSQL & " AND ub.orgid = " & session("orgid")
 sSQL = sSQL & " AND UPPER(ub.posting_type) = '" & UCASE(lcl_sc_list_type) & "' "
 sSQL = sSQL & " AND jb.end_date <> '1/1/1900' "
 sSQL = sSQL & lcl_where_clause
'------------------------------------------------------------------------------
 sSQL = sSQL & " UNION ALL "
'------------------------------------------------------------------------------
 sSQL = sSQL & " SELECT ub.userbidid, ub.posting_id, ub.posting_type, ub.userid, ub.orgid, ub.submitdate, ub.uploadid, ub.filelocation, "
 sSQL = sSQL & " ub.filename, jb.jobbid_id, jb.title, u.userfname + ' ' + u.userlname AS user_fullname, "
 sSQL = sSQL & " u.userbusinessname, '' AS end_date, '2' AS orderby_enddate, "
 sSQL = sSQL & " (select status_name from egov_statuses where status_id = jb.status_id) AS posting_status, "
 sSQL = sSQL & " isnull(ub.userbid_firstviewedby,0) AS firstviewedby, "
 sSQL = sSQL & " (select FirstName + ' ' + LastName from users where userid = isnull(ub.userbid_firstviewedby,0)) AS firstviewedby_name, "
 sSQL = sSQL & " isnull(ub.userbid_firstvieweddate,'') AS firstvieweddate, ub.userlabel "
 sSQL = sSQL & " FROM egov_jobs_bids_userbids ub, egov_jobs_bids jb, egov_users u "
 sSQL = sSQL & " WHERE ub.posting_id = jb.posting_id "
 sSQL = sSQL & " AND ub.userid = u.userid "
 sSQL = sSQL & " AND ub.orgid = " & session("orgid")
 sSQL = sSQL & " AND UPPER(ub.posting_type) = '" & UCASE(lcl_sc_list_type) & "' "
 sSQL = sSQL & " AND jb.end_date = '1/1/1900' "
 sSQL = sSQL & lcl_where_clause

 sSQL = sSQL & "ORDER BY " & lcl_orderby

	set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 3, 1

	if not oList.eof then
%>
          <div class="shadow" style="width:<%=lcl_table_width%>">
          <table border="0" cellpadding="5" cellspacing="0" class="tablelist" style="width:<%=lcl_table_width%>">
            <tr align="left">
                <th><%=lcl_sc_list_type%> ID</th>
                <th>Title</th>
                <th>Status</th>
                <th>End Date</th>
                <th>Submitted By</th>
                <th>Business</th>
                <th>Bid Upload ID</th>
                <th>&nbsp;</th>
                <th align="center">Initially<br />Viewed By</th>
            </tr>
<%
    lcl_bgcolor = "#ffffff"
    iRowCount   = 0
    while not oList.eof
       lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")
       iRowCount   = iRowCount + 1

      'Setup end_date
       if oList("end_date") <> "" then
          if CDate(oList("end_date")) = CDate("1/1/1900") then
             lcl_end_date = ""
          else
             lcl_end_date = oList("end_date")
          end if
       else
          lcl_end_date = ""
       end if

      'Setup firstvieweddate
       if oList("firstvieweddate") <> "" then
          if CDate(oList("firstvieweddate")) = CDate("1/1/1900") then
             lcl_firstviewedby_date = ""
          else
             lcl_firstviewedby_date = oList("firstvieweddate")
          end if
       else
          lcl_firstviewedby_date = ""
       end if

      'Check for permission to edit job/bid posting
       lcl_display_edit = "Y"

       'lcl_onclick = lcl_edit_url
       lcl_onclick = ""

       response.write "<tr bgcolor=""" & lcl_bgcolor & """ id=""" & iRowCount & """ valign=""top"" onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
       response.write "    <td" & lcl_onclick & ">" & oList("jobbid_id")             & "</td>" & vbcrlf
       response.write "    <td" & lcl_onclick & ">" & oList("title")                 & "</td>" & vbcrlf
       response.write "    <td" & lcl_onclick & ">" & oList("posting_status")        & "</td>" & vbcrlf
       response.write "    <td" & lcl_onclick & " nowrap=""nowrap"">" & lcl_end_date & "</td>" & vbcrlf
       response.write "    <td" & lcl_onclick & ">" & oList("user_fullname")         & "</td>" & vbcrlf
       response.write "    <td" & lcl_onclick & ">" & oList("userbusinessname")      & "</td>" & vbcrlf
       response.write "    <td" & lcl_onclick & ">[" & oList("uploadid") & "]<br /><span style=""color:#800000"">" & oList("userlabel") & "</span></td>" & vbcrlf
       response.write "    <td align=""center"">" & vbcrlf

      'Determine if the current date is > then end_date of the posting
      'If "yes" then show the "View Bid" button
       if lcl_end_date <> "" then
          if datediff("s",lcl_end_date,lcl_local_datetime) > 0 then

            'Build the filepath
             lcl_filelocation = oList("filelocation")
             lcl_filelocation = replace(lcl_filelocation,"\custom\pub\","")

             lcl_file_url = Application("userbids_upload_directory")
             lcl_file_url = lcl_file_url & "/public_documents300/"
             lcl_file_url = lcl_file_url & lcl_filelocation
             lcl_file_url = lcl_file_url & oList("filename")
             lcl_file_url = replace(lcl_file_url,"\","/")

             response.write "        <input type=""button"" name=""viewbid" & iRowCount & """ id=""viewbid" & iRowCount & """ value=""View Bid"" class=""button"" onclick=""viewBid(" & iRowCount & "," & oList("userbidid") & ")"" />" & vbcrlf
             'response.write "        <input type=""hidden"" name=""filepath" & iRowcount & """ id=""filepath" & iRowcount & """ value=""" & oList("filelocation") & oList("filename") & """ size=""10"" />" & vbcrlf
             'response.write "        <input type=""hidden"" name=""filepath" & iRowcount & """ id=""filepath" & iRowcount & """ value=""" & replace(oList("filelocation"),"custom\pub\" & session("sitename") & "\","") & oList("filename") & """ size=""10"" />" & vbcrlf
             response.write "        <input type=""hidden"" name=""filepath" & iRowcount & """ id=""filepath" & iRowcount & """ value=""" & lcl_file_url & """ size=""10"" />" & vbcrlf
          else
             response.write "&nbsp;" & vbcrlf
          end if
       else
          response.write "&nbsp;" & vbcrlf
       end if

       response.write "    </td>" & vbcrlf
       response.write "    <td" & lcl_onclick & " align=""center"" nowrap=""nowrap"">" & oList("firstviewedby_name") & "<br />" & lcl_firstviewedby_date & "</td>" & vbcrlf
       response.write "</tr>" & vbcrlf

    			oList.movenext
    wend

    response.write "</table>" & vbcrlf
    response.write "</div>" & vbcrlf

		  oList.close
  		Set oList = Nothing 
 else
  		response.write "<font color=""red""><strong>No " & lcl_list_label & " currenty exist.</strong></font>"
 end if
%>
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
sub setFirstViewedByInfo(iUserBidID)
  if iUserBidID <> "" then
     if isnumeric(iUserBidID) then
       'First check to see if a date exists.
       'If not then populate the current date.
        sSQL = "SELECT userbid_firstviewedby, isnull(userbid_firstvieweddate,'1/1/1900') AS firstvieweddate "
        sSQL = sSQL & " FROM egov_jobs_bids_userbids "
        sSQL = sSQL & " WHERE userbidid = " & iUserBidID

        set oUserBidCheck = Server.CreateObject("ADODB.Recordset")
       	oUserBidCheck.Open sSQL, Application("DSN"), 3, 1

        if not oUserBidCheck.eof then
           lcl_firstviewed_date = oUserBidCheck("firstvieweddate")

           if CDate(lcl_firstviewed_date) = CDate("1/1/1900") then
              sSQL = "UPDATE egov_jobs_bids_userbids SET "
              sSQL = sSQL & " userbid_firstvieweddate = '" & lcl_local_datetime & "', "
              sSQL = sSQL & " userbid_firstviewedby = " & session("userid")
              sSQL = sSQL & " WHERE userbidid = " & iUserBidID

              set oUserBidViewed = Server.CreateObject("ADODB.Recordset")
    	         oUserBidViewed.Open sSQL, Application("DSN"), 3, 1

              set oUserBidViewed = nothing

           end if
        end if

        oUserBidCheck.close
        set oUserBidCheck = nothing

     end if
  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb (notes) VALUES ('" & replace(p_value,"'","''") & "')"
  set rsi = Server.CreateObject("ADODB.Recordset")
 	rsi.Open sSQLi, Application("DSN"), 3, 1

  set rsi = nothing

end Sub
%>