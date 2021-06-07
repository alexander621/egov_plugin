<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: dl_mgmt.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 03/21/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  04/17/06  John Stullenberger - Initial Version
' 1.1  04/17/06  Terry Foster - Made functional
' 1.2	 10/05/06	 Steve Loar - Security, Header and nav changed
' 1.3  11/29/07  David Boyer - Removed apostrophe from the Distribution Name within the javascript link
' 2.0  01/23/08  David Boyer - Redesigned layout, added Job/Bid Postings, added isFeatureOffline check
' 2.1  01/26/09  David Boyer - Fixed bug raised when clicking the "Edit" button and trying to pass in the list name with an apostrophe.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("subscriptions,job_postings,bid_postings") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=hide

 if not userhaspermission(session("userid"), "distribution lists" ) _
    AND not userhaspermission(session("userid"), "job_postings") _
    AND not userhaspermission(session("userid"), "bid_postings") then
  	     response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the search parameters
 lcl_sc_name              = request("sc_name")
 lcl_sc_publicly_viewable = request("sc_publicly_viewable")
 lcl_sc_show_postings     = request("sc_show_postings")
 lcl_sc_list_type         = request("sc_list_type")

'Set up the link parameters for the return url for the search criteria options
 lcl_return_url_parameters = ""
 if lcl_sc_name <> "" then
    lcl_return_url_parameters = "sc_name=" & lcl_sc_name
 end if

 if lcl_sc_publicly_viewable <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_publicly_viewable=" & lcl_sc_publicly_viewable
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_publicly_viewable=" & lcl_sc_publicly_viewable
    end if
 end if

 if lcl_sc_show_postings <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_show_postings=" & lcl_sc_show_postings
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_show_postings=" & lcl_sc_show_postings
    end if
 end if

 if lcl_sc_list_type <> "" then
    if lcl_return_url_parameters = "" then
       lcl_return_url_parameters = "sc_list_type=" & lcl_sc_list_type
    else
       lcl_return_url_parameters = lcl_return_url_parameters & "&sc_list_type=" & lcl_sc_list_type
    end if
 end if

 if lcl_return_url_parameters <> "" then
    lcl_return_url_parameters = "&" & REPLACE(lcl_return_url_parameters,"%","<<PER>>")
 end if

'Convert the % in the search criteria
 lcl_sc_name        = REPLACE(request("sc_name"),"<<PER>>","%")

'Determine the list type and what label should be displayed on the page
 if UCASE(lcl_sc_list_type) = "JOB" then
    lcl_list_label        = "Job Postings"
    lcl_page_title        = "Subsription Job Postings: Categories"
    lcl_new_list_label    = "Category"
 elseif UCASE(lcl_sc_list_type) = "BID" then
    lcl_list_label        = "Bid Postings"
    lcl_page_title        = "Subscription Bid Postings: Categories/Sub-Categories"
    lcl_new_list_label    = "Category/Sub-Category"
 else
    lcl_list_label        = "Distribution Lists"
    lcl_page_title        = "Subscription Distribution Lists"
    lcl_new_list_label    = "Distribution List"
 end if

'This determines if the word "Parent Name" is shown in the results list column header.
 lcl_show_parentcolumntitle = "N"
%>
<html>
<head>
	<title>E-Gov Administration Console (<%=lcl_check%>)</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../recreation/facility.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />
	
<script language="Javascript" src="tablesort.js"></script>
<script language="Javascript">
<!--
//function deleteconfirm(ID, sName) {
function deleteconfirm(ID) {
  lcl_name = document.getElementById("distribution_name_"+ID).innerHTML;

// 	if(confirm('Do you wish to delete \'' + sName + '\'?')) {
 	if(confirm('Do you wish to delete \'' + lcl_name + '\'?')) {
     lcl_redirect_url = "&sc_name=<%=lcl_sc_name%>&sc_publicy_viewable=<%=lcl_sc_publicly_viewable%>&sc_show_postings=<%=lcl_sc_show_postings%>&sc_list_type=<%=lcl_sc_list_type%>";

 				window.location="dl_delete.asp?idlid="+ID+lcl_redirect_url;
		}
}

function openWin2(url, name) {
  lcl_width  = 800;
  lcl_height = 360;
  lcl_left   = (screen.availWidth/2) - (lcl_width/2);
  lcl_top    = (screen.availHeight/2) - (lcl_height/2);
		popupWin = window.open(url, name,"resizable,width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top);
}
//-->
</script>
</head>
<body onload="javascript:document.search_sort_form.sc_name.focus()">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<p>
<div id="centercontent">
<font size="+1"><strong><%=Session("sOrgName")%>&nbsp;<%=lcl_page_title%></strong></font><p>
<table border="0" cellspacing="0" cellpadding="5" width="800px">
  <tr>
      <td colspan="2">
          <fieldset>
            <legend><strong>Search/Sorting Option(s)&nbsp;</strong></legend><br />
          <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <form name="search_sort_form" value="dl_mgmt.asp">
              <input type="<%=lcl_hidden%>" name="sc_list_type" value="<%=lcl_sc_list_type%>" size="15" maxlength="100">
            <tr valign="top">
                <td>
                    <table border="0" cellspacing="0" cellpadding="2">
                      <tr>
                          <td nowrap="nowrap"><%=lcl_new_list_label%>:</td>
                          <td><input type="text" name="sc_name" value="<%=lcl_sc_name%>" size="30" maxlength="512"></td>
                      </tr>
                      <tr>
                          <td nowrap="nowrap">Publicly Viewable:</td>
                          <td>
                              <select name="sc_publicly_viewable">
                                <option value=""></option>
                              <%
                                if lcl_sc_publicly_viewable <> "" then
                                   if lcl_sc_publicly_viewable then
                                      lcl_selected_true  = "selected"
                                      lcl_selected_false = ""
                                   else
                                      lcl_selected_true  = ""
                                      lcl_selected_false = "selected"
                                   end if
                                else
                                   lcl_selected_true  = ""
                                   lcl_selected_false = ""
                                end if
                              %>
                                <option value="1" <%=lcl_selected_true%>>True</option>
                                <option value="0" <%=lcl_selected_false%>>False</option>
                              </select>
                          </td>
                      </tr>
                    </table>
                </td>
                <td>
                    <table border="0" cellspacing="0" cellpadding="2">
                      <tr>
                      <% if lcl_sc_list_type <> "" then %>
                          <td nowrap="nowrap">Show <%=lcl_list_label%>:</td>
                          <td>
                              <select name="sc_show_postings">
                              <%
                                if lcl_sc_show_postings <> "" then
                                   if lcl_sc_show_postings = "Y" then
                                      lcl_selected_yes = " selected"
                                      lcl_selected_no  = ""
                                   else
                                      lcl_selected_yes = ""
                                      lcl_selected_no  = " selected"
                                   end if
                                else
                                   lcl_selected_yes = ""
                                   lcl_selected_no  = " selected"
                                end if
                              %>
                                <option value="Y"<%=lcl_selected_yes%>>Yes</option>
                                <option value="N"<%=lcl_selected_no%>>No</option>
                              </select>
                          </td>
                      <% else %>
                          <td colspan="2">&nbsp;</td>
                      <% end if %>
                      </tr>
                    </table>
                </td>
            </tr>
            <tr><td colspan="2"><input type="submit" value="SEARCH" method="post" action="dl_mgmt.asp" class="button" /></td></tr>
            </form>
          </table>
          </fieldset>
          <p>
      </td>
  </tr>
  <tr>
      <td id="functionlinks">
          <input type="button" name="newDistList" id="newDistList" value="New <%=lcl_new_list_label%>" class="button" onclick="location.href='dl_edit.asp?dlid=0<%=lcl_return_url_parameters%>'" />
      </td>
      <td align="right">
      <%
        lcl_message = ""

        if request("success") = "SN" then
           lcl_message = "<strong style=""color:#FF0000"">*** Successfully Created... ***</strong>"
        elseif request("success") = "SD" then
           lcl_message = "<strong style=""color:#FF0000"">*** Successfully Deleted... ***</strong>"
        else
           lcl_message = "&nbsp;"
        end if

        if lcl_message <> "" then
           response.write lcl_message
        end if
      %>
      </td>
  </tr>
  <tr>
      <td colspan="2">
          <div class="shadow">
          <table border="0" cellpadding="5" cellspacing="0" width="100%" class="tableadmin">
            <tr align="center" valign="bottom">
                <th align="left"><%=lcl_new_list_label%></th>
                <th align="left" id="parentcolumn">&nbsp;</th>
                <th align="left" width="20%">Description</th>
                <th>Publicly<br />Viewable</th>
                <th># of<br />Subscribers</th><th>Subscribers</th>
                <th>Delete</th>
            </tr>
<%
 Dim sSql, oList

'Are we performing a search for specific criteria or returning the entire list?
'If NOT then return the entire list.
'------------------------------------------------------------------------------
 if lcl_sc_name = "" AND lcl_sc_publicly_viewable = "" then
'------------------------------------------------------------------------------
   'Retrieve all of the distribution lists for the org
   	sSQL = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
    sSQL = sSQL & " distributionlistdisplay, orgid, parentid, "
    sSQL = sSQL & " (CASE WHEN distributionlisttype IS NULL OR distributionlisttype = '' THEN '' ELSE distributionlisttype END) AS distributionlisttype "
    sSQL = sSQL & " FROM egov_class_distributionlist "
    sSQL = sSQL & " WHERE orgid = " & session("orgid")

    if lcl_sc_list_type <> "" then
       sSQL = sSQL & " AND UPPER(distributionlisttype) LIKE ('" & UCASE(lcl_sc_list_type) & "') "

      'If the listtype = "BID" then only pull the distributionlistids that have do NOT have a parentid associated to them.
       if lcl_sc_list_type = "BID" then
          sSQL = sSQL & " AND (parentid IS NULL OR parentid = '') "
       end if
    else
      'This returns all of the original distributionlists (pre-job/bid postings)
       sSQL = sSQL & " AND (distributionlisttype = '' OR distributionlisttype is null) "
    end if

    sSQL = sSQL & " ORDER BY UPPER(distributionlistname)"

   	Set oList = Server.CreateObject("ADODB.Recordset")
   	oList.Open sSQL, Application("DSN"), 0, 1

   	if not oList.eof then
       lcl_bgcolor = "#ffffff"

       do while not oList.eof
          lcl_bgcolor          = changeBGColor(lcl_bgcolor,"","")
          lcl_subscribercount1 = GetSubscriberCount(oList("distributionlistid"))

          if lcl_subscribercount1 < 1 then
             lcl_subscribercount1 = "&nbsp;"
          end if

          response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
          response.write "      <td><a href=""dl_edit.asp?dlid=" & oList("distributionlistid") & lcl_return_url_parameters & """><strong><span id=""distribution_name_" & oList("distributionlistid") & """>" & oList("distributionlistname") & "</span></strong></a></td>" & vbcrlf
          response.write "      <td>&nbsp;</td>" & vbcrlf
          response.write "      <td align=""left"">"   & oList("distributionlistdescription")   & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & trim(oList("distributionlistdisplay")) & "</td>" & vbcrlf
          response.write "      <td align=""center"">" & lcl_subscribercount1                   & "</td>" & vbcrlf
          'response.write "      <td align=""center""><a href=""javascript:openWin2('dl_manage_subscribers.asp?idlid=" & oList("distributionlistid") & "','_blank')"">Edit</a></td>" & vbcrlf
          'response.write "      <td align=""center""><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""deleteconfirm(" & oList("distributionlistid") & ")""></td>" & vbcrlf
          response.write "      <td align=""center""><input type=""button"" name=""sEditSub" & oList("distributionlistid") & """ id=""sEditSub" & oList("distributionlistid") & """ value=""Edit"" class=""button"" onclick=""openWin2('dl_manage_subscribers.asp?idlid=" & oList("distributionlistid") & "','_blank')"" /></td>" & vbcrlf
          response.write "      <td align=""center""><input type=""button"" name=""sDeleteCat" & oList("distributionlistid") & """ id=""sDeleteCat" & oList("distributionlistid") & """ value=""Delete"" class=""button"" onclick=""deleteconfirm(" & oList("distributionlistid") & ")"" /></td>" & vbcrlf
          response.write "  </tr>" & vbcrlf

          if lcl_sc_show_postings = "Y" then
             displayJobsBids oList("distributionlistid"), lcl_sc_list_type, lcl_return_parameters, lcl_bgcolor
          end if

         'Retrieve all sub-categories if available
         	sSQLs = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
          sSQLs = sSQLs & " distributionlistdisplay, parentid, "
          sSQLs = sSQLs & " (CASE WHEN distributionlisttype IS NULL OR distributionlisttype = '' THEN '' ELSE distributionlisttype END) AS distributionlisttype "
          sSQLs = sSQLs & " FROM egov_class_distributionlist "
          sSQLs = sSQLs & " WHERE orgid = "  & session("orgid")
          sSQLs = sSQLs & " AND parentid = " & oList("distributionlistid")
          sSQLs = sSQLs & " ORDER BY UPPER(distributionlistname)"

         	set rs = Server.CreateObject("ADODB.Recordset")
         	rs.Open sSQLs, Application("DSN"), 0, 1

          if not rs.eof then
             while not rs.eof
                lcl_bgcolor          = changeBGColor(lcl_bgcolor,"","")
                lcl_subscribercount2 = GetSubscriberCount(rs("distributionlistid"))

                if lcl_subscribercount2 < 1 then
                   lcl_subscribercount2 = "&nbsp;"
                end if

                response.write "  <tr align=""center"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
                response.write "      <td align=""left"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""dl_edit.asp?dlid=" & rs("distributionlistid") & lcl_return_url_parameters & """><span id=""distribution_name_" & rs("distributionlistid") & """>" & trim(rs("distributionlistname")) & "</span></a></td>" & vbcrlf
                response.write "      <td>&nbsp;</td>" & vbcrlf
                response.write "      <td align=""left"">"   & rs("distributionlistdescription")   & "</td>" & vbcrlf
                response.write "      <td>"                  & trim(rs("distributionlistdisplay")) & "</td>" & vbcrlf
                response.write "      <td align=""center"">" & lcl_subscribercount2                & "</td>" & vbcrlf
                'response.write "      <td><a href=""javascript:openWin2('dl_manage_subscribers.asp?idlid=" & rs("distributionlistid") & "','_blank')"">Edit</a></td>" & vbcrlf
                'response.write "      <td><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""deleteconfirm(" & rs("distributionlistid") & ")""></td>" & vbcrlf
                response.write "      <td><input type=""button"" name=""sEditSub" & rs("distributionlistid") & """ id=""sEditSub" & rs("distributionlistid") & """ value=""Edit"" class=""button"" onclick=""openWin2('dl_manage_subscribers.asp?idlid=" & rs("distributionlistid") & "','_blank')"" /></td>" & vbcrlf
                response.write "      <td><input type=""button"" name=""sDeleteCat" & rs("distributionlistid") & """ id=""sDeleteCat" & rs("distributionlistid") & """ value=""Delete"" class=""button"" onclick=""deleteconfirm(" & rs("distributionlistid") & ")"" /></td>" & vbcrlf
                response.write "  </tr>" & vbcrlf

                if lcl_sc_show_postings = "Y" then
                   displayJobsBids rs("distributionlistid"), lcl_sc_list_type, lcl_return_parameters, lcl_bgcolor
                end if

                rs.movenext
             wend
          end if
       			oList.movenext
       loop
   		  oList.close
     		Set oList = Nothing
   	end if
'------------------------------------------------------------------------------
 else  'If we are performing a search then limit done the results
'------------------------------------------------------------------------------
   'Retrieve all of the distribution lists for the org
   	sSQL = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
    sSQL = sSQL & " distributionlistdisplay, orgid, parentid, "
    sSQL = sSQL & " (CASE WHEN distributionlisttype IS NULL OR distributionlisttype = '' THEN '' ELSE distributionlisttype END) AS distributionlisttype "
    sSQL = sSQL & " FROM egov_class_distributionlist "
    sSQL = sSQL & " WHERE orgid = " & session("orgid")

   'Evaluate search criteria
    if lcl_sc_name <> "" then
       sSQL = sSQL & " AND UPPER(distributionlistname) LIKE ('%" & UCASE(lcl_sc_name) & "%') "
    end if

    if lcl_sc_publicly_viewable <> "" then
       sSQL = sSQL & " AND UPPER(distributionlistdisplay) LIKE (" & UCASE(lcl_sc_publicly_viewable) & ") "
    end if

    if lcl_sc_list_type <> "" then
       sSQL = sSQL & " AND UPPER(distributionlisttype) LIKE ('" & UCASE(lcl_sc_list_type) & "') "
    else
       sSQL = sSQL & " AND (distributionlisttype = '' OR distributionlisttype is null) "
    end if

    sSQL = sSQL & " ORDER BY UPPER(distributionlistname)"

   	Set oList = Server.CreateObject("ADODB.Recordset")
   	oList.Open sSQL, Application("DSN"), 0, 1

   	if not oList.eof then
       lcl_bgcolor = "#ffffff"
       do while not oList.eof
          lcl_bgcolor                = changeBGColor(lcl_bgcolor,"","")
          lcl_show_parentcolumntitle = "Y"
          lcl_subscribercount3       = GetSubscriberCount(oList("distributionlistid"))

          if lcl_subscribercount3 < 1 then
             lcl_subscribercount3 = "&nbsp;"
          end if

          response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf

         'Category
          if oList("parentid") = "" OR isnull(oList("parentid")) then
             lcl_parent_name = "&nbsp;"
         'Sub-Category
          else
             lcl_parent_name = "<strong><i>Belongs to: </i></strong>" & getCategoryName(oList("parentid"))
          end if

          response.write "    <td nowrap=""nowrap"">" & vbcrlf
          response.write "        <a href=""dl_edit.asp?dlid=" & oList("distributionlistid") & lcl_return_url_parameters & """>"
          response.write "        <span id=""distribution_name_" & oList("distributionlistid") & """>" & trim(oList("distributionlistname")) & "</span></a>" & vbcrlf
          response.write "    </td>" & vbcrlf
          response.write "    <td align=""left"">"   & lcl_parent_name                        & "</td>" & vbcrlf
          response.write "    <td align=""left"">"   & oList("distributionlistdescription")   & "</td>" & vbcrlf
          response.write "    <td align=""center"">" & trim(oList("distributionlistdisplay")) & "</td>" & vbcrlf
          response.write "    <td align=""center"">" & lcl_subscribercount3                   & "</td>" & vbcrlf
          'response.write "    <td align=""center"">"
          'response.write "        <a href=""javascript:openWin2('dl_manage_subscribers.asp?idlid=" & oList("distributionlistid") & "','_blank')"">Edit</a>"
          'response.write "    </td>" & vbclrf
          'response.write "    <td align=""center"">"
          'response.write "        <img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""deleteconfirm(" & oList("distributionlistid") & ")"">"
          'response.write "    </td>" & vbclrf
          response.write "      <td align=""center""><input type=""button"" name=""sEditSub" & oList("distributionlistid") & """ id=""sEditSub" & oList("distributionlistid") & """ value=""Edit"" class=""button"" onclick=""openWin2('dl_manage_subscribers.asp?idlid=" & oList("distributionlistid") & "','_blank')"" /></td>" & vbcrlf
          response.write "      <td align=""center""><input type=""button"" name=""sDeleteCat" & oList("distributionlistid") & """ id=""sDeleteCat" & oList("distributionlistid") & """ value=""Delete"" class=""button"" onclick=""deleteconfirm(" & oList("distributionlistid") & ")"" /></td>" & vbcrlf
          response.write "</tr>" & vbclrf

          if lcl_sc_show_postings = "Y" then
             displayJobsBids oList("distributionlistid"), lcl_sc_list_type, lcl_return_parameters, lcl_bgcolor
          end if

       			oList.movenext
       loop
		     oList.close
     		Set oList = Nothing 
   '---------------------------------------------------------------------------
   'If search criteria has been entered and no records are found on the categories then
   'if this is a listype = BID we still need to search through the sub-categories.
    else
   '---------------------------------------------------------------------------
        if lcl_sc_list_type = "BID" then
          	sSQLs = "SELECT distributionlistid, distributionlistname, distributionlistdescription, distributionlistdisplay, parentid, "
           sSQLs = sSQLs & " (CASE WHEN distributionlisttype IS NULL OR distributionlisttype = '' THEN '' ELSE distributionlisttype END) AS distributionlisttype "
           sSQLs = sSQLs & " FROM egov_class_distributionlist "
           sSQLs = sSQLs & " WHERE orgid = "  & session("orgid")
           sSQLs = sSQLs & " AND (parentid IS NOT NULL OR parentid <> '') "
           sSQLs = sSQLs & " AND UPPER(distributionlisttype) LIKE ('" & UCASE(lcl_sc_list_type) & "') "

           if lcl_sc_name <> "" then
              sSQLs = sSQLs & " AND UPPER(distributionlistname) LIKE ('%" & UCASE(lcl_sc_name) & "%') "
           end if

           if lcl_sc_publicly_viewable <> "" then
              sSQLs = sSQLs & " AND UPPER(distributionlistdisplay) LIKE (" & UCASE(lcl_sc_publicly_viewable) & ") "
           end if

           sSQLs = sSQLs & " ORDER BY UPPER(distributionlistname)"

          	set rs = Server.CreateObject("ADODB.Recordset")
          	rs.Open sSQLs, Application("DSN"), 0, 1

           if not rs.eof then
              while not rs.eof
                 lcl_bgcolor          = changeBGColor(lcl_bgcolor,"","")
                 lcl_subscribercount4 = GetSubscriberCount(rs("distributionlistid"))

                 if lcl_subscribercount4 < 1 then
                    lcl_subscribercount4 = "&nbsp;"
                 end if

                 response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
                 response.write "      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""dl_edit.asp?dlid=" & rs("distributionlistid") & lcl_return_url_parameters & """><span id=""distribution_name_" & rs("distributionlistid") & """>" & rs("distributionlistname") & "</span></a></td>" & vbcrlf
                 response.write "      <td>&nbsp;</td>" & vbcrlf
                 response.write "      <td align=""left"">"   & oList("distributionlistdescription") & "</td>" & vbcrlf
                 response.write "      <td align=""center"">" & trim(rs("distributionlistdisplay"))  & "</td>" & vbcrlf
                 response.write "      <td align=""center"">" & lcl_subscribercount4                 & "</td>" & vbcrlf
                 'response.write "      <td align=""center""><a href=""javascript:openWin2('dl_manage_subscribers.asp?idlid=" & rs("distributionlistid") & "','_blank')"">Edit</a></td>" & vbcrlf
                 'response.write "      <td align=""center""><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""deleteconfirm(" & rs("distributionlistid") & ")""></td>" & vbcrlf
                 response.write "      <td align=""center""><input type=""button"" name=""sEditSub" & rs("distributionlistid") & """ id=""sEditSub" & rs("distributionlistid") & """ value=""Edit"" class=""button"" onclick=""openWin2('dl_manage_subscribers.asp?idlid=" & rs("distributionlistid") & "','_blank')"" /></td>" & vbcrlf
                 response.write "      <td align=""center""><input type=""button"" name=""sDeleteCat" & rs("distributionlistid") & """ id=""sDeleteCat" & rs("distributionlistid") & """ value=""Delete"" class=""button"" onclick=""deleteconfirm(" & rs("distributionlistid") & ")"" /></td>" & vbcrlf
                 response.write "  </tr>" & vbcrlf

                 if lcl_sc_show_postings = "Y" then
                    displayJobsBids rs("distributionlistid"), lcl_sc_list_type, lcl_return_parameters, lcl_bgcolor
                 end if

                 rs.movenext
              wend
           else
            		response.write "<tr><td colspan=""6""><font color=""#ff0000""><strong>No " & lcl_list_label & " currenty exist.</strong></font></td></tr>" & vbcrlf
           end if
       else
         		response.write "<tr><td colspan=""6""><font color=""#ff0000""><strong>No " & lcl_list_label & " currenty exist.</strong></font></td></tr>" & vbcrlf
       end if
    end if
 end if
%>
          </table>
          </div>
      </td>
  </tr>
</table>
</div>
<!--#Include file="../admin_footer.asp"-->  

<%
 'This determines if "Parent" should display on the column header or not.
  if lcl_show_parentcolumntitle = "Y" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write "  document.getElementById(""parentcolumn"").style.width = '30%';" & vbcrlf
     response.write "  document.getElementById(""parentcolumn"").innerHTML = 'Parent';" & vbcrlf
     response.write "</script>" & vbcrlf
  end if
%>
</body>
</html>
<%
'------------------------------------------------------------------------------
function getCategoryName(p_value)
  sSQL = "SELECT distributionlistname "
  sSQL = sSQL & " FROM egov_class_distributionlist "
  sSQL = sSQL & " WHERE orgid = "  & session("orgid")
  sSQL = sSQL & " AND distributionlistid = " & p_value

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 0, 1

  if not rs.eof then
     lcl_return_value = rs("distributionlistname")
  else
     lcl_return_value = ""
  end if

  getCategoryName = lcl_return_value

end function

'------------------------------------------------------------------------------
sub displayJobsBids(p_dlid, p_sc_list_type, p_return_parameters, p_bgcolor)
 'Retrieve all jobs/bids for this category/sub-category
  sSQLjb = "SELECT jb.posting_id, jb.jobbid_id, jb.title, jb.active_flag, "
  sSQLjb = sSQLjb & " (select s.status_name from egov_statuses s where s.status_id = jb.status_id) AS status_name "
  sSQLjb = sSQLjb & " FROM egov_jobs_bids jb, egov_distributionlists_jobbids djb "
  sSQLjb = sSQLjb & " WHERE jb.posting_id = djb.posting_id "
  sSQLjb = sSQLjb & " AND jb.posting_type = '" & p_sc_list_type & "'"
  sSQLjb = sSQLjb & " AND djb.distributionlistid = " & p_dlid
  sSQLjb = sSQLjb & " ORDER BY jb.title "

 	set rsjb = Server.CreateObject("ADODB.Recordset")
	rsjb.Open sSQLjb, Application("DSN"), 0, 1

  if not rsjb.eof then
     while not rsjb.eof

        response.write "  <tr bgcolor=""" & p_bgcolor & """>" & vbcrlf
        response.write "      <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- " & rsjb("title") & "</td>" & vbcrlf
        response.write "      <td><strong><i>" & lcl_sc_list_type & " ID: </i></strong>" & rsjb("jobbid_id") & "</td>" & vbcrlf
        response.write "      <td><strong><i>Status: </i></strong>" & rsjb("status_name") & "</td>" & vbcrlf
        response.write "      <td><strong><i>Active: </i></strong>" & rsjb("active_flag") & "</td>" & vbcrlf
        response.write "      <td colspan=""2"">&nbsp;</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        rsjb.movenext
     wend
  end if

  rsjb.close
  set rsjb = nothing
end Sub

'------------------------------------------------------------------------------
function GetSubscriberCount(iDistributionListID)

 lcl_return = 0

 if iDistributionListID <> "" then
   	sSQL = "SELECT count(u.userid) as total_subscribers "
 	  sSQL = sSQL & " FROM egov_users u "
   	sSQL = sSQL & " INNER JOIN egov_class_distributionlist_to_user ug ON u.userid = ug.userid "
 	  sSQL = sSQL & " WHERE (ug.distributionlistid = '" & iDistributionListID & "') "
    sSQL = sSQL & " AND u.isdeleted = 0 "

   	'sSQL = "SELECT COUNT(useremail) AS total_subscribers "
    'sSQL = sSQL & " FROM egov_dl_user_list "
    'sSQL = sSQL & " WHERE distributionlistid = " & iDistributionListID

   	set oSubscriberCount = Server.CreateObject("ADODB.Recordset")
   	oSubscriberCount.Open sSQL, Application("DSN"), 3, 1

   	if not oSubscriberCount.eof then
     		lcl_return = CLng(oSubscriberCount("total_subscribers"))
    end if

   	oSubscriberCount.close
   	set oSubscriberCount = nothing 
 end if

 GetSubscriberCount = lcl_return

end function
%>