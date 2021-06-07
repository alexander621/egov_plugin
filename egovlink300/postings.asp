 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="postings_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: postings.asp
' AUTHOR:   David Boyer
' CREATED:  01/04/2008
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the Job/Bid Postings
'
' MODIFICATION HISTORY
' 1.0  02/05/08	 David Boyer - Initial Version
' 1.1  10/14/08  David Boyer - Added "Required Login to view Postings" feature check.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim orgHasWordPress

'Check to see if the feature is offline
 if isFeatureOffline("job_postings,bid_postings") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 if not OrgHasFeature(iorgid,"job_postings") and not OrgHasFeature(iorgid,"bid_postings") then response.redirect "default.asp"

 Dim oPostingsOrg
 Set oPostingsOrg = New classOrganization

 lcl_hidden = "HIDDEN"  'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide

'BEGIN: Check for apotrophes --------------------------------------------------
 lcl_listtype        = ""
 lcl_sc_category_id  = ""
 lcl_sc_status_id    = ""
 lcl_sc_show_expired = ""

'Check required fields
 if request("listtype") <> "" then
    if not containsApostrophe(request("listtype")) then
       lcl_list_type = request("listtype")

      'Verify that the p_list_type is a valid status_type for the org
       lcl_valid = checkStatusType(lcl_list_type)

       if lcl_valid <> "Y" then
          response.redirect sEgovWebsiteURL
       end if

    end if
 end if

 if lcl_list_type = "" then
    response.redirect sEgovWebsiteURL
 end if

'Retrieve the search criteria parameters
 if request("sc_category_id") <> "" then
    if not containsApostrophe(request("sc_category_id")) then
       lcl_sc_category_id = request("sc_category_id")
    end if
 end if

 if request("sc_status_id") <> "" then
    if not containsApostrophe(request("sc_status_id")) and (isnumeric(request("sc_status_id")) or request("sc_status_id") = "none") then
       lcl_sc_status_id = request("sc_status_id")
    end if
 end if

 if request("sc_show_expired") <> "" then
    if not containsApostrophe(request("sc_show_expired")) then
       lcl_sc_show_expired = request("sc_show_expired")
    end if
 end if
'END: Check for apotrophes ----------------------------------------------------

'Set up ORG variables
 if lcl_list_type = "JOB" then
    lcl_feature_name = oPostingsOrg.GetOrgFeatureName("job_postings")
 elseif lcl_list_type = "BID" then
    lcl_feature_name = oPostingsOrg.GetOrgFeatureName("bid_postings")
 end if

 lcl_org_name        = oPostingsOrg.GetOrgName()
 lcl_org_state       = oPostingsOrg.GetState()
 lcl_org_featurename = lcl_feature_name

'Set up local variables based on posting (list) type.
 if lcl_list_type = "JOB" then
    lcl_list_label = "Job"
    'lcl_list_title = "Job Postings"
    lcl_sc_label   = "Categories"
 elseif lcl_list_type = "BID" then
    lcl_list_label = "Bid"
    'lcl_list_title = "Bid Postings"
    lcl_sc_label   = "Category/Sub-Category"
 end if

'Retrieve the search parameters
 'lcl_sc_category_id  = request("sc_category_id")
 'lcl_sc_status_id    = request("sc_status_id")
 'lcl_sc_show_expired = request("sc_show_expired")

 if lcl_sc_status_id = "" then
    lcl_sc_status_id = getStatusDefault(lcl_list_type)
 end if

 if lcl_sc_show_expired = "" then
    lcl_sc_show_expired = "N"
 end if

'Set up search criteria session variables
 session("sc_category_id")  = lcl_sc_category_id
 session("sc_status_id")    = lcl_sc_status_id
 session("sc_show_expired") = lcl_sc_show_expired

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)
 
 orgHasWordPress = orghasfeature( iOrgID, "wordpress public interface" )
 
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
<p>
<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;">
  <tr valign="top">
      <td>
          <% oPostingsOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename %>

          <!--<font class="pagetitle">Welcome to the <%'oPostingsOrg.GetOrgName()%>, <%'oPostingsOrg.GetState()%>, <%'lcl_feature_name %></font>-->
      </td>
      <td align="right">
          <table border="0" cellspacing="0" cellpadding="2">
        <%
          if check_for_jobbid_categories(lcl_list_type) = "Y" then
            'If the user HAS signed in then set the link to their manage subscriptions account page.
            'If the user is NOT signed in then set the link to the public subscriptions page.
           		if request.cookies("userid") <> "" AND request.cookies("userid") <> "-1" then
                lcl_subscriptions_url = "manage_mail_lists.asp"
             else
                If orgHasWordPress Then  
                  lcl_subscriptions_url = getOrganization_WP_URL( iOrgID, "wp_subscriptions_url" ) & "#subscriptionlist/" & lcl_list_type
                Else 
                  lcl_subscriptions_url = "subscriptions/subscribe.asp?listtype=" & lcl_list_type
                End If
             end if

             response.write "            <tr>" & vbcrlf
             'response.write "                <td>[<a href=""" & lcl_subscriptions_url & """>Want to be notified about <strong>New " & lcl_list_title & "</strong>...</a>]</td>" & vbcrlf
             response.write "                <td>[<a href=""" & lcl_subscriptions_url & """>Want to be notified about <strong>New " & lcl_org_featurename & "</strong>...</a>]</td>" & vbcrlf
             response.write "            </tr>" & vbcrlf

             if lcl_list_type = "JOB" then
               'Get the postings_email set up for the ORG
                sSQLe = "SELECT postings_email, defaultemail FROM organizations WHERE orgid = " & iorgid
                set rse = Server.CreateObject("ADODB.Recordset")
                rse.Open sSQLe, Application("DSN"), 0, 1

                if not rse.eof then
                  'First check for the postings_email.  If it is blank then use the org default_email
                   if rse("postings_email") = "" OR isnull(rse("postings_email")) then
                      lcl_email = rse("defaultemail")
                   else
                      lcl_email = rse("postings_email")
                   end if
                else
                   lcl_email = ""
                end if

                rse.close
                set rse = nothing

                'response.write "<tr align=""center""><td>[<a href=""mailto:" & lcl_email & "?subject=" & lcl_list_title & " Resume""><strong>Email your resume</strong>...</a>]</td></tr>" & vbcrlf

             end if
          else
             response.write "            <tr><td>&nbsp;</td></tr>" & vbcrlf
          end if
        %>
          </table>
      </td>
  </tr>
</table>
</p>

<% RegisteredUserDisplay("") %>

<div id="content">
  <div id="centercontent">

<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;">
  <form name="postings" action="postings.asp" method="post">
    <input type="<%=lcl_hidden%>" name="listtype" value="<%=lcl_list_type%>" size="5" maxlength="5" />
  <tr>
      <td>
          <p>
          <fieldset>
            <p><legend><strong>Search Option(s)</strong>&nbsp;</legend></p>
            <table border="0" cellspacing="0" cellpadding="2" width="100%">
              <tr>
                  <td width="200"><strong><%=lcl_sc_label%>:</strong></td>
                  <td>
                      <select name="sc_category_id">
                        <option value="">All</option>
                        <%
                         'Retrieve all of the categories that are active
                      	   sSQL = "SELECT DISTINCT dl.distributionlistname, dl.distributionlistid, dl.parentid, "
                          sSQL = sSQL & " dl.distributionlisttype, dl.distributionlistdisplay "
               	          sSQL = sSQL & " FROM egov_class_distributionlist dl, egov_distributionlists_jobbids djb "
               	          sSQL = sSQL & " WHERE dl.distributionlistid = djb.distributionlistid "
               	          sSQL = sSQL & " AND dl.orgid = " & iorgid
               	          sSQL = sSQL & " AND UPPER(dl.distributionlisttype) = '" & UCASE(lcl_list_type) & "' "
                          sSQL = sSQL & " AND (dl.parentid IS NULL OR dl.parentid = '') "
               	          sSQL = sSQL & " ORDER BY dl.distributionlistname, dl.distributionlistid, dl.parentid, "
                          sSQL = sSQL & " dl.distributionlisttype, dl.distributionlistdisplay "

               	         	set oCategory = Server.CreateObject("ADODB.Recordset")
               	         	oCategory.Open sSQL, Application("DSN"), 0, 1

               	          if not oCategory.eof then
               	             while not oCategory.eof
                                if oCategory("distributionlistdisplay") then
                                   if jobbid_per_dlist_count(oCategory("distributionlistid"),oCategory("distributionlisttype"),"","",lcl_sc_show_expired) > CLng(0) then
                                      if lcl_sc_category_id = "C" & oCategory("distributionlistid") then
                                         lcl_category_selected = " selected"
                                      else
                                         lcl_category_selected = ""
                                      end if

                                      response.write "  <option value=""C" & oCategory("distributionlistid") & """" & lcl_category_selected & ">" & oCategory("distributionlistname") & "</option>" & vbcrlf
                                   end if

                                  'Retrieve all of the sub-categories that are active
                               	   sSQL = "SELECT DISTINCT dl.distributionlistname, dl.distributionlistid, dl.parentid, "
                                   sSQL = sSQL & " dl.distributionlistdisplay, dl.distributionlisttype "
               	                   sSQL = sSQL & " FROM egov_class_distributionlist dl, egov_distributionlists_jobbids djb "
                        	          sSQL = sSQL & " WHERE dl.distributionlistid = djb.distributionlistid "
                        	          sSQL = sSQL & " AND dl.orgid = " & iorgid
                        	          sSQL = sSQL & " AND UPPER(dl.distributionlisttype) = '" & UCASE(lcl_list_type) & "' "
                                   sSQL = sSQL & " AND dl.parentid = " & oCategory("distributionlistid")
                        	          sSQL = sSQL & " ORDER BY dl.distributionlistname, dl.distributionlistid, dl.parentid, "
                                   sSQL = sSQL & " dl.distributionlistdisplay, dl.distributionlisttype "

                        	         	set oSubCat = Server.CreateObject("ADODB.Recordset")
                        	         	oSubCat.Open sSQL, Application("DSN"), 0, 1

                        	          if not oSubCat.eof then
                        	             while not oSubCat.eof
                                         if oSubCat("distributionlistdisplay") then
                                            if jobbid_per_dlist_count(oSubCat("distributionlistid"),oSubCat("distributionlisttype"),"","",lcl_sc_show_expired) > CLng(0) then
                                               if lcl_sc_category_id = "SC" & oSubCat("distributionlistid") then
                                                  lcl_subcategory_selected = " selected"
                                               else
                                                  lcl_subcategory_selected = ""
                                               end if

                                               lcl_parent_name = getCategoryName(oSubCat("parentid"),"N")
                                               response.write "  <option value=""SC" & oSubCat("distributionlistid") & """" & lcl_subcategory_selected & ">" & lcl_parent_name & ": " & oSubCat("distributionlistname") & "</option>" & vbcrlf
                                            end if
                                         end if
                                         oSubCat.movenext
                                      wend
                                   end if

                                   set oSubCat = nothing

                                end if
                                oCategory.movenext
                             wend
                          end if

                          oCategory.close
                          set oCategory = nothing
                        %>
                      </select>
                  </td>
              </tr>
              <tr>
                  <td><strong>Status:</strong></td>
                  <td>
                      <select name="sc_status_id">
                        <option value="none">All</option>
                        <%
                          sSQL = "SELECT DISTINCT s.status_name, jb.status_id, s.status_order "
                          sSQL = sSQL & " FROM egov_jobs_bids jb, egov_statuses s "
                          sSQL = sSQL & " WHERE jb.status_id = s.status_id "
                          sSQL = sSQL & " AND jb.posting_type = '" & lcl_list_type & "' "
                          sSQL = sSQL & " AND jb.orgid = " & iorgid
                          sSQL = sSQL & " AND jb.active_flag = 'Y' "
                          sSQL = sSQL & " ORDER BY s.status_order "

                         	set oStatus = Server.CreateObject("ADODB.Recordset")
                         	oStatus.Open sSQL, Application("DSN"), 0, 1

                          if not oStatus.eof then
                             while not oStatus.eof
                                if lcl_sc_status_id <> "" AND lcl_sc_status_id <> "none" then
                                      lcl_selected = ""
				      on error resume next
                                   	if clng(lcl_sc_status_id) = clng(oStatus("status_id")) then lcl_selected = "selected"
				      on error goto 0
                                else
                                   lcl_selected = ""
                                end if

                                response.write "  <option value=""" & oStatus("status_id") & """ " & lcl_selected & ">" & oStatus("status_name") & "</option>" & vbcrlf
                                oStatus.movenext
                             wend
                          end if

                          oStatus.close
                          set oStatus = nothing
                        %>
                      </select>
                  </td>
              </tr>
              <tr>
                  <td width="200"><strong>Show Expired Postings:</strong></td>
                  <td>
                    <%
                      if lcl_sc_show_expired <> "" then
                         if lcl_sc_show_expired = "Y" then
                            lcl_selected_yes = "selected"
                            lcl_selected_no  = ""
                         else
                            lcl_selected_yes = ""
                            lcl_selected_no  = "selected"
                         end if
                      else
                         lcl_selected_yes = ""
                         lcl_selected_no  = "selected"
                      end if
                    %>
                      <select name="sc_show_expired">
                        <option value="Y"<%=lcl_selected_yes%>>Yes</option>
                        <option value="N"<%=lcl_selected_no%>>No</option>
                      </select>
                  </td>
              </tr>
              <tr>
                  <td colspan="2">&nbsp;</td>
              </tr>
              <tr>
                  <td colspan="2"><input type="submit" name="sAction" value="Search" class="button" /></td>
              </tr>
            </table>
          </fieldset>
          </p>
      </td>
  </tr>
  <tr>
      <td>
<!-- <table cellspacing="0" cellpadding="5" width="900" style="background-color:#9c192f;border-top: solid #000000 1px;border-left: solid #000000 1px;border-right: solid #000000 1px;border-bottom: solid #000000 1px;"> -->
  <%
   '---------------------------------------------------------------------------
   '-- No search criteria entered ---------------------------------------------
   '---------------------------------------------------------------------------
    session("lcl_total_postings") = 0

    if lcl_sc_category_id = "" AND lcl_sc_status_id = "none" then
   '---------------------------------------------------------------------------
      'Retrieve all of the categories for this org
   	   sSQL = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
       sSQL = sSQL & " distributionlistdisplay, orgid, distributionlisttype, parentid "
       sSQL = sSQL & " FROM egov_class_distributionlist "
       sSQL = sSQL & " WHERE orgid = " & iorgid
       sSQL = sSQL & " AND UPPER(distributionlisttype) = '" & UCASE(lcl_list_type) & "' "
       sSQL = sSQL & " AND (parentid IS NULL OR parentid = '') "
       sSQL = sSQL & " ORDER BY UPPER(distributionlistname)"

      	set oOrgCategories = Server.CreateObject("ADODB.Recordset")
      	oOrgCategories.Open sSQL, Application("DSN"), 0, 1

       if not oOrgCategories.eof then
          while not oOrgCategories.eof
             if oOrgCategories("distributionlistname") <> "" then
                lcl_distribution_name = oOrgCategories("distributionlistname")
             else
                lcl_distribution_name = "&nbsp;"
             end if

             if oOrgCategories("distributionlistdescription") <> "" then
                lcl_distribution_desc = "&nbsp;-&nbsp;" & oOrgCategories("distributionlistdescription")
             else
                lcl_distribution_desc = ""
             end if

             if oOrgCategories("parentid") <> "" AND not isnull(oOrgCategories("parentid")) then
                lcl_category_name = getCategoryName(oOrgCategories("parentid"),"N") & ": "
             else
                lcl_category_name = ""
             end if

             if oOrgCategories("distributionlistdisplay") then
                if jobbid_per_dlist_count(oOrgCategories("distributionlistid"),oOrgCategories("distributionlisttype"),lcl_sc_category_id,lcl_sc_status_id,lcl_sc_show_expired) > CLng(0) then
                   response.write "<p><table cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#000000"">" & vbcrlf
                   response.write "  <tr align=""left"" bgcolor=""#C0C0C0"">" & vbcrlf
                   response.write "      <td>" & lcl_category_name & "<strong>" & lcl_distribution_name & "</strong>" & lcl_distribution_desc & "</td>" & vbcrlf
                   response.write "      <td align=""center"" width=""60""><strong>" & lcl_list_label & "s: " & jobbid_per_dlist_count(oOrgCategories("distributionlistid"),oOrgCategories("distributionlisttype"),lcl_sc_catgory_id,lcl_sc_status_id,lcl_sc_show_expired) & "</strong></td>" & vbcrlf
                   response.write "  </tr>" & vbcrlf
                   response.write "</table>" & vbcrlf
                end if

               'Retrieve all of the job/bid postings for this category
                displayPostings oOrgCategories("distributionlistid"), oOrgCategories("distributionlisttype"), lcl_sc_status_id, lcl_sc_show_expired

             else
                session("lcl_total_postings") = session("lcl_total_postings")
             end if

            'Retrieve all sub-categories if available
            	sSQL = "SELECT distributionlistid, distributionlistname, distributionlistdescription, "
             sSQL = sSQL & " distributionlistdisplay, distributionlisttype, parentid "
             sSQL = sSQL & " FROM egov_class_distributionlist "
             sSQL = sSQL & " WHERE orgid = "  & iorgid
             sSQL = sSQL & " AND parentid = " & oOrgCategories("distributionlistid")
             sSQL = sSQL & " ORDER BY UPPER(distributionlistname)"

            	set oOrgSubCat = Server.CreateObject("ADODB.Recordset")
            	oOrgSubCat.Open sSQL, Application("DSN"), 0, 1

             if not oOrgSubCat.eof then
                while not oOrgSubCat.eof
                   if oOrgSubCat("distributionlistdisplay") then
                      if jobbid_per_dlist_count(oOrgSubCat("distributionlistid"),oOrgSubCat("distributionlisttype"),lcl_sc_category_id,lcl_sc_status_id,lcl_sc_show_expired) > CLng(0) then

                         if oOrgSubCat("distributionlistdescription") <> "" then
                            lcl_distribution_desc = "&nbsp;-&nbsp;" & oOrgSubCat("distributionlistdescription")
                         else
                            lcl_distribution_desc = ""
                         end if

                         response.write "<p><table cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#000000"">" & vbcrlf
                         response.write "  <tr align=""left"" bgcolor=""#C0C0C0"">" & vbcrlf
                         response.write "      <td>" & getCategoryName(oOrgSubCat("parentid"),"N") & ":&nbsp;<strong>" & oOrgSubCat("distributionlistname") & "</strong>" & lcl_distribution_desc & "</td>" & vbcrlf
                         response.write "      <td align=""center"" width=""60""><strong>" & lcl_list_label & "s: " & jobbid_per_dlist_count(oOrgSubCat("distributionlistid"),oOrgSubCat("distributionlisttype"),lcl_sc_category_id,lcl_sc_status_id,lcl_sc_show_expired) & "</strong></td>" & vbcrlf
                         response.write "  </tr>" & vbcrlf
                         response.write "</table>" & vbcrlf

                        'Retrieve all of the job/bid postings for this category
                         displayPostings oOrgSubCat("distributionlistid"), oOrgSubCat("distributionlisttype"), lcl_sc_status_id, lcl_sc_show_expired
                      end if

                   else
                      session("lcl_total_postings") = session("lcl_total_postings")
                   end if

                   oOrgSubCat.movenext
                wend

                set oOrgSubCat = nothing
             end if

             oOrgCategories.movenext
          wend

          oOrgCategories.close
          set oOrgCategories = nothing

       end if
   '---------------------------------------------------------------------------
    else  'Search criteria has been entered -----------------------------------
   '---------------------------------------------------------------------------
      'Retrieve all of the categories/sub-categories for this org
   	   sSQL = "SELECT DISTINCT dl.distributionlistid, dl.distributionlistname, dl.distributionlistdescription, "
       sSQL = sSQL & " dl.distributionlistdisplay, dl.orgid, dl.distributionlisttype, parentid "
       sSQL = sSQL & " FROM egov_class_distributionlist dl, egov_jobs_bids jb, egov_distributionlists_jobbids djb "
       sSQL = sSQL & " WHERE dl.orgid = " & iorgid
       sSQL = sSQL & " AND dl.distributionlistid = djb.distributionlistid "
       sSQL = sSQL & " AND djb.posting_id = jb.posting_id "

       if lcl_sc_status_id <> "none" AND lcl_sc_status_id <> "" then
          sSQL = sSQL & " AND jb.status_id = " & lcl_sc_status_id
       end if
       sSQL = sSQL & " AND UPPER(dl.distributionlisttype) = '" & UCASE(lcl_list_type) & "' "

      'Evaluate the search criteria, if it exists
       if lcl_sc_category_id <> "" then
	  q_catid_tmp = track_dbsafe(Replace(Replace(lcl_sc_category_id,"SC",""),"C",""))
	  if isnumeric(q_catid_tmp) then
		q_catid = 0
		on error resume next
			q_catid = clng(q_catid_tmp)
		on error goto 0
           	sSQL = sSQL & " AND dl.distributionlistid = '" & q_catid & "'"
	  end if
       end if

       sSQL = sSQL & " ORDER BY dl.distributionlistname, dl.distributionlistid, dl.distributionlistdescription, "
       sSQL = sSQL & " dl.distributionlistdisplay, dl.orgid, dl.distributionlisttype, dl.parentid "

      	set rs = Server.CreateObject("ADODB.Recordset")
      	rs.Open sSQL, Application("DSN"), 0, 1

       if not rs.eof then
          while not rs.eof
             if rs("distributionlistname") <> "" then
                lcl_distribution_name = rs("distributionlistname")
             else
                lcl_distribution_name = "&nbsp;"
             end if

             if rs("distributionlistdescription") <> "" then
                lcl_distribution_desc = "&nbsp;-&nbsp;" & rs("distributionlistdescription")
             else
                lcl_distribution_desc = ""
             end if

             if rs("parentid") <> "" AND not isnull(rs("parentid")) then
                lcl_category_name = getCategoryName(rs("parentid"),"N") & ": "
             else
                lcl_category_name = ""
             end if

             if rs("distributionlistdisplay") then
                if jobbid_per_dlist_count(rs("distributionlistid"),rs("distributionlisttype"),lcl_sc_category_id,lcl_sc_status_id,lcl_sc_show_expired) > CLng(0) then
                   response.write "<p><table cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#000000"">" & vbcrlf
                   response.write "  <tr align=""left"" bgcolor=""#C0C0C0"">" & vbcrlf
                   response.write "      <td>" & lcl_category_name & "<strong>" & lcl_distribution_name & "</strong>" & lcl_distribution_desc & "</td>" & vbcrlf
                   response.write "      <td align=""center"" width=""60""><strong>" & lcl_list_label & "s: " & jobbid_per_dlist_count(rs("distributionlistid"),rs("distributionlisttype"),lcl_sc_category_id,lcl_sc_status_id,lcl_sc_show_expired) & "</strong></td>" & vbcrlf
                   response.write "  </tr>" & vbcrlf
                   response.write "</table>" & vbcrlf

                  'Retrieve all of the job/bid postings for this category
                   displayPostings rs("distributionlistid"), rs("distributionlisttype"), lcl_sc_status_id, lcl_sc_show_expired
                end if

             else
                session("lcl_total_postings") = session("lcl_total_postings")
             end if

             rs.movenext
          wend
       end if

   '---------------------------------------------------------------------------
    end if
   '---------------------------------------------------------------------------
    if  session("lcl_total_postings") = 0 _
    and jobbid_per_dlist_count("",lcl_list_type,lcl_sc_category_id,lcl_sc_status_id,lcl_sc_show_expired) = 0 then
       response.write "<p><table cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#000000"">" & vbcrlf
       response.write "  <tr align=""left"" bgcolor=""#C0C0C0"">" & vbcrlf
       'response.write "      <td><strong>No " & lcl_list_title & " Available</strong></td>" & vbcrlf
       response.write "      <td><strong>No " & lcl_org_featurename & " Available</strong></td>" & vbcrlf
       response.write "  </tr>" & vbcrlf
       response.write "</table>" & vbcrlf
    end if
  %>
          <p>
      </td>
  </tr>
<%
'  <tr>
'      <td>
'          if jobbid_per_dlist_count("",lcl_list_type,lcl_sc_category_id,lcl_sc_status_id,"Y") > 0 then
'             response.write "<table cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#000000"">" & vbcrlf
'             response.write "  <tr align=""left"" bgcolor=""#C0C0C0"">" & vbcrlf
'             response.write "      <td><strong>EXPIRED " & lcl_list_title & "</strong></td>" & vbcrlf
'             response.write "  </tr>" & vbcrlf
'             response.write "</table>" & vbcrlf

'             displayExpiredPostings lcl_list_type, lcl_sc_status_id
'          end if
'      </td>
'  </tr>
%>
  </form>
</table>
<p>&nbsp;</p>

  </div>
</div>
<!-- #include file="include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub displayPostings(p_dlistid, p_listtype, p_sc_status_id, p_sc_show_expired)

 'Retrieve all of the job/bid postings for the distribution (category) that are:
 ' a. Active = "Y"
 ' b. The Start Date <= Current Date
 ' c. The Current Date <= End Date
  sSQLjb = "SELECT jb.posting_id, jb.jobbid_id, jb.posting_type, jb.title, jb.start_date, jb.end_date, jb.status_id, "
  sSQLjb = sSQLjb & " jb.additional_status_info, jb.active_flag "
  sSQLjb = sSQLjb & " FROM egov_jobs_bids jb, egov_distributionlists_jobbids djb "
  sSQLjb = sSQLjb & " WHERE jb.posting_id = djb.posting_id "
  sSQLjb = sSQLjb & " AND jb.active_flag = 'Y' "
  sSQLjb = sSQLjb & " AND (jb.start_date <= '" & lcl_local_datetime & "' "
  sSQLjb = sSQLjb & " AND DATEDIFF(d,jb.start_date,'1/1/1900') <> 0) "
  sSQLjb = sSQLjb & " AND jb.orgid = " & iorgid
  sSQLjb = sSQLjb & " AND djb.distributionlistid = " & p_dlistid
  sSQLjb = sSQLjb & " AND jb.posting_type = '"       & p_listtype & "' "

 'Evaluate the search criteria
  if p_sc_status_id <> "" AND UCASE(p_sc_status_id) <> "NONE" then
     sSQLjb = sSQLjb & " AND jb.status_id = " & lcl_sc_status_id
  end if

  if p_sc_show_expired = "N" then
     sSQLjb = sSQLjb & " AND ('" & lcl_local_datetime & "' <= jb.end_date "
     sSQLjb = sSQLjb & " OR jb.end_date = '') "
  else
     sSQLjb = sSQLjb & " AND ('" & lcl_local_datetime & "' <= jb.end_date "
     sSQLjb = sSQLjb & " OR jb.end_date = '') "
     sSQLjb = sSQLjb & "UNION ALL "
     sSQLjb = sSQLjb & "SELECT jb.posting_id, jb.jobbid_id, jb.posting_type, jb.title, jb.start_date, jb.end_date, jb.status_id, "
     sSQLjb = sSQLjb & " jb.additional_status_info, jb.active_flag "
     sSQLjb = sSQLjb & " FROM egov_jobs_bids jb, egov_distributionlists_jobbids djb "
     sSQLjb = sSQLjb & " WHERE jb.posting_id = djb.posting_id "
     sSQLjb = sSQLjb & " AND jb.active_flag = 'Y' "
     sSQLjb = sSQLjb & " AND '" & lcl_local_datetime & "' > jb.end_date "
     sSQLjb = sSQLjb & " AND DATEDIFF(d,jb.end_date,'1/1/1900') <> 0 "
     sSQLjb = sSQLjb & " AND jb.end_date <> '' "
     sSQLjb = sSQLjb & " AND jb.orgid = " & iorgid
     sSQLjb = sSQLjb & " AND djb.distributionlistid = " & p_dlistid
     sSQLjb = sSQLjb & " AND jb.posting_type = '"       & p_listtype & "' "

    'Evaluate the search criteria
     if p_sc_status_id <> "" AND UCASE(p_sc_status_id) <> "NONE" then
        sSQLjb = sSQLjb & " AND jb.status_id = " & lcl_sc_status_id
     end if
  end if

  set rsjb = Server.CreateObject("ADODB.Recordset")
  rsjb.Open sSQLjb, Application("DSN"), 0, 1

  if not rsjb.eof then
     lcl_line_count  = 0
     lcl_status_name = ""
     lcl_bgcolor     = "#eeeeee"

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <td>" & vbcrlf

     while not rsjb.eof
        lcl_line_count = lcl_line_count + 1
        'lcl_bgcolor    = changeBGColor(lcl_bgcolor,"","")

        if rsjb("status_id") <> "" then
           'lcl_status_name = getStatusName(rsjb("status_id"))
           sSQLsn = "SELECT status_name "
           sSQLsn = sSQLsn & " FROM egov_statuses "
           sSQLsn = sSQLsn & " WHERE status_id = "  & rsjb("status_id")
           sSQLsn = sSQLsn & " AND status_type = '" & rsjb("posting_type") & "' "
           sSQLsn = sSQLsn & " AND orgid = " & iorgid

           set rssn = Server.CreateObject("ADODB.Recordset")
           rssn.Open sSQLsn, Application("DSN"), 0, 1

           if not rssn.eof then
              lcl_status_name = rssn("status_name")
           else
              lcl_status_name = ""
           end if

           set rssn = nothing
        else
           lcl_status_name = ""
        end if

        if rsjb("end_date") = "" OR isnull(rsjb("end_date")) then
           lcl_close_date = ""
        else
           if CDate(rsjb("end_date")) = CDate("01/01/1900") then
              lcl_close_date = ""
           else
              lcl_close_date = rsjb("end_date")
           end if
        end if

        if rsjb("title") <> "" then
           lcl_title = rsjb("title")
        else
           lcl_title = "[No Title Available - View Details]"
        end if

        response.write "          <table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#ffffff"">" & vbcrlf
        response.write "            <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "                <td width=""75%"">" & vbcrlf
        response.write "                    <a href=""postings_info.asp?posting_id=" & rsjb("posting_id") & "&dlistid=" & p_dlistid & "&listtype=" & rsjb("posting_type") & "&sc_category_id=" & lcl_sc_category_id & "&sc_status_id=" & lcl_sc_status_id & "&sc_show_expired=" & lcl_sc_show_expired & """><font color=""#0000ff"">" & lcl_title & "</font></a><br />" & vbcrlf

       'Get the description of the job/bid posting
        sSQLdesc = "SELECT description FROM egov_jobs_bids WHERE posting_id = " & rsjb("posting_id")
        set rsdesc = Server.CreateObject("ADODB.Recordset")
        rsdesc.Open sSQLdesc, Application("DSN"), 0, 1

        if not rsdesc.eof then
           lcl_description = rsdesc("description")
        else
           lcl_description = ""
        end if

        if lcl_description <> "" then
           if LEN(lcl_description) > 300 then
              lcl_description = LEFT(lcl_description,300) & "... "
           else
              lcl_description = lcl_description & "&nbsp;"
           end if

          'Format the description so that it displays properly
           lcl_description = replace(lcl_description,chr(10),"<br />")

           response.write lcl_description & vbcrlf
           response.write "<a href=""postings_info.asp?posting_id=" & rsjb("posting_id") & "&dlistid=" & p_dlistid & "&listtype=" & rsjb("posting_type") & "&sc_category_id=" & lcl_sc_category_id & "&sc_status_id=" & lcl_sc_status_id & "&sc_show_expired=" & lcl_sc_show_expired & """><font color=""#0000ff"">[more]</font></a>" & vbcrlf
        end if

        response.write "                </td>" & vbcrlf
        response.write "                <td>" & vbcrlf

        if lcl_status_name <> "" then
           response.write "                    <strong>Status: </strong>" & lcl_status_name & "<br />" & vbcrlf
        end if

        if lcl_close_date <> "" then
           lcl_expired_date_total = lcl_local_datetime - lcl_close_date
           if lcl_expired_date_total > 0 then
              response.write "                    <strong>Closed: </strong>" & lcl_close_date & "<br />" & vbcrlf
              response.write "                    <font color=""#ff0000""><strong>EXPIRED</strong></font>" & vbcrlf
           else
              response.write "                    <strong>Closes: </strong>" & lcl_close_date & "<br />" & vbcrlf
           end if
        end if

        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "          </table>" & vbcrlf

        rsjb.movenext
     wend

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

     session("lcl_total_postings") = session("lcl_total_postings") + 1
  else
     session("lcl_total_postings") = session("lcl_total_postings")
  end if

  set rsjb = nothing

end sub

'------------------------------------------------------------------------------
sub displayExpiredPostings(p_listtype, p_sc_status_id)
'Retrieve all of the job/bid postings for the distribution (category) that are:
'  a. Active = "Y"
'  b. The Start Date <= Current Date
'  c. The Current Date > end Date
  sSQLjb = "SELECT distinct jb.posting_id, jb.jobbid_id, jb.posting_type, jb.title, jb.start_date, jb.end_date, jb.status_id, "
  sSQLjb = sSQLjb & " jb.additional_status_info, jb.active_flag "
  sSQLjb = sSQLjb & " FROM egov_jobs_bids jb, egov_distributionlists_jobbids djb "
  sSQLjb = sSQLjb & " WHERE jb.posting_id = djb.posting_id "
  sSQLjb = sSQLjb & " AND jb.active_flag = 'Y' "
'  sSQLjb = sSQLjb & " AND (jb.start_date <= '" & lcl_local_datetime & "' "
'  sSQLjb = sSQLjb & " AND DATEDIFF(d,jb.start_date,'1/1/1900') <> 0) "
  sSQLjb = sSQLjb & " AND ('" & lcl_local_datetime & "' > jb.end_date "
  sSQLjb = sSQLjb & " AND DATEDIFF(d,jb.end_date,'1/1/1900') <> 0) "
  sSQLjb = sSQLjb & " AND jb.orgid = " & iorgid
  sSQLjb = sSQLjb & " AND jb.posting_type = '" & p_listtype & "' "

 'Evaluate the search criteria
  if p_sc_status_id <> "" AND UCASE(p_sc_status_id) <> "NONE" then
     sSQLjb = sSQLjb & " AND jb.status_id = " & lcl_sc_status_id
  end if

  set rsjb = Server.CreateObject("ADODB.Recordset")
  rsjb.Open sSQLjb, Application("DSN"), 0, 1

  if not rsjb.eof then
     lcl_line_count  = 0
     lcl_status_name = ""
     lcl_bgcolor     = "#eeeeee"

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <td>" & vbcrlf

     while not rsjb.eof
        lcl_line_count = lcl_line_count + 1
'        lcl_bgcolor    = changeBGColor(lcl_bgcolor,"","")

        if rsjb("status_id") <> "" then
'           lcl_status_name = getStatusName(rsjb("status_id"))
           sSQLsn = "SELECT status_name "
           sSQLsn = sSQLsn & " FROM egov_statuses "
           sSQLsn = sSQLsn & " WHERE status_id = "  & rsjb("status_id")
           sSQLsn = sSQLsn & " AND status_type = '" & rsjb("posting_type") & "' "
           sSQLsn = sSQLsn & " AND orgid = " & iorgid

           set rssn = Server.CreateObject("ADODB.Recordset")
           rssn.Open sSQLsn, Application("DSN"), 0, 1

           if not rssn.eof then
              lcl_status_name = rssn("status_name")
           else
              lcl_status_name = ""
           end if

           set rssn = nothing
        else
           lcl_status_name = ""
        end if

        if rsjb("end_date") = "" OR isnull(rsjb("end_date")) then
           lcl_close_date = ""
        else
           if CDate(rsjb("end_date")) = CDate("01/01/1900") then
              lcl_close_date = ""
           else
              lcl_close_date = rsjb("end_date")
           end if
        end if

        if rsjb("title") <> "" then
           lcl_title = rsjb("title")
        else
           lcl_title = "[No Title Available - View Details]"
        end if

       'Get the description of the job/bid posting
        sSQLdesc = "SELECT description FROM egov_jobs_bids WHERE posting_id = " & rsjb("posting_id")
        set rsdesc = Server.CreateObject("ADODB.Recordset")
        rsdesc.Open sSQLdesc, Application("DSN"), 0, 1

        if not rsdesc.eof then
           lcl_description = rsdesc("description")
        else
           lcl_description = ""
        end if

        response.write "          <table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""100%"" bgcolor=""#ffffff"">" & vbcrlf
        response.write "            <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "                <td width=""75%"">" & vbcrlf
        response.write "                    <a href=""postings_info.asp?posting_id=" & rsjb("posting_id") & "&dlistid=" & p_dlistid & "&listtype=" & rsjb("posting_type") & "&sc_category_id=" & lcl_sc_category_id & "&sc_status_id=" & lcl_sc_status_id & "&sc_show_expired=" & lcl_sc_show_expired & """><font color=""#0000ff"">" & lcl_title & "</font></a><br />" & vbcrlf

        if lcl_description <> "" then
           if LEN(lcl_description) > 300 then
              lcl_description = LEFT(lcl_description,300) & "... "
           else
              lcl_description = lcl_description & "&nbsp;"
           end if

           response.write lcl_description & "<a href=""postings_info.asp?posting_id=" & rsjb("posting_id") & "&dlistid=" & p_dlistid & "&listtype=" & rsjb("posting_type") & "&sc_category_id=" & lcl_sc_category_id & "&sc_status_id=" & lcl_sc_status_id & "&sc_show_expired=" & lcl_sc_show_expired & """><font color=""#0000ff"">[more]</font></a>" & vbcrlf
        end if

        response.write "                </td>" & vbcrlf
        response.write "                <td>" & vbcrlf

        if lcl_status_name <> "" then
           response.write "                    <strong>Status: </strong>" & lcl_status_name & "<br />" & vbcrlf
        end if

        if lcl_close_date <> "" then
           response.write "                    <strong>Closes: </strong>" & lcl_close_date  & "<br />" & vbcrlf
        end if

        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "          </table>" & vbcrlf

        rsjb.movenext
     wend

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

     session("lcl_total_postings") = session("lcl_total_postings") + 1
  else
     session("lcl_total_postings") = session("lcl_total_postings")
  end if

  set rsjb = nothing

end sub

'------------------------------------------------------------------------------
function jobbid_per_dlist_count(p_dlistid, p_listtype, p_sc_category_id, p_sc_status_id, p_sc_show_expired)
  lcl_return  = 0
  lcl_return2 = 0

 'Setup the query for the count of NON-EXPIRED postings
  sSQLcnt = "SELECT count(jb.posting_id) AS total_postings "
  sSQLcnt = sSQLcnt & " FROM egov_jobs_bids jb, egov_distributionlists_jobbids djb "
  sSQLcnt = sSQLcnt & " WHERE jb.posting_id = djb.posting_id "
  sSQLcnt = sSQLcnt & " AND (jb.start_date <= '" & lcl_local_datetime & "' "
  sSQLcnt = sSQLcnt & " AND DATEDIFF(d,jb.start_date,'1/1/1900') <> 0) "
  sSQLcnt = sSQLcnt & " AND jb.active_flag = 'Y' "
  sSQLcnt = sSQLcnt & " AND jb.orgid = " & iorgid
  sSQLcnt = sSQLcnt & " AND jb.posting_type = '" & p_listtype & "' "
  sSQLcnt = sSQLcnt & " AND ('" & lcl_local_datetime & "' <= jb.end_date "
  sSQLcnt = sSQLcnt & " OR jb.end_date = '') "

  if p_dlistid <> "" then
     sSQLcnt = sSQLcnt & " AND djb.distributionlistid = " & p_dlistid
  end if

 'Evaluate the search criteria if any has been entered.
  if p_sc_category_id <> "" then
     q_catid_tmp = track_dbsafe(Replace(Replace(lcl_sc_category_id,"SC",""),"C",""))
	  if isnumeric(q_catid_tmp) then
		q_catid = 0
		on error resume next
			q_catid = clng(q_catid_tmp)
		on error goto 0
     		sSQLcnt = sSQLcnt & " AND djb.distributionlistid = '" & q_catid & "' "
	end if
  end if

  if p_sc_status_id <> "" AND UCASE(p_sc_status_id) <> "NONE" then
     sSQLcnt = sSQLcnt & " AND jb.status_id = " & p_sc_status_id
  end if

  set rscnt = Server.CreateObject("ADODB.Recordset")
  rscnt.Open sSQLcnt, Application("DSN"), 0, 1

 'Get the total count of NON-EXPIRED postings
  if not rscnt.eof then
     lcl_return = rscnt("total_postings")
  else
     lcl_return = 0
  end if

 '-----------------------------------------------------------------------------
 'Count all of the EXPIRED postings that fit the criteria, if any has been entered
 'Also, ONLY perform this count if the user has selected to DISPLAY the expired postings
 '-----------------------------------------------------------------------------
  if UCASE(p_sc_show_expired) = "Y" then
    'Setup the query for the count of EXPIRED postings
     sSQLcnt2 = "SELECT count(jb.posting_id) AS total_postings "
     sSQLcnt2 = sSQLcnt2 & " FROM egov_jobs_bids jb, egov_distributionlists_jobbids djb "
     sSQLcnt2 = sSQLcnt2 & " WHERE jb.posting_id = djb.posting_id "
     sSQLcnt2 = sSQLcnt2 & " AND (jb.start_date <= '" & lcl_local_datetime & "' "
     sSQLcnt2 = sSQLcnt2 & " AND DATEDIFF(d,jb.start_date,'1/1/1900') <> 0) "
     sSQLcnt2 = sSQLcnt2 & " AND jb.active_flag = 'Y' "
     sSQLcnt2 = sSQLcnt2 & " AND jb.orgid = " & iorgid
     sSQLcnt2 = sSQLcnt2 & " AND jb.posting_type = '" & p_listtype & "' "

     if p_dlistid <> "" then
        sSQLcnt2 = sSQLcnt2 & " AND djb.distributionlistid = " & p_dlistid
     end if

    'Evaluate the search criteria if any has been entered.
     if p_sc_category_id <> "" then
     	q_catid_tmp = track_dbsafe(Replace(Replace(lcl_sc_category_id,"SC",""),"C",""))
	  if isnumeric(q_catid_tmp) then
		q_catid = 0
		on error resume next
			q_catid = clng(q_catid_tmp)
		on error goto 0
     		sSQLcnt2 = sSQLcnt2 & " AND djb.distributionlistid = '" & q_catid & "' "
	end if
     end if

     if p_sc_status_id <> "" AND UCASE(p_sc_status_id) <> "NONE" then
        sSQLcnt2 = sSQLcnt2 & " AND jb.status_id = " & p_sc_status_id
     end if

     sSQLcnt2 = sSQLcnt2 & " AND '" & lcl_local_datetime & "' > jb.end_date "
     sSQLcnt2 = sSQLcnt2 & " AND jb.end_date <> '1/1/1900' "
     sSQLcnt2 = sSQLcnt2 & " AND jb.end_date <> '' "

     set rscnt2 = Server.CreateObject("ADODB.Recordset")
     rscnt2.Open sSQLcnt2, Application("DSN"), 0, 1

    'Get the total count of EXPIRED postings
     if not rscnt2.eof then
        lcl_return2 = rscnt2("total_postings")
     else
        lcl_return2 = 0
     end if
  end if

 'Combine the totals
  lcl_return = lcl_return + lcl_return2

  jobbid_per_dlist_count = CLng(lcl_return)

end function
%>
