<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="postings_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: postings_info.asp
' AUTHOR:   David Boyer
' CREATED:  01/04/2008
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays detail information for job/bid posting
'
' MODIFICATION HISTORY
' 1.0  02/06/08	David Boyer - Initial Version
' 1.1  10/14/08 David Boyer - Added "Requires Login" check to "Download Available" field
' 1.2  05/20/09 David Boyer - Added "Click Counter" to "Download Available" links.
' 1.3  08/17/09 David Boyer - Added check for required fields.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim orgHasWordPress

'BEGIN: Check for apotrophes --------------------------------------------------
 lcl_dlistid         = ""
 lcl_listtype        = ""
 lcl_posting_id      = ""
 lcl_sc_category_id  = ""
 lcl_sc_status_id    = ""
 lcl_sc_show_expired = ""

'Check required fields
 if request("listtype") <> "" then
    if not containsApostrophe(request("listtype")) then
       lcl_list_type = request("listtype")
    end if
 end if

 if lcl_list_type = "" then
    response.redirect sEgovWebsiteURL
 end if

 if request("dlistid") <> "" then
    if not containsApostrophe(request("dlistid")) then
       lcl_dlistid = request("dlistid")
    end if
 end if

 if lcl_dlistid = "" then
    response.redirect sEgovWebsiteURL
 end if

lcl_posting_id = ""
 if request("posting_id") <> "" then
    if not containsApostrophe(request("posting_id")) then
	    on error resume next
       lcl_posting_id = clng(request("posting_id"))
       	on error goto 0
    end if
 end if

 if lcl_posting_id = "" then
    response.redirect sEgovWebsiteURL
 end if

'Retrieve the search criteria parameters
 if request("sc_category_id") <> "" then
    if not containsApostrophe(request("sc_category_id")) then
       lcl_sc_category_id = request("sc_category_id")
    end if
 end if

 if request("sc_status_id") <> "" then
    if not containsApostrophe(request("sc_status_id")) then
       lcl_sc_status_id = request("sc_status_id")
    end if
 end if

 if request("sc_show_expired") <> "" then
    if not containsApostrophe(request("sc_show_expired")) then
       lcl_sc_show_expired = request("sc_show_expired")
    end if
 end if
'END: Check for apotrophes ----------------------------------------------------

'To help prevent hacks.
 'if NOT isnumeric(lcl_posting_id) then
    'response.redirect "postings.asp?listtype=" & lcl_list_type
 'end if

'Check to see if the feature is offline
 if isFeatureOffline("job_postings,bid_postings") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

'Retrieve the org_group_id of the organization group that is to be maintained.
'If no value exists then redirect them back to the main results screen
 'if lcl_posting_id <> "" then
    'lcl_posting_id = CLng(lcl_posting_id)
 'end if

 Dim oPostingsOrg
 Set oPostingsOrg = New classOrganization

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
 elseif lcl_list_type = "BID" then
    lcl_list_label = "Bid"
    'lcl_list_title = "Bid Postings"
 end if

'Set up org features.
 lcl_orghasfeature_bidpostings_requirepubliclogin            = orghasfeature(iorgid,lcase(lcl_list_type)&"postings_requirepubliclogin")
 lcl_orghasfeature_bidpostings_upload_userbids               = orghasfeature(iorgid,lcase(lcl_list_type)&"postings_upload_userbids")
 lcl_orghasfeature_clickcounter_postings                     = orghasfeature(iorgid,"clickcounter_postings")
 lcl_orghasfeature_bidpostings_viewplanholders_requirefields = orghasfeature(iorgid,"bidpostings_viewplanholders_requirefields")
 orgHasWordPress                                             = orghasfeature( iorgid, "wordpress public interface" )

'Determine if the user has subscribed to this categoryid
 lcl_isCategoryAssigned = isCategoryAssigned(request.cookies("userid"),lcl_dlistid)

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

'Setup the session variable in case the user needs to log in.
 session("RedirectPage") = request.servervariables("script_name") & "?" & request.querystring()

'Check to see if the user has the Business Name and Work Phone populated on their account.
'If "yes" then allow them to click on the "download available" link(s) and/or "Submit Bid" button
'If "no" then show an error message.
 lcl_userbusinessname     = ""
 lcl_userworkphone        = ""
 lcl_hasAllRequiredFields = "N"

 if request.cookies("userid") <> "" then
    sSQL = "SELECT userbusinessname, userworkphone "
    sSQL = sSQL & " FROM egov_users "
    sSQL = sSQL & " WHERE userid = " & request.cookies("userid")
    sSQL = sSQL & " AND orgid = "    & iOrgID

    set oGetUser_RequiredFields = Server.CreateObject("ADODB.Recordset")
    oGetUser_RequiredFields.Open sSQL, Application("DSN"), 3, 1

    if not oGetUser_RequiredFields.eof then
       lcl_userbusinessname = trim(oGetUser_RequiredFields("userbusinessname"))
       lcl_userworkphone    = trim(oGetUser_RequiredFields("userworkphone"))
    end if

    oGetUser_RequiredFields.close
    set oGetUser_RequiredFields = nothing

   'Now verify that BOTH fields are populated.
    if lcl_userbusinessname <> "" AND lcl_userworkphone <> "" then
       lcl_hasAllRequiredFields = "Y"
    end if
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

<style>
  .posting_table          { border-top: solid #000000 1px;border-left: solid #000000 1px;border-right: solid #000000 1px;border-bottom: solid #000000 1px; }
  .posting_header         { background-color:#9c192f;font-size:12px;color: #ffffff }
  .posting_title          { font-size:16px;color: #9c192f }
  .posting_text_highlight { color: #9c192f }
  .posting_fieldset {
     width:  800px;
     border: 1pt solid #808080;
     -webkit-border-radius: 5px;
     -moz-border-radius:    5px;
  }
</style>

<script language="javascript">
function submitUserBid() {
//function fOpenWin(page,p_wintype,p_width,p_height) {
  var lcl_wintype  = "new";
  var lcl_width    = 700;
  var lcl_height   = 200;
  var lcl_left_pos = (screen.availWidth/2) - (lcl_width/2);
  var lcl_top_pos  = (screen.availHeight/2) - (lcl_height/2);

  OpenWin = window.open("postings_submit_userbids.asp?posting_id=<%=lcl_posting_id%>&listtype=<%=lcl_list_type%>&dlistid=<%=lcl_dlistid%>", lcl_wintype, "status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes,width="+lcl_width+",height="+lcl_height+",left="+lcl_left_pos+",top="+lcl_top_pos);
  if (document.images) {OpenWin.focus();}
}

function viewPlanHolders() {
//function fOpenWin(page,p_wintype,p_width,p_height) {
  var lcl_wintype  = "new";
  var lcl_width    = 950;
  var lcl_height   = 400;
  var lcl_left_pos = (screen.availWidth/2) - (lcl_width/2);
  var lcl_top_pos  = (screen.availHeight/2) - (lcl_height/2);

  OpenWin = window.open("view_planholders.asp?posting_id=<%=lcl_posting_id%>&listtype=<%=lcl_list_type%>&dlistid=<%=lcl_dlistid%>", lcl_wintype, "status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes,width="+lcl_width+",height="+lcl_height+",left="+lcl_left_pos+",top="+lcl_top_pos);
  if (document.images) {OpenWin.focus();}
}

<% if lcl_orghasfeature_clickcounter_postings then %>
function countClick(p_fieldid) {
  //Build the parameter string
		var sParameter = 'isAjaxRoutine=Y';
  sParameter    += '&orgid='      + encodeURIComponent('<%=iorgid%>');
  sParameter    += '&userid='     + encodeURIComponent('<%=request.cookies("userid")%>');
  sParameter    += '&posting_id=' + encodeURIComponent(document.getElementById("posting_id").value);

  if(p_fieldid != "") {
     sParameter += '&linkID='   + encodeURIComponent(p_fieldid);
     sParameter += '&linkText=' + encodeURIComponent(document.getElementById(p_fieldid).innerHTML);
     sParameter += '&linkURL='  + encodeURIComponent(document.getElementById(p_fieldid).href);
  }

  //doAjax('clickcounter/updateclickcounter_postings.asp', sParameter, 'displayScreenMsg', 'post', '0');
doAjax('clickcounter/updateclickcounter_postings.asp', sParameter, 'displayScreenMsg', 'post', '0');
}
<% end if %>

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "&nbsp;";
}
</script>
</head>
<!-- <body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"> -->
<!--#include file="include_top.asp"-->

<%
  oPostingsOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
  response.write "<br />" & vbcrlf

  RegisteredUserDisplay("")

 'If the user HAS signed in then set the link to their manage subscriptions account page.
 'If the user is NOT signed in then set the link to the public subscriptions page.
  if request.cookies("userid") <> "" AND request.cookies("userid") <> "-1" then
     lcl_subscriptions_url = "manage_mail_lists.asp"
  else
      If orgHasWordPress Then  
        lcl_subscriptions_url = getOrganization_WP_URL( iorgid, "wp_subscriptions_url" ) & "#subscriptionlist/" & lcl_list_type
      Else 
        lcl_subscriptions_url = "subscriptions/subscribe.asp?listtype=" & lcl_list_type
      End If 
  end if
%>

<div id="content">
  <div id="centercontent">
<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;" class="start">
  <form name="postings_info" method="post" action="postings_info.asp">
    <input type="hidden" name="posting_id" id="posting_id" value="<%=lcl_posting_id%>" size="5" maxlength="5" />
    <input type="hidden" name="orgid" id="orgid" value="<%=lcl_orgid%>" size="4" maxlength="10" />
  <tr>
      <td>
          <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr><td colspan="2" align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;">&nbsp;</span></td></tr>
            <tr valign="top">
                <td><a href="postings.asp?listtype=<%=lcl_list_type%>&sc_category_id=<%=lcl_sc_category_id%>&sc_status_id=<%=lcl_sc_status_id%>&sc_show_expired=<%=lcl_sc_show_expired%>">Return to <%=lcl_feature_name%></a></td>
                <td align="right">
                    <table border="0" cellspacing="0" cellpadding="2">
                      <tr><td>[<a href="<%=lcl_subscriptions_url%>">Want to be notified about <strong>New <%=lcl_feature_name%></strong>...</a>]</td></tr>
                   <%
                      if lcl_list_type = "JOB" then
                        'Check to see if the posting is related to an Action Line Request
                         sSQLr = "SELECT public_apply_for_position_actionline FROM egov_jobs_bids WHERE posting_id = '" & CLng(lcl_posting_id) & "'"
                         set rsr = Server.CreateObject("ADODB.Recordset")
                         rsr.Open sSQLr, Application("DSN"), 0, 1

                         if not rsr.eof then
                            if rsr("public_apply_for_position_actionline") <> "" then
                               lcl_public_apply_for_position_actionline = rsr("public_apply_for_position_actionline")

                               response.write "<tr align=""center"">" & vbcrlf
                               response.write "    <td>" & vbcrlf
                               If orgHasWordPress Then
                                  response.write "[<a href=""" & getOrganization_WP_URL( iorgid, "wp_actionline_url" ) & "#form/" & lcl_public_apply_for_position_actionline & """>" & vbcrlf
                               Else
                                  response.write "[<a href=""action.asp?actionid=" & lcl_public_apply_for_position_actionline & """>" & vbcrlf
                               End If
                               response.write "        <strong>Apply for Position</strong>...</a>]" & vbcrlf
                               response.write "    </td>" & vbcrlf
                               response.write "</tr>" & vbcrlf

                            else
                               lcl_apply_url = ""

                              'Get the postings_email set up for the ORG
                               sSQLe = "SELECT postings_email, defaultemail FROM organizations WHERE orgid = " & iorgid
                               set rse = Server.CreateObject("ADODB.Recordset")
                               rse.Open sSQLe, Application("DSN"), 0, 1

                               if not rse.eof and iorgid <> 152 then
                                 'First check for the postings_email.  If it is blank then use the org default_email
                                  if rse("postings_email") = "" OR isnull(rse("postings_email")) then
                                     lcl_email = rse("defaultemail")
                                  else
                                     lcl_email = rse("postings_email")
                                  end if

                                  response.write "<tr align=""center"">" & vbcrlf
                                  response.write "    <td>" & vbcrlf
                                  response.write "        [<a href=""mailto:" & lcl_email & "?subject=" & getPostingTitle(lcl_posting_id) & " - " & lcl_feature_name & " Resume"">" & vbcrlf
                                  response.write "        <strong>Email your resume</strong>...</a>]" & vbcrlf
                                  response.write "    </td>" & vbcrlf
                                  response.write "</tr>" & vbcrlf

                               else
                                  lcl_email = ""
                               end if
                            end if
                         else
                            lcl_email = ""
                         end if
                      end if
                   %>
                    </table>
                </td>
            </tr>
          </table>
      </td>
  </tr>
  <tr valign="top">
      <td>
          <%
            'Retrieve all of the job/bid posting data
             sSQL = "SELECT posting_id, jobbid_id, posting_type, title, isnull(start_date,'1/1/1900') as start_date, isnull(end_date,'1/1/1900') as end_date, status_id, "
             sSQL = sSQL & " additional_status_info, description, qualifications, special_requirements, "
             sSQL = sSQL & " misc_info, active_flag, job_salary, bid_publication_info, bid_submittal_info, "
             sSQL = sSQL & " bid_opening_info, bid_recipient, bid_addendum_date, bid_pre_bid_meeting, "
             sSQL = sSQL & " bid_contact_person, download_available, bid_fee, bid_plan_spec_available, "
             sSQL = sSQL & " bid_business_hours, bid_fax_number, bid_plan_holders "
             sSQL = sSQL & " FROM egov_jobs_bids "
             sSQL = sSQL & " WHERE posting_id = '" & CLng(lcl_posting_id) & "'"
             'response.write sSQL & "<br /><br />"

             set rs = Server.CreateObject("ADODB.Recordset")
             rs.Open sSQL, Application("DSN"), 3, 1

             if not rs.eof then
                lcl_posting_id              = rs("posting_id")
                lcl_jobbid_id               = rs("jobbid_id")
                lcl_posting_type            = rs("posting_type")
                lcl_title                   = rs("title")
                lcl_start_date              = replace(rs("start_date"),"1/1/1900","")
                lcl_end_date                = replace(rs("end_date"),"1/1/1900","")
                lcl_status_id               = rs("status_id")
                lcl_additional_status_info  = rs("additional_status_info")
                lcl_description             = rs("description")
                lcl_qualifications          = rs("qualifications")
                lcl_special_requirements    = rs("special_requirements")
                lcl_misc_info               = rs("misc_info")
                lcl_active_flag             = rs("active_flag")
                lcl_job_salary              = rs("job_salary")
                lcl_bid_publication_info    = rs("bid_publication_info")
                lcl_bid_submittal_info      = rs("bid_submittal_info")
                lcl_bid_opening_info        = rs("bid_opening_info")
                lcl_bid_recipient           = rs("bid_recipient")
                if lcl_bid_addendum_date <> "" then
                   lcl_bid_addendum_date    = replace(rs("bid_addendum_date"),"1/1/1900","")
                else
                   lcl_bid_addendum_date    = ""
                end if
                lcl_bid_pre_bid_meeting     = rs("bid_pre_bid_meeting")
                lcl_bid_contact_person      = rs("bid_contact_person")
                lcl_download_available      = rs("download_available")
                lcl_bid_fee                 = rs("bid_fee")
                lcl_bid_plan_spec_available = rs("bid_plan_spec_available")
                lcl_bid_business_hours      = rs("bid_business_hours")
                lcl_bid_fax_number          = rs("bid_fax_number")
                lcl_bid_plan_holders        = rs("bid_plan_holders")

               'BEGIN: Format textarea fields so they display properly --------
                if lcl_description <> "" then
                   lcl_description = replace(rs("description"),chr(10),"<br />")
                end if

                if lcl_qualifications <> "" then
                   lcl_qualifications = replace(rs("qualifications"),chr(10),"<br />")
                end if

                if lcl_special_requirements <> "" then
                   lcl_special_requirements = replace(rs("special_requirements"),chr(10),"<br />")
                end if

                if lcl_misc_info <> "" then
                   lcl_misc_info = replace(rs("misc_info"),chr(10),"<br />")
                end if

                if lcl_job_salary <> "" then
                   lcl_job_salary = replace(rs("job_salary"),chr(10),"<br />")
                end if

                if lcl_bid_contact_person <> "" then
                   lcl_bid_contact_person = replace(rs("bid_contact_person"),chr(10),"<br />")
                end if

                if lcl_download_available <> "" then
                   lcl_download_available = replace(rs("download_available"),chr(10),"<br />")
                end if

                if lcl_bid_plan_spec_available <> "" then
                   lcl_bid_plan_spec_available = replace(rs("bid_plan_spec_available"),chr(10),"<br />")
                end if

                if lcl_bid_plan_holders <> "" then
                   lcl_bid_plan_holders = replace(rs("bid_plan_holders"),chr(10),"<br />")
                end if
               'END: Format textarea fields so they display properly ----------

                lcl_status_name = getStatusName(lcl_status_id,lcl_posting_type)

                if lcl_dlistid = "" then
                   lcl_category_name = "EXPIRED"
                else
                   lcl_category_name = getCategoryName(lcl_dlistid,"Y")
                end if
             else
                response.redirect "postings.asp?listtype=" & lcl_list_type
             end if

             if lcl_title <> "" then
                lcl_title = lcl_title
             else
                lcl_title = "[No Title Available]"
             end if
          %>
          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <tr>
                <td>
                    <fieldset class="posting_fieldset">
                      <legend class="posting_title"><%=lcl_title%></legend><p>
                      <table border="0" cellspacing="0" cellpadding="2" width="100%">
                        <tr valign="top">
                            <td>
                                <p>
                                <table border="0" cellspacing="0" cellpadding="2">
                                  <tr id="jobbid_id" style="display: none">
                                      <td width="150"><strong><%=lcl_list_label%> Number:</strong></td>
                                      <td><%=lcl_jobbid_id%></td>
                                  </tr>
                                  <tr>
                                      <td width="150"><strong>Category:</strong></td>
                                      <td><%=lcl_category_name%></td>
                                  </tr>
                                  <tr>
                                      <td width="150"><strong>Status:</strong></td>
                                      <td><%=lcl_status_name%></td>
                                  </tr>
                                  <tr id="additional_status_info" style="display: none" valign="top">
                                      <td width="150"><strong>Additional Status Information:</strong></td>
                                      <td><%=lcl_additional_status_info%></td>
                                  </tr>
                                  <tr id="job_salary" style="display: none">
                                      <td width="150"><strong>Salary:</strong></td>
                                      <td><%=lcl_job_salary%></td>
                                  </tr>
                                  <tr id="bid_recipient" style="display: none">
                                      <td width="150"><strong>Bid Recipient:</strong></td>
                                      <td><%=lcl_bid_recipient%></td>
                                  </tr>
                                </table>
                                </p>
                            </td>
                            <td align="center">
                            <%
                              lcl_show_submitbid_btn = "N"

                             'Determine if the org allows citizens to upload their bids
                             'If "yes", check to see if the org requires the user to be logged in.
                             'If "yes", check to see if the user has logged in.
                             'If "yes", check to see if the user logged in has subscribed/registered to the category of the posting.
                             'If "yes", if the posting's close date is > current date or if the closing date is NULL
                             'If "yes", if all required fields on the user's profile have been populated (business name and work phone
                             'If "yes", show the "Submit Bid" button.
                              if lcl_orghasfeature_bidpostings_upload_userbids then
                                 if  lcl_orghasfeature_bidpostings_requirepubliclogin _
                                 AND request.cookies("userid") <> "" _
                                 AND lcl_isCategoryAssigned _
                                 AND lcl_hasAllRequiredFields = "Y" then
                                     if lcl_end_date = "" then
                                        lcl_show_submitbid_btn = "Y"
                                     else
                                        if datediff("s",lcl_local_datetime,lcl_end_date) > 0 then
                                           lcl_show_submitbid_btn = "Y"
                                        end if
                                     end if
                                 end if
                              end if

                             'Check to see if we show the Submit Bid button
                              if lcl_show_submitbid_btn = "Y" then
                                 response.write "<input type=""button"" name=""submituserbid"" value=""Submit Bid"" class=""button"" onclick=""submitUserBid()"" />" & vbcrlf
                              end if
                            %>
                            </td>
                        </tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
          </table>
          <table border="0" cellspacing="0" cellpadding="2" width="100%">
            <tr valign="top" id="description" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Description:&nbsp;</strong></legend><p>
                      <%=lcl_description%><p>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="start_date" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Publication Date/Time:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_start_date%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_publication_info" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Publication Information:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_publication_info%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="end_date" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Closing Date/Time:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_end_date%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_submittal_info" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Submittal Information:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_submittal_info%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_opening_info" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Bid Opening Information:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_opening_info%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_addendum_date" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Addendum Date/Time:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_addendum_date%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_pre_bid_meeting" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Pre-Bid Meeting:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_pre_bid_meeting%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_contact_person" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Contact Person:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_contact_person%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="download_available" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Download Available:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr>
                            <td class="posting_text_highlight">
                            <%
                             'Determine if:
                             ' 1. the org requires user to be logged on before viewing posting download(s)
                             ' 2. if "yes" then check to see if the user IS logged in
                             ' 3. if "yes" then see if the user has registered for the category on the posting.
                             ' 4. if "yes" then if all required fields have been populated on the user's profile (business name and work phone)
                              if lcl_orghasfeature_bidpostings_requirepubliclogin then
                                 if request.cookies("userid") = "" then
                                    displaySignInLinks(lcl_list_label)
                                 else
                                    if lcl_isCategoryAssigned then
                                       if lcl_orghasfeature_bidpostings_viewplanholders_requirefields then
                                          if lcl_hasAllRequiredFields = "Y" then
                                             response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""750"">" & vbcrlf
                                             response.write "  <tr valign=""top"">" & vbcrlf
                                             response.write "      <td width=""90%"">" & lcl_download_available & "</td>" & vbcrlf
                                             response.write "      <td align=""right"">" & vbcrlf
                                             response.write "          <input type=""button"" name=""viewPlanHoldersButton"" id=""viewPlanHoldersButton"" value=""View Plan Holders"" class=""button"" onclick=""viewPlanHolders();"" />" & vbcrlf
                                             response.write "      </td>" & vbcrlf
                                             response.write "  </tr>" & vbcrlf
                                             response.write "</table>" & vbcrlf
                                          else
                                             displayPostingsRequiredFields lcl_posting_id, lcl_list_type, lcl_dlistid, lcl_hasAllRequiredFields, lcl_userbusinessname, lcl_userworkphone
                                          end if
                                       else
                                          response.write lcl_download_available & vbcrlf
                                       end if
                                    else
                                       displaySubscribeLink lcl_dlistid
                                       if lcl_hasAllRequiredFields <> "Y" then
                                          displayPostingsRequiredFields lcl_posting_id, lcl_list_type, lcl_dlistid, lcl_hasAllRequiredFields, lcl_userbusinessname, lcl_userworkphone
                                       end if
                                    end if
                                 end if
                              else
                                 response.write lcl_download_available & vbcrlf
                              end if
                            %>
                            </td>
                        </tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_fee" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Fee:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_fee%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_plan_spec_available" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Plan and Spec Available:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_plan_spec_available%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_business_hours" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Business Hours:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_business_hours%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_fax_number" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Fax Number:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_fax_number%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="bid_plan_holders" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Plan Holders List:&nbsp;</strong></legend>
                      <table border="0" cellspacing="0" cellpadding="5">
                        <tr><td class="posting_text_highlight"><%=lcl_bid_plan_holders%></td></tr>
                      </table>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="qualifications" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Qualifications:&nbsp;</strong></legend><p>
                      <%=lcl_qualifications%><p>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="special_requirements" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Special Requirements:&nbsp;</strong></legend><p>
                      <%=lcl_special_requirements%><p>
                    </fieldset>
                </td>
            </tr>
            <tr valign="top" id="misc_info" style="display: none">
                <td>
                    <fieldset class="posting_fieldset">
                      <legend><strong>Miscellaneous:&nbsp;</strong></legend><p>
                      <%=lcl_misc_info%><p>
                    </fieldset>
                </td>
            </tr>
          </table>
      </td>
  </tr>
  <script language="javascript">
  <%
    response.write "  //Common Fields - Display only fields that are populated." & vbcrlf

    if lcl_additional_status_info <> "" then
       response.write "  document.getElementById(""additional_status_info"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_jobbid_id <> "" then
       response.write "  document.getElementById(""jobbid_id"").style.display="""";" & vbcrlf
    end if

    if lcl_start_date <> "" then
       response.write "  document.getElementById(""start_date"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_end_date <> "" then
       response.write "  document.getElementById(""end_date"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_description <> "" then
       response.write "  document.getElementById(""description"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_qualifications <> "" then
       response.write "  document.getElementById(""qualifications"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_special_requirements <> "" then
       response.write "  document.getElementById(""special_requirements"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_misc_info <> "" then
       response.write "  document.getElementById(""misc_info"").style.display=""table-row"";" & vbcrlf
    end if

    if lcl_download_available <> "" then
       response.write "  document.getElementById(""download_available"").style.display=""table-row"";" & vbcrlf
    end if

    response.write "  //Display the correct fields depending on the posting type." & vbcrlf
    response.write "  //Also, only display fields if there is a value in it." & vbcrlf

    if lcl_list_type = "JOB" then
       response.write "  //-- JOB -----------------------------------------------------" & vbcrlf

       if lcl_job_salary <> "" then
          response.write "  document.getElementById(""job_salary"").style.display="""";" & vbcrlf
       end if

    elseif lcl_list_type = "BID" then
       response.write "  //-- BID -----------------------------------------------------" & vbcrlf

       if lcl_bid_recipient <> "" then
          response.write "  document.getElementById(""bid_recipient"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_publication_info <> "" then
          response.write "  document.getElementById(""bid_publication_info"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_submittal_info <> "" then
          response.write "  document.getElementById(""bid_submittal_info"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_opening_info <> "" then
          response.write "  document.getElementById(""bid_opening_info"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_addendum_date <> "" then
          response.write "  document.getElementById(""bid_addendum_date"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_pre_bid_meeting <> "" then
          response.write "  document.getElementById(""bid_pre_bid_meeting"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_contact_person <> "" then
          response.write "  document.getElementById(""bid_contact_person"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_fee <> "" then
          response.write "  document.getElementById(""bid_fee"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_plan_spec_available <> "" then
          response.write "  document.getElementById(""bid_plan_spec_available"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_business_hours <> "" then
          response.write "  document.getElementById(""bid_business_hours"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_fax_number <> "" then
          response.write "  document.getElementById(""bid_fax_number"").style.display=""table-row"";" & vbcrlf
       end if

       if lcl_bid_plan_holders <> "" then
          response.write "  document.getElementById(""bid_plan_holders"").style.display=""table-row"";" & vbcrlf
       end if

    end if
  %>
  </script>
  </form>
</table>
  </div>
</div>
<!-- #include file="include_bottom.asp" -->
