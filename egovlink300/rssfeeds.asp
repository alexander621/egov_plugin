<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="rss/rss_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rssfeeds.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module display all of the RSS Feeds available
'
' MODIFICATION HISTORY
' 1.0 03/25/09	David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

'Verify the org has the feature assigned
 if not orghasfeature(iorgid,"rssfeeds") then
    response.redirect sEgovWebsiteURL
 end if

 dim oRSSOrg
 set oRSSOrg = New classOrganization
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

 	<title>E-Gov Services - <%=sOrgName%> RSS Feeds</title>

	 <link rel="stylesheet" type="text/css" href="css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="global.css" />
	 <link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	 <script language="javascript" src="scripts/modules.js"></script>
	 <script language="javascript" src="scripts/easyform.js"></script> 

  <script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
  <script type="text/javascript">var addthis_pub="cschappacher";</script>

<style>
  .rssTable {
    border: 1pt solid #000000;
  }

  .rssTableHeaders {
    background-color: #c0c0c0;
    border-bottom:    1pt solid #000000;
  }
</style>
</head>

<!--#Include file="include_top.asp"-->

<p>
<table border="0" cellspacing="0" cellpadding="0" style="max-width:800px;">
  <tr>
      <td>
        <%
          lcl_org_name        = oRSSOrg.GetOrgName()
          lcl_org_state       = oRSSOrg.GetState()
          lcl_org_featurename = "RSS Feeds"

          oRSSOrg.buildWelcomeMessage iorgid, lcl_orghasdisplay_action_page_title, lcl_org_name, lcl_org_state, lcl_org_featurename
        %>
          <!--<font class="pagetitle">Welcome to the <%'sOrgName%> RSS Feeds</font>-->
      </td>
      <td align="right">
          <% displayAddThisButton iorgid %>
      </td>
  </tr>
</table>
<% RegisteredUserDisplay( "" ) %>
</p>
<%
 'BEGIN: Display Page Content --------------------------------------------------
 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""max-width:800px;"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td align=""left"">" & vbcrlf
  response.write "          <p>" & vbcrlf
  response.write "          Stay up-to-date on the latest news with RSS (Really Simple Syndication).  RSS feeds "
  response.write "          provide a way to get the most current information delivered straight to you!  Subscribe to any/all of "
  response.write "          the RSS feeds below, the service is FREE!<br />"
  response.write "          </p>" & vbcrlf
  response.write "          <p>" & vbcrlf
  response.write "                 <strong>How to view one of the RSS feeds below:</strong>" & vbcrlf
  response.write "                 <p>Copy a URL below and paste it into your preferred RSS reader.</p>" & vbcrlf
  response.write "          </p>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
                            displayRSSFeeds
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "<p><br />&nbsp;<br />&nbsp;</p>" & vbcrlf
 'END: Display Page Content ---------------------------------------------------

  response.write "<p>&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;<br />&nbsp;</p>" & vbcrlf
%>
<!--#Include file="include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
sub displayRSSFeeds()

 'Retrieve all of the active RSS Feeds that have at least a single item in them.
  sSQL = "SELECT f.feedid, isnull(t.orgtitle,f.title) AS title, f.feedurl, f.isActive, f.feature "
  sSQL = sSQL & " FROM egov_rssfeeds f "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_rssfeeds_orgtitles t ON f.feedid = t.feedid AND t.orgid = " & iorgid
  sSQL = sSQL & " WHERE f.isActive = 1 "
  sSQL = sSQL & " AND (select count(rssid) "
  sSQL = sSQL &      " from egov_rss r "
  sSQL = sSQL &      " where r.orgid = " & iorgid
  sSQL = sSQL &      " and r.feedid = f.feedid) >= 1 "
  sSQL = sSQL & " ORDER BY f.feedname "

	 set oRSSFeeds = Server.CreateObject("ADODB.Recordset")
	 oRSSFeeds.Open sSQL, Application("DSN"), 3, 1

  if not oRSSFeeds.eof then
     response.write "<div class=""rssContainer"">" & vbcrlf
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""100%"" class=""rssTable liquidtable"">" & vbcrlf
     response.write "  <thead>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""rssTableHeaders"">RSS Feeds</td>" & vbcrlf
     response.write "      <td class=""rssTableHeaders"">Copy URL to RSS Reader</td>" & vbcrlf
     response.write "      <td class=""rssTableHeaders"" align=""center"">View Feed</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  </thead>" & vbcrlf

     lcl_bgcolor = "#eeeeee"

     do while not oRSSFeeds.eof
        if orghasfeature(iorgid,oRSSFeeds("feature")) then
           lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
           lcl_feedurl = sEgovWebsiteURL & oRSSFeeds("feedurl")

           response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
           response.write "      <td class=""repeatheaders"">RSS Feed</td>" & vbcrlf
           response.write "      <td>" & oRSSFeeds("title") & "</td>" & vbcrlf
           response.write "      <td class=""repeatheaders"">Copy URL to RSS Reader</td>" & vbcrlf
           response.write "      <td><a href=""" & lcl_feedurl & """ target=""_blank"">" & lcl_feedurl & "</a></td>" & vbcrlf
           response.write "      <td class=""repeatheaders"">View Feed</td>" & vbcrlf
           response.write "      <td align=""center"">" & vbcrlf
           response.write "          <a href=""" & lcl_feedurl & """ target=""_blank""><img src=""images/icon_rss.jpg"" border=""0"" width=""16"" height=""16"" alt=""Click to view rss feed"" /></a>" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "  </tr>" & vbcrlf
        end if

        oRSSFeeds.movenext
     loop

     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf

  end if

  oRSSFeeds.close
  set oRSSFeeds = nothing

end sub
%>
