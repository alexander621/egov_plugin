<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rssfeeds_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the rss feeds
'
' MODIFICATION HISTORY
' 1.0 04/08/09 David Boyer - Initial Version
' 1.1 01/07/10 David Boyer - Modified security to now check to see if the user is a "root admin"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("rssfeeds") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 'if not userhaspermission(session("userid"),"rssfeeds_maint") then
 '  	response.redirect sLevel & "permissiondenied.asp"
 'end if

 lcl_pagetitle = "RSS Feeds"

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 else
    lcl_isRootAdmin = False
 end if

'Retrieve the search options
 'lcl_sc_fromcreatedate = ""

 'if request("sc_fromcreatedate") <> "" then
 '   lcl_sc_fromcreatedate = request("sc_fromcreatedate")
 'end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--
function confirm_delete(iFeedID, iTotalItems) {
  var lcl_rssTitle = document.getElementById("rssfeed"+iFeedID).innerHTML;

  if(iTotalItems > 0) {
     lcl_msg  = '"' + lcl_rssTitle + '" cannot be deleted as there are RSS Items associated to it.\n';
     lcl_msg += 'Set the RSS Feed to "inactive".';

     alert(lcl_msg);
  }else{
    	if (confirm("Are you sure you want to delete '" + lcl_rssTitle + "' ?")) { 
  	   			//DELETE HAS BEEN VERIFIED
   		  		location.href='rssfeeds_action.asp?user_action=DELETE&feedid='+ iFeedID;
     }
		}
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
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;">
            <table border="0" cellspacing="0" cellpadding="0" width="1000px">
              <tr>
                  <td><font size="+1"><strong><%=Session("sOrgName")%>&nbsp;<%=lcl_pagetitle%></strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
              </tr>
            </table>
            <br />
          <%
           'If the user is a "root admin" then allow them to add a new RSS Feed.
            if lcl_isRootAdmin then
               response.write "<input type=""button"" name=""newButton"" id=""newButton"" value=""New RSS Feed"" class=""button"" onclick=""window.location='rssfeeds_maint.asp';"" />" & vbcrlf
            end if

            response.write "<p>" & vbcrlf
            listRSSFeeds lcl_isRootAdmin
            response.write "</p>" & vbcrlf
          %>
          </div>
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
sub listRSSFeeds(iIsRootAdmin)
 	Dim iRowCount

 	iRowCount = 0

'  sSQL = "SELECT f.feedid, f.feedname, f.isActive, f.title, f.description , f.feedurl, f.lastbuilddate, t.orgtitle "
'  sSQL = sSQL & " FROM egov_rssfeeds f "
'  sSQL = sSQL &      " LEFT OUTER JOIN egov_rssfeeds_orgtitles t ON f.feedid = t.feedid AND t.orgid = " & session("orgid")
'  sSQL = sSQL & " ORDER BY feedname "

  sSQL = "SELECT rf.feedid, rf.feedname, rf.isActive, rf.title, rf.description , rf.feedurl, rf.lastbuilddate, "
  sSQL = sSQL & " (select orgtitle "
  sSQL = sSQL &  " from egov_rssfeeds_orgtitles t "
  sSQL = sSQL &  " where rf.feedid = t.feedid "
  sSQL = sSQL &  " and t.orgid = " & session("orgid") & ") as orgtitle "
  sSQL = sSQL & " FROM egov_rssfeeds rf, egov_organization_features f, egov_organizations_to_features FO "
  sSQL = sSQL & " WHERE UPPER(rf.feature) = UPPER(f.feature) "
  sSQL = sSQL & " AND f.featureid = fo.featureid "
  sSQL = sSQL & " AND FO.orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY title "

 	set oRSSFeeds = Server.CreateObject("ADODB.Recordset")
	 oRSSFeeds.Open sSQL, Application("DSN"), 3, 1
	
 	if not oRSSFeeds.eof then
   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:800px"">" & vbcrlf
   		response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <th>Title</th>" & vbcrlf
     response.write "      <th>Org Title</th>" & vbcrlf
     response.write "      <th>Feed Name</th>" & vbcrlf
     response.write "      <th>Feed URL</th>" & vbcrlf
     response.write "      <th>Last Build Date</th>" & vbcrlf
     response.write "      <th align=""center"">Total<br />RSS Items</th>" & vbcrlf
     response.write "      <th align=""center"">Active</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     lcl_bgcolor             = "#ffffff"

     do while not oRSSFeeds.eof
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount    = iRowCount + 1

       'Determine if the rss feed is/isn't active
        if oRSSFeeds("isActive") then
           lcl_active = "Y"
        else
           lcl_active = "&nbsp;"
        end if

       'Count total RSS articles per feed.
        lcl_total_rss_items = getTotalRSSItems(oRSSFeeds("feedid"))

       'Setup the javascript events
       'ONLY users that are "root admins" can edit RSS feeds.
        if iIsRootAdmin then
           lcl_row_onmouseover = " onMouseOver=""mouseOverRow(this);"""
           lcl_row_onmouseout  = " onMouseOut=""mouseOutRow(this);"""
           lcl_row_onclick     = " onclick=""location.href='rssfeeds_maint.asp?feedid=" & oRSSFeeds("feedid") & "';"""
        else
           lcl_row_onmouseover = ""
           lcl_row_onmouseout  = ""
           lcl_row_onclick     = ""
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """" & lcl_row_onmouseover & lcl_row_onmouseout & " valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & " nowrap=""nowrap""><span id=""rssfeed" & oRSSFeeds("feedid") & """>" & oRSSFeeds("title") & "</span></td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & " nowrap=""nowrap"">" & oRSSFeeds("orgtitle") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & ">" & oRSSFeeds("feedname")                   & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & ">" & oRSSFeeds("feedurl")                    & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & ">" & oRSSFeeds("lastbuilddate")              & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & " align=""center"">" & lcl_total_rss_items    & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit""" & lcl_row_onclick & " align=""center"">" & lcl_active             & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" align=""center"">" & vbcrlf

        if lcl_total_rss_items > 0 then
           response.write "&nbsp;" & vbcrlf
        else
           response.write "          <input type=""button"" name=""delete" & iRowCount & """ id=""delete" & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oRSSFeeds("feedid") & "'," & lcl_total_rss_items & ");"" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

       oRSSFeeds.movenext
   loop

 		response.write "</table>" & vbcrlf
	  response.write "</div>" & vbcrlf
   response.write "<div style=""text-align:right;"">Total RSS Feeds: [" & iRowCount & "]</div>" & vbcrlf

 else
  		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No RSS Feeds have been created.</p>" & vbcrlf
	end if

	oRSSFeeds.close
	set oRSSFeeds = nothing 

end sub

'------------------------------------------------------------------------------
function getTotalRSSItems(iFeedID)

  lcl_return = 0

  if iFeedID <> "" then
     sSQL = "SELECT count(rssid) AS total_rss "
     sSQL = sSQL & " FROM egov_rss "
     sSQL = sSQL & " WHERE orgid = " & session("orgid")
     sSQL = sSQL & " AND feedid = "  & iFeedID

    	set oRSSCnt = Server.CreateObject("ADODB.Recordset")
   	 oRSSCnt.Open sSQL, Application("DSN"), 3, 1

     if not oRSSCnt.eof then
        lcl_return = oRSSCnt("total_rss")
     end if

     oRSSCnt.close
     set oRSSCnt = nothing

  end if

  getTotalRSSItems = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "RSS Feed does not exist..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>