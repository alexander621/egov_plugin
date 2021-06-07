 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="../class/classOrganization.asp" //-->
<!-- #include file="../include_top_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: news_info.asp
' AUTHOR:   David Boyer
' CREATED:  04/08/2009
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the News article selected to view
'
' MODIFICATION HISTORY
' 1.0	04/08/2009	David Boyer - Initial Version
' 1.1	06/10/2013	Steve Loar - Added regular expression match to month and year to block intrusion attack attempts
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim re, matches

Set re = New RegExp
re.Pattern = "^\d+$"

'Check to see if the feature is offline
if isFeatureOffline("news_items") = "Y" then
	response.redirect "outage_feature_offline.asp"
end if

If request("id") = "" Then
	response.redirect "news.asp"
End If 

'Set up local variables based on posting (list) type.
lcl_list_label   = "News"
lcl_list_title   = "News"

'Retrieve the "newsitemid"
If request("id") <> "" Then 
	Set matches = re.Execute(request("id"))
	If matches.Count > 0 Then
		lcl_newsitemid = CLng(request("id"))
	Else
		lcl_newsitemid = 0
	End If 
Else 
	lcl_newsitemid = 0
End If 

'Retrieve the news item data.
sSql = "SELECT newsitemid, itemtitle, isnull(publicationstart,itemdate) AS articledate, itemtext, itemlinkurl "
sSql = sSql & " FROM egov_news_items "
sSql = sSql & " WHERE itemdisplay = 1 "
sSql = sSql & " AND UPPER(newstype) = 'NEWS' "
sSql = sSql & " AND orgid = " & iorgid
sSql = sSql & " AND newsitemid = " & lcl_newsitemid

'response.write sSql & "<br /><br />"

set oNewsInfo = Server.CreateObject("ADODB.Recordset")
oNewsInfo.Open sSql, Application("DSN"), 3, 1

if not oNewsInfo.eof then
	lcl_itemtitle   = oNewsInfo("itemtitle")
	lcl_articledate = oNewsInfo("articledate")
	lcl_itemtext    = oNewsInfo("itemtext")
	lcl_itemlinkurl = oNewsInfo("itemlinkurl")

	'Build the link
	if lcl_itemlinkurl <> "" then
		lcl_itemlinkurl = "<a href=""" & lcl_itemlinkurl & """ style=""color:#0000ff"" target=""_blank"">[More Info]</a>" & vbcrlf
	end if
else
	lcl_itemtitle   = ""
	lcl_articledate = ""
	lcl_itemtext    = "Nothing was found in our system that matched this news item."
	lcl_itemlinkurl = ""
end if

oNewsInfo.close
set oNewsInfo = nothing
%>

<html>
<head>
	<meta name="viewport" content="width=device-width, initial-scale=1" />
 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

 	<script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/easyform.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/setfocus.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>

  <script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
  <script type="text/javascript">var addthis_pub="cschappacher";</script>

<style>
  .articleDate {
    font-size:     16px;
    font-family:   verdana;
    color:         #800000;
  }

  .articleTitle {
    font-size:     16px;
    font-family:   verdana;
    border-bottom: 1pt solid #000000;
  }

  .closeButton {
    border-top: 1pt solid #000000;
  }
</style>
</head>
<body bgcolor="#c0c0c0">
<table border="0" cellspacing="0" cellpadding="5" width="100%">
  <tr valign="top">
      <td class="articleDate">
          <%=lcl_articledate%>
      </td>
      <td align="right">
          <table border="0" cellspacing="0" cellpadding="2">
            <tr valign="top">
                <td><% displayYahooBuzzButton iorgid %></td>
                <td><% displayAddThisButton iorgid %></td>
            </tr>
          </table>
      </td>
  </tr>
  <tr valign="top">
      <td colspan="2" class="articleTitle">
          <%=lcl_itemtitle%>
      </td>
  </tr>
  <tr valign="top">
      <td colspan="2" bgcolor="#ffffff">
          <p style="padding-top:5px; padding-bottom:5px;">
          <%=lcl_itemtext%>
          <%=lcl_itemlinkurl%>
          </p>
      </td>
  </tr>
  <tr>
      <td colspan="2" align="center" class="closeButton">
          <!--input type="button" name="closeButton" id="closeButton" value="Close Window" class="button" onclick="parent.close()" /-->
      </td>
  </tr>
</table>
</body>
</html>
