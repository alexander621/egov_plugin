 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: faq_info.asp
' AUTHOR:   David Boyer
' CREATED:  05/12/2009
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays a specific FAQ/Rumor Mill detail.
'
' MODIFICATION HISTORY
' 1.0  05/12/09	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim re, matches

Set re = New RegExp
re.Pattern = "^\d+$"

'Check for the faqtype
if request("faqtype") <> "" then
	lcl_faqtype = UCASE(request("faqtype"))
else
	lcl_faqtype = "FAQ"
end if

'Based on the faqtype check for the proper permission
if lcl_faqtype = "RUMORMILL" then
	lcl_pagetitle   = "Rumor Mill"
	lcl_screenlabel = lcl_pagetitle
	lcl_feature     = "rumormill"
else
	lcl_pagetitle   = "FAQ"
	lcl_screenlabel = "Frequently Asked Questions"
	lcl_feature     = "faq"
end if

lcl_feature_name = GetFeatureName(lcl_feature)

'Check to see if the feature is offline
if isFeatureOffline(lcl_feature) = "Y" then
	response.redirect "outage_feature_offline.asp"
end if

Dim oActionOrg

set oActionOrg = New classOrganization

'Show/Hide all hidden fields.  TEXT=Show, HIDDEN=Hide
lcl_hidden = "HIDDEN"

lcl_faqid = request("id")
Set matches = re.Execute(lcl_faqid)
If matches.Count > 0 Then
	lcl_faqid = CLng(lcl_faqid)
Else
	response.redirect "faq.asp?faqtype=" & lcl_faqtype
End If 

if lcl_faqid <> "" then
	if not isnumeric(lcl_faqid) then
		response.redirect "faq.asp?faqtype=" & lcl_faqtype
	end if
end if
%>
<html>
<head>
 	<title>E-Gov Services - <%=sOrgName & " " & lcl_screenlabel%></title>

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

 	<script language="javascript" src="scripts/modules.js"></script>
 	<script language="javascript" src="scripts/easyform.js"></script>
    <script language="javascript" src="scripts/ajaxLib.js"></script>
    <script language="javascript" src="scripts/setfocus.js"></script>
    <script language="javascript" src="scripts/removespaces.js"></script>

    <script type="text/javascript" src="https://s7.addthis.com/js/200/addthis_widget.js"></script>
    <script type="text/javascript">var addthis_pub="cschappacher";</script>

</head>
<!--#include file="include_top.asp"-->
<p>
<table border="0" cellspacing="0" cellpadding="0" width="800">
  <tr><td><font class="pagetitle"><%=lcl_feature_name%></font></td></tr>
</table>
</p>

<% RegisteredUserDisplay("") %>

<div id="content">
  <div id="centercontent">

<table border="0" cellspacing="0" cellpadding="2" width="800">
  <form name="faq_rumormill" action="faq_info.asp" method="post">
  <tr valign="top">
      <td width="600">
        <% displayFAQ iorgid, lcl_faqid %>
      </td>
  </tr>
  </form>
</table>
<p>&nbsp;</p>

  </div>
</div>
<!-- #include file="include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub displayFAQ( ByVal p_orgid, ByVal p_faqid)
	Dim sSql, oFAQInfo

  sSql = "SELECT faqid, faqq, faqa, faqcategoryid, publicationstart, publicationend "
  sSql = sSql & " FROM faq  WHERE orgid = " & p_orgid
  sSql = sSql & " AND faqid = "   & p_faqid

 	set oFAQInfo = Server.CreateObject("ADODB.Recordset")
  oFAQInfo.Open sSql, Application("DSN"), 3, 1

  if not oFAQInfo.eof then
     do while not oFAQInfo.eof

        response.write "<p>" & vbcrlf
        response.write "  <fieldset>" & vbcrlf
        response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
        response.write "    <tr>" & vbcrlf
        response.write "        <td valign=""top"">" & vbcrlf
        response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
        response.write "              <tr>" & vbcrlf
        response.write "                  <td style=""font-size:16px; font-weight:bold;"">" & oFAQInfo("faqq") & "</td>" & vbcrlf
        response.write "              </tr>" & vbcrlf
        response.write "            </table>" & vbcrlf
        response.write "            <br />" & vbcrlf
        response.write "            <p>" & oFAQInfo("faqa") & "</p>" & vbcrlf
        response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
        response.write "              <tr>" & vbcrlf
        response.write "                  <td align=""right"">" & vbcrlf
        response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
        response.write "                        <tr valign=""top"">" & vbcrlf
        response.write "                            <td>" & vbcrlf
                                                        displayYahooBuzzButton p_orgid
        response.write "                            </td>" & vbcrlf
        response.write "                            <td>" & vbcrlf
                                                        displayAddThisButton p_orgid
        response.write "                            </td>" & vbcrlf
        response.write "                        </tr>" & vbcrlf
        response.write "                      </table>" & vbcrlf
        response.write "                      </a>" & vbcrlf
        response.write "                  </td>" & vbcrlf
        response.write "              </tr>" & vbcrlf
        response.write "            </table>" & vbcrlf
        response.write "        </td>" & vbcrlf
        response.write "    </tr>" & vbcrlf
        response.write "  </table>" & vbcrlf
        response.write "  </fieldset>" & vbcrlf
        response.write "</p>" & vbcrlf

        oFAQInfo.movenext
     loop
  end if

  oFAQInfo.close
  set oFAQInfo = nothing

end sub
%>
