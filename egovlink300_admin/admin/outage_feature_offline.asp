<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  outage_feature_offline.asp
' AUTHOR:    David Boyer
' CREATED:   01/21/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Outage screen that users are redirected to for any high-level (parent) feature that has been "turned off"
'               using the Outage Maintenance screen.
'
' MODIFICATION HISTORY
' 1.0   01/21/08    David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
 Dim sError
 sLevel = "../" 'Override of value from common.asp
%>
<html>
<head>
  <title>E-Gov - Maintenance</title>

  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
  <script language="javaScript" src="../scripts/ajaxLib.js"></script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!-- #include file="../menu/menu.asp" -->

<!--BEGIN PAGE CONTENT-->
<div id="content">
  <div id="centercontent">
<table id="bodytable" border="0" cellpadding="5" cellspacing="0" class="start">
  <tr valign="top">
      <td>
          <table border="0" cellspacing="0" cellpadding="5" class="tableadmin" style="width: 400px">
            <tr>
                <th align="left">Features Offline</th>
            </tr>
          <%
           'Retreive all of the features that are offline
            sSQLf = "SELECT o.featurename "
            sSQLf = sSQLf & " FROM egov_organization_features o, egov_organizations_to_features f "
            sSQLf = sSQLf & " WHERE o.featureid = f.featureid "
            sSQLf = sSQLf & " AND f.orgid = " & session("orgid")
            sSQLf = sSQLf & " AND o.feature_offline = 'Y' "

            set rsf = Server.CreateObject("ADODB.Recordset")
            rsf.Open sSQLf, Application("DSN"),3,1

            if not rsf.eof then
               while not rsf.eof
                  response.write "            <tr><td>" & rsf("featurename") & "</td></tr>"
                  rsf.movenext
               wend
            end if
          %>
          </table>
 	    </td>
      <td>
          <table border="0" cellspacing="0" cellpadding="5" class="tableadmin" style="width: 400px">
            <tr>
                <th align="center">Offline Message</th>
            </tr>
          <%
           'Retreive all of the features that are offline
            sSQLm = "SELECT outage_message "
            sSQLm = sSQLm & " FROM organizations "
            sSQLm = sSQLm & " WHERE orgid = " & session("orgid")

            set rsm = Server.CreateObject("ADODB.Recordset")
            rsm.Open sSQLm, Application("DSN"),3,1

            if not rsm.eof then
               response.write "            <tr><td align=""center"">" & rsm("outage_message") & "<p></td></tr>"
            end if
          %>
          </table>
 	    </td>
  </tr>
</table>
  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  
</body>
</html>
