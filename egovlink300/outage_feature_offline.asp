<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  outage_feature_offline.asp
' AUTHOR:    David Boyer
' CREATED:   01/22/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Outage screen that users are redirected to for any high-level (parent) feature that has been "turned off"
'               using the Outage Maintenance screen.
'
' MODIFICATION HISTORY
' 1.0   01/22/08    David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sError 

'Show/Hide hidden fields.  To Hide = "HIDDEN", To Show = "TEXT"
lcl_hidden = "hidden"
%>
<html>
<head>

	<%If iorgid = 7 Then %>
		<title><%=sOrgName%></title>
	<%Else%>
		<title>E-Gov Services <%=sOrgName%></title>
	<%End If%>

 	<link rel="stylesheet" type="text/css" href="css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="global.css" />
 	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

 	<script language="Javascript" src="scripts/modules.js"></script>
 	<script language="Javascript" src="scripts/easyform.js"></script>
  <script language="JavaScript" src="scripts/ajaxLib.js"></script>
  <script language="JavaScript" src="scripts/removespaces.js"></script>
  <script language="JavaScript" src="scripts/setfocus.js"></script>

</head>
<!--#Include file="include_top.asp"-->

<!--BODY CONTENT-->
<div id="content">
 	<div id="centercontent">
<table id="bodytable" border="0" cellpadding="5" cellspacing="0" class="start">
  <tr valign="top">
      <td>
          <table border="0" cellspacing="0" cellpadding="5" style="width: 400px">
            <caption align="left"><font color="#ff0000"><strong>Features Not Available at this Time</strong></font><hr size="1"></caption>
          <%
           'Retreive all of the features that are offline
            sSQLf = "SELECT o.featurename "
            sSQLf = sSQLf & " FROM egov_organization_features o, egov_organizations_to_features f "
            sSQLf = sSQLf & " WHERE o.featureid = f.featureid "
            sSQLf = sSQLf & " AND f.orgid = " & iorgid
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
          <table border="0" cellspacing="0" cellpadding="5" style="width: 400px">
          <%
           'Retreive all of the features that are offline
            sSQLm = "SELECT outage_message "
            sSQLm = sSQLm & " FROM organizations "
            sSQLm = sSQLm & " WHERE orgid = " & iorgid

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
<p>
  </div>
</div>

<!--#Include file="include_bottom.asp"-->