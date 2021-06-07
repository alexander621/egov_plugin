<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: subscriptionslog_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the rss feeds
'
' MODIFICATION HISTORY
' 1.0 06/29/09 David Boyer - Initial Version
' 1.1 08/05/09 David Boyer - Added "listtype"
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%> - Email Log}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
 	<script language="javascript" src="../scripts/getdates.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;width:900px;">
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr>
                  <td><font size="+1"><strong><%=Session("sOrgName")%>&nbsp;<%=lcl_pagetitle%> - Email Complaint/Bounce Log (Last 30 days)</strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
              </tr>
            </table>
            <br />
			<p>This log reports messages anytime we send an email and it bounces (i.e. isn’t able to be delivered), or receives a complaint (i.e. the recipient marked it as spam, and the software reported back to us that the user took that action).  In order for us to maintain a good reputation as a non-spaming mail sender, it’s important for us to take action on each of these.</p>
            <p><% displayEmailLog %></p>
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
sub displayEmailLog
%>
<div class="shadow">
	<table cellspacing="0" cellpadding="3" class="tablelist" border="0" style="width:900px">
		<tr align="left">
			<th nowrap="nowrap">Date of Complaint/Bounce</th>
			<th nowrap="nowrap">Type</th>
			<th nowrap="nowrap">Email Address</th>
			<th nowrap="nowrap">Action Taken</th>
		</tr>
				<style> tr.nopointer {cursor:default !important;}</style>
		<%
     		lcl_bgcolor             = "#ffffff"

			sSQL = "SELECT DateRecorded, MessageType, EmailAddress, ActionTaken FROM zEmailHandlingLog " _
					& " WHERE DateRecorded > '" & DateAdd("d", -30, date()) & "' AND ActionTaken <> 'Unhandled' AND OrgID = " & session("orgid") _
					& " ORDER BY DateRecorded DESC "
			set oRs = Server.CreateObject("ADODB.RecordSet")
			oRs.Open sSQL, Application("DSN"), 3, 1
			Do While not oRs.EOF
        		lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
				%>
				<tr bgcolor="<%=lcl_bgcolor%>" onMouseOver="mouseOverRow(this);" onMouseOut="mouseOutRow(this);" valign="top" class="nopointer">
					<td class="formlist" nowrap="nowrap"><%=oRs("DateRecorded")%></td>
					<td class="formlist" nowrap="nowrap"><%=oRs("MessageType")%></td>
					<td class="formlist" nowrap="nowrap"><%=oRs("EmailAddress")%></td>
					<td class="formlist" nowrap="nowrap"><%=Replace(oRs("ActionTaken"), " - Transient","")%></td>
				</tr>
				<%oRs.MoveNext
			loop
			oRs.Close
			Set oRs = Nothing

		%>
	</table>
</div>
<%
end sub
%>
