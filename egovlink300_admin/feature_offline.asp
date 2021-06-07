<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<%
Dim sError

sLevel = "" ' Override of value from common.asp
%>
<html>
<head>
  <title>E-Gov - Maintenance</title>

  <link rel="stylesheet" type="text/css" href="global.css" />
  <link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />

  <script language="Javascript" src="scripts/modules.js"></script>
  <script language="JavaScript" src="scripts/ajaxLib.js"></script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
<div id="centercontent">
<%
 'Set up the values
  lcl_feature        = "Create Action Line Requests"
  lcl_start_datetime = "12-17-07 08:00 AM EST"
  lcl_end_datetime   = "12-17-07 10:30 AM EST"
%>
<table id="bodytable" border="0" cellpadding="0" cellspacing="0" class="start">
  <tr>
      <td align="center">
          The feature "<font color="#800000" style="font-size: 11px"><strong><%=lcl_feature%></strong></font>" will be unavailable between<br>
          <font color="#800000" style="font-size: 11px"><strong><%=lcl_start_datetime%></strong></font> to 
          <font color="#800000" style="font-size: 11px"><strong><%=lcl_end_datetime%></strong></font><br>
          due to scheduled maintenance outage.
 	    </td>
  </tr>
</table>
</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="admin_footer.asp"-->  
</body>
</html>
