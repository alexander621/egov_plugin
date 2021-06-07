<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<% 
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "groups" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html>
<head>
	<title><%=langBSCommittees%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script src="../scripts/selectAll.js"></script>

	<script language="Javascript">
	  <!--

		function UpdateFamily( sUserId )
		{
			location.href='../dirs/family_members.asp?userid=' + sUserId;
		}
	//-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
    <tr>
      <td background="../images/back_main.jpg">
          <%  'DrawTabs tabCommittees,2  %>

			<% ShowHeader sLevel %>
			<!--#Include file="../menu/menu.asp"--> 

      </td>
    </tr>

  </table>

<!-- #include file="dir_constants.asp"-->
<div id="content">
	<div id="centercontent">
<!------------------------------------- -->

<%
'if not HasPermission("CanRegisterCommittee") then
'	response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleRegisterCommittee)
'end if
%>


  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td><font size="+1"><b>Registration: New Citizen Group</b></font><br><br><img src="../images/arrow_2back.gif" align="absmiddle">&nbsp;<A HREF="display_citizen_groups.asp">Back to Group List</A></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td  valign="top">
        <!--#include file='new_citizengroup.html'-->
	</td>
	  <td width="200">&nbsp;</td>
    </tr>
  </table>

 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->