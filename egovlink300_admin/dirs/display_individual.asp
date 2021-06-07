<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<% 
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit users" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 
%>

<html>
<head>
  <title><%=langBSCommittees%></title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

  <script src="../scripts/selectAll.js"></script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
    <tr>
      <td background="../images/back_main.jpg">

			<% ShowHeader sLevel %>
			<!--#Include file="../menu/menu.asp"--> 

      </td>
    </tr>

  </table>

<!-- #include file="dir_constants.asp"-->

<div id="content">
	<div id="centercontent">

  <table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
    <tr>
      <td><font size="+1"><b><%=langUserDetailedInformation%></b></font><br>
	  <% if request.querystring("s")<>"" then %>
	  <img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;&nbsp;
	   <A HREF="javascript:history.go(-1)"><%=langGoback%></a>  
	  <%else%>
	  <div id="backto">
	  <img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;&nbsp;
	  <A HREF="display_member.asp"><%=langBackToUserDisplay%></a></div>
	   <% end if %>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="1" valign="top">
<% 'if HasPermission("CanEditUser") or Session("UserID") = clng(trim(request.querystring("userid"))) then  %>
&nbsp;<img src="../images/cut.gif"> 
	  <a href="update_user.asp?userid=<%=CLng(trim(request.querystring("userid")))%>"><%=langEdit%></a> 
	  &nbsp;&nbsp;
	  <img src="../images/small_delete.gif" align="absmiddle">&nbsp; 
	  <a href="delete_user.asp?userid=<%=CLng(trim(request.querystring("userid")))%>" onClick="javascript: return confirm('<%=langDelteSingleUser%>');"><%=langDelete%></a>
<%	If UserHasPermission( Session("UserId"), "user permission" ) Then %>	  
	  &nbsp;&nbsp;
	  <img src="../images/newpermission.gif" height="16" width="16" alt="" border="0" />&nbsp;
	  <a href="../security/edit_user_security.asp?iuserid=<%=clng(trim(request("userid")))%>">User Permissions</a>
<%	End If %>
<br /><br />
<% 'end if %>
          <!--#include file='display_individual.html'-->
		  </td>
  <td width='200'>&nbsp;</td>
    </tr>	
  </table>
 </div>
 </div>

<br /><br /><br />
  <!--#Include file="../admin_footer.asp"-->  

  <!--#include file='footer.asp'-->
