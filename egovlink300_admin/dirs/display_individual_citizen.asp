<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->

<% 
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit citizens" ) Then
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

		function FamilyList( sUserId )
		{
			location.href='family_list.asp?userid=' + sUserId;
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
	Session("RedirectPage") = "../dirs/display_individual_citizen.asp?userid=" & request("userid") 
	Session("RedirectLang") = "Return to Detailed User Information"
%>

<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
	<tr>
      <td><font size="+1"><b>Registration: Detailed User Information</b></font><br>
	  <% if request.querystring("s")<>"" then %>
	  <img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;&nbsp;
	   <A HREF="javascript:history.go(-1)"><%=langGoback%></a>
	  <%else%>
	  <div id="backto">
	  <img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;&nbsp;
	  <A HREF="display_citizen.asp"><%=langBackToUserDisplay%></a></div>
	   <% end if %>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="1" valign="top">
			<% 'if HasPermission("CanEditUser") or Session("UserID") = CLng(trim(request.querystring("userid"))) then  %>
			&nbsp;<img src=../images/cut.gif> <A HREF="update_citizen.asp?userid=<%=CLng(trim(request.querystring("userid")))%>"><%=langEdit%></A> &nbsp;&nbsp;
			<img src="../images/small_delete.gif" align="absmiddle">&nbsp; <A  HREF="delete_citizen.asp?userid=<%=CLng(trim(request.querystring("userid")))%>" onClick="javascript: return confirm('<%=langDelteSingleUser%>');"><%=langDelete%></a>
<%			if OrgHasFeature("hasfamily") Then %>
				&nbsp;&nbsp;
				<img src="<%=RootPath%>images/newgroup.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="javascript:FamilyList('<%=request("userid")%>');">Family Members</a>
<%			Else 
				If OrgHasFeature("activities") Then %>
					&nbsp;&nbsp;
					<img src="<%=RootPath%>images/newgroup.gif" width="16" height="16" align="absmiddle">&nbsp;<a href="javascript:UpdateFamily('<%=request("userid")%>');">Family Members</a>
<%				End If 
			End If %>
			<br /><br />
			<% 'end if %>
          <!--#include file='display_individual_citizen.html'-->
	  </td>
	  <td width='200'>&nbsp;</td>
    </tr>
 </table>

 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->