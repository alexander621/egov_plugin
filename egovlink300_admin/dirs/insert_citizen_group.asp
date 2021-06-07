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
dim conn,cmd,resultid,strSuccess
'if not HasPermission("CanRegisterCommittee") then
'response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleRegisterCommittee)
' end if %> 
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center">&nbsp;<!--<img src="../images/icon_directory.jpg">--></td>
      <td><font size="+1"><b>Registration: New Group</b></font>
	  <br>
	  <div id="goback" name="goback">
	  <img src='../images/arrow_back.gif' align='absmiddle'><a href='display_citizen_groups.asp'>Back to Group List</a>
	  </div>
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink_citizen.asp'-->      
      </td>
      <td colspan="2" valign="top">
        
<%
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set cmd=Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection=conn
cmd.commandtext="NewCitizengroup"
cmd.commandtype=&H0004
cmd.Parameters.Refresh
With request
cmd.parameters(1)=clng(.form("orgid"))
cmd.parameters(2)=left(.form("groupname"),44)
cmd.parameters(3)=left(.form("groupdescription"),150)
cmd.parameters(4)=clng(.form("grouptype"))
cmd.execute
end with
ResultID=cmd.parameters(0)
conn.close
set conn=nothing
set cmd=nothing
'response.write "<br>ResultID="&ResultID
select case ResultID
case -100
response.write "<br><li>"&langInsertDatabaseError&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 0
response.write "<br><li>"&langInsertCommittee2&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 2
response.write "<br><li>"&langInsertCommittee3&"</li>"
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
case 1
response.write "<br><li>"&langInsertCommittee4&"</li>"
strSuccess="<br><img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<A HREF=display_citizen_groups.asp>Back to Group List</A>"
response.write "<script>document.all.goback.innerHTML="""&strSuccess&"""</script>"
end select
%>
</td></tr></table>

 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->

