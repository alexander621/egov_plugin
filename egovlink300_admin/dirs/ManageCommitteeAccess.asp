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
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL="select groupname from citizengroups g where g.orgid="&Session("OrgID")&" and g.groupid="&clng(request("groupid"))
rs.Open strSQL
If Not rs.EOF Then 
	CommitteeName = rs("groupname")
End If 
rs.close
set rs=nothing
%>
  <table border="0" cellpadding="5" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center">&nbsp;<!--<img src="../images/icon_directory.jpg">--></td>
      <td><font size="+1"><b><%=langCommittee&": "&CommitteeName%></b></font>
	  <br><img src='../images/arrow_back.gif' align='absmiddle'> <a href='display_committee.asp'><%=langBackToCommittee%></a>
	  <br><%="<br><br><B>"&request.querystring("strMessage")&"</B>"%>
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink.asp'-->      
      </td>
      <td colspan="2" valign="top">
<%
strAction="<div style='font-size:10px; padding-bottom:5px;'><img src='../images/cancel.gif' align='absmiddle'>&nbsp;<a href='javascript:history.back();'>"&langCancel&"</a>&nbsp;&nbsp;&nbsp;&nbsp;<img src='../images/go.gif' align='absmiddle'>&nbsp;<a href='javascript:document.all.ManageCommitteAccess.submit();' >"&langUpdate&"</a></div>"

thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set cmd=Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection=conn
cmd.commandtext="GetDirectoryPermissionType"
cmd.commandtype=&H0004

set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 2 
'strSQL = "select g.groupname, sum(u.userid) as entry from groups g, usersgroups as u where u.groupid=g.groupid group by g.groupname"
strSQL = "select groupid,groupname from citizengroups where orgid="&Session("OrgID")&"  order by groupname"
rs.Open strSQL,,, 2 
'------- the following code dealing with the recordcount=0-------

if rs.recordcount>0 then
	displayform
end if

rs.close
conn.close
set rs=nothing
set conn=nothing
set cmd=nothing
'--------------------------------------------------------------------


sub displayform()
	groupid=request.querystring("groupid")
	rs.movefirst
	delimeter="_qwcdfsddewrf45435ds_"
	If sBgcolor = "#ffffff" Then 
		sBgcolor = "#eeeeee" 
	Else 
		sBgcolor = "#ffffff"
	End If 
	Do While Not rs.EOF
		sChecked0="" 
		sChecked1="" 
		sChecked2="" 
		cmd.Parameters.Refresh
		cmd.parameters(1)=rs("groupid")
		cmd.parameters(2)=groupid
		cmd.execute
		ResultID=cmd.parameters(0)
		select case ResultID
			case 0
				sChecked0="checked"
			case 1 
				sChecked1="checked" 
			case 2 
				sChecked2="checked" 
		end select

		sGroups = sGroups & "<tr bgcolor=""" & sBgcolor & """><td>"&rs("groupname")&"</td>"&vbcrlf
		sGroups = sGroups & "<td><input type=radio name=""field_"&rs("groupid")&""" "&sChecked0&" value=""0"">&nbsp;</td>"&vbcrlf
		sGroups = sGroups & "<td><input type=radio name=""field_"&rs("groupid")&""" "&sChecked1&" value=""1"">&nbsp;</td>"&vbcrlf
		sGroups = sGroups & "<td><input type=radio name=""field_"&rs("groupid")&""" "&sChecked2&" value=""2"">&nbsp;</td></tr>"&vbcrlf
		rs.movenext
	Loop 
	response.write "<form name=""ManageCommitteAccess"" action=""Insert_GroupPermissions.asp?url="&thisname&"&groupid="&groupid&""" method=""post"" >"
	response.write strAction
	response.write "<TABLE cellpadding=5 width=400 cellspacing=0 class=""tablelist"">"
	response.write "<th align=""left"">Group</th>"
	response.write "<th align=""left"">Hidden</th>"
	response.write "<th align=""left"">View</th>"
	response.write "<th align=""left"">Edit</th>"
	response.write  sGroups
	response.write "</TABLE>" 
	response.write strAction
	response.write "</form>"
 end sub 
 %>
 
</td></tr></table>
 </div>
 </div>

<!--#Include file="../admin_footer.asp"-->  

<!--#include file='footer.asp'-->
