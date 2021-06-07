<!--#include virtual='/boardsite/includes/common.asp'-->
<%
'Option implicit 
%>


<html>
<head>
  <title>BoardSite {Committees}</title>
  <link href="../global.css" rel="stylesheet" type="text/css">
  <script language="Javascript">
  <!--
    function doNewDisc() {
      x = (screen.width - 475)/2;
      y = (screen.height - 515)/2;
      window.open("htmleditor/newdir.html", "newdir", "width=475,height=515,scrollbars=no,status=no,toolbar=no,menubar=no,left="+x+",top="+y);
      return false;
    }
  //-->
  </script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="menu">
    <tr>
      <td background="../images/back_main.jpg">
          <%  DrawTabs tabCommittees,2  %>
      </td>
    </tr>

  </table>


<%
const langAdminUserExtendedTitle= "Manage User Extended properties"
const langAdminCloseWindow	= "Close Window"
const langAdminNewProperty	= "New property"
const langAdminAllProperty	= "All properties"
fields_description_extended	= array("Extended ID","User ID","Property","Value","Added Time")
fields_description_committee=array("GroupID","Organizatioin ID","GroupName","Description","Added Time")


const langquicklink1="Manage All Users"
const langquicklink2="Manage Users-Roles"
const langquicklink3="Manage All Groups"
const langquicklink4="Manage Groups-Roles"
const langquicklink5="Manage the Roles(Permissions)"
const langquicklink6="Browse Role-Permission"
const langquicklink7="Manage the permissions"
const langquicklink8="2"
thisname=request.servervariables("script_name")
%>
<%
if instr(thisname,"/users/index.asp") then 				strAnchor1="<!a" 	else	strAnchor1="<a" 
if instr(thisname,"/UsersRoles/index.asp") then 		strAnchor2="<!a" 	else	strAnchor2="<a" 
if instr(thisname,"/groups/index.asp") then 			strAnchor3="<!a" 	else	strAnchor3="<a" 
if instr(thisname,"/groupsroles/index.asp") then 		strAnchor4="<!a" 	else	strAnchor4="<a" 
if instr(thisname,"/Roles/index.asp") then 				strAnchor5="<!a" 	else	strAnchor5="<a" 
if instr(thisname,"/RolesPermissions/index.asp") then 	strAnchor6="<!a" 	else	strAnchor6="<a" 
if instr(thisname,"/permissions/index.asp") then     	strAnchor7="<!a" 	else	strAnchor7="<a" 
if instr(thisname,"/UsersRoles2/index.asp") then 		strAnchor8="<!a" 	else	strAnchor8="<a" 
%>
<table border=0 cellpadding=1 class='tablelist' cellspacing=1 align=center width=850 bgcolor=#DEEAFE>
<tr>
    <td width="20%"><%=strAnchor1%> href="../users/index.asp?iOfaction=6"><%=langquicklink1 %></a></td>
    <td width="20%">
	<%=strAnchor2%> href="../UsersRoles/index.asp?iOfaction=6"><%=langquicklink2 %></a>
		<%=strAnchor8%> href="../UsersRoles2/index.asp?iOfaction=6"><%=langquicklink8 %></a>
	</td>
    <td width="20%" rowspan="2"><%=strAnchor5%> href="../Roles/index.asp?iOfaction=6"><%=langquicklink5 %></a></td>
    <td width="20%" rowspan="2"><%=strAnchor6%> href="../RolesPermissions/index.asp?iOfaction=6"><%=langquicklink6 %></a></td>
    <td width="20%" rowspan="2"><%=strAnchor7%> href="../permissions/index.asp?iofaction=6"><%=langquicklink7 %></a></td>
  </tr>
  <tr>
    <td width="20%"><%=strAnchor3%> href="../groups/index.asp?iOfaction=6"><%=langquicklink3 %></a></td>
    <td width="20%"><%=strAnchor4%> href="../groupsroles/index.asp?iOfaction=6"><%=langquicklink4 %></a></td>
  </tr>
</table>
<br>
