<!--#include file='header.asp'-->
<% if not HasPermission("CanEditRoles") then
'response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleUpdateRole)
 end if %> 
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
      <td><font size="+1"><b><%=lanUpdateRoleTitle%></b></font>
	  <br><img src='../images/arrow_back.gif' align='absmiddle'> <a href="display_roles.asp"><%=langBackToRoleDisplay%></a>
	  </td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink.asp'-->      
      </td>
      <td colspan="2" valign="top">
<%
dim thisname,conn,rs,strSQL,rolename,title,roledescription,newrolename,newroledescription
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")

if request.form("start")="Y" then
'==========================================================
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3
strSQL = "select * from roles where roleid="&clng(trim(request.form("roleid")))
rs.Open strSQL
rolename=rs("rolename")
roledescription=rs("roledescription")
newrolename=replace(trim(request.form("rolename")),"'","''")
newroledescription=replace(trim(request.form("roledescription")),"'","''")
rs.close
if trim(rolename)<>trim(newrolename) then
Title="<li>Role name: <B>"&rolename&"</B> is sucessfully updated to <B>"&newrolename&"</B></li>"
end if
if trim(roledescription)<>trim(newroledescription) then
Title=Title+"<li>Role description is sucessfully updated!!</li>"
else
if trim(rolename)=trim(newrolename) then Title="<li>No changes have been made!!</li>"
end if


strSQL = "update roles set rolename='"&newrolename&"',  roledescription='"&newroledescription&"'  where roleid="&clng(trim(request.form("roleid")))
conn.execute(strSQL)
response.write "<br>"&Title
'===========================================================
else 
'===========================================================
'-- check is the role id is entered or not ---------
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3

if trim(request.querystring("roleid"))="" then
response.write "<br>No roleid is entered, end program here"
response.end
else
strSQL = "select *  from roles where roleid="&clng(trim(request.querystring("roleid")))
end if
'--------------------------------------------
rs.Open strSQL
'-----------------
if rs.recordcount=0 then
response.write "<br>Cannot find role in the system name in database"
response.write "<a href='javascript:> Go back</a>"
response.end
end if
'---------------------------
rolename=rs("rolename")
Title="Update Directory:"&rolename
%>
<FORM METHOD=POST name=UpdateRole ACTION="update_role.asp" >

<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateRole.submit();" onclick="return CheckCommitteeField();"><%=langUpdate%></a></div>

<table border="0" width="100%"  class='tablelist' cellpadding='5' cellspacing='0'>
 <tr><th  width=110  align=left><%=langProperty%></th><th   align=left><%=langValue%></th><tr>   

<tr>
	<td ><%=langrole%></td>
    <td align=left><input type=text name="rolename" value="<%=rs("rolename")%>" size=30 maxlength=50></td>
    </tr>
	          
	<tr>
	<td><%=langDescription%></td>
    <td><textarea rows="3" cols="30" name="roleDescription"><%=rs("roleDescription")%></textarea></td>
     </tr>  

	<tr>
	<td >Permissions</td>
    <td >
	<div id="permission" name="permission" bgcolor="#DDDDDD"></div>
	<% response.write "<a href=""javascript:openWin1('ManageRolePermission.asp?roleid="&rs("roleid")&"','_blank')""><img src='../images/newpermission.gif' align='absmiddle' border=0>&nbsp;"&langedit&"</a> " %></td>
     </tr>  

</table>

<div style="font-size:10px; padding-bottom:5px;"><img src="../images/cancel.gif" align="absmiddle">&nbsp;<a href="javascript:history.back();"><%=langCancel%></a>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/go.gif" align="absmiddle">&nbsp;<a href="javascript:document.all.UpdateRole.submit();" onclick="return CheckCommitteeField();"><%=langUpdate%></a></div>	
	
	<input type=hidden name="roleid" value="<%=rs("roleid")%>">
	<input type=hidden name="start" value="Y">


</FORM>

<%
'=================================================
rs.close
set rs=nothing
conn.close
set conn=nothing
end if %>
</td></tr></table>
 <!--#include file='footer.asp'-->


 <script language="JavaScript">
function CheckCommitteeField()
				{
					if (document.UpdateRole.rolename.value == "")
					{
						alert("Role Name is required");
						document.UpdateRole.rolename.focus();
					return false;				
					}					
					return true;
				}

</script>

<script language=javascript>
	function openWin1(url, name) {
  popupWin = window.open(url, name,
"resizable,width=550,height=450");
}
</script>
