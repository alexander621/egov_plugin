
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
<%
thisname=request.servervariables("script_name")
if trim(request.querystring("roleid"))="" then
response.write "<br>No Role is entered, end program here"
response.end
end if

thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL="select distinct roleID from  ManageRolesPermissions RP where RP.roleID="&clng(request.querystring("roleid"))
rs.Open strSQL
roleID=rs("roleID")
rs.close
set rs=nothing
%>

      <td><font size="+1">Manage Role Permissions <b> <%=roleID%> </b></font><br></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
&nbsp;
      </td>
      <td colspan="2" valign="top">
  
<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs1 = Server.CreateObject("ADODB.Recordset")
set rs1.ActiveConnection = conn
rs1.CursorLocation = 3 
rs1.CursorType = 3 
set rs2 = Server.CreateObject("ADODB.Recordset")
set rs2.ActiveConnection = conn
rs2.CursorLocation = 3 
rs2.CursorType = 3 

strSQL1 = "select PermissionID , permissionName  from ManageRolesPermissions RP where RP.roleID="&clng(request.querystring("roleid")) &" order by permissionName"
strSQL2 = "select PermissionID  ,permissionName  from  permissions P where P.PermissionID not in (select permissionID from RolesPermissions R where R.roleID="&clng(trim(request.querystring("roleid")))&")" &" order by permissionName"


rs1.Open strSQL1
rs2.Open strSQL2

response.write "<TABLE border=0 cellpadding=5 cellspacing=0 width=350 class='tablelist'> <TR>"
response.write "<TD>"
response.write "Existing Roles"
call ExistingRoleList
response.write "</TD><td>"
response.write "<A HREF='javascript:document.c1.submit();'><b>></b></A>"
response.write "<br><br>"
response.write "<A HREF='javascript:document.r1.submit();'><b><</b></A>"
response.write "</td><TD>"
response.write "Available Roles"
call TheRemainingMemberList
response.write "</TD>"
response.write "</TR> </TABLE>"
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing

conn.close
set conn=nothing

sub ExistingRoleList
   response.write "<form name=c1 method='POST' action='RoleDelete.asp?roleid="&trim(request.querystring("roleid"))&"'>"
   response.write "<select size='20' name='ExistingList' multiple>"
	   for i=0 to rs1.recordcount-1
	   memberstr=rs1("permissionName")
'	   if trim(memberstr)="" then memberstr="** "&rs1("userid") &" **"
    	response.write "<option value="&rs1("PermissionID")&">"&memberstr&"</option>"
		rs1.movenext
		next 
    response.write   "</select></p>"
     response.write  "</form>"

end sub

sub  TheRemainingMemberList
      response.write "<form name=r1 method='POST' action='RoleAdd.asp?roleid="&trim(request.querystring("roleid"))&"'>"
    response.write "<select size='20' name='RemainingList' multiple>"
	   for i=0 to rs2.recordcount-1
		 memberstr=rs2("permissionName")
'	   if trim(memberstr)="" then memberstr="** "&rs2("userid") &" **"
    	response.write "<option value="&rs2("PermissionID")&">"&memberstr&"</option>"
				rs2.movenext
		next
    response.write   "</select></p>"	 
     response.write  "</form>"
end sub

%>

