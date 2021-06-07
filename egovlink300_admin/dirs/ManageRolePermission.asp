<!--#include file='header_simply.asp'-->
<body  bgcolor="#c9def0">
  <table border="0" cellpadding="10" cellspacing="0" width="100%" bgcolor="#c9def0">
    <tr>
      <td colspan="2" valign="top">
<%
dim thisname,conn,rs,rolename,rs1,rs2,index,strSQL1,strSQL2,strSQL,i,memberstr,rolestr,strList
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
strSQL="select rolename from roles R where R.roleID="&clng(request.querystring("roleid"))
rs.Open strSQL
rolename=rs("rolename")
rs.close
set rs=nothing
%>

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

response.write "<CENTER><br>"
response.write "<FONT SIZE=+1>"&langRole&"<B>:"&rolename&"</B></FONT>"
response.write "<br><a href='javascript:self.close();'><FONT SIZE=2>"&langAdminCloseWindow&"</FONT></a></font></CENTER>"
response.write "<TABLE border=0  cellpadding=10 cellspacing=0 width=350  align=center>"
response.write "<tr><TD>"
call ExistingRoleList
response.write "</TD><td align=center>"

response.write "<A HREF='javascript:document.c1.submit();'><b><img src='../images/ieforward.gif' align='absmiddle' border=0></b></A>"
response.write "<br><br>"
response.write "<A HREF='javascript:document.r1.submit();'><b><img src='../images/ieback.gif' align='absmiddle' border=0></b></A>"
response.write "</td><TD>"
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
response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=130 >"
response.write "<TR><Td height=20><B>"&langExistingPermissions&"</B></Td></TR>"
response.write "<TR><TD>"

   response.write "<form name=c1 method='POST' action='RoleDelete.asp?roleid="&trim(request.querystring("roleid"))&"'>"
   response.write "<select size='25'  WIDTH=""200"" STYLE=""width:200px"" name='ExistingList' multiple>"
	   for i=0 to rs1.recordcount-1
	   memberstr=rs1("permissionName")
'	   if trim(memberstr)="" then memberstr="** "&rs1("userid") &" **"
    	response.write "<option value="&rs1("PermissionID")&">"&memberstr&"</option>"
		rs1.movenext
		next 
    response.write   "</select></p>"
     response.write  "</form>"
response.write "</TD></TR></TABLE>"  
end sub

sub  TheRemainingMemberList
index=1
response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=130 >"
response.write "<TR><Td height=20><B>"&langExistingPermissions2&"</B></Td></TR>"
response.write "<TR><TD>"
      response.write "<form name=r1 method='POST' action='RoleAdd.asp?roleid="&trim(request.querystring("roleid"))&"'>"
    response.write "<select size='25' WIDTH=""200"" STYLE=""width:200px"" name='RemainingList' multiple>"
	   for i=0 to rs2.recordcount-1
		 memberstr=rs2("permissionName")
'	   if trim(memberstr)="" then memberstr="** "&rs2("userid") &" **"
    	response.write "<option value="&rs2("PermissionID")&">"&memberstr&"</option>"
		if i<index then strList=strList+rs2("permissionName")+"," 
				rs2.movenext
		next
    response.write   "</select></p>"	 
     response.write  "</form>"
	 response.write "</TD></TR></TABLE>"  
	'response.write "<script>opener.document.all.permission.innerHTML = """&strList & """</script>"
 
end sub
%>

</TD>
</TR>
</TABLE>
</body>
