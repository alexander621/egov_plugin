<%
if trim(request.querystring("userid"))="" then
response.write "<br>No User ID is entered, end program here"
response.end
end if

thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL="select firstname,lastname,userid from users u where u.userid="&clng(request.querystring("userid"))
rs.Open strSQL
Title="<B>"&rs("firstname")&" "&rs("lastname")&"("&rs("userid")&")</b>"
rs.close
set rs=nothing
%>

<TABLE border=0 cellspacing=0 cellpadding=5 align=center>
<TR><TD align=center><font size="+1">Manage User: <b> <%=Title %> </b></font><br></TD></TR>
<TR><TD>
<%
thisname=request.servervariables("script_name")
set rs1 = Server.CreateObject("ADODB.Recordset")
set rs1.ActiveConnection = conn
rs1.CursorLocation = 3 
rs1.CursorType = 3 
strSQL1 = "select RoleID,RoleName "& _
"from UsersGroupsRoles UGR, groups g where UGR.userid="&clng(request.querystring("userid")) &" and UGR.groupid=g.groupid"
rs1.Open strSQL1
'response.write "<br>strSQL1="&strSQL1
while not rs1.eof
'response.write "<br>"&h <B>Roles</b>:"&rs1("RoleName")
'response.write "<br>"&rs1("GroupName")&"("&rs1("GroupID")&") with <B>Roles</b>:"&rs1("RoleName")
wend
rs1.close
set rs1=nothing
conn.close
set conn=nothing

%>

</TD>
</TR>
</TABLE>
