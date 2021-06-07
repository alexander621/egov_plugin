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

strSQL1 = "select ugp.groupid,ugp.groupname from UsersGroupsPlus ugp where ugp.userid="&clng(request.querystring("userid"))
strSQL2 = "select g.groupid,g.groupname from Groups g where g.groupid not in ( select ugp.groupid from users u inner join UsersGroupsPlus ugp on u.userid=ugp.userid where  u.userid="&clng(request.querystring("userid")) &")"

rs1.Open strSQL1
rs2.Open strSQL2

response.write "<TABLE border=0 cellpadding=5 cellspacing=0 width=350 class='tablelist'> <TR>"
response.write "<TD>"
response.write "Already in Groups"
call ExistingRoleList
response.write "</TD><td>"
response.write "<A HREF='javascript:document.c1.submit();'><b>></b></A>"
response.write "<br><br>"
response.write "<A HREF='javascript:document.r1.submit();'><b><</b></A>"
response.write "</td><TD>"
response.write "Available Groups"
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
   response.write "<form name=c1 method='POST' action='DeleteGroup.asp?userid="&trim(request.querystring("userid"))&"'>"
   response.write "<select size='10' name='ExistingList' multiple>"
	   for i=0 to rs1.recordcount-1
	   memberstr=rs1("GroupName")&"("&rs1("GroupID")&")"
'	   if trim(memberstr)="" then memberstr="** "&rs1("userid") &" **"
    	response.write "<option value="&rs1("GroupID")&">"&memberstr&"</option>"
		rs1.movenext
		next 
    response.write   "</select></p>"
     response.write  "</form>"

end sub

sub  TheRemainingMemberList
      response.write "<form name=r1 method='POST' action='AddGroup.asp?userid="&trim(request.querystring("userid"))&"'>"
    response.write "<select size='10' name='RemainingList' multiple>"
	   for i=0 to rs2.recordcount-1
		  memberstr=rs2("GroupName")&"("&rs2("GroupID")&")"
'	   if trim(memberstr)="" then memberstr="** "&rs2("userid") &" **"
    	response.write "<option value="&rs2("GroupID")&">"&memberstr&"</option>"
				rs2.movenext
		next
    response.write   "</select></p>"	 
     response.write  "</form>"
end sub
%>

</TD>
</TR>
</TABLE>
