<!--#include file='../../header.asp'-->
  <table border="0" cellpadding="10" cellspacing="0" width="100%">
    <tr>
      <td width="151" align="center"><img src="../images/icon_directory.jpg"></td>
<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL="select groupname from groups g where g.groupid="&clng(request.querystring("groupid"))
rs.Open strSQL
CommitteeName=rs("groupname")
rs.close
set rs=nothing
%>

      <td><font size="+1"><b><%=langcommittee%>:<%=CommitteeName%></b></font><br><A HREF="display_committee.asp"><%=langCommittees%></A></td>
      <td width="200">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top">
        <!--#include file='quicklink.asp'-->      
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

if trim(request.querystring("groupid"))="" then
response.write "<br>No Groupid is entered, end program here"
response.end
else
strSQL1 = "select u.userid,lastname,firstname  from users u, usersgroups ug where u.userid=ug.userid and ug.groupid="&clng(trim(request.querystring("groupid")))
strSQL2 = "select *  from users u where u.userid not in (select userid from usersgroups ug where ug.groupid="&clng(trim(request.querystring("groupid")))&")"
end if

rs1.Open strSQL1
rs2.Open strSQL2

response.write "<TABLE border=0 cellpadding=5 cellspacing=0 width=350 class='tablelist'> <TR>"
response.write "<TD>"
response.write langManageCommitteeMember1
call CommitteeMemberList
response.write "</TD><td>"
response.write "<A HREF='javascript:document.c1.submit();'><b>--></b></A>"
response.write "<br><br>"
response.write "<A HREF='javascript:document.r1.submit();'><b><--</b></A>"
response.write "</td><TD>"
response.write langManageCommitteeMember2
call TheRemainingMemberList
response.write "</TD>"
response.write "</TR> </TABLE>"
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing

conn.close
set conn=nothing

sub CommitteeMemberList
   response.write "<form name=c1 method='POST' action='Committee_deletemember.asp?groupid="&trim(request.querystring("groupid"))&"'>"
   response.write "<select size='20' name='committeelist' multiple>"
	   for i=0 to rs1.recordcount-1
	   memberstr=rs1("lastname")&" "&rs1("firstname")
	   if trim(memberstr)="" then memberstr="** "&rs1("userid") &" **"
    	response.write "<option value="&rs1("userid")&">"&memberstr&"</option>"
		rs1.movenext
		next 
    response.write   "</select></p>"
     response.write  "</form>"

end sub

sub  TheRemainingMemberList
      response.write "<form name=r1 method='POST' action='Committee_AddMember.asp?groupid="&trim(request.querystring("groupid"))&"'>"
   response.write "<select size='20' name='OtherList' multiple>"
	   for i=0 to rs2.recordcount-1
		 memberstr=rs2("lastname")&" "&rs2("firstname")
	   if trim(memberstr)="" then memberstr="** "&rs2("userid") &" **"
    	response.write "<option value="&rs2("userid")&">"&memberstr&"</option>"
				rs2.movenext
		next
    response.write   "</select></p>"	 
     response.write  "</form>"
end sub

%>
<!--#include file='footer.asp'-->
