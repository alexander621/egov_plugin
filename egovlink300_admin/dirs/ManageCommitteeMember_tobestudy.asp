<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs1 = Server.CreateObject("ADODB.Recordset")
set rs1.ActiveConnection = conn
rs1.CursorLocation = 3 
rs1.CursorType = 2 
set rs2 = Server.CreateObject("ADODB.Recordset")
set rs2.ActiveConnection = conn
rs2.CursorLocation = 3 
rs2.CursorType = 2 

if trim(request.querystring("groupid"))="" then
response.write "<br>No Groupid is entered, end program here"
response.end
else
strSQL1 = "select u.userid,lastname,firstname  from users u, usersgroups ug where u.userid=ug.userid and ug.groupid="&clng(trim(request.querystring("groupid")))
strSQL2 = "select *  from users u where u.userid not in (select userid from usersgroups ug where ug.groupid="&clng(trim(request.querystring("groupid")))&")"
end if

rs1.Open strSQL1
rs2.Open strSQL2

'if rs.recordcount=0 then
'response.write "The user you request doesn't exist in the database, or user's organization id doesn't exist<br>"
'response.end
'end if
'rs.movefirst
'call CommitteeMemberList
'call TheRemainingMemberList
 response.write "<form method='POST' action='Committee_deletemember.asp'>"
   response.write "<select size='20' name='committeelist' multiple>"
	   'for i=0 to rs1.recordcount
	   while not rs1.eof
    	response.write "<option value="&rs1("userid")&">"&rs1("lastname")&" "&rs1("firstname")&"</option>"
		rs1.MoveNext
wend
		'next 
    response.write   "</select></p>"
     response.write  "</form>"

rs1.close
set rs1=nothing

  response.write "<form method='POST' action='Committee_AddMember.asp'>"
   response.write "<select size='20' name='OtherList' multiple>"
	   for i=0 to rs2.recordcount
	   	  ' while not rs2.eof
    	response.write "<option value="&rs2("userid")&">"&rs2("lastname")&" "&rs2("firstname")&"</option>"
				rs2.movenext
			'	wend
		next
    response.write   "</select></p>"
     response.write  "</form>"

rs2.close
set rs2=nothing

conn.close
set conn=nothing

sub CommitteeMemberList
  

end sub

sub  TheRemainingMemberList
    
end sub

%>
