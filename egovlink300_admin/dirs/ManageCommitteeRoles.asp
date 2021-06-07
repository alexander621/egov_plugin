<!--#include file='header_simply.asp'-->
<%
dim conn,rs1,rs2,rs,strSQL1,strSQL2,strSQL,i,thisname,committeeName,memberstr,rolestr
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


<% if not HasPermission("CanEditCommittee") and not HasPermission("CanEdit"&CommitteeName) then
response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleEditCommittee)
 end if %>

<body onload="javasript:opener.location.reload(true);">


  <table border="0" cellpadding="10" cellspacing="0" width="100%" bgcolor="#c9def0">


    <tr>

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
'response.write "<br>No Groupid is entered, end program here"
response.end
else
strSQL1 = "select r.roleid,rolename  from groupsroles gr, roles r where gr.roleid=r.roleid and gr.groupid="&clng(trim(request.querystring("groupid"))) &" and r.OrgID="&Session("OrgID")&" order by rolename"
strSQL2 = "select r.roleid,rolename from roles r where roleid not in (select r.roleid  from groupsroles gr, roles r where gr.roleid=r.roleid and gr.groupid="&clng(trim(request.querystring("groupid"))) &") and r.OrgID="&Session("OrgID")&" order by rolename"
end if
rs1.Open strSQL1
rs2.Open strSQL2


'response.write "Click on the arrow to make the transfer"
response.write "<CENTER>"
response.write "<FONT SIZE=+1>Directory:<B>"&CommitteeName&"</B></FONT>"
response.write "<br><a href='javascript:self.close();'><FONT SIZE=2>"&langAdminCloseWindow&"</FONT></a></font></CENTER>"
response.write "<TABLE border=0  cellpadding=13 cellspacing=0 width=350 >"
response.write "<tr><TD>"
'response.write langManageCommitteerole1
call CommitteeroleList
response.write "</TD><td align=center>"
response.write "&nbsp;<A HREF='javascript:document.c1.submit();'><img src='../images/ieforward.gif' align='absmiddle' border=0></A>"
response.write "<br><br>"
response.write "<A HREF='javascript:document.r1.submit();'><img src='../images/ieback.gif' align='absmiddle' border=0></A>"
response.write "</td><TD>"
'response.write langManageCommitteerole2
call TheRemainingroleList
response.write "</TD>"
response.write "</TR>"
response.write "</TABLE>"
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing

conn.close
set conn=nothing

sub CommitteeroleList
response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=130 >"
response.write "<TR><Td height=20><B>"&langManageRoles1&"</B></Td></TR>"
response.write "<TR><TD>"
 response.write "<form name=c1 method='POST' action='Committee_deleterole.asp?groupid="&trim(request.querystring("groupid"))&"'>"
   response.write "<select size='14' border=0 WIDTH=""130"" STYLE=""width:130px"" name='committeelist' multiple>"
	   for i=0 to rs1.recordcount-1
	   rolestr=trim(rs1("rolename"))
	   if trim(rolestr)="" then rolestr="** "&rs1("roleid") &" **"
    	response.write "<option value="&rs1("roleid")&">"&rolestr&"</option>"
		rs1.movenext
		next 
    response.write   "</select></p>"
     response.write  "</form>"
response.write "</TD></TR></TABLE>"  
end sub

sub  TheRemainingroleList
response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=130>"
response.write "<TR><Td height=20><B>"&langManageRoles2&"</B></Td></TR>"
response.write "<TR><TD>"
 response.write "<form name=r1 method='POST' action='Committee_Addrole.asp?groupid="&trim(request.querystring("groupid"))&"'>"
   response.write "<select size='14' WIDTH=""130"" STYLE=""width:130px"" name='OtherList' multiple>"
	   for i=0 to rs2.recordcount-1
	   rolestr=trim(rs2("rolename"))
	   if trim(rolestr)="" then rolestr="** "&rs2("roleid") &" **"
    	response.write "<option value="&rs2("roleid")&">"&rolestr&"</option>"
				rs2.movenext
		next
    	response.write "<option value='')>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>"
    response.write   "</select></p>"	 
     response.write  "</form>"
response.write "</TD></TR></TABLE>"      
end sub

%>
<!--#include file='footer.asp'-->
</body>
