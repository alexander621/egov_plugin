<!--#include file='header_simply.asp'-->
<%
dim conn,rs1,rs2,rs,strSQL1,strSQL2,strSQL,i,thisname,committeeName,memberstr
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
%>

<% if not HasPermission("CanEditCommittee") and not HasPermission("CanEdit"&CommitteeName) then
'response.redirect "InvalidRole.asp?error="&server.urlencode(langInvalidRoleEditCommittee)
 end if %>
<body onload="javasript:opener.location.reload(true);" bgcolor="#c9def0">

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

tablename=request.querystring("tablename")
'response.write "<br>"&tablename
strSQL1 = "select tablefield as name from SubScribedItems where tablename='"&tablename&"'"
rs1.Open strSQL1
strSQL2 = "SELECT syscolumns.name as name FROM sysobjects INNER JOIN syscolumns ON sysobjects.id = syscolumns.id "& _
         "where sysobjects.name = '"&tablename&"'"& _
         " and syscolumns.name not in (select tablefield from SubScribedItems where tablename='"&tablename&"')"& _
    	 " ORDER BY syscolumns.colid"
rs2.Open strSQL2
'response.write "<br>"&strSQL1
'response.write "<br>"&strSQl2
response.write "<CENTER>"
response.write "<FONT SIZE=+1>Directory:<B>"&CommitteeName&"</B></FONT>"
response.write "<br><a href='javascript:self.close();'><FONT SIZE=2>"&langAdminCloseWindow&"</FONT></a></font></CENTER>"
response.write "<TABLE cellpadding=10 cellspacing=0 width=350  border=0 >"
response.write "<tr><TD>"
'response.write langManageCommitteeMember1
call CommitteeMemberList
response.write "</TD><td align=center>"
response.write "&nbsp;<A HREF='javascript:document.c1.submit();'><img src='../images/ieforward.gif' align='absmiddle' border=0></A>"
response.write "<br><br>"
response.write "<A HREF='javascript:document.r1.submit();'><img src='../images/ieback.gif' align='absmiddle' border=0></A>"
response.write "</td><TD>"
'response.write langManageCommitteeMember2
call TheRemainingMemberList
response.write "</TD>"
response.write "</TR>"
response.write "</TABLE>"
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing

conn.close
set conn=nothing

sub CommitteeMemberList
response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=130 >"
response.write "<TR><Td height=20><B>"&langManageCommitteeMember1&"</B></Td></TR>"
response.write "<TR><TD>"
 response.write "<form name=c1 method='POST' action='SubscribeTableField_delete.asp?tablename="&tablename&"'>"
   response.write "<select size='15' border=0 WIDTH=""140"" STYLE=""width:140px"" name='committeelist' multiple>"
   response.write rs1.recordcount
	   for i=0 to rs1.recordcount-1
	   memberstr=trim(rs1("name"))
	   if trim(memberstr)="" then memberstr="** "&rs1("name") &" **"
    	response.write "<option value="&rs1("name")&">"&memberstr&"</option>"
		rs1.movenext
		next 
    response.write   "</select></p>"
     response.write  "</form>"
response.write "</TD></TR></TABLE>"  
end sub

sub  TheRemainingMemberList
response.write "<TABLE border=0 cellpadding=0 cellspacing=0 width=130>"
response.write "<TR><Tdheight=20><B>"&langManageCommitteeMember2&"</B></Td></TR>"
response.write "<TR><TD>"
 response.write "<form name=r1 method='POST' action='SubscribeTableField_Add.asp?tablename="&tablename&"'>"
   response.write "<select size='15' WIDTH=""140"" STYLE=""width:140px"" name='OtherList' multiple>"
	   for i=0 to rs2.recordcount-1
	   memberstr=trim(rs2("name"))
	   if trim(memberstr)="" then memberstr="** "&rs2("name") &" **"
    	response.write "<option value="&rs2("name")&">"&memberstr&"</option>"
				rs2.movenext
		next
    	response.write "<option value='')>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>"
    response.write   "</select></p>"	 
     response.write  "</form>"
response.write "</TD></TR></TABLE>"      
end sub


%>