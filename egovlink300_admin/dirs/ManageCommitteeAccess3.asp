<!--#include file='header_simply.asp'-->
<%
dim conn,rs,strSQL,thisname,cmd,groupid,delimeter,sGroups,strResult,sBgcolor,sChecked0,sChecked1,sChecked2,bDisabled,CommitteeName,ResultID
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL="select groupname from groups g where g.orgid="&Session("OrgID")&" and g.groupid="&clng(request.querystring("groupid"))
rs.Open strSQL
CommitteeName=rs("groupname")
rs.close
set rs=nothing
%>
<br>
<body onload="javasript:opener.location.reload(true);">
<CENTER><FONT SIZE="+1" COLOR=""><B>Update Directory abilities for :<%=CommitteeName%></B></FONT><br>
<a href='javascript:self.close(-1)'>Close Window</a></CENTER>
<TABLE  cellpadding=5 width=400 cellspacing=0 align=center>
<TR>
	<TD width='90%'>
	

<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
set cmd=Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection=conn
cmd.commandtext="GetDirectoryPermissionType"
cmd.commandtype=&H0004

set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 2 
'strSQL = "select g.groupname, sum(u.userid) as entry from groups g, usersgroups as u where u.groupid=g.groupid group by g.groupname"
strSQL = "select groupid,groupname from groups where orgid="&Session("OrgID")&"  order by groupname"
rs.Open strSQL,,, 2 
'------- the following code dealing with the recordcount=0-------
if rs.recordcount>0 then
call displayform
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
set cmd=nothing
'--------------------------------------------------------------------
sub displayform
groupid=request.querystring("groupid")
rs.movefirst
delimeter="_qwcdfsddewrf45435ds_"
If sBgcolor = "#ffffff" Then sBgcolor = "#eeeeee" Else sBgcolor = "#ffffff"
strResult=request.querystring("strMessage")
While Not rs.EOF
sChecked0="" 
sChecked1="" 
sChecked2="" 
bDisabled=""
cmd.Parameters.Refresh
cmd.parameters(1)=groupid
cmd.parameters(2)=rs("groupid")
cmd.execute
ResultID=cmd.parameters(0)
select case ResultID
case 0
sChecked0="checked"
case 1 
sChecked1="checked" 
case 2 
sChecked2="checked" 
case 3
sChecked2="checked" 
bDisabled="Disabled"
end select
sGroups = sGroups & "<tr bgcolor=""" & sBgcolor & """><td>"&rs("groupname")&"</td>"&vbcrlf
sGroups = sGroups & "<td><input type=radio "&bDisabled&" name=""field_"&rs("groupid")&""" "&sChecked0&" value=""0"">&nbsp;</td>"&vbcrlf
sGroups = sGroups & "<td><input type=radio "&bDisabled&" name=""field_"&rs("groupid")&""" "&sChecked1&" value=""1"">&nbsp;</td>"&vbcrlf
sGroups = sGroups & "<td><input type=radio "&bDisabled&" name=""field_"&rs("groupid")&""" "&sChecked2&" value=""2"">&nbsp;</td></tr>"&vbcrlf
rs.movenext
wend
response.write "<form name=""ManageCommitteAccess"" action=""Insert_GroupPermissions.asp?url="&thisname&"&groupid="&groupid&""" method=""post"" >"
'response.write strAction
response.write "<B><FONT COLOR=red>"&strResult&"</FONT></B>"
response.write "<TABLE cellpadding=5 width=400 cellspacing=0 class=""tablelist"">"
response.write "<th align=""left"">Group</th>"
response.write "<th align=""left"">Hidden</th>"
response.write "<th align=""left"">View</th>"
response.write "<th align=""left"">Edit</th>"
response.write  sGroups
response.write "</TABLE>" 
response.write "<CENTER><INPUT TYPE='submit' value='Update'></CENTER>"
'response.write strAction
response.write "</form>"
 end sub 
 %>
 
	</TD>
</TR>
</TABLE>
</body>
