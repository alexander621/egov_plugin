<%
response.buffer=true
dim conn,strSQL,thisname,currentpage,pagesize,totalpages,delete,id
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
for each delete in request.form("delete")
'response.write delete
id=clng(delete)
strSQL = "delete from roles where roleid="&id
conn.execute(strSQL)
'response.write "<br><FONT COLOR=red>Delte Group with GroupID="&delete&"</FONT>"
next
'if request.form("delete").count=0 then response.write langNoDelete
'response.write "<br><a href='javascript:history.go(-1)'>"&langGoBack&"</a>"
response.write "<br>"
previousURL=request.querystring("previousURL")
if request.querystring("extra")<>"" then previousURL=previousURL&"?"&request.querystring("Extra")
'response.write previousURL
response.redirect(previousURL)
%>

