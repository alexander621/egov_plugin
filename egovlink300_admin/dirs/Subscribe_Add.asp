<%
response.buffer=true
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")

for each selectname in request.form("OtherList")
'response.write "<br>listvalue="&request.form("OtherList")
'response.write "<br>selectname="&selectname
strSQL="insert into subscribe(name) values('"&selectname&"')"
'response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
set conn=nothing
URL="ManageSubscribe.asp"
'response.write url
response.redirect(url)
%>
