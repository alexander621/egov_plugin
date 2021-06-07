<%
response.buffer=true
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")

for each selectname in request.form("OtherList")
'response.write "<br>listvalue="&request.form("OtherList")
'response.write "<br>selectname="&selectname
strSQL="insert into SubScribedItems(tablename,tablefield) values('"&trim(request.querystring("tablename"))&"','"&selectname&"')"
response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
set conn=nothing
URL="ManageSubscribeTableFields.asp?tablename="&trim(request.querystring("tablename"))
'response.write url
response.redirect(url)
%>
