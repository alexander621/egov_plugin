<%
response.buffer=true
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")

for each selectname in request.form("committeelist")
'response.write "<br>listvalue="&request.form("OtherList")
response.write "<br>selectname="&selectname
strSQL="delete from  SubScribedItems where tablename='"&trim(request.querystring("tablename"))&"' and tablefield='"&selectname&"'"
response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
'set conn=nothing
URL="ManageSubscribeTableFields.asp?tablename="&trim(request.querystring("tablename"))
'response.write url
response.redirect(url)
%>
