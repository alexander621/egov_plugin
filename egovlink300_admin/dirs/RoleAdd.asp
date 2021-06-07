<%
response.buffer=true
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
if trim(request.querystring("roleid"))="" then
response.write "<br>No Role ID is entered, end program here"
response.end
else
roleid=trim(request.querystring("roleid"))
end if

for each selectid in request.form("RemainingList")
'response.write "<br>listvalue="&request.form("RemainingList")
response.write "<br>selectid="&selectid
strSQL="insert into rolespermissions(roleid,permissionid) values("&clng(roleid)&","&clng(selectid)&")"
'response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
set conn=nothing
URL="ManageRolePermission.asp?roleid="&roleid
response.write url
response.redirect(url)
%>
