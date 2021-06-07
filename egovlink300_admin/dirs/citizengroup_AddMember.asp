<%
response.buffer=true
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
if trim(request.querystring("groupid"))="" then
response.end
else
groupid=trim(request.querystring("groupid"))
end if

for each selectid in request.form("OtherList")
'response.write "<br>listvalue="&request.form("OtherList")
response.write "<br>selectid="&selectid
strSQL="insert into citizentogroups(citizenid,groupid) values("&CLng(selectid)&","&CLng(groupid)&")"
'response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
set conn=nothing
URL="ManageCitizenGroupMember.asp?groupid="&groupid
response.write url
response.redirect(url)
%>
