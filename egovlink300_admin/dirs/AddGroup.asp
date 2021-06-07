<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
if trim(request.querystring("userid"))="" then
response.write "<br>No User ID is entered, end program here"
response.end
else
userid=trim(request.querystring("userid"))
end if

for each selectid in request.form("RemainingList")
'response.write "<br>listvalue="&request.form("RemainingList")
response.write "<br>selectid="&selectid
strSQL="insert into UsersGroups(userid,GroupID) values("&clng(userid)&","&clng(selectid)&")"
'response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
set conn=nothing
URL="ManageMemberGroup.asp?userid="&userid
response.write url
response.redirect(url)
%>
