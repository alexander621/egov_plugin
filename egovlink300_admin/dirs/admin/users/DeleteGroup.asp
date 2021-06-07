<%
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
if trim(request.querystring("userid"))="" then
response.write "<br>No Role ID is entered, end program here"
response.end
else
userid=trim(request.querystring("userid"))
end if

for each selectid in request.form("ExistingList")
'response.write "<br>listvalue="&request.form("OtherList")
response.write "<br>selectid="&selectid
strSQL="delete from UsersGroups where Groupid="&clng(selectid)&" and userid="&clng(userid)
response.write "<br>strSQL="&strSQL
conn.execute strSQL, lngRecs
next
conn.close
set conn=nothing
URL="ManageGroup.asp?userid="&userid
response.write url
response.redirect(url)
%>
