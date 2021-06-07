<!-- #include file="../includes/common.asp" //-->
<%
On Error Resume Next
Dim conn,rs,strSQL,SubscribeID,howoften

if session("userid")=0 or session("userid")="" then
response.write "Session Expired"
response.end
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")

set rs = Server.CreateObject("ADODB.Recordset")
set rs.ActiveConnection = conn
rs.CursorLocation = 3 
rs.CursorType = 3 
strSQL = "select subscribeid,name,source from subscribe"
rs.Open strSQL,,, 2 


While Not rs.EOF
SubScribeID=rs("subscribeid")
howoften=request.form("howoften"&rs("subscribeid"))
userid=session("userid")

if lcase(howoften)="no"  or howoften="" then
strSQL = "delete from UserSubscribe where userid="&userid&" and subscribeid="&SubscribeID
else
strSQL="if (not exists(select * from UserSubscribe where userid="&userid&" and subscribeid="&SubscribeID&" )) "&vbcrlf 
strSQL =strSQL+ "insert UserSubscribe(subscribeID,howoften,userid) values("&SubscribeID&",'"&howoften&"',"&userid&")"&vbcrlf
strSQL =strSQL+ " else " &vbcrlf
strSQL =strSQL+ "update UserSubscribe set howoften='"&howoften&"' where userid="&userid&" and subscribeid="&SubscribeID&vbcrlf
end if
response.write "<br>"&strSQL
conn.execute(strSQL)
for each error1 in conn.Errors
response.write "<br>error="&error1.description
next
rs.movenext
wend
'Response.Redirect "../announcements"
%>