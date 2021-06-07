<!--#include file='dir_constants.asp'-->
<%
response.buffer=true
sGroupID= request.querystring("groupid")
response.write "<br>groupid="&groupid
With request
for each name in Request.Form
'response.write "<br>name="&name&" value="&.form(name)
AccessedGroupid=replace(name,"field_","")
iname=clng(.form(name))
'response.write "<br>*sGroupID="&sGroupID &"  accessedgroupid="&AccessedGroupid
'Do some verification here"
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
str="EXEC NewGroupPermissions "&clng(sGroupID)&","&clng(AccessedGroupid)&","&iname
'response.write "<br>"&str
conn.execute str
Next 
end with
conn.close
set conn=nothing
strMessage=langSucessUpdate
url=request.querystring("url")&"?strMessage="&strMessage&"&groupid="&request.querystring("groupid")
response.redirect url
%>


