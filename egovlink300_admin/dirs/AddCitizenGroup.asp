<%
' INITIALIZE VALUES
thisname=request.servervariables("script_name")
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")


' CHECK FOR VALID CITIZEN ID PASSED VIA QUERYSTRING
If trim(request.querystring("userid"))="" then
	response.write "<br>Error: Missing citizen ID!"
	response.end
Else
	userid=trim(request.querystring("userid"))
End If


' LOOP THRU EACH SELECTED GROUP AND ADD USER
For each selectid in request.form("RemainingList")
	'DEBUG CODE: response.write "<br>selectid="&selectid
	strSQL="insert into CitizentoGroups(citizenid,GroupID) values(" & CLng(userid) & "," & CLng(selectid)&")"
	'DEBUG CODE: response.write "<br>strSQL="&strSQL
	conn.execute strSQL, lngRecs
Next


' CLEAN UP OBJECTS
conn.close
set conn=nothing


' REDIRECT TO REFRESH LIST
URL="ManageCitizenGroup.asp?userid="&userid
response.write url
response.redirect(url)
%>
