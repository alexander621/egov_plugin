<%
Dim conn, strSQL, selectid


' INITIALIZE VALUES
thisname = request.servervariables("script_name")



' CHECK FOR VALID CITIZEN ID PASSED VIA QUERYSTRING
If trim(request.querystring("userid")) = "" then
	response.write "<br>Error: Missing citizen ID!"
	response.end
Else
	userid = trim(request("userid"))
End If

set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")

' LOOP THRU EACH SELECTED GROUP AND REMOVE MEMBER FROM LIST
For each selectid in request("ExistingList")
	' DEBUG CODE: response.write "<br>selectid="&selectid
	strSQL = "DELETE FROM CitizentoGroups WHERE Groupid = " & CLng(selectid) & " AND citizenid = " & CLng(userid)
	' DEBUG CODE:response.write "<br>strSQL="&strSQL
	conn.execute strSQL, lngRecs
Next


' CLEAN UP OBJECTS
conn.close
Set conn = Nothing 


' REDIRECT TO REFRESH LIST
URL = "ManageCitizenGroup.asp?userid=" & userid
'response.write url
response.redirect(url)

%>
