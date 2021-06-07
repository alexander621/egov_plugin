<%
'if instr(request.servervariables("HTTP_REFERER"),"wp-admin") > 0 then
	url = replace(lcase(request.servervariables("HTTP_REFERER")),"http://","")
	url = "http://" & left(url,instr(url,"/"))
	url = DBSafe(url)

	'response.write "Site: " & url
	'response.write "<br />"
	sSQL = "SELECT orgid,OrgEgovWebsiteURL FROM Organizations WHERE OrgPublicWebsiteURL = '" & url & "'"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then 
		redirectionURL = oRs("OrgEgovWebsiteURL")
		if right(redirectionURL,1) <> "/" then redirectionURL = redirectionURL & "/"
		redirectionURL = redirectionURL & "admin/login.asp"
		'response.write "E-Gov URL: " & redirectionURL & "<br />"
		'response.write "Email: " & request.querystring("userid")
		'response.redirect redirectionURL
		'response.write "<meta http-equiv=""refresh"" content=""0; url=" & redirectionURL & """>"
		sSQL = "SELECT username,password FROM users WHERE email = '" & dbsafe(request.querystring("userid")) & "' and orgid = '" & oRs("orgid") & "'"
		oRs.Close
		oRs.Open sSQL, Application("DSN"), 3, 1
		if not oRs.EOF then %>
		<form action="<%=redirectionURL%>" method="POST" name="frm">
			<input type="hidden" name="_task" value="login" />
			<input type="hidden" name="username" value="<%=oRs("username")%>" />
			<input type="hidden" name="password" value="<%=oRs("password")%>" />
		</form>
		<script>
			document.frm.submit();
			//window.history.back();
		</script>
		<%else %>
			<h1>Sorry, we could not find an E-Gov user that matched your WordPress email address</h1>
		<% end if %>
		<%
	end if
	oRs.Close
	Set oRs = Nothing

'else
	'for each item in request.servervariables
		'response.write "<b>" & item & "</b> = " & request.servervariables(item) & "<br />"
	'next
'end if

Function DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString
End Function

%>
