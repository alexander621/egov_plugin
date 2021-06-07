<%

	querystring = ""
	if request.servervariables("query_string") <> "" then querystring = "?googleauth=true&" & request.servervariables("query_string")
if Split(Request.ServerVariables("SCRIPT_NAME"), "/")(1) <> "eclink" then
	response.redirect "basic_login.asp" & querystring
else
	response.redirect "basic_chooseorg.asp" & querystring
end if
%>
