<%

if request.cookies("preview") = "true" then
	response.cookies("preview") = ""
	response.redirect "action_line_list.asp"
else
	response.cookies("preview") = "true"
	response.cookies("preview").Expires = dateadd("d",0, "7/4/2018")
	response.redirect "rd_action_line_list.asp"
end if



%>
