<%
session("jobbid_id") = request.querystring("id")
session("jobbid_title") = request.querystring("title")
session("email_dlids") = ""

response.redirect "dl_sendmail.asp?listtype=" & request.querystring("type") & "&screen_mode=AUTOSEND"
%>
