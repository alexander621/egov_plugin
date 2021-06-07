<%
Session.Abandon
Response.Cookies("User")("UserID") = ""
Response.Cookies("User")("FullName") = ""
Response.Cookies("User").Expires = Now() - 1
Response.Redirect "default.asp"
%>