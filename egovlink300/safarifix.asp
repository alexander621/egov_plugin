<%
response.cookies("dummycookie") = "something"
if request.servervariables("HTTP_REFERER") <> "" then response.redirect request.servervariables("HTTP_REFERER")

%>
