Steps 2 & 5: Read Cookies in Classic ASP
<br />
<%
response.write "The dupname cookie value is: " & request.cookies("dupname")
response.write "<hr>"
for each item in request.cookies
	
	response.write "<b>" & item & "</b>: " & request.cookies(item) & "<br />"
next


%>
<a href="cookie_final.aspx">Steps 3 & 6: Read Cookie in ASP.NET</a>
