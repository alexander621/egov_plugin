<%
For each selectid in request.form("availablelist")
	response.write "<br>selectid="&selectid
Next
%>
