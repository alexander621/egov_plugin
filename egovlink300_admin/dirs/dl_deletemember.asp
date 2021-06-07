<%
For each selectid in request.form("committeelist")
	response.write "<br>selectid="&selectid
Next
%>
