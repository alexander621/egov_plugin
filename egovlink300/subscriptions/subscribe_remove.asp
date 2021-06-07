<html>
<body onload="showCoords(event)">
	Are you sure you want to unsubscribe from this list?
	<br />
	<button onclick="window.location='process_subscribe_remove.asp?<%=request.servervariables("Query_String")%>'">YES</button>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<button onclick="window.close()">NO</button>
	
</body>
</html>
