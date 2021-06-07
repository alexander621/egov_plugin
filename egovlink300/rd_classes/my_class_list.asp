<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="utf-8" />
</head>
<body>
<%

	Dim sSql, oRs

	' GET SELECTED class INFORMATION
	' Display to public added - SJL 5/19/2006
	sSql = "SELECT classid, classtypeid, isparent, classname, categorytitle, classdescription, " 
	sSql = sSql & " ISNULL(startdate,0) AS startdate, ISNULL(enddate,0) AS enddate "
	sSql = sSql & " FROM egov_class_to_categories WHERE "
	sSql = sSql & " (('" & date() & "' BETWEEN publishstartdate AND publishenddate) OR noenddate = 1) "
	sSql = sSql & " AND displaytopublic = 1 AND statusname = 'ACTIVE' AND orgid = 60 " 
	sSql = sSql & " ORDER BY classname, noenddate DESC, startdate, isparent DESC"

	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Response.write "<table border=""1"" cellspacing=""0"" cellpadding=""0"">"
	' DISPLAY class INFORMATION
	Do while Not oRs.EOF 
	
		Response.Write vbcrlf & "<tr><td>" & oRs("classid") & "</td><td>" & oRs("classname") & "</td><td>" & oRs("classdescription") & "</td></tr>"
		
		response.flush

		oRs.MoveNext
	Loop
	Response.write "</table>"

	' CLOSE OBJECTS
	oRs.Close
	Set oRs = Nothing 

%>
</body>
</html>