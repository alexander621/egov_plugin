<%
sSQL = "SELECT TOP 1 action_autoid FROM egov_actionline_requests ORDER BY action_autoid DESC"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	If not oRs.EOF then
		response.write "CURRENT COUNT: " & oRs("action_autoid") & " (" & 1000000 - oRs("action_autoid") & " remaining)<br />"
	end if

	oRs.Close
	Set oRs = Nothing
sSQL = "SELECT *, DATEADD(d,NumDaysLeft,GetDate()) as Midnight " _
	& " FROM  ( " _
	& " SELECT (COUNT(action_autoid)/365) as AvgPerDay, (1000000 - MAX(action_autoid)) / (COUNT(action_autoid)/365) as NumDaysLeft " _
	& " FROM egov_actionline_requests " _
	& " WHERE submit_date > DATEADD(d,-365,GetDate()) " _
	& " ) a "
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	If not oRs.EOF then
		response.write "Avg Per Day: " & oRs("AvgPerDay") & "<br />"
		response.write "Number of Days Until Midnight: " & oRs("NumDaysLeft") & "<br />"
		response.write "Midnight: " & oRs("Midnight") & "<br />"
	end if

	oRs.Close
	Set oRs = Nothing

%>
