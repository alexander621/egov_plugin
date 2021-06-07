<%

code = request.querystring("code")


sSQL = "SELECT linktourl FROM linktracker WHERE linkcode = '" & code & "'"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
if not oRs.EOF then
	sSQL = "UPDATE linktracker SET clickeddate = '" & now() & "'"
	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.Open Application("DSN")
	oCmd.Execute(sSQL)
	oCmd.Close
	Set oCmd = Nothing

	response.redirect oRs("linktourl")
else
	response.write "Sorry, your link is broken."
end if
oRs.Close
Set oRs = Nothing
%>
