<% Response.AddHeader "Access-Control-Allow-Origin","*"  %>
<%
on error resume next

'check to see if this IP has upvoted
'sSQL = "SELECT * FROM egov_actionline_requests_upvotelog WHERE IP = '" & dbsafe(request.servervariables("REMOTE_ADDR")) & "' AND action_autoid = '" & dbsafe(request.querystring("id")) & "'"
sSQL = "SELECT action_autoid, IP FROM egov_actionline_requests_upvotelog WHERE IP = '" & dbsafe(request.servervariables("REMOTE_ADDR")) & "' AND action_autoid = '" & dbsafe(request.querystring("id")) & "' " _
	& " UNION ALL " _
	& " SELECT action_autoid, submittedby_remoteaddress FROM egov_actionline_requests WHERE submittedby_remoteaddress = '" & dbsafe(request.servervariables("REMOTE_ADDR")) & "' AND action_autoid = '" & dbsafe(request.querystring("id")) & "'"


Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
if oRs.EOF then


	sSQL = "UPDATE egov_actionline_requests SET upvotes = upvotes + 1 WHERE action_autoid = '" & dbsafe(request.querystring("id")) & "'"
	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.Open Application("DSN")
	oCmd.Execute(sSQL)

	sSQL = "INSERT INTO egov_actionline_requests_upvotelog (IP, action_autoid) VALUES('" & dbsafe(request.servervariables("REMOTE_ADDR")) & "','" & dbsafe(request.querystring("id")) & "')"
	oCmd.Execute(sSQL)
	
	oCmd.Close
	Set oCmd = Nothing
	If Err.Number <> 0 Then
		response.write "FAIL"
		response.end
	end if
	on error goto 0
end if
	
response.write "success"
	
Function DBsafe( ByVal strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	DBsafe = sNewString
End Function

%>
