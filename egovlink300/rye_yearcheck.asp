<%
strAddress = DBSafe(request.querystring("address"))
strStartDate = DBSafe(request.querystring("date"))

sSQL = "SELECT ad.answer as Address,sd.answer as StartDate, DATEADD(yyyy,1,DATEADD(d,29,CONVERT(datetime,sd.answer))) as EndDate " _
	& " FROM egov_actionline_requests ar  " _
	& " INNER JOIN egov_users u ON u.userid = ar.userid  " _
	& " INNER JOIN action_submitted_questions_and_answers ad ON ad.action_autoid = ar.action_autoid and ad.question = 'Address'  " _
	& " INNER JOIN action_submitted_questions_and_answers sd ON sd.action_autoid = ar.action_autoid and sd.question = 'Start Date'  " _
	& " WHERE ar.category_id = '17890' AND ar.status = 'RESOLVED' " _
	& " and ad.answer = '" & strAddress & "' " _
	& " AND sd.answer <= '" & strStartDate & "' " _
	& " AND DATEADD(yyyy,1,CONVERT(datetime,sd.answer)) >= '" & strStartDate & "'  " _
	& " ORDER BY sd.answer DESC"


Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN") , 3, 1
if not oRs.EOF then
	response.write DateAdd("yyyy",1,oRs("StartDate"))
else
	response.write "PASS"
end if
oRs.Close
Set oRs = Nothing

Function DBsafe( ByVal strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	DBsafe = sNewString
End Function
%>
