<%

sSQL = "SELECT u.userfname + ' ' + u.userlname as Name, ad.answer as Address, MIN(a.parcelidnumber) as ParcelID, " _
		& " sd.answer as StartDate, DATEADD(d,29,CONVERT(datetime,sd.answer)) as EndDate, rm.answer as RemovalType " _
	& " FROM egov_actionline_requests ar " _
	& " INNER JOIN egov_users u ON u.userid = ar.userid " _
	& " INNER JOIN action_submitted_questions_and_answers ad ON ad.action_autoid = ar.action_autoid and ad.question = 'Address' " _
	& " INNER JOIN egov_residentaddresses a ON ad.answer = a.residentstreetnumber + ' ' + a.residentstreetname " _
	& " INNER JOIN action_submitted_questions_and_answers sd ON sd.action_autoid = ar.action_autoid and sd.question = 'Start Date' " _
	& " INNER JOIN action_submitted_questions_and_answers rm ON rm.action_autoid = ar.action_autoid and rm.question = 'Type of Removal' " _
	& " WHERE ar.category_id = '17890' and DATEADD(d,29,CONVERT(datetime,sd.answer)) >= GETDATE() AND ar.status = 'RESOLVED' " _
	& " GROUP BY a.residentstreetname,a.residentstreetnumber, u.userfname, u.userlname, ad.answer, sd.answer, rm.answer " _
	& " ORDER BY a.residentstreetname,a.residentstreetnumber " 
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN") , 3, 1
%>
<table cellpadding="3" cellspacing="0" border="1">
	<tr>
		<th>Address</th>
		<th>Parcel ID (S/B/L)</th>
		<th>Rock Removal Type</th>
		<th>Start Date</th>
		<th>End Date</th>
	</tr>
	<%
	Do While Not oRs.EOF
		%><tr><td><%=oRs("Address")%></td><td><%=oRs("ParcelID")%></td><td><%=oRs("RemovalType")%></td><td><%=oRs("StartDate")%></td><td><%=oRs("EndDate")%></td></tr><%
		oRs.MoveNext
	loop
	%>
</table>
<%
oRs.Close
Set oRs = Nothing
%>
