<% Response.AddHeader "Access-Control-Allow-Origin","*"  %>
<%
response.flush

	sSQL = "SELECT r.[OrgID] ,category_id, action_autoid ,[mobileoption_latitude] ,[mobileoption_longitude] ,f.action_form_name " _
  		& " FROM egov_actionline_requests r " _
  		& " INNER JOIN egov_action_request_forms f ON f.action_form_id = r.category_id " _
  		& " WHERE r.orgid = " & dbsafe(request("orgid")) & " and r.category_id = '" & dbsafe(request("actionid")) & "' and mobileoption_latitude IS NOT NULL and mobileoption_latitude <> 0 " _
  		& " AND status IN ('SUBMITTED','IN PROGRESS','WAITING') "
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then
    		response.write "["
		strJS = ""
		Do While Not oRs.EOF
       			strJS = strJS & "["""", " & oRs("mobileoption_latitude") & "," & oRs("mobileoption_longitude") & "],"
			oRs.MoveNext
		loop
		response.write left(strJS,len(strJS)-1)
   		response.write "]"
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

