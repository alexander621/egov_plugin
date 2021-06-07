<!-- #include file="../includes/common.asp" //-->
<%

Dim sSql, oRs

' pull the docs with more than one record in the database
sSql = "select count(documenturl) AS hits, documenturl, orgid, parentfolderid from documents where documenturl in (select documenturl from documents) "
sSql = sSql & " group by documenturl, orgid, parentfolderid having count(documenturl) > 1 "
sSql = sSql & " order by orgid, parentfolderid, documenturl"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

Do While Not oRs.EOF
	CleanDocs oRs("orgid"), oRs("documenturl"), oRs("parentfolderid")
	oRs.MoveNext 
Loop 

oRs.CLose
Set oRs = Nothing 



Sub CleanDocs( sOrgId, sDocURL, sParentFolderId )
	Dim sSql, oRs

	sSql = "SELECT documentid FROM documents WHERE orgid = " & sOrgId & " AND parentfolderid = " & sParentFolderId & " AND documenturl = '" & sDocURL & "' ORDER BY documentid"
	response.write "<b>" &  sSql & "</b><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	' Skip the first one. We want to keep that one
	response.write "Keeping: " & oRs("documentid") & "<br />"
	oRs.MoveNext
	Do While Not oRs.EOF
		' Delete the file
		sSql = "DELETE FROM documents WHERE documentid = " & oRs("documentid")
		response.write sSql & "<br />"
		RunSQLStatement sSql 
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	response.write "<br />"
	response.flush 
End Sub 

%>