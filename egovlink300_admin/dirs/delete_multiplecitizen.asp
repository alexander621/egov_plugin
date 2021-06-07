<%
response.buffer=true
dim conn,strSQL,thisname,currentpage,pagesize,totalpages,delete,id
set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("DSN")
for each delete in request.form("delete")
	'response.write delete
	id = CLng( delete )
	If Not hasAccountBalance( id ) Then 
		'strSQL = "delete from egov_users where userid="&id
		strSQL = "UPDATE egov_users SET isdeleted = 1, deleteddate = GETDATE(), deletedbyuserid = " & session("UserID") & " WHERE userid = " & id
		conn.execute(strSQL)
		' delete the family members row.
		strSQL = "UPDATE egov_familymembers SET isdeleted = 1 WHERE userid = " & id
		conn.execute(strSQL)
	End If 
Next 

Set conn = Nothing 

'response.write "<br>"
previousURL = request.querystring("previousURL")
if request.querystring("extra") <> "" then previousURL = previousURL & "?" & request.querystring("Extra")
response.redirect(previousURL)


'--------------------------------------------------------------------------------------------------
'  boolean hasAccountBalance( userId )
'--------------------------------------------------------------------------------------------------
Function hasAccountBalance( ByVal userId )
	Dim sSql, oRs
	
	sSql = "SELECT ISNULL(accountbalance,0) AS accountbalance FROM egov_users WHERE userid = " & userId
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		If CDbl(oRs("accountbalance")) <> CDbl(0) Then
			hasAccountBalance = true
		Else
			hasAccountBalance = false
		End If 
	Else 
		hasAccountBalance = false 
	End If 
	
	oRs.Close
	Set oRs = Nothing
	
End Function 


%>

