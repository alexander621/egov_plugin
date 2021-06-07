
<%
if request("type") = "glaccounts" then ShowAccounts

Sub ShowAccounts( )
	Dim sSql, oRs

	sSql = "SELECT accountid, accountname, ISNULL(accountnumber,'') AS accountnumber FROM egov_accounts WHERE orgid = " & session("orgid" )
	sSql = sSql & " AND accountstatus = 'A' ORDER BY accountname"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<option value=""0"" selected=""selected"" >Select an Account</option>"
		Do While Not oRs.EOF 
			response.write vbcrlf & "<option value=""" & oRs("accountid") & """ "
			response.write " >" & oRs("accountname")
			If oRs("accountnumber") <> "" Then
				response.write " (" & oRs("accountnumber") & ")"
			End If 
			response.write "</option>"
			oRs.MoveNext 
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 
%>
