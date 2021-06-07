<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: save_accounts.asp
' AUTHOR: Steve Loar
' CREATED: 02/15/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page saves changes to gl accounts
'
' MODIFICATION HISTORY
' 1.0   02/15/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oCmd, iAccountId, sAccountName, sAccountNumber, sNewStatus

Set oCmd = Server.CreateObject("ADODB.Command")

With oCmd
	.ActiveConnection = Application("DSN")
	
	If request("action") = "delete" Then
		' Delete the checked ones
		For Each iAccountId In Request("accountid")
			sSql = "Delete from egov_accounts where accountid = " & iAccountId
			response.write sSql & "<br />"
			.CommandText = sSql
			.execute
		Next 
	ElseIf request("action") = "deactivate" Then
		' Deactivate the checked ones
		For Each iAccountId In Request("accountid")
			If Request("accountstatus"&iAccountId) = "A" Then
				sNewStatus = "D"
			Else
				sNewStatus = "A"
			End If 
			sSql = "Update egov_accounts Set accountstatus = '" & sNewStatus & "' where accountid = " & iAccountId

			response.write sSql & "<br />"
			.CommandText = sSql
			.execute
		Next 
	ElseIf request("action") = "save" Then 
		' Save all records
		For Each iAccountId In Request("accountid")
			sSql = "Update egov_accounts Set accountname = '" & dbsafe(Request("accountname"&iAccountId)) & "', accountnumber = '" & dbsafe(Request("accountnumber"&iAccountId)) & "' where accountid = " & iAccountId
			response.write sSql & "<br />"
			.CommandText = sSql
			.execute
		Next 
	Else 
		' New account 
		sSql = "Insert Into egov_accounts ( orgid, accountname, accountnumber, accountstatus ) Values ( " & session("orgid") & ", '" & dbsafe(Request("accountname")) & "', '" & dbsafe(Request("accountnumber")) & "', 'A')"
		response.write sSql & "<br />"
		.CommandText = sSql
		.execute
	End If 
End with 

Set oCmd = Nothing

' Return to the account management
response.redirect "gl_account_mgmt.asp"


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	DBsafe = Replace( strDB, "'", "''" )
End Function


%>