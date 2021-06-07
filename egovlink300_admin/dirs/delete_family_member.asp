<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: delete_family_member.asp
' AUTHOR: Steve Loar
' CREATED: 1/3/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This Deletes a family member
'
' MODIFICATION HISTORY
' 1.0   1/3/2007	Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserId, oCmd, sReturnFlag

iUserId = CLng(request("iUserId"))

If hasAccountBalance( iUserId ) Then
	sReturnFlag = "deletefail"
Else
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")

		.CommandText = "UPDATE egov_users SET isdeleted = 1, deleteddate = GETDATE(), deletedbyuserid = " & session("UserID") & " WHERE userid = " & iUserId
		.Execute

		.CommandText = "UPDATE egov_familymembers SET isdeleted = 1 WHERE userid = " & iUserId
		.Execute
	End With

	Set oCmd = Nothing
	sReturnFlag = "deletesuccess"
End If 

response.redirect "family_list.asp?userid=" & request("iReturn") & "&status=" & sReturnFlag



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