<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setpermitreviewer.asp
' AUTHOR: Steve Loar	
' CREATED: 01/18/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the permit reviewers for admin users
'
' MODIFICATION HISTORY
' 1.0   01/18/2008   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, sNewFlag

If GetIsPermitReviewer( CLng(request("userid")) ) Then 
	sNewFlag = 0
Else
	sNewFlag = 1
End If 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")
oCmd.CommandText = "UPDATE users SET ispermitreviewer = " & sNewFlag & " WHERE userid = " & request("userid")
oCmd.Execute
Set oCmd = Nothing

response.write request("userid") & ": " & sNewFlag 


'--------------------------------------------------------------------------------------------------
' Function GetIsPermitReviewer( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetIsPermitReviewer( iUserId )
	Dim sSql, oUser
	
	sSql = "SELECT ispermitreviewer FROM users WHERE userid = " & iUserID

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open  sSQL, Application("DSN"), 3, 1

	If Not oUser.EOF Then 
		GetIsPermitReviewer = oUser("ispermitreviewer")
	Else
		GetIsPermitReviewer = False 
	End If 

	oUser.close
	Set oUser = Nothing 

End Function 


%>