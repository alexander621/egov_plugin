<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setpermitinspector.asp
' AUTHOR: Steve Loar	
' CREATED: 01/17/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the permit inspectors for admin users
'
' MODIFICATION HISTORY
' 1.0   01/17/2008   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, sNewFlag

If GetIsPermitInspector( CLng(request("userid")) ) Then 
	sNewFlag = 0
Else
	sNewFlag = 1
End If 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")
oCmd.CommandText = "UPDATE users SET ispermitinspector = " & sNewFlag & " WHERE userid = " & request("userid")
oCmd.Execute
Set oCmd = Nothing

response.write request("userid") & ": " & sNewFlag 


'--------------------------------------------------------------------------------------------------
' Function GetIsPermitInspector( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetIsPermitInspector( iUserId )
	Dim sSql, oUser
	
	sSql = "SELECT ispermitinspector FROM users WHERE userid = " & iUserID

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open  sSQL, Application("DSN"), 3, 1

	If Not oUser.EOF Then 
		GetIsPermitInspector = oUser("ispermitinspector")
	Else
		GetIsPermitInspector = False 
	End If 

	oUser.close
	Set oUser = Nothing 

End Function 
%>