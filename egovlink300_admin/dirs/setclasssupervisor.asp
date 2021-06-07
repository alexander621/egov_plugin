<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setclasssupervisor.asp
' AUTHOR: Steve Loar	
' CREATED: 02/21/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the class supervisor for users
'
' MODIFICATION HISTORY
' 1.0   02/21/2007   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, sNewFlag, sUserId

sUserId = CLng(request("userid"))

If IsClassSupervisor( sUserId ) Then 
	sNewFlag = 0
Else
	sNewFlag = 1
End If 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")
oCmd.CommandText = "UPDATE users SET isclasssupervisor = " & sNewFlag & " WHERE userid = " & sUserId
oCmd.Execute
Set oCmd = Nothing

'response.write request("userid") & ": " & sNewFlag 


'--------------------------------------------------------------------------------------------------
' Function IsClassSupervisor( iUserId )
'--------------------------------------------------------------------------------------------------
Function IsClassSupervisor( iUserId )
	Dim sSql, oRs
	
	sSql = "SELECT isclasssupervisor FROM users WHERE userid = " & iUserID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		IsClassSupervisor = oRs("isclasssupervisor")
	Else
		IsClassSupervisor = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>