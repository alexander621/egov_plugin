<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setrentalsupervisor.asp
' AUTHOR: Steve Loar	
' CREATED: 08/17/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the rental supervisors for admin users
'
' MODIFICATION HISTORY
' 1.0   08/17/2009   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, sNewFlag, sUserId

sUserId = CLng(request("userid"))

If IsRentalSupervisor( sUserId ) Then 
	sNewFlag = 0
Else
	sNewFlag = 1
End If 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")
oCmd.CommandText = "UPDATE users SET isrentalsupervisor = " & sNewFlag & " WHERE userid = " & sUserId
oCmd.Execute
Set oCmd = Nothing

'response.write request("userid") & ": " & sNewFlag 


'--------------------------------------------------------------------------------------------------
' Function IsRentalSupervisor( iUserId )
'--------------------------------------------------------------------------------------------------
Function IsRentalSupervisor( iUserId )
	Dim sSql, oRs
	
	sSql = "SELECT isrentalsupervisor FROM users WHERE userid = " & iUserID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		IsRentalSupervisor = oRs("isrentalsupervisor")
	Else
		IsRentalSupervisor = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>