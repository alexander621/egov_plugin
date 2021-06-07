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
Dim oCmd, sNewFlag

If GetIsClassSupervisor( request("userid") ) Then 
	sNewFlag = 0
Else
	sNewFlag = 1
End If 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")
oCmd.CommandText = "Update users set isclasssupervisor = " & sNewFlag & " where userid = " & request("userid")
oCmd.Execute
Set oCmd = Nothing

response.write request("userid") & ": " & sNewFlag 


'--------------------------------------------------------------------------------------------------
' Function GetIsClassSupervisor( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetIsClassSupervisor( iUserId )
	Dim sSql, oUser
	
	sSql = "Select isclasssupervisor from users where userid = " & iUserID

	Set oUser = Server.CreateObject("ADODB.Recordset")
	oUser.Open  sSQL, Application("DSN"), 3, 1

	If Not oUser.EOF Then 
		GetIsClassSupervisor = oUser("isclasssupervisor")
	Else
		GetIsClassSupervisor = False 
	End If 

	oUser.close
	Set oUser = Nothing 

End Function 
%>