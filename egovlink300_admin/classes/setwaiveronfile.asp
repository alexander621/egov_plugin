<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: setwaiveronfile.asp
' AUTHOR: Steve Loar	
' CREATED: 05/04/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sets the waivers on file flag for students
'
' MODIFICATION HISTORY
' 1.0   05/04/2007   Steve Loar - Initial code  
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oCmd, sNewFlag

If WaiverIsOnFile( request("classlistid") ) Then 
	sNewFlag = 0
Else
	sNewFlag = 1
End If 

Set oCmd = Server.CreateObject("ADODB.Command")
oCmd.ActiveConnection = Application("DSN")
oCmd.CommandText = "Update egov_class_list set waiveronfile = " & sNewFlag & " where classlistid = " & request("classlistid")
oCmd.Execute
Set oCmd = Nothing

response.write request("classlistid") & ": " & sNewFlag 


'--------------------------------------------------------------------------------------------------
' Function WaiverIsOnFile( iClassListId )
'--------------------------------------------------------------------------------------------------
Function WaiverIsOnFile( iClassListId )
	Dim sSql, oWaiver
	
	sSql = "Select waiveronfile from egov_class_list where classlistid = " & iClassListId

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open  sSQL, Application("DSN"), 3, 1

	If Not oWaiver.EOF Then 
		WaiverIsOnFile = oWaiver("waiveronfile")
	Else
		WaiverIsOnFile = False 
	End If 

	oWaiver.close
	Set oWaiver = Nothing 

End Function 
%>