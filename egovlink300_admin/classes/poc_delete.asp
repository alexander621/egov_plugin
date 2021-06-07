<%
'--------------------------------------------------------------------------------------------------
' AUTHOR: Steve Loar
' CREATED: 05/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
	Dim sSql, oCmd
	
	sSQL = "DELETE FROM egov_class_pointofcontact WHERE pocid = " &  request("pocid") 
'	response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSQL
		.Execute
		.CommandText = "Update egov_class set pocid = NULL Where pocid = " & request("pocid")
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO poc management page
	response.redirect "poc_mgmt.asp"

%>
