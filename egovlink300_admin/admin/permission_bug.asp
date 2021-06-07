<%
	Dim sSql, oUsers
	Dim iCOunt
	iCount = 0

	sSql = "Select userid from egov_users_to_features where featureid = 43 "

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSQL, Application("DSN"), 0, 1

	Do While Not oUsers.EOF
		If UserDoesNotHaveIt( oUsers("userid") ) Then 
			iCount = iCount + 1
			InsertPermission oUsers("userid")
		End If 
		oUsers.movenext
	Loop 

	oUsers.close
	Set oUsers = Nothing
	response.write "<br />Inserted: " & iCount



Function UserDoesNotHaveIt( iUserId )
	Dim sSql, oFeature

	sSql = "Select count(permissionid) as hits from egov_users_to_features where featureid = 125 and userid = " & iUserId

	Set oFeature = Server.CreateObject("ADODB.Recordset")
	oFeature.Open sSQL, Application("DSN"), 0, 1

	If Not oFeature.EOF Then
		If clng(oFeature("hits")) = 0 Then 
			UserDoesNotHaveIt = True 
		Else
			UserDoesNotHaveIt = False 
		End If 
	Else
		UserDoesNotHaveIt = True 
	End If 

	oFeature.close
	Set oFeature = Nothing

End Function 



Sub InsertPermission( iUserId ) 
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		response.write "<br />Inserting : " & iUserId
		sSql = "Insert INTO egov_users_to_features ( featureid, permissionid, userid ) values ( 125, 1, " & iUserId & " )"
		.CommandText = sSql
		.execute
	End With
	Set oCmd = Nothing
End Sub 
%>	



