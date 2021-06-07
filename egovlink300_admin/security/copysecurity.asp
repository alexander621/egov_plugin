<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: copysecurity.ASP
' AUTHOR: Steve Loar
' CREATED: 10/02/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This copies the security of one user to another
'
' MODIFICATION HISTORY
' 1.0   10/02/2006	Steve Loar - INITIAL VERSION
' 1.1	03/03/2011	Steve Loar - Cleaned up and added RunSQLStatement call from common
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, item, iPermissionid, oRs

' Remove current security permissions
sSql = "DELETE FROM egov_users_to_features WHERE userid = " & CLng(request("iToUserID"))
'response.write sSql  & "<br />"

RunSQLStatement sSql


' Get Source permissions
sSql = "SELECT userid, featureid, permissionid, ISNULL(permissionlevelid,0) AS permissionlevelid "
sSql = sSql & "FROM egov_users_to_features WHERE userid = " & CLng(request("iFromUserID"))

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1
 
Do While Not oRs.EOF 
	' Insert new view permissions
	sSql = "INSERT INTO egov_users_to_features ( userid, featureid, permissionid, permissionlevelid ) VALUES ( " & CLng(request("iToUserID") )
	sSql = sSql & ", " & oRs("featureid") & ", "  & oRs("permissionid") & ", "
	If clng(oRs("permissionlevelid")) > 0 Then 
		sSql = sSql & oRs("permissionlevelid")
	Else
		sSql = sSql & "NULL"
	End If 
	sSql = sSql & " )"

	'response.write sSql  & "<br />"

	RunSQLStatement sSql

	oRs.MoveNext 
Loop 

oRs.Close
Set oRs = Nothing 

response.redirect "edit_user_security.asp?s=u&iUserId=" & CLng(request("iToUserID"))


%>
