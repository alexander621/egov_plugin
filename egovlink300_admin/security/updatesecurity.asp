<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: updatesecurity.ASP
' AUTHOR: Steve Loar
' CREATED: 09/28/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This saves the changes to a users permissions 
' This approach allows the retension of Root Admin assignable features, which get deleted in a blanket delete
'
' MODIFICATION HISTORY
' 1.0	09/28/2006	Steve Loar - INITIAL VERSION
' 1.1	12/07/2006	Steve Loar - Changed to loop through features on delete
' 1.2	08/12/2009 David Boyer - Added "screen msg" parameters to redirect url
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, item, iPermissionid, iFeatureId, oRs

'Get the features the user has now that have permissions
sSql = "SELECT U.featureid "
sSql = sSql & " FROM egov_users_to_features U, egov_organization_features O " 
sSql = sSql & " WHERE U.featureid = O.featureid "
sSql = sSql & " AND O.haspermissions = 1 "
sSql = sSql & " AND U.userid = " & request("iUserId")

'If they are not the root admin, then do not pull those features that only a root admin can assign
If Not request("bIsRootAdmin") Then 
	sSql = sSql & " AND O.rootadminrequired = 0 "
End If 

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

Do While Not oRs.EOF
	'Delete those features that were pulled
	DeleteUserFeature request("iUserId"), oRs("featureid")
	oRs.MoveNext
Loop 

oRs.Close
Set oRs = Nothing 

'Insert selected user features
iPermissionid = GetPermissionId("permission")

For Each item In request("viewPermission")
	sSql = "INSERT INTO egov_users_to_features (userid, featureid, permissionid) values ( "
	sSql = sSql & request("iUserId") & ", "
	sSql = sSql & item & ", " 
	sSql = sSql & iPermissionid & " ) "

	RunSQLStatement sSql

	'If there is a permissionlevel then set it
	if request("edit_oda" & item) <> "" then
		sSql = "UPDATE egov_users_to_features "
		sSql = sSql & " SET permissionlevelid = " & request("edit_oda" & item) 
		sSql = sSql & " WHERE userid = " & request("iUserId")
		sSql = sSql & " AND featureid = " & item
		sSql = sSql & " AND permissionid = " & iPermissionid

		RunSQLStatement sSql
	End If 
Next 

response.redirect "edit_user_security.asp?iUserId=" & request("iUserId") & "&success=SU"


'------------------------------------------------------------------------------
Function GetPermissionId( ByVal sPermission )
	Dim sSql, oRs

	sSql = "SELECT permissionid FROM egov_feature_permissions WHERE permission = '" & sPermission & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
  		GetPermissionId = oRs("permissionid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
Sub DeleteUserFeature( ByVal iUserId, ByVal iFeatureId )
	Dim sSql

	'Remove current security permissions
	sSql = "DELETE FROM egov_users_to_features "
	sSql = sSql & " WHERE userid = " & iUserId
	sSql = sSql & " AND featureid = " & iFeatureId

	RunSQLStatement sSql

End Sub 

%>