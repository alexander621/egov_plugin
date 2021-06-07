<%
iUserID = request("iUserId")
iOrgID = 106
adOpenStatic = 3
adLockReadOnly = 1

bIsRootAdmin = true

'Check for a screen message
lcl_onload = ""
lcl_success = request("success")

If lcl_success <> "" Then 
	lcl_msg = setupScreenMsg(lcl_success)
	lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
End If
 
If request("s") = "u" Then
	lcl_onload = "displayScreenMsg('User Permissions Were Successfully Copied.');"
End If

ShowAdminUserPicks iorgid, iUserID, bIsRootAdmin

displayFeatureList iorgid, iUserID, bIsRootAdmin


'------------------------------------------------------------------------------
Sub ShowAdminUserPicks( ByVal iOrgID, ByVal iUserID, ByVal bIsRootAdmin )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users "
	sSql = sSql & "WHERE orgid = " & iOrgID & " AND isdeleted = 0"

	If Not bIsRootAdmin Then 
		sSql = sSql & " AND (isrootadmin IS NULL OR isrootadmin = 0) "
	End If 

	sSql = sSql & " ORDER BY lastname, firstname"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 

		Do While Not oRs.EOF
			If clng(iUserID) = clng(oRs("userid")) Then 
				lcl_selected_user = " selected=""selected"""
				response.write oRs("lastname") & ", " & oRs("firstname") & "<br />" & vbcrlf
			Else 
				lcl_selected_user = ""
			End If 


			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

'------------------------------------------------------------------------------
Sub displayFeatureList( ByVal iOrgID, ByVal iUserID, ByVal bIsRootAdmin )
	Dim sChecked, sSql, oRs

	sChecked = ""

	'Get the features that the organization has
	sSql = "SELECT F.featureid, P.permission, F.featurename, F.haspermissionlevels "
	sSql = sSql & " FROM egov_organization_features F, egov_feature_permissions P, "
	sSql = sSql & " egov_features_to_permissions FP, egov_organizations_to_features FO "
	sSql = sSql & " WHERE F.parentfeatureid = 0 "
	sSql = sSql & " AND F.haspermissions = 1 "
	sSql = sSql & " AND FP.permissionid = P.permissionid "
	sSql = sSql & " AND FP.featureid = F.featureID "
	sSql = sSql & " AND FO.featureid = F.featureid "
	sSql = sSql & " AND FO.orgid = " & iOrgID

	If Not bIsRootAdmin Then 
		sSql = sSql & " AND rootadminrequired = 0 "
	End If 

	sSql = sSql & " ORDER BY F.admindisplayorder, F.featurename"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 

		Do While Not oRs.EOF
			response.write  oRs("FeatureName") & ":"

			'View Display Check
			If lCase(oRs("permission")) = "permission" Then 
				sChecked = GetPermissionSetValue(UserHasPermissionToFeature(oRs("featureid"),iUserID,"permission"))

				if sChecked <> "" then response.write " YES"
				response.write "<br />" & vbcrlf

				sChecked = ""
			Else 
			End If 

			GetChildrenFeatureList iOrgID, oRs("Featureid"), iUserID, bIsRootAdmin, oRs("FeatureName")

			oRs.MoveNext
		Loop 


	End If 

	oRs.Close
	set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Sub GetChildrenFeatureList( ByVal iOrgID, ByVal iParentID, ByVal iUserID, ByVal bIsRootAdmin, ByVal FeatureName )
	Dim sClass, sSql

	'Get the child features that the organization has
	sSql = "SELECT F.featureid, P.permission, F.featurename, F.haspermissionlevels "
	sSql = sSql & "FROM egov_organization_features F, egov_feature_permissions P, "
	sSql = sSql & "egov_features_to_permissions FP, egov_organizations_to_features FO "
	sSql = sSql & "WHERE F.haspermissions = 1 "
	sSql = sSql & "AND parentfeatureid = " & iParentID
	sSql = sSql & " AND FP.permissionid = P.permissionid "
	sSql = sSql & "AND FP.featureid = F.featureID "
	sSql = sSql & "AND FO.featureid = F.featureid "
	sSql = sSql & "AND FO.orgid = " & iOrgID
	'response.write sSQL

	If Not bIsRootAdmin Then 
		sSql = sSql & " AND rootadminrequired = 0 "
	End If 

	sSql = sSql & " ORDER BY securitydisplayorder, F.featurename"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write FeatureName & " - " & oRs("FeatureName") & ":"

			'View Display Check
			If lCase(oRs("permission")) = "permission" Then 
				sChecked = GetPermissionSetValue(UserHasPermissionToFeature(oRs("featureid"), iUserID, "permission"))
				if sChecked <> "" then response.write " YES"


				sChecked = ""
			Else 
			End If   


			'ODA level CHECK
			If oRs("haspermissionlevels") Then 
				DisplayPermissionLevels GetUserPermissionLevel(oRs("featureid"), iuserid, "permission"), "edit_oda" & oRs("featureid")
			Else 
			End If 
			response.write "<br />" & vbcrlf

			oRs.MoveNext
		Loop 
	End If 

End Sub 


'------------------------------------------------------------------------------
Sub DisplayPermissionLevels( ByVal iSelected, ByVal sName )
	Dim oRs, sSql

	sSql = "SELECT * FROM egov_feature_permission_levels ORDER BY permissionlevelid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 

		Do While Not oRs.EOF
			If Clng(iSelected) = oRs("permissionlevelid") Then 
				response.write " - " & oRs("permissionlevel")
			End If 
			oRs.MoveNext
		Loop 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
Function UserHasPermissionToFeature( ByVal iFeatureID, ByVal iUserID, ByVal sPermission )
	Dim sSql, oRs, iPermissionID

	iPermissionID = GetPermissionId( sPermission )

	UserHasPermissionToFeature = False

	sSql = "SELECT permissionid "
	sSql = sSql & " FROM egov_users_to_features "
	sSql = sSql & " WHERE userid = " & iUserID
	sSql = sSql & " AND featureid = " & iFeatureID
	sSql = sSql & " AND permissionid = " & iPermissionID 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		UserHasPermissionToFeature = True
	End If 

End Function 


'------------------------------------------------------------------------------
Function GetPermissionId( ByVal sPermission )
	Dim sSql, oRs

	sSql = "SELECT permissionid "
	sSql = sSql & " FROM egov_feature_permissions "
	sSql = sSql & " WHERE permission = '" & sPermission & "' "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		GetPermissionId = oRs("permissionid")
	End If 

	oRs.Close
	set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
Function GetPermissionSetValue( ByVal blnValue )
	
	sReturnValue = ""

	If blnValue Then 
		sReturnValue = " checked=""checked"" "
	End If 

	GetPermissionSetValue = sReturnValue

End Function


'------------------------------------------------------------------------------
Function GetUserPermissionLevel( ByVal iFeatureID, ByVal iUserID, ByVal sPermission )
	Dim sSql, oRs

	GetUserPermissionLevel = 0

	sSql = "SELECT isnull(F.permissionlevelid,0) as permissionlevelid "
	sSql = sSql & " FROM egov_users_to_features F, egov_feature_permissions P "
	sSql = sSql & " WHERE F.userid = " & iUserID
	sSql = sSql & " AND F.featureid = " & iFeatureID 
	sSql = sSql & " AND F.permissionid = P.permissionid "
	sSql = sSql & " AND P.permission = '" & sPermission & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oRs.EOF Then 
		GetUserPermissionLevel = oRs("permissionlevelid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
Sub displayButtons()

	response.write "<p>" & vbcrlf
	response.write "<input class=""button"" type=""button"" value=""Copy Permissions"" onClick=""javascript: CopyUser(" & iUserID & ");"" />" & vbcrlf
	response.write "<input class=""button"" type=""submit"" value=""Save Changes"" />" & vbcrlf
	response.write "</p>" & vbcrlf

End Sub


'------------------------------------------------------------------------------
Function setupScreenMsg( ByVal iSuccess )
	Dim lcl_return

	lcl_return = ""

	If iSuccess <> "" then
		iSuccess = UCase(iSuccess)

		If iSuccess = "SU" Then 
			lcl_return = "Successfully Updated..."
		ElseIf iSuccess = "SA" Then 
			lcl_return = "Successfully Created..."
		ElseIf iSuccess = "SR" Then 
			lcl_return = "Successfully Reordered..."
		ElseIf iSuccess = "SD" Then 
			lcl_return = "Successfully Deleted..."
		End If 
	End If 

	setupScreenMsg = lcl_return

End Function


%>
