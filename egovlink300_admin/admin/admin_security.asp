<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: admin_security.asp
' AUTHOR: Steve Loar
' CREATED: 02/26/2007
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates security for the admins of a city
'
' MODIFICATION HISTORY
' 1.0   02/26/2007	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oUsers, iFeatureOrgId, sOldOrgName, oAdmins, bSkipAll, iOrgId

iOrgId = request("orgid")

' Permissions for Admin Users
sOldOrgName = "p"

sSql = "select distinct U.orgid, O.orgname, UG.userid, U.firstname, U.lastname from groupsroles GR, groups G, usersgroups UG, users U, organizations O "
sSql = sSql & " where roleid = 53 and GR.groupid = G.groupid and G.groupid = UG.groupid and UG.userid = U.userid and U.orgid = O.orgid and U.orgid = " & iOrgId & " order by U.orgid, O.orgname, UG.userid"

Set oAdmins = Server.CreateObject("ADODB.Recordset")
oAdmins.Open sSQL, Application("DSN"), 0, 1

Do While Not oAdmins.EOF 
	iFeatureOrgId = clng(oAdmins("orgid"))
	If sOldOrgName <> oAdmins("orgname") Then 
		sOldOrgName = oAdmins("orgname")
	End If 
	ProcessFeatures iFeatureOrgId, oAdmins("userid")
'	ProcessAdminFeatures iFeatureOrgId, oAdmins("userid")
	oAdmins.MoveNext
Loop 

oAdmins.close
Set oAdmins = Nothing

' Back to the edit page
response.redirect "featureselection.asp?orgid=" & iOrgId



'--------------------------------------------------------------------------------------------------
' Sub ProcessFeatures( iOrgId, iUserId )
'--------------------------------------------------------------------------------------------------
Sub ProcessFeatures( iOrgId, iUserId )
	Dim sSql, oFeatures, iPermissionLevel

	sSql = "select OTF.orgid, FO.featureid, FO.parentfeatureid, FO.feature, FO.featurename, FO.haspermissions, FO.haspermissionlevels "
	sSql = sSql & " from egov_organization_features FO, egov_organizations_to_features OTF Where "
'	sSql = sSql & " FO.featureid not in (8,11) and FO.parentfeatureid not in (8,11) and "
	sSql = sSql & " FO.featureid = OTF.featureid and FO.haspermissions = 1 and orgid = " & iOrgId 
	sSql = sSql & " order by FO.featureid"


	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 0, 1

	If Not oFeatures.EOF Then 
'		response.write vbcrlf & "<ul>"
		Do While Not oFeatures.EOF 
			If Not HasUserFeature( oFeatures("feature"), iUserId ) Then
				If oFeatures("haspermissionlevels") Then
					iPermissionLevel = 1
				Else
					iPermissionLevel = "NULL"
				End If 
				GiveUserFeature iUserId, oFeatures("featureid"), iPermissionLevel, oFeatures("featurename")
				'response.write vbcrlf & "<li>Needs " & oFeatures("featurename") & " (" & oFeatures("featureid") & ")" & "</li>"
			Else
'				response.write vbcrlf & vbtab & "<li>Has " & oFeatures("featurename") & " (" & oFeatures("featureid") & ")" & "</li>"
			End If 
			oFeatures.MoveNext
		Loop 
'		response.write vbcrlf & "</ul>"
	End If 

	oFeatures.close
	Set oFeatures = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ProcessAdminFeatures( iOrgId, iUserId )
'--------------------------------------------------------------------------------------------------
Sub ProcessAdminFeatures( iOrgId, iUserId )
	Dim sSql, oFeatures, iPermissionLevel

	sSql = "select OTF.orgid, FO.parentfeatureid, FO.featureid, FO.feature, FO.featurename, FO.haspermissions, FO.haspermissionlevels "
	sSql = sSql & " from egov_organization_features FO, egov_organizations_to_features OTF "
	sSql = sSql & " where (FO.featureid in (8,11) or FO.parentfeatureid in (8,11)) "
	sSql = sSql & " and FO.featureid = OTF.featureid and FO.haspermissions = 1 and orgid = " & iOrgId 
	sSql = sSql & " order by OTF.orgid, FO.parentfeatureid, FO.featureid"


	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 0, 1

	If Not oFeatures.EOF Then 
'		response.write vbcrlf & "<ul>"
		Do While Not oFeatures.EOF 
			If Not HasUserFeature( oFeatures("feature"), iUserId ) Then
				If oFeatures("haspermissionlevels") Then
					iPermissionLevel = 1
				Else
					iPermissionLevel = "NULL"
				End If 
				GiveUserFeature iUserId, oFeatures("featureid"), iPermissionLevel, oFeatures("featurename")
				'response.write vbcrlf & "<li>Needs " & oFeatures("featurename") & " (" & oFeatures("featureid") & ")" & "</li>"
'			Else
'				response.write vbcrlf & vbtab & "<li>Has " & oFeatures("featurename") & " (" & oFeatures("featureid") & ")" & "</li>"
			End If 
			oFeatures.MoveNext
		Loop 
'		response.write vbcrlf & "</ul>"
	End If 

	oFeatures.close
	Set oFeatures = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GiveUserFeature( iUserId, iFeatureId, iPermissionLevel, sFeatureName )
'--------------------------------------------------------------------------------------------------
Sub GiveUserFeature( iUserId, iFeatureId, iPermissionLevel, sFeatureName )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "Insert Into egov_users_to_features ( featureid, userid, permissionid, permissionlevelid ) Values ( " & iFeatureId & ", " & iUserId & ", 1, " & iPermissionLevel & " ) "
		.Execute
	End With
	Set oCmd = Nothing

'	response.write vbcrlf & vbtab & "<li>Added " & sFeatureName & " (" & iFeatureId & ")" & "</li>"
	
End Sub 


'--------------------------------------------------------------------------------------------------
' FUNCTION HasUserFeature( sFeature, iOrgId )
'--------------------------------------------------------------------------------------------------
Function HasUserFeature( sFeature, iUserId )
	Dim sSql, oFeatureAccess, blnReturnValue

	' SET DEFAULT
	HasUserFeature = False

	' LOOKUP passed FEATURE FOR the current ORGANIZATION 
	sSql = "SELECT count(FO.featureid) as feature_count FROM egov_users_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid and userid = " & iUserId & " AND F.feature = '" & sFeature & "' "

	Set oFeatureAccess = Server.CreateObject("ADODB.Recordset")
	oFeatureAccess.Open  sSQL, Application("DSN"), 3, 1
	
	If clng(oFeatureAccess("feature_count")) > clng(0) Then
		' the ORGANIZATION HAS the FEATURE
		HasUserFeature = True
	End If
	
	oFeatureAccess.close 
	Set oFeatureAccess = Nothing

End Function
%>
