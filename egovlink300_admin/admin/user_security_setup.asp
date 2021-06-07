<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: user_security_setup.asp
' AUTHOR: Steve Loar
' CREATED: 10/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module gives permissions to users for new city setups.
'				Run this after giving the organization its features and creating its admin users.
'
' MODIFICATION HISTORY
' 1.0   10/19/06	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

' Set this for the new city
iNewOrgId = 69

bSkipAll = False 

%>

<html>
<head>
	<title>E-Gov Users Setup Script</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="Javascript">
<!--
	// Put any JavaScript here

//-->
</script>

</head>

<body>
 
<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<p><%=Now()%></p>

<% If Not bSkipAll Then %>
	<p><h4>Permissions for All Users</h4></p>
	
<%
	Dim sSql, oUsers, iFeatureOrgId, sOldOrgName, oAdmins

	sOldOrgName = "p"

	sSql = "select U.userid, U.orgid, O.orgname, U.firstname, U.lastname from users U, organizations O "
	sSql = sSql & " where U.orgid = O.orgid and U.orgid = " & iNewOrgId & " order by U.orgid, U.lastname, U.firstname"

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSQL, Application("DSN"), 0, 1

	Do While Not oUsers.EOF 
		iFeatureOrgId = clng(oUsers("orgid"))
		If sOldOrgName <> oUsers("orgname") Then 
			If sOldOrgName <> "p" Then
				response.write vbcrlf & "</ul>"
			End If 
			sOldOrgName = oUsers("orgname")
			response.write vbcrlf & "<h4>" & oUsers("orgname") & " (" & oUsers("orgid") & ")</h4>"
			response.write vbcrlf & "<ul>"
		End If 
		response.write vbcrlf & vbtab & "<li>" & oUsers("firstname") & " " & oUsers("lastname") & " (" & oUsers("userid") & ")</li>"
		response.flush
		ProcessFeatures iFeatureOrgId, oUsers("userid")
		oUsers.MoveNext
	Loop 
	response.write "</ul>"

	oUsers.close
	Set oUsers = Nothing
%>
	<p><hr /></p>
	<p><%=Now()%></p>

<% End If %>
	
	<p><h4>Extra Permissions for Admin Users</h4></p>

<%
	sOldOrgName = "p"

	sSql = "select distinct U.orgid, O.orgname, UG.userid, U.firstname, U.lastname from groupsroles GR, groups G, usersgroups UG, users U, organizations O "
	sSql = sSql & " where roleid = 53 and GR.groupid = G.groupid and G.groupid = UG.groupid and UG.userid = U.userid and U.orgid = O.orgid and U.orgid = " & iNewOrgId & " order by U.orgid, O.orgname, UG.userid"

	Set oAdmins = Server.CreateObject("ADODB.Recordset")
	oAdmins.Open sSQL, Application("DSN"), 0, 1

	Do While Not oAdmins.EOF 
		iFeatureOrgId = clng(oAdmins("orgid"))
		If sOldOrgName <> oAdmins("orgname") Then 
			If sOldOrgName <> "p" Then
				response.write vbcrlf & "</ul>"
			End If 
			sOldOrgName = oAdmins("orgname")
			response.write vbcrlf & "<h4>" & oAdmins("orgname") & " (" & oAdmins("orgid") & ")</h4>"
			response.write vbcrlf & "<ul>"
		End If 
		response.write vbcrlf & vbtab & "<li>" & oAdmins("firstname") & " " & oAdmins("lastname") & " (" & oAdmins("userid") & ")</li>"
		response.flush
		ProcessAdminFeatures iFeatureOrgId, oAdmins("userid")
		oAdmins.MoveNext
	Loop 
	response.write vbcrlf & "</ul>"

	oAdmins.close
	Set oAdmins = Nothing
%>
	<p><hr /></p>
	<p><%=Now()%></p>


	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ProcessFeatures( iOrgId, iUserId )
'--------------------------------------------------------------------------------------------------
Sub ProcessFeatures( iOrgId, iUserId )
	Dim sSql, oFeatures, iPermissionLevel

	sSql = "select OTF.orgid, FO.featureid, FO.parentfeatureid, FO.feature, FO.featurename, FO.haspermissions, FO.haspermissionlevels "
	sSql = sSql & " from egov_organization_features FO, egov_organizations_to_features OTF "
	sSql = sSql & " where FO.featureid not in (8,11) and FO.parentfeatureid not in (8,11) "
	sSql = sSql & " and FO.featureid = OTF.featureid and FO.haspermissions = 1 and orgid = " & iOrgId 
	sSql = sSql & " order by FO.featureid"


	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 0, 1

	If Not oFeatures.EOF Then 
		response.write vbcrlf & "<ul>"
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
				response.write vbcrlf & vbtab & "<li>Has " & oFeatures("featurename") & " (" & oFeatures("featureid") & ")" & "</li>"
			End If 
			oFeatures.MoveNext
		Loop 
		response.write vbcrlf & "</ul>"
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
		response.write vbcrlf & "<ul>"
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
				response.write vbcrlf & vbtab & "<li>Has " & oFeatures("featurename") & " (" & oFeatures("featureid") & ")" & "</li>"
			End If 
			oFeatures.MoveNext
		Loop 
		response.write vbcrlf & "</ul>"
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

	response.write vbcrlf & vbtab & "<li>Added " & sFeatureName & " (" & iFeatureId & ")" & "</li>"
	
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
