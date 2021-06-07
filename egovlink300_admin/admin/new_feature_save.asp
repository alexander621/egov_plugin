<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: new_feature_save.asp
' AUTHOR: Steve Loar
' CREATED: 12/13/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is where new features are saved
'
' MODIFICATION HISTORY
' 1.0   12/13/06   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oCmd, sFeature, sFeatureName, sFeatureDescription, sHasPublicView, sPublicUrl, sPublicImageUrl
Dim sHasAdminView, sAdminUrl, sParentFeatureId, sFeatureType, sHasPermissions, sHasPermissionLevels
Dim sRootAdminRequired, sIsDefault, sPublicDisplayOrder, sAdminDisplayOrder, sSecurityDisplayOrder
Dim iFeatureId

iFeatureId = request("iFeatureId")

If request("haspublicview") = "on" Then
	sHasPublicView = True 
	If clng(iFeatureId) = clng(0) Then 
		sPublicDisplayOrder = GetNextPublicDisplayOrder()
	End If 
Else 
	sHasPublicView = False 
	If clng(iFeatureId) = clng(0) Then 
		sPublicDisplayOrder = 0
	End If 
End If 

If request("hasadminview") = "on" Then
	sHasAdminView = True 
Else 
	sHasAdminView = False 
End If 

If clng(iFeatureId) = clng(0) Then 
	If clng(request("parentfeatureid")) = clng(0) Then 
		sAdminDisplayOrder = GetNextAdminDisplayOrder() ' These are for top level items
		sSecurityDisplayOrder = 0 
	Else
		sAdminDisplayOrder = 0
		sSecurityDisplayOrder = GetNextSecurityDisplayOrder( request("parentfeatureid") ) ' This is for sub items
	End If 
End If 

If request("haspermissions") = "on" Then 
	sHasPermissions = True 
Else
	sHasPermissions = False 
End If 

If request("haspermissionlevels") = "on" Then 
	sHasPermissionLevels = True 
Else
	sHasPermissionLevels = False 
End If 

If request("rootadminrequired") = "on" Then 
	sRootAdminRequired = True 
Else
	sRootAdminRequired = False 
End If 

If request("isdefault") = "on" Then 
	sIsDefault = True 
Else
	sIsDefault = False 
End If 


Set oCmd = Server.CreateObject("ADODB.Command")

With oCmd
	.ActiveConnection = Application("DSN")
	.CommandType = 4

	If clng(iFeatureId) = clng(0) Then 
		.CommandText = "NewOrganizationFeature"
		'	@feature varchar(50),
		 '   @featurename varchar(255),
		  '  @featuredescription text = NULL,
		'    @haspublicview bit,
		'    @publicdisplayorder int = NULL,
		'    @publicurl varchar(512) = NULL,
		'    @publicimageurl varchar(255) = NULL,
		'    @hasadminview bit,
		'    @admindisplayorder int = NULL,
		'    @adminurl varchar(255) = NULL,
		'    @parentfeatureid int,
		'    @featuretype char(1),
		'    @haspermissions bit,
		'    @haspermissionlevels bit,
		'    @securitydisplayorder int = NULL,
		'    @rootadminrequired bit,
		'    @isdefault bit
		
		.Parameters.Append oCmd.CreateParameter("@feature", 200, 1, 50, LCase(request("feature")))
		.Parameters.Append oCmd.CreateParameter("@featurename", 200, 1, 255, request("featurename"))

		If request("featurenotes") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@featurenotes", 201, 1, Len(request("featurenotes")), request("featurenotes"))
		Else
			.Parameters.Append oCmd.CreateParameter("@featurenotes", 201, 1, 1, NULL)
		End If 
		
		If request("featuredescription") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@featuredescription", 201, 1, Len(request("featuredescription")), request("featuredescription"))
		Else
			.Parameters.Append oCmd.CreateParameter("@featuredescription", 201, 1, 1, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@haspublicview", 11, 1, 1, sHasPublicView)

		If sPublicDisplayOrder <> 0 Then 
			.Parameters.Append oCmd.CreateParameter("@publicdisplayorder", 3, 1, 4, sPublicDisplayOrder)
		Else
			.Parameters.Append oCmd.CreateParameter("@publicdisplayorder", 3, 1, 4, NULL)
		End If 

		If request("publicurl") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@publicurl", 200, 1, 512, request("publicurl"))
		Else
			.Parameters.Append oCmd.CreateParameter("@publicurl", 200, 1, 512, NULL)
		End If 

		If request("publicimageurl") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@publicimageurl", 200, 1, 255, request("publicimageurl"))
		Else
			.Parameters.Append oCmd.CreateParameter("@publicimageurl", 200, 1, 255, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@hasadminview", 11, 1, 1, sHasAdminView)

		If sAdminDisplayOrder <> 0 Then 
			.Parameters.Append oCmd.CreateParameter("@admindisplayorder", 3, 1, 4, sAdminDisplayOrder)
		Else
			.Parameters.Append oCmd.CreateParameter("@admindisplayorder", 3, 1, 4, NULL)
		End If 

		If request("adminurl") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@adminurl", 200, 1, 255, request("adminurl"))
		Else
			.Parameters.Append oCmd.CreateParameter("@adminurl", 200, 1, 255, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@parentfeatureid", 3, 1, 4, request("parentfeatureid"))

		.Parameters.Append oCmd.CreateParameter("@featuretype", 129, 1, 1, request("featuretype"))

		.Parameters.Append oCmd.CreateParameter("@haspermissions", 11, 1, 1, sHasPermissions)

		.Parameters.Append oCmd.CreateParameter("@haspermissionlevels", 11, 1, 1, sHasPermissionLevels)

		If sSecurityDisplayOrder <> 0 Then 
			.Parameters.Append oCmd.CreateParameter("@securitydisplayorder", 3, 1, 4, sSecurityDisplayOrder)
		Else
			.Parameters.Append oCmd.CreateParameter("@securitydisplayorder", 3, 1, 4, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@rootadminrequired", 11, 1, 1, sRootAdminRequired)
		.Parameters.Append oCmd.CreateParameter("@isdefault", 11, 1, 1, sIsDefault)
	Else
		.CommandText = "UpdateOrganizationFeature"
'		@FeatureId int,
'		@featurename varchar(255),
'		@featurenotes text = NULL,
'		@featuredescription text = NULL,
'		@haspublicview bit,
'		@publicurl varchar(512) = NULL,
'		@publicimageurl varchar(255) = NULL,
'		@hasadminview bit,
'		@adminurl varchar(255) = NULL,
'		@parentfeatureid int,
'		@featuretype char(1),
'		@haspermissions bit,
'		@haspermissionlevels bit,
'		@rootadminrequired bit,
'		@isdefault bit

		.Parameters.Append oCmd.CreateParameter("@featureid", 3, 1, 4, iFeatureId)
		.Parameters.Append oCmd.CreateParameter("@featurename", 200, 1, 255, request("featurename"))

		If request("featurenotes") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@featurenotes", 201, 1, Len(request("featurenotes")), request("featurenotes"))
		Else
			.Parameters.Append oCmd.CreateParameter("@featurenotes", 201, 1, 1, NULL)
		End If 
		
		If request("featuredescription") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@featuredescription", 201, 1, Len(request("featuredescription")), request("featuredescription"))
		Else
			.Parameters.Append oCmd.CreateParameter("@featuredescription", 201, 1, 1, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@haspublicview", 11, 1, 1, sHasPublicView)

		If request("publicurl") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@publicurl", 200, 1, 512, request("publicurl"))
		Else
			.Parameters.Append oCmd.CreateParameter("@publicurl", 200, 1, 512, NULL)
		End If 

		If request("publicimageurl") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@publicimageurl", 200, 1, 255, request("publicimageurl"))
		Else
			.Parameters.Append oCmd.CreateParameter("@publicimageurl", 200, 1, 255, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@hasadminview", 11, 1, 1, sHasAdminView)

		If request("adminurl") <> "" Then 
			.Parameters.Append oCmd.CreateParameter("@adminurl", 200, 1, 255, request("adminurl"))
		Else
			.Parameters.Append oCmd.CreateParameter("@adminurl", 200, 1, 255, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@parentfeatureid", 3, 1, 4, request("parentfeatureid"))
		.Parameters.Append oCmd.CreateParameter("@featuretype", 129, 1, 1, request("featuretype"))
		.Parameters.Append oCmd.CreateParameter("@haspermissions", 11, 1, 1, sHasPermissions)
		.Parameters.Append oCmd.CreateParameter("@haspermissionlevels", 11, 1, 1, sHasPermissionLevels)
		.Parameters.Append oCmd.CreateParameter("@rootadminrequired", 11, 1, 1, sRootAdminRequired)
		.Parameters.Append oCmd.CreateParameter("@isdefault", 11, 1, 1, sIsDefault)
	End If 

	.Execute
End With

Set oCmd = Nothing
	
' Return to the edit page
response.redirect "manage_features.asp?orgid=" & request("orgid")
'response.redirect "new_feature.asp?orgid=" & request("orgid")


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
	DBsafe = Replace( strDB, "'", "''" )
End Function


'--------------------------------------------------------------------------------------------------
' Function GetNextPublicDisplayOrder()
'--------------------------------------------------------------------------------------------------
Function GetNextPublicDisplayOrder()
	Dim sSql, oDisplay, iMax

	sSql = "Select max(publicdisplayorder) as maxpublicdisplayorder from egov_organization_features"

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSQL, Application("DSN"), 0, 1

	If Not oDisplay.EOF Then 
		iMax = CLng(oDisplay("maxpublicdisplayorder")) + 1
	Else 
		iMax = 1
	End If 

	oDisplay.close
	Set oDisplay = Nothing 

	GetNextPublicDisplayOrder = iMax

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetNextAdminDisplayOrder()
'--------------------------------------------------------------------------------------------------
Function GetNextAdminDisplayOrder()
	Dim sSql, oDisplay, iMax

	sSql = "Select max(admindisplayorder) as maxadmindisplayorder from egov_organization_features"

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSQL, Application("DSN"), 0, 1

	If Not oDisplay.EOF Then 
		iMax = CLng(oDisplay("maxadmindisplayorder")) + 1
	Else 
		iMax = 1
	End If 

	oDisplay.close
	Set oDisplay = Nothing 

	GetNextAdminDisplayOrder = iMax
End Function


'--------------------------------------------------------------------------------------------------
' Function GetNextSecurityDisplayOrder()
'--------------------------------------------------------------------------------------------------
Function GetNextSecurityDisplayOrder( iParentFeatureId )
Dim sSql, oDisplay, iMax

	sSql = "Select max(securitydisplayorder) as maxsecuritydisplayorder from egov_organization_features Where parentfeatureid = " & iParentFeatureId

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSQL, Application("DSN"), 0, 1

	If Not oDisplay.EOF Then 
		If Not IsNull(oDisplay("maxsecuritydisplayorder")) Then 
			iMax = CLng(oDisplay("maxsecuritydisplayorder")) + 1
		Else
			iMax = 1
		End If 
	Else 
		iMax = 1
	End If 

	oDisplay.close
	Set oDisplay = Nothing 

	GetNextSecurityDisplayOrder = iMax
End Function 


%>
