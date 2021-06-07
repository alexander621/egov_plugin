<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: featureupdate.asp
' AUTHOR: Steve Loar
' CREATED: 9/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is where features are updated and created
'
' MODIFICATION HISTORY
' 1.0  09/12/08 Steve Loar - Initial Version
' 1.1  04/17/09 David Boyer - Added "CommunityLinkOn" options
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oCmd, sFeature, sFeatureName, sFeatureDescription, sHasPublicView, sPublicUrl, sPublicImageUrl
Dim sHasAdminView, sAdminUrl, sParentFeatureId, sFeatureType, sHasPermissions, sHasPermissionLevels
Dim sRootAdminRequired, sIsDefault, sPublicDisplayOrder, sAdminDisplayOrder, sSecurityDisplayOrder
Dim iFeatureId, sReturnTo
Dim sHasMobileView, sMobileURL, sIsMobileNavOnly, sMobileDefaultItemCount, sMobileDefaultDisplayOrder
Dim sMobileDefaultListCount

iFeatureId = request("iFeatureId")

If LCase(request("haspublicview")) = "on" Then
	sHasPublicView = True 
	If clng(iFeatureId) = clng(0) Then 
		' this is new so get the next value
		sPublicDisplayOrder = GetNextPublicDisplayOrder()
	Else
		' this is not new, so get the current value or a new value
		sPublicDisplayOrder = GetPublicDisplayOrder( iFeatureId )
	End If 
	'response.write sPublicDisplayOrder 
	'response.End 
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
		' this is new so get the next value
		sAdminDisplayOrder = GetNextAdminDisplayOrder() ' These are for top level items
		sSecurityDisplayOrder = 0 
	Else
		sAdminDisplayOrder = 0
		' These are the features below the top most ones. This order repeats for each grouping by top level feature
		' this is new so get the next value
		sSecurityDisplayOrder = GetNextSecurityDisplayOrder( request("parentfeatureid") ) ' This is for sub items
	End If 
Else
	' Need to handle those features moved to and from a top level

	If clng(request("parentfeatureid")) = clng(0) Then
		sSecurityDisplayOrder = 0 
		' this is not new, so get the current value or a new value
		sAdminDisplayOrder = GetAdminDisplayOrder( iFeatureId )
	Else
		sAdminDisplayOrder = 0
		' this is not new, so get the current value or a new value
		sSecurityDisplayOrder = GetSecurityDisplayOrder( iFeatureId, request("parentfeatureid") )
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

If request("communitylinkon") = "on" Then 
	sCommunityLinkOn = True
Else 
	sCommunityLinkOn = False
End If 

If request("CL_numListItems") <> "" Then 
	sCL_numListItems = request("CL_numListItems")
Else 
	sCL_numListItems = 0
End If 

If request("CL_portaltype") <> "" Then 
	sCL_portaltype = trim(request("CL_portaltype"))
Else 
	sCL_portaltype = "NULL"
End If 

If request("hasmobileview") = "on" Then
	sHasMobileView = "1"
Else
	sHasMobileView = "0"
End If 

If request("mobileurl") <> "" Then 
	sMobileURL = request("mobileurl")
Else
	sMobileURL = "NULL"
End If 

If request("ismobilenavonly") = "on" Then
	sIsMobileNavOnly = "1"
Else
	sIsMobileNavOnly = "0"
End If 

If request("mobiledefaultitemcount") <> "" Then 
	sMobileDefaultItemCount = request("mobiledefaultitemcount")
Else
	sMobileDefaultItemCount = "NULL"
End If

If request("mobiledefaultlistcount") <> "" Then 
	sMobileDefaultListCount = request("mobiledefaultlistcount")
Else
	sMobileDefaultListCount = "NULL"
End If 

If request("mobiledefaultdisplayorder") <> "" Then 
	sMobileDefaultDisplayOrder = request("mobiledefaultdisplayorder")
Else
	sMobileDefaultDisplayOrder = "NULL"
End If 



Set oCmd = Server.CreateObject("ADODB.Command")

With oCmd
.ActiveConnection = Application("DSN")
.CommandType = 4

If clng(iFeatureId) = clng(0) Then 
	.CommandText = "NewOrganizationFeature"
	' Parameter list in order
	'	 @feature varchar(50),
	'	 @featurename varchar(255),
	'	 @featuredescription text = NULL,
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
	'    @hasmobileview bit
	'    @mobileurl varchar(1025) = NULL
	'    @ismobilenavonly bit
	'    @mobiledefaultitemcount int = NULL
	'    @mobiledefaultlistcount int = NULL
	'    @mobiledefaultdisplayorder int = NULL
		
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

		If CLng(sPublicDisplayOrder) > CLng(0) Then 
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
		.Parameters.Append oCmd.CreateParameter("@communitylinkon", 11, 1, 1, sCommunityLinkOn)
		.Parameters.Append oCmd.CreateParameter("@CL_numListItems", 3, 1, 4, sCL_numListItems)
		.Parameters.Append oCmd.CreateParameter("@CL_portaltype", 129, 1, 50, sCL_portaltype)

		' Mobile Properties
		.Parameters.Append oCmd.CreateParameter("@hasmobileview", 11, 1, 1, sHasMobileView)
		If sMobileURL <> "NULL" Then
			.Parameters.Append oCmd.CreateParameter("@mobileurl", 201, 1, Len(sMobileURL), sMobileURL)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobileurl", 201, 1, 1, NULL)
		End If 
		.Parameters.Append oCmd.CreateParameter("@ismobilenavonly", 11, 1, 1, sIsMobileNavOnly)
		If sMobileDefaultItemCount <> "NULL" Then 
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultitemcount", 3, 1, 4, sMobileDefaultItemCount)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultitemcount", 3, 1, 4, NULL)
		End If 
		If sMobileDefaultListCount <> "NULL" Then 
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultlistcount", 3, 1, 4, sMobileDefaultListCount)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultlistcount", 3, 1, 4, NULL)
		End If 
		If sMobileDefaultDisplayOrder <> "NULL" Then 
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultdisplayorder", 3, 1, 4, sMobileDefaultDisplayOrder)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultdisplayorder", 3, 1, 4, NULL)
		End If 


		sReturnTo = "managefeatures.asp?success=SA"

	Else
		.CommandText = "UpdateOrganizationFeature"
		' Parameter list in order
'		@FeatureId int,
'		@featurename varchar(255),
'		@featurenotes text = NULL,
'		@featuredescription text = NULL,
'		@haspublicview bit,
'	    @publicdisplayorder int = NULL,
'	    @publicdisplayorder int = NULL,
'		@publicurl varchar(512) = NULL,
'		@publicimageurl varchar(255) = NULL,
'		@hasadminview bit,
'	    @admindisplayorder int = NULL,
'		@adminurl varchar(255) = NULL,
'		@parentfeatureid int,
'		@featuretype char(1),
'		@haspermissions bit,
'		@haspermissionlevels bit,
'	    @securitydisplayorder int = NULL,
'		@rootadminrequired bit,
'		@isdefault bit
'       @hasmobileview bit
'       @mobileurl varchar(1025) = NULL
'       @ismobilenavonly bit
'       @mobiledefaultitemcount int = NULL
'       @mobiledefaultlistcount int = NULL
'       @mobiledefaultdisplayorder int  = NULL


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

		If CLng(sPublicDisplayOrder) > CLng(0) Then 
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

		If CLng(sAdminDisplayOrder) > CLng(0) Then 
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

		If CLng(sSecurityDisplayOrder) <> CLng(0) Then 
			.Parameters.Append oCmd.CreateParameter("@securitydisplayorder", 3, 1, 4, sSecurityDisplayOrder)
		Else
			.Parameters.Append oCmd.CreateParameter("@securitydisplayorder", 3, 1, 4, NULL)
		End If 

		.Parameters.Append oCmd.CreateParameter("@rootadminrequired", 11, 1, 1, sRootAdminRequired)
		.Parameters.Append oCmd.CreateParameter("@isdefault", 11, 1, 1, sIsDefault)
		.Parameters.Append oCmd.CreateParameter("@communitylinkon", 11, 1, 1, sCommunityLinkOn)
		.Parameters.Append oCmd.CreateParameter("@CL_numListItems", 3, 1, 4, sCL_numListItems)

		If request("CL_portaltype") <> "" Then 
   			.Parameters.Append oCmd.CreateParameter("@CL_portaltype", 129, 1, 50, sCL_portaltype)
		Else 
  			.Parameters.Append oCmd.CreateParameter("@CL_portaltype", 129, 1, 50, NULL)
		End If 

		' Mobile Properties
		.Parameters.Append oCmd.CreateParameter("@hasmobileview", 11, 1, 1, sHasMobileView)
		If sMobileURL <> "NULL" Then
			.Parameters.Append oCmd.CreateParameter("@mobileurl", 201, 1, Len(sMobileURL), sMobileURL)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobileurl", 201, 1, 1, NULL)
		End If 
		.Parameters.Append oCmd.CreateParameter("@ismobilenavonly", 11, 1, 1, sIsMobileNavOnly)
		If sMobileDefaultItemCount <> "NULL" Then 
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultitemcount", 3, 1, 4, sMobileDefaultItemCount)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultitemcount", 3, 1, 4, NULL)
		End If 
		If sMobileDefaultListCount <> "NULL" Then 
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultlistcount", 3, 1, 4, sMobileDefaultListCount)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultlistcount", 3, 1, 4, NULL)
		End If 
		If sMobileDefaultDisplayOrder <> "NULL" Then 
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultdisplayorder", 3, 1, 4, sMobileDefaultDisplayOrder)
		Else
			.Parameters.Append oCmd.CreateParameter("@mobiledefaultdisplayorder", 3, 1, 4, NULL)
		End If 

		sReturnTo = "featureedit.asp?featureid=" & iFeatureId & "&success=SU"
	End If 

	.Execute
End With

Set oCmd = Nothing
	
'Return to the manage features page
response.redirect sReturnTo




'------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

	DBsafe = Replace( strDB, "'", "''" )

End Function


'------------------------------------------------------------------------------
' integer GetNextPublicDisplayOrder( )
'------------------------------------------------------------------------------
Function GetNextPublicDisplayOrder()
	Dim sSql, oRs, iMax

	sSql = "SELECT MAX(publicdisplayorder) AS maxpublicdisplayorder FROM egov_organization_features"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		iMax = CLng(oRs("maxpublicdisplayorder")) + 1
	Else 
		iMax = 1
	End If 

	oRs.close
	Set oRs = Nothing 

	GetNextPublicDisplayOrder = iMax

End Function 


'------------------------------------------------------------------------------
' integer GetNextAdminDisplayOrder( )
'------------------------------------------------------------------------------
Function GetNextAdminDisplayOrder()
	Dim sSql, oRs, iMax

	sSql = "Select max(admindisplayorder) as maxadmindisplayorder from egov_organization_features"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		iMax = CLng(oRs("maxadmindisplayorder")) + 1
	Else 
		iMax = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetNextAdminDisplayOrder = iMax

End Function


'------------------------------------------------------------------------------
' integer GetNextSecurityDisplayOrder( iParentFeatureId )
'------------------------------------------------------------------------------
Function GetNextSecurityDisplayOrder( ByVal iParentFeatureId )
	Dim sSql, oRs, iMax

	sSql = "SELECT MAX(securitydisplayorder) AS maxsecuritydisplayorder "
	sSql = sSql & "FROM egov_organization_features WHERE parentfeatureid = " & iParentFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If Not IsNull(oRs("maxsecuritydisplayorder")) Then 
			iMax = CLng(oRs("maxsecuritydisplayorder")) + 1
		Else
			iMax = 1
		End If 
	Else 
		iMax = 1
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetNextSecurityDisplayOrder = iMax

End Function


'------------------------------------------------------------------------------
' integer GetPublicDisplayOrder( iFeatureId )
'------------------------------------------------------------------------------
Function GetPublicDisplayOrder( ByVal iFeatureId )
	Dim sSql, oRs
	' The feature has a public view, so it either has an order or needs one

	sSql = "SELECT ISNULL(publicdisplayorder,0) AS publicdisplayorder "
	sSql = sSql & "FROM egov_organization_features WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("publicdisplayorder")) > CLng(0) Then 
			GetPublicDisplayOrder = oRs("publicdisplayorder")
		Else
			GetPublicDisplayOrder = GetNextPublicDisplayOrder()
		End If 
	Else 
		GetPublicDisplayOrder = GetNextPublicDisplayOrder()
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetAdminDisplayOrder( iFeatureId )
'------------------------------------------------------------------------------
Function GetAdminDisplayOrder( ByVal iFeatureId )
	Dim sSql, oRs

	' it has one, or we need to get one
	sSql = "SELECT ISNULL(admindisplayorder,0) AS admindisplayorder "
	sSql = sSql & "FROM egov_organization_features WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("admindisplayorder")) > CLng(0) Then 
			GetAdminDisplayOrder = oRs("admindisplayorder")
		Else
			GetAdminDisplayOrder = GetNextAdminDisplayOrder()
		End If 
	Else 
		GetAdminDisplayOrder = GetNextAdminDisplayOrder()
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetSecurityDisplayOrder( iFeatureId, iParentfeatureid )
'------------------------------------------------------------------------------
Function GetSecurityDisplayOrder( ByVal iFeatureId, ByVal iParentfeatureid )
	Dim sSql, oRs

	' it either has one, or we need to get one
	sSql = "SELECT ISNULL(securitydisplayorder,0) AS securitydisplayorder "
	sSql = sSql & "FROM egov_organization_features WHERE featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("securitydisplayorder")) > CLng(0) Then 
			GetSecurityDisplayOrder = oRs("securitydisplayorder")
		Else
			GetSecurityDisplayOrder = GetNextSecurityDisplayOrder( iParentfeatureid )
		End If 
	Else 
		GetSecurityDisplayOrder = GetNextSecurityDisplayOrder( iParentfeatureid )
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function




%>
