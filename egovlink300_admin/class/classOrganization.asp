<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: classOrganization.asp
' AUTHOR: Steve Loar
' CREATED: 07/11/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the Organization class
'
' MODIFICATION HISTORY
' 1.0   07/11/2006   Steve Loar - Initial code 
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Class classOrganization
	Dim iOrgId

	
	'------------------------------------------------------------------------------------------------------------
	' Private Sub Class_Initialize()
	'------------------------------------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		If IsNumeric(session("orgid")) And Not isEmpty(session("orgid")) Then 
			iOrgId = CLng(session("orgid"))
		Else
			iOrgId = 0
		End If 
	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Sub SetOrgId( iNewOrgId ) 
	'------------------------------------------------------------------------------------------------------------
	Public Sub SetOrgId( iNewOrgId ) 
		If IsNumeric(iNewOrgId) Then 
			iOrgId = CLng(iNewOrgId)
		Else
			iOrgId = 0
		End If 
	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetEgovURL()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetEgovURL()
		Dim sSQL, oURL

		sSQL = "Select isnull(OrgEgovWebsiteURL,'') as OrgEgovWebsiteURL FROM organizations WHERE orgid = " & iOrgId

		Set oURL = Server.CreateObject("ADODB.Recordset")
		oURL.Open sSQL, Application("DSN"), 3, 1

		If Not oURL.EOF Then 
			GetEgovURL = oURL("OrgEgovWebsiteURL")
		End If
			
		oURL.close
		Set oURL = Nothing
	End Function 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgBanner()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgBanner()
		Dim sSQL, oBanner

		sSQL = "Select OrgName, OrgTopGraphicLeftURL, OrgHeaderSize FROM organizations WHERE orgid = " & iOrgId

		Set oBanner = Server.CreateObject("ADODB.Recordset")
		oBanner.Open sSQL, Application("DSN"), 3, 1

		If Not oBanner.EOF Then 
			GetOrgBanner = "<img src=""" & oBanner("OrgTopGraphicLeftURL") & """ border=""0"" height=""" & oBanner("OrgHeaderSize") & """ alt=""" & oBanner("OrgName") & " e-Government Services"" />"
		End If
			
		oBanner.close
		Set oBanner = Nothing
	End Function 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgBannerBckgrnd()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgBannerBckgrnd()
		Dim sSQL, oBanner

		sSQL = "Select OrgTopGraphicRightURL FROM organizations WHERE orgid = " & iOrgId

		Set oBanner = Server.CreateObject("ADODB.Recordset")
		oBanner.Open sSQL, Application("DSN"), 3, 1

		If Not oBanner.EOF Then 
			GetOrgBannerBckgrnd = "background:url(" & oBanner("OrgTopGraphicRightURL") & ") repeat;"
		End If
			
		oBanner.close
		Set oBanner = Nothing
	End Function 
	
	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgName()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgName()
		Dim sSQL, oName

		sSQL = "Select orgname FROM organizations WHERE orgid = " & iorgid
'		response.write sSql
'		response.end

		Set oName = Server.CreateObject("ADODB.Recordset")
		oName.Open sSQL, Application("DSN"), 3, 1

		If Not oName.EOF Then 
			GetOrgName = oName("orgname")
		End If
			
		oName.close
		Set oName = Nothing
	End Function 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgURL()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgURL()
		Dim sSQL, oURL

		sSQL = "Select isnull(OrgPublicWebsiteURL,'') as OrgPublicWebsiteURL FROM organizations WHERE orgid = " & iOrgId

		Set oURL = Server.CreateObject("ADODB.Recordset")
		oURL.Open sSQL, Application("DSN"), 3, 1

		If Not oURL.EOF Then 
			GetOrgURL = oURL("OrgPublicWebsiteURL")
		End If
			
		oURL.close
		Set oURL = Nothing
	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetState()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetState()
		Dim sSQL, oState

		sSQL = "Select isnull(orgstate,'') as orgstate FROM organizations WHERE orgid = " & iOrgId

		Set oState = Server.CreateObject("ADODB.Recordset")
		oState.Open sSQL, Application("DSN"), 3, 1

		If Not oState.EOF Then 
			GetState = oState("orgstate")
		End If
			
		oState.close
		Set oState = Nothing
	End Function 


	'--------------------------------------------------------------------------------------------------
	' Public Function OrgHasFeature( sFeature )
	'--------------------------------------------------------------------------------------------------
	Public Function OrgHasFeature( sFeature )
		Dim sSql, oRs

		OrgHasFeature = False

		' Lookup the passed feature for the organization 
		sSql = "SELECT count(FO.featureid) as feature_count FROM egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & " WHERE FO.featureid = F.featureid and orgid = " & iOrgId & " AND F.feature = '" & sFeature & "' "
		session("OrgHasFeatureSQL") = sSql

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open  sSQL, Application("DSN"), 3, 1
		
		If clng(oRs("feature_count")) > 0 Then
			' the Organization has the feature
			OrgHasFeature = True
		End If
		
		oRs.Close 
		Set oRs = Nothing

		session("OrgHasFeatureSQL") = ""

	End Function


	'--------------------------------------------------------------------------------------------------
	' Public Function OrgHasMembership( sFeature )
	'--------------------------------------------------------------------------------------------------
	Public Function OrgHasMembership( sMembership )
		Dim sSql, oMembership

		OrgHasMembership = False

		' Lookup the passed feature for the organization 
		sSql = "SELECT count(membershipid) as membership_count FROM egov_memberships "
		sSql = sSql & " WHERE orgid = " & iOrgId & " AND membership = '" & sMembership & "' "

		Set oMembership = Server.CreateObject("ADODB.Recordset")
		oMembership.Open  sSQL, Application("DSN"), 3, 1
		
		If clng(oMembership("membership_count")) > 0 Then
			' the Organization has the membership
			OrgHasMembership = True
		End If
		
		oMembership.close 
		Set oMembership = Nothing

	End Function


	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicDropDownMenu()
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicDropDownMenu()
		Dim sSQL, oNav, sNav

		sSQL = "Select O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename "
		sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSQL = sSQL & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oNav = Server.CreateObject("ADODB.Recordset")
		oNav.Open sSQL, Application("DSN"), 3, 1

		Do While Not oNav.EOF 
			If UCase(Left(oNav("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oNav("publicURL")
			Else
				' start from our page
				sNav = oNav("OrgEgovWebsiteURL") & "/" & oNav("publicURL")
			End If 
			response.write vbcrlf & vbtab & "<li><a href=""" & sNav & """>" & oNav("featurename") & "</a></li>"
			oNav.MoveNext
		Loop 
			
		oNav.close
		Set oNav = Nothing
	End Sub
	

	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicFooterNav(iCount)
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicFooterNav(iCount)
		Dim sSQL, oNav, sNav, iTotalCount

		sSQL = "Select O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename "
		sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSQL = sSQL & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oNav = Server.CreateObject("ADODB.Recordset")
		oNav.Open sSQL, Application("DSN"), 3, 1

		iTotalCount = oNav.recordcount

		Do While Not oNav.EOF 
			If (iCount Mod 6) = 0 Then 
				response.write "<br />"
			Else 
				response.write " | "
			End If 
			If UCase(Left(oNav("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oNav("publicURL")
			Else
				' start from our page
				sNav = oNav("OrgEgovWebsiteURL") & "/" & oNav("publicURL")
			End If 
			response.write vbcrlf & vbtab & "<a href=""" & sNav & """ class=""afooter"">" & oNav("featurename") & "</a>"
			
			iCount = iCount + 1
			oNav.MoveNext
		Loop 
			
		oNav.close
		Set oNav = Nothing
	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicDefaultFooterNav(iCount)
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicDefaultFooterNav(iCount)
		Dim sSQL, oNav, sNav, iTotalCount

		sSQL = "Select O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename "
		sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSQL = sSQL & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oNav = Server.CreateObject("ADODB.Recordset")
		oNav.Open sSQL, Application("DSN"), 3, 1

		iTotalCount = oNav.recordcount

		Do While Not oNav.EOF 
			If (iCount Mod 6) = 0 Then 
				response.write "<br />"
			Else 
				response.write " | "
			End If 
			If UCase(Left(oNav("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oNav("publicURL")
			Else
				' start from our page
				sNav = oNav("OrgEgovWebsiteURL") & "/" & oNav("publicURL")
			End If 
			response.write vbcrlf & vbtab & "<a href=""" & sNav & """ class=""adefaultfooter"">" & oNav("featurename") & "</a>"
			
			iCount = iCount + 1
			oNav.MoveNext
		Loop 
			
		oNav.close
		Set oNav = Nothing
	End Sub 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicLeftNav()
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicLeftNav()
		Dim sSQL, oNav, sNav

		sSQL = "Select O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename "
		sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSQL = sSQL & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oNav = Server.CreateObject("ADODB.Recordset")
		oNav.Open sSQL, Application("DSN"), 3, 1

		Do While Not oNav.EOF 
			If UCase(Left(oNav("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oNav("publicURL")
			Else
				' start from our page
				sNav = oNav("OrgEgovWebsiteURL") & "/" & oNav("publicURL")
			End If 
			response.write vbcrlf & vbtab & "<p><a href=""" & sNav & """>" & oNav("featurename") & "</a></p>"
'			response.write vbcrlf & vbtab & "<p><a href=""" & sNav & """><img src=""images/btn.gif"" height=""9"" width=""6"" border=""0"" alt="""" /> &nbsp; " & oNav("featurename") & "</a></p>"
			oNav.MoveNext
		Loop 
			
		oNav.close
		Set oNav = Nothing

		' Add the login link for those that have this
		If OrgHasFeature("registration") Then
			response.write vbcrlf & vbtab & "<p><a href=""user_login.asp"">Login</a></p>"
		End If 

	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicMainNav()
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicMainNav()
		Dim sSQL, oNav, sNav, sAlign, bHasImage

		sAlign = "imgleft"
		sSQL = "Select O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, isnull(FO.publicimageurl,F.publicimageurl) as publicimageurl, "
		sSQL = sSQL & "isnull(FO.featurename,F.featurename) as featurename, isnull(FO.featuredescription,F.featuredescription) as featuredescription "
		sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSQL = sSQL & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSQL = sSQL & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oNav = Server.CreateObject("ADODB.Recordset")
		oNav.Open sSQL, Application("DSN"), 3, 1

		Do While Not oNav.EOF 
			If sAlign = "imgleft" Then 
				sAlign = "imgright"
			Else
				sAlign = "imgleft"
			End If 
			 
			If UCase(Left(oNav("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oNav("publicURL")
			Else
				' start from our page
				sNav = oNav("OrgEgovWebsiteURL") & "/" & oNav("publicURL")
			End If 
			Response.Write vbcrlf & vbtab & "<div class=""featuregroup"" onClick=""location.href='" & sNav & "';"">"
			response.write vbcrlf & vbtab & vbtab & "<h2><a href=""" & sNav & """>" & oNav("featurename") & "</a></h2>"

			response.write vbcrlf & vbtab & vbtab & "<div class=""features"" onClick=""location.href='" & sNav & "';"">"
			
			If oNav("publicimageurl") <> "" Then
				response.write "<img class=""" & sAlign & """ src=""" & oNav("publicimageurl") & """ alt="""" />"
				bHasImage = True 
			Else
				bHasImage = False 
			End If
			response.write "<p"
			If bHasImage = True Then
				response.write " class=""hasimage"" "
			End If 
			response.write ">"
			response.write oNav("featuredescription") 
			response.write "</p></div>"
			response.write vbcrlf & vbtab & "</div>"
			oNav.MoveNext
		Loop 
			
		oNav.close
		Set oNav = Nothing
	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgDisplayName( sDisplay )
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgDisplayName( sDisplay )
		Dim sDisplayname

		sDisplayname = ""

		' Get the org override 
			sDisplayname = GetOrgSpecificDisplayname( sDisplay )

		' If no override, try to get a default
		If sDisplayname = "" Then
			sDisplayname = GetDefaultDisplayName( sDisplay )
		End If 

		GetOrgDisplayName = sDisplayname

	End Function 


	'------------------------------------------------------------------------------------------------------------
	' Private Function GetOrgSpecificDisplayname( sDisplay )
	'------------------------------------------------------------------------------------------------------------
	Private Function GetOrgSpecificDisplayname( sDisplay )
		Dim sSql, oDisplay

		If CLng(iOrgid) > CLng(0) Then 
			sSql = "select O.displayname from egov_organizations_to_displays O, egov_organization_displays D "
			sSql = sSql & " where O.displayid = D.displayid and D.display = '" & sDisplay & "' and O.orgid = " & iOrgId

			Set oDisplay = Server.CreateObject("ADODB.Recordset")
			oDisplay.Open sSQL, Application("DSN"), 3, 1

			If Not oDisplay.EOF Then
				GetOrgSpecificDisplayname = oDisplay("displayname")
			Else
				GetOrgSpecificDisplayname = ""
			End If 
			oDisplay.close
			Set oDisplay = Nothing 
		Else
			GetOrgSpecificDisplayname = ""
		End If 

	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Private Function GetDefaultDisplayName( sDisplay )
	'------------------------------------------------------------------------------------------------------------
	Private Function GetDefaultDisplayName( sDisplay )
		Dim sSql, oDisplay

		sSql = "select displayname from egov_organization_displays "
		sSql = sSql & " where display = '" & sDisplay & "'" 

		Set oDisplay = Server.CreateObject("ADODB.Recordset")
		oDisplay.Open sSQL, Application("DSN"), 3, 1

		If Not oDisplay.EOF Then
			GetDefaultDisplayName = oDisplay("displayname")
		Else
			GetDefaultDisplayName = ""
		End If 

		oDisplay.close
		Set oDisplay = Nothing 
	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetDefaultState()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetDefaultState()
		Dim sSQL, oState

		sSQL = "Select isnull(defaultstate,'') as defaultstate FROM organizations WHERE orgid = " & iOrgId

		Set oState = Server.CreateObject("ADODB.Recordset")
		oState.Open sSQL, Application("DSN"), 3, 1

		If Not oState.EOF Then 
			GetDefaultState = oState("defaultstate")
		End If
			
		oState.close
		Set oState = Nothing
	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetDefaultCity()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetDefaultCity()
		Dim sSQL, oCity

		sSQL = "Select isnull(defaultcity,'') as defaultcity FROM organizations WHERE orgid = " & iOrgId

		Set oCity = Server.CreateObject("ADODB.Recordset")
		oCity.Open sSQL, Application("DSN"), 3, 1

		If Not oCity.EOF Then 
			GetDefaultCity = oCity("defaultcity")
		End If
			
		oCity.close
		Set oCity = Nothing
	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetDefaultZip()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetDefaultZip()
		Dim sSQL, oZip

		sSQL = "Select isnull(defaultzip,'') as defaultzip FROM organizations WHERE orgid = " & iOrgId

		Set oZip = Server.CreateObject("ADODB.Recordset")
		oZip.Open sSQL, Application("DSN"), 3, 1

		If Not oZip.EOF Then 
			GetDefaultZip = oZip("defaultzip")
		End If
			
		oZip.close
		Set oZip = Nothing
	End Function 


End Class 
%>
