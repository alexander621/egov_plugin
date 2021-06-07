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
	
	Private Sub Class_Initialize()

	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetEgovURL()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetEgovURL()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(OrgEgovWebsiteURL,'') AS OrgEgovWebsiteURL FROM organizations WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetEgovURL = oRs("OrgEgovWebsiteURL")
		Else
			GetEgovURL = ""
		End If
			
		oRs.Close
		Set oRs = Nothing

	End Function 


	'------------------------------------------------------------------------------
	Public Function GetOrgBanner()
		Dim sSql, oRs, lcl_return

		lcl_return = ""

		sSql = "SELECT OrgName, OrgTopGraphicLeftURL, OrgHeaderSize "
		sSql = sSql & " FROM organizations "
		sSql = sSql & " WHERE orgid = " & iOrgId

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			If oRs("orgTopGraphicLeftURL") <> "" Then 
				lcl_return = "<img src=""" & replace(oRs("OrgTopGraphicLeftURL"),"http://www.egovlink.com","") & """ border=""0"" height=""" & oRs("OrgHeaderSize") & """ alt=""" & oRs("OrgName") & " E-Gov Services"" />" & vbcrlf
			End If 
		End If 

		oRs.Close
		Set oRs = Nothing 

		GetOrgBanner = lcl_return

	End Function 


'-- Original ------------------------------------------------------------------
'	Public Function GetOrgBanner()
'		Dim sSql, oRs

'		sSql = "Select OrgName, OrgTopGraphicLeftURL, OrgHeaderSize FROM organizations WHERE orgid = " & iOrgId

'		Set oRs = Server.CreateObject("ADODB.Recordset")
'		oRs.Open sSql, Application("DSN"), 3, 1

'		If Not oRs.EOF Then 
'			GetOrgBanner = "<img src=""" & oRs("OrgTopGraphicLeftURL") & """ border=""0"" height=""" & oRs("OrgHeaderSize") & """ alt=""" & oRs("OrgName") & " e-Government Services"" />"
'		End If
			
'		oRs.close
'		Set oRs = Nothing
'	End Function 

'------------------------------------------------------------------------------
	public function GetOrgBannerBckgrnd()
		 dim sSQL, oRs, lcl_return

   lcl_return = ""

		 sSQL = "SELECT OrgTopGraphicRightURL, "
   sSQL = sSQL & " OrgHeaderSize "
   sSQL = sSQL & " FROM organizations "
   sSQL = sSQL & " WHERE orgid = " & iOrgId

		 set oRs = Server.CreateObject("ADODB.Recordset")
		 oRs.Open sSQL, Application("DSN"), 0, 1

		 if not oRs.eof then
			   lcl_return = lcl_return & "background:url(" & oRs("OrgTopGraphicRightURL") & ") repeat;" & vbcrlf
      lcl_return = lcl_return & "height:" & oRs("OrgHeaderSize") & "px;" & vbcrlf
		 end if

		 oRs.close
		 set oRs = nothing

   GetOrgBannerBckgrnd = lcl_return

 end function	

	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgName()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgName()
		Dim sSql, oRs
		
		If iOrgid <> "" Then 
			sSql = "SELECT orgname FROM organizations WHERE orgid = " & iOrgId

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 0, 1

			If Not oRs.EOF Then 
				GetOrgName = oRs("orgname")
			End If
				
			oRs.close
			Set oRs = Nothing
		Else
			GetOrgName = ""
		End If 

	End Function 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgURL()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgURL()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(OrgPublicWebsiteURL,'') AS OrgPublicWebsiteURL FROM organizations WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetOrgURL = oRs("OrgPublicWebsiteURL")
		End If
			
		oRs.Close
		Set oRs = Nothing

	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetState()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetState()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(orgstate,'') AS orgstate FROM organizations WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetState = oRs("orgstate")
		End If
			
		oRs.Close
		Set oRs = Nothing

	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetDefaultState()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetDefaultState()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(defaultstate,'') AS defaultstate FROM organizations WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetDefaultState = oRs("defaultstate")
		End If
			
		oRs.close
		Set oRs = Nothing

	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetDefaultCity()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetDefaultCity()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(defaultcity,'') AS defaultcity FROM organizations WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetDefaultCity = oRs("defaultcity")
		End If
			
		oRs.close
		Set oRs = Nothing

	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Function GetDefaultZip()
	'------------------------------------------------------------------------------------------------------------
	Public Function GetDefaultZip()
		Dim sSql, oRs

		sSql = "SELECT ISNULL(defaultzip,'') AS defaultzip FROM organizations WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			GetDefaultZip = oRs("defaultzip")
		End If
			
		oRs.Close
		Set oRs = Nothing

	End Function 


	'--------------------------------------------------------------------------------------------------
	' Public Function OrgHasFeature( sFeature )
	'--------------------------------------------------------------------------------------------------
	Public Function OrgHasFeature( ByVal sFeature )
		Dim sSql, oRs

		OrgHasFeature = False

		' Lookup the passed feature for the organization 
		sSql = "SELECT COUNT(FO.featureid) AS feature_count "
		sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & " WHERE FO.featureid = F.featureid AND F.feature = '" & sFeature & "' AND orgid = " & iOrgId

		session("sSql") = sSql	' For debugging
		session("sPathURL") = GetPathURL()

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open  sSql, Application("DSN"), 0, 1
		session("sSql") = ""
		session("sPathURL") = ""
		
		If clng(oRs("feature_count")) > 0 Then
			' the Organization has the feature
			OrgHasFeature = True
		End If
		
		oRs.Close 
		Set oRs = Nothing

	End Function


	'--------------------------------------------------------------------------------------------------
	' Private Function GetPathURL()
	'--------------------------------------------------------------------------------------------------
	Private Function GetPathURL()

		prot = "http" 
		https = lcase(request.ServerVariables("HTTPS")) 
		if https <> "off" then prot = "https" 
		domainname = Request.ServerVariables("SERVER_NAME") 
		filename = Request.ServerVariables("SCRIPT_NAME") 
		querystring = Request.ServerVariables("QUERY_STRING") 
		GetPathURL = prot & "://" & domainname & filename & "?" & querystring 

	End Function 


	'--------------------------------------------------------------------------------------------------
	' Public Function OrgHasMembership( sFeature )
	'--------------------------------------------------------------------------------------------------
	Public Function OrgHasMembership( ByVal sMembership )
		Dim sSql, oRs

		OrgHasMembership = False

		' Lookup the passed feature for the organization 
		sSql = "SELECT COUNT(membershipid) AS membership_count FROM egov_memberships "
		sSql = sSql & " WHERE orgid = " & iOrgId & " AND membership = '" & sMembership & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open  sSql, Application("DSN"), 3, 1
		
		If clng(oRs("membership_count")) > 0 Then
			' the Organization has the membership
			OrgHasMembership = True
		End If
		
		oRs.Close 
		Set oRs = Nothing

	End Function


	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicDropDownMenu()
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicDropDownMenu()
		Dim sSql, oRs, sNav

		sSql = "SELECT O.OrgEgovWebsiteURL, ISNULL(FO.publicurl,F.publicURL) AS publicURL, "
		sSql = sSql & "ISNULL(FO.featurename,F.featurename) AS featurename "
		sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & "WHERE FO.publiccanview = 1 AND F.haspublicview = 1 and O.orgid = FO.orgid "
		sSql = sSql & " AND FO.featureid = F.featureid AND O.orgid = " & iOrgId
		sSql = sSql & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		Do While Not oRs.EOF 
			If UCase(Left(oRs("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oRs("publicURL")
			Else
				' start from our page
				sNav = oRs("OrgEgovWebsiteURL") & "/" & oRs("publicURL")
     				if request.servervariables("HTTPS") = "on" then
					sNav = replace(sNav,"http:","https:")
				end if
			End If 
			response.write vbcrlf & vbtab & "<li><a href=""" & sNav & """>" & oRs("featurename") & "</a></li>"
			oRs.MoveNext
		Loop 
			
		oRs.Close
		Set oRs = Nothing

	End Sub
	

	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicFooterNav(iCount)
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicFooterNav( ByRef iCount )
		Dim sSql, oRs, sNav, iTotalCount

		sSql = "SELECT O.OrgEgovWebsiteURL, ISNULL(FO.publicurl,F.publicURL) AS publicURL, "
		sSql = sSql & "ISNULL(FO.featurename,F.featurename) AS featurename "
		sSql = sSql & " FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & " WHERE FO.publiccanview = 1 AND F.haspublicview = 1 AND O.orgid = FO.orgid "
		sSql = sSql & " AND FO.featureid = F.featureid AND O.orgid = " & iOrgId
		sSql = sSql & " ORDER BY FO.publicdisplayorder, F.publicdisplayorder"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		iTotalCount = oRs.recordcount

		Do While Not oRs.EOF
			If (iCount Mod 6) = 0 Then 
				response.write "<br />"
			Else 
				response.write " | "
			End If 

			If UCase(Left(oRs("publicURL"),4)) = "HTTP" then
				'They have their own page to start from
				sNav = oRs("publicURL")
			Else 
				'Start from our page
				sNav = oRs("OrgEgovWebsiteURL") & "/" & oRs("publicURL")
			End If 

			response.write "<a href=""" & sNav & """ class=""afooter"" target=""_top"">" & oRs("featurename") & "</a>" & vbcrlf

			iCount = iCount + 1
			oRs.MoveNext
		Loop   

		oRs.Close
		set oRs = Nothing 

	End Sub 

	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicDefaultFooterNav(iCount)
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicDefaultFooterNav( ByRef iCount )
		Dim sSql, oRs, sNav, iTotalCount

		sSql = "SELECT O.OrgEgovWebsiteURL, ISNULL(FO.publicurl,F.publicURL) AS publicURL, "
		sSql = sSql & "ISNULL(FO.featurename,F.featurename) AS featurename "
		sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSql = sSql & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		iTotalCount = oRs.recordcount

		Do While Not oRs.EOF 
			If (iCount Mod 6) = 0 Then 
				response.write "<br />"
			Else 
				response.write " | "
			End If 
			If UCase(Left(oRs("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oRs("publicURL")
			Else
				' start from our page
				sNav = oRs("OrgEgovWebsiteURL") & "/" & oRs("publicURL")
			End If 
			response.write vbcrlf & "<a href=""" & replace(sNav,"http://www.egovlink.com","") & """ class=""adefaultfooter"">" & oRs("featurename") & "</a>"
			
			iCount = iCount + 1
			oRs.MoveNext
		Loop 
			
		oRs.close
		Set oRs = Nothing

	End Sub 

	
	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicLeftNav()
	'------------------------------------------------------------------------------------------------------------
	public sub ShowPublicLeftNav()
		Dim sSql, oRs, sNav

		sSql = "SELECT O.OrgEgovWebsiteURL, ISNULL(FO.publicurl,F.publicURL) AS publicURL, "
		sSql = sSql & " ISNULL(FO.featurename,F.featurename) AS featurename "
		sSql = sSql & " FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & " WHERE FO.publiccanview = 1 AND F.haspublicview = 1 AND O.orgid = FO.orgid "
		sSql = sSql & " AND FO.featureid = F.featureid AND O.orgid = " & iOrgId
		sSql = sSql & " ORDER BY FO.publicdisplayorder, F.publicdisplayorder "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		Do While Not oRs.EOF
			If UCase(Left(oRs("publicURL"),4)) = "HTTP" Then 
				'They have their own page to start from
				sNav = oRs("publicURL")
			Else 
				'start from our page
				sNav = oRs("OrgEgovWebsiteURL") & "/" & oRs("publicURL")
     				if request.servervariables("HTTPS") = "on" then
					sNav = replace(sNav,"http:","https:")
				end if
			End If 

			response.write "<p><a href=""" & sNav & """>" & oRs("featurename") & "</a></p>" & vbcrlf
			'response.write vbcrlf & vbtab & "<p><a href=""" & sNav & """><img src=""images/btn.gif"" height=""9"" width=""6"" border=""0"" alt="""" /> &nbsp; " & oRs("featurename") & "</a></p>"
			oRs.MoveNext
		Loop 

		oRs.Close
		Set oRs = Nothing 

		'Add the login link for those that have this
		If OrgHasFeature("registration") Then 
			response.write "<p><a href=""user_login.asp"">Login</a></p>" & vbcrlf
		End If 

	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Sub ShowPublicMainNav()
	'------------------------------------------------------------------------------------------------------------
	Public Sub ShowPublicMainNav()
		Dim sSql, oRs, sNav, sAlign, bHasImage

		sAlign = "imgleft"
		sSql = "Select O.OrgEgovWebsiteURL, isnull(FO.publicurl,F.publicURL) as publicURL, isnull(FO.publicimageurl,F.publicimageurl) as publicimageurl, "
		sSql = sSql & "isnull(FO.featurename,F.featurename) as featurename, isnull(FO.featuredescription,F.featuredescription) as featuredescription "
		sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & "WHERE FO.publiccanview = 1 and F.haspublicview = 1 and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & iOrgId
		sSql = sSql & " Order By FO.publicdisplayorder,F.publicdisplayorder"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		Do While Not oRs.EOF 
			If sAlign = "imgleft" Then 
				sAlign = "imgright"
			Else
				sAlign = "imgleft"
			End If 
			 
			If UCase(Left(oRs("publicURL"),4)) = "HTTP" Then 
				' They have their own page to start from
				sNav = oRs("publicURL")
			Else
				' start from our page
				sNav = oRs("OrgEgovWebsiteURL") & "/" & oRs("publicURL")

				sNav = replace(sNav,"http://www.egovlink.com","")
			End If 
			Response.Write vbcrlf & vbtab & "<div class=""featuregroup"" onClick=""location.href='" & sNav & "';"">"
			response.write vbcrlf & vbtab & vbtab & "<h2><a href=""" & sNav & """>" & oRs("featurename") & "</a></h2>"

			response.write vbcrlf & vbtab & vbtab & "<div class=""features"" onClick=""location.href='" & sNav & "';"">"
			
			If oRs("publicimageurl") <> "" Then
				response.write "<img class=""" & sAlign & """ src=""" & replace(oRs("publicimageurl"),"http://www.egovlink.com","") & """ alt="""" />"
				bHasImage = True 
			Else
				bHasImage = False 
			End If
			response.write "<p"
			If bHasImage = True Then
				response.write " class=""hasimage"" "
			End If 
			response.write ">"
			response.write oRs("featuredescription") 
			response.write "</p></div>"
			response.write vbcrlf & vbtab & "</div>"
			oRs.MoveNext
		Loop 
			
		oRs.Close
		Set oRs = Nothing

	End Sub 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgDisplayName( sDisplay )
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgDisplayName( ByVal sDisplay )
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
	Private Function GetOrgSpecificDisplayname( ByVal sDisplay )
		Dim sSql, oRs

		sSql = "select O.displayname from egov_organizations_to_displays O, egov_organization_displays D "
		sSql = sSql & " where O.displayid = D.displayid and D.display = '" & sDisplay & "' and O.orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			GetOrgSpecificDisplayname = oRs("displayname")
		Else
			GetOrgSpecificDisplayname = ""
		End If 

		oRs.Close
		Set oRs = Nothing 

	End Function 

	
	'------------------------------------------------------------------------------------------------------------
	' Private Function GetDefaultDisplayName( sDisplay )
	'------------------------------------------------------------------------------------------------------------
	Private Function GetDefaultDisplayName( ByVal sDisplay )
		Dim sSql, oRs

		sSql = "select displayname from egov_organization_displays "
		sSql = sSql & " where display = '" & sDisplay & "'" 

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then
			GetDefaultDisplayName = oRs("displayname")
		Else
			GetDefaultDisplayName = ""
		End If 

		oRs.Close
		Set oRs = Nothing 

	End Function 


	'------------------------------------------------------------------------------------------------------------
	' Public Function GetOrgFeatureName( sFeature )
	'------------------------------------------------------------------------------------------------------------
	Public Function GetOrgFeatureName( ByVal sFeature )
		Dim sSql, oRs

		sSql = "SELECT ISNULL(FO.featurename,F.featurename) AS featurename "
		sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & " where FO.featureid = F.featureid and FO.orgid = " & iOrgId & " and feature = '" & sFeature & "'" 

		'response.write sSql
		'response.End 

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then
			GetOrgFeatureName = oRs("featurename")
		Else
			GetOrgFeatureName = ""
		End If 

		oRs.Close
		Set oRs = Nothing 

	End Function 


	'--------------------------------------------------------------------------------------------------
	' Public Function OrgHasNeighborhoods( )
	'--------------------------------------------------------------------------------------------------
	Public Function OrgHasNeighborhoods( )
		Dim sSql, oRs

		sSql = "SELECT COUNT(neighborhoodid) AS hits FROM egov_neighborhoods WHERE orgid = " & iOrgId 
		
		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If clng(oRs("hits")) > 0 Then
			OrgHasNeighborhoods = True 
		Else
			OrgHasNeighborhoods = False 
		End if
		
		oRs.Close
		Set oRs = Nothing
		
	End Function 


	'--------------------------------------------------------------------------------------------------
	' FUNCTION OrgHasDisplay( iorgid, sDisplay )
	'--------------------------------------------------------------------------------------------------
	Public Function OrgHasDisplay( ByVal sDisplay )
		Dim sSql, oRs, blnReturnValue

		' SET DEFAULT
		blnReturnValue = False

		' LOOKUP passed display FOR the current ORGANIZATION 
		sSql = "SELECT COUNT(OD.displayid) AS display_count FROM egov_organizations_to_displays OD, egov_organization_displays D "
		sSql = sSql & " WHERE OD.displayid = D.displayid AND orgid = " & iOrgId & " AND D.display = '" & sDisplay & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open  sSql, Application("DSN"), 0, 1
		
		If clng(oRs("display_count")) > 0 Then
			' the ORGANIZATION HAS the Display
			blnReturnValue = True
		End If
		
		oRs.close 
		Set oRs = Nothing

		' set the RETURN  value
		OrgHasDisplay = blnReturnValue

	End Function


	'--------------------------------------------------------------------------------------------------
	' FUNCTION GetOrgDisplay( iorgid, sDisplay )
	'--------------------------------------------------------------------------------------------------
	Public Function GetOrgDisplay( ByVal sDisplay )
		Dim sSql, oRs

		' SET DEFAULT
		GetOrgDisplay = ""

		' LOOKUP passed Display FOR the passed Organization 
		sSql = "select ISNULL(OD.displaydescription, D.displaydescription) AS displaydescription "
		sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
		sSql = sSql & " WHERE OD.displayid = D.displayid AND orgid = " & iOrgId & " AND D.display = '" & sDisplay & "' "

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open  sSql, Application("DSN"), 3, 1
		
		If Not oRs.EOF Then
			' the ORGANIZATION HAS the Display
			GetOrgDisplay = oRs("displaydescription")
		End If
		
		oRs.Close 
		Set oRs = Nothing

	End Function


	'------------------------------------------------------------------------------
	public function checkMenuOptionEnabled( ByVal iField )
		Dim lcl_return, sSql, oRs

		lcl_return = True

		sSql = "SELECT public_menuopt_" & iField & "home_enabled AS MenuOpt_Enabled "
		sSql = sSql & " FROM organizations "
		sSql = sSql & " WHERE orgid = " & iOrgId

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			lcl_return = oRs("MenuOpt_Enabled")
		End If 

		oRs.close
		Set oRs = Nothing 

		checkMenuOptionEnabled = lcl_return

	End function

	'------------------------------------------------------------------------------
	Public Function getMenuOptionLabel( ByVal iField )
		Dim lcl_return, sSql, oRs

		lcl_return = ""

		sSql = "SELECT public_menuopt_" & iField & "home_label AS MenuOpt_Label "
		sSql = sSql & " FROM organizations "
		sSql = sSql & " WHERE orgid = " & iOrgId

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 3, 1

		If Not oRs.EOF Then 
			lcl_return = oRs("MenuOpt_Label")
		End If 

		oRs.Close
		Set oRs = Nothing 

		getMenuOptionLabel = lcl_return

	End Function 

	'------------------------------------------------------------------------------
	Sub buildWelcomeMessage( ByVal p_orgid, ByVal iOrgHasDisplay_ActionPageTitle, ByVal iOrgName, ByVal iOrgState, ByVal iOrgFeatureName )
		Dim lcl_welcome_message

		If iOrgHasDisplay_ActionPageTitle Then 
			lcl_welcome_message = GetOrgDisplay("action page title")
		Else 
			lcl_welcome_message = ""

			If iOrgName <> "" Then 
				lcl_welcome_message = iOrgName
			End If 

			If iOrgState <> "" Then 
				If lcl_welcome_message <> "" Then 
					lcl_welcome_message = lcl_welcome_message & ", " & iOrgState
				Else 
					lcl_welcome_message = iOrgState
				End If 
			End If 

			If iOrgFeatureName <> "" Then 
				If lcl_welcome_message <> "" Then 
					lcl_welcome_message = lcl_welcome_message & ", " & iOrgFeatureName
				Else 
					lcl_welcome_message = iOrgFeatureName
				End If 
			end if

			If lcl_welcome_message <> "" Then 
				lcl_welcome_message = " to the " & lcl_welcome_message
			End If 

			lcl_welcome_message = "Welcome" & lcl_welcome_message

		End If 

		response.write "<font class=""pagetitle"">" & lcl_welcome_message & "</font>" & vbcrlf

	End Sub 


End Class
%>
