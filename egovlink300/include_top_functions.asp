<%
'------------------------------------------------------------------------------------------------------------
'  integer SetOrganizationParameters()
'------------------------------------------------------------------------------------------------------------
Function SetOrganizationParameters()
	Dim sSql, oRs, iReturnValue, sProtocol, sCurrent, sServer

	' SET DEFAULT RETURN VALUE
	iReturnValue = CLng(0)

	' BUILD CURRENT URL
	If request.servervariables("HTTPS") = "on" Then
		sProtocol = "https://"
	Else
		sProtocol = "http://"
	End If
	
	sServer = request.servervariables("SERVER_NAME")
	
	' Translate secure payment URL to regular URL for Org lookup
	If LCase(sServer) = "secure.egovlink.com" or LCase(sServer) = "www.egovlink.com" Then
		sCurrent = "http://www.egovlink.com/" & GetVirtualDirectyName()
	Else
		sCurrent = sProtocol & sServer & "/" & GetVirtualDirectyName()
	End if 
	
	' LOOKUP CURRENT URL IN DATABASE
	'sSql = "SELECT * FROM Organizations INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID WHERE OrgEgovWebsiteURL = '" & sCurrent & "'"
	sSql = "SELECT OrgID, "
	sSql = sSql & " OrgName, "
	sSql = sSql & " OrgPublicWebsiteURL, "
	sSql = sSql & " OrgEgovWebsiteURL, "
	sSql = sSql & " OrgTopGraphicLeftURL, "
	sSql = sSql & " OrgTopGraphicRightURL,"
	sSql = sSql & " OrgWelcomeMessage, "
	sSql = sSql & " OrgActionLineDescription, "
	sSql = sSql & " OrgPaymentDescription, "
	sSql = sSql & " OrgHeaderSize, "
	sSql = sSql & " OrgTagline, "
	sSql = sSql & " OrgPaymentGateway, "
	sSql = sSql & " OrgActionOn, "
	sSql = sSql & " OrgPaymentOn, "
	sSql = sSql & " OrgDocumentOn, "
	sSql = sSql & " OrgCalendarOn, "
	sSql = sSql & " OrgFaqOn, "
	sSql = sSql & " orgVirtualSiteName, "
	sSql = sSql & " OrgActionName, "
	sSql = sSql & " OrgPaymentName, "
	sSql = sSql & " OrgCalendarName, "
	sSql = sSql & " OrgDocumentName, "
	sSql = sSql & " OrgRegistration, "
	sSql = sSql & " OrgRequestCalOn, "
	sSql = sSql & " OrgRequestCalForm, "
	sSql = sSql & " OrgPublicWebsiteTag, "
	sSql = sSql & " OrgEgovWebsiteTag, "
	sSql = sSql & " OrgCustomButtonsOn, "
	sSql = sSql & " gmtoffset, "
	sSql = sSql & " orgDisplayMenu, "
	sSql = sSql & " orgDisplayFooter, "
	sSql = sSql & " orgCustomMenu, "
	sSql = sSql & " defaultphone, "
	sSql = sSql & " defaultemail, "
	sSql = sSql & " defaultstate, "
	sSql = sSql & " defaultcity, "
	sSql = sSql & " defaultzip, "
	sSql = sSql & " separate_index_catalog, "
	sSql = sSql & " orgwaivertext, "
	sSql = sSql & " OrgGoogleAnalyticAccnt, "
	sSql = sSql & " latitude,longitude "
	sSql = sSql & " FROM organizations O, timezones T "
	sSql = sSql & " WHERE O.OrgTimeZoneID = T.TimeZoneID "
	sSql = sSql & " AND O.OrgEgovWebsiteURL = '" & sCurrent & "'"

	session("OrgParametersSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	session("OrgParametersSql") = ""
	
	If Not oRs.EOF Then
		iOrgID             = oRs("OrgID")
		sOrgName           = oRs("OrgName")
		sHomeWebsiteURL    = oRs("OrgPublicWebsiteURL")
		sEgovWebsiteURL    = oRs("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oRs("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oRs("OrgTopGraphicRightURL")
		dblLat		   = oRs("latitude")
		dblLng		   = oRs("longitude")
		intZoom = 12

		If request.servervariables("HTTPS") = "on" Then
			 'ADJUST FOR PAYMENT URL
			  sTopGraphicLeftURL = Replace(oRs("OrgTopGraphicLeftURL"),"http:","https:")
			  sTopGraphicRighURL = Replace(oRs("OrgTopGraphicRightURL"),"http:","https:")
			sEgovWebsiteURL    = Replace(oRs("OrgEgovWebsiteURL"),"http:","https:")
		End If

		sWelcomeMessage      = oRs("OrgWelcomeMessage")
		sActionDescription   = oRs("OrgActionLineDescription")
		sPaymentDescription  = oRs("OrgPaymentDescription")
		iHeaderSize          = oRs("OrgHeaderSize")
		sTagline             = oRs("OrgTagline")
		iPaymentGatewayID    = oRs("OrgPaymentGateway")
		blnOrgAction         = oRs("OrgActionOn")
		blnOrgPayment        = oRs("OrgPaymentOn")
		blnOrgDocument       = oRs("OrgDocumentOn")
		blnOrgCalendar       = oRs("OrgCalendarOn")
		blnOrgFaq            = oRs("OrgFaqOn")
		sorgVirtualSiteName  = oRs("orgVirtualSiteName")
		sOrgActionName       = oRs("OrgActionName")
		sOrgPaymentName      = oRs("OrgPaymentName")
		sOrgCalendarName     = oRs("OrgCalendarName")
		sOrgDocumentName     = oRs("OrgDocumentName")
		'sOrgFaqName          = oRs("OrgFaqName")
		sOrgRegistration     = oRs("OrgRegistration")
		blnCalRequest        = oRs("OrgRequestCalOn")
		iCalForm             = oRs("OrgRequestCalForm")
		sHomeWebsiteTag      = oRs("OrgPublicWebsiteTag")
		sEgovWebsiteTag      = oRs("OrgEgovWebsiteTag")
		bCustomButtonsOn     = oRs("OrgCustomButtonsOn")
		iTimeOffset          = oRs("gmtoffset")
		blnMenuOn            = oRs("orgDisplayMenu")
		blnFooterOn          =	oRs("orgDisplayFooter")
		blnCustomMenu        = oRs("orgCustomMenu")
		sDefaultPhone        = oRs("defaultphone")
		sDefaultEmail        = oRs("defaultemail")
		sDefaultState        = oRs("defaultstate")
		sDefaultCity         = oRs("defaultcity")
		sDefaultZip          = oRs("defaultzip")
		blnSeparateIndex     = oRs("separate_index_catalog")
		sWaiverText	         = oRs("orgwaivertext")
		sGoogleAnalyticAccnt = oRs("OrgGoogleAnalyticAccnt")
	Else
		' The Org could not be found due to a bad URL so take a shot at another URL before you crash - SJL - 1/5/2007
		
		' Close things before you leave
		oRs.Close
		Set oRs = Nothing 

		' Take a guess at what the URL should be and redirect them there.
		sCurrent = "http://www.egovlink.com/" & GetVirtualDirectyName()
		response.redirect sCurrent
	End If

	oRs.Close
	Set oRs = Nothing 

	If Not IsNull(iOrgID) Then 
		iReturnValue = iOrgID
	End If

	' RETURN VALUE
	SetOrganizationParameters = iReturnValue
	
End Function


'------------------------------------------------------------------------------------------------------------
' string GetPageName()
'------------------------------------------------------------------------------------------------------------
Function GetPageName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	For Each arr in strURL 
		sReturnValue = arr 
	Next 
	
	GetPageName = sReturnValue

End Function


'------------------------------------------------------------------------------------------------------------
' string GetVirtualDirectyName()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function


'------------------------------------------------------------------------------------------------------------
' void fnInserGoogleAnalytics( sGoogleAccount )
'------------------------------------------------------------------------------------------------------------
Function fnInserGoogleAnalytics( ByVal sGoogleAccount )
	
	' IF ACCNT IS NOT EMPTY THEN POPULATE GOOGLE ANALYTIC TRACKING CODE
	If Not IsNull(sAccnt) OR sAccnt <> "" Then
		If request.servervariables("HTTPS") <> "on" Then
   'response.write "<script src=""http://www.google-analytics.com/urchin.js"" type=""text/javascript"">"
   'response.write "</script>"
   'response.write "<script type=""text/javascript"">"
   'response.write "_uacct = """ & sGoogleAccount & """;"
   'response.write "urchinTracker();"
   'response.write "</script>"
			response.write "<script type=""text/javascript"">" & vbcrlf
			response.write "  var gaJsHost = ((""https:"" == document.location.protocol) ? ""https://ssl."" : ""http://www."");" & vbcrlf
			response.write "  document.write(unescape(""%3Cscript src='"" + gaJsHost + ""google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E""));" & vbcrlf
			response.write "</script>" & vbcrlf
			response.write "<script type=""text/javascript"">" & vbcrlf
			response.write "  var pageTracker = _gat._getTracker(""" & sGoogleAccount & """);" & vbcrlf
			response.write "  pageTracker._initData();" & vbcrlf
			response.write "  pageTracker._trackPageview();" & vbcrlf
			response.write "</script>" & vbcrlf
		End If
	End If

End Function


'--------------------------------------------------------------------------------------------------
' boolean HasFeature( iFeatureid, iorgid )
'--------------------------------------------------------------------------------------------------
Function HasFeature( ByVal iFeatureid, ByVal iorgid )
	Dim sSql, oFeatureAccess, blnReturnValue

	' SET DEFAULT
	blnReturnValue = False

	' LOOKUP FEATUREID FOR ORGANIZATION ID SUPPLIED
	sSql = "SELECT featureid FROM egov_organizations_to_features WHERE orgid = " & iorgid & " AND featureid = " & iFeatureid

	Set oFeatureAccess = Server.CreateObject("ADODB.Recordset")
	oFeatureAccess.Open  sSql, Application("DSN"), 3, 1
	
	If Not oFeatureAccess.EOF Then
		' ORGANIZATION HAS FEATURE
		blnReturnValue = True
	End If
	
	oFeatureAccess.Close 
	Set oFeatureAccess = Nothing

	' RETURN 
	HasFeature = blnReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' boolean OrgHasFeature( iorgid, sFeature )
'--------------------------------------------------------------------------------------------------
Function OrgHasFeature( ByVal iOrgId, ByVal sFeature )
	Dim sSql, oRs, blnReturnValue

	' SET DEFAULT
	blnReturnValue = False

	' LOOKUP passed FEATURE FOR the current ORGANIZATION 
	sSql = "SELECT COUNT(FO.featureid) AS feature_count FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid AND orgid = " & iOrgId & " AND F.feature = '" & sFeature & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	session("sSql") = sSql  ' this is set to try and catch missing orgid values - 11/8/06
	oRs.Open  sSql, Application("DSN"), 0, 1
	session("sSql") = ""
	
	If clng(oRs("feature_count")) > clng(0) Then
		' the ORGANIZATION HAS the FEATURE
		blnReturnValue = True
	End If
	
	oRs.Close 
	Set oRs = Nothing

	' set the RETURN  value
	OrgHasFeature = blnReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' boolean CitizenAddressIsMissing( iUserId )
'--------------------------------------------------------------------------------------------------
Function CitizenAddressIsMissing( ByVal iUserId )
	Dim sSql, oRs, blnReturnValue

	' SET DEFAULT
	blnReturnValue = False

	' Pull the citizen address 
	sSql = "SELECT ISNULL(useraddress,'') AS useraddress FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 0, 1
	
	If Trim(oRs("useraddress")) = "" Then
		blnReturnValue = True
	End If
	
	oRs.Close 
	Set oRs = Nothing

	CitizenAddressIsMissing = blnReturnValue

End Function


'--------------------------------------------------------------------------------------------------
' boolean FeatureIsTurnedOnForPublic( iorgid, sFeature )
'--------------------------------------------------------------------------------------------------
Function FeatureIsTurnedOnForPublic( ByVal iOrgId, ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT FO.publiccanview FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid and orgid = " & iOrgId & " AND F.feature = '" & sFeature & "' "
	'response.write sSql & "<br />"
	session("FeatureIsTurnedOnForPublicSql") = sSql  

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 3, 1
	session("FeatureIsTurnedOnForPublicSql") = ""

	If Not oRs.EOF Then
		If oRs("publiccanview") Then 
			FeatureIsTurnedOnForPublic = True 
		Else
			FeatureIsTurnedOnForPublic = False
		End If 
	Else
		FeatureIsTurnedOnForPublic = False 
	End If 

	oRs.Close
	Set oRs = Nothing 


End Function 


'--------------------------------------------------------------------------------------------------
' boolean OrgHasDisplay( iorgid, sDisplay )
'--------------------------------------------------------------------------------------------------
Function OrgHasDisplay( ByVal iOrgId, ByVal sDisplay )
	Dim sSql, oDisplay, blnReturnValue

	' SET DEFAULT
	blnReturnValue = False

	' LOOKUP passed display FOR the current ORGANIZATION 
	sSql = "SELECT count(OD.displayid) as display_count FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid and orgid = " & iOrgId & " AND D.display = '" & sDisplay & "' "
	'response.write sSql & "<br />"
	'response.End 

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 3, 1
	
	If clng(oDisplay("display_count")) > 0 Then
		' the ORGANIZATION HAS the Display
		blnReturnValue = True
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing

	' set the RETURN  value
	OrgHasDisplay = blnReturnValue
End Function


'------------------------------------------------------------------------------
' string GetOrgDisplay( iorgid, sDisplay )
'------------------------------------------------------------------------------
Function GetOrgDisplay( ByVal iOrgId, ByVal sDisplay )
	Dim sSql, oDisplay

	' SET DEFAULT
	GetOrgDisplay = ""

	' LOOKUP passed Display FOR the passed Organization 
	sSql = "SELECT ISNULL(OD.displaydescription, D.displaydescription) AS displaydescription "
	sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid AND orgid = " & iOrgId & " AND D.display = '" & sDisplay & "' "

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 3, 1
	
	If Not oDisplay.EOF Then
		' the ORGANIZATION HAS the Display
		GetOrgDisplay = oDisplay("displaydescription")
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing

End Function


'------------------------------------------------------------------------------
Function RegisteredUserDisplay( ByVal sPath )
	Dim sSql, oRs, sUserName

	'sPath should be "" for root scripts and "../" for those in directories
	If sOrgRegistration Then 

		'If cookie found display menu
		If request.cookies("userid") <> "" And request.cookies("userid") <> "-1" and isnumeric(request.cookies("userid")) Then 
			sSql = "SELECT userfname + ' ' + userlname AS username FROM egov_users WHERE userid = " & CLng(request.cookies("userid"))

			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open sSql, Application("DSN"), 3, 1

			If Not oRs.EOF Then 
				sUserName = UCase(oRs("username"))
			Else 
				sUserName = "NOT KNOWN"
			End If 

			oRs.CLose 
			Set oRs = Nothing 

			'Build Account Manage Menu
			response.write "<div id=""accountmenu"">" & vbcrlf

			'Welcome Message
			response.write "    <img class=""accountmenu"" src=""" & sPath & "images/accountmenu.jpg"" />" & vbcrlf
			response.write "    Welcome" 

			If Trim(sUserName) <> "" Then 
				response.write ", <strong>" & sUserName & "</strong>"
			End If 

			response.write "!<br />" & vbcrlf

			'Menu Links
			ShowLoggedinLinks sPath

			response.write "</div>" & vbcrlf
		Else 
			'No cookie do not display menu
			response.write "<div id=""datetagline""><font class=""datetagline"">Today is " & FormatDateTime(Date(), vbLongDate) & ". " & sTagline & " </font></div>" & vbcrlf
		End If 
	Else 
		response.write "<div id=""datetagline""><font class=""datetagline"">Today is " & FormatDateTime(Date(), vbLongDate) & ". " & sTagline & " </font></div>" & vbcrlf
	End If 

End Function 


'------------------------------------------------------------------------------
sub ShowLoggedinLinks( ByVal sPath )
 'Org Features
  lcl_orghasfeature_payments     = orghasfeature(iOrgID, "payments")
  lcl_orghasfeature_action_line  = orghasfeature(iOrgID, "action line")
  lcl_orghasfeature_activities   = orghasfeature(iOrgID, "activities")
  lcl_orghasfeature_facilities   = orghasfeature(iOrgID, "facilities")
  lcl_orghasfeature_memberships  = orghasfeature(iOrgID, "memberships")
  lcl_orghasfeature_gifts        = orghasfeature(iOrgID, "gifts")
  lcl_orghasfeature_bid_postings = orghasfeature(iOrgID, "bid_postings")
  lcl_orghasfeature_donotknock   = orghasfeature(iOrgID, "donotknock")
  lcl_orghasfeature_rentals	 = orghasfeature(iOrgID, "rentals")

 'PublicCanView Features
  lcl_publicCanViewFeature_payments     = publicCanViewFeature(iOrgID, "payments")
  lcl_publicCanViewFeature_action_line  = publicCanviewFeature(iOrgID, "action line")
  lcl_publicCanViewFeature_activities   = publicCanViewFeature(iOrgID, "activities")
  lcl_publicCanViewFeature_facilities   = publicCanViewFeature(iOrgID, "facilities")
  lcl_publicCanViewFeature_memberships  = publicCanViewFeature(iOrgID, "memberships")
  lcl_publicCanViewFeature_gifts        = publicCanViewFeature(iOrgID, "gifts")
  lcl_publicCanViewFeature_bid_postings = publicCanViewFeature(iOrgID, "bid_postings")
  lcl_publicCanViewFeature_donotknock   = publicCanViewFeature(iOrgID, "donotknock")
  lcl_publicCanViewFeature_rentals   = publicCanViewFeature(iOrgID, "rentals")

if iorgid = "228" then
	retvariables = ""
	for each x in Request.querystring
		if retvariables <> "" then retvariables = retvariables & "&"
  		retvariables = retvariables & x & "=" & request.querystring(x)

	next
	for each x in Request.form
		if retvariables <> "" then retvariables = retvariables & "&"
  		retvariables = retvariables & x & "=" & request.form(x)

	next

	if instr(request.servervariables("SCRIPT_NAME"), "manage_account") < 1 then
		retURL = request.servervariables("SCRIPT_NAME") 
		if retvariables <> "" then retURL = retURL & "?" & retvariables
		session("retfromtop") = retURL
		'response.write "retfrommaurl=" & request.servervariables("SCRIPT_NAME") & retvariables
	end if
end if
	'Manage Account Link
	 response.write "<a class=""accountmenu"" href=""" & sPath & "manage_account.asp"">MANAGE ACCOUNT</a> | " & vbcrlf

	'View Standard EGov Payments Link
	if lcl_orghasfeature_payments AND lcl_publicCanViewFeature_payments then
       response.write "<a class=""accountmenu"" href=""" & sPath & "user_home.asp?trantype=1"">VIEW PAYMENTS</a> | " & vbcrlf
	end if

	'View Submitted Action Line Requests Link
	if lcl_orghasfeature_action_line AND lcl_publicCanViewFeature_action_line then
       response.write "<a class=""accountmenu"" href=""" & sPath & "user_home.asp?trantype=0"">VIEW REQUESTS</a> | " & vbcrlf
	end if

	'View Shopping Cart (Purchases) Link
	if lcl_orghasfeature_activities AND lcl_publicCanViewFeature_activities then
       'response.write "<a class=""accountmenu"" href=""" & sPath & "classes/class_cart.asp"">VIEW CART</a> | " & vbcrlf
    response.write "<a class=""accountmenu"" href=""" & sPath & "rd_classes/class_cart.aspx"">VIEW CART</a> | " & vbcrlf
	end if

	if (lcl_orghasfeature_facilities  AND lcl_publicCanViewFeature_facilities) _ 
	OR (lcl_orghasfeature_activities  AND lcl_publicCanViewFeature_activities) _
	OR (lcl_orghasfeature_memberships AND lcl_publicCanViewFeature_memberships) _
	OR (lcl_orghasfeature_rentals AND lcl_publicCanViewFeature_rentals) _
	OR (lcl_orghasfeature_gifts       AND lcl_publicCanViewFeature_gifts) then
		response.write "<a class=""accountmenu"" href=""" & sPath & "purchases_report/purchases_list.asp"">VIEW PURCHASES</a> | " & vbcrlf
	end if

	'View Bids (Bid Postings) Link
 	if lcl_orghasfeature_bid_postings AND lcl_publicCanViewFeature_bid_postings then
     'response.write "<a class=""accountmenu"" href=""" & sPath & "classes/class_cart.asp"">VIEW CART</a> | " & vbcrlf
     response.write "<a class=""accountmenu"" href=""" & sPath & "view_bids.asp"">VIEW BIDS</a> | " & vbcrlf
 	end if

 'Do Not Knock List Link
  lcl_userid            = request.cookies("userid")
  lcl_canViewPeddlers   = checkAccessToList(lcl_userid, iorgid, "peddlers")
  lcl_canViewSolicitors = checkAccessToList(lcl_userid, iorgid, "solicitors")

  if   lcl_orghasfeature_donotknock AND lcl_publicCanViewFeature_donotknock _
  AND (lcl_canViewPeddlers OR lcl_canViewSolicitors) then
       response.write "<a class=""accountmenu"" href=""" & sPath & "view_donotknock.asp"">VIEW ""DO NOT KNOCK LIST""</a> | " & vbcrlf
  end if

 'Logout Link
  response.write "<a class=""accountmenu"" href=""" & sPath & "logout.asp"">LOGOUT</a>"  & vbcrlf

end sub

'------------------------------------------------------------------------------
'Sub ShowLoggedinLinks( ByVal sPath )
	'Manage Account Link
'	response.write "<a class=""accountmenu"" href=""" & sPath & "manage_account.asp"">MANAGE ACCOUNT</a> | " & vbcrlf

	'View Standard EGov Payments Link
'	If OrgHasFeature( iOrgId, "payments" ) And PublicCanViewFeature( iOrgId, "payments" ) Then 
'		response.write "<a class=""accountmenu"" href=""" & sPath & "user_home.asp?trantype=1"">VIEW PAYMENTS</a> | " & vbcrlf
'	End If 

	'View Submitted Action Line Requests Link
'	If OrgHasFeature( iOrgId, "action line" ) And PublicCanViewFeature( iOrgId, "action line" ) Then 
'		response.write "<a class=""accountmenu"" href=""" & sPath & "user_home.asp?trantype=0"">VIEW REQUESTS</a> | " & vbcrlf
'	End If 

	'View Shopping Cart (Purchases) Link
'	If OrgHasFeature( iOrgId, "activities" ) And PublicCanViewFeature( iOrgId, "activities" ) Then 
'		response.write "<a class=""accountmenu"" href=""" & sPath & "classes/class_cart.asp"">VIEW CART</a> | " & vbcrlf
'	End If 

'	If (OrgHasFeature( iOrgId, "facilities" )  And PublicCanViewFeature( iOrgId, "facilities" )) _ 
'	Or (OrgHasFeature( iOrgId, "activities" )  And PublicCanViewFeature( iOrgId, "activities" )) _
'	Or (OrgHasFeature( iOrgId, "memberships" ) And PublicCanViewFeature( iOrgId, "memberships" )) _
'	Or (OrgHasFeature( iOrgId, "gifts" )       And PublicCanViewFeature( iOrgId, "gifts" )) Then 
'		response.write "<a class=""accountmenu"" href=""" & sPath & "purchases_report/purchases_list.asp"">VIEW PURCHASES</a> | " & vbcrlf
'	End If 

	'View Bids (Bid Postings) Link
'	If orghasfeature(iorgid,"bid_postings") And publiccanviewfeature(iorgid,"bid_postings") Then 
	'   response.write "<a class=""accountmenu"" href=""" & sPath & "classes/class_cart.asp"">VIEW CART</a> | " & vbcrlf
'		response.write "<a class=""accountmenu"" href=""" & sPath & "view_bids.asp"">VIEW BIDS</a> | " & vbcrlf
'	End If 

	'Logout Link
'	response.write "<a class=""accountmenu"" href=""" & sPath & "logout.asp"">LOGOUT</a>"  & vbcrlf

'End Sub 


'------------------------------------------------------------------------------
' boolean PublicCanViewFeature( iOrgId, sFeature )
'------------------------------------------------------------------------------
Function PublicCanViewFeature( ByVal iOrgId, ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT FO.publiccanview FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid and orgid = " & iOrgId & " AND F.feature = '" & sFeature & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		PublicCanViewFeature = oRs("publiccanview")
	Else
		PublicCanViewFeature = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string RegisteredUserDisplay2()
'------------------------------------------------------------------------------
Function RegisteredUserDisplay2()
	' This is to allow the function to work from subdirectories

	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
		
		sSql = "SELECT useremail,USERFNAME,USERLNAME FROM egov_users WHERE userid = " & request.cookies("userid")

		Set name = Server.CreateObject("ADODB.Recordset")
		name.Open sSql, Application("DSN") , 3, 1
									
		response.write "<DIV style="" padding-left:5px;""><B>YOU ARE LOGGED IN AS: " & UCASE(NAME("USERFNAME")) & " " & UCASE(NAME("USERLNAME")) & ". &nbsp;&nbsp;"
		response.write "<A HREF=""../manage_account.ASP"">MANAGE ACCOUNT</A> | "
		
		If blnOrgPayment Then
			response.write "<A HREF=""../USER_HOME.ASP?trantype=1"">VIEW PAYMENTS</A> | "
		End If
		
		If blnOrgAction Then
			response.write "<A HREF=""../USER_HOME.ASP?trantype=0"">VIEW REQUESTS</A> | "
		End If

		' IF RECREATION MODULE ON - SHOW VIEW CART LINK
		If HasFeature(9,iorgid) Then
			response.write "<A HREF=""../classes/class_cart.asp"">VIEW CART</A> | "
		End If
		name.close
		set name = nothing

		response.write "<A HREF=""../LOGOUT.ASP"">LOGOUT</A></B></DIV>"

	End If 

End Function


'--------------------------------------------------------------------------------------------------
' string RegisteredUserDisplayWLevels( iLevels )
'--------------------------------------------------------------------------------------------------
Function RegisteredUserDisplayWLevels( ByVal iLevels )
	Dim sSql, oRs, sUserName

	' IF COOKIE FOUND DISPLAY MENU
	If request.cookies("userid") <> "" And request.cookies("userid") <> "-1" Then

		' GET FOLDER PATH			
		For l=1 to iLevels 
			sLevel = sLevel & "../"
		Next
		
		sSql = "SELECT userfname + ' ' + userlname as username FROM egov_users WHERE userid = '" & request.cookies("userid") & "'"

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN") , 3, 1

		If Not oRs.EOF Then
			sUserName = ucase(oRs("username"))
		Else
			sUserName = "NOT KNOWN"
		End If
		oRs.Close 
		Set oRs = Nothing
									
		' BUILD ACCOUNT MANAGE MENU
		response.write "<DIV class=accountmenu style="" padding-left:5px;"">"
		
		' WELCOME MESSAGE
		response.write "<img class=accountmenu src=""" & sLevel & "images/accountmenu.jpg"">"
		response.write "Welcome, <B>" & sUserName & "</B>!<br>"
		
		' MENU LINKS

		' MANAGE ACCOUNT LINK
		response.write "<A class=accountmenu HREF=""" & sLevel & "manage_account.asp"">MANAGE ACCOUNT</A> | "
		
		' VIEW STANDARD EGOV PAYMENTS LINK
		If blnOrgPayment Then
			response.write "<A class=accountmenu HREF=""" & sLevel & "user_home.asp?trantype=1"">VIEW PAYMENTS</A> | "
		End If
		
		' VIEW SUBMITTED ACTION LINE REQUESTS LINK
		If blnOrgAction Then
			response.write "<A class=accountmenu HREF=""" & sLevel & "user_home.asp?trantype=0"">VIEW REQUESTS</A> | "
		End If

		' VIEW SHOPPING CART LINK
		If HasFeature(9,iorgid) Then
			'response.write "<A class=accountmenu HREF=""" & sLevel & "classes/class_cart.asp"">VIEW CART</A> | "
			response.write "<A class=accountmenu HREF=""" & sLevel & "rd_classes/class_cart.aspx"">VIEW CART</A> | "
		End If
		
		' LOGOUT LINK
		response.write "<A class=accountmenu HREF=""" & sLevel & "logout.asp"">LOGOUT</A></DIV>"
	Else

		' NO COOKIE DON'T DISPLAY MENU
	End If 


End Function


'------------------------------------------------------------------------------------------------------------
' string GetOrgFeatureName( sFeature )
'------------------------------------------------------------------------------------------------------------
Function GetOrgFeatureName( ByVal sFeature )
	Dim sSql, oRs

	sSql = "SELECT isnull(FO.featurename,F.featurename) as featurename "
	sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " where FO.featureid = F.featureid and FO.orgid = " & iOrgId & " and feature = '" & sFeature & "'" 

	'response.write sSql
	'response.End 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetOrgFeatureName = oRs("featurename")
	Else
		GetOrgFeatureName = ""
	End If 

	oRs.close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' boolean OrgHasNeighborhoods( iOrgId )
'--------------------------------------------------------------------------------------------------
Function OrgHasNeighborhoods( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(neighborhoodid) AS hits FROM egov_neighborhoods WHERE orgid = " & iorgid 
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If CLng(oRs("hits")) > 0 Then
		OrgHasNeighborhoods = True 
	Else
		OrgHasNeighborhoods = False 
	End If 
	
	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' void ShowStatePicks sDefaultState, sState 
'--------------------------------------------------------------------------------------------------
Sub ShowStatePicks( ByVal sDefaultState, ByVal sState )
	Dim sSql, oRs, sPickState

	' if the citizen does not have a state, then use the state of the organization
	If sState = "" Then 
		sPickState = UCase(Trim(sDefaultState))
	Else
		sPickState = UCase(Trim(sState))
	End If 

	sSql = "SELECT statecode FROM states ORDER BY statecode"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("statecode") & """"
		If sPickState = UCase(Trim(oRs("statecode"))) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("statecode") & "</option>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' string GetPaymentGatewayName()
'------------------------------------------------------------------------------
Function GetPaymentGatewayName()
	Dim sSql, oRs

	sSql = "SELECT admingatewayname FROM egov_payment_gateways WHERE paymentgatewayid = " & iPaymentGatewayID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetPaymentGatewayName = oRs("admingatewayname")
	Else
		GetPaymentGatewayName = "PayPal"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetPaymentImage( sRelPath )
'------------------------------------------------------------------------------
Function GetPaymentImage( ByVal sRelPath )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(logopath,'') AS logopath FROM egov_payment_gateways "
	sSql = sSql & "WHERE paymentgatewayid = " & iPaymentGatewayID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("logopath") <> "" Then 
			GetPaymentImage = sRelPath & oRs("logopath")
		Else
			GetPaymentImage = ""
		End If 
	Else
		GetPaymentImage = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetProcessingRoute()
'------------------------------------------------------------------------------
Function GetProcessingRoute()
	Dim sSql, oRs

	sSql = "SELECT ISNULL(processingroute,'') AS processingroute FROM egov_payment_gateways "
	sSql = sSql & "WHERE paymentgatewayid = " & iPaymentGatewayID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetProcessingRoute = oRs("processingroute")
	Else
		GetProcessingRoute = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' void ShowCreditCardPicks
'------------------------------------------------------------------------------
Sub ShowCreditCardPicks()
	Dim sSql, oRs

	sSql = "SELECT C.creditcard FROM CreditCards C, egov_organizations_to_creditcards O "
	sSql = sSql & " WHERE O.creditcardid = C.creditcardid AND O.orgid = " & iOrgId & " ORDER BY creditcard"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""cardtype"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & LCase(oRs("creditcard")) & """>" & oRs("creditcard") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' string FormatEmailAsDecimal( sEmailAddress )
'------------------------------------------------------------------------------
Function FormatEmailAsDecimal( ByVal sEmailAddress )
	Dim x, sString

	If sEmailAddress <> "" Then 
		For x = 1 To Len(sEmailAddress)
			sString = sString & "&#" & Asc(Mid(sEmailAddress, x, 1)) & ";"
		Next 
		FormatEmailAsDecimal = sString
	Else
		FormatEmailAsDecimal = ""
	End If 

End Function 


'------------------------------------------------------------------------------
' string FormatMailToAsJavascript( sEmailAddress )
'------------------------------------------------------------------------------
Function FormatMailToAsJavascript( ByVal sEmailAddress )
	Dim aEmail

	If sEmailAddress <> "" Then
		If InStr(sEmailAddress, "@") > 0 Then 
			aEmail = Split(sEmailAddress, "@")

			FormatMailToAsJavascript = "href=""#"" onclick=""a='@'; this.href='mail'+'to:" & aEmail(0) & "'+a+'" & aEmail(1) & "';"""
		Else
			' THis is not an email address, so just make it an href
			FormatMailToAsJavascript = "href=""" & sEmailAddress & """"
		End If 
	Else
		FormatMailToAsJavascript = ""
	End If 

End Function


'------------------------------------------------------------------------------
' void MobileCheck
'------------------------------------------------------------------------------
Sub MobileCheck()
	' Check if this is a mobile device and route accordingly
	Dim sMobileRootUrl

	'Session("deviceViewMode") = ""

	If Session("deviceViewMode") <> "S" Then 
		If request("mobile") <> "no" Then
			' if org has mobile features turned on 
			If OrgHasMobileFeatures( iOrgId ) Then 
				' if this is a mobile user agent go to the mobile site
				'response.write "http_user_agent: " & request.servervariables("http_user_agent") & "<br /><br />"
				If UserAgentIsMobile( request.servervariables("http_user_agent") ) Then
					Session("deviceViewMode") = "M"
				Else
					' because devices upgrade constantly and the WRUFL does not, we need to compare to generic devices before giving up
					If UserAgentIsGenericMobileDevice( request.servervariables("http_user_agent") ) Then
						Session("deviceViewMode") = "M"
					Else 
						Session("deviceViewMode") = "S"
					End If 
				End If 
			Else
				Session("deviceViewMode") = "S"
			End If 
		Else
			Session("deviceViewMode") = "S"
		End If 
	End If 


	If Session("deviceViewMode") = "M" Then
		' redirect them to the root level mobile site page
		sMobileRootUrl = GetMobileRootURL( iOrgID )
		If sMobileRootUrl <> "" Then 
			'response.write "sMobileRootUrl: " & sMobileRootUrl & "<br /><br />"
			'response.End
			
			response.redirect sMobileRootUrl
		Else
			Session("deviceViewMode") = "S"
		End If 
	End If 

End Sub 


'------------------------------------------------------------------------------
' string GetMobileRootURL( iOrgID )
'------------------------------------------------------------------------------
Function GetMobileRootURL( ByVal iOrgID )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(orgmobilewebsiteurl,'') AS orgmobilewebsiteurl "
	sSql = sSql & "FROM Organizations WHERE orgid = " & iOrgId
	response.write sSql &  "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMobileRootURL = oRs("orgmobilewebsiteurl") 
	Else
		GetMobileRootURL = "" 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' boolean UserAgentIsMobile( sUserAgent )
'------------------------------------------------------------------------------
Function UserAgentIsMobile( ByVal sUserAgent )
	Dim sSql, oRs

	sSql = "SELECT deviceid FROM mobile_devices WHERE useragent = '" & Trim(Replace(sUserAgent,"'","''")) & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' if it is in the table then it is a mobile device of some kind
		UserAgentIsMobile = True 
	Else
		UserAgentIsMobile = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' boolean UserAgentIsGenericMobileDevice( sUserAgent )
'------------------------------------------------------------------------------
Function UserAgentIsGenericMobileDevice( ByVal sUserAgent )
	Dim sSql, oRs

	sSql = "SELECT genericdeviceid FROM mobile_generic_devices WHERE '" & LCase(Trim(Replace(sUserAgent,"'","''"))) & "' LIKE '%' + genericuseragent + '%'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' if it is in the table then it is a mobile device of some kind
		UserAgentIsGenericMobileDevice = True 
	Else
		UserAgentIsGenericMobileDevice = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' boolean OrgHasMobileFeatures( iOrgId )
'------------------------------------------------------------------------------
Function OrgHasMobileFeatures( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT hasmobilepages FROM Organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("hasmobilepages") Then
			OrgHasMobileFeatures = True 
		Else
			OrgHasMobileFeatures = False 
		End If 
	Else
		OrgHasMobileFeatures = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 





%>
