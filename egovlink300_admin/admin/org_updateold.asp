<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: org_update.asp
' AUTHOR: Steve Loar
' CREATED: 2/22/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the organization properties
'
' MODIFICATION HISTORY
' 1.0	02/22/2007	Steve Loar - INITIAL VERSION
' 1.1	10/22/08	David Boyer - Added "UserBids Upload Email"
' 1.2	08/28/2009	Steve Loar - Added Facility Reserve Periods
' 1.3	07/13/2010	Steve Loar - Added the deactivation save code
' 1.4	01/17/2011	Steve Loar - Added payment type fields for rentals out of the box functionality
' 1.4	01/17/2011	Steve Loar - Added View Full Site URL for Mobile
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Const adExecuteNoRecords = 128
Const adCmdStoredProc    = 4
Const adCmdText          = 1
Const adInteger          = 3
Const adVarChar          = 200
Const adLongVarChar      = 201
Const adDateTime         = 135
Const adParamReturnValue = 4
Const adParamInput       = 1
Const adParamOutput      = 2
Const adOpenStatic       = 3
Const adUseClient        = 3
Const adLockReadOnly     = 1
Const adStateOpen        = 1

Dim iOrgId, sOrgname, sOrgCity, sOrgState, sOrgPublicWebsiteURL, sOrgEgovWebsiteURL, sOrgTopGraphicRightURL, sOrgTopGraphicLeftURL
Dim sOrgWelcomeMessage, sOrgActionLineDescription, sOrgHeaderSize, sOrgPaymentDescription, sOrgPaymentGateway, sOrgVirtualSiteName1
Dim sOrgRequestCalOn, sOrgRequestCalForm, sOrgActionName, sOrgPaymentName, sOrgDocumentName, sOrgCalendarName
Dim sOrgCustomButtonsOn, sOrgRegistration, sOrgTimeZoneID, sOrgDisplayFooter, sOrgDisplayMenu, sOrgCustomMenu
Dim sOrgFacilityHoldDays, sDefaultEmail, sDefaultPhone, sdefaultcity, sdefaultstate, sdefaultzip, sinternal_default_contact
Dim sinternal_default_phone, sinternal_default_email, sLatitude, sLongitude, sSql, oCmd, sFirstName, sLastName, sUsername, sPassword
Dim iRootId, iAdminId, iAdminGroupId, iCityGroupId, sUsesAfter5Adjustment, sUsesWeekDays, sAllowedUnresolvedDays, sDefaultareacode
Dim iFacilityRootId, iMembershipId, iFacilityCategoryId, iGiftId, iGiftGroupId, sOrgGoogleAnalyticAccnt, spostings_email
Dim sEvaluationFormId, sFacilitySurveyFormId, sGoogleMapApiKey, sGoogleSearchID_documents, spostings_userbids_notifyemail
Dim sOrgPaymentOn, iResidentReservePeriod, iNonResidentReservePeriod, iPublicMenuOptCityHome_isEnabled, iPublicMenuOptEGovHome_isEnabled
Dim iPublicMenuOptCityHome_Label, iPublicMenuOptEGovHome_Label, sRentalSurveyFormId
Dim sOrgIsDeactivated, sShowMobileNavagation, sCitizenPaysFee
Dim sHasMobilePages, sOrgMobileWebsiteURL, sMobileLogo, sPublicDocumentsRoot, sViewFullSiteURL, sViewFullSiteLabel

iOrgId                    = request("orgid")
sOrgname                  = dbsafe(request("OrgName"))
sOrgCity                  = dbsafe(request("orgCity"))
sOrgState                 = UCase(request("OrgState"))
sOrgPublicWebsiteURL      = dbsafe(request("OrgPublicWebsiteURL"))
sOrgEgovWebsiteURL        = dbsafe(request("OrgEgovWebsiteURL"))
sOrgTopGraphicRightURL    = dbsafe(request("OrgTopGraphicRightURL"))
sOrgTopGraphicLeftURL     = dbsafe(request("OrgTopGraphicLeftURL"))
sOrgWelcomeMessage        = dbsafe(request("OrgWelcomeMessage"))
sOrgActionLineDescription = dbsafe(request("OrgActionLineDescription"))
sGoogleMapApiKey          = dbsafe(request("googlemapapikey"))
sGoogleSearchID_documents = dbsafe(request("googlesearchid_documents"))

sOrgHeaderSize = request("OrgHeaderSize")
If sOrgHeaderSize = "" Then 
	sOrgHeaderSize = "NULL"
End If 

sOrgPaymentDescription =  DBsafe(request("OrgPaymentDescription"))
sOrgPaymentGateway     = clng(request("OrgPaymentGateway"))
If clng(sOrgPaymentGateway) = clng(0) Then
	sOrgPaymentGateway = "NULL"
End If 
sOrgVirtualSiteName1 = dbsafe(request("OrgVirtualSiteName"))
If request("OrgRequestCalOn") = "on" Then 
	sOrgRequestCalOn = "1"
Else
	sOrgRequestCalOn = "0"
End If 
sOrgRequestCalForm = request("OrgRequestCalForm")
If sOrgRequestCalForm = "" Then 
	sOrgRequestCalForm = "NULL"
End If 
If request("OrgCustomButtonsOn") = "on" Then
	sOrgCustomButtonsOn = "1"
Else
	sOrgCustomButtonsOn = "0"
End If 
If request("OrgRegistration") = "on" Then 
	sOrgRegistration = "1"
Else
	sOrgRegistration = "0"
End If 

If request("citizenpaysfee") = "on" Then
	sCitizenPaysFee = "1"
Else
	sCitizenPaysFee = "0"
End If 

sOrgTimeZoneID = request("OrgTimeZoneID")

If request("OrgDisplayFooter") = "on" Then
	sOrgDisplayFooter = "1"
Else
	sOrgDisplayFooter = "0"
End If 

If request("OrgDisplayMenu") = "on" Then 
	sOrgDisplayMenu = "1"
Else
	sOrgDisplayMenu = "0"
End If 

If request("OrgCustomMenu") = "on" Then
	sOrgCustomMenu = "1"
Else
	sOrgCustomMenu = "0"
End If 

sOrgFacilityHoldDays = request("OrgFacilityHoldDays")
If sOrgFacilityHoldDays = "" Then 
	sOrgFacilityHoldDays = "NULL"
End If 

sDefaultEmail    = dbsafe(request("DefaultEmail"))
sDefaultPhone    = request("DefaultPhone")
sdefaultcity     = dbsafe(request("defaultcity"))
sdefaultstate    = dbsafe(UCase(request("defaultstate")))
sdefaultzip      = request("defaultzip")
sDefaultareacode = request("defaultareacode")

If CLng(iOrgId) > CLng(0) Then
	sinternal_default_contact = dbsafe(request("internal_default_contact"))
Else
	sinternal_default_contact = dbsafe(request("firstname")) & " " & dbsafe(request("lastname"))
End If

sinternal_default_phone = request("internal_default_phone")
sinternal_default_email = dbsafe(request("internal_default_email"))

If request("separate_index_catalog") = "on" Then 
	sseparate_index_catalog = "1"
Else
	sseparate_index_catalog = "0"
End If 

If request("usesafter5adjustment") = "on" Then 
	sUsesAfter5Adjustment = 1
Else
	sUsesAfter5Adjustment = 0
End If 

sLatitude = request("latitude")
If sLatitude = "" Then 
	sLatitude = "NULL"
End If 

sLongitude = request("longitude")
If sLongitude = "" Then 
	sLongitude = "NULL"
End If 

sFirstName = request("firstname")
sLastName  = request("lastname")
sUsername  = request("username")
sPassword  = request("password")

If request("usesweekdays") = "on" Then
	sUsesWeekDays = 1
Else
	sUsesWeekDays = 0
End If 

If request("OrgGoogleAnalyticAccnt") = "" Then 
  	sOrgGoogleAnalyticAccnt = "NULL"
Else
  	sOrgGoogleAnalyticAccnt = "'" & dbsafe(request("OrgGoogleAnalyticAccnt")) & "'"
End If 

If request("postings_email") = "" Then 
   spostings_email = "NULL"
Else 
   spostings_email = "'" & dbsafe(request("postings_email")) & "'"
End If

If request("postings_userbids_notifyemail") = "" Then 
   spostings_userbids_notifyemail = "NULL"
Else 
   spostings_userbids_notifyemail = "'" & dbsafe(request("postings_userbids_notifyemail")) & "'"
End If

If request("EvaluationFormId") = "" Then
	sEvaluationFormId = "NULL"
Else
	sEvaluationFormId = CLng(request("EvaluationFormId"))
End If 

If request("facilitysurveyformid") = "" Then
	sFacilitySurveyFormId = "NULL"
Else
	sFacilitySurveyFormId = CLng(request("facilitysurveyformid"))
End If

If request("rentalsurveyformid") = "" Then
	sRentalSurveyFormId = "NULL"
Else
	sRentalSurveyFormId = CLng(request("rentalsurveyformid"))
End If

If request("membershipcard_printer") = "" Then 
   smembercard_printer = 1
Else 
   smembershipcard_printer = CLng(request("membershipcard_printer"))
End If

If request("OrgPaymentOn") = "on" Then 
	sOrgPaymentOn = "1"
Else
	sOrgPaymentOn = "0"
End If 

If request("residentreserveperiod") <> "" Then 
	iResidentReservePeriod = request("residentreserveperiod")
Else
	' Old hard coded default
	iResidentReservePeriod = 12
End If

If request("residentreserveperiod") <> "" Then 
	iNonResidentReservePeriod = request("nonresidentreserveperiod")
Else
	' Old hard coded default
	iNonResidentReservePeriod = 6
End If

If request("public_menuopt_cityhome_enabled") = "on" Then 
  	iPublicMenuOptCityHome_isEnabled = "1"
Else 
  	iPublicMenuOptCityHome_isEnabled = "0"
End If 

If request("public_menuopt_egovhome_enabled") = "on" Then 
  	iPublicMenuOptEGovHome_isEnabled = "1"
Else 
  	iPublicMenuOptEGovHome_isEnabled = "0"
End If 

If request("public_menuopt_cityhome_label") = "" Then 
   iPublicMenuOptCityHome_Label = "'City Home'"
Else 
   iPublicMenuOptCityHome_Label = "'" & dbsafe(request("public_menuopt_cityhome_label")) & "'"
End If 

If request("public_menuopt_egovhome_label") = "" Then 
   iPublicMenuOptEGovHome_Label = "'E-Gov Home'"
Else 
   iPublicMenuOptEGovHome_Label = "'" & dbsafe(request("public_menuopt_egovhome_label")) & "'"
End If 

If request("isdeactivated") = "on" Then 
	sOrgIsDeactivated = "1"
Else
	sOrgIsDeactivated = "0"
End If 

'sAllowedUnresolvedDays = clng(request("allowedunresolveddays"))

If request("hasmobilepages") = "on" Then
	sHasMobilePages = "1"
Else
	sHasMobilePages = "0"
End If 

If request("orgmobilewebsiteurl") <> "" Then 
	sOrgMobileWebsiteURL = "'" & dbsafe(request("orgmobilewebsiteurl")) & "'"
Else
	sOrgMobileWebsiteURL = "NULL"
End If 

If request("mobilelogo") <> "" Then 
	sMobileLogo = "'" & dbsafe(request("mobilelogo")) & "'"
Else
	sMobileLogo = "NULL"
End If 

If request("publicdocumentsroot") <> "" Then 
	sPublicDocumentsRoot = "'" & dbsafe(request("publicdocumentsroot")) & "'"
Else
	sPublicDocumentsRoot = "NULL"
End If 

If request("viewfullsiteurl") <> "" Then 
	sViewFullSiteURL = "'" & dbsafe(request("viewfullsiteurl")) & "'"
Else
	sViewFullSiteURL = "NULL"
End If 

If request("viewfullsitelabel") <> "" Then 
	sViewFullSiteLabel = "'" & dbsafe(request("viewfullsitelabel")) & "'"
Else
	sViewFullSiteLabel = "NULL"
End If 

If request("showmobilenavagation") = "on" Then
	sShowMobileNavagation = "1"
Else
	sShowMobileNavagation = "0"
End If 


If CLng(iOrgId) > CLng(0) Then 
	' Update the existing one
	sSql = "UPDATE organizations SET "
	sSql = sSql & " OrgName = '"                        & sOrgname                         & "', "
	sSql = sSql & " orgCity = '"                        & sOrgCity                         & "', "
	sSql = sSql & " OrgState = '"                       & sOrgState                        & "', "
	sSql = sSql & " OrgPublicWebsiteURL = '"            & sOrgPublicWebsiteURL             & "', "
	sSql = sSql & " OrgEgovWebsiteURL = '"              & sOrgEgovWebsiteURL               & "', "
	sSql = sSql & " OrgTopGraphicRightURL = '"          & sOrgTopGraphicRightURL           & "', "
	sSql = sSql & " OrgTopGraphicLeftURL = '"           & sOrgTopGraphicLeftURL            & "', "
	sSql = sSql & " OrgWelcomeMessage = '"              & sOrgWelcomeMessage               & "', "
	sSql = sSql & " OrgActionLineDescription = '"       & sOrgActionLineDescription        & "', "
	sSql = sSql & " OrgHeaderSize = "                   & sOrgHeaderSize                   & ", "
	sSql = sSql & " OrgPaymentDescription = '"          & sOrgPaymentDescription           & "', "
	sSql = sSql & " OrgPaymentGateway = "               & sOrgPaymentGateway               & ", "
	sSql = sSql & " citizenpaysfee = "                  & sCitizenPaysFee                  & ", "
	sSql = sSql & " OrgVirtualSiteName = '"             & sOrgVirtualSiteName1             & "', "
	sSql = sSql & " OrgRequestCalOn = "                 & sOrgRequestCalOn                 & ", "
	sSql = sSql & " OrgPaymentOn = "                    & sOrgPaymentOn                    & ", "
	sSql = sSql & " OrgRequestCalForm = "               & sOrgRequestCalForm               & ", "
	sSql = sSql & " OrgTimeZoneID = "                   & sOrgTimeZoneID                   & ", "
	sSql = sSql & " OrgCustomButtonsOn = "              & sOrgCustomButtonsOn              & ", "
	sSql = sSql & " OrgRegistration = "                 & sOrgRegistration                 & ", "
	sSql = sSql & " OrgDisplayFooter = "                & sOrgDisplayFooter                & ", "
	sSql = sSql & " OrgDisplayMenu = "                  & sOrgDisplayMenu                  & ", "
	sSql = sSql & " OrgCustomMenu = "                   & sOrgCustomMenu                   & ", "
	sSql = sSql & " OrgFacilityHoldDays = "             & sOrgFacilityHoldDays             & ", "
	sSql = sSql & " DefaultEmail = '"                   & sDefaultEmail                    & "', "
	sSql = sSql & " DefaultPhone = '"                   & sDefaultPhone                    & "', "
	sSql = sSql & " defaultcity = '"                    & sdefaultcity                     & "', "
	sSql = sSql & " defaultstate = '"                   & sdefaultstate                    & "', "
	sSql = sSql & " defaultzip = '"                     & sdefaultzip                      & "', "
	sSql = sSql & " internal_default_contact = '"       & sinternal_default_contact        & "', "
	sSql = sSql & " internal_default_phone = '"         & sinternal_default_phone          & "', "
	sSql = sSql & " internal_default_email = '"         & sinternal_default_email          & "', "
	sSql = sSql & " separate_index_catalog = "          & sseparate_index_catalog          & ", "
	sSql = sSql & " latitude = "                        & sLatitude                        & ", "
	sSql = sSql & " longitude = "                       & sLongitude                       & ", "
	sSql = sSql & " usesafter5adjustment = "            & sUsesAfter5Adjustment            & ", "
	sSql = sSql & " usesweekdays = "                    & sUsesWeekDays                    & ", "
	sSql = sSql & " defaultareacode = '"                & sDefaultareacode                 & "', "
	sSql = sSql & " OrgGoogleAnalyticAccnt = "          & sOrgGoogleAnalyticAccnt          & ", "
	sSql = sSql & " postings_email = "                  & spostings_email			               & ", "
	sSql = sSql & " membershipcard_printer = "          & smembershipcard_printer          & ", "
	sSql = sSql & " EvaluationFormId = "                & sEvaluationFormId	               & ", "
	sSql = sSql & " googlemapapikey = '"                & sGoogleMapApiKey                 & "', "
 sSql = sSql & " googlesearchid_documents = '"       & sGoogleSearchID_documents        & "', "
	sSql = sSql & " facilitysurveyformid = "            & sFacilitySurveyFormId            & ", "
	sSql = sSql & " rentalsurveyformid = "              & sRentalSurveyFormId              & ", "
	sSql = sSql & " residentreserveperiod = "           & iResidentReservePeriod           & ", "
	sSql = sSql & " nonresidentreserveperiod = "        & iNonResidentReservePeriod        & ", "
	sSql = sSql & " postings_userbids_notifyemail = "   & spostings_userbids_notifyemail   & ", "
	sSql = sSql & " public_menuopt_cityhome_enabled = " & iPublicMenuOptCityHome_isEnabled & ", "
	sSql = sSql & " public_menuopt_egovhome_enabled = " & iPublicMenuOptEGovHome_isEnabled & ", "
	sSql = sSql & " public_menuopt_cityhome_label = "   & iPublicMenuOptCityHome_Label     & ", "
	sSql = sSql & " public_menuopt_egovhome_label = "   & iPublicMenuOptEGovHome_Label		& ", "
	sSql = sSql & " hasmobilepages = "					& sHasMobilePages					& ", "
	sSql = sSql & " orgmobilewebsiteurl = "				& sOrgMobileWebsiteURL				& ", "
	sSql = sSql & " mobilelogo = "						& sMobileLogo						& ", "
	sSql = sSql & " publicdocumentsroot = "				& sPublicDocumentsRoot				& ", "
	sSql = sSql & " viewfullsiteurl = "	  				& sViewFullSiteURL					& ", "
	sSql = sSql & " viewfullsitelabel = "	  			& sViewFullSiteLabel				& ", "
	sSql = sSql & " showmobilenavagation = "	  		& sShowMobileNavagation				& ", "
	sSql = sSql & " isdeactivated = "                   & sOrgIsDeactivated
	sSql = sSql & " WHERE orgid = " & iOrgId

	'response.write sSql
	RunSQL sSql 

	sSql = "DELETE FROM egov_organizations_to_creditcards WHERE orgid = " & iOrgId
	RunSQL sSql 

	'Set return success message
	lcl_success = "SU"

Else
	' Insert a new Org
	'response.write "Insert a new Org<br />"
	sSql = "INSERT INTO organizations ("
	sSql = sSql & " OrgName, "
	sSql = sSql & " orgCity, "
	sSql = sSql & " OrgState, "
	sSql = sSql & " OrgPublicWebsiteURL, "
	sSql = sSql & " OrgEgovWebsiteURL, "
	sSql = sSql & " OrgTopGraphicRightURL, "
	sSql = sSql & " OrgTopGraphicLeftURL, "
	sSql = sSql & " OrgWelcomeMessage, "
	sSql = sSql & " OrgActionLineDescription, "
	sSql = sSql & " OrgHeaderSize, "
	sSql = sSql & " OrgPaymentDescription, "
	sSql = sSql & " OrgPaymentGateway, "
	sSql = sSql & " citizenpaysfee, "
	sSql = sSql & " OrgVirtualSiteName, "
	sSql = sSql & " OrgRequestCalOn, "
	sSql = sSql & " OrgPaymentOn, "
	sSql = sSql & " OrgRequestCalForm, "
	sSql = sSql & " OrgTimeZoneID, "
	sSql = sSql & " OrgCustomButtonsOn, "
	sSql = sSql & " OrgRegistration, "
	sSql = sSql & " OrgDisplayFooter, "
	sSql = sSql & " OrgDisplayMenu, "
	sSql = sSql & " OrgCustomMenu, "
	sSql = sSql & " OrgFacilityHoldDays, "
	sSql = sSql & " DefaultEmail, "
	sSql = sSql & " DefaultPhone, "
	sSql = sSql & " defaultcity, "
	sSql = sSql & " defaultstate, "
	sSql = sSql & " defaultzip, "
	sSql = sSql & " defaultareacode, "
	sSql = sSql & " internal_default_contact, "
	sSql = sSql & " internal_default_phone, "
	sSql = sSql & " internal_default_email, "
	sSql = sSql & " separate_index_catalog, "
	sSql = sSql & " latitude, "
	sSql = sSql & " longitude, "
	sSql = sSql & " usesafter5adjustment, "
	sSql = sSql & " usesweekdays, "
	sSql = sSql & " OrgGoogleAnalyticAccnt, "
	sSql = sSql & " postings_email, "
	sSql = sSql & " membershipcard_printer, "
	sSql = sSql & " EvaluationFormId, "
	sSql = sSql & " googlemapapikey, "
 sSql = sSql & " googlesearchid_documents, "
	sSql = sSql & " facilitysurveyformid, "
	sSql = sSql & " rentalsurveyformid, "
	sSql = sSql & " residentreserveperiod, "
	sSql = sSql & " nonresidentreserveperiod, "
	sSql = sSql & " postings_userbids_notifyemail, "
	sSql = sSql & " public_menuopt_cityhome_enabled, "
	sSql = sSql & " public_menuopt_egovhome_enabled, "
	sSql = sSql & " public_menuopt_cityhome_label, "
	sSql = sSql & " public_menuopt_egovhome_label, "
	sSql = sSql & " hasmobilepages, "
	sSql = sSql & " orgmobilewebsiteurl, "
	sSql = sSql & " mobilelogo, "
	sSql = sSql & " publicdocumentsroot, "
	sSql = sSql & " viewfullsiteurl, "
	sSql = sSql & " viewfullsitelabel, "
	sSql = sSql & " showmobilenavagation, "
	sSql = sSql & " isdeactivated "
	sSql = sSql & " ) VALUES ( "
	sSql = sSql & "'" & sOrgname                         & "', "
	sSql = sSql & "'" & sOrgCity                         & "', "
	sSql = sSql & "'" & sdefaultstate                    & "', "
	sSql = sSql & "'" & sOrgPublicWebsiteURL             & "', "
	sSql = sSql & "'" & sOrgEgovWebsiteURL               & "', "
	sSql = sSql & "'" & sOrgTopGraphicRightURL           & "', "
	sSql = sSql & "'" & sOrgTopGraphicLeftURL            & "', "
	sSql = sSql & "'" & sOrgWelcomeMessage               & "', "
	sSql = sSql & "'" & sOrgActionLineDescription        & "', "
	sSql = sSql &       sOrgHeaderSize                   & ", "
	sSql = sSql & "'" & sOrgPaymentDescription           & "', "
	sSql = sSql &       sOrgPaymentGateway               & ", "
	sSql = sSql &       sCitizenPaysFee                  & ", "
	sSql = sSql & "'" & sOrgVirtualSiteName1             & "', "
	sSql = sSql &       sOrgRequestCalOn                 & ", "
	sSql = sSql &       sOrgPaymentOn                    & ", "
	sSql = sSql &       sOrgRequestCalForm               & ", "
	sSql = sSql &       sOrgTimeZoneID                   & ", "
	sSql = sSql &       sOrgCustomButtonsOn              & ", "
	sSql = sSql &       sOrgRegistration                 & ", "
	sSql = sSql &       sOrgDisplayFooter                & ", "
	sSql = sSql &       sOrgDisplayMenu                  & ", "
	sSql = sSql &       sOrgCustomMenu                   & ", "
	sSql = sSql &       sOrgFacilityHoldDays             & ", "
	sSql = sSql & "'" & sDefaultEmail                    & "', "
	sSql = sSql & "'" & sDefaultPhone                    & "', "
	sSql = sSql & "'" & sdefaultcity                     & "', "
	sSql = sSql & "'" & sdefaultstate                    & "', "
	sSql = sSql & "'" & sdefaultzip                      & "', "
	sSql = sSql & "'" & sDefaultareacode                 & "', "
	sSql = sSql & "'" & sinternal_default_contact        & "', "
	sSql = sSql & "'" & sinternal_default_phone          & "', "
	sSql = sSql & "'" & sinternal_default_email          & "', "
	sSql = sSql &       sseparate_index_catalog          & ", "
	sSql = sSql &       sLatitude                        & ", "
	sSql = sSql &       sLongitude                       & ", "
	sSql = sSql &       sUsesAfter5Adjustment            & ", "
	sSql = sSql &       sUsesWeekDays                    & ", "
	sSql = sSql &       sOrgGoogleAnalyticAccnt          & ", "
	sSql = sSql &       spostings_email                  & ", "
	sSql = sSql &       smembershipcard_printer          & ", "
	sSql = sSql &       sEvaluationFormId                & ", "
	sSql = sSql & "'" & sGoogleMapApiKey                 & "', "
 sSql = sSql & "'" & sGoogleSearchID_documents        & "', "
	sSql = sSql &       sFacilitySurveyFormId            & ", "
	sSql = sSql &       sRentalSurveyFormId              & ", "
	sSql = sSql &       iResidentReservePeriod           & ", "
	sSql = sSql &       iNonResidentReservePeriod        & ", "
	sSql = sSql &       spostings_userbids_notifyemail   & ", "
	sSql = sSql &       iPublicMenuOptCityHome_isEnabled & ", "
	sSql = sSql &       iPublicMenuOptEGovHome_isEnabled & ", "
	sSql = sSql &       iPublicMenuOptCityHome_Label     & ", "
	sSql = sSql &       iPublicMenuOptEGovHome_Label     & ", "
	sSql = sSql &       sHasMobilePages					& ", "
	sSql = sSql &       sOrgMobileWebsiteURL			& ", "
	sSql = sSql &       sMobileLogo						& ", "
	sSql = sSql &       sPublicDocumentsRoot			& ", "
	sSql = sSql &       sViewFullSiteURL				& ", "
	sSql = sSql &       sViewFullSiteLabel				& ", "
	sSql = sSql &       sShowMobileNavagation			& ", "
	sSql = sSql &	    sOrgIsDeactivated
	sSql = sSql & " )"
	'response.write sSql

	iOrgId = RunIdentityInsert( sSql )
	'response.write iOrgId
	'response.flush

	' Set up the root user
	'response.write "<br />Set up the root user<br />"
	sSql = "Insert Into users ( orgid, username, password, lastname, firstname, isrootadmin, pagesize ) Values ( " & iOrgId & ", 'eclink', 'ecl1nk05', 'ECLink', 'Admin', 1, 20)"
	iRootId = RunIdentityInsert( sSql )
	'response.write iRootId
	'response.flush

	' Set up the Admin User
	'response.write "<br />Set up the Admin User<br />"
	sSql = "Insert Into users ( orgid, username, password, lastname, firstname, isrootadmin, email, pagesize ) Values ( " & iOrgId & ", '" & sUsername & "', '" & sPassword & "', '" & sLastName & "', '" & sFirstName & "', 0, '" & sinternal_default_email & "', 20)"
	iAdminId = RunIdentityInsert( sSql )
	'response.write iAdminId
	'response.flush

	' Create the Admin Group
	'response.write "<br />Create the Admin Group<br />"
	sSql = "Insert Into groups( orgid, groupname, groupdescription, grouptype ) Values ( " & iOrgId & ", 'Administrators', 'This Group has full administrative privledges', 2)"
	iAdminGroupId = RunIdentityInsert( sSql )

	' Create the GroupsRoles for the admins - This makes initial Permissions assignment work
	sSql = "Insert Into groupsroles ( groupid, roleid ) Values ( " & iAdminGroupId & ", 53 )"
	RunSQL sSql

	' Create the City Employees Group
	'response.write "<br />Create the City Employees Group<br />"
	sSql = "Insert Into groups ( orgid, groupname, groupdescription, grouptype ) Values ( " & iOrgId & ", 'City Employees', 'This Department contains City Employees', 2)"
	iCityGroupId = RunIdentityInsert( sSql )

	' Create the Admin UsersGroups Row
	'response.write "<br />Create the Admin UsersGroups Row<br />"
	sSql = "Insert Into usersgroups ( userid, groupid, isprimarygroup ) Values ( " & iRootId & ", " & iAdminGroupId & ", 0 )"
	RunSQL sSql
	sSql = "Insert Into usersgroups ( userid, groupid, isprimarygroup ) Values ( " & iAdminId & ", " & iAdminGroupId & ", 0 )"
	RunSQL sSql

	' Create the City UserGorups Row
	'response.write "<br />Create the City UserGorups Row<br />"
	'response.flush
	sSql = "Insert Into usersgroups ( userid, groupid, isprimarygroup ) Values ( " & iRootId & ", " & iCityGroupId & ", 0 )"
	RunSQL sSql
	sSql = "Insert Into usersgroups ( userid, groupid, isprimarygroup ) Values ( " & iAdminId & ", " & iCityGroupId & ", 0 )"
	RunSQL sSql

	' Copy ActionLine Forms
	'response.write "<br />Copy ActionLine Forms<br />"
	'response.flush
	CopyAllForms iOrgId, iAdminId, iCityGroupId

	' Put Forms into Categories
	'response.write "<br />Put Forms into Categories<br />"
	'response.flush
	PutFormsIntoCategories iOrgId 

	' Sync the Files in the documents folders
	'response.write "<br />Sync the Files in the documents folders<br />"
	'response.flush
	'FileSync iOrgId, sOrgVirtualSiteName1


	'***********************************************************************************
	' Recreation Table Setup so they have basic functionality out of the box.
	'***********************************************************************************
	' Facilities
	sSql = "INSERT INTO egov_recreation_categories ( categorytitle, categorydescription, isroot, sequenceid, orgid) VALUES ( 'Facility Reservations', 'Contact City Hall', 1, 1, " & iOrgId & " )"
	iFacilityRootId = RunIdentityInsert( sSql )
	
	sSql = "INSERT INTO egov_recreation_categories ( categorytitle, categorydescription, isroot, sequenceid, orgid) VALUES ( 'Lodges', 'We have some lodges to rent.', 0, 2, " & iOrgId & " )"
	iFacilityCategoryId = RunIdentityInsert( sSql )

	sSql = "INSERT INTO egov_recreation_category_to_subcategory ( recreationcategoryid, recreationsubcategoryid ) VALUES ( " & iFacilityRootId & ", " & iFacilityCategoryId & " )"
	RunSQL sSql

	' Classes
	sSql = "INSERT INTO egov_class_categories ( categorytitle, categorydescription, isroot, sequenceid, orgid ) VALUES ( 'Classes/Events', 'Classes/Events', 1, 1, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_price_types ( pricetype, pricetypename, pricetypedescription, displayorder, "
	sSql = sSql & "checkresidency, isresident, isactiveforfacility, isactiveforclasses, needsregistrationstartdate, "
	sSql = sSql & "orgid, isforrentals, alwaysadd ) VALUES ( "
	sSql = sSql & "'R', 'Resident', 'Resident', 1, 1, 1, 1, 1, 1, " & iOrgId & ", 1, 1 )"
	RunSQL sSql

	sSql = "INSERT INTO egov_price_types ( pricetype, pricetypename, pricetypedescription, displayorder, "
	sSql = sSql & "checkresidency, isresident, isactiveforfacility, isactiveforclasses, needsregistrationstartdate, "
	sSql = sSql & "orgid, isforrentals, alwaysadd ) VALUES ( "
	sSql = sSql & "'N', 'Nonresident', 'Nonresident', 2, 1, 0, 1, 1, 1, " & iOrgId & ", 1, 0 )"
	RunSQL sSql

	sSql = "INSERT INTO egov_price_types ( pricetype, pricetypename, pricetypedescription, displayorder, "
	sSql = sSql & "checkresidency, isresident, isactiveforfacility, isactiveforclasses, needsregistrationstartdate, "
	sSql = sSql & "orgid, isforrentals, alwaysadd ) VALUES ( "
	sSql = sSql & "'E', 'Everyone', 'Everyone', 3, 0, 0, 0, 1, 1, " & iOrgId & ", 1, 1 )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, orgid ) VALUES ( 1, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, orgid ) VALUES ( 2, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, orgid ) VALUES ( 3, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, orgid ) VALUES ( 4, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, orgid ) VALUES ( 5, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, orgid ) VALUES ( 6, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_organizations_to_paymenttypes ( paymenttypeid, defaultamount, orgid ) VALUES ( 7, 15.00, " & iOrgId & " )"
	RunSQL sSql

	' Memberships
	sSql = "INSERT INTO egov_memberships ( membership, membershipdesc, introtext, orgid ) VALUES ( 'pool', 'Pool', 'Purchase your Pool Pass', " & iOrgId & " )"
	iMembershipId = RunIdentityInsert( sSql )

	sSql = "INSERT INTO egov_poolpassresidenttypes ( resident_type, description, public_display, displayorder, orgid ) VALUES ( 'R', 'Resident', 1, 1, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_poolpassresidenttypes ( resident_type, description, public_display, displayorder, orgid ) VALUES ( 'N', 'Non Resident', 1, 2, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_membership_rate_displays ( resident_type, public_display, membershipid ) VALUES ( 'R', 1, " & iMembershipId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_membership_rate_displays ( resident_type, public_display, membershipid ) VALUES ( 'N', 1, " & iMembershipId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_membership_periods ( period_desc, is_seasonal, orgid ) VALUES ( 'Season', 1, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_familymember_relationships ( relationship, displayorder, selftag, isdefault, orgid ) VALUES ( 'Spouse', 1, 'Yourself', 1, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_familymember_relationships ( relationship, displayorder, selftag, isdefault, orgid ) VALUES ( 'Child', 2, 'Child', 0, " & iOrgId & " )"
	RunSQL sSql

	' Gifts
	sSql = "INSERT INTO egov_gift ( giftname, giftdescription, paymentformid, amount, orgid ) VALUES ( 'Commemorative Brick Program', '<p><table border=""0"" cellpadding=""3"" cellspacing=""3"" id=""table2"" width=""100%""><tbody><tr><td align=""right"" valign=""top"" width=""46%""><img border=""0"" height=""257"" src=""http://www.egovlink.com/parkcity/admin/custom/pub/parkcity/unpublished_documents/gifts_bricks.jpg"" width=""310""/></td><td valign=""top"" width=""54%""><font size=""2"">A $25.00 gift will install a Commemorative Brick in honor of a special someone in your life.</font> <p><font size=""2"">Your gift will provide beauty and value, while enhancing the City Hall building.&nbsp; Most of all, your gift will be a legacy to our City, helping create an attractive and pleasant place to live, work, and play.</font></p><p><font size=""2"">An acknowledgement card will be sent to the person or family of the individual honored, as well as to the contributor.</font></p></td></tr></tbody></table></p>', 1, 25.00, " & iOrgId & " )"
	iGiftId = RunIdentityInsert( sSql )

	sSql = "INSERT INTO egov_gift_group ( giftgroupname, sequence, orgid, giftid ) VALUES ( 'Commemorative Plaque Information', 1, " & iOrgId & ", " & iGiftId & " )"
	iGiftGroupId = RunIdentityInsert( sSql )

	sSql = "INSERT INTO egov_gift_fields ( fieldprompt, fieldtype, isrequired, validation, helptext, maxfieldsize, giftid, groupid ) VALUES ( 'Line 1', 1, 1, 'thirteenchar', 'Max. 12 Characters', 12, " & iGiftId & ", " & iGiftGroupId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_gift_fields ( fieldprompt, fieldtype, isrequired, validation, helptext, maxfieldsize, giftid, groupid ) VALUES ( 'Line 2', 1, 1, 'thirteenchar', 'Max. 12 Characters', 12, " & iGiftId & ", " & iGiftGroupId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_gift_fields ( fieldprompt, fieldtype, isrequired, validation, helptext, maxfieldsize, giftid, groupid ) VALUES ( 'Line 3', 1, 1, 'thirteenchar', 'Max. 12 Characters', 12, " & iGiftId & ", " & iGiftGroupId & " )"
	RunSQL sSql

	' Other tables that make it all work. These will need to be manually updated later.
	sSql = "INSERT INTO egov_paymentservices ( paymentservicename, paymentservicedescription, paymentserviceenabled, paymentservice_type, orgid ) VALUES ( 'Facility Reservation', 'recreation', 0, 3, " & iOrgId & " )"
	RunSQL sSql
	
	sSql = "INSERT INTO egov_paymentservices ( paymentservicename, paymentservicedescription, paymentserviceenabled, paymentservice_type, orgid ) VALUES ( 'Commerative Gift', 'recreation', 0, 1, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_paymentservices ( paymentservicename, paymentservicedescription, paymentserviceenabled, paymentservice_type, orgid ) VALUES ( 'Pool Pass', 'recreation', 0, 2, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_paymentservices ( paymentservicename, paymentservicedescription, paymentserviceenabled, paymentservice_type, orgid ) VALUES ( 'Classes and Events', 'recreation', 0, 4, " & iOrgId & " )"
	RunSQL sSql

	sSql = "INSERT INTO egov_verisign_options ( vendor, [user], password, partner, IsLive, liveurl, orgid ) VALUES ( '" & sOrgVirtualSiteName1 & "', '" & sOrgVirtualSiteName1 & "', '" & sOrgVirtualSiteName1 & "', 'paypal', 1, 'payflow.verisign.com', " & iOrgId & " )"
	RunSQL sSql

'	response.write "<p>All Done!</p>"
'	response.flush

	'Set return success message
	lcl_success = "SA"

End If 

' Add in the credit card picks
For Each item In request("creditcardid")
	sSql = "INSERT INTO egov_organizations_to_creditcards ( orgid, creditcardid ) VALUES ( " & iOrgId & ", " & item & " )"
	RunSQL sSql
Next 

' Back to the edit page
response.redirect "edit_org.asp?orgid=" & iOrgId & "&success=" & lcl_success


'------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then 
		DBsafe = strDB 
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 

End Function


'------------------------------------------------------------------------------
' void RunSQL sSql 
'------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


'------------------------------------------------------------------------------
' integer RunIdentityInsert( sInsertStatement )
'------------------------------------------------------------------------------
Function RunIdentityInsert( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oRs

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"
	'dtb_debug(sSql)
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oRs("ROWID")
	oRs.close
	Set oRs = Nothing

	RunIdentityInsert = iReturnValue

End Function


'------------------------------------------------------------------------------
' void CopyAllForms iNewOrg, iAdminId, iDeptId 
'------------------------------------------------------------------------------
Sub CopyAllForms( ByVal iNewOrg, ByVal iAdminId, ByVal iDeptId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_action_request_forms WHERE orgid = 47 " 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 2
		
	If Not oRs.EOF Then
		Do While Not oRs.EOF 
			subCopyForm oRs("action_form_id"), iNewOrg, iAdminId, iDeptId 
			oRs.MoveNext
		Loop 
	End If

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void subCopyForm iFormID, iNewOrg, iAdminId, iDeptId
'------------------------------------------------------------------------------
Sub subCopyForm( ByVal iFormID, ByVal iNewOrg, ByVal iAdminId, ByVal iDeptId )
	Dim iNewFormID, iNewCategoryID
	
	' COPY the FORMS
	iNewFormID = fnCopySQLDataRow( "action_form_id", iFormID, "egov_action_request_forms", iNewOrg, iAdminId, iDeptId )

	' COPY FORM QUESTIONS
	SubCopyFormQuestions iFormID, iNewOrg, iNewFormID 

	' GET CATEGORY ID 
	iNewCategoryID = GetCategoryID( iFormID )

	'CODE TO ASSIGN FORM TO CATEGORY
	subAssignFormtoCategory iNewFormID, iNewCategoryID, iNewOrg 

End Sub


'------------------------------------------------------------------------------
' integer fnCopySQLDataRow( sPrimaryKey, iPrimaryKeyID, sTableName, iNewOrg, iAdminId, iDeptId )
'------------------------------------------------------------------------------
Function fnCopySQLDataRow( ByVal sPrimaryKey, ByVal iPrimaryKeyID, ByVal sTableName, ByVal iNewOrg, ByVal iAdminId, ByVal iDeptId )
	Dim iReturnValue, oRs, sSql, fldLoop, iCount, sValueList, sInsertStatement

	iReturnValue = 0 
	iCount = 0

	sSql = "SELECT * FROM " & sTableName & " WHERE " & sPrimaryKey & " = '" & iPrimaryKeyID & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then

		' BUILD COLUMN NAME LIST
		For Each fldLoop In oRs.Fields
			iCount = iCount + 1 
			If iCount <> 1 Then
				sFieldList = sFieldList & fldLoop.Name 
				If iCount <> oRs.Fields.Count Then
					sFieldList = sFieldList & ","
				End If
			End If
		Next

		' WRITE DATA ROWS
		Do while Not oRs.EOF 
			iCount = 0
			sValueList = ""
			If Not oRs.EOF Then
				For Each fldLoop in oRs.Fields
					iCount = iCount + 1 
					If iCount <> 1 Then ' SKIP FIRST FIELD (AUTO IDENTITY)
						
						' CUSTOM FIELD HANDLING
						Select Case LCase(fldLoop.Name)

							'Case "action_form_name"
								'sValueList = sValueList & "'NEW ACTION FORM'"

							Case "orgid" ' Try to use the new orgid
								sValueList = sValueList & "'" & iNewOrg & "'"
							
							Case "assigned_userid"
								sValueList = sValueList & "'" & iAdminId & "'"

							Case "assigned_userid2"
								sValueList = sValueList & "'" & 0 & "'"
							
							Case "assigned_userid3"
								sValueList = sValueList & "'" & 0 & "'"

							Case "deptid"
								sValueList = sValueList & "'" & iDeptId & "'"

							Case Else

								' ADD DATA TO STRING
								If fldLoop.Type = 11 Then
									sValueList = sValueList & "'" & fnBitConvert(fldLoop.Value) & "'"
								Else
									sValueList = sValueList & "'" & DBSafe(fldLoop.Value) & "'"
								End If

						End Select 

						' ADD TRAILING COMMA IF NECESSARY
						If iCount <> oRs.Fields.Count Then
							sValueList = sValueList & ","
						End If
					End If
				Next

				sInsertStatement = "INSERT INTO " & sTableName & " (" & sFieldList & ") VALUES (" & sValueList & ")"
				'DEBUG DATA: 
				'response.write sInsertStatement & "<br /><br />" & vbcrlf
				'response.flush
				
				iReturnValue = RunIdentityInsert( sInsertStatement )

				'DEBUG DATA: response.write "(" & iReturnValue & ")"
				
				oRs.MoveNext
			End If
		Loop

	End If

	oRs.Close
	Set oRs = Nothing 

	fnCopySQLDataRow = iReturnValue

End Function


'------------------------------------------------------------------------------
' SUB SUBCOPYFORMQUESTIONS(IFORMID,IORGID,INEWFORMID)
'------------------------------------------------------------------------------
Sub SubCopyFormQuestions( ByVal iFormID, ByVal iOrgID, ByVal iNewFormID )
	Dim sSql, oRs, iTempReturnID

	sSql = "SELECT * FROM egov_action_form_questions WHERE formid = '" & iFormID & "' ORDER BY SEQUENCE"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 2
	
	If Not oRs.EOF Then
		Do While NOT oRs.EOF 
			iTempReturnID = fnCopySQLQuestionRow( "questionid", oRs("questionid"), "egov_action_form_questions", iNewFormID, iOrgID )
			oRs.MoveNext
		Loop 
	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' integer  FNBITCONVERT( BLNVALUE )
'------------------------------------------------------------------------------
Function  fnBitConvert( ByVal blnValue )
	Dim iReturnValue

	iReturnValue = 0

	If blnValue Then
		iReturnValue = 1
	End If	

	fnBitConvert = iReturnValue

End Function


'------------------------------------------------------------------------------
' integer fnCopySQLQuestionRow( SPRIMARYKEY, IPRIMARYKEYID, STABLENAME, IFORMID, IORGID ) 
'------------------------------------------------------------------------------
Function fnCopySQLQuestionRow( ByVal sPrimaryKey, ByVal iPrimaryKeyID, ByVal sTableName, ByVal iFormID, ByVal iOrgID ) 
	Dim iReturnValue, sSql, oRs, fldLoop, iCount, sValueList, sValue, sInsertStatement

	iReturnValue = 0 
	iCount = 0

	sSql = "SELECT * FROM " & sTableName & " WHERE " & sPrimaryKey & " = '" & iPrimaryKeyID & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then

		' BUILD COLUMN NAME LIST
		For Each fldLoop in oRs.Fields
			iCount = iCount + 1 
			If iCount <> 1 Then
				sFieldList = sFieldList & fldLoop.Name 
				If iCount <> oRs.Fields.Count Then
					sFieldList = sFieldList & ","
				End If
			End If
		Next

		' WRITE DATA ROWS
		Do while Not oRs.EOF 
			iCount = 0
			sValueList = ""
			If Not oRs.EOF Then
				For Each fldLoop In oRs.Fields
					iCount = iCount + 1 
					If iCount <> 1 Then ' SKIP FIRST FIELD (AUTO IDENTITY)
						
						sValue = DbSafe(fldLoop.Value)

						' CHANGE FORMID AND ORGID VALUES
						If fldLoop.Name = "formid" Then
							sValue = iFormID
						End If

						If fldLoop.Name = "orgid" Then
							sValue = iOrgID
						End If

						' ADD DATA TO STRING
						If fldLoop.Type = 11 Then
							sValueList = sValueList & "'" & fnBitConvert(sValue) & "'"
						Else
							sValueList = sValueList & "'" & sValue & "'"
						End If

						' ADD TRAILING COMMA IF NECESSARY
						If iCount <> oRs.Fields.Count Then
							sValueList = sValueList & ","
						End If
					End If
				Next

				sInsertStatement = "INSERT INTO " & sTableName & " (" & sFieldList & ") VALUES (" & sValueList & ")"
				'DEBUG DATA: response.write sInsertStatement & "<BR>"
				
				' INSERT NEW ROW INTO DATABASE AND GET ROWID
				iReturnValue = RunIdentityInsert( sInsertStatement )

				oRs.MoveNext
			End If
		Loop

	End If

	oRs.Close
	Set oRs = Nothing 

	fnCopySQLQuestionRow = iReturnValue

End Function


'------------------------------------------------------------------------------
' void subAssignFormtoCategory iFormID, iCategoryID, iOrgId
'------------------------------------------------------------------------------
Sub subAssignFormtoCategory( ByVal iFormID, ByVal iCategoryID, ByVal iOrgId )
	Dim sSql, oRs

	' INSERT NEW 
	sSql = "INSERT INTO egov_forms_to_categories (form_category_id,action_form_id, orgid) VALUES ('" & iCategoryID & "','" & iFormID & "','" & iOrgId & "')"
	RunSQL sSql

	' INSERT NEW
	sSql = "INSERT INTO egov_organizations_to_forms (orgid,action_form_id,action_form_enabled) VALUES ('" & iOrgId & "','" & iFormID & "','1')"
	RunSQL sSql

End Sub


'------------------------------------------------------------------------------
' integer GetCategoryID( iFormID )
'------------------------------------------------------------------------------
Function GetCategoryID( ByVal iFormID )
	Dim iReturnValue, sSql, oRs

	iReturnValue = 0

	sSql = "SELECT form_category_id FROM egov_forms_to_categories "
	sSql = sSql & "WHERE action_form_id = " & iFormID 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		iReturnValue = oRs("form_category_id")
	End If

	oRs.Close
	Set oRs = Nothing 

	GetCategoryID = iReturnValue

End Function


'------------------------------------------------------------------------------
' void PutFormsIntoCategories  iOrgId 
'------------------------------------------------------------------------------
Sub PutFormsIntoCategories( ByVal iOrgId )
	Dim iCitizenId, iOtherId, iRequestId, iRepairId, iNuisanceId, iAnimal, iLicenseId, iFormCategoryId, sSql, oRs

	iCitizenId = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Citizen Comments and Concerns',1," & iOrgid & ")")
	response.write "Citizen Comments and Concerns &ndash; " & iCitizenId & "<br/>"
'	response.flush

	iRequestId = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Requests for Information',2," & iOrgid & ")")
	response.write "Requests for Information &ndash; " & iRequestId & "<br/>"
'	response.flush

	iRepairId = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Repairs and Requests for Service',3," & iOrgid & ")")
	response.write "Repairs and Requests for Service &ndash; " & iRepairId & "<br/>"
'	response.flush

	iNuisanceId = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Nuisance/Code Violations',4," & iOrgid & ")")
	response.write "Nuisance/Code Violations &ndash; " & iNuisanceId & "<br/>"
'	response.flush

	iAnimal = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Animal Control',5," & iOrgid & ")")
	response.write "Animal Control &ndash; " & iAnimal & "<br/>"
'	response.flush

	iLicenseId = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Licenses',6," & iOrgid & ")")
	response.write "Licenses &ndash; " & iLicenseId & "<br/>"
'	response.flush

	iOtherId = RunIdentityInsert("insert into egov_form_categories (form_category_name, form_category_sequence, orgid) values ('Other',7," & iOrgid & ")")
	response.write "Other &ndash; " & iOtherId & "<br/>"
'	response.flush

	response.write "<p><b>Updating the Form Categories.</b></p>"
	' Update the egov_form_categories table with the new catagories

	sSql = "SELECT egov_rowid, form_category_id, form_category_name "
	sSql = sSql & "FROM egov_forms_categories_view WHERE orgid = " & iorgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			Select Case oRs("form_category_name")
				Case "Citizen Comments and Concerns"
					iFormCategoryId = iCitizenId
				Case "Requests for Information"
					iFormCategoryId = iRequestId
				Case "Repairs and Requests for Service"
					iFormCategoryId = iRepairId
				Case "Nuisance/Code Violations"
					iFormCategoryId = iNuisanceId
				Case "Animal Control"
					iFormCategoryId = iAnimal
				Case "Licenses"
					iFormCategoryId = iLicenseId
				Case "Other"
					iFormCategoryId = iOtherId
				Case Else 
					iFormCategoryId = iOtherId
			End Select 

			response.write "(" & oRs("egov_rowid") & ") " & oRs("form_category_name") & " &ndash; " & iFormCategoryId & "<br />"
			UpdateCategoryId oRs("egov_rowid"), iFormCategoryId  
			oRs.movenext
		Loop 
	Else
		response.write "<h2>Error &ndash; No forms were found to place in the categories.</h2>Clear out the categories, edit and run copy_all_ques.asp, then run this script again.<br /><br />"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void UpdateCategoryId iEgovRowId, iFormCategoryId 
'------------------------------------------------------------------------------
Sub  UpdateCategoryId( ByVal iEgovRowId, ByVal iFormCategoryId )
	Dim sSql, oCmd

	' update the formcategoryid
	sSql = "UPDATE egov_forms_to_categories SET form_category_id = " & iFormCategoryId
	sSql = sSql & " WHERE egov_rowid =" & iEgovRowId  

	RunSQL sSql

End Sub 


'------------------------------------------------------------------------------
' Sub FileSync( iOrgId, sOrgVirtualSiteName )
'------------------------------------------------------------------------------
Sub FileSync( ByVal iOrgId, ByVal sOrgVirtualSiteName )
	' Stolen from John's filesync.asp
	' INITIALIZE VALUES AND OBJECTS

	Dim sPath, sVirtualPath, blnOverwrite

	blnOverwrite = True
	sPath ="D:\wwwroot\www.cityegov.com\egovlink300_admin\custom\pub\" & sOrgVirtualSiteName
	sVirtualPath = "/public_documents300/custom/pub/" & sOrgVirtualSiteName

'	sDSN = "Driver={SQL Server}; Server=ISPS0014; Database=egovlink300; UID=egovsa; PWD=egov_4303;"
	Set FSO = CreateObject("Scripting.FileSystemObject")


'	response.write "<div style=""background-color:#e0e0e0;border: solid 1px #000000;padding:10px;FONT-FAMILY: Verdana,Tahoma,Arial;font-size:10px;"">"
'	response.flush


	' APPENDING OR OVERWRITE
	If blnOverwrite = True Then
		' DELETE EXISTING FILE AND FOLDER DATABASE ROWS SO SCRIPT CAN REBUILD FROM SCRATCH
		response.write "<p><b>Clearing existing database rows for orgid(" & iorgid & ")...</b><br>"
		ClearDocuments iOrgID
		response.write "<b>Done clearing database rows.</b><br></p>"
		response.flush
		
		' ADD ROOT PATH
		AddFolder iOrgID, sVirtualPath  

'	Else

		' APPEND EXISTING TO EXISTING DATABASE ROWS
'		response.write "<P><b>This option has not been coded.  Please contact a developer.</b><br></p>"
'		response.flush
	End If


	' ENUMERATE FOLDERS AND FILES ADDING ENTRIES TO DATABASE
	response.write "<b>Adding Files and Folders...</b><br>"
'	response.flush


	' ENUMERATE FOLDERS AND FILES ADDING TO DATABASE
	SyncDocuments iOrgID, FSO.GetFolder( sPath )

	response.write "<b>Done Adding Files and Folders.</b><br />"
'	response.write "</div>"
	response.flush

End Sub 


'-----------------------------------------------------------------------------------
' void SYNCDOCUMENTS(FOLDER)
'-----------------------------------------------------------------------------------
Sub SyncDocuments( ByVal iOrgID, ByVal Folder )
	Dim Subfolder, sSubPath, sDocumentPath
   
	For Each Subfolder In Folder.SubFolders
		sSubPath = replace(Subfolder.Path,sPath,"")
		sSubPath = replace(sSubPath,"\","/")
		sSubPath = sVirtualPath & sSubPath
		response.write "<b>Adding Folder: </b>" & sSubPath & "<br>"
		AddFolder iOrgID, sSubPath 
		
		For Each File In Subfolder.Files 
			sDocumentPath = sSubPath & "/" & File.Name 
			response.write  "<b>Adding File: </b>" &sDocumentPath & "<BR>"
			response.flush
			AddDocument iOrgID, sDocumentPath
		Next 
        
		SyncDocuments Subfolder
    Next

End Sub


'-----------------------------------------------------------------------------------
' void AddDocument IORGID,SPATH
'-----------------------------------------------------------------------------------
Sub AddDocument( ByVal iOrgID, ByVal sPath )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "NewDocument"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iOrgID)
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
		.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, sPath)
		.Parameters.Append oCmd.CreateParameter("LinkURL", adVarChar, adParamInput, 300, null)
		.Execute
	End With

	Set oCmd = Nothing

End Sub


'-----------------------------------------------------------------------------------
' void AddFolder iOrgid, sPath
'-----------------------------------------------------------------------------------
Sub AddFolder( ByVal iorgid, ByVal sPath )
	Dim oCmd

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "NewFolder"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("OrgID", adInteger, adParamInput, 4, iorgid)
		.Parameters.Append oCmd.CreateParameter("CreatorID", adInteger, adParamInput, 4, Session("UserID"))
		.Parameters.Append oCmd.CreateParameter("FolderPath", adVarChar, adParamInput, 300, sPath)
		.Execute
	End With
    
	Set oCmd = Nothing

End Sub


'-----------------------------------------------------------------------------------
' void ClearDocuments iorgid
'-----------------------------------------------------------------------------------
Sub ClearDocuments( ByVal iorgid )
	Dim sSql

	sSql = "DELETE FROM DOCUMENTS WHERE ORGID = " & iorgid 
	RunSQL sSql 

	sSql = "DELETE FROM DOCUMENTFOLDERS WHERE ORGID = " & iorgid 
	RunSQL sSql 
	
End Sub


sub dtb_debug( ByVal p_value)
	sSqli = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "') "
 	RunSQL sSqli
end sub

%>
