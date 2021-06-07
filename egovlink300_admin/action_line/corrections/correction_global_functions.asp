<%
'------------------------------------------------------------------------------
'Sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sPrefix, ByRef sAddress, ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, ByRef sZip, ByRef sLatitude, ByRef sLongitude, ByRef sCounty, ByRef sParcelID )
sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sPrefix, ByRef sAddress, ByRef sSuffix, ByRef sDirection, _
                    ByRef sLatitude, ByRef sLongitude, ByRef sCity, ByRef sState, ByRef sZip, ByRef sCounty, ByRef sParcelID, _
                    ByRef sListedOwner, ByRef sLegalDescription, ByRef sResidentType, ByRef sRegisteredUserID, ByRef sValidStreet )

 sValidStreet = "N"

	sSQL = "SELECT residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude, residentcity, residentstate, residentzip, "
 sSQL = sSQL & " county, parcelidnumber, listedowner, legaldescription, residenttype, registereduserid "
 sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE residentaddressid = " & sResidentAddressId

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 3, 1
	
	if not oAddress.eof then
  		sNumber           = trim(oAddress("residentstreetnumber"))
    sPrefix           = oAddress("residentstreetprefix")
  		sAddress          = oAddress("residentstreetname")
    sSuffix           = oAddress("streetsuffix")
    sDirection        = oAddress("streetdirection")
		  sLatitude         = oAddress("latitude")
  		sLongitude        = oAddress("longitude")
    sCity             = oAddress("residentcity")
    sState            = oAddress("residentstate")
    sZip              = oAddress("residentzip")
    sCounty           = oAddress("county")
    sParcelID         = oAddress("parcelidnumber")
    sListedOwner      = oAddress("listedowner")
    sLegalDescription = oAddress("legaldescription")
    sResidentType     = oAddress("residenttype")
    sRegisteredUserID = oAddress("registereduserid")
    sValidStreet      = "Y"
	end if

	oAddress.close
	set oAddress = nothing

end sub

'------------------------------------------------------------------------------
sub GetAddressInfoLarge( ByVal sStreetNumber, ByVal sStreetName, ByRef sNumber, ByRef sPrefix, ByRef sAddress, ByRef sSuffix, _
                         ByRef sDirection, ByRef sLatitude, ByRef sLongitude, ByRef sCity, ByRef sState, ByRef sZip, ByRef sCounty, _
                         ByRef sParcelID, ByRef sListedOwner, ByRef sLegalDescription, ByRef sResidentType, ByRef sRegisteredUserID, _
                         ByRef sValidStreet )

 sValidStreet = "N"

	sSQL = "SELECT residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude, residentcity, residentstate, residentzip, "
 sSQL = sSQL & " county, parcelidnumber, listedowner, legaldescription, residenttype, registereduserid "
 sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND UPPER(residentstreetnumber) = UPPER('" & dbsafe(sStreetNumber) & "') "
 sSQL = sSQL & " AND (residentstreetname = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sStreetName) & "' "
 sSQL = sSQL & " )"

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 3, 1
	
	if not oAddress.eof then
  		sNumber           = trim(oAddress("residentstreetnumber"))
    sPrefix           = oAddress("residentstreetprefix")
  		sAddress          = oAddress("residentstreetname")
    sSuffix           = oAddress("streetsuffix")
    sDirection        = oAddress("streetdirection")
		  sLatitude         = oAddress("latitude")
  		sLongitude        = oAddress("longitude")
    sCity             = oAddress("residentcity")
    sState            = oAddress("residentstate")
    sZip              = oAddress("residentzip")
    sCounty           = oAddress("county")
    sParcelID         = oAddress("parcelidnumber")
    sListedOwner      = oAddress("listedowner")
    sLegalDescription = oAddress("legaldescription")
    sResidentType     = oAddress("residenttype")
    sRegisteredUserID = oAddress("registereduserid")
    sValidStreet      = "Y"
	end if

	oAddress.close
	set oAddress = nothing

end sub

'------------------------------------------------------------------------------
function checkForIssueLocationOnForm(p_requestid)
  lcl_return = False

  if p_requestid <> "" then
    'Determine if the form is diplaying the issue/problem location section
     sSQLr = "SELECT f.action_form_display_issue "
     sSQLr = sSQLr & " FROM egov_action_request_forms f, egov_actionline_requests r "
     sSQLr = sSQLr & " WHERE f.action_form_id = r.category_id "
     sSQLr = sSQLr & " AND r.action_autoid = " & p_requestid

     set oForm = Server.CreateObject("ADODB.Recordset")
     oForm.Open sSQLr, Application("DSN"), 3, 1

     if not oForm.eof then
        lcl_return = oForm("action_form_display_issue")
     end if
  end if

  oForm.close
  set oForm = nothing

  checkForIssueLocationOnForm = lcl_return

end function

'------------------------------------------------------------------------------
function checkForHideIssueLocAddInfo(p_requestid)
  lcl_return = False

  if p_requestid <> "" then
    'Determine if the form is diplaying the issue/problem location section
     sSQLr = "SELECT f.hideIssueLocAddInfo "
     sSQLr = sSQLr & " FROM egov_action_request_forms f, egov_actionline_requests r "
     sSQLr = sSQLr & " WHERE f.action_form_id = r.category_id "
     sSQLr = sSQLr & " AND r.action_autoid = " & p_requestid

     set oCheckAddInfo = Server.CreateObject("ADODB.Recordset")
     oCheckAddInfo.Open sSQLr, Application("DSN"), 3, 1

     if not oCheckAddInfo.eof then
        lcl_return = oCheckAddInfo("hideIssueLocAddInfo")
     end if
  end if

  oCheckAddInfo.close
  set oCheckAddInfo = nothing

  checkForHideIssueLocAddInfo = lcl_return

end function
%>