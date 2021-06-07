<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
  if request("user_action") = "" then
     response.redirect "mappoints_list.asp"
  end if

  if request("f") <> "" then
     lcl_feature = request("f")
  else
     lcl_feature = ""
  end if

 'Check for org features
  lcl_orghasfeature_large_address_list = orghasfeature("large address list")

 'Setup variables
  lcl_useraction      = ""
  lcl_mappointid      = 0
  lcl_mappoint_typeid = 0
  lcl_mappointcolor   = ""
  lcl_orgid           = request("orgid")
  lcl_street_number   = request("residentstreetnumber")
  lcl_street_address  = request("streetaddress")
  sNumber             = ""
  sPrefix             = ""
  sAddress            = ""
  sSuffix             = ""
  sDirection          = ""
  sValidStreet        = request("validstreet")
  sCity               = request("city")
  sState              = request("state")
  sZip                = request("zip")
  sLatitude           = 0.00
  sLongitude          = 0.00
  'sCounty             = ""
  'sParcelID           = ""
  'sListedOwner        = ""
  'sLegalDescription   = ""
  'sResidentType       = ""
  'sRegisteredUserID   = 0
  sSortStreetName     = ""
  lcl_isActive        = 1
  'sStatusID           = 0
  lcl_userid          = session("userid")
  lcl_current_date    = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  lcl_redirect_url    = "mappoints_list.asp"

  'oSave("streetunit")       = request("streetunit")
  'oSave("county")           = request("county")
  'oSave("parcelidnumber")   = request("parcelidnumber")
  'oSave("listedowner")      = request("listedowner")
  'oSave("residenttype")     = request("residenttype")
  'oSave("legaldescription") = request("legaldescription")

  'if request("registereduserid") = "" then
  '   oSave("registereduserid") = 0
  'else
  '   oSave("registereduserid") = request("registereduserid")
  'end if

  if request("user_action") <> "" then
     lcl_useraction = UCASE(request("user_action"))
  end if

  if request("mappointid") <> "" then
     lcl_mappointid = request("mappointid")
  end if

  if request("mappoint_typeid") <> "" then
     lcl_mappoint_typeid = request("mappoint_typeid")
  end if

  if request("mappointcolor") <> "" then
     lcl_mappointcolor = request("mappointcolor")
  else
     lcl_mappointcolor = getMapPointTypePointColor(lcl_mappoint_typeid)
  end if

 'Retrieve the search options
  lcl_sc_mappoint_typeid = ""

  if request("sc_mappoint_typeid") <> "" then
     lcl_sc_mappoint_typeid = request("sc_mappoint_typeid")
  end if

 'Build return parameters
  lcl_url_parameters = ""
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_mappoint_typeid", lcl_sc_mappoint_typeid)

  if lcl_feature <> "mappoints_maint" then
     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
  end if

 'Execute the user's action
  if lcl_useraction = "DELETE" then
    'First: delete all "field values" associated to this Map-Point
     deleteMapPointValue "mappointid", lcl_mappointid

    'Second: delete the Map-Point
     sSQL = "DELETE FROM egov_mappoints WHERE mappointid = " & lcl_mappointid

   		set oDeleteMapPoint = Server.CreateObject("ADODB.Recordset")
    	oDeleteMapPoint.Open sSQL, Application("DSN"), 3, 1

     set oDeleteMapPoint = nothing

     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SD")
     lcl_redirect_url   = "mappoints_list.asp" & lcl_url_parameters

  else
     if request("isActive") = "Y" then
        lcl_isActive = 1
     else
        lcl_isActive = 0
     end if

    'BEGIN: Status ------------------------------------------------------------
     'if request("statusid") <> "" then
     '   sStatusID = request("statusid")
     'else
     '   sStatusID = 0
     'end if

     'lcl_mpv_statusname = getMapPointStatusName(sStatusID)
     'lcl_mpv_statusname = ""
    'END: Status --------------------------------------------------------------

    'BEGIN: Check to see if the address is valid or custom --------------------
    	if trim(request("ques_issue2")) <> "" then
        sAddress        = request("ques_issue2")
        lcl_mpv_address = sAddress
     else
        getAddressInfoNew lcl_orghasfeature_large_address_list, lcl_orgid, lcl_street_number, lcl_street_address, _
                          sNumber, sPrefix, sAddress, sSuffix, sDirection, sLatitude, sLongitude, sCity, sState, sZip, _
                          sCounty, sParcelID, sListedOwner, sLegalDescription, sResidentType, sRegisteredUserID, sValidStreet

        lcl_mpv_address = buildStreetAddress(sNumber, sPrefix, sAddress, sSuffix, sDirection)

     end if

    'COMPARE VALUES - IF CHANGED UPDATE AND LOG
     'lcl_original_address     = buildStreetAddress(oSave("streetnumber"), oSave("streetprefix"), oSave("streetaddress"), oSave("streetsuffix"), oSave("streetdirection"))
     'lcl_sAddress_street_name = buildStreetAddress(sNumber, sPrefix, sAddress, sSuffix, sDirection)

    'Re-build the SortStreetName
     sSortStreetName = trim(sAddress)

     if trim(sSuffix) <> "" then
        if sSortStreetName <> "" then
           sSortStreetName = sSortStreetName & " " & sSuffix
        else
           sSortStreetName = sSuffix
        end if
     end if

     if trim(sDirection) <> "" then
        if sSortStreetName <> "" then
           sSortStreetName = sSortStreetName & " " & sDirection
        else
           sSortStreetName = sDirection
        end if
     end if

     if trim(sPrefix) <> "" then
        if sSortStreetName <> "" then
           sSortStreetName = sSortStreetName & " " & sPrefix
        else
           sSortStreetName = sPrefix
        end if
     end if

    'BEGIN: Latitude ----------------------------------------------------------
     if request("latitude") <> "" then
        sLatitude = request("latitude")
     else
        sLatitude = 0.00
     end if
    'END: Latitude ------------------------------------------------------------

    'BEGIN: Longitude ---------------------------------------------------------
     if request("longitude") <> "" then
        sLongitude = request("longitude")
     else
        sLongitude = 0.00
     end if
    'END: Longitude -----------------------------------------------------------


    'BEGIN: Format the columns for the table ----------------------------------
     'sLatitude       = formatFieldforInsertUpdate(sLatitude)
     'sLongitude      = formatFieldforInsertUpdate(sLongitude)
     sSortStreetName = formatFieldforInsertUpdate(sSortStreetName)
     sNumber         = formatFieldforInsertUpdate(sNumber)
     sPrefix         = formatFieldforInsertUpdate(sPrefix)
     sAddress        = formatFieldforInsertUpdate(sAddress)
     sSuffix         = formatFieldforInsertUpdate(sSuffix)
     sDirection      = formatFieldforInsertUpdate(sDirection)
     sValidStreet    = formatFieldforInsertUpdate(sValidStreet)
     sCity           = formatFieldforInsertUpdate(sCity)
     sState          = formatFieldforInsertUpdate(sState)
     sZip            = formatFieldforInsertUpdate(sZip)
     sMapPointColor  = formatFieldforInsertUpdate(lcl_mappointcolor)
    'END: Format the columns for the table ------------------------------------

     if lcl_useraction = "UPDATE" then

      		sSQL = "UPDATE egov_mappoints SET "
        sSQL = sSQL & "mappoint_typeid = "    & lcl_mappoint_typeid & ", "
        sSQL = sSQL & "lastmodifiedbyid = "   & lcl_userid          & ", "
        sSQL = sSQL & "lastmodifiedbydate = " & lcl_current_date    & ", "
        sSQL = sSQL & "isActive = "           & lcl_isActive        & ", "
        sSQL = sSQL & "streetnumber = "       & sNumber             & ", "
        sSQL = sSQL & "streetprefix = "       & sPrefix             & ", "
        sSQL = sSQL & "streetaddress = "      & sAddress            & ", "
        sSQL = sSQL & "streetsuffix = "       & sSuffix             & ", "
        sSQL = sSQL & "streetdirection = "    & sDirection          & ", "
        sSQL = sSQL & "sortstreetname = "     & sSortStreetName     & ", "
        sSQL = sSQL & "city = "               & sCity               & ", "
        sSQL = sSQL & "state = "              & sState              & ", "
        sSQL = sSQL & "zip = "                & sZip                & ", "
        sSQL = sSQL & "validstreet = "        & sValidStreet        & ", "
        'sSQL = sSQL & "statusid = "           & sStatusID           & ", "
        sSQL = sSQL & "latitude = "           & sLatitude           & ", "
        sSQL = sSQL & "longitude = "          & sLongitude          & ", "
        sSQL = sSQL & "mappointcolor = "      & sMapPointColor
        sSQL = sSQL & " WHERE mappointid = " & lcl_mappointid

      		set oCreateMapPoint = Server.CreateObject("ADODB.Recordset")
	      	oCreateMapPoint.Open sSQL, Application("DSN"), 3, 1

        set oCreateMapPoint = nothing

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "mappointid", lcl_mappointid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SU")
        lcl_redirect_url   = "mappoints_maint.asp" & lcl_url_parameters

    '---------------------------------------------------------------------------
     else  'New MapPoint
    '---------------------------------------------------------------------------
        sCreatedByID   = lcl_userid
        sCreatedByDate = lcl_current_date

     		'Insert the new Map-Point
   	   	sSQL = "INSERT INTO egov_mappoints ("
        sSQL = sSQL & "mappoint_typeid, "
        sSQL = sSQL & "orgid, "
        sSQL = sSQL & "createdbyid, "
        sSQL = sSQL & "createdbydate, "
        sSQL = sSQL & "lastmodifiedbyid, "
        sSQL = sSQL & "lastmodifiedbydate, "
        sSQL = sSQL & "isActive, "
        sSQL = sSQL & "streetnumber, "
        sSQL = sSQL & "streetprefix, "
        sSQL = sSQL & "streetaddress, "
        sSQL = sSQL & "streetsuffix, "
        sSQL = sSQL & "streetdirection, "
        sSQL = sSQL & "sortstreetname, "
        sSQL = sSQL & "city, "
        sSQL = sSQL & "state, "
        sSQL = sSQL & "zip, "
        sSQL = sSQL & "validstreet, "
        'sSQL = sSQL & "statusid, "
        sSQL = sSQL & "latitude, "
        sSQL = sSQL & "longitude, "
        sSQL = sSQL & "mappointcolor"
        sSQL = sSQL & ") VALUES ("
        sSQL = sSQL & lcl_mappoint_typeid & ", "
        sSQL = sSQL & lcl_orgid           & ", "
        sSQL = sSQL & sCreatedByID        & ", "
        sSQL = sSQL & sCreatedByDate      & ", "
        sSQL = sSQL & "NULL,NULL"         & ", "
        sSQL = sSQL & lcl_isActive        & ", "
        sSQL = sSQL & sNumber             & ", "
        sSQL = sSQL & sPrefix             & ", "
        sSQL = sSQL & sAddress            & ", "
        sSQL = sSQL & sSuffix             & ", "
        sSQL = sSQL & sDirection          & ", "
        sSQL = sSQL & sSortStreetName     & ", "
        sSQL = sSQL & sCity               & ", "
        sSQL = sSQL & sState              & ", "
        sSQL = sSQL & sZip                & ", "
        sSQL = sSQL & sValidStreet        & ", "
        'sSQL = sSQL & sStatusID           & ", "
        sSQL = sSQL & sLatitude           & ", "
        sSQL = sSQL & sLongitude          & ", "
        sSQL = sSQL & sMapPointColor
        sSQL = sSQL & ")"

     		'Get the MapPointID
    	  	lcl_mappointid = RunIdentityInsert(sSQL)

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "mappointid", lcl_mappointid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SA")
        'lcl_redirect_url   = "mappoints_maint.asp?success=SA"
        lcl_redirect_url   = "mappoints_maint.asp" & lcl_url_parameters

        'if lcl_useraction = "ADD" then
        '   lcl_redirect_url = lcl_redirect_url & "&mappointid=" & lcl_mappointid
        'end if
     end if

    'BEGIN: Map-Point Field Values --------------------------------------------
     if request("totalFields") <> "" then
        lcl_totalfields = request("totalFields")
     else
        lcl_totalfields = 0
     end if

    'Clear all existing values for the MapPointID
     if lcl_totalfields > 0 then

       'BEGIN: Delete Map-Point Values ----------------------------------------
        sSQL = "DELETE FROM egov_mappoints_values "
        sSQL = sSQL & " WHERE mappointid = " & lcl_mappointid

      		set oDeleteMPValues = Server.CreateObject("ADODB.Recordset")
	      	oDeleteMPValues.Open sSQL, Application("DSN"), 3, 1

        set oDeleteMPValues = nothing
       'END: Delete Map-Point Values ------------------------------------------

       'BEGIN: Insert Map-Point Values ----------------------------------------
        for v = 1 to lcl_totalfields
         'if request.form("deleteField" & f) = "Y" then
         '   if request("mp_fieldid" & f) <> "" then
         '      deleteMapPointTypeField request("mp_fieldid" & f)
         '   end if
         'else

           'Determine if we pull the value from the screen or if we generate it.
            if request("fieldtype" & v) = "ADDRESS" then
               lcl_fieldvalue = lcl_mpv_address
            'elseif request("fieldtype" & v) = "STATUS" then
            '   lcl_fieldvalue = lcl_mpv_statusname
            elseif request("fieldtype" & v) = "LATITUDE" then
               lcl_fieldvalue = sLatitude
            elseif request("fieldtype" & v) = "LONGITUDE" then
               lcl_fieldvalue = sLongitude
            else
               lcl_fieldvalue = request("mp_fieldvalue" & v)
            end if

            'maintainMapPointValues lcl_orgid, lcl_mappoint_typeid, lcl_mappointid, request("mp_fieldid" & v), "", request("fieldtype" & v), _
            '                       request("fieldname" & v), lcl_fieldvalue, request("displayInResults" & v), request("resultsOrder" & v)
            maintainMapPointValues lcl_orgid, lcl_mappoint_typeid, lcl_mappointid, request("mp_fieldid" & v), "", request("fieldtype" & v), lcl_fieldvalue
         'end if
        next
       'END: Insert Map-Point Values ------------------------------------------

     end if
    'END: Map-Point Field Values ----------------------------------------------

  end if

  response.redirect lcl_redirect_url
%>