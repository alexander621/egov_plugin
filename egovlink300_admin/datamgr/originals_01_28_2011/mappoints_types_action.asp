<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'for each x in request.form
'    response.write(x&" = """& Server.URLEncode(request.form(x)) &""" ("&len(request.form(x))&")") & "<br />"
'next

lcl_feature    = request("f")
lcl_isTemplate = request("t")
lcl_return_url = ""

if lcl_feature <> "" then
      lcl_return_url = "&f=" & lcl_feature
end if

if request("user_action") <> "" then
   if request("user_action") <> "DELETE" then

     'This maintains the MapPoint Type data (top of maintenance screen)
      updateMapPointType request("user_action"), request("mappoint_typeid"), request("description"), request("isActive"), _
                         request("mappointcolor"), request("feature_public"), request("feature_maintain"), _
                         request("feature_maintain_fields"), request("displayMap"), request("useAdvancedSearch"), _
                         request("latitude"), request("longitude"), request("mappoints_types_defaultzoomlevel"), _
                         request("isTemplate"), request("MPTemplateID"), request("orgid"), session("userid"), lcl_redirect_url

     'This maintains the custom fields data for the MapPoint Type (bottom of maintenance screen)
      for f = 1 to request("totalFields")
         if request.form("deleteField" & f) = "Y" then
            if request("mp_fieldid" & f) <> "" then
              'First: delete all "field values" associated to this Map-Point Type Field
               deleteMapPointValue "mp_fieldid", request("mp_fieldid" & f)

              'Second: delete Map-Point Type Field
               deleteMapPointTypeField request("mp_fieldid" & f)
            end if
         else
           'This inserts/updates the data for each custom field

'dtb_debug("[" & request("mp_fieldid" & f) & "] - [" & request("fieldname" & f) & "] - [" & request("fieldname5") & "]")

            maintainMapPointTypeField request("mp_fieldid" & f), request("mappoint_typeid"), request("orgid"), request("fieldname" & f), _
                                      request("fieldtype" & f), request("hasAddLinkButton" & f), request("isMultiLine" & f), _
                                      request("displayInResults" & f), request("displayInInfoPage" & f), request("resultsOrder" & f), _
                                      request("inPublicSearch" & f)
         end if

      next

     'Reorder the Map-Point Type fields
      reorderMapPointTypeFields request("mappoint_typeid")

      response.redirect lcl_redirect_url & lcl_return_url

   else
      deleteMapPointType request("mappoint_typeid"), lcl_isTemplate
   end if
else
   response.redirect "mappoints_types_list.asp" & replace(lcl_return_url,"&","?")
end if

'------------------------------------------------------------------------------
sub updateMapPointType(ByVal iAction, ByVal iMapPointTypeID, ByVal iDescription, ByVal iIsActive, ByVal iMapPointColor, _
                       ByVal iFeaturePublic, ByVal iFeatureMaintain, ByVal iFeatureMaintainFields, ByVal iDisplayMap, _
                       ByVal iUseAdvancedSearch, ByVal iLatitude, ByVal iLongitude, ByVal iMapPoints_DefaultZoomLevel, _
                       ByVal iIsTemplate, ByVal iMPTemplateID, ByVal iOrgID, ByVal iUserID, ByRef lcl_redirect_url)

 sOrgID = iOrgID

 if iMapPointTypeID <> "" then
    sMapPointTypeID = CLng(iMapPointTypeID)
 else
    sMapPointTypeID = 0
 end if

 if iDescription = "" then
  		sDescription = "NULL"
 else
  		sDescription = "'" & dbsafe(iDescription) & "'"
 end if

 if iIsActive = "Y" then
    sIsActive = 1
 else
    sIsActive = 0
 end if

 if iDisplayMap = "Y" then
    sDisplayMap = 1
 else
    sDisplayMap = 0
 end if

 if iUseAdvancedSearch = "Y" then
    sUseAdvancedSearch = 1
 else
    sUseAdvancedSearch = 0
 end if

 if iIsTemplate = "Y" then
    sIsTemplate = 1
    sOrgID      = 0  'Set the OrgID to Zero when this is a template
    sIsTemplate_url = "&t=Y"
 else
    sIsTemplate = 0
    sIsTemplate_url = ""
 end if

 if iLatitude <> "" then
    sLatitude = iLatitude
 else
    sLatitude = "0.00"
 end if

 if iLongitude <> "" then
    sLongitude = iLongitude
 else
    sLongitude = "0.00"
 end if

 if iMapPoints_DefaultZoomLevel <> "" then
    sMapPoints_DefaultZoomLevel = iMapPoints_DefaultZoomLevel
 else
    sMapPoints_DefaultZoomLevel = "13"
 end if

 if iUserID <> "" then
    sUserID = CLng(iUserID)
 else
    sUserID = 0
 end if

 if iMapPointColor <> "" then
    sMapPointColor = "'" & dbsafe(iMapPointColor) & "'"
 else
    sMapPointColor = "'green'"
 end if

 if iFeaturePublic <> "" then
    sFeaturePublic = "'" & dbsafe(iFeaturePublic) & "'"
 else
    sFeaturePublic = "NULL"
 end if

 if iFeatureMaintain <> "" then
    sFeatureMaintain = "'" & dbsafe(iFeatureMaintain) & "'"
 else
    sFeatureMaintain = "NULL"
 end if

 if iFeatureMaintainFields <> "" then
    sFeatureMaintainFields = "'" & dbsafe(iFeatureMaintainFields) & "'"
 else
    sFeatureMaintainFields = "NULL"
 end if

'The mappointtype exists, so update it
 if iAction = "UPDATE" then
  		sSQL = "UPDATE egov_mappoints_types SET "
    sSQL = sSQL & "description = "                & sDescription                & ", "
    sSQL = sSQL & "isActive = "                   & sIsActive                   & ", "
    sSQL = sSQL & "mappointcolor = "              & sMapPointColor              & ", "
    sSQL = sSQL & "feature_public = "             & sFeaturePublic              & ", "
    sSQL = sSQL & "feature_maintain = "           & sFeatureMaintain            & ", "
    sSQL = sSQL & "feature_maintain_fields = "    & sFeatureMaintainFields      & ", "
    sSQL = sSQL & "displayMap = "                 & sDisplayMap                 & ", "
    sSQL = sSQL & "useAdvancedSearch = "          & sUseAdvancedSearch          & ", "
    sSQL = sSQL & "latitude = "                   & sLatitude                   & ", "
    sSQL = sSQL & "longitude = "                  & sLongitude                  & ", "
    sSQL = sSQL & "mappoints_defaultzoomlevel = " & sMapPoints_DefaultZoomLevel & ", "
    sSQL = sSQL & "lastmodifiedbyid = "           & sUserID                     & ", "
    sSQL = sSQL & "lastmodifiedbydate = '"        & dbsafe(ConvertDateTimetoTimeZone()) & "', "
    sSQL = ssQL & "isTemplate = "                 & sIsTemplate
    sSQL = sSQL & " WHERE mappoint_typeid = " & sMapPointTypeID

  		set oMPTypes = Server.CreateObject("ADODB.Recordset")
	  	oMPTypes.Open sSQL, Application("DSN"), 3, 1

    set oMPTypes = nothing

    lcl_redirect_url = "mappoints_types_maint.asp?mappoint_typeid=" & sMapPointTypeID & "&success=SU" & sIsTemplate_url

'------------------------------------------------------------------------------
 else  'New MapPointType
'------------------------------------------------------------------------------
    sCreatedByID   = iUserID
    sCreatedByDate = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

 		'Insert the new Blog
  		sSQL = "INSERT INTO egov_mappoints_types ("
    sSQL = sSQL & "orgid, "
    sSQL = sSQL & "description, "
    sSQL = sSQL & "isActive, "
    sSQL = sSQL & "mappointcolor, "
    sSQL = sSQL & "createdbyid, "
    sSQL = sSQL & "createdbydate, "
    sSQL = sSQL & "lastmodifiedbyid, "
    sSQL = sSQL & "lastmodifiedbydate, "
    sSQL = sSQL & "feature_public, "
    sSQL = sSQL & "feature_maintain, "
    sSQL = sSQL & "feature_maintain_fields, "
    sSQL = sSQL & "displayMap, "
    sSQL = sSQL & "useAdvancedSearch, "
    sSQL = sSQL & "latitude, "
    sSQL = sSQL & "longitude, "
    sSQL = sSQL & "mappoints_defaultzoomlevel, "
    sSQL = sSQL & "isTemplate "
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & sOrgID                      & ", "
    sSQL = sSQL & sDescription                & ", "
    sSQL = sSQL & sIsActive                   & ", "
    sSQL = sSQL & sMapPointColor              & ", "
    sSQL = sSQL & sCreatedByID                & ", "
    sSQL = sSQL & sCreatedByDate              & ", "
    sSQL = sSQL & "NULL,NULL,"
    sSQL = sSQL & sFeaturePublic              & ", "
    sSQL = sSQL & sFeatureMaintain            & ", "
    sSQL = sSQL & sFeatureMaintainFields      & ", "
    sSQL = sSQL & sDisplayMap                 & ", "
    sSQL = sSQL & sUseAdvancedSearch          & ", "
    sSQL = sSQL & sLatitude                   & ", "
    sSQL = sSQL & sLongitude                  & ", "
    sSQL = sSQL & sMapPoints_DefaultZoomLevel & ", "
    sSQL = sSQL & sIsTemplate
    sSQL = sSQL & ")"

 		'Get the MapPointTypeID
	  	sMapPointTypeID = RunIdentityInsert(sSQL)

    if iIsTemplate <> "Y" then
      'Insert the default Map-Point Type fields
       'maintainMapPointTypeField "", sMapPointTypeID, sOrgID, "Address",   "ADDRESS",   "0", "0", "1", "1", 1, "0"
       'maintainMapPointTypeField "", sMapPointTypeID, sOrgID, "Latitude",  "LATITUDE",  "0", "0", "1", "1", 2, "0"
       'maintainMapPointTypeField "", sMapPointTypeID, sOrgID, "Longitude", "LONGITUDE", "0", "0", "1", "1", 3, "0"

      'If a "fields template" has been selected then get all of the fields from the template to add to the new MapPoint Type
       if iMPTemplateID <> "" then
          sSQL = "SELECT mp_fieldid, "
          sSQL = sSQL & " mappoint_typeid, "
          sSQL = sSQL & " orgid, "
          sSQL = sSQL & " fieldname, "
          sSQL = sSQL & " isnull(fieldtype,'') as fieldtype, "
          sSQL = sSQL & " hasAddLinkButton, "
          sSQL = sSQL & " isMultiLine, "
          sSQL = sSQL & " displayInResults, "
          sSQL = sSQL & " displayInInfoPage, "
          sSQL = sSQL & " resultsOrder, "
          sSQL = sSQL & " inPublicSearch "
          sSQL = sSQL & " FROM egov_mappoints_types_fields "
          sSQL = sSQL & " WHERE mappoint_typeid = " & iMPTemplateID
          sSQL = sSQL & " ORDER BY resultsOrder, mp_fieldid "

          set oGetMPTFields = Server.CreateObject("ADODB.Recordset")
          oGetMPTFields.Open sSQL, Application("DSN"), 3, 1

          if not oGetMPTFields.eof then
             do while not oGetMPTFields.eof

                lcl_template_fieldname         = oGetMPTFields("fieldname")
                lcl_template_fieldtype         = oGetMPTFields("fieldtype")
                lcl_template_hasAddLinkButton  = "0"
                lcl_template_isMultiLine       = "0"
                lcl_template_displayInResults  = "0"
                lcl_template_displayInInfoPage = "0"
                lcl_template_resultsOrder      = 0
                lcl_template_inPublicSearch    = "0"

                if oGetMPTFields("hasAddLinkButton") then
                   lcl_template_hasAddLinkButton = "1"
                end if

                if oGetMPTFields("isMultiLine") then
                   lcl_template_isMultiLine = "1"
                end if

                if oGetMPTFields("displayInResults") then
                   lcl_template_displayInResults = "1"
                end if

                if oGetMPTFields("displayInInfoPage") then
                   lcl_template_displayInInfoPage = "1"
                end if

                lcl_template_resultsOrder = oGetMPTFields("resultsOrder")

                if oGetMPTFields("inPublicSearch") then
                   lcl_template_inPublicSearch = "1"
                end if

                maintainMapPointTypeField "", sMapPointTypeID, sOrgID, lcl_template_fieldname, lcl_template_fieldtype, _
                                          lcl_template_hasAddLinkButton, lcl_template_isMultiLine, _
                                          lcl_template_displayInResults, lcl_template_displayInInfoPage, _
                                          lcl_template_results_order, lcl_template_inPublicSearch

                oGetMPTFields.movenext
             loop
          end if

          oGetMPTFields.close
          set oGetMPTFields = nothing
       end if
    end if

    lcl_redirect_url = "mappoints_types_maint.asp?success=SA" & sIsTemplate_url

    if iAction = "ADD" then
       lcl_redirect_url = lcl_redirect_url & "&mappoint_typeid=" & sMapPointTypeID
    end if

 end if

end sub

'------------------------------------------------------------------------------
sub deleteMapPointType(iMapPointTypeID, iIsTemplate)

  if iMapPointTypeID <> "" then
     sMapPointTypeID = CLng(iMapPointTypeID)
  else
     sMapPointTypeID = 0
  end if

 'Delete all of the Map-Point Values
  sSQL = "DELETE FROM egov_mappoints_values "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPValues = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPValues.Open sSQL, Application("DSN"), 3, 1

 'Delete the Map-Points
  sSQL = "DELETE FROM egov_mappoints "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPoints = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPoints.Open sSQL, Application("DSN"), 3, 1

 'Delete all of the Map-Point Types Fields for the Map-Point Type
  sSQL = "DELETE FROM egov_mappoints_types_fields "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPTFields = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPTFields.Open sSQL, Application("DSN"), 3, 1

 'Delete the Map-Point Type
  sSQL = "DELETE FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE mappoint_typeid = " & sMapPointTypeID

	 set oDeleteMPType = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPType.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPValues  = nothing
  set oDeleteMPoints   = nothing
  set oDeleteMPTFields = nothing
  set oDeleteMPType    = nothing

  lcl_isTemplate_url = ""

  if iIsTemplate = "Y" then
     lcl_isTemplate_url = "&t=Y"
  end if

  response.redirect "mappoints_types_list.asp?success=SD" & lcl_isTemplate_url

end sub

'------------------------------------------------------------------------------
sub reorderMapPointTypeFields(iMapPointTypeID)

  lcl_order = 0

  if iMapPointTypeID <> "" then
     sSQL = "SELECT mp_fieldid "
     sSQL = sSQL & " FROM egov_mappoints_types_fields "
     sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID
     sSQL = sSQL & " ORDER BY resultsOrder "

   	 set oReorderMPTFields = Server.CreateObject("ADODB.Recordset")
    	oReorderMPTFields.Open sSQL, Application("DSN"), 3, 1

     if not oReorderMPTFields.eof then
        do while not oReorderMPTFields.eof

           lcl_order = lcl_order + 1

           sSQL = "UPDATE egov_mappoints_types_fields SET resultsOrder = " & lcl_order
           sSQL = sSQL & " WHERE mp_fieldid = " & oReorderMPTFields("mp_fieldid")

   	       set oUpdateOrderBy = Server.CreateObject("ADODB.Recordset")
          	oUpdateOrderBy.Open sSQL, Application("DSN"), 3, 1

           set oUpdateOrderBy = nothing

           oReorderMPTFields.movenext
        loop
     end if

     oReorderMPTFields.close
     set oReorderMPTFields = nothing
  end if

end sub
%>