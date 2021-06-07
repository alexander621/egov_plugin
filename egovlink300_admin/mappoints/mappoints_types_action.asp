<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%

lcl_feature    = request("f")
lcl_return_url = ""

if lcl_feature <> "" then
      lcl_return_url = "&f=" & lcl_feature
end if

if request("user_action") <> "" then
   if request("user_action") <> "DELETE" then
      updateMapPointType request("user_action"), request("mappoint_typeid"), request("description"), request("isActive"), _
                         request("mappointcolor"), request("feature_public"), request("feature_maintain"), _
                         request("feature_maintain_fields"), request("displayMap"), request("useAdvancedSearch"), _
                         request("orgid"), session("userid"), lcl_redirect_url

      for f = 1 to request("totalFields")
         if request.form("deleteField" & f) = "Y" then
            if request("mp_fieldid" & f) <> "" then
              'First: delete all "field values" associated to this Map-Point Type Field
               deleteMapPointValue "mp_fieldid", request("mp_fieldid" & f)

              'Second: delete Map-Point Type Field
               deleteMapPointTypeField request("mp_fieldid" & f)
            end if
         else
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
      deleteMapPointType request("mappoint_typeid")
   end if
else
   response.redirect "mappoints_types_list.asp" & replace(lcl_return_url,"&","?")
end if

'------------------------------------------------------------------------------
sub updateMapPointType(ByVal iAction, ByVal iMapPointTypeID, ByVal iDescription, ByVal iIsActive, ByVal iMapPointColor, _
                       ByVal iFeaturePublic, ByVal iFeatureMaintain, ByVal iFeatureMaintainFields, ByVal iDisplayMap, _
                       ByVal iUseAdvancedSearch, ByVal iOrgID, ByVal iUserID, ByRef lcl_redirect_url)

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
    sSQL = sSQL & "description = "             & sDescription           & ", "
    sSQL = sSQL & "isActive = "                & sIsActive              & ", "
    sSQL = sSQL & "mappointcolor = "           & sMapPointColor         & ", "
    sSQL = sSQL & "feature_public = "          & sFeaturePublic         & ", "
    sSQL = sSQL & "feature_maintain = "        & sFeatureMaintain       & ", "
    sSQL = sSQL & "feature_maintain_fields = " & sFeatureMaintainFields & ", "
    sSQL = sSQL & "displayMap = "              & sDisplayMap            & ", "
    sSQL = sSQL & "useAdvancedSearch = "       & sUseAdvancedSearch     & ", "
    sSQL = sSQL & "lastmodifiedbyid = "        & sUserID                & ", "
    sSQL = sSQL & "lastmodifiedbydate = '"     & dbsafe(ConvertDateTimetoTimeZone()) & "' "
    sSQL = sSQL & " WHERE mappoint_typeid = " & sMapPointTypeID

  		set oMPTypes = Server.CreateObject("ADODB.Recordset")
	  	oMPTypes.Open sSQL, Application("DSN"), 3, 1

    set oMPTypes = nothing

    lcl_redirect_url = "mappoints_types_maint.asp?mappoint_typeid=" & sMapPointTypeID & "&success=SU"

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
    sSQL = sSQL & "useAdvancedSearch "
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & iOrgID                 & ", "
    sSQL = sSQL & sDescription           & ", "
    sSQL = sSQL & sIsActive              & ", "
    sSQL = sSQL & sMapPointColor         & ", "
    sSQL = sSQL & sCreatedByID           & ", "
    sSQL = sSQL & sCreatedByDate         & ", "
    sSQL = sSQL & "NULL,NULL,"
    sSQL = sSQL & sFeaturePublic         & ", "
    sSQL = sSQL & sFeatureMaintain       & ", "
    sSQL = sSQL & sFeatureMaintainFields & ", "
    sSQL = sSQL & sDisplayMap            & ", "
    sSQL = sSQL & sUseAdvancedSearch
    sSQL = sSQL & ")"

 		'Get the MapPointTypeID
	  	sMapPointTypeID = RunIdentityInsert(sSQL)

   'Insert the default Map-Point Type fields
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Address",   "ADDRESS",   "0", "0", "1", "1", 1, "0"
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Latitude",  "LATITUDE",  "0", "0", "1", "1", 2, "0"
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Longitude", "LONGITUDE", "0", "0", "1", "1", 3, "0"
    'maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Status",    "STATUS",    "0", "0", "1", "1", 4, "0"

    lcl_redirect_url = "mappoints_types_maint.asp?success=SA"

    if iAction = "ADD" then
       lcl_redirect_url = lcl_redirect_url & "&mappoint_typeid=" & sMapPointTypeID
    end if

 end if

end sub

'------------------------------------------------------------------------------
sub deleteMapPointType(iMapPointTypeID)

  if iMapPointTypeID <> "" then
     sMapPointTypeID = CLng(iMapPointTypeID)
  else
     sMapPointTypeID = 0
  end if

 'BEGIN: Delete all of the Map-Point Values -----------------------------------
  sSQL = "DELETE FROM egov_mappoints_values "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPValues = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPValues.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPValues = nothing
 'END: Delete all of the Map-Point Values -------------------------------------

 'BEGIN: Delete the Map-Points ------------------------------------------------
  sSQL = "DELETE FROM egov_mappoints "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPoints = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPoints.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPoints = nothing
 'END: Delete the Map-Points --------------------------------------------------

 'BEGIN: Delete all of the Map-Point Types Fields for the Map-Point Type ------
  sSQL = "DELETE FROM egov_mappoints_types_fields "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPTFields = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPTFields.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPTFields = nothing
 'END: Delete all of the Map-Point Types Fields for the Map-Point Type --------

 'BEGIN: Delete the Map-Point Type --------------------------------------------
  sSQL = "DELETE FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE mappoint_typeid = " & sMapPointTypeID

	 set oDeleteMPType = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPType.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPType = nothing
 'END: Delete the Map-Point Type ----------------------------------------------

  response.redirect "mappoints_types_list.asp?success=SD"

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

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,"'","''")
  end if

  dbsafe = lcl_return

end function
%>