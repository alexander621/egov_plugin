<!-- #include file="../includes/common.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
if request("user_action") <> "" then
   if request("user_action") <> "DELETE" then
      updateMapPointType request("user_action"), request("mappoint_typeid"), request("description"), request("isActive"), request("orgid"), session("userid"), lcl_redirect_url

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
                                      request("fieldtype" & f), request("displayInResults" & f), request("resultsOrder" & f)
         end if

      next

        'Update all of the Map-Point Values with the current Map-Point Type data.
         updateMapPointValues request("mappoint_typeid")

      response.redirect lcl_redirect_url

   else
      deleteMapPointType request("mappoint_typeid")
   end if
else
   response.redirect "mappoints_types_list.asp"
end if

'------------------------------------------------------------------------------
sub updateMapPointType(ByVal iAction, ByVal iMapPointTypeID, ByVal iDescription, ByVal iIsActive, ByVal iOrgID, ByVal iUserID, ByRef lcl_redirect_url)

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
    sIsInActive = 0
 else
    sIsInActive = 1
 end if

 if iUserID <> "" then
    sUserID = CLng(iUserID)
 else
    sUserID = 0
 end if

'The mappointtype exists, so update it
 if iAction = "UPDATE" then
  		sSQL = "UPDATE egov_mappoints_types SET "
    sSQL = sSQL & "description = "         & sDescription     & ", "
    sSQL = sSQL & "isInactive = "          & sIsInActive      & ", "
    sSQL = sSQL & "lastmodifiedbyid = "    & sUserID          & ", "
    sSQL = sSQL & "lastmodifiedbydate = '" & dbsafe(ConvertDateTimetoTimeZone()) & "' "
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
    sSQL = sSQL & "isInactive, "
    sSQL = sSQL & "createdbyid, "
    sSQL = sSQL & "createdbydate, "
    sSQL = sSQL & "lastmodifiedbyid,"
    sSQL = sSQL & "lastmodifiedbydate"
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & iOrgID           & ", "
    sSQL = sSQL & sDescription     & ", "
    sSQL = sSQL & sIsInActive      & ", "
    sSQL = sSQL & sCreatedByID     & ", "
    sSQL = sSQL & sCreatedByDate   & ", "
    sSQL = sSQL & "NULL,NULL"
    sSQL = sSQL & ")"

 		'Get the MapPointTypeID
	  	sMapPointTypeID = RunIdentityInsert(sSQL)

   'Insert the default Map-Point Type fields
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Address",   "ADDRESS",   "1", 1
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Latitude",  "LATITUDE",  "1", 2
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Longitude", "LONGITUDE", "1", 3
    maintainMapPointTypeField "", sMapPointTypeID, iOrgID, "Status",    "STATUS",    "1", 4

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

 'Delete all of the Map-Point Types Fields for the Map-Point Type
  sSQL = "DELETE FROM egov_mappoints_types_fields "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

	 set oDeleteMPTFields = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPTFields.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPTFields = nothing

 'Delete the Map-Point Type
  sSQL = "DELETE FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE mappoint_typeid = " & sMapPointTypeID

	 set oDeleteMPType = Server.CreateObject("ADODB.Recordset")
 	oDeleteMPType.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMPType = nothing

  response.redirect "mappoints_types_list.asp?success=SD"

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