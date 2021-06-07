<%
'------------------------------------------------------------------------------
sub GetCityPoint(ByVal p_orgid, ByRef sLat, ByRef sLng, ByRef sZoom )
    sLat  = ""
    sLng  = ""
    sZoom = 13

   'Get the point to center the map
    sSQL = "SELECT latitude, longitude, mappoints_defaultzoomlevel "
    sSQL = sSQL & " FROM organizations "
    sSQL = sSQL & " WHERE orgid = " & p_orgid

    set oCityPoint = Server.CreateObject("ADODB.Recordset")
    oCityPoint.Open sSQL, Application("DSN"), 3, 1

    if not oCityPoint.eof then
       sLat = oCityPoint("latitude")
       sLng = oCityPoint("longitude")

       if oCityPoint("mappoints_defaultzoomlevel") <> "" or not isnull(oCityPoint("mappoints_defaultzoomlevel")) then
          sZoom = CLng(oCityPoint("mappoints_defaultzoomlevel"))
       end if

    end if

    oCityPoint.close
    set oCityPoint = nothing

end sub

'------------------------------------------------------------------------------
sub getMapPointsTypeInfo(ByVal iMapPointTypeID, ByVal p_orgid, ByVal iFeature, ByRef lcl_total_mptypes, _
                         ByRef lcl_mappoint_typeid, ByRef lcl_description, ByRef lcl_mappointcolor, _
                         ByRef lcl_displayMap, ByRef lcl_useAdvancedSearch)
  lcl_mappoint_typeid   = 0
  lcl_feature           = ""
  lcl_description       = ""
  lcl_mappointcolor     = "green"
  lcl_total_mptypes     = 0
  lcl_displayMap        = True
  lcl_useAdvancedSearch = False

  if iFeature <> "" then
     lcl_feature = iFeature
  end if

 'Get the Map-Point Type ID
  if iMapPointTypeID <> "" then
     lcl_mappoint_typeid = CLng(iMapPointTypeID)
  else
    'If a feature has been passed in and the Map-Point Type could not be found then it means that the 
    'feature has not yet been assigned to the Map-Point Type.  Therefore, we should not show the MapPoints.

     if lcl_feature = "" then
       'Check to see if org has only one Map-Point Type.
       'If "yes" then show the Map-Points for that Map-Point Type
       'If "no" then grab the first one in the list (ordered by description)
        lcl_mappointtypes = getMapPointTypes(p_orgid)

        if lcl_mappointtypes <> "" then
           'sSQL = "SELECT DISTINCT mappoint_typeid "
           'sSQL = sSQL & " FROM egov_mappoints_types "
           'sSQL = sSQL & " WHERE mappoint_typeid IN (" & lcl_mappointtypes & ") "

           sSQL = "SELECT distinct mpt.mappoint_typeid "
           sSQL = sSQL & " FROM egov_mappoints_types mpt, egov_mappoints mp "
           sSQL = sSQL & " WHERE mpt.mappoint_typeid = mp.mappoint_typeid "
           sSQL = sSQL & " AND mp.orgid = " & p_orgid
           sSQL = sSQL & " AND mpt.mappoint_typeid IN (" & lcl_mappointtypes & ") "
           sSQL = sSQL & " AND mpt.isActive = 1 "
           sSQL = sSQL & " AND mp.isActive = 1 "
           sSQL = sSQL & " AND mp.latitude is not null "
           sSQL = sSQL & " AND mp.latitude <> 0.00 "
           sSQL = sSQL & " AND mp.longitude is not null "
           sSQL = sSQL & " AND mp.longitude <> 0.00 "

          	set oGetDefaultMPTypeID = Server.CreateObject("ADODB.Recordset")
           oGetDefaultMPTypeID.Open sSQL, Application("DSN"), 3, 1

           if not oGetDefaultMPTypeID.eof then
              lcl_mappoint_typeid = oGetDefaultMPTypeID("mappoint_typeid")
           end if

           oGetDefaultMPTypeID.close
           set oGetDefaultMPTypeID = nothing
        end if
     end if
  end if

 'Get the Map-Point Type Info
  if lcl_mappoint_typeid <> "" then
     sSQL = "SELECT distinct mappoint_typeid, "
     sSQL = sSQL & " description, "
     sSQL = sSQL & " isnull(mappointcolor, 'green') as mappointcolor, "
     sSQL = sSQL & " displayMap, "
     sSQL = sSQL & " useAdvancedSearch "
     sSQL = sSQL & " FROM egov_mappoints_types "
     sSQL = sSQL & " WHERE mappoint_typeid = '" & lcl_mappoint_typeid & "'"

    	set oGetMapPointTypeInfo = Server.CreateObject("ADODB.Recordset")
     oGetMapPointTypeInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetMapPointTypeInfo.eof then
        lcl_mappoint_typeid   = oGetMapPointTypeInfo("mappoint_typeid")
        lcl_description       = oGetMapPointTypeInfo("description")
        lcl_mappointcolor     = oGetMapPointTypeInfo("mappointcolor")
        lcl_displayMap        = oGetMapPointTypeInfo("displayMap")
        lcl_useAdvancedSearch = oGetMapPointTypeInfo("useAdvancedSearch")
     end if

     oGetMapPointTypeInfo.close
     set oGetMapPointTypeInfo = nothing

    'Find the total available, active, Map-Point Types for the org
     lcl_total_mptypes = getTotalMapPointTypes(p_orgid)
  end if

end sub

'------------------------------------------------------------------------------
function getMapPointTypes(p_orgid)

  lcl_return = ""

  if p_orgid <> "" then
     sSQL = "SELECT distinct mpt.mappoint_typeid "
     sSQL = sSQL & " FROM egov_mappoints_types mpt, egov_mappoints mp "
     sSQL = sSQL & " WHERE mpt.mappoint_typeid = mp.mappoint_typeid "
     sSQL = sSQL & " AND mp.orgid = " & p_orgid
     sSQL = sSQL & " AND mpt.isActive = 1 "
     sSQL = sSQL & " AND mp.isActive = 1 "
     sSQL = sSQL & " AND mp.latitude is not null "
     sSQL = sSQL & " AND mp.latitude <> 0.00 "
     sSQL = sSQL & " AND mp.longitude is not null "
     sSQL = sSQL & " AND mp.longitude <> 0.00 "

     'sSQL = "SELECT mappoint_typeid "
     'sSQL = sSQL & " FROM egov_mappoints_types "
     'sSQL = sSQL & " WHERE orgid = " & p_orgid
     'sSQL = sSQL & " ORDER BY description "

     set oGetMPTypes = Server.CreateObject("ADODB.Recordset")
     oGetMPTypes.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPTypes.eof then
        do while not oGetMPTypes.eof

           if lcl_return <> "" then
              lcl_return = lcl_return & "," & oGetMPTypes("mappoint_typeid")
           else
              lcl_return = oGetMPTypes("mappoint_typeid")
           end if

           oGetMPTypes.movenext
        loop
     end if

     oGetMPTypes.close
     set oGetMPTypes = nothing

  end if

  getMapPointTypes = lcl_return

end function

'------------------------------------------------------------------------------
function getTotalMapPointTypes(p_orgid)
  lcl_return = 0

  if p_orgid <> "" then
     'sSQL = "SELECT count(mappoint_typeid) as total_mptypes "
     'sSQL = sSQL & " FROM egov_mappoints_types "
     'sSQL = sSQL & " WHERE orgid = " & p_orgid
     'sSQL = sSQL & " AND isActive = 1 "

     sSQL = "SELECT count(distinct mpt.mappoint_typeid) as total_mptypes "
     sSQL = sSQL & " FROM egov_mappoints_types mpt, egov_mappoints mp "
     sSQL = sSQL & " WHERE mpt.mappoint_typeid = mp.mappoint_typeid "
     sSQL = sSQL & " AND mp.orgid = " & p_orgid
     sSQL = sSQL & " AND mpt.isActive = 1 "
     sSQL = sSQL & " AND mp.isActive = 1 "
     sSQL = sSQL & " AND mp.latitude is not null "
     sSQL = sSQL & " AND mp.latitude <> 0.00 "
     sSQL = sSQL & " AND mp.longitude is not null "
     sSQL = sSQL & " AND mp.longitude <> 0.00 "

     set oGetTotalMPTypes = Server.CreateObject("ADODB.Recordset")
     oGetTotalMPTypes.Open sSQL, Application("DSN"), 3, 1

     if not oGetTotalMPTypes.eof then
        lcl_return = oGetTotalMPTypes("total_mptypes")
     end if

     oGetTotalMPTypes.close
     set oGetTotalMPTypes = nothing

  end if

  getTotalMapPointTypes = lcl_return

end function

'------------------------------------------------------------------------------
sub displayMapPointTypes(p_orgid, iMapPointTypeID)

  if iMapPointTypeID <> "" then
     lcl_mappointtypeid = CLng(iMapPointTypeID)
  else
     lcl_mappointtypeid = 0
  end if

 'sSQL = "SELECT mp.mappoint_typeid, "
 'sSQL = sSQL & " mpt.description "
 'sSQL = sSQL & " FROM egov_mappoints mp "
 'sSQL = sSQL &      " INNER JOIN egov_mappoints_types mpt ON mp.mappoint_typeid = mpt.mappoint_typeid "
 'sSQL = sSQL &      " LEFT OUTER JOIN egov_mappoints_values mpv ON mp.mappointid = mpv.mappointid "
 'sSQL = sSQL &                  " AND mpv.displayInResults = 1 "
 'sSQL = sSQL & " WHERE mp.orgid = " & p_orgid
 'sSQL = sSQL & " AND mpt.isActive = 1 "
 'sSQL = sSQL & " AND mp.latitude is not null "
 'sSQL = sSQL & " AND mp.latitude <> 0.00 "
 'sSQL = sSQL & " AND mp.longitude is not null "
 'sSQL = sSQL & " AND mp.longitude <> 0.00 "
 'sSQL = sSQL & " ORDER BY mpt.description "

 sSQL = "SELECT distinct mpt.mappoint_typeid, "
 sSQL = sSQL & " mpt.description "
 sSQL = sSQL & " FROM egov_mappoints_types mpt, egov_mappoints mp  "
 sSQL = sSQL & " WHERE mpt.mappoint_typeid = mp.mappoint_typeid "
 sSQL = sSQL & " AND mp.orgid = " & p_orgid
 sSQL = sSQL & " AND mpt.isActive = 1 "
 sSQL = sSQL & " AND mp.isActive = 1 "
 sSQL = sSQL & " AND mp.latitude is not null "
 sSQL = sSQL & " AND mp.latitude <> 0.00 "
 sSQL = sSQL & " AND mp.longitude is not null "
 sSQL = sSQL & " AND mp.longitude <> 0.00 "
 sSQL = sSQL & " ORDER BY mpt.description "

  set oDisplayMPTypes = Server.CreateObject("ADODB.Recordset")
  oDisplayMPTypes.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayMPTypes.eof then
     do while not oDisplayMPTypes.eof

        if oDisplayMPTypes("mappoint_typeid") = lcl_mappointtypeid then
           lcl_selected_mappointtype = " selected=""selected"""
        else
           lcl_selected_mappointtype = ""
        end if

        response.write "  <option value=""" & oDisplayMPTypes("mappoint_typeid") & """" & lcl_selected_mappointtype & ">" & oDisplayMPTypes("description") & "</option>" & vbcrlf

        oDisplayMPTypes.movenext
     loop
  end if

end sub

'------------------------------------------------------------------------------
function getMapPointTypeByFeature(p_orgid, p_feature)

  lcl_return = ""

  if p_feature <> "" then
     sSQL = "SELECT mappoint_typeid "
     sSQL = sSQL & " FROM egov_mappoints_types "
     sSQL = sSQL & " WHERE UPPER(feature_public) = '" & UCASE(p_feature) & "' "
     sSQL = sSQL & " AND orgid = " & p_orgid

     set oGetMPTID = Server.CreateObject("ADODB.Recordset")
     oGetMPTID.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPTID.eof then
        lcl_return = oGetMPTID("mappoint_typeid")
     end if

     oGetMPTID.close
     set oGetMPTID = nothing

  end if

  getMapPointTypeByFeature = lcl_return

end function

'------------------------------------------------------------------------------
function formatFieldValue(iFieldValue)
  lcl_return = ""

  if iFieldValue <> "" then
     lcl_return = iFieldValue
     lcl_return = replace(lcl_return,chr(10),"")
     lcl_return = replace(lcl_return,chr(13),"<br />")
     lcl_return = replace(lcl_return,"'","\'")
  end if

  formatFieldValue = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(iValue)
  lcl_return = ""

  if iValue <> "" then
     lcl_return = iValue
     lcl_return = replace(lcl_return,"'","''")
  end if

  dbsafe = lcl_return

end function
%>
