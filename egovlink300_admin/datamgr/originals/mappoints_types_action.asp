<!-- #include file="../includes/common.asp" //-->
<%
if request("user_action") <> "" then
   if request("user_action") <> "DELETE" then
      updateMapPoint request("user_action"),request("mappoint_typeid"), request("description"), request("mappointtype"), _
                     request("feature"), request("feature_maintain"), session("userid")
   else
      deleteMapPointType request("mappoint_typeid")
   end if
else
   response.redirect "mappoints_types_list.asp"
end if

'------------------------------------------------------------------------------
sub updateMapPoint(iAction, iMapPointTypeID, iDescription, iMapPointType, iFeature, iFeatureMaintain, iUserID)

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

 if iMapPointType = "" then
  		sMapPointType = "NULL"
 else
  		sMapPointType = "'" & dbsafe(iMapPointType) & "'"
 end if

 if iFeature = "" then
  		sFeature = "NULL"
 else
  		sFeature = "'" & dbsafe(iFeature) & "'"
 end if

 if iFeatureMaintain = "" then
  		sFeatureMaintain = "NULL"
 else
  		sFeatureMaintain = "'" & dbsafe(iFeatureMaintain) & "'"
 end if

 if iUserID <> "" then
    sUserID = CLng(iUserID)
 else
    sUserID = 0
 end if

'The mappointtype exists, so update it
 if iAction = "UPDATE" then
  		sSQL = "UPDATE egov_mappoints_types SET "
    sSQL = sSQL & "mappointtype = "        & sMapPointType    & ", "
    sSQL = sSQL & "description = "         & sDescription     & ", "
    sSQL = sSQL & "feature = "             & sFeature         & ", "
    sSQL = sSQL & "feature_maintain = "    & sFeatureMaintain & ", "
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
    sSQL = sSQL & "mappointtype, "
    sSQL = sSQL & "description, "
    sSQL = sSQL & "feature, "
    sSQL = sSQL & "feature_maintain, "
    sSQL = sSQL & "createdbyid, "
    sSQL = sSQL & "createdbydate, "
    sSQL = sSQL & "lastmodifiedbyid,"
    sSQL = sSQL & "lastmodifiedbydate"
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & sMapPointType    & ", "
    sSQL = sSQL & sDescription     & ", "
    sSQL = sSQL & sFeature         & ", "
    sSQL = sSQL & sFeatureMaintain & ", "
    sSQL = sSQL & sCreatedByID     & ", "
    sSQL = sSQL & sCreatedByDate   & ", "
    sSQL = sSQL & "NULL,NULL"
    sSQL = sSQL & ")"

 		'Get the BlogID
	  	sMapPointTypeID = RunIdentityInsert(sSQL)

    lcl_redirect_url = "mappoints_types_maint.asp?success=SA"

    if iAction = "ADD" then
       lcl_redirect_url = lcl_redirect_url & "&mappoint_typeid=" & sMapPointTypeID
    end if

 end if

 response.redirect lcl_redirect_url

end sub

'------------------------------------------------------------------------------
function RunIdentityInsert( sInsertStatement )
	 Dim sSQL, iReturnValue, oInsert

	 iReturnValue = 0

	'Insert new row into database and get rowid
 	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

 	set oInsert = Server.CreateObject("ADODB.Recordset")
	 oInsert.Open sSQL, Application("DSN"), 3, 3

 	iReturnValue = oInsert("ROWID")

 	oInsert.close
	 set oInsert = nothing

 	RunIdentityInsert = iReturnValue

end function

'------------------------------------------------------------------------------
sub deleteMapPointType(iMapPointTypeID)

  if iMapPointTypeID <> "" then
     sMapPointTypeID = CLng(iMapPointTypeID)
  else
     sMapPointTypeID = 0
  end if

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