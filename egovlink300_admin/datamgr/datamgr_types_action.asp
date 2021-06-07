<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
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

     'This maintains the DM Type data (top of maintenance screen)
      updateDMType request("user_action"), _
                   request("dm_typeid"), _
                   request("description"), _
                   request("isActive"), _
                   request("enableOwnerMaint"), _
                   request("mappointcolor"), _
                   request("feature_public"), _
                   request("feature_maintain"), _
                   request("feature_maintain_fields"), _
                   request("feature_owners"), _
                   request("assignedto"), _
                   request("displayMap"), _
                   request("useAdvancedSearch"), _
                   request("latitude"), _
                   request("longitude"), _
                   request("defaultzoomlevel"), _
                   request("googleMapType"), _
                   request("googleMapMarker"), _
                   request("isTemplate"), _
                   request("DMTemplateID"), _
                   request("layoutid"), _
                   request("accountInfoSectionID"), _
                   request("defaultcategoryid"), _
                   request("includeBlankCategoryOption"), _
                   request("intro_message"), _
                   request("orgid"), _
                   session("userid"), _
                   lcl_redirect_url

     'This maintains the custom fields data for the DM Type (bottom of maintenance screen)
      for f = 1 to request("totalFields")
         if request.form("deleteField" & f) = "Y" then
            if request("dm_fieldid" & f) <> "" then
              'First: delete all "field values" associated to this DM Type Field
               deleteDMValue "dm_fieldid", request("dm_fieldid" & f)

              'Second: delete DM Type Field
               deleteDMTypeField request("dm_fieldid" & f)
            end if
         else
           'This inserts/updates the data for each custom field

'dtb_debug("[" & request("dm_fieldid" & f) & "] - [" & request("fieldname" & f) & "] - [" & request("fieldname5") & "]")

            'maintainMapPointTypeField request("dm_fieldid" & f), request("dm_typeid"), request("orgid"), request("fieldname" & f), _
            '                          request("fieldtype" & f), request("hasAddLinkButton" & f), request("isMultiLine" & f), _
            '                          request("displayInResults" & f), request("displayInInfoPage" & f), request("resultsOrder" & f), _
            '                          request("inPublicSearch" & f)

            maintainDMTypeField request("dm_fieldid" & f), request("dm_typeid"), request("dm_sectionid" & f), _
                                request("section_fieldid" & f), request("orgid"), request("displayInResults" & f), _
                                request("displayInInfoPage" & f), request("resultsOrder" & f), _
                                request("inPublicSearch" & f), request("displayFieldName" & f), _
                                request("isSidebarLink" & f)
         end if

      next

     'Reorder the DM Type fields
      reorderDMTypeFields request("dm_typeid")

      response.redirect lcl_redirect_url & lcl_return_url

   else
      deleteDMType request("dm_typeid"), lcl_isTemplate
   end if
else
   response.redirect "datamgr_types_list.asp" & replace(lcl_return_url,"&","?")
end if

'------------------------------------------------------------------------------
sub updateDMType(ByVal iAction, _
                 ByVal iDMTypeID, _
                 ByVal iDescription, _
                 ByVal iIsActive, _
                 ByVal iEnableOwnerMaint, _
                 ByVal iMapPointColor, _
                 ByVal iFeaturePublic, _
                 ByVal iFeatureMaintain, _
                 ByVal iFeatureMaintainFields, _
                 ByVal iFeatureOwners, _
                 ByVal iAssignedTo, _
                 ByVal iDisplayMap, _
                 ByVal iUseAdvancedSearch, _
                 ByVal iLatitude, _
                 ByVal iLongitude, _
                 ByVal iDefaultZoomLevel, _
                 ByVal iGoogleMapType, _
                 ByVal iGoogleMapMarker, _
                 ByVal iIsTemplate, _
                 ByVal iDMTemplateID, _
                 ByVal iLayoutID, _
                 ByVal iAccountInfoSectionID, _
                 ByVal iDefaultCategoryID, _
                 ByVal iIncludeBlankCategoryOption, _
                 ByVal iIntro_Message, _
                 ByVal iOrgID, _
                 ByVal iUserID, _
                 ByRef lcl_redirect_url)

 sOrgID = iOrgID

 if iDMTypeID <> "" then
    sDMTypeID = CLng(iDMTypeID)
 else
    sDMTypeID = 0
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

 if iEnableOwnerMaint = "Y" then
    sEnableOwnerMaint = 1
 else
    sEnableOwnerMaint = 0
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

 if iDefaultZoomLevel <> "" then
    sDefaultZoomLevel = iDefaultZoomLevel
 else
    sDefaultZoomLevel = "13"
 end if

 if iGoogleMapType <> "" then
    sGoogleMapType = ucase(iGoogleMapType)
    sGoogleMapType = dbsafe(sGoogleMapType)
 else
    sGoogleMapType = "ROADMAP"
 end if

 sGoogleMapType = "'" & sGoogleMapType & "'"

 if iGoogleMapMarker <> "" then
    sGoogleMapMarker = ucase(iGoogleMapMarker)
    sGoogleMapMarker = dbsafe(sGoogleMapMarker)
 else
    sGoogleMapMarker = "GOOGLE"
 end if

 sGoogleMapMarker = "'" & sGoogleMapMarker & "'"

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

 if iFeatureOwners <> "" then
    sFeatureOwners = "'" & dbsafe(iFeatureOwners) & "'"
 else
    sFeatureOwners = "NULL"
 end if

 if iAssignedTo <> "" then
    sAssignedTo = clng(iAssignedTo)
 else
    sAssignedTo = 0
 end if

 if iLayoutID <> "" then
    sLayoutID = CLng(iLayoutID)
 else
    sLayoutID = getOriginalLayoutID()
 end if

 if iAccountInfoSectionID <> "" then
    sAccountInfoSectionID = CLng(iAccountInfoSectionID)
 else
    sAccountInfoSectionID = 0
 end if

 if iDefaultCategoryID <> "" then
    sDefaultCategoryID = CLng(iDefaultCategoryID)
 else
    sDefaultCategoryID = 0
 end if

 if iIncludeBlankCategoryOption = "Y" then
    sIncludeBlankCategoryOption = 1
 else
    sIncludeBlankCategoryOption = 0
 end if

 if iIntro_Message <> "" then
    sIntro_Message = "'" & dbsafe(iIntro_Message) & "'"
 else
    sIntro_Message = "NULL"
 end if

'The dm type exists, so update it
 if iAction = "UPDATE" then
  		sSQL = "UPDATE egov_dm_types SET "
    sSQL = sSQL & "description = "                & sDescription                        & ", "
    sSQL = sSQL & "isActive = "                   & sIsActive                           & ", "
    sSQL = sSQL & "enableOwnerMaint = "           & sEnableOwnerMaint                   & ", "
    sSQL = sSQL & "mappointcolor = "              & sMapPointColor                      & ", "
    sSQL = sSQL & "feature_public = "             & sFeaturePublic                      & ", "
    sSQL = sSQL & "feature_maintain = "           & sFeatureMaintain                    & ", "
    sSQL = sSQL & "feature_maintain_fields = "    & sFeatureMaintainFields              & ", "
    sSQL = sSQL & "feature_owners = "             & sFeatureOwners                      & ", "
    sSQL = sSQL & "assignedto = "                 & sAssignedTo                         & ", "
    sSQL = sSQL & "displayMap = "                 & sDisplayMap                         & ", "
    sSQL = sSQL & "useAdvancedSearch = "          & sUseAdvancedSearch                  & ", "
    sSQL = sSQL & "latitude = "                   & sLatitude                           & ", "
    sSQL = sSQL & "longitude = "                  & sLongitude                          & ", "
    sSQL = sSQL & "defaultzoomlevel = "           & sDefaultZoomLevel                   & ", "
    sSQL = sSQL & "googleMapType = "              & sGoogleMapType                      & ", "
    sSQL = sSQL & "googleMapMarker = "            & sGoogleMapMarker                    & ", "
    sSQL = sSQL & "lastmodifiedbyid = "           & sUserID                             & ", "
    sSQL = sSQL & "lastmodifiedbydate = '"        & dbsafe(ConvertDateTimetoTimeZone()) & "', "
    sSQL = ssQL & "isTemplate = "                 & sIsTemplate                         & ", "
    sSQL = sSQL & "layoutid = "                   & sLayoutID                           & ", "
    sSQL = sSQL & "accountInfoSectionID = "       & sAccountInfoSectionID               & ", "
    sSQL = sSQL & "defaultcategoryid = "          & sDefaultCategoryID                  & ", "
    sSQL = sSQL & "includeBlankCategoryOption = " & sIncludeBlankCategoryOption         & ", "
    sSQL = sSQL & "intro_message = "              & sIntro_Message
    sSQL = sSQL & " WHERE dm_typeid = " & sDMTypeID

  		set oDMTypes = Server.CreateObject("ADODB.Recordset")
	  	oDMTypes.Open sSQL, Application("DSN"), 3, 1

    set oDMTypes = nothing

    lcl_redirect_url = "datamgr_types_maint.asp?dm_typeid=" & sDMTypeID & "&success=SU" & sIsTemplate_url

'------------------------------------------------------------------------------
 else  'New DM Type
'------------------------------------------------------------------------------
    sCreatedByID   = iUserID
    sCreatedByDate = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

   'If we are creating this DM Type from a template then we need to get:
   ' - LayoutID
   ' - AccountInfoSectionID
    if iIsTemplate <> "Y" then
       if iDMTemplateID <> "" then
          sLayoutID             = getDMTLayoutID(iDMTemplateID)
          sAccountInfoSectionID = getAccountInfoSectionID(iDMTemplateID)
       end if
    end if

 		'Insert the new DM Type
  		sSQL = "INSERT INTO egov_dm_types ("
    sSQL = sSQL & "orgid, "
    sSQL = sSQL & "description, "
    sSQL = sSQL & "isActive, "
    sSQL = sSQL & "enableOwnerMaint, "
    sSQL = sSQL & "mappointcolor, "
    sSQL = sSQL & "createdbyid, "
    sSQL = sSQL & "createdbydate, "
    sSQL = sSQL & "lastmodifiedbyid, "
    sSQL = sSQL & "lastmodifiedbydate, "
    sSQL = sSQL & "feature_public, "
    sSQL = sSQL & "feature_maintain, "
    sSQL = sSQL & "feature_maintain_fields, "
    sSQL = sSQL & "feature_owners, "
    sSQL = sSQL & "assignedto, "
    sSQL = sSQL & "displayMap, "
    sSQL = sSQL & "useAdvancedSearch, "
    sSQL = sSQL & "latitude, "
    sSQL = sSQL & "longitude, "
    sSQL = sSQL & "defaultzoomlevel, "
    sSQL = sSQL & "googleMapType, "
    sSQL = sSQL & "googleMapMarker, "
    sSQL = sSQL & "isTemplate, "
    sSQL = sSQL & "layoutid, "
    sSQL = sSQL & "accountInfoSectionID, "
    sSQL = sSQL & "defaultcategoryid, "
    sSQL = sSQL & "includeBlankCategoryOption, "
    sSQL = sSQL & "intro_message"
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & sOrgID                      & ", "
    sSQL = sSQL & sDescription                & ", "
    sSQL = sSQL & sIsActive                   & ", "
    sSQL = sSQL & sEnableOwnerMaint           & ", "
    sSQL = sSQL & sMapPointColor              & ", "
    sSQL = sSQL & sCreatedByID                & ", "
    sSQL = sSQL & sCreatedByDate              & ", "
    sSQL = sSQL & "NULL,NULL,"
    sSQL = sSQL & sFeaturePublic              & ", "
    sSQL = sSQL & sFeatureMaintain            & ", "
    sSQL = sSQL & sFeatureMaintainFields      & ", "
    sSQL = sSQL & sFeatureOwners              & ", "
    sSQL = sSQL & sAssignedTo                 & ", "
    sSQL = sSQL & sDisplayMap                 & ", "
    sSQL = sSQL & sUseAdvancedSearch          & ", "
    sSQL = sSQL & sLatitude                   & ", "
    sSQL = sSQL & sLongitude                  & ", "
    sSQL = sSQL & sDefaultZoomLevel           & ", "
    sSQL = sSQL & sGoogleMapType              & ", "
    sSQL = sSQL & sGoogleMapmarker            & ", "
    sSQL = sSQL & sIsTemplate                 & ", "
    sSQL = sSQL & sLayoutID                   & ", "
    sSQL = sSQL & sAccountInfoSectionID       & ", "
    sSQL = sSQL & sDefaultCategoryID          & ", "
    sSQL = sSQL & sIncludeBlankCategoryOption & ", "
    sSQL = sSQL & sIntro_Message
    sSQL = sSQL & ")"

 		'Get the DMTypeID
	  	sDMTypeID = RunIdentityInsert(sSQL)

   'If a "fields template" has been selected then get all of the sections and 
   'section fields from the template to add to the new DM Type
    if iIsTemplate <> "Y" then
       if iDMTemplateID <> "" then
          buildDMTypeFromTemplate sOrgID, sDMTypeID, iDMTemplateID
       end if
    end if

    lcl_redirect_url = "datamgr_types_maint.asp?success=SA" & sIsTemplate_url

    if iAction = "ADD" then
       lcl_redirect_url = lcl_redirect_url & "&dm_typeid=" & sDMTypeID
    end if

 end if

'If an AccountInfoSectionID has been selected then we need to make sure that the section fields
'for the AccountInfo section exist for the DMTypeID.
 if sAccountInfoSectionID > 0 then
    lcl_dm_sectionid  = 0
    lcl_sectionactive = "Y"

    maintainDMTSectionFields sDMTypeID, sOrgID, lcl_dm_sectionid, sAccountInfoSectionID, lcl_sectionactive
 end if

end sub

'------------------------------------------------------------------------------
sub deleteDMType(iDMTypeID, iIsTemplate)

  if iDMTypeID <> "" then
     sDMTypeID = CLng(iDMTypeID)
  else
     sDMTypeID = 0
  end if

 'Delete all of the DM Category Assignments
  sSQL = "DELETE FROM egov_dmdata_to_dmcategories "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

	 set oDeleteDMCatAssignments = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMCatAssignments.Open sSQL, Application("DSN"), 3, 1

 'Delete all of the DM Categories
  sSQL = "DELETE FROM egov_dm_categories "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

	 set oDeleteDMCategories = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMCategories.Open sSQL, Application("DSN"), 3, 1

 'Delete all of the DM Values
  sSQL = "DELETE FROM egov_dm_values "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

	 set oDeleteDMValues = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMValues.Open sSQL, Application("DSN"), 3, 1

 'Delete the DM Data
  sSQL = "DELETE FROM egov_dm_data "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

	 set oDeleteDMData = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMData.Open sSQL, Application("DSN"), 3, 1

 'Delete all of the DM Types Fields for the DM Type
  sSQL = "DELETE FROM egov_dm_types_fields "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

	 set oDeleteDMTFields = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMTFields.Open sSQL, Application("DSN"), 3, 1

 'Delete all of the DM Types Sections for the DM Type
  sSQL = "DELETE FROM egov_dm_types_sections "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

	 set oDeleteDMTSections = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMTSections.Open sSQL, Application("DSN"), 3, 1

 'Delete the DM Type
  sSQL = "DELETE FROM egov_dm_types "
  sSQL = sSQL & " WHERE dm_typeid = " & sDMTypeID

	 set oDeleteDMType = Server.CreateObject("ADODB.Recordset")
 	oDeleteDMType.Open sSQL, Application("DSN"), 3, 1

  set oDeleteDMCatAssignments = nothing
  set oDeleteDMCategories     = nothing
  set oDeleteDMValues         = nothing
  set oDeleteDMData           = nothing
  set oDeleteDMTFields        = nothing
  set oDeleteDMTSections      = nothing
  set oDeleteDMType           = nothing

  lcl_isTemplate_url = ""

  if iIsTemplate = "Y" then
     lcl_isTemplate_url = "&t=Y"
  end if

  response.redirect "datamgr_types_list.asp?success=SD" & lcl_isTemplate_url

end sub

'------------------------------------------------------------------------------
sub reorderDMTypeFields(iDMTypeID)

  lcl_order = 0

  if iDMTypeID <> "" then
     sSQL = "SELECT dm_fieldid "
     sSQL = sSQL & " FROM egov_dm_types_fields "
     sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID
     sSQL = sSQL & " ORDER BY resultsOrder "

   	 set oReorderDMTFields = Server.CreateObject("ADODB.Recordset")
    	oReorderDMTFields.Open sSQL, Application("DSN"), 3, 1

     if not oReorderDMTFields.eof then
        do while not oReorderDMTFields.eof

           lcl_order = lcl_order + 1

           sSQL = "UPDATE egov_dm_types_fields SET resultsOrder = " & lcl_order
           sSQL = sSQL & " WHERE dm_fieldid = " & oReorderDMTFields("dm_fieldid")

   	       set oUpdateOrderBy = Server.CreateObject("ADODB.Recordset")
          	oUpdateOrderBy.Open sSQL, Application("DSN"), 3, 1

           set oUpdateOrderBy = nothing

           oReorderDMTFields.movenext
        loop
     end if

     oReorderDMTFields.close
     set oReorderDMTFields = nothing
  end if

end sub
%>