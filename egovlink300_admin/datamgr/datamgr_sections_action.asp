<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  if request("user_action") = "" then
     response.redirect "datamgr_sections_list.asp"
  end if

  lcl_useraction = ""
  lcl_sectionid  = 0

  if request("user_action") <> "" then
     lcl_useraction = UCASE(request("user_action"))
  end if

  if request("sectionid") <> "" then
     lcl_sectionid = request("sectionid")
  end if

 'Execute the user's action
  if lcl_useraction = "DELETE" then

    'Delete all of the MapPoint Values
     sSQL = "DELETE FROM egov_dm_values "
     sSQL = sSQL & " WHERE dm_sectionid IN (select dm_sectionid "
     sSQL = sSQL &                        " from egov_dm_types_sections "
     sSQL = sSQL &                        " where sectionid = " & lcl_sectionid & ") "

    'Delete all of the MapPoint Type Fields
     sSQL1 = "DELETE FROM egov_dm_types_fields "
     sSQL1 = sSQL1 & " WHERE dm_sectionid IN (select dm_sectionid "
     sSQL1 = sSQL1 &                        " from egov_dm_types_sections "
     sSQL1 = sSQL1 &                        " where sectionid = " & lcl_sectionid & ") "

    'Delete all of the fields associated to the section
     sSQL2 = "DELETE FROM egov_dm_sections_fields WHERE sectionid = " & lcl_sectionid

    'Delete the section
     sSQL3 = "DELETE FROM egov_dm_sections WHERE sectionid = " & lcl_sectionid

   		set oDeleteDMSection = Server.CreateObject("ADODB.Recordset")

    	oDeleteDMSection.Open sSQL,  Application("DSN"), 3, 1
    	oDeleteDMSection.Open sSQL1, Application("DSN"), 3, 1
    	oDeleteDMSection.Open sSQL2, Application("DSN"), 3, 1
    	oDeleteDMSection.Open sSQL3, Application("DSN"), 3, 1

     set oDeleteDMSection = nothing

     lcl_url_parameters = ""
     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SD")
     lcl_redirect_url   = "datamgr_sections_list.asp" & lcl_url_parameters

  else
     'This maintains the MapPoint Section data (top of maintenance screen)
      updateDMTSection lcl_useraction, _
                       lcl_sectionid, _
                       request("sectionname"), _
                       request("sectiontype"), _
                       request("description"), _
                       request("isActive"), _
                       request("isAccountInfoSection"), _
                       request("displaySectionName"), _
                       request("section_orgid"), _
                       session("userid"), _
                       lcl_redirect_url

     'This maintains the custom fields data for the Section (bottom of maintenance screen)
      for f = 1 to request("totalFields")
         if request.form("deleteField" & f) = "Y" then
            if request("section_fieldid" & f) <> "" then
               deleteDMTSectionField request("section_fieldid" & f)
            end if
         else
           'This inserts/updates the data for each custom field
            maintainDMSectionField request("section_fieldid" & f), _
                                   lcl_sectionid, _
                                   request("fieldname" & f), _
                                   request("fieldtype" & f), _
                                   request("sectionfield_isActive" & f), _
                                   request("displayFieldName" & f), _
                                   request("isMultiLine" & f), _
                                   request("hasAddLinkButton" & f), _
                                   request("displayOrder" & f), _
                                   session("orgid"), _
                                   session("userid")
         end if

      next

     'Reorder the MPT Section fields
      reorderDMTSectionFields request("sectionid")

  end if

  response.redirect lcl_redirect_url


'------------------------------------------------------------------------------
sub updateDMTSection(ByVal iAction, ByVal iSectionID, ByVal iSectionName, ByVal iSectionType, ByVal iDescription, _
                     ByVal iIsActive, ByVal iIsAccountInfoSection, ByVal iDisplaySectionName, ByVal iSectionOrgID, _
                     ByVal iUserID, ByRef lcl_redirect_url)

 if iSectionID <> "" then
    sSectionID = CLng(iSectionID)
 else
    sSectionID = 0
 end if

 if iSectionName = "" then
  		sSectionName = "NULL"
 else
  		sSectionName = "'" & dbsafe(iSectionName) & "'"
 end if

 if iSectionType = "" then
  		sSectionType = "NULL"
 else
    sSectionType = ucase(iSectionType)
  		sSectionType = "'" & dbsafe(sSectionType) & "'"
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

 if iIsAccountInfoSection = "Y" then
    sIsAccountInfoSection = 1
 else
    sIsAccountInfoSection = 0
 end if

 if iDisplaySectionName = "Y" then
    sDisplaySectionName = 1
 else
    sDisplaySectionName = 0
 end if

 if iSectionOrgID <> "" then
    sSectionOrgID = clng(iSectionOrgID)
 else
    sSectionOrgID = 0
 end if

 if iUserID <> "" then
    sUserID = CLng(iUserID)
 else
    sUserID = 0
 end if

 lcl_current_date = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

'-- Update DM Section ---------------------------------------------------------
 if lcl_useraction = "UPDATE" then
'------------------------------------------------------------------------------

    sSQL = "UPDATE egov_dm_sections SET "
    sSQL = sSQL & "sectionname = "          & sSectionName          & ", "
    sSQL = sSQL & "sectiontype = "          & sSectionType          & ", "
    sSQL = sSQL & "description = "          & sDescription          & ", "
    sSQL = sSQL & "isActive = "             & sIsActive             & ", "
    sSQL = sSQL & "isAccountInfoSection = " & sIsAccountInfoSection & ", "
    sSQL = sSQL & "displaySectionName = "   & sDisplaySectionName   & ", "
    sSQL = sSQL & "section_orgid = "        & sSectionOrgID         & ", "
    sSQL = sSQL & "lastmodifiedbyid = "     & sUserID               & ", "
    sSQL = sSQL & "lastmodifiedbydate = "   & lcl_current_date
    sSQL = sSQL & " WHERE sectionid = " & sSectionID

  		set oUpdateDMSections = Server.CreateObject("ADODB.Recordset")
   	oUpdateDMSections.Open sSQL, Application("DSN"), 3, 1

    set oUpdateDMSections = nothing

    lcl_url_parameters = ""
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sectionid", lcl_sectionid)
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SU")
    lcl_redirect_url   = "datamgr_sections_maint.asp" & lcl_url_parameters

'---------------------------------------------------------------------------
 else  'New DM Section
'---------------------------------------------------------------------------
   	sSQL = "INSERT INTO egov_dm_sections ("
    sSQL = sSQL & "sectionname, "
    sSQL = sSQL & "sectiontype, "
    sSQL = sSQL & "description, "
    sSQL = sSQL & "isActive, "
    sSQL = sSQL & "isAccountInfoSection, "
    sSQL = sSQL & "displaySectionName, "
    sSQL = sSQL & "section_orgid, "
    sSQL = sSQL & "createdbyid, "
    sSQL = sSQL & "createdbydate, "
    sSQL = sSQL & "lastmodifiedbyid, "
    sSQL = sSQL & "lastmodifiedbydate "
    sSQL = sSQL & ") VALUES ("
    sSQL = sSQL & sSectionName          & ", "
    sSQL = sSQL & sSectionType          & ", "
    sSQL = sSQL & sDescription          & ", "
    sSQL = sSQL & sIsActive             & ", "
    sSQL = sSQL & sIsAccountInfoSection & ", "
    sSQL = sSQL & sDisplaySectionName   & ", "
    sSQL = sSQL & sSectionOrgID         & ", "
    sSQL = sSQL & sUserID               & ", "
    sSQL = sSQL & lcl_current_date      & ", "
    sSQL = sSQL & "NULL,NULL"
    sSQL = sSQL & ")"

   'Get the SectionID
	  	lcl_sectionid = RunIdentityInsert(sSQL)

    lcl_url_parameters = ""
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sectionid", lcl_sectionid)
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SA")
    lcl_redirect_url   = "datamgr_sections_maint.asp" & lcl_url_parameters
 end if

end sub

'------------------------------------------------------------------------------
sub deleteDMTSectionField(iSectionFieldID)

 'Delete all of the MapPoint Values
  sSQL = "DELETE FROM egov_dm_values "
  sSQL = sSQL & " WHERE dm_fieldid IN (select dm_fieldid "
  sSQL = sSQL &                        " from egov_dm_types_fields "
  sSQL = sSQL &                        " where section_fieldid = " & iSectionFieldID & ") "

 'Delete all of the MapPoint Type Fields
  sSQL1 = "DELETE FROM egov_dm_types_fields WHERE section_fieldid = " & iSectionFieldID

 'Delete all of the fields associated to the section
  sSQL2 = "DELETE FROM egov_dm_sections_fields WHERE section_fieldid = " & iSectionFieldID

	 set oDeleteDMTSectionField = Server.CreateObject("ADODB.Recordset")

 	oDeleteDMTSectionField.Open sSQL,  Application("DSN"), 3, 1
 	oDeleteDMTSectionField.Open sSQL1, Application("DSN"), 3, 1
 	oDeleteDMTSectionField.Open sSQL2, Application("DSN"), 3, 1

  set oDeleteDMTSectionField = nothing

end sub

'------------------------------------------------------------------------------
sub maintainDMSectionField(iSectionFieldID, iSectionID, iFieldName, iFieldType, iIsActive, iDisplayFieldName, _
                           iIsMultiLine, iHasAddLinkButton, iDisplayOrder, iOrgID, iUserID)

  if iFieldName <> "" then
     lcl_fieldname = "'" & dbsafe(iFieldName) & "'"
  else
     lcl_fieldname = "NULL"
  end if

  if iFieldType <> "" then
     lcl_fieldtype = "'" & dbsafe(UCASE(iFieldType)) & "'"
  else
     lcl_fieldtype = "NULL"
  end if

  if iIsActive = "Y" then
     lcl_isActive = 1
  else
     lcl_isActive = 0
  end if

  if iDisplayFieldName = "Y" then
     lcl_displayFieldName = 1
  else
     lcl_displayFieldName = 0
  end if

  if iIsMultiLine <> "" then
     lcl_isMultiLine = iIsMultiLine
  else
     lcl_isMultiLine = 0
  end if

  if iHasAddLinkButton <> "" then
     lcl_hasAddLinkButton = iHasAddLinkButton
  else
     lcl_hasAddLinkButton = 0
  end if

  if iDisplayOrder <> "" then
     lcl_displayOrder = iDisplayOrder
  else
     lcl_displayOrder = 0
  end if

  if iUserID <> "" then
     sUserID = CLng(iUserID)
  else
     sUserID = 0
  end if

  lcl_current_date = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

  lcl_insertIntoDMTSectionFields = False

 'Determine if the DM Section Field is to be added or updated
  if iSectionFieldID <> "" then
     sSQL = "UPDATE egov_dm_sections_fields SET "
     sSQL = sSQL & "sectionid = "          & iSectionID           & ", "
     sSQL = sSQL & "fieldname = "          & lcl_fieldname        & ", "
     sSQL = sSQL & "fieldtype = "          & lcl_fieldtype        & ", "
     sSQL = sSQL & "isActive = "           & lcl_isActive         & ", "
     sSQL = sSQL & "displayFieldName = "   & lcl_displayFieldName & ", "
     sSQL = sSQL & "isMultiLine = "        & lcl_isMultiLine      & ", "
     sSQL = sSQL & "hasAddLinkButton = "   & lcl_hasAddLinkButton & ", "
     sSQL = sSQL & "displayOrder = "       & lcl_displayOrder     & ", "
     sSQL = sSQL & "lastmodifiedbyid = "   & sUserID              & ", "
     sSQL = sSQL & "lastmodifiedbydate = " & lcl_current_date
     sSQL = sSQL & " WHERE section_fieldid = " & iSectionFieldID

   	 set oUpdateSectionField = Server.CreateObject("ADODB.Recordset")
    	oUpdateSectionField.Open sSQL, Application("DSN"), 3, 1

     set oUpdateSectionField = nothing

     lcl_section_fieldid = iSectionFieldID

    'If this field is "active" and does NOT exist on MapPoint Type Fields then insert the record
     'if iIsActive = "Y" then
     '   lcl_insertIntoDMTSectionFields = True
     'end if

  else
     sSQL = "INSERT INTO egov_dm_sections_fields ("
     sSQL = sSQL & "sectionid,"
     sSQL = sSQL & "fieldname,"
     sSQL = sSQL & "fieldtype,"
     sSQL = sSQL & "isActive, "
     sSQL = sSQL & "displayFieldName, "
     sSQL = sSQL & "isMultiLine,"
     sSQL = sSQL & "hasAddLinkButton,"
     sSQL = sSQL & "displayOrder, "
     sSQL = sSQL & "createdbyid, "
     sSQL = sSQL & "createdbydate, "
     sSQL = sSQL & "lastmodifiedbyid, "
     sSQL = sSQL & "lastmodifiedbydate "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & iSectionID           & ", "
     sSQL = sSQL & lcl_fieldname        & ", "
     sSQL = sSQL & lcl_fieldtype        & ", "
     sSQL = sSQL & lcl_isActive         & ", "
     sSQL = sSQL & lcl_displayFieldName & ", "
     sSQL = sSQL & lcl_isMultiLine      & ", "
     sSQL = sSQL & lcl_hasAddLinkButton & ", "
     sSQL = sSQL & lcl_displayOrder     & ", "
     sSQL = sSQL & sUserID              & ", "
     sSQL = sSQL & lcl_current_date     & ", "
     sSQL = sSQL & "NULL,NULL"
     sSQL = sSQL & ")"

 	  	lcl_section_fieldid = RunIdentityInsert(sSQL)

    'If this field is "active" then we need to add it to all MapPoint Type Fields that have this section.
     'if iIsActive = "Y" then
     '   lcl_insertIntoDMTSectionFields = True
     'end if
  end if

 'Determine if we are to insert the Section Field into existing MapPoint Type Fields
  if iIsActive = "Y" then
     insertNewDMSectionField iOrgID, iSectionID, lcl_section_fieldid
  end if

end sub

'------------------------------------------------------------------------------
sub insertNewDMSectionField(iOrgID, iSectionID, iSectionFieldID)
 'Get a distinct list of DM_SectionIDs.  The following must be TRUE:
 '  1. the section for the field is "ACTIVE"
 '  2. the MapPoint Type Section is "ACTIVE"
  sSQL = "SELECT distinct dmts.dm_sectionid, dmts.dm_typeid "
  sSQL = sSQL & " FROM egov_dm_types_sections dmts "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
  sSQL = sSQL &                 " ON dmts.sectionid = dms.sectionid "
  sSQL = sSQL &                 " AND dms.isActive = 1 "
  sSQL = sSQL &                 " AND dms.sectionid = " & iSectionID
  sSQL = sSQL & " WHERE dmts.isActive = 1 "
  sSQL = sSQL & " AND dmts.dm_sectionid NOT IN (select distinct dmtf.dm_sectionid "
  sSQL = sSQL &                               " from egov_dm_types_fields dmtf "
  sSQL = sSQL &                               " where dmtf.section_fieldid = " & iSectionFieldID & ") "

  set oInsertNewDMSField = Server.CreateObject("ADODB.Recordset")
  oInsertNewDMSField.Open sSQL, Application("DSN"), 3, 1

  if not oInsertNewDMSField.eof then
     do while not oInsertNewDMSField.eof

        lcl_displayInResults  = "0"
        lcl_displayInInfoPage = "1"
        lcl_resultsOrder      = 1
        lcl_inPublicSearch    = "0"
        lcl_displayFieldName  = "0"
        lcl_dmfieldid         = ""
        lcl_isSidebarLink     = ""

        maintainDMTypeField lcl_dmfieldid, oInsertNewDMSField("dm_typeid"), oInsertNewDMSField("dm_sectionid"), _
                            iSectionFieldID, iOrgID, lcl_displayInResults, lcl_displayInInfoPage, _
                            lcl_resultsOrder, lcl_inPublicSearch, lcl_displayFieldName, lcl_isSidebarLink

        oInsertNewDMSField.movenext
     loop
  end if

  oInsertNewDMSField.close
  set oInsertNewDMSField = nothing

end sub

'------------------------------------------------------------------------------
sub reorderDMTSectionFields(iSectionID)

  lcl_order = 0

  if iSectionID <> "" then
     sSQL = "SELECT section_fieldid "
     sSQL = sSQL & " FROM egov_dm_sections_fields "
     sSQL = sSQL & " WHERE sectionid = " & iSectionID
     sSQL = sSQL & " ORDER BY displayOrder "

   	 set oReorderDMTSectionFields = Server.CreateObject("ADODB.Recordset")
    	oReorderDMTSectionFields.Open sSQL, Application("DSN"), 3, 1

     if not oReorderDMTSectionFields.eof then
        do while not oReorderDMTSectionFields.eof

           lcl_order = lcl_order + 1

           sSQL = "UPDATE egov_dm_sections_fields SET displayOrder = " & lcl_order
           sSQL = sSQL & " WHERE section_fieldid = " & oReorderDMTSectionFields("section_fieldid")

   	       set oUpdateOrderBy = Server.CreateObject("ADODB.Recordset")
          	oUpdateOrderBy.Open sSQL, Application("DSN"), 3, 1

           set oUpdateOrderBy = nothing

           oReorderDMTSectionFields.movenext
        loop
     end if

     oReorderDMTSectionFields.close
     set oReorderDMTSectionFields = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub

%>