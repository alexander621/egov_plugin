<!-- #include file="datamgr_build_sections_functions.asp" //-->
<%
'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  dim lcl_return, lcl_orgid, lcl_dmid

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "Does not exist..."
     elseif iSuccess = "ERROR" then
        lcl_return = "ERROR"
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function setupUserMaintLogInfo(iName, iDate)
  dim lcl_return

  lcl_return = ""

  if iName <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & iName
     else
        lcl_return = iName
     end if
  end if

  if iDate <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & " on " & iDate
     else
        lcl_return = iDate
     end if
  end if

  setupUserMaintLogInfo = lcl_return

end function

'------------------------------------------------------------------------------
'sub maintainMapPointTypeField(iMPFieldID, iMapPointTypeID, iMPSectionID, iSectionFieldID, iOrgID, _
'                              iDisplayInResults, iDisplayInInfoPage, iResultsOrder, iInPublicSearch, _
'                              iDisplayFieldName)

sub maintainDMTypeField(iDMFieldID, iDMTypeID, iDMSectionID, iSectionFieldID, iOrgID, _
                        iDisplayInResults, iDisplayInInfoPage, iResultsOrder, iInPublicSearch, _
                        iDisplayFieldName, iIsSidebarLink)

  dim lcl_dm_fieldid, lcl_displayInResults, lcl_displayInInfoPage, lcl_resultsOrder
  dim lcl_inPublicSearch, lcl_displayFieldName, lcl_isSidebarLink

  lcl_dm_fieldid        = iDMFieldID
  lcl_displayInResults  = 0
  lcl_displayInInfoPage = 0
  lcl_resultsOrder      = 0
  lcl_inPublicSearch    = 0
  lcl_displayFieldName  = 0
  lcl_isSidebarLink     = 0

  if iDisplayInResults <> "" then
     lcl_displayInResults = iDisplayInResults
  end if

  if iDisplayInInfoPage <> "" then
     lcl_displayInInfoPage = iDisplayInInfoPage
  end if

  if iResultsOrder <> "" then
     lcl_resultsOrder = iResultsOrder
  end if

  if iInPublicSearch <> "" then
     lcl_inPublicSearch = iInPublicSearch
  end if

  if iDisplayFieldName <> "" then
     lcl_displayFieldName = iDisplayFieldName
  end if

  if iIsSidebarLink <> "" then
     lcl_isSidebarLink = iIsSidebarLink
  end if

 'Determine if the Map-Point is to be added or updated
  if lcl_dm_fieldid <> "" then
     sSQL = "UPDATE egov_dm_types_fields SET "
     sSQL = sSQL & " displayInResults = "  & lcl_displayInResults  & ", "
     sSQL = sSQL & " displayInInfoPage = " & lcl_displayInInfoPage & ", "
     sSQL = sSQL & " resultsOrder = "      & lcl_resultsOrder      & ", "
     sSQL = sSQL & " inPublicSearch = "    & lcl_inPublicSearch    & ", "
     sSQL = sSQL & " displayFieldName = "  & lcl_displayFieldName  & ", "
     sSQL = sSQL & " isSidebarLink = "     & lcl_isSidebarLink
     sSQL = sSQL & " WHERE dm_fieldid = " & lcl_dm_fieldid

   	 set oMaintainDMTypeField = Server.CreateObject("ADODB.Recordset")
    	oMaintainDMTypeField.Open sSQL, Application("DSN"), 3, 1

     set oMaintainDMTypeField = nothing

  else
     sSQL = "INSERT INTO egov_dm_types_fields ("
     sSQL = sSQL & "dm_typeid,"
     sSQL = sSQL & "orgid,"
     sSQL = sSQL & "dm_sectionid,"
     sSQL = sSQL & "section_fieldid, "
     sSQL = sSQL & "displayInResults,"
     sSQL = sSQL & "displayInInfoPage,"
     sSQL = sSQL & "resultsOrder, "
     sSQL = sSQL & "inPublicSearch, "
     sSQL = sSQL & "displayFieldName, "
     sSQL = sSQL & "isSidebarLink "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & iDMTypeID       & ", "
     sSQL = sSQL & iOrgID                & ", "
     sSQL = sSQL & iDMSectionID          & ", "
     sSQL = sSQL & iSectionFieldID       & ", "
     sSQL = sSQL & lcl_displayInResults  & ", "
     sSQL = sSQL & lcl_displayInInfoPage & ", "
     sSQL = sSQL & lcl_resultsOrder      & ", "
     sSQL = sSQL & lcl_inPublicSearch    & ", "
     sSQL = sSQL & lcl_displayFieldName  & ", "
     sSQL = sSQL & lcl_isSidebarLink
     sSQL = sSQL & ")"

     lcl_dm_fieldid = RunIdentityInsert(sSQL)
  end if

'CANNOT DO THIS BECAUSE WE DO NOT HAVE A DMID!!! ------------------------
 'Now we need to check to see if there is a MapPoints Value record for this field.
 'If not then we need to create one.
'  lcl_mpvalue_exists = checkMPValueExists(lcl_dm_fieldid)

'  if not lcl_mpvalue_exists then
'     lcl_dmid = 0

'     maintainDMValues(iOrgID, iDMTypeID, iDMID, iDMSectionID, iDMFieldID, iDMValueID, iFieldValue)
'  end if

end sub

'------------------------------------------------------------------------------
'function checkMPValueExists(iDMFieldID)

'  lcl_return = False

'  if iDMFieldID <> "" then
'     sSQL = "SELECT count(dm_valueid) as total_count "
'     sSQL = sSQL & " FROM egov_dm_values "
'     sSQL = sSQL & " WHERE dm_fieldid = " & iDMFieldID

'   	 set oCheckMPVExists = Server.CreateObject("ADODB.Recordset")
'    	oCheckMPVExists.Open sSQL, Application("DSN"), 3, 1

'     if not oCheckMPVExists.eof then
'        if oCheckMPVExists("total_count") > 0 then
'           lcl_return = True
'        end if
'     end if

'     oCheckMPVExists.close
'     set oCheckMPVExists = nothing
'  end if

'  checkMPValueExists = lcl_return

'end function

'------------------------------------------------------------------------------
'function checkMapPointValueExists(p_dm_fieldid)

'  lcl_return = False
'  lcl_exists = "N"

'  if p_dm_fieldid <> "" then

'     sSQL = "SELECT distinct 'Y' as mpvalue_exists "
'     sSQL = sSQL & " FROM egov_dm_values "
'     sSQL = sSQL & " WHERE dm_fieldid = " & p_dm_fieldid

'   	 set oMPValueExists = Server.CreateObject("ADODB.Recordset")
'    	oMPValueExists.Open sSQL, Application("DSN"), 3, 1

'     if not oMPValueExists.eof then
'        lcl_exists = oMPValueExists("mpvalue_exists")
'     end if

'     oMPValueExists.close
'     set oMPValueExists = nothing

'  end if

'  if lcl_exists = "Y" then
'     lcl_return = True
'  end if

'  checkMapPointValueExists = lcl_return

'end function

'------------------------------------------------------------------------------
'function getMPValueID(iMPValueID, iMPTypeID, iMapPointID, iMPSectionID, iMPFieldID)
function getDMValueID(iDMValueID, iDMTypeID, iDMID, iDMSectionID, iDMFieldID)

  dim lcl_return, lcl_dm_valueid, lcl_dmid, lcl_dmtid, lcl_dmsid, lcl_dmfid

  lcl_return = 0

  sSQL = "SELECT dm_valueid  "
  sSQL = sSQL & " FROM egov_dm_values "

  if iDMValueID <> "" AND iDMValueID > 0 then
     lcl_dm_valueid = iDMValueID

     if iDMValueID = "" then
        lcl_dm_valueid = 0
     end if

     sSQL = sSQL & " WHERE dm_valueid = " & lcl_dm_valueid

  else
     'lcl_dmid  = 0
     'lcl_dmtid = 0
     'lcl_dmsid = 0
     'lcl_dmfid = 0

     lcl_dmid  = iDMID
     lcl_dmtid = iDMTypeID
     lcl_dmsid = iDMSectionID
     lcl_dmfid = iDMFieldID

     if iDMID = "" then
        lcl_dmid = 0
     end if

     if iDMTypeID = "" then
        lcl_dmtid = 0
     end if

     if iDMSectionID = "" then
        lcl_dmsid = 0
     end if

     if iDMFieldID = "" then
        lcl_dmfid = 0
     end if

     sSQL = sSQL & " WHERE dm_typeid = "  & lcl_dmtid
     sSQL = sSQL & " AND dmid = "         & lcl_dmid
     sSQL = sSQL & " AND dm_sectionid = " & lcl_dmsid
     sSQL = sSQL & " AND dm_fieldid = "   & lcl_dmfid
  end if

	 set oGetDMValueID = Server.CreateObject("ADODB.Recordset")
 	oGetDMValueID.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMValueID.eof then
     lcl_return = oGetDMValueID("dm_valueid")
  end if

  oGetDMValueID.close
  set oGetDMValueID = nothing

  getDMValueID = lcl_return

end function

'------------------------------------------------------------------------------
'sub deleteMapPointTypeField(iDMFieldID)
sub deleteDMTypeField(iDMFieldID)

  sSQL = "DELETE FROM egov_dm_types_fields WHERE dm_fieldid = " & iDMFieldID

	 set odeleteDMTypeField = Server.CreateObject("ADODB.Recordset")
 	odeleteDMTypeField.Open sSQL, Application("DSN"), 3, 1

  set odeleteDMTypeField = nothing

end sub

'------------------------------------------------------------------------------
'sub deleteMapPointValue(iDBColumnName, iID)
sub deleteDMValue(iDBColumnName, iID)

  sSQL = "DELETE FROM egov_dm_values WHERE " & iDBColumnName & " = " & iID

	 set oDelDMValue = Server.CreateObject("ADODB.Recordset")
 	oDelDMValue.Open sSQL, Application("DSN"), 3, 1

  set oDelDMValue = nothing

end sub

'------------------------------------------------------------------------------
function RunIdentityInsert( sInsertStatement )
	 dim sSQL, iReturnValue, oInsert

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
'function checkForMapPointsByMapPointTypeID(iDMTypeID)
function checkForDMDataByDMTypeID(iDMTypeID)
  dim lcl_return

  lcl_return = true

  if iDMTypeID <> "" then
     sSQL = "SELECT DISTINCT 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_dm_data "
     sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

    	set oCheckForDMData = Server.CreateObject("ADODB.Recordset")
   	 oCheckForDMData.Open sSQL, Application("DSN"), 3, 3

     if not oCheckForDMData.eof then
        lcl_return = false
     end if

     oCheckForDMData.close
     set oCheckForDMData = nothing

  end if

  checkForDMDataByDMTypeID = lcl_return

end function

'------------------------------------------------------------------------------
function checkForDefaultCategoryOnDMTypes(iCategoryID)
  dim lcl_return

  lcl_return = False

  if iCategoryID <> "" then
     sSQL = "SELECT DISTINCT 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_dm_types "
     sSQL = sSQL & " WHERE defaultcategoryid = " & iCategoryID

    	set oCheckForDefaultCat = Server.CreateObject("ADODB.Recordset")
   	 oCheckForDefaultCat.Open sSQL, Application("DSN"), 3, 3

     if not oCheckForDefaultCat.eof then
        lcl_return = True
     end if

     oCheckForDefaultCat.close
     set oCheckForDefaultCat = nothing

  end if

  checkForDefaultCategoryOnDMTypes = lcl_return

end function

'------------------------------------------------------------------------------
'sub maintainMPTSection(iDMTypeID, iOrgID, iDMSectionID, iSectionID, iSectionLocation, _
sub maintainDMTSection(iDMTypeID, iOrgID, iDMSectionID, iSectionID, iSectionLocation, _
                       iSectionOrder, iSectionActive)

  dim sDMTypeID, sOrgID, sDMSectionID, sSectionID, sSectionLocation
  dim sSectionOrder, sSectionActive, sSQL

  sDMTypeID        = 0
  sOrgID           = iOrgID
  sDMSectionID     = 0
  sSectionID       = 0
  sSectionLocation = "L"
  sSectionOrder    = 1 
  sSectionActive   = 0
  sSQL             = ""

  if iDMTypeID <> "" then
     sDMTypeID = CLng(iDMTypeID)
  end if

  if iDMSectionID <> "" then
     sDMSectionID = iDMSectionID
  end if

  if iSectionID <> "" then
     sSectionID = CLng(iSectionID)
  end if

  if iSectionLocation <> "" then
     sSectionLocation = ucase(iSectionLocation)
     sSectionLocation = "'" & dbsafe(sSectionLocation) & "'"
  end if

  if iSectionOrder <> "" then
     sSectionOrder = CLng(iSectionOrder)
  end if

  if iSectionActive = "Y" then
     sSectionActive = 1
  end if

  if sDMTypeID > 0 AND sSectionID > 0 then
    'Check to see if the section already exists on the DMType.
    'If "yes" then perform an "update".
    'If "no" then perform an "insert".
     if sDMSectionID = 0 then
        sDMSectionID = getDMSectionID(iDMTypeID, sDMSectionID, sSectionID)
     end if

     if sDMSectionID > 0 then
        sSQL = "UPDATE egov_dm_types_sections SET "
        sSQL = sSQL & " sectionlocation = " & sSectionLocation & ", "
        sSQL = sSQL & " sectionorder = "    & sSectionOrder    & ", "
        sSQL = sSQL & " isActive = "        & sSectionActive
        sSQL = sSQL & " WHERE dm_sectionid = " & sDMSectionID

       	set oUpdateDMTSection = Server.CreateObject("ADODB.Recordset")
      	 oUpdateDMTSection.Open sSQL, Application("DSN"), 3, 1

        set oUpdateDMTSection = nothing

       'Insert/Update any section fields associated to this section.
       'This will keep all of the section fields up-to-date
        'if iIsActiveByOrg = "Y" then
        if sSectionActive then
           maintainDMTSectionFields sDMTypeID, sOrgID, sDMSectionID, sSectionID, sSectionActive
        end if
     else
       'ONLY insert the section if it is "active"
        'if iIsActiveByOrg = "Y" then
        if sSectionActive then
           sSQL = "INSERT INTO egov_dm_types_sections ("
           sSQL = sSQL & "dm_typeid, "
           sSQL = sSQL & "orgid, "
           sSQL = sSQL & "sectionid, "
           sSQL = sSQL & "sectionlocation, "
           sSQL = sSQL & "sectionorder, "
           sSQL = sSQL & "isActive "
           sSQL = sSQL & ") VALUES ("
           sSQL = sSQL & sDMTypeID  & ", "
           sSQL = sSQL & sOrgID           & ", "
           sSQL = sSQL & sSectionID       & ", "
           sSQL = sSQL & sSectionLocation & ", "
           sSQL = sSQL & sSectionOrder    & ", "
           sSQL = sSQL & sSectionActive
           sSQL = sSQL & ") "

           lcl_dm_sectionid = RunIdentityInsert(sSQL)

          'Add any/all section field(s) to the DMType that belong to this section
           insertDMTSectionFields sDMTypeID, sOrgID, lcl_dm_sectionid

        end if
     end if
  end if

end sub

'------------------------------------------------------------------------------
'sub updateMapPointTypeLayout(iMapPointTypeID, iLayoutID)
sub updateDMTypeLayout(iDMTypeID, iLayoutID)

  dim sDMTypeID, sLayoutID

  if iDMTypeID <> "" then
     sDMTypeID = CLng(iDMTypeID)
  else
     sDMTypeID = 0
  end if

  if iLayoutID <> "" then
     sLayoutID = CLng(iLayoutID)
  else
     sLayoutID = getOriginalLayoutID()
  end if

 	sSQL = "UPDATE egov_dm_types SET "
  sSQL = sSQL & "layoutid = " & sLayoutID
  sSQL = sSQL & " WHERE dm_typeid = " & sDMTypeID

 	set oUpdateDMTLayout = Server.CreateObject("ADODB.Recordset")
	 oUpdateDMTLayout.Open sSQL, Application("DSN"), 3, 1

  set oUpdateDMTLayout = nothing

  'lcl_redirect_url = "mappoints_types_layout_maint.asp"
  'lcl_redirect_url = lcl_redirect_url & "?dm_typeid=" & sDMTypeID
  'lcl_redirect_url = lcl_redirect_url & "&layoutid=" & sLayoutID
  'lcl_redirect_url = lcl_redirect_url & "&success=SU"

end sub

'------------------------------------------------------------------------------
'function getMPSectionID(iDMTypeID, sDMSectionID, sSectionID)
function getDMSectionID(iDMTypeID, iDMSectionID, iSectionID)

  dim lcl_return

  lcl_return = 0

  sSQL = "SELECT dm_sectionid "
  sSQL = sSQL & " FROM egov_dm_types_sections "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

  if iDMSectionID > 0 then
     sSQL = sSQL & " AND dm_sectionid = " & iDMSectionID
  else
     sSQL = sSQL & " AND sectionid = " & iSectionID
  end if

 	set oGetDMSectionID = Server.CreateObject("ADODB.Recordset")
	 oGetDMSectionID.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMSectionID.eof then
     lcl_return = oGetDMSectionID("dm_sectionid")
  end if

  oGetDMSectionID.close
  set oGetDMSectionID = nothing

  getDMSectionID = lcl_return

end function

'------------------------------------------------------------------------------
'sub maintainMPTSectionFields(iDMTypeID, iOrgID, iDMSectionID, iSectionID, iSectionActive)
sub maintainDMTSectionFields(iDMTypeID, iOrgID, iDMSectionID, iSectionID, iSectionActive)

 'Loop through all of the section fields and determine if any/all need to be inserted.
 'ONLY insert records if the section is "active".
 'This query gives a side-by-side comparison of all of the section fields associated to the section
 '  and all of the section fields associated to a DM Type.  Records WITHOUT a value in the 
 '  "dmtf.section_fieldid" are ones that the section field has NOT been associated to the DM Type.
  sSQL = "SELECT sf.section_fieldid, "
  sSQL = sSQL & " sf.fieldname, "
  sSQL = sSQL & " sf.isActive, "
  sSQL = sSQL & " sf.displayOrder, "
  sSQL = ssQL & " dmts.dm_sectionid "
  sSQL = sSQL & " FROM egov_dm_sections_fields sf "
  sSQL = sSQL & "      INNER JOIN egov_dm_sections s "
  sSQL = sSQL & "              ON sf.sectionid = s.sectionid "
  sSQL = sSQL & "             AND s.sectionid = " & iSectionID
  sSQL = sSQL & "             AND s.isActive = 1 "
  sSQL = sSQL & "      INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL & "              ON sf.sectionid = dmts.sectionid "
  sSQL = sSQL & "             AND dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL & "             AND dmts.isActive = 1 "
  sSQL = sSQL & " WHERE sf.isActive = 1 "
  sSQL = sSQL & "   AND sf.section_fieldid NOT IN (select dmtf.section_fieldid "
  sSQL = sSQL & "                                  from egov_dm_types_fields dmtf, "
  sSQL = sSQL & "                                       egov_dm_types_sections dmts2 "
  sSQL = sSQL & "                                  where dmtf.dm_sectionid = dmts2.dm_sectionid "
  sSQL = sSQL & "                                  and dmts2.orgid = " & iOrgID
  sSQL = sSQL & "                                  and dmts2.dm_typeid = " & iDMTypeID
  sSQL = sSQL & "                                  and dmts2.sectionid = " & iSectionID
  sSQL = sSQL & "                                  and dmts2.isActive = 1 "
  sSQL = sSQL & "                                 ) "
  sSQL = sSQL & " ORDER BY sf.displayOrder "

  set oGetDMTSectionFields = Server.CreateObject("ADODB.Recordset")
  oGetDMTSectionFields.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMTSectionFields.eof then
     do while not oGetDMTSectionFields.eof

        sSQL = "INSERT INTO egov_dm_types_fields ("
        sSQL = sSQL & "dm_typeid, "
        sSQL = sSQL & "orgid, "
        sSQL = sSQL & "dm_sectionid, "
        sSQL = sSQL & "section_fieldid, "
        sSQL = sSQL & "displayInResults, "
        sSQL = sSQL & "displayInInfoPage, "
        sSQL = sSQL & "resultsOrder, "
        sSQL = sSQL & "inPublicSearch "
        sSQL = sSQL & ") VALUES ("
        sSQL = sSQL & iDMTypeID                               & ", "
        sSQL = sSQL & iOrgID                                  & ", "
        'sSQL = sSQL & iDMSectionID & ", "
        sSQL = sSQL & oGetDMTSectionFields("dm_sectionid")    & ", "
        sSQL = sSQL & oGetDMTSectionFields("section_fieldid") & ", "
        sSQL = sSQL & "1, "
        sSQL = sSQL & "1, "
        sSQL = sSQL & oGetDMTSectionFields("displayOrder")    & ", "
        sSQL = sSQL & "1 "
        sSQL = sSQL & ") "

        set oMaintainDMTSectionField = Server.CreateObject("ADODB.Recordset")
        oMaintainDMTSectionField.Open sSQL, Application("DSN"), 3, 1

        set oMaintainDMTSectionField = nothing

        oGetDMTSectionFields.movenext
     loop
  end if

  oGetDMTSectionFields.close
  set oGetDMTSectionFields = nothing

end sub

'------------------------------------------------------------------------------
'sub insertMPTSectionFields(iDMTypeID, iOrgID, iDMSectionID)
sub insertDMTSectionFields(iDMTypeID, iOrgID, iDMSectionID)

 'Get all of the "active" fields for the section
  sSQL = "SELECT section_fieldid "
  sSQL = sSQL & " FROM egov_dm_sections_fields "
  sSQL = sSQL & " WHERE sectionid IN (select sectionid "
  sSQL = sSQL &                     " from egov_dm_types_sections "
  sSQL = sSQL &                     " where dm_sectionid = " & iDMSectionID & ") "
  sSQL = sSQL & " AND isActive = 1 "
  sSQL = sSQL & " ORDER BY displayOrder "

  set oGetSectionFields = Server.CreateObject("ADODB.Recordset")
  oGetSectionFields.Open sSQL, Application("DSN"), 3, 1

  if not oGetSectionFields.eof then
     do while not oGetSectionFields.eof

        sSQLi = "INSERT INTO egov_dm_types_fields ("
        sSQLi = sSQLi & "dm_typeid,"
        sSQLi = sSQLi & "orgid,"
        sSQLi = sSQLi & "dm_sectionid,"
        sSQLi = sSQLi & "section_fieldid,"
        sSQLi = sSQLi & "displayInResults,"
        sSQLi = sSQLi & "displayInInfoPage,"
        sSQLi = sSQLi & "resultsOrder,"
        sSQLi = sSQLi & "inPublicSearch"
        sSQLi = sSQLi & ") VALUES ("
        sSQLi = sSQLi & iDMTypeID & ", "
        sSQLi = sSQLi & iOrgID          & ", "
        sSQLi = sSQLi & iDMSectionID    & ", "
        sSQLi = sSQLi & oGetSectionFields("section_fieldid") & ", "
        sSQLi = sSQLi & "1, 1, 1, 1 "
        sSQLi = sSQLi & ") "

        lcl_dm_fieldid = RunIdentityInsert(sSQLi)

        oGetSectionFields.movenext
     loop
  end if

  oGetSectionFields.close
  set oGetSectionFields = nothing

end sub

'------------------------------------------------------------------------------
sub buildDMTypeFromTemplate(iOrgID, iDMTypeID, iDMTemplateID)

  dim lcl_dm_sectionid, lcl_sectionactive, lcl_dmid, lcl_catid

  if iDMTypeID <> "" AND iDMTemplateID <> "" then

    'BEGIN: Get any/all sections associated to the DM Type Template -----------
     sSQL = "SELECT dmts.dm_sectionid, "
     sSQL = sSQL & " dmts.sectionid, "
     sSQL = sSQL & " dmts.sectionlocation, "
     sSQL = sSQL & " dmts.sectionorder, "
     sSQL = sSQL & " dmts.isActive "
     sSQL = sSQL & " FROM egov_dm_types_sections dmts "
     sSQL = sSQL & " WHERE dmts.dm_typeid = " & iDMTemplateID
     sSQL = sSQL & " ORDER BY dmts.sectionorder "

     set oGetDMTSections = Server.CreateObject("ADODB.Recordset")
     oGetDMTSections.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTSections.eof then
        do while not oGetDMTSections.eof

          'Insert the template section for the the DM TypeID.
          'NOTE: If/When a section is added the "maintainDMTSection" procedure will automatically 
          'add to the DM Type any fields associated with the section being inserted.
           lcl_dm_sectionid  = ""
           lcl_sectionactive = ""

           if oGetDMTSections("isActive") then
              lcl_sectionactive = "Y"
           end if

           maintainDMTSection iDMTypeID, iOrgID, lcl_dm_sectionid, oGetDMTSections("sectionid"), _
                              oGetDMTSections("sectionlocation"), oGetDMTSections("sectionorder"), _
                              lcl_sectionactive

           oGetDMTSections.movenext
        loop
     end if

     oGetDMTSections.close
     set oGetDMTSections = nothing
    'END: Get any/all sections associated to the DM Type Template -------------

    'BEGIN: Get any/all categories associated to the DM Type Template ---------
    'First pull all of the parent categories
     sSQL = "SELECT categoryid, "
     sSQL = sSQL & " categoryname, "
     sSQL = sSQL & " isActive, "
     sSQL = sSQL & " parent_categoryid, "
     sSQL = sSQL & " isApproved, "
     sSQL = sSQL & " approvedeniedbyid, "
     sSQL = sSQL & " approvedeniedbydate, "
     sSQL = sSQL & " mappointcolor "
     sSQL = sSQL & " FROM egov_dm_categories "
     sSQL = sSQL & " WHERE dm_typeid = " & iDMTemplateID
     sSQL = sSQL & " AND isActive = 1 "
     sSQL = sSQL & " AND parent_categoryid = 0 "

     set oGetTemplateCategories = Server.CreateObject("ADODB.Recordset")
     oGetTemplateCategories.Open sSQL, Application("DSN"), 3, 1

     if not oGetTemplateCategories.eof then
        do while not oGetTemplateCategories.eof

           lcl_catid         = 0
           lcl_new_catid     = 0
           lcl_new_subcatid  = 0
           lcl_dmid          = 0
           lcl_delete        = ""
           lcl_mergecat      = ""
           lcl_assignsubcat  = ""
           lcl_cat_active    = ""
           lcl_subcat_active = ""

           if oGetTemplateCategories("isActive") then
              lcl_cat_active = "Y"
           end if

           lcl_new_catid = maintainSubCategory(iOrgID, iDMTypeID, lcl_dmid, session("userid"), lcl_delete, lcl_mergecat, _
                                               lcl_catid, oGetTemplateCategories("categoryname"), lcl_cat_active, _
                                               oGetTemplateCategories("parent_categoryid"), lcl_assignsubcat)

          'Now see if there are any sub-categories associated to parent category
           sSQLsc = "SELECT categoryid, "
           sSQLsc = sSQLsc & " categoryname, "
           sSQLsc = sSQLsc & " isActive, "
           sSQLsc = sSQLsc & " parent_categoryid, "
           sSQLsc = sSQLsc & " isApproved, "
           sSQLsc = sSQLsc & " approvedeniedbyid, "
           sSQLsc = sSQLsc & " approvedeniedbydate, "
           sSQLsc = sSQLsc & " mappointcolor "
           sSQLsc = sSQLsc & " FROM egov_dm_categories "
           sSQLsc = sSQLsc & " WHERE dm_typeid = " & iDMTemplateID
           sSQLsc = sSQLsc & " AND isActive = 1 "
           sSQLsc = sSQLsc & " AND parent_categoryid = " & oGetTemplateCategories("categoryid")

           set oGetTemplateSubCategories = Server.CreateObject("ADODB.Recordset")
           oGetTemplateSubCategories.Open sSQLsc, Application("DSN"), 3, 1

           if not oGetTemplateSubCategories.eof then
              do while not oGetTemplateSubCategories.eof
                 if oGetTemplateSubCategories("isActive") then
                    lcl_subcat_active = "Y"
                 end if

                 lcl_new_subcatid = maintainSubCategory(iOrgID, iDMTypeID, lcl_dmid, session("userid"), lcl_delete, _
                                                        lcl_mergecat, lcl_catid, oGetTemplateSubCategories("categoryname"), _
                                                        lcl_subcat_active, lcl_new_catid, lcl_assignsubcat)

                 oGetTemplateSubCategories.movenext
              loop
           end if

           set oGetTemplateSubCategories = nothing

           oGetTemplateCategories.movenext
        loop
     end if

     oGetTemplateCategories.close
     set oGetTemplateCategories = nothing


'LEFT OFF HERE!!!
'1. pull all of the parent categories from the template
'2. pull all of the sub-categories for each parent category from the template
'** NO assignments needed. we just need to pull the values.


    'END: Get any/all categories associated to the DM Type Template -----------
  end if

end sub

'------------------------------------------------------------------------------
function DisplayAddress(p_orgid, p_street_number, p_street_name)
	
	dim sNumber, oAddressList, blnFound
 dim lcl_streetnumber, lcl_streetname, lcl_new_street_name

 lcl_streetnumber    = p_street_number
 lcl_streetname      = p_street_name
 lcl_new_street_name = buildStreetAddress(lcl_streetnumber, "", lcl_streetname, "", "")

'Get list of addresses
	sSQL = "SELECT residentaddressid, "
 sSQL = sSQL & " isnull(residentstreetnumber,'') as residentstreetnumber, "
 sSQL = sSQL & " residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
 sSQL = sSQL & " streetsuffix, "
 sSQL = sSQL & " streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, "
 sSQL = sSQL & " isnull(longitude,0.00) as longitude "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid=" & p_orgid
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND residentstreetname is not null "
 sSQL = sSQL & " ORDER BY sortstreetname, residentstreetprefix"

	set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

 if not oAddressList.eof then
    response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value="""" />" & vbcrlf
   	'response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkImportAddressBtn();"">" & vbcrlf
   	response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
   	response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
	
   	do while not oAddressList.eof
      'Build the original full street address
       lcl_original_street_name = buildStreetAddress(oAddressList("residentstreetnumber"), _
                                                     oAddressList("residentstreetprefix"), _
                                                     oAddressList("residentstreetname"), _
                                                     oAddressList("streetsuffix"), _
                                                     oAddressList("streetdirection")_
                                                    )

       if UCASE(lcl_streetname) = UCASE(oAddressList("residentstreetname")) then
          sSelected = " selected=""selected"""
       else
          sSelected = ""
       end if

      	response.write "  <option value=""" & oAddressList("residentaddressid") & """" & sSelected & ">" & lcl_original_street_name & "</option>" & vbcrlf

  	   	oAddressList.MoveNext
   	loop

   	response.write "</select>" & vbcrlf
 else
response.write "small address list<br />" & vbcrlf
    response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" />" & vbcrlf
    response.write "<input type=""hidden"" name=""streetaddress"" id=""streetaddress"" value=""0000"" />" & vbcrlf
   	'response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
   	'response.write "  <option value=""0000"">No addresses available</option>" & vbcrlf
    'response.write "</select>" & vbcrlf
 end if

	oAddressList.close
	set oAddressList = nothing

end function

'------------------------------------------------------------------------------
sub displayLargeAddressList(p_orgid, p_street_number, p_prefix, p_street_name, p_suffix, p_direction)
 dim sSql, oAddressList
 dim lcl_streetnumber, lcl_prefix, lcl_streetname, lcl_suffix, lcl_direction, lcl_compare_address

 lcl_streetnumber    = p_street_number
 lcl_prefix          = p_prefix
 lcl_streetname      = p_street_name
 lcl_suffix          = p_suffix
 lcl_direction       = p_direction
 lcl_compare_address = buildStreetAddress("", lcl_prefix, lcl_streetname, lcl_suffix, lcl_direction)

	sSQL = "SELECT DISTINCT sortstreetname, "
 sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, "
 sSQL = sSQL & " ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " ORDER BY sortstreetname "
	
 set oAddressList = Server.CreateObject("ADODB.Recordset")
 oAddressList.Open sSQL, Application("DSN"), 3, 1

 if not oAddressList.eof then
  		'response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" onchange=""save_address();"" /> &nbsp; " & vbcrlf
 	 	'response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkAddressButtons()"">" & vbcrlf
  		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
 	 	response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
  		response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

    do while not oAddressList.eof

      'Build the full street address
       lcl_streetaddress = buildStreetAddress("", oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

      'Determine if the option is selected
     		if UCASE(lcl_streetaddress) = UCASE(lcl_compare_address) then
       			lcl_selected_address = " selected=""selected"""
       else
          lcl_selected_address = ""
     		end if

       response.write "<option value=""" & lcl_streetaddress & """" & lcl_selected_address & ">" & lcl_streetaddress & "</option>" & vbcrlf

    			oAddressList.MoveNext
    loop

    response.write "</select>&nbsp;" & vbcrlf
    response.write "<input type=""button"" id=""validateAddress"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
 else
    'response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""font-size:9pt"">" & vbcrlf
    'response.write "  <tr>" & vbcrlf
    'response.write "      <td>Street Number:</td>" & vbcrlf
    'response.write "      <td><input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /></td>" & vbcrlf
 	 	'response.write "  </tr>" & vbcrlf
    'response.write "  <tr>" & vbcrlf
    'response.write "      <td>Street Name:</td>" & vbcrlf
    'response.write "      <td><input type=""text"" name=""streetaddress"" id=""streetaddress"" size=""30"" maxlength=""84"" /></td>" & vbcrlf
    'response.write "  </tr>" & vbcrlf
    'response.write "</table>" & vbcrlf
response.write "large address list<br />" & vbcrlf
  		response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
    response.write "<input type=""hidden"" name=""streetaddress"" id=""streetaddress"" value=""0000"" />" & vbcrlf
 	 	'response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
  		'response.write "  <option value=""0000"">No addresses available</option>" & vbcrlf
    'response.write "</select>" & vbcrlf

 end if

 oAddressList.close
 set oAddressList = nothing

end sub

'------------------------------------------------------------------------------
function formatFieldforInsertUpdate(p_value)
  dim lcl_return

  lcl_return = "NULL"

  if trim(p_value) <> "" then
     lcl_return = "'" & dbsafe(p_value) & "'"
  end if

  formatFieldforInsertUpdate = lcl_return

end function

'------------------------------------------------------------------------------
'function getMapPointTypeDescription(iDMTypeID)
function getDMTypeDescription(iDMTypeID)

  dim lcl_return, lcl_dmtypeid

  lcl_return   = ""
  lcl_dmtypeid = 0

  if iDMTypeID <> "" then
     lcl_dmtypeid = CLng(iDMTypeID)
  end if

  sSQL = "SELECT description "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE dm_typeid = " & lcl_dmtypeid

  set oGetDMTDesc = Server.CreateObject("ADODB.Recordset")
  oGetDMTDesc.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMTDesc.eof then
     lcl_return = oGetDMTDesc("description")
  end if

  oGetDMTDesc.close
  set oGetDMTDesc = nothing

  getDMTypeDescription = lcl_return

end function

'------------------------------------------------------------------------------
'sub displayMapPointTypes(iOrgID, iDMTypeID, iFeature)
sub displayDMTypes(iOrgID, iDMTypeID, iFeature)

  dim lcl_dmtypeid

  if iDMTypeID <> "" then
     lcl_dmtypeid = CLng(iDMTypeID)
  else
     lcl_dmtypeid = 0
  end if

  sSQL = "SELECT dm_typeid, description "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND isActive = 1 "

  if iFeature <> "" AND iFeature <> "datamgr_maint" AND iFeature <> "datamgr_owners" then
     sSQL = sSQL & " AND UPPER(feature_maintain) = '" & UCASE(iFeature) & "' "
  end if

  sSQL = sSQL & " ORDER BY description "

  set oDisplayDMTypes = Server.CreateObject("ADODB.Recordset")
  oDisplayDMTypes.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayDMTypes.eof then
     do while not oDisplayDMTypes.eof

        if oDisplayDMTypes("dm_typeid") = lcl_dmtypeid then
           lcl_selected_dmtype = " selected=""selected"""
        else
           lcl_selected_dmtype = ""
        end if

        response.write "  <option value=""" & oDisplayDMTypes("dm_typeid") & """" & lcl_selected_dmtype & ">" & oDisplayDMTypes("description") & "</option>" & vbcrlf

        oDisplayDMTypes.movenext
     loop
  end if

end sub

'------------------------------------------------------------------------------
'sub maintainMapPointValues(iUserID, iOrgID, iDMTypeID, iDMID, iDMSectionID, iDMFieldID, iDMValueID, iFieldValue)
sub maintainDMValues(iUserID, _
                     iOrgID, _
                     iDMTypeID, _
                     iDMID, _
                     iDMSectionID, _
                     iDMFieldID, _
                     iDMValueID, _
                     iFieldValue, _
                     iMPValueID, _
                     iDMImportID)

  dim lcl_userid, lcl_orgid, lcl_dm_typeid, lcl_dm_sectionid
  dim lcl_dm_fieldid, lcl_dm_valueid, lcl_fieldvalue
  dim lcl_mp_valueid, lcl_dm_importid

  lcl_userid       = iUserID
  lcl_orgid        = iOrgID
  lcl_dm_typeid    = iDMTypeID
  lcl_dmid         = iDMID
  lcl_dm_sectionid = iDMSectionID
  lcl_dm_fieldid   = iDMFieldID
  lcl_dm_valueid   = iDMValueID
  lcl_fieldvalue   = formatFieldforInsertUpdate(iFieldValue)
  lcl_mp_valueid   = iMPValueID
  lcl_dm_importid  = iDMImportID

 'If a dm_valueid exists then update the DM Data Value.  Otherwise, insert it.
  'lcl_dm_valueid = getDMValueID(lcl_dm_valueid, lcl_dm_typeid, lcl_dmid, lcl_dm_sectionid, lcl_dm_fieldid)
  lcl_dm_valueid = getDMValueID(lcl_orgid, _
                                lcl_dm_typeid, _
                                lcl_dmid, _
                                lcl_dm_sectionid, _
                                lcl_dm_fieldid)

  if lcl_dm_valueid = "" then
     lcl_dm_valueid = 0
  end if

  if lcl_mp_valueid = "" then
     lcl_mp_valueid = 0
  end if

  if lcl_dm_importid = "" then
     lcl_dm_importid = 0
  end if

  if lcl_dm_valueid > 0 then
     sSQL = "UPDATE egov_dm_values SET "
     sSQL = sSQL & " fieldvalue = " & lcl_fieldvalue

     if lcl_mp_valueid <> "" then
        if lcl_mp_valueid > 0 then
           sSQL = sSQL & ", mp_valueid = " & lcl_mp_valueid
        end if
     end if

     if lcl_dm_importid <> "" then
        if lcl_dm_importid > 0 then
           sSQL = sSQL & ", dm_importid = " & lcl_dm_importid
        end if
     end if

     sSQL = sSQL & " WHERE dm_valueid = " & lcl_dm_valueid

     set oMaintainDMValues = Server.CreateObject("ADODB.Recordset")
     oMaintainDMValues.Open sSQL, Application("DSN"), 3, 1

     set oMaintainDMValues = nothing

  else
     sSQL = "INSERT INTO egov_dm_values ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "dm_typeid, "
     sSQL = sSQL & "dmid, "
     sSQL = sSQL & "dm_sectionid, "
     sSQL = sSQL & "dm_fieldid, "
     sSQL = sSQL & "fieldvalue, "
     sSQL = sSQL & "mp_valueid, "
     sSQL = sSQL & "dm_importid "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & lcl_orgid        & ", "
     sSQL = sSQL & lcl_dm_typeid    & ", "
     sSQL = sSQL & lcl_dmid         & ", "
     sSQL = sSQL & lcl_dm_sectionid & ", "
     sSQL = sSQL & lcl_dm_fieldid   & ", "
     sSQL = sSQL & lcl_fieldvalue   & ", "
     sSQL = sSQL & lcl_mp_valueid   & ", "
     sSQL = sSQL & lcl_dm_importid
     sSQL = sSQL & ")"

  		'Get the DMID
 	  	lcl_dm_valueid = RunIdentityInsert(sSQL)
  end if
end sub

'------------------------------------------------------------------------------
'sub maintainMapPoint(ByVal iUserID, ByVal iOrgID, ByVal iDMID, ByVal iDMTypeID, ByVal iDMSectionID, _
sub maintainDMData(ByVal iUserID, ByVal iOrgID, ByVal iDMID, ByVal iDMTypeID, ByVal iDMSectionID, _
                   ByVal iDMFieldID, ByVal iIsActive, ByVal iCategoryID, ByVal iDMImportID, ByRef lcl_dmid)

  dim sUserID, sDMID, sDMTypeID, sDMSectionID, sDMFieldID, sDMImportID
  dim sCategoryID, sIsActive, lcl_current_date

  sUserID      = iUserID
  sDMID        = 0
  sDMTypeID    = ""
  sDMSectionID = ""
  sDMFieldID   = ""
  sCategoryID  = ""
  sDMImportID  = "NULL"
  sIsActive    = ""
  lcl_dmid     = 0
  lcl_current_date = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iDMSectionID <> "" then
     sDMSectionID = clng(iDMSectionID)
  end if

  if iDMFieldID <> "" then
     sDMFieldID = clng(iDMFieldID)
  end if

  if iCategoryID <> "" then
     sCategoryID = clng(iCategoryID)
  end if

  if iDMImportID <> "" then
     sDMImportID = clng(iDMImportID)
  end if

  if iIsActive <> "" then
     sIsActive = iIsActive
  end if

 'BEGIN: Update ---------------------------------------------------------------
  if sDMID > 0 then

   		sSQL = "UPDATE egov_dm_data SET "
     sSQL = sSQL & "lastmodifiedbyid = "   & sUserID          & ", "
     sSQL = sSQL & "lastmodifiedbydate = " & lcl_current_date & ", "
     sSQL = sSQL & "dm_typeid = "          & sDMTypeID        & ", "
     sSQL = sSQL & "categoryid = "         & sCategoryID      & ", "
     sSQL = sSQL & "isActive = "           & sIsActive
     sSQL = sSQL & " WHERE dmid = " & sDMID

   		set oUpdateDMData = Server.CreateObject("ADODB.Recordset")
    	oUpdateDMData.Open sSQL, Application("DSN"), 3, 1

     lcl_dmid = sDMID

 'BEGIN: Insert ---------------------------------------------------------------
  else
     if sDMTypeID = "" then
        sDMID = 0
     end if

     if sDMSectionID = "" then
        sDMSectionID = 0
     end if

     if sDMFieldID = "" then
        sDMFieldID = 0
     end if

     if sCategoryID = "" then
        sCategoryID = 0
     end if

     if sIsActive = "" then
        sIsActive = 1
     end if

     sCreatedByID          = sUserID
     sCreatedByDate        = lcl_current_date
     sApprovedDeniedByID   = sUserID
     sApprovedDeniedByDate = lcl_current_date

    'Insert the new Map-Point
     sSQL = "INSERT INTO egov_dm_data ("
     sSQL = sSQL & "dm_typeid, "
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "categoryid, "
     sSQL = sSQL & "createdbyid, "
     sSQL = sSQL & "createdbydate, "
     sSQL = sSQL & "isApproved, "
     sSQL = sSQL & "approvedeniedbyid, "
     sSQL = sSQL & "approvedeniedbydate, "
     sSQL = sSQL & "lastmodifiedbyid, "
     sSQL = sSQL & "lastmodifiedbydate, "
     sSQL = sSQL & "isActive, "
     sSQL = sSQL & "dm_importid "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & sDMTypeID  & ", "
     sSQL = sSQL & iOrgID                & ", "
     sSQL = sSQL & sCategoryID           & ", "
     sSQL = sSQL & sCreatedByID          & ", "
     sSQL = sSQL & sCreatedByDate        & ", "
     sSQL = sSQL & "1, "
     sSQL = sSQL & sApprovedDeniedByID   & ", "
     sSQL = sSQL & sApprovedDeniedByDate & ", "
     sSQL = sSQL & "NULL,NULL"           & ", "
     sSQL = sSQL & sIsActive             & ", "
     sSQL = sSQL & sDMImportID
     sSQL = sSQL & ")"

    'Get the DMID
     lcl_dmid = RunIdentityInsert(sSQL)
  end if

end sub

'------------------------------------------------------------------------------
'sub maintainMPSection_address(ByVal iUserID, ByVal iOrgID, ByVal iDMID, ByVal iDMTypeID, ByVal iDMSectionID, _
sub maintainDMSection_address(ByVal iUserID, ByVal iOrgID, ByVal iDMID, ByVal iDMTypeID, ByVal iDMSectionID, _
                              ByVal iDMFieldID, ByVal iStreetNumber, ByVal iStreetPrefix, ByVal iStreetAddress, _
                              ByVal iStreetSuffix, ByVal iStreetDirection, ByVal iSortStreetName, ByVal iCity, _
                              ByVal iState, ByVal iZip, ByVal iValidStreet, ByVal iLatitude, ByVal iLongitude, _
                              ByRef lcl_dmid)

  dim sUserID, sDMID, sDMTypeID, sDMSectionID, lcl_current_date
  dim sSortStreetName, sNumber, sPrefix, sAddress, sSuffix, sDirection
  dim sValidStreet, sCity, sState, sZip, sMapPointColor, sLatitude, sLongitude

  sUserID          = iUserID
  sDMID            = 0
  sDMTypeID        = ""
  sDMSectionID     = ""
  sDMFieldID       = ""
  lcl_dmid         = 0
  lcl_current_date = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"

  if iDMID <> "" then
     sDMID = iDMID
  end if

  if iDMTypeID <> "" then
     sDMTypeID = iDMTypeID
  end if

  if iDMSectionID <> "" then
     sDMSectionID = iDMSectionID
  end if

  if iDMFieldID <> "" then
     sDMFieldID = iDMFieldID
  end if

 'BEGIN: Format the columns for the table ----------------------------------
  sSortStreetName = formatFieldforInsertUpdate(iSortStreetName)
  sNumber         = formatFieldforInsertUpdate(iStreetNumber)
  sPrefix         = formatFieldforInsertUpdate(iStreetPrefix)
  sAddress        = formatFieldforInsertUpdate(iStreetAddress)
  sSuffix         = formatFieldforInsertUpdate(iStreetSuffix)
  sDirection      = formatFieldforInsertUpdate(iStreetDirection)
  sValidStreet    = formatFieldforInsertUpdate(iValidStreet)
  sCity           = formatFieldforInsertUpdate(iCity)
  sState          = formatFieldforInsertUpdate(iState)
  sZip            = formatFieldforInsertUpdate(iZip)
  sMapPointColor  = formatFieldforInsertUpdate(lcl_mappointcolor)
  sLatitude       = 0.00
  sLongitude      = 0.00
 'END: Format the columns for the table ------------------------------------

 'BEGIN: Latitude ----------------------------------------------------------
  if iLatitude <> "" then
     sLatitude = iLatitude
  end if
 'END: Latitude ------------------------------------------------------------

 'BEGIN: Longitude ---------------------------------------------------------
  if iLongitude <> "" then
     sLongitude = iLongitude
  end if
 'END: Longitude -----------------------------------------------------------

 'BEGIN: Update ---------------------------------------------------------------
  if sDMID > 0 then

   		sSQLdu = "UPDATE egov_dm_data SET "
     sSQLdu = sSQLdu & "lastmodifiedbyid = "   & sUserID          & ", "
     sSQLdu = sSQLdu & "lastmodifiedbydate = " & lcl_current_date & ", "
     sSQLdu = sSQLdu & "dm_typeid = "          & sDMTypeID        & ", "
     sSQLdu = sSQLdu & "streetnumber = "       & sNumber          & ", "
     sSQLdu = sSQLdu & "streetprefix = "       & sPrefix          & ", "
     sSQLdu = sSQLdu & "streetaddress = "      & sAddress         & ", "
     sSQLdu = sSQLdu & "streetsuffix = "       & sSuffix          & ", "
     sSQLdu = sSQLdu & "streetdirection = "    & sDirection       & ", "
     sSQLdu = sSQLdu & "sortstreetname = "     & sSortStreetName  & ", "
     sSQLdu = sSQLdu & "city = "               & sCity            & ", "
     sSQLdu = sSQLdu & "state = "              & sState           & ", "
     sSQLdu = sSQLdu & "zip = "                & sZip             & ", "
     sSQLdu = sSQLdu & "validstreet = "        & sValidStreet     & ", "
     sSQLdu = sSQLdu & "latitude = "           & sLatitude        & ", "
     sSQLdu = sSQLdu & "longitude = "          & sLongitude
     sSQLdu = sSQLdu & " WHERE dmid = " & sDMID

   		'set oUpdateDMSectionAddress = Server.CreateObject("ADODB.Recordset")
    	'oUpdateDMSectionAddress.Open sSQLdu, Application("DSN"), 3, 1

     'set oUpdateDMSectionAddress = nothing

     RunSQLStatement sSQLdu

     lcl_dmid = sDMID

 'BEGIN: Insert ---------------------------------------------------------------
  else
     if sDMTypeID = "" then
        sDMID = 0
     end if

     if sDMSectionID = "" then
        sDMSectionID = 0
     end if

     if sDMFieldID = "" then
        sDMFieldID = 0
     end if

     sCreatedByID   = sUserID
     sCreatedByDate = lcl_current_date

    'Insert the new Map-Point
     sSQLdi = "INSERT INTO egov_dm_data ("
     sSQLdi = sSQLdi & "dm_typeid, "
     sSQLdi = sSQLdi & "orgid, "
     sSQLdi = sSQLdi & "createdbyid, "
     sSQLdi = sSQLdi & "createdbydate, "
     sSQLdi = sSQLdi & "lastmodifiedbyid, "
     sSQLdi = sSQLdi & "lastmodifiedbydate, "
     sSQLdi = sSQLdi & "streetnumber, "
     sSQLdi = sSQLdi & "streetprefix, "
     sSQLdi = sSQLdi & "streetaddress, "
     sSQLdi = sSQLdi & "streetsuffix, "
     sSQLdi = sSQLdi & "streetdirection, "
     sSQLdi = sSQLdi & "sortstreetname, "
     sSQLdi = sSQLdi & "city, "
     sSQLdi = sSQLdi & "state, "
     sSQLdi = sSQLdi & "zip, "
     sSQLdi = sSQLdi & "validstreet, "
     sSQLdi = sSQLdi & "latitude, "
     sSQLdi = sSQLdi & "longitude "
     sSQLdi = sSQLdi & ") VALUES ("
     sSQLdi = sSQLdi & sDMTypeID  & ", "
     sSQLdi = sSQLdi & iOrgID           & ", "
     sSQLdi = sSQLdi & sCreatedByID     & ", "
     sSQLdi = sSQLdi & sCreatedByDate   & ", "
     sSQLdi = sSQLdi & "NULL,NULL"      & ", "
     sSQLdi = sSQLdi & sStreetNumber    & ", "
     sSQLdi = sSQLdi & sStreetPrefix    & ", "
     sSQLdi = sSQLdi & sStreetAddress   & ", "
     sSQLdi = sSQLdi & sStreetSuffix    & ", "
     sSQLdi = sSQLdi & sStreetDirection & ", "
     sSQLdi = sSQLdi & sSortStreetName  & ", "
     sSQLdi = sSQLdi & sCity            & ", "
     sSQLdi = sSQLdi & sState           & ", "
     sSQLdi = sSQLdi & sZip             & ", "
     sSQLdi = sSQLdi & sValidStreet     & ", "
     sSQLdi = sSQLdi & sLatitude        & ", "
     sSQLdi = sSQLdi & sLongitude
     sSQLdi = sSQLdi & ")"

    'Get the DMID
     lcl_dmid = RunIdentityInsert(sSQLdi)
  end if

end sub

'------------------------------------------------------------------------------
function setupUrlParameters(iURLParameters, iFieldName, iFieldValue)
  dim lcl_return

  lcl_return = ""

  if trim(iURLParameters) <> "" then
     lcl_return = iURLParameters
  end if

  if iFieldValue <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & "&"
     else
        lcl_return = lcl_return & "?"
     end if

     lcl_return = lcl_return & iFieldName & "=" & iFieldValue

  end if

  setupUrlParameters = lcl_return

end function

'------------------------------------------------------------------------------
'sub updateMapPointValues(iDMTypeID)
sub updateDMValues(iDMTypeID)

  if iDMTypeID <> "" then
     sSQL = "SELECT dm_fieldid, fieldtype, fieldname, displayInResults, resultsOrder "
     sSQL = sSQL & " FROM egov_dm_types_fields "
     sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

     set oGetDMTFields = Server.CreateObject("ADODB.Recordset")
     oGetDMTFields.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTFields.eof then
        do while not oGetDMTFields.eof

           lcl_fieldtype        = "NULL"
           lcl_fieldname        = "NULL"
           lcl_displayInResults = 0
           lcl_resultsOrder     = 1

           if oGetDMTFields("fieldtype") <> "" then
              lcl_fieldtype = "'" & dbsafe(oGetDMTFields("fieldtype")) & "'"
           end if

           if oGetDMTFields("fieldname") <> "" then
              lcl_fieldname = "'" & dbsafe(oGetDMTFields("fieldname")) & "'"
           end if

           if oGetDMTFields("displayInResults") then
              lcl_displayInResults = 1
           end if

           if oGetDMTFields("resultsOrder") <> "" then
              lcl_resultsOrder = oGetDMTFields("resultsOrder")
           end if


           sSQL = "UPDATE egov_dm_values SET "
           sSQL = sSQL & " fieldtype = "        & lcl_fieldtype        & ", "
           sSQL = sSQL & " fieldname = "        & lcl_fieldname        & ", "
           sSQL = sSQL & " displayInResults = " & lcl_displayInResults & ", "
           sSQL = sSQL & " resultsOrder = "     & lcl_resultsOrder
           sSQL = sSQL & " WHERE dm_fieldid = " & oGetDMTFields("dm_fieldid")

           set oUpdateDMValues = Server.CreateObject("ADODB.Recordset")
           oUpdateDMValues.Open sSQL, Application("DSN"), 3, 1

           set oUpdateDMValues = nothing

           oGetDMTFields.movenext
        loop
     end if

     oGetDMTFields.close
     set oGetDMTFields = nothing

  end if

end sub

'------------------------------------------------------------------------------
'sub displayMPTCategories(iOrgID, iCategoryID)
sub displayDMTCategories(iOrgID, iDMTypeID, iParentCategoryID, iCategoryID)

  sSQL = "SELECT categoryid, categoryname "
  sSQL = sSQL & " FROM egov_dm_categories "
  sSQL = sSQL & " WHERE isActive = 1 "
  sSQL = sSQL & " AND orgid = " & iOrgID
  sSQL = sSQL & " AND dm_typeid = " & iDMTypeID
  sSQL = sSQL & " AND parent_categoryid = " & iParentCategoryID
  sSQL = sSQL & " ORDER BY upper(categoryname) "

  set oGetDMTCategories = Server.CreateObject("ADODB.Recordset")
  oGetDMTCategories.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMTCategories.eof then
     do while not oGetDMTCategories.eof
        sCategoryID           = oGetDMTCategories("categoryid")
        sCategoryName         = oGetDMTCategories("categoryname")
        lcl_selected_category = ""

        if iCategoryID = sCategoryID then
           lcl_selected_category = " selected=""selected"""
        end if

        response.write "  <option value=""" & sCategoryID & """" & lcl_selected_category & ">" & sCategoryName & "</option>" & vbcrlf

        oGetDMTCategories.movenext
     loop
  end if

  oGetDMTCategories.close
  set oGetDMTCategories = nothing

end sub

'------------------------------------------------------------------------------
sub displayAllCategoriesOptions(iOrgID, iDMTypeID, iParentCategoryID, iCategoryID)

  sSQL = "SELECT categoryid, "
  sSQL = sSQL & " categoryname "
  sSQL = sSQL & " FROM egov_dm_categories "
  sSQL = sSQL & " WHERE isActive = 1 "
  sSQL = sSQL & " AND orgid = " & iOrgID
  sSQL = sSQL & " AND dm_typeid = " & iDMTypeID
  sSQL = sSQL & " AND parent_categoryid = " & iParentCategoryID

  if iCategoryID > 0 then
     sSQL = sSQL & " AND categoryid <> " & iCategoryID
  end if

  sSQL = sSQL & " ORDER BY upper(categoryname) "

  set oGetAllCategories = Server.CreateObject("ADODB.Recordset")
  oGetAllCategories.Open sSQL, Application("DSN"), 3, 1

  if not oGetAllCategories.eof then
     do while not oGetAllCategories.eof
        sCategoryID           = oGetAllCategories("categoryid")
        sCategoryName         = oGetAllCategories("categoryname")
        lcl_categorytype      = "PC"  'PC: Parent Category - SC: Sub-Category
        lcl_selected_category = ""


       'If we are showing a sub-category option then indent the option
        if iParentCategoryID > 0 then
           lcl_categorytype = "SC"
           sCategoryName    = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sCategoryName
        end if

        if iCategoryID = sCategoryID then
           lcl_selected_category = " selected=""selected"""
        end if

        response.write "  <option value=""" & lcl_categorytype & "" & sCategoryID & """" & lcl_selected_category & ">" & sCategoryName & "</option>" & vbcrlf

        displayAllCategoriesOptions iOrgID, iDMTypeID, oGetAllCategories("categoryid"), iCategoryID

        oGetAllCategories.movenext
     loop
  end if

  oGetAllCategories.close
  set oGetAllCategories = nothing

end sub

'------------------------------------------------------------------------------
sub displayMapPointColors(iMapPointColor)

 'Determine which color is selected
  if iMapPointColor = "blue" then
     lcl_selected_mpcolor_blue   = " selected=""selected"""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = ""
  elseif iMapPointColor = "orange" then
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = " selected=""selected"""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = ""
  elseif iMapPointColor = "pink" then
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = " selected=""selected"""
     lcl_selected_mpcolor_red    = ""
  elseif iMapPointColor = "red" then
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = " selected=""selected"""
  else
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = " selected=""selected"""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = ""
  end if

  response.write "  <option value=""blue"""   & lcl_selected_mpcolor_blue   & ">Blue</option>"   & vbcrlf
  response.write "  <option value=""green"""  & lcl_selected_mpcolor_green  & ">Green</option>"  & vbcrlf
  response.write "  <option value=""orange""" & lcl_selected_mpcolor_orange & ">Orange</option>" & vbcrlf
  response.write "  <option value=""pink"""   & lcl_selected_mpcolor_pink   & ">Pink</option>"   & vbcrlf
  response.write "  <option value=""red"""    & lcl_selected_mpcolor_red    & ">Red</option>"    & vbcrlf

end sub

'------------------------------------------------------------------------------
'function getMapPointTypePointColor(iDMTypeID)
function getDMTypePointColor(iDMTypeID)

  dim lcl_return

  lcl_return = "green"

  sSQL = "SELECT mappointcolor "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

  set oGetDMTypeColor = Server.CreateObject("ADODB.Recordset")
  oGetDMTypeColor.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMTypeColor.eof then
     lcl_return = oGetDMTypeColor("mappointcolor")
  end if

  oGetDMTypeColor.close
  set oGetDMTypeColor = nothing

  getDMTypePointColor = lcl_return

end function

'------------------------------------------------------------------------------
'function getMapPointTypeByFeature(p_orgid, p_featuresearch, p_feature)
function getDMTypeByFeature(p_orgid, p_featuresearch, p_feature)

  dim lcl_return, lcl_featuresearch

  lcl_return        = 0
  lcl_featuresearch = "feature_maintain"

  if p_featuresearch <> "" then
     lcl_featuresearch = p_featuresearch
  end if

  if p_feature <> "" then
     sSQL = "SELECT dm_typeid "
     sSQL = sSQL & " FROM egov_dm_types "
     sSQL = sSQL & " WHERE UPPER(" & lcl_featuresearch & ") = '" & UCASE(p_feature) & "' "
     sSQL = sSQL & " AND orgid = " & p_orgid

     set oGetDMTID = Server.CreateObject("ADODB.Recordset")
     oGetDMTID.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTID.eof then
        lcl_return = oGetDMTID("dm_typeid")
     end if

     oGetDMTID.close
     set oGetDMTID = nothing

  end if

  getDMTypeByFeature = lcl_return

end function

'------------------------------------------------------------------------------
function getFeatureFromDMType(iDMTypeID, iFeatureToRetrieve)
  dim lcl_return, sDMTypeID, sFeatureToRetrieve

  lcl_return         = "datamgr_maint"
  sDMTypeID          = 0
  sFeatureToRetrieve = "feature_maintain"

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iFeatureToRetrieve <> "" then
     sFeatureToRetrieve = iFeatureToRetrieve
     sFeatureToRetrieve = dbsafe(sFeatureToRetrieve)
  end if

  sSQL = "SELECT " & sFeatureToRetrieve & " as return_feature "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE dm_typeid = " & sDMTypeID

  set oGetFeatureFromDMType = Server.CreateObject("ADODB.Recordset")
  oGetFeatureFromDMType.Open sSQL, Application("DSN"), 3, 1

  if not oGetFeatureFromDMType.eof then
     lcl_return = oGetFeatureFromDMType("return_feature")
  end if

  oGetFeatureFromDMType.close
  set oGetFeatureFromDMType = nothing

  getFeatureFromDMType = lcl_return

end function

'------------------------------------------------------------------------------
sub getDMOwnerEditorInfo(ByVal iDMID, ByVal iUserID, ByRef lcl_ownerid, ByRef lcl_ownertype, _
                         ByRef lcl_isOwner, ByRef lcl_isApproved, ByRef lcl_isWaitingApproval)
  dim sSQL, sDMID, sUserID

  lcl_ownerid             = 0
  lcl_ownertype           = ""
  lcl_isOwner             = false
  lcl_isApproved          = false
  lcl_isWaitingApproval   = false
  lcl_approvedeniedbydate = ""

  if iDMID <> "" then
     sDMID = clng(iDMID)

     if iUserID <> "" then
        sUserID = clng(iUserID)
     else
        sUserID = 0
     end if

     sSQL = "SELECT userid, "
     sSQL = sSQL & " ownertype, "
     sSQL = sSQL & " isApproved, "
     sSQL = sSQL & " approvedeniedbydate "
     sSQL = sSQL & " FROM egov_dm_owners "
     sSQL = sSQL & " WHERE dmid = " & sDMID
     sSQL = sSQL & " AND userid = " & sUserID

     set oGetDMOwnerEditorInfo = Server.CreateObject("ADODB.Recordset")
     oGetDMOwnerEditorInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMOwnerEditorInfo.eof then
        lcl_ownerid             = oGetDMOwnerEditorInfo("userid")
        lcl_ownertype           = oGetDMOwnerEditorInfo("ownertype")
        lcl_isApproved          = oGetDMOwnerEditorInfo("isApproved")
        lcl_approvedeniedbydate = oGetDMOwnerEditorInfo("approvedeniedbydate")

        if lcl_ownertype = "OWNER" then
           lcl_isOwner = true
        end if

        if not lcl_isApproved AND lcl_approvedeniedbydate = "" then
           lcl_isWaitingApproval = true
        end if

     end if

     oGetDMOwnerEditorInfo.close
     set oGetDMOwnerEditorInfo = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub GetCityPoint(ByVal p_orgid, ByRef sLat, ByRef sLng )

    sLat = ""
    sLng = ""

   'Get the point to center the map
    sSQL = "SELECT latitude, longitude "
    sSQL = sSQL & " FROM organizations "
    sSQL = sSQL & " WHERE orgid = " & p_orgid

    set oCityPoint = Server.CreateObject("ADODB.Recordset")
    oCityPoint.Open sSQL, Application("DSN"), 3, 1

    if not oCityPoint.eof then
       sLat = oCityPoint("latitude")
       sLng = oCityPoint("longitude")
    end if

    oCityPoint.close
    set oCityPoint = nothing

end sub

'------------------------------------------------------------------------------
sub displayTemplateOptions()

  response.write "  <option value=""""></option>" & vbcrlf

  sSQL = "SELECT dm_typeid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE isTemplate = 1 "
  sSQL = sSQL & " AND isActive = 1 "

  set oGetDMTemplates = Server.CreateObject("ADODB.Recordset")
  oGetDMTemplates.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMTemplates.eof then
     do while not oGetDMTemplates.eof

        response.write "  <option value=""" & oGetDMTemplates("dm_typeid") & """>" & oGetDMTemplates("description") & "</option>" & vbcrlf

        oGetDMTemplates.movenext
     loop
  end if

  oGetDMTemplates.close
  set oGetDMTemplates = nothing

end sub

'------------------------------------------------------------------------------
sub displayLayoutOptions(iLayoutID)

  sSQL = "SELECT layoutid, "
  sSQL = sSQL & " layoutname "
  sSQL = sSQL & " FROM egov_dm_layouts "
  sSQL = sSQL & " WHERE isActive = 1 "

  set oGetDMLayoutOptions = Server.CreateObject("ADODB.Recordset")
  oGetDMLayoutOptions.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMLayoutOptions.eof then
     do while not oGetDMLayoutOptions.eof
        sLayoutID           = oGetDMLayoutOptions("layoutid")
        sLayoutName         = oGetDMLayoutOptions("layoutname")
        lcl_selected_layout = ""

        if clng(iLayoutID) = clng(sLayoutID) then
           lcl_selected_layout = " selected=""selected"""
        end if

        response.write "  <option value=""" & sLayoutid & """" & lcl_selected_layout & ">" & sLayoutName & "</option>" & vbcrlf

        oGetDMLayoutOptions.movenext
     loop
  end if

  oGetDMLayoutOptions.close
  set oGetDMLayoutOptions = nothing

end sub

'------------------------------------------------------------------------------
sub getLayoutInfo(ByVal iLayoutID, ByRef lcl_layoutname, ByRef lcl_isOriginalLayout, _
                  ByRef lcl_useLayoutSections, ByRef lcl_totalcolumns, ByRef lcl_columnwidth_left, _
                  ByRef lcl_columnwidth_middle, ByRef lcl_columnwidth_right)

  'dim lcl_layoutid, lcl_layoutname, lcl_isOriginalLayout, lcl_useLayoutSections
  'dim lcl_totalcolumns, lcl_columnwidth_left, lcl_columnwidth_middle, lcl_columnwidth_right
  dim lcl_layoutid, lcl_layoutExists

  lcl_layoutid           = 0
  lcl_layoutname         = ""
  lcl_layoutExists       = False
  lcl_isOriginalLayout   = False
  lcl_useLayoutSections  = True
  lcl_totalcolumns       = 1
  lcl_columnwidth_left   = 0
  lcl_columnwidth_middle = 0
  lcl_columnwidth_right  = 0

  if iLayoutID <> "" then
     lcl_layoutExists = checkLayoutExists(iLayoutID)
  end if

  if lcl_layoutExists then
     lcl_layoutid = iLayoutID
  else
     lcl_layoutid = getOriginalLayoutID()
  end if

 'Get Layout information
  sSQL = "SELECT layoutname, "
  sSQL = sSQL & " isOriginalLayout, "
  sSQL = sSQL & " useLayoutSections, "
  sSQL = sSQL & " totalcolumns, "
  sSQL = sSQL & " columnwidth_left, "
  sSQL = sSQL & " columnwidth_middle, "
  sSQL = sSQL & " columnwidth_right "
  sSQL = sSQL & " FROM egov_dm_layouts "
  sSQL = sSQL & " WHERE layoutid = " & lcl_layoutid

  set oGetLayoutInfo = Server.CreateObject("ADODB.Recordset")
  oGetLayoutInfo.Open sSQL, Application("DSN"), 3, 1

  if not oGetLayoutInfo.eof then
     lcl_layoutname         = oGetLayoutInfo("layoutname")
     lcl_isOriginalLayout   = oGetLayoutInfo("isOriginalLayout")
     lcl_useLayoutSections  = oGetLayoutInfo("useLayoutSections")
     lcl_totalcolumns       = oGetLayoutInfo("totalcolumns")
     lcl_columnwidth_left   = oGetLayoutInfo("columnwidth_left")
     lcl_columnwidth_middle = oGetLayoutInfo("columnwidth_middle")
     lcl_columnwidth_right  = oGetLayoutInfo("columnwidth_right")
  end if

  oGetLayoutInfo.close
  set oGetLayoutInfo = nothing

end sub

'------------------------------------------------------------------------------
'function getMPTLayoutID(iDMTypeID)
function getDMTLayoutID(iDMTypeID)

  dim lcl_return

  lcl_return = 0

  if iDMTypeID <> "" then
     sSQL = "SELECT layoutid "
     sSQL = sSQL & " FROM egov_dm_types "
     sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

     set oGetDMTLayoutID = Server.CreateObject("ADODB.Recordset")
     oGetDMTLayoutID.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTLayoutID.eof then
        lcl_return = oGetDMTLayoutID("layoutid")
     end if

     oGetDMTLayoutID.close
     set oGetDMTLayoutID = nothing

  end if

  getDMTLayoutID = lcl_return

end function

'------------------------------------------------------------------------------
function checkLayoutExists(iLayoutID)
  dim lcl_return

  lcl_return = false

  if iLayoutID <> "" then
     sSQL = "SELECT 'Y' AS lcl_exists "
     sSQL = sSQL & " FROM egov_dm_layouts "
     sSQL = sSQL & " WHERE layoutid = " & iLayoutID

     set oCheckLayoutExists = Server.CreateObject("ADODB.Recordset")
     oCheckLayoutExists.Open sSQL, Application("DSN"), 3, 1

     if not oCheckLayoutExists.eof then
        if oCheckLayoutExists("lcl_exists") = "Y" then
           lcl_return = true
        end if
     end if

     oCheckLayoutExists.close
     set oCheckLayoutExists = nothing
  end if

  checkLayoutExists = lcl_return

end function

'------------------------------------------------------------------------------
function getOriginalLayoutID()

  dim lcl_return

  lcl_return = 0

  sSQL = "SELECT layoutid "
  sSQL = sSQL & " FROM egov_dm_layouts "
  sSQL = sSQL & " WHERE isOriginalLayout = 1 "

  set oGetOriginalLayout = Server.CreateObject("ADODB.Recordset")
  oGetOriginalLayout.Open sSQL, Application("DSN"), 3, 1

  if not oGetOriginalLayout.eof then
     lcl_return = oGetOriginalLayout("layoutid")
  end if

  oGetOriginalLayout.close
  set oGetOriginalLayout = nothing

  getOriginalLayoutID = lcl_return

end function

'------------------------------------------------------------------------------
function getColumnNumber(iTotalColumns, iColumnLocation)

  dim lcl_return

  lcl_return = 1

  if iColumnLocation <> "" then
     if ucase(iColumnLocation) = "M" then
        lcl_return = 2
     elseif ucase(iColumnLocation) = "R" then
        if iTotalColumns = 3 then
           lcl_return = 3
        else
           lcl_return = 2
        end if
     end if
  end if

  getColumnNumber = lcl_return

end function

'------------------------------------------------------------------------------
function getColumnLocation(iTotalColumns, iColumnLocation)

  dim lcl_return, lcl_totalcolumns, lcl_columnlocation

  lcl_return         = ""
  lcl_totalcolumns   = 1
  lcl_columnlocation = ""

  if iTotalColumns <> "" then
     lcl_totalcolumns = iTotalColumns
  end if

  if iColumnLocation <> "" then
     lcl_columnlocation = ucase(iColumnLocation)
  end if

 'We have to handle if there is a "middle" columns, but only a 2-column layout.
 'If this is the case then move the "middle" column sections to the "right"
  if lcl_columnlocation = "M" then
     if lcl_totalcolumns = 2 then
        lcl_columnlocation = "R"
     end if
  end if

  if lcl_columnlocation = "" then
     lcl_columnlocation = "L"
  end if

  lcl_return = lcl_columnlocation

  getColumnLocation = lcl_return

end function

'------------------------------------------------------------------------------
function getAccountInfoSectionID(iDMTypeID)
  dim lcl_return

  lcl_return = 0

  if iDMTypeID <> "" then
     sSQL = "SELECT accountInfoSectionID "
     sSQL = sSQL & " FROM egov_dm_types "
     sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

     set oGetAccountInfoSID = Server.CreateObject("ADODB.Recordset")
     oGetAccountInfoSID.Open sSQL, Application("DSN"), 3, 1

     if not oGetAccountInfoSID.eof then
        lcl_return = oGetAccountInfoSID("accountInfoSectionID")
     end if

     oGetAccountInfoSID.close
     set oGetAccountInfoSID = nothing
  end if

  getAccountInfoSectionID = lcl_return

end function

'------------------------------------------------------------------------------
'sub getAddressInfo_byMapPointID(ByVal iDMID, ByRef sNumber, ByRef sPrefix, ByRef sAddress, _
sub getAddressInfo_byDMID(ByVal iDMID, ByRef sNumber, ByRef sPrefix, ByRef sAddress, _
                          ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, ByRef sZip, _
                          ByRef sValidStreet, ByRef sLatitude, ByRef sLongitude)

  'dim sDMID, sNumber, sPrefix, sAddress, sSuffix, sDirection
  'dim sCity, sState, sZip, sValidStreet, sLatitude, sLongitude

  dim sDMID

  sDMID        = 0
  sNumber      = ""
  sPrefix      = ""
  sAddress     = ""
  sSuffix      = ""
  sDirection   = ""
  sCity        = ""
  sState       = ""
  sZip         = ""
  sValidStreet = "N"
  sLatitude    = ""
  sLongitude   = ""

  if iDMID <> "" then
     sDMID = iDMID
  end if

  if sDMID > 0 then
     sSQL = "SELECT streetnumber, "
     sSQL = sSQL & " streetprefix, "
     sSQL = sSQL & " streetaddress, "
     sSQL = sSQL & " streetsuffix, "
     sSQL = sSQL & " streetdirection, "
     sSQL = sSQL & " city, "
     sSQL = sSQL & " state, "
     sSQL = sSQL & " zip, "
     sSQL = sSQL & " validstreet, "
     sSQL = sSQL & " latitude, "
     sSQL = sSQL & " longitude "
     sSQL = sSQL & " FROM egov_dm_data "
     sSQL = sSQL & " WHERE dmid = " & sDMID

     set oGetAddressInfoByMPID = Server.CreateObject("ADODB.Recordset")
     oGetAddressInfoByMPID.Open sSQL, Application("DSN"), 3, 1

     if not oGetAddressInfoByMPID.eof then
        sNumber       = oGetAddressInfoByMPID("streetnumber")
        sPrefix       = oGetAddressInfoByMPID("streetprefix")
        sAddress      = oGetAddressInfoByMPID("streetaddress")
        sSuffix       = oGetAddressInfoByMPID("streetsuffix")
        sDirection    = oGetAddressInfoByMPID("streetdirection")
        sCity         = oGetAddressInfoByMPID("city")
        sState        = oGetAddressInfoByMPID("state")
        sZip          = oGetAddressInfoByMPID("zip")
        sValidStreet  = oGetAddressInfoByMPID("validstreet")
        sLatitude     = oGetAddressInfoByMPID("latitude")
        sLongitude    = oGetAddressInfoByMPID("longitude")
     end if

     oGetAddressInfoByMPID.close
     set oGetAddressInfoByMPID = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sPrefix, ByRef sAddress, _
                    ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, ByRef sZip, _
                    ByRef sValidStreet, ByRef sLatitude, ByRef sLongitude )

 'dim sValidStreet

 sValidStreet = "N"

	sSQL = "SELECT residentstreetnumber, "
 sSQL = sSQL & " residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
 sSQL = sSQL & " streetsuffix, "
 sSQL = sSQL & " streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, "
 sSQL = sSQL & " isnull(longitude,0.00) as longitude, "
 sSQL = sSQL & " residentcity, "
 sSQL = sSQL & " residentstate, "
 sSQL = sSQL & " residentzip, "
 sSQL = sSQL & " latitude, "
 sSQL = sSQL & " longitude "
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
    sValidStreet      = "Y"
    sLatitude         = oAddress("latitude")
    sLongitude        = oAddress("longitude")
	end if

	oAddress.close
	set oAddress = nothing

end sub

'------------------------------------------------------------------------------
sub GetAddressInfoLarge( ByVal iOrgID, ByVal sStreetNumber, ByVal sStreetName, ByRef sNumber, ByRef sPrefix, _
                         ByRef sAddress, ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, _
                         ByRef sZip, ByRef sValidStreet, ByRef sLatitude, ByRef sLongitude )

 'dim sValidStreet, lcl_streetnumber, lcl_streetname

 sValidStreet     = "N"
 lcl_streetnumber = "''"
 lcl_streetname   = "''"

 if sStreetNumber <> "" then
    lcl_streetnumber = sStreetNumber
    lcl_streetnumber = dbsafe(lcl_streetnumber)
    lcl_streetnumber = ucase(lcl_streetnumber)
    lcl_streetnumber = "'" & lcl_streetnumber & "'"
 end if

 if sStreetName <> "" then
    lcl_streetname = sStreetName
    lcl_streetname = dbsafe(lcl_streetname)
    lcl_streetname = "'" & lcl_streetname & "'"
 end if 

	sSQL = "SELECT residentstreetnumber, "
 sSQL = sSQL & " residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
 sSQL = sSQL & " streetsuffix, "
 sSQL = sSQL & " streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, "
 sSQL = sSQL & " isnull(longitude,0.00) as longitude, "
 sSQL = sSQL & " residentcity, "
 sSQL = sSQL & " residentstate, "
 sSQL = sSQL & " residentzip, "
 sSQL = sSQL & " latitude, "
 sSQL = sSQL & " longitude "
 sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE orgid = " & iOrgID
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND UPPER(residentstreetnumber) = " & lcl_streetnumber
 sSQL = sSQL & " AND (residentstreetname = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = " & lcl_streetname
 sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " & lcl_streetname
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
    sValidStreet      = "Y"
    sLatitude         = oAddress("latitude")
    sLongitude        = oAddress("longitude")
	end if

	oAddress.close
	set oAddress = nothing

end sub

'------------------------------------------------------------------------------
function checkForAddressFieldInSection(iSectionID)
  dim lcl_return

  lcl_return = false

  if iSectionID <> "" then
     sSQL = "SELECT distinct 'Y' lcl_exists  "
     sSQL = sSQL & " FROM egov_dm_sections_fields "
     sSQL = sSQL & " WHERE sectionid = " & iSectionID
     sSQL = sSQL & " AND fieldtype = 'ADDRESS' "

     set oCheckForAddress = Server.CreateObject("ADODB.Recordset")
     oCheckForAddress.Open sSQL, Application("DSN"), 3, 1

     if not oCheckForAddress.eof then
        lcl_return = true
     end if

     oCheckForAddress.close
     set oCheckForAddress = nothing
  end if

  checkForAddressFieldInSection = lcl_return

end function

'------------------------------------------------------------------------------
function maintainSubCategory(iOrgID, iDMTypeID, iDMID, iUserID, iSubDelete, _
                             iSubMergeIntoCategory, iSubCategoryID, iSubCategoryName, _
                             iSubisActive, iParentCategoryID, iAssignSubCategory)

  dim lcl_return, lcl_orgid, lcl_dm_typeid, lcl_userid, lcl_current_date
  dim lcl_subcategoryid, lcl_subcategoryname, lcl_subisActive, lcl_subisApproved
  dim lcl_parentcategoryid, lcl_subdelete, lcl_submergeintocategory, lcl_assign_subcategory
  dim sCreatedByID, sCreatedByDate, sApprovedByID, sApprovedByDate

  lcl_return               = 0
  lcl_orgid                = iOrgID
  lcl_dm_typeid            = iDMTypeID
  lcl_dmid                 = iDMID
  lcl_userid               = iUserID
  lcl_current_date         = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  lcl_subcategoryid        = 0
  lcl_subcategoryname      = iSubCategoryName
  lcl_subisActive          = 0
  lcl_subisApproved        = 1
  lcl_parentcategoryid     = 0
  lcl_subdelete            = ""
  lcl_submergeintocategory = 0
  lcl_assign_subcategory   = false

  if iSubDelete <> "" then
     lcl_subdelete = iSubDelete
     lcl_subdelete = ucase(lcl_subdelete)
  end if

 'Determine if this category is being merged
  if iSubMergeIntoCategory <> "" then
     lcl_submergeintocategory = replace(iSubMergeIntoCategory,"SC","")

     if isnumeric(lcl_submergeintocategory) then
        lcl_submergeintocategory = clng(lcl_submergeintocategory)
        lcl_subdelete            = "Y"
     end if
  end if

  if iSubCategoryID <> "" then
     if isnumeric(iSubCategoryID) then
        lcl_subcategoryid = clng(iSubCategoryID)
     end if
  end if

  if iAssignSubCategory <> "" then
     lcl_assign_subcategory = iAssignSubCategory
  end if

 'Check to see if we are deleting/merging sub-categories or updating/inserting a sub-category
  if lcl_subdelete = "Y" then
     if lcl_subcategoryid > 0 then
       'Determine if the current category is to be merged into another category.
       '  First: merge the category assignments into selected category
       '  Second: delete the category
        if lcl_submergeintocategory > 0 then
           mergeCategoryAssignments lcl_orgid, lcl_dm_typeid, lcl_subcategoryid, lcl_submergeintocategory
        else
          'Delete the category assignments ONLY if we are not merging
           sSQLa = "DELETE FROM egov_dmdata_to_dmcategories "
           sSQLa = sSQLa & " WHERE categoryid = " & lcl_subcategoryid

           set oDeleteSCAssignments = Server.CreateObject("ADODB.Recordset")
           oDeleteSCAssignments.Open sSQLa, Application("DSN"), 3, 1

           set oDeleteSCAssignments = nothing

        end if

       'Delete the category
        sSQLc1 = "DELETE FROM egov_dm_categories "
        sSQLc1 = sSQLc1 & " WHERE categoryid = " & lcl_subcategoryid

        set oDeleteSubCategory = Server.CreateObject("ADODB.Recordset")
        oDeleteSubCategory.Open sSQLc1, Application("DSN"), 3, 1

        set oDeleteSubCategory = nothing

     end if
  else
     lcl_orgid     = clng(lcl_orgid)
     lcl_dm_typeid = clng(lcl_dm_typeid)
     lcl_dmid      = clng(lcl_dmid)
     lcl_userid    = clng(lcl_userid)

     if iSubisActive = "Y" then
        lcl_subisActive = 1
     end if

     if iParentCategoryID <> "" then
        lcl_parentcategoryid = clng(iParentCategoryID)
        lcl_subisApproved    = 1
        sApprovedByID        = lcl_userid
        sApprovedByDate      = lcl_current_date
     end if

     if lcl_subcategoryid > 0 then
        if iSubCategoryName <> "" then
           lcl_subcategoryname =  dbsafe(iSubCategoryName)
           lcl_subcategoryname = "'" & lcl_subcategoryname & "'"
        end if

        sSQLs = "UPDATE egov_dm_categories SET "
        sSQLs = sSQLs & "categoryname = "       & lcl_subcategoryname  & ", "
        sSQLs = sSQLs & "isActive = "           & lcl_subisActive      & ", "
        sSQLs = sSQLs & "lastmodifiedbyid = "   & lcl_userid           & ", "
        sSQLs = sSQLs & "lastmodifiedbydate = " & lcl_current_date     & ", "
        sSQLs = sSQLs & "parent_categoryid = "  & lcl_parentcategoryid
        'sSQLs = sSQLs & "isApproved = "         & lcl_isApproved      & ", "
        'sSQLs = sSQLs & "mappointcolor= "       & lcl_mappointcolor
        sSQLs = sSQLs & " WHERE categoryid = " & lcl_subcategoryid

        set oUpdateSubCategory = Server.CreateObject("ADODB.Recordset")
	       oUpdateSubCategory.Open sSQLs, Application("DSN"), 3, 1

        set oUpdateSubCategory = nothing

    '--------------------------------------------------------------------------
     else  'New Sub Category
    '--------------------------------------------------------------------------
       'Check to make sure that a duplicate categoryname doesn't exist
        lcl_subcategory_exists = checkSubCategoryExistsByCategoryName(lcl_subcategoryname, lcl_dm_typeid, lcl_parentcategoryid)

        if lcl_subcategory_exists then
           response.write "already exists"
        else
           if iSubCategoryName <> "" then
              lcl_subcategoryname =  dbsafe(iSubCategoryName)
              lcl_subcategoryname = "'" & lcl_subcategoryname & "'"
           end if

           sCreatedByID    = lcl_userid
           sCreatedByDate  = lcl_current_date

        		'Insert the new Category
	         	sSQLs = "INSERT INTO egov_dm_categories ("
           sSQLs = sSQLs & "categoryname, "
           sSQLs = sSQLs & "orgid, "
           sSQLs = sSQLs & "dm_typeid, "
           sSQLs = sSQLs & "isActive, "
           sSQLs = sSQLs & "createdbyid, "
           sSQLs = sSQLs & "createdbydate, "
           sSQLs = sSQLs & "lastmodifiedbyid, "
           sSQLs = sSQLs & "lastmodifiedbydate, "
           sSQLs = sSQLs & "parent_categoryid, "
           sSQLs = sSQLs & "isApproved, "
           sSQLs = sSQLs & "approvedeniedbyid, "
           sSQLs = sSQLs & "approvedeniedbydate "
           'sSQLs = sSQLs & "mappointcolor"
           sSQLs = sSQLs & ") VALUES ("
           sSQLs = sSQLs & lcl_subcategoryname   & ", "
           sSQLs = sSQLs & lcl_orgid             & ", "
           sSQLs = sSQLs & lcl_dm_typeid         & ", "
           sSQLs = sSQLs & lcl_subisActive       & ", "
           sSQLs = sSQLs & sCreatedByID          & ", "
           sSQLs = sSQLs & sCreatedByDate        & ", "
           sSQLs = sSQLs & "NULL,NULL"           & ", "
           sSQLs = sSQLs & lcl_parentcategoryid  & ", "
           sSQLs = sSQLs & lcl_subisApproved     & ", "
           sSQLs = sSQLs & sApprovedByID         & ", "
           sSQLs = sSQLs & sApprovedByDate
           'sSQLs = sSQLs & lcl_mappointcolor
           sSQLs = sSQLs & ")"

        		'Get the Sub-CategoryID
 	        	lcl_subcategoryid = RunIdentityInsert(sSQLs)

        end if
     end if

  end if

 'Set up the return sub_categoryid
  lcl_return = lcl_subcategoryid

  maintainSubCategory = lcl_return

end function

'------------------------------------------------------------------------------
sub maintainSubCategoryAssignments(iOrgID, iDM_TypeID, iDMID, iSubCategoryID)

  dim lcl_orgid, lcl_sub_categoryid

  lcl_orgid          = 0
  lcl_sub_categoryid = ""

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        lcl_orgid = iOrgID
     end if
  end if

  if iSubCategoryID <> "" then
     if not containsApostrophe(iSubCategoryID) then
        lcl_sub_categoryid = iSubCategoryID
     else
        sLevel = "../"  'Override of value from common.asp
       	response.redirect sLevel & "permissiondenied.asp"
     end if
  end if

  if lcl_orgid <> "" AND lcl_sub_categoryid <> "" then
     sSQL = "SELECT categoryid "
     sSQL = sSQL & " FROM egov_dm_categories "
     sSQL = sSQL & " WHERE orgid = " & lcl_orgid
     sSQL = sSQL & " AND categoryid IN (" & lcl_sub_categoryid & ") "

     set oGetSubCategories = Server.CreateObject("ADODB.Recordset")
     oGetSubCategories.Open sSQL, Application("DSN"), 3, 1

     if not oGetSubCategories.eof then
        do while not oGetSubCategories.eof

           lcl_dm_importid = ""

           addSubCategoryAssignment lcl_orgid, _
                                    iDM_TypeID, _
                                    iDMID, _
                                    oGetSubCategories("categoryid"), _
                                    lcl_dm_importid

           oGetSubCategories.movenext
        loop
     end if

     oGetSubCategories.close
     set oGetSubCategories = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub addSubCategoryAssignment(iOrgID, iDMTypeID, iDMID, iSubCategoryID, iDMImportID)

  dim sOrgID, sDMTypeID, sDMID, sSubCategoryID

  sOrgID         = 0
  sDMTypeID      = 0
  sDMID          = 0
  sSubCategoryID = ""
  sDMImportID    = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iSubCategoryID <> "" then
     sSubCategoryID = iSubCategoryID
     sSubCategoryID = dbsafe(sSubCategoryID)
     sSubCategoryID = "'" & sSubCategoryID & "'"
  end if

  if iDMImportID <> "" then
     sDMImportID = clng(iDMImportID)
  end if

  if sDMTypeID > 0 AND sDMID > 0 AND sSubCategoryID <> "" then  
     sSQL = "INSERT INTO egov_dmdata_to_dmcategories ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "dm_typeid, "
     sSQL = sSQL & "dmid, "
     sSQL = sSQL & "categoryid, "
     sSQL = sSQL & "dm_importid "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & sOrgID         & ", "
     sSQL = sSQL & sDMTypeID      & ", "
     sSQL = sSQL & sDMID          & ", "
     sSQL = sSQL & sSubCategoryID & ", "
     sSQL = sSQL & sDMImportID
     sSQL = sSQL & ") "

     set oAddSCAssignment = Server.CreateObject("ADODB.Recordset")
     oAddSCAssignment.Open sSQL, Application("DSN"), 3, 1

     set oAddSCAssignment = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub deleteSubCategoryAssignments(iDMID, iSubCategoryID)

  dim sDMID, sSubCategoryID

  sDMID          = 0
  sSubCategoryID = ""

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iSubCategoryID <> "" then
     if not containsApostrophe(iSubCategoryID) then
        sSubCategoryID = iSubCategoryID
     end if
  end if

  if sDMID > 0 then  
     sSQL = "DELETE FROM egov_dmdata_to_dmcategories "
     sSQL = sSQL & " WHERE dmid = " & sDMID

     if sSubCategoryID <> "" then
        sSQL = sSQL & " AND categoryid IN (" & sSubCategoryID & ") "
     end if

     set oRemoveSCAssignment = Server.CreateObject("ADODB.Recordset")
     oRemoveSCAssignment.Open sSQL, Application("DSN"), 3, 1

     set oRemoveSCAssignment = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub deleteDMOwners(iDMID)

  dim sDMID

  sDMID = 0

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if sDMID > 0 then
     sSQL = "DELETE FROM egov_dm_owners "
     sSQL = sSQL & " WHERE dmid = " & sDMID

     set oDeleteDMOwners = Server.CreateObject("ADODB.Recordset")
     oDeleteDMOwners.Open sSQL, Application("DSN"), 3, 1

     set oDeleteDMOwners = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub mergeCategoryAssignments(iOrgID, iDM_TypeID, iCategoryID, iMergeIntoCategoryID)

  dim sOrgID, lcl_dmtypeid, lcl_categoryid, lcl_mergeIntoCategoryID, lcl_dmid_list

  sOrgID                  = 0
  lcl_dmtypeid            = 0
  lcl_categoryid          = 0
  lcl_mergeIntoCategoryID = 0
  lcl_dmid_list           = ""

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iDM_TypeID <> "" then
     if isnumeric(iDM_TypeID) then
        lcl_dmtypeid = clng(iDM_TypeID)
     end if
  end if

  if iCategoryID <> "" then
     if isnumeric(iCategoryID) then
        lcl_categoryid = clng(iCategoryID)
     end if
  end if

  if iMergeIntoCategoryID <> "" then
     if isnumeric(iMergeIntoCategoryID) then
        lcl_mergeIntoCategoryID = clng(iMergeIntoCategoryID)
     end if
  end if

 'Combine the category ids to get a single list
  lcl_categoryid_list = ""

  if lcl_categoryid <> "" then
     if lcl_categoryid_list <> "" then
        lcl_categoryid_list = lcl_categoryid_list & "," & lcl_categoryid
     else
        lcl_categoryid_list = lcl_categoryid
     end if
  end if

  if lcl_mergeIntoCategoryID <> "" then
     if lcl_categoryid_list <> "" then
        lcl_categoryid_list = lcl_categoryid_list & "," & lcl_mergeIntoCategoryID
     else
        lcl_categoryid_list = lcl_mergeIntoCategoryID
     end if
  end if

 'BEGIN: Merge category assignments -------------------------------------------
  if lcl_categoryid_list <> "" then
     lcl_dmid_list = getDMIDCategoryAssignments(sOrgID, lcl_categoryid_list)

     if lcl_dmid_list <> "" then
       'Delete all of the current assignments
        sSQLd = "DELETE FROM egov_dmdata_to_dmcategories "
        sSQLd = sSQLd & " WHERE categoryid IN (" & lcl_categoryid_list & ") "

        set oDeleteDMIDAssignments = Server.CreateObject("ADODB.Recordset")
        oDeleteDMIDAssignments.Open sSQLd, Application("DSN"), 3, 1

        set oDeleteDMIDAssignments = nothing

       'Loop through all of the dmids and create new assignments to the merge-into-categoryid
        sSQL2 = "SELECT distinct dmid "
        sSQL2 = sSQL2 & " FROM egov_dm_data "
        sSQL2 = sSQL2 & " WHERE dmid IN (" & lcl_dmid_list & ") "

        set oGetDMIDAssignments = Server.CreateObject("ADODB.Recordset")
        oGetDMIDAssignments.Open sSQL2, Application("DSN"), 3, 1

        if not oGetDMIDAssignments.eof then
           do while not oGetDMIDAssignments.eof
              sSQLi = "INSERT INTO egov_dmdata_to_dmcategories ("
              sSQLi = sSQLi & "orgid,"
              sSQLi = sSQLi & "dm_typeid,"
              sSQLi = sSQLi & "dmid,"
              sSQLi = sSQLi & "categoryid"
              sSQLi = sSQLi & ") VALUES ("
              sSQLi = sSQLi & sOrgID                      & ", "
              sSQLi = sSQLi & lcl_dmtypeid                & ", "
              sSQLi = sSQLi & oGetDMIDAssignments("dmid") & ", "
              sSQLi = sSQLi & lcl_mergeIntoCategoryID
              sSQLi = sSQLi & ") "

              lcl_dmid_categoryid = RunIdentityInsert(sSQLi)

              oGetDMIDAssignments.movenext
           loop
        end if

        oGetDMIDAssignments.close
        set oGetDMIDAssignments = nothing
     end if
  end if
 'END: Merge category assignments ---------------------------------------------

end sub

'------------------------------------------------------------------------------
function getDMIDCategoryAssignments(iOrgID, iCategoryID)

  dim lcl_return, lcl_orgid, lcl_categoryid, lcl_dmids

  lcl_return     = ""
  lcl_orgid      = 0
  lcl_categoryid = ""
  lcl_dmids      = ""

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        lcl_orgid = clng(iOrgID)
     end if
  end if

  if iCategoryID <> "" then
     lcl_categoryid = dbsafe(iCategoryID)
  end if

 'Get all of the assignments for the categoryid
  sSQL = "SELECT distinct dmid "
  sSQL = sSQL & " FROM egov_dmdata_to_dmcategories "
  sSQL = sSQL & " WHERE orgid = " & lcl_orgid
  sSQL = sSQL & " AND categoryid IN (" & lcl_categoryid & ") "

  set oGetDMIDCategoryAssignments = Server.CreateObject("ADODB.Recordset")
  oGetDMIDCategoryAssignments.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMIDCategoryAssignments.eof then
     do while not oGetDMIDCategoryAssignments.eof
        if lcl_dmids <> "" then
           lcl_dmids = lcl_dmids & "," & oGetDMIDCategoryAssignments("dmid")
        else
           lcl_dmids = oGetDMIDCategoryAssignments("dmid")
        end if

        oGetDMIDCategoryAssignments.movenext
     loop

     lcl_return = lcl_dmids

  end if

  set oGetDMIDCategoryAssignments = nothing

  getDMIDCategoryAssignments = lcl_return

end function

'------------------------------------------------------------------------------
function checkSubCategoryExistsByCategoryName(iSubCategoryName, iDM_TypeID, iParentCategoryID)

  dim lcl_return, sSubCategoryName, sDM_TypeID, sParentCategoryID

  lcl_return        = false
  sSubCategoryName  = ""
  sDM_TypeID        = 0
  sParentCategoryID = 0

  if iSubCategoryName <> "" then
     sSubCategoryName = trim(iSubCategoryName)
     sSubCategoryName = ucase(sSubCategoryName)
     sSubCategoryName = dbsafe(sSubCategoryName)
     sSubCategoryName = "'" & sSubCategoryName & "'"
  end if

  if iDM_TypeID <> "" then
     if isnumeric(iDM_TypeID) then
        sDM_TypeID = clng(iDM_TypeID)
     end if
  end if

  if iParentCategoryID <> "" then
     sParentCategoryID = clng(iParentCategoryID)
  end if

  if sDM_TypeID > 0 AND sSubCategoryName <> "" then
     sSQL = "SELECT distinct 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_dm_categories "
     sSQL = sSQL & " WHERE dm_typeid = " & sDM_TypeID
     sSQL = sSQL & " AND UPPER(categoryname) = " & sSubCategoryName

     if sParentCategoryID > 0 then
        sSQL = sSQL & " AND parent_categoryid = " & sParentCategoryID
     end if

     set oCheckSCExistsByCategoryName = Server.CreateObject("ADODB.Recordset")
     oCheckSCExistsByCategoryName.Open sSQL, Application("DSN"), 3, 1

     if not oCheckSCExistsByCategoryName.eof then
        if oCheckSCExistsByCategoryName("lcl_exists") = "Y" then
           lcl_return = true
        end if
     end if

     oCheckSCExistsByCategoryName.close
     set oCheckSCExistsByCategoryName = nothing

  end if

  checkSubCategoryExistsByCategoryName = lcl_return

end function

'------------------------------------------------------------------------------
function checkSubCategoryAssignmentExists(iSubCategoryID, iDM_TypeID, iDMID)

  dim lcl_return, sSubCategoryID, sDM_TypeID, sDMID

  lcl_return        = false
  sSubCategoryID    = 0
  sDM_TypeID        = 0
  sDMID             = 0

  if iSubCategoryID <> "" then
     if isnumeric(iSubCategoryID) then
        sSubCategoryID = clng(iSubCategoryID)
     end if
  end if

  if iDM_TypeID <> "" then
     if isnumeric(iDM_TypeID) then
        sDM_TypeID = clng(iDM_TypeID)
     end if
  end if

  if iDMID <> "" then
     if isnumeric(iDMID) then
        sDMID = clng(iDMID)
     end if
  end if

  if sSubCategoryID > 0 AND sDM_TypeID > 0 AND sDMID > 0 then
     sSQL = "SELECT distinct 'Y' as subcategory_exists "
     sSQL = sSQL & " FROM egov_dmdata_to_dmcategories "
     sSQL = sSQL & " WHERE categoryid = " & sSubCategoryID
     sSQL = sSQL & " AND dm_typeid = "    & sDM_TypeID
     sSQL = sSQL & " AND dmid = "         & sDMID

     set oCheckSCAssignmentExists = Server.CreateObject("ADODB.Recordset")
     oCheckSCAssignmentExists.Open sSQL, Application("DSN"), 3, 1

     if not oCheckSCAssignmentExists.eof then
        if oCheckSCAssignmentExists("subcategory_exists") = "Y" then
           lcl_return = true
        end if
     end if

     oCheckSCAssignmentExists.close
     set oCheckSCAssignmentExists = nothing

  end if

  checkSubCategoryAssignmentExists = lcl_return

end function

'------------------------------------------------------------------------------
function getDMID_by_mappointid(iMapPointID)
  dim lcl_return, lcl_mappointid

  lcl_return = ""

  if iMapPointID <> "" then
     lcl_mappointid = clng(iMapPointID)

     sSQL = "SELECT dmid "
     sSQL = sSQL & " FROM egov_dm_data "
     sSQL = sSQL & " WHERE mappointid = " & lcl_mappointid

     set oGetDMIDByMPID = Server.CreateObject("ADODB.Recordset")
     oGetDMIDByMPID.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMIDByMPID.eof then
        lcl_return = oGetDMIDByMPID("dmid")
     end if

     oGetDMIDByMPID.close
     set oGetDMIDByMPID = nothing
  end if

  getDMID_by_mappointid = lcl_return

end function

'------------------------------------------------------------------------------
function getDMValueID_by_MPValueID(iMPValueID)
  dim lcl_return, lcl_mp_valueid

  lcl_return = ""

  if iMPValueID <> "" then
     lcl_mp_valueid = clng(iMPValueID)

     sSQL = "SELECT dm_valueid "
     sSQL = sSQL & " FROM egov_dm_values "
     sSQL = sSQL & " WHERE mp_valueid = " & lcl_mp_valueid

     set oGetDMValueIDByMPValueID = Server.CreateObject("ADODB.Recordset")
     oGetDMValueIDByMPValueID.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMValueIDByMPValueID.eof then
        lcl_return = oGetDMValueIDByMPValueID("dm_valueid")
     end if

     oGetDMValueIDByMPValueID.close
     set oGetDMValueIDByMPValueID = nothing

  end if

  getDMValueID_by_MPValueID = lcl_return

end function

'------------------------------------------------------------------------------
function getFieldTypeByDMFieldID(iDMTypeID, iDMFieldID)

  dim lcl_return, sDMTypeID, sDMFieldID, sSectionFieldID

  lcl_return = ""
  sDMTypeID  = 0
  sDMFieldID = 0

  if iDMTypeID <> "" then
     if isnumeric(iDMTypeID) then
        sDMTypeID = clng(iDMTypeID)
     end if
  end if

  if iDMFieldID <> "" then
     if isnumeric(iDMFieldID) then
        sDMFieldID = clng(iDMFieldID)
     end if
  end if

  if sDMTypeID > 0 AND sDMFieldID > 0 then
     sSQL = "SELECT dmsf.fieldtype "
     sSQL = sSQL & " FROM egov_dm_sections_fields dmsf "
     sSQL = sSQL & " WHERE dmsf.section_fieldid = (select dmtf.section_fieldid "
     sSQL = sSQL &                               " from egov_dm_types_fields dmtf "
     sSQL = sSQL &                               " where dmtf.dm_typeid = " & sDMTypeID
     sSQL = sSQL &                               " and dmtf.dm_fieldid = "  & sDMFieldID & ")"

     set oGetFieldType = Server.CreateObject("ADODB.Recordset")
     oGetFieldType.Open sSQL, Application("DSN"), 3, 1

     if not oGetFieldType.eof then
        lcl_return = oGetFieldType("fieldtype")
     end if

     set oGetFieldType = nothing

  end if

  getFieldTypeByDMFieldID = lcl_return


end function

'------------------------------------------------------------------------------
function getDMValueID(iOrgID, iDMTypeID, iDMID, iDMSectionID, iDMFieldID)

  dim lcl_return, sOrgID, sDMTypeID, sDMID, sDMSectionID, sDMFieldID

  lcl_return   = 0
  sOrgID       = 0
  sDMTypeID    = 0
  sDMID        = 0
  sDMSectionID = 0
  sDMFieldID   = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iDMSectionID <> "" then
     sDMSectionID = clng(iDMSectionID)
  end if

  if iDMFieldID <> "" then
     sDMFieldID = clng(iDMFieldID)
  end if

  sSQL = "SELECT dm_valueid "
  sSQL = sSQL & " FROM egov_dm_values "
  sSQL = sSQL & " WHERE orgid = "      & sOrgID
  sSQL = sSQL & " AND dm_typeid = "    & sDMTypeID
  sSQL = sSQL & " AND dmid = "         & sDMID
  sSQL = sSQL & " AND dm_sectionid = " & sDMSectionID
  sSQL = sSQL & " AND dm_fieldid = "   & iDMFieldID

	 set oGetDMValueID = Server.CreateObject("ADODB.Recordset")
 	oGetDMValueID.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMValueID.eof then
     lcl_return = oGetDMValueID("dm_valueid")
  end if

  oGetDMValueID.close
  set oGetDMValueID = nothing

  getDMValueID = lcl_return

end function

'------------------------------------------------------------------------------
function formatAdminActionsInfo(iActionName, iActionDate)

  dim lcl_return, lcl_action_name, lcl_action_date

  lcl_return = ""
  lcl_action_name = trim(iActionName)
  lcl_action_date = trim(iActionDate)

  if lcl_action_name <> "" OR lcl_action_date <> "" then
     if lcl_action_name <> "" then
        lcl_return = lcl_action_name
     end if

     if lcl_action_date <> "" then
        if lcl_return <> "" then
           lcl_return = lcl_return & "<br />" & vbcrlf
        end if

        lcl_return = lcl_return & "<span style=""color:#800000;"">[" & lcl_action_date & "]</span>" & vbcrlf
     end if
  end if

  formatAdminActionsInfo = lcl_return

end function

'------------------------------------------------------------------------------
function buildURLDisplayValue(iFieldType, iFieldValue)
  dim lcl_return, sFieldType, sFieldValue, lcl_url_value, lcl_current_value, lcl_total_urls
  dim lcl_display_value, lcl_display_url, lcl_display_text
  dim lcl_comma_position, lcl_url_start, lcl_url_end, lcl_url_length
  dim lcl_text_start, lcl_text_end, lcl_text_length, lcl_website_url, lcl_website_text

  lcl_return  = ""
  sFieldType  = ""
  sFieldValue = ""

  if iFieldType <> "" then
     if not containsApostrophe(iFieldType) then
        sFieldType = ucase(iFieldType)
     end if
  end if

  if iFieldValue <> "" then
     sFieldValue = iFieldValue
  end if

  'Break out the values.  Websites and Emails are stored in the following format:
  '1. URL/Email address
  '2. Display Text (clickable link - value CAN be NULL)
  '3. URLs will be surrounded by [].
  '4. Display Text will be surrounded by <>.
  '5. Multiple websites and emails will be seperated by a comma.
  '6. If Display Text is NULL then the URL/Email will be used as the "clickable link".
  '   i.e. [www.mywebsite.com]<My Website>,[www.anotherwebsite.com]<>
   lcl_url_value     = sFieldValue
   lcl_current_value = ""
   lcl_display_value = ""
   lcl_display_url   = ""
   lcl_display_text  = ""
   lcl_total_urls    = 0

   if lcl_url_value <> "" then
      do until lcl_url_value = ""
         lcl_comma_position = 0
         lcl_url_start      = 0
         lcl_url_end        = 0
         lcl_url_length     = 0
         lcl_text_start     = 0
         lcl_text_end       = 0
         lcl_text_length    = 0
         lcl_website_url    = ""
         lcl_website_text   = ""

         if lcl_url_value <> "" then
            lcl_total_urls     = lcl_total_urls + 1
            'lcl_comma_position = instr(lcl_url_value,",")
            lcl_comma_position = instr(lcl_url_value,">,[")

            if lcl_comma_position > 0 then
               lcl_current_value = mid(lcl_url_value,1,lcl_comma_position+1)
            else
               lcl_current_value = lcl_url_value
            end if

            lcl_url_start    = instr(lcl_current_value,"[")
            lcl_url_end      = instr(lcl_current_value,"]")
            lcl_url_length   = lcl_url_end - lcl_url_start

            lcl_text_start   = instr(lcl_current_value,"<")
            lcl_text_end     = instr(lcl_current_value,">")
            lcl_text_length  = lcl_text_end - lcl_text_start

            if lcl_url_start > -1 AND lcl_url_length > 0 then
               lcl_website_url  = mid(lcl_current_value,lcl_url_start,lcl_url_length)
               lcl_website_url  = replace(lcl_website_url,"[","")
               lcl_website_url  = replace(lcl_website_url,"]","")
            end if

            if lcl_text_start > -1 AND lcl_text_length > 0 then
               lcl_website_text = mid(lcl_current_value,lcl_text_start,lcl_text_length)
               lcl_website_text = replace(lcl_website_text,"<","")
               lcl_website_text = replace(lcl_website_text,">","")
            end if

            if lcl_website_text <> "" then
               lcl_display_text = lcl_website_text
            else
               lcl_display_text = lcl_website_url
            end if

            lcl_url_value   = replace(lcl_url_value,lcl_current_value,"")

           'Build the "display_url"
            lcl_display_url = "<a href="""

            if instr(sFieldType,"EMAIL") > 0 then
               lcl_display_url = lcl_display_url & "mailto:"
            else
               if instr(lcl_website_url,"http://") = 0 AND instr(lcl_website_url,"https://") = 0 then
                  lcl_display_url = lcl_display_url & "http://"
               end if
            end if

            lcl_display_url = lcl_display_url & lcl_website_url & """ target=""_blank"">" & lcl_display_text & "</a>"

            if lcl_display_value <> "" then
               lcl_display_value = lcl_display_value & "<br />" & lcl_display_url
            else
               lcl_display_value = lcl_display_url
            end if
         end if
      loop

      if lcl_display_value <> "" then
         lcl_return = lcl_display_value
      end if
   end if

   buildURLDisplayValue = lcl_return

end function

'------------------------------------------------------------------------------
function isCheckboxChecked(iValue)

  dim lcl_return

  lcl_return = ""

  if iValue then
     lcl_return = " checked=""checked"""
  end if

  isCheckboxChecked = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(iValue)

  dim lcl_return

  lcl_return = ""

  if iValue <> "" then
     lcl_return = replace(iValue,"'","''")
  end if

  dbsafe = lcl_return


end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub
%>
