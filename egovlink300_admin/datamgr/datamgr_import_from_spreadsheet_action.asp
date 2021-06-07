<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  sLevel = "../"  'Override of value from common.asp

 'Determine if the parent feature is "offline"
  if isFeatureOffline("datamgr") = "Y" then
     response.redirect sLevel & "permissiondenied.asp"
  end if

 'Determine if the user is a "root admin"
  lcl_isRootAdmin = False

  if UserIsRootAdmin(session("userid")) then
     lcl_isRootAdmin = True
  end if

  if not lcl_isRootAdmin then
    	response.redirect sLevel & "permissiondenied.asp"
  end if

 'Retreive the values
  'dim lcl_dm_sectionid_current, lcl_dm_sectionid_new
  'dim lcl_dm_fieldid_current, lcl_dm_fieldid_new
  'dim lcl_mp_fieldid, lcl_transfer_field_data, lcl_dm_sectionid, lcl_dm_fieldid
  dim lcl_feature, lcl_userid, lcl_orgid, lcl_dm_typeid, lcl_mp_valueid, lcl_dm_importid
  dim lcl_sc_dm_importid, lcl_includeOnlyImportedNVA
  dim i, lcl_error_msg, lcl_return_msg, lcl_totalfields, lcl_overrideValues
  dim lcl_sub_categorytype, lcl_dbcolumn_name, lcl_validate_addresses
  dim lcl_linecount, lcl_address_row, lcl_sc_latitude, lcl_sc_longitude
  dim lcl_latitude_dm_valueid, lcl_longitude_dm_valueid, lcl_dm_sectionid, lcl_dm_fieldid
  dim lcl_fieldvalue, lcl_latitude_fieldvalue, lcl_longitude_fieldvalue

  lcl_feature                  = ""
  lcl_userid                   = session("userid")
  lcl_orgid                    = ""
  lcl_dmid                     = ""
  lcl_dm_typeid                = ""
  lcl_dm_valueid               = 0
  lcl_dm_sectionid             = ""
  lcl_dm_fieldid               = ""
  lcl_mp_valueid               = ""
  lcl_dm_importid              = ""
  lcl_sc_dm_importid           = ""
  lcl_includeOnlyImportedNVA   = ""
  lcl_sub_categorytype         = ""
  lcl_dbcolumn_name            = ""
  lcl_validate_addresses       = "N"
  lcl_linecount                = 0
  lcl_address_row              = ""
  lcl_nvaddresses_address      = ""
  lcl_nvaddresses_city         = ""
  lcl_nvaddresses_state        = ""
  lcl_nvaddresses_latitude     = ""
  lcl_nvaddresses_longitude    = ""
  lcl_sc_latitude              = ""
  lcl_sc_longitude             = ""
  lcl_latitude_dm_typeid       = ""
  lcl_latitude_dm_sectionid    = ""
  lcl_latitude_dm_fieldid      = ""
  lcl_longitude_dm_typeid      = ""
  lcl_longitude_dm_sectionid   = ""
  lcl_longitude_dm_fieldid     = ""
  lcl_fieldvalue               = ""
  lcl_latitude_fieldvalue      = ""
  lcl_longitude_fieldvalue     = ""
  lcl_action                   = ""
  lcl_overrideValues           = "N"
  lcl_error_msg                = ""
  lcl_return_msg               = ""

  if request("f") <> "" then
     if not containsApostrophe(request("f")) then
        lcl_feature = ucase(request("f"))
     end if
  end if

  'if request("userid") <> "" then
  '   lcl_userid = request("userid")

  '   if not isnumeric(lcl_userid) then
  '      lcl_error_msg = "INVALID VALUE: Non-numeric value in 'userid'"
  '   else
  '      lcl_userid = clng(lcl_userid)
  '   end if
  'end if

  if request("orgid") <> "" then
     lcl_orgid = request("orgid")

     if not isnumeric(lcl_orgid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'orgid'"
     else
        lcl_orgid = clng(lcl_orgid)
     end if
  end if

  if request("dm_typeid") <> "" then
     lcl_dm_typeid = request("dm_typeid")

     if not isnumeric(lcl_dm_typeid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'dm_typeid'"
     else
        lcl_dm_typeid = clng(lcl_dm_typeid)
     end if
  end if

  if request("dm_importid") <> "" then
     lcl_dm_importid = request("dm_importid")

     if not isnumeric(lcl_dm_importid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'dm_importid'"
     else
        lcl_dm_importid = clng(lcl_dm_importid)
     end if
  end if

  if request("sc_dm_importid") <> "" then
     lcl_sc_dm_importid = request("sc_dm_importid")

     if not isnumeric(lcl_sc_dm_importid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'DM ImportID'"
     else
        lcl_sc_dm_importid = clng(lcl_sc_dm_importid)
     end if
  end if

  if request("includeOnlyImportedNVA") <> "" then
     if not containsApostrophe(request("includeOnlyImportedNVA")) then
        lcl_includeOnlyImportedNVA = ucase(request("includeOnlyImportedNVA"))
     end if
  end if

'  if request("totalfields") <> "" then
'     lcl_totalfields = request("totalfields")

'     if not isnumeric(lcl_totalfields) then
'        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'totalfields'"
'     else
'        lcl_totalfields = clng(lcl_totalfields)
'     end if
'  end if

'  if request("mp_fieldid") <> "" then
'     lcl_mp_fieldid = request("mp_fieldid")

'     if not isnumeric(lcl_mp_fieldid) then
'        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'mp_fieldid'"
'     else
'        lcl_mp_fieldid = clng(lcl_mp_fieldid)
'     end if
'  end if

  if request("overrideValues") <> "" then
     if not containsApostrophe(request("overrideValues")) then
        lcl_overrideValues = ucase(request("overrideValues"))
     end if
  end if

  if request("action") <> "" then
     if not containsApostrophe(request("action")) then
        lcl_action = ucase(request("action"))
     end if
  end if

  if request("subcategorytype") <> "" then
     if not containsApostrophe(request("subcategorytype")) then
        lcl_sub_categorytype = request("subcategorytype")
        lcl_sub_categorytype = ucase(lcl_sub_categorytype)
     end if
  end if

 'Check for an error message.
  if lcl_error_msg <> "" then
     response.write lcl_error_msg
  else
     if lcl_action <> "" then

       'BEGIN: Determine which action to take in the import process -----------
       '-----------------------------------------------------------------------
        if lcl_action = "START_IMPORT" then
           lcl_current_date = ConvertDateTimetoTimeZone()

          'Create the import record
           sSQLi = "INSERT INTO egov_dm_import ("
           sSQLi = sSQLi & "orgid, "
           sSQLi = sSQLi & "status, "
           sSQLi = sSQLi & "importedbyid, "
           sSQLi = sSQLi & "importstartdate "
           sSQLi = sSQLi & ") VALUES ( "
           sSQLi = sSQLi & lcl_orgid              & ", "
           sSQLi = sSQLi & "'IMPORT STARTED', "
           sSQLi = sSQLi & lcl_userid             & ", "
           sSQLi = sSQLi & "'" & lcl_current_date & "'"
           sSQLi = sSQLi & ") "

           lcl_dm_importid = RunInsertStatement(sSQLi)

           response.write lcl_dm_importid

       '-----------------------------------------------------------------------
        elseif lcl_action = "ASSIGN_ORG_DMTYPE" then

          'Create a DMID for each import data record
           sSQLid = "SELECT dm_importdata_id "
           sSQLid = sSQLid & " FROM egov_dm_import_data "
           sSQLid = sSQLid & " ORDER BY dm_importdata_id "

          	set oGetImportDataIDs = Server.CreateObject("ADODB.Recordset")
         	 oGetImportDataIDs.Open sSQLid, Application("DSN"), 3, 1

           if not oGetImportDataIDs.eof then
              do while not oGetImportDataIDs.eof

                'Create a DMID for each import record
                 sDMID        = ""
                 sDMSectionID = ""
                 sDMFieldID   = ""
                 sIsActive    = 1
                 sCategoryID  = ""

                 maintainDMData lcl_userid, _
                                lcl_orgid, _
                                sDMID, _
                                lcl_dm_typeid, _
                                sDMSectionID, _
                                sDMFieldID, _
                                sIsActive, _
                                sCategoryID, _
                                lcl_dm_importid, _
                                lcl_dmid

                'Update the dm_importid, orgid, dm_typeid, and dmid columns on egov_dm_import.
                 sSQLu = "UPDATE egov_dm_import_data SET "
                 sSQLu = sSQLu & " dm_importid = " & lcl_dm_importid & ", "
                 sSQLu = sSQLu & " orgid = "       & lcl_orgid       & ", "
                 sSQLu = sSQLu & " dm_typeid = "   & lcl_dm_typeid   & ", "
                 sSQLu = sSQLu & " dmid = "        & lcl_dmid
                 sSQLu = sSQLu & " WHERE dm_importdata_id = " & oGetImportDataIDs("dm_importdata_id")

                	set oUpdateOrgDMType = Server.CreateObject("ADODB.Recordset")
               	 oUpdateOrgDMType.Open sSQLu, Application("DSN"), 3, 1

                 set oUpdateOrgDMType = nothing

                 oGetImportDataIDs.movenext
              loop
           end if

           set oGetImportDataIDs = nothing

           response.write "Organization and DM Type Assigned"

           response.write lcl_return_msg

       '-----------------------------------------------------------------------
        elseif lcl_action = "SETUP_CATEGORIES" then
           lcl_total_categories_created = 0
           lcl_categoryid               = 0
           lcl_category_created         = 0

          '1. Pull a distinct list of all categories from import list
          '2. Cycle through list and determine if any of the categories already exist for the orgid and dm_typeid
          '3. If "no" then create the new sub-category and assign the categoryid to the row(s) in the import
          '4. If "yes" then retrieve the categoryid and assign it to the row(s) in the import
           sSQLc = "SELECT distinct category, "
           sSQLc = sSQLc & " dmid "
           sSQLc = sSQLc & " FROM egov_dm_import_data "
           sSQLc = sSQLc & " WHERE dm_importid = " & lcl_dm_importid
           sSQLc = sSQLc & " AND category <> '' "
           sSQLc = sSQLc & " AND category IS NOT NULL "
           sSQLc = sSQLc & " ORDER BY category "

          	set oGetImportCategoriesList = Server.CreateObject("ADODB.Recordset")
         	 oGetImportCategoriesList.Open sSQLc, Application("DSN"), 3, 1

           if not oGetImportCategoriesList.eof then
              do while not oGetImportCategoriesList.eof
                 getCategoryID lcl_userid, _
                               lcl_orgid, _
                               lcl_dm_typeid, _
                               lcl_dm_importid, _
                               oGetImportCategoriesList("category"), _
                               lcl_categoryid, _
                               lcl_category_created

                 lcl_update_column = "categoryid"
                 lcl_update_value  = lcl_categoryid
                 lcl_search_column = "category"
                 lcl_search_value  = oGetImportCategoriesList("category")

                 updateDMImport lcl_update_column, lcl_update_value, lcl_search_column, lcl_search_value

                 lcl_total_categories_created = lcl_total_categories_created + lcl_category_created

                'Now update egov_dm_data (dmid) and assign the category to the DMID
                 sSQLd = "UPDATE egov_dm_data SET "
                 sSQLd = sSQLd & " categoryid = " & lcl_categoryid
                 sSQLd = sSQLd & " WHERE dmid = " & oGetImportCategoriesList("dmid")

                	set oUpdateDMIDCategoryID = Server.CreateObject("ADODB.Recordset")
               	 oUpdateDMIDCategoryID.Open sSQLd, Application("DSN"), 3, 1

                 set oUpdateDMIDCategoryID = nothing
 
                 oGetImportCategoriesList.movenext
              loop
           end if

           oGetImportCategoriesList.close
           set oGetImportCategoriesList = nothing

          'Build return message
           if lcl_total_categories_created > 0 then
              lcl_return_msg = "Categories Created <span class=""redText"">[" & lcl_total_categories_created & "]</span>"
           else
              lcl_return_msg = "COMPLETED"
           end if

           response.write lcl_return_msg

       '-----------------------------------------------------------------------
        elseif lcl_action = "SETUP_SUBCATEGORIES" then
           lcl_total_subcategories_created = 0
           lcl_subcategoryid               = 0
           lcl_subcategory_created         = 0

          '1. Pull a distinct list of all sub-categories from import list
          '2. If this is a "1to1" sub-category type, then cycle through list and determine if any of the 
          '   sub-categories already exist for the orgid and dm_typeid and the category for the row(s)
          '3. If "no" then create the new sub-category and assign the subcategoryid to the row(s) in the import
          '4. If "yes" then retrieve the subcategoryid and assign it to the row(s) in the import
           sSQLsc = "SELECT distinct subcategory, "
           sSQLsc = sSQLsc & " category, "
           sSQLsc = sSQLsc & " categoryid, "
           'sSQLsc = sSQLsc & " dm_importdata_id, "
           sSQLsc = sSQLsc & " dmid "
           sSQLsc = sSQLsc & " FROM egov_dm_import_data "
           sSQLsc = sSQLsc & " WHERE dm_importid = " & lcl_dm_importid
           sSQLsc = sSQLsc & " AND subcategory <> '' "
           sSQLsc = sSQLsc & " AND subcategory IS NOT NULL "
           sSQLsc = sSQLsc & " AND category <> '' "
           sSQLsc = sSQLsc & " AND category IS NOT NULL "
           sSQLsc = sSQLsc & " ORDER BY category, subcategory "

          	set oGetImportSubCategoriesList = Server.CreateObject("ADODB.Recordset")
         	 oGetImportSubCategoriesList.Open sSQLsc, Application("DSN"), 3, 1

           if not oGetImportSubCategoriesList.eof then
              do while not oGetImportSubCategoriesList.eof

                'Insert/Update the Sub-Category
                 getSubCategoryID lcl_userid, _
                                  lcl_orgid, _
                                  lcl_dm_typeid, _
                                  lcl_dm_importid, _
                                  oGetImportSubCategoriesList("subcategory"), _
                                  oGetImportSubCategoriesList("categoryid"), _
                                  lcl_subcategoryid, _
                                  lcl_subcategory_created

                'Assign the Sub-Category to the DM Type ONLY if an assignment does NOT exist
                 lcl_subcategory_assignment_exists = false
                 lcl_subcategory_assignment_exists = checkSubCategoryAssignmentExists(lcl_subcategoryid, _
                                                                                      lcl_dm_typeid, _
                                                                                      oGetImportSubCategoriesList("dmid"))

                 if not lcl_subcategory_assignment_exists then
                    addSubCategoryAssignment lcl_orgid, _
                                             lcl_dm_typeid, _
                                             oGetImportSubCategoriesList("dmid"), _
                                             lcl_subcategoryid, _
                                             lcl_dm_importid
                 end if

                'Update the import with the sub-category id
                 lcl_update_column = "subcategoryid"
                 lcl_update_value  = lcl_subcategoryid
                 lcl_search_column = "subcategory"
                 lcl_search_value  = oGetImportSubCategoriesList("subcategory")

                 updateDMImport lcl_update_column, _
                                lcl_update_value, _
                                lcl_search_column, _
                                lcl_search_value

                 lcl_total_subcategories_created = lcl_total_subcategories_created + lcl_subcategory_created

                 oGetImportSubCategoriesList.movenext
              loop
           end if

           oGetImportSubCategoriesList.close
           set oGetImportSubCategoriesList = nothing

          'Build return message
           if lcl_total_subcategories_created > 0 then
              lcl_return_msg = "Sub-Categories Created <span class=""redText"">[" & lcl_total_subcategories_created & "]</span>"
           else
              lcl_return_msg = "COMPLETED"
           end if

           response.write lcl_return_msg

       '-----------------------------------------------------------------------
        elseif lcl_action = "BUILD_DM_TRANSFERFIELD_OPTIONS" then
           lcl_transferFieldsOptions = ""

           sSQL = sSQL & "SELECT DISTINCT "
           sSQL = sSQL & " dmtf.dm_sectionid, "
           sSQl = sSQL & " dms.sectionname, "
           sSQL = sSQL & " dmtf.dm_fieldid, "
           sSQL = sSQL & " dmtf.section_fieldid, "
           sSQL = sSQL & " dmsf.fieldname "
           sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
           sSQL = sSQL &      " INNER JOIN egov_dm_types dmt "
           sSQL = sSQL &            " ON dmt.dm_typeid = dmtf.dm_typeid "
           sSQL = sSQL &            " AND dmt.isActive = 1 "
           sSQL = sSQL &            " AND dmt.isTemplate = 0 "
           sSQL = sSQL &            " AND dmt.orgid = " & lcl_orgid
           sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
           sSQL = sSQL &            " ON dmts.dm_sectionid = dmtf.dm_sectionid "
           sSQL = sSQL &            " AND dmts.isActive = 1 "
           sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
           sSQL = sSQL &            " ON dms.sectionid = dmts.sectionid "
           sSQL = sSQL &            " AND dms.isActive = 1 "
           sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
           sSQL = sSQL &            " ON dmsf.section_fieldid = dmtf.section_fieldid "
           sSQL = sSQL &            " AND dmsf.isActive = 1 "
           sSQL = sSQL & " WHERE dmtf.orgid = " & lcl_orgid
           sSQL = sSQL & " AND dmtf.dm_typeid = " & lcl_dm_typeid
           sSQL = sSQL & " ORDER BY dms.sectionname, dmsf.fieldname "

          	set oDMTransferFieldsOptions = Server.CreateObject("ADODB.Recordset")
         	 oDMTransferFieldsOptions.Open sSQL, Application("DSN"), 3, 1
	
          	if not oDMTransferFieldsOptions.eof then
              do while not oDMTransferFieldsOptions.eof

                 lcl_transferFieldsOptions = lcl_transferFieldsOptions & "  <option value=""dmsectionid" & oDMTransferFieldsOptions("dm_sectionid") & "_dmfieldid" & oDMTransferFieldsOptions("dm_fieldid") & """>" & oDMTransferFieldsOptions("sectionname") & ": " & oDMTransferFieldsOptions("fieldname") & "</option>" & vbcrlf

                 oDMTransferFieldsOptions.movenext
              loop
           end if

           oDMTransferFieldsOptions.close
           set oDMTransferFieldsOptions = nothing

           response.write lcl_transferFieldsOptions

       '-----------------------------------------------------------------------
        elseif lcl_action = "IMPORT_SPREADSHEET_VALUES" then
           lcl_orghasfeature_issue_location     = orghasfeature("issue location")
           lcl_orghasfeature_large_address_list = orghasfeature("large address list")

           lcl_dmid         = 0
           lcl_sectionid    = 0
           lcl_dm_sectionid = 0
           lcl_dm_fieldid   = 0
           lcl_totalfields  = 0
           'lcl_feature     = "datamgr_maint"
           lcl_featurename  = getFeatureName(lcl_feature)

           if request("dmid") <> "" then
              lcl_dmid = request("dmid")
           end if

           if request("sectionid") <> "" then
              lcl_sectionid = request("sectionid")

              if not isnumeric(lcl_sectionid) then
                 lcl_error_msg = "INVALID VALUE: Non-numeric value in 'sectionid'"
              else
                 lcl_sectionid = clng(lcl_sectionid)
              end if
           end if

           if request("dbcolumn_name") <> "" then
              if not containsApostrophe(request("dbcolumn_name")) then
                 lcl_dbcolumn_name = ucase(request("dbcolumn_name"))
              end if
           end if

           if request("validateAddresses") <> "" then
              if not containsApostrophe(request("validateAddresses")) then
                 lcl_validate_addresses = ucase(request("validateAddresses"))
              end if
           end if

          'BEGIN: Cycle through all of the records in the import table. -------
          'We need to perform an initial check to see if a record exists for each field on the 
          '   egov_datamgr_values table.  This table is where we store the value for each field.
          'If there is NOT a "dm_valueid" then we need to INSERT a record.
          'If there IS a "dm_valueid" then we need to update the record.
           lcl_fieldtype = ""

           if request("transfer_field_data") <> "" then
              if not containsApostrophe(request("transfer_field_data")) then
                 lcl_transfer_field_data = split(request("transfer_field_data"), "_")
                 lcl_dm_sectionid        = replace(lcl_transfer_field_data(0),"dmsectionid","")
                 lcl_dm_fieldid          = replace(lcl_transfer_field_data(1),"dmfieldid","")
              end if
           end if

           lcl_fieldtype = getFieldTypeByDMFieldID(lcl_dm_typeid, lcl_dm_fieldid)

           if lcl_dbcolumn_name <> "" then
              sSQLsi = "SELECT " & lcl_dbcolumn_name & " as dbcolumn_name, "
              sSQLsi = sSQLsi & " dm_importdata_id, "
              sSQLsi = sSQLsi & " dm_importid, "
              sSQLsi = sSQLsi & " orgid, "
              sSQLsi = sSQLsi & " dm_typeid, "
              sSQLsi = sSQLsi & " dmid "
              'sSQLsi = sSQLsi & " categoryid, "
              'sSQLsi = sSQLsi & " category, "
              'sSQLsi = sSQLsi & " subcategoryid, "
              'sSQLsi = sSQLsi & " subcategory "
              sSQLsi = sSQLsi & " FROM egov_dm_import_data "
              sSQLsi = sSQLsi & " WHERE orgid = " & lcl_orgid
              sSQLsi = sSQLsi & " AND dm_importid = " & lcl_dm_importid
              sSQLsi = sSQLsi & " ORDER BY dm_importdata_id "
'dtb_debug(sSQLsi)
              set oGetImportRecords = Server.CreateObject("ADODB.Recordset")
              oGetImportRecords.Open sSQLsi, Application("DSN"), 3, 1

              if not oGetImportRecords.eof then
                 do while not oGetImportRecords.eof

                    lcl_dm_valueid    = 0
                    lcl_fieldvalue    = ""
                    lcl_dmid          = oGetImportRecords("dmid")
                    lcl_dmv_address   = ""
                    lcl_address       = ""
                    lcl_streetnumber  = ""
                    lcl_streetaddress = ""
                    lcl_latitude      = ""
                    lcl_longitude     = ""
                    sNumber           = ""
                    sPrefix           = ""
                    sAddress          = ""
                    sSuffix           = ""
                    sDirection        = ""
                    sLatitude         = ""
                    sLongitude        = ""
                    sCity             = ""
                    sState            = ""
                    sZip              = ""
                    sCounty           = ""
                    sParcelID         = ""
                    sListedOwner      = ""
                    sLegalDescription = ""
                    sResidentType     = ""
                    sRegisteredUserID = ""
                    sValidStreet      = "N"

                   'BEGIN: Check to see if the address is valid or custom --------------------
                    if instr(lcl_fieldtype,"ADDRESS") > 0 OR instr(lcl_fieldtype,"LATITUDE") > 0 OR instr(lcl_fieldtype,"LONGITUDE") > 0 then
                       if lcl_orghasfeature_issue_location then

                          if oGetImportRecords("dbcolumn_name") <> "" then
                             if lcl_address <> "" then
                                lcl_address = lcl_address & " " & oGetImportRecords("dbcolumn_name")
                             else
                                lcl_address = oGetImportRecords("dbcolumn_name")
                             end if
                          end if
'dtb_debug("lcl_address BEFORE BREAKOUT: [" & lcl_address & "] - sStreetNumber: [" & sStreetNumber& "] - sStreetName: [" & sStreetName & "]")
                          breakOutAddress lcl_address, sStreetNumber, sStreetName
'dtb_debug("lcl_address AFTER BREAKOUT: [" & lcl_address & "] - sStreetNumber: [" & sStreetNumber& "] - sStreetName: [" & sStreetName & "]")

'dtb_debug("1. [" & lcl_address & "] - [" & sStreetNumber & "] - [" & sStreetName & "] - lcl_validate_addresses: [" & lcl_validate_addresses & "]")
                         'Are we to validate the addresses: yes/no?
                         'If "yes" then:
                         '   Check to see if the address exists on the "valid address" table.
                         '   If it doesn't exist then mark the address as "non-valid"
                         'If "no" then bypass table "check"
                          if lcl_validate_addresses = "Y" then
                             getAddressInfoNew lcl_orghasfeature_large_address_list, _
                                               lcl_orgid, _
                                               sStreetNumber, _
                                               sStreetName, _
                                               sNumber, _
                                               sPrefix, _
                                               sAddress, _
                                               sSuffix, _
                                               sDirection, _
                                               sLatitude, _
                                               sLongitude, _
                                               sCity, _
                                               sState, _
                                               sZip, _
                                               sCounty, _
                                               sParcelID, _
                                               sListedOwner, _
                                               sLegalDescription, _
                                               sResidentType, _
                                               sRegisteredUserID, _
                                               sValidStreet
'response.write "2. [" & lcl_address & "] - [" & sNumber & "] - [" & sAddress & "]"
                             lcl_dmv_address = buildStreetAddress(sNumber, sPrefix, sAddress, sSuffix, sDirection)
                          end if
'dtb_debug("2. sAddress: [" & sAddress & "] - lcl_address: [" & lcl_address & "] - lcl_dmv_address: [" & lcl_dmv_address & "]")
                         'If the address is NOT a valid address then we need to populate the proper variables
                         'so that the record is built properly
                         '  1. lcl_dmv_address = Address value to be used in the fieldvalue for egov_dm_values
                         '  2. sAddress        = Address value to be used in the "streetaddress" column on egov_dm_data
                         '  3. sSortStreetName = Address value to be used to sort the address values on egov_dm_data
                          if trim(lcl_dmv_address) = "" then
                             lcl_dmv_address = lcl_address
                             sAddress        = lcl_address
                             sSortStreetName = sStreetName
                             sValidStreet    = "N"
                          else
                            'BEGIN: Re-build the SortStreetName ------------------
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
                            'END: Re-build the SortStreetName --------------------
                          end if

'dtb_debug("3. sAddress: [" & sAddress & "] - lcl_address: [" & lcl_address & "] - lcl_dmv_address: [" & lcl_dmv_address & "]")

                       else
                          if instr(lcl_fieldtype,"ADDRESS") > 0 then
                             sNumber    = ""
                             sAddress   = oGetImportRecords("dbcolumn_name")
                             sPrefix    = ""
                             sSuffix    = ""
                             sDirection = ""

                             lcl_dmv_address = sAddress
                          elseif instr(lcl_fieldtype,"LATITUDE") > 0 then
                             sLatitude = oGetImportRecords("dbcolumn_name")
'dtb_debug(lcl_fieldtype & " [" & sLatitude & "] - [" & isnull(oGetImportRecords("dbcolumn_name")) & "]")
                          elseif instr(lcl_fieldtype,"LONGITUDE") > 0 then
                             sLongitude = oGetImportRecords("dbcolumn_name")
'dtb_debug(lcl_fieldtype & " [" & sLatitude & "] - [" & isnull(oGetImportRecords("dbcolumn_name")) & "]")
                          end if
                       end if

                      'Determine if we pull the value from the screen or if we generate it.
                       if instr(lcl_fieldtype,"ADDRESS") > 0 then
                          lcl_fieldvalue = lcl_dmv_address
                       elseif instr(lcl_fieldtype,"LATITUDE") > 0 then
                          lcl_fieldvalue = sLatitude
                       elseif instr(lcl_fieldtype,"LONGITUDE") > 0 then
                          lcl_fieldvalue = sLongitude
                       'else
                       '   lcl_fieldvalue = request("dm_fieldvalue" & v)
                       end if

                      'Make one last check for the Latitude/Longitude fields.
                      'If they are NULL then default the value to 0.00
                       if instr(lcl_fieldtype,"LATITUDE") > 0 OR instr(lcl_fieldtype,"LONGITUDE") > 0 then
                          if lcl_fieldvalue = "" OR isnull(lcl_fieldvalue) then
                             lcl_fieldvalue = "0.00"
                          end if
                       end if

                       'lcl_fieldvalue = lcl_address
                       if instr(lcl_fieldtype,"ADDRESS") > 0 then
'dtb_debug("maintainDMSection_address - sAddress: [" & sAddress & "]")
                          maintainDMSection_address lcl_userid, _
                                                    lcl_orgid, _
                                                    lcl_dmid, _
                                                    lcl_dm_typeid, _
                                                    lcl_dm_sectionid, _
                                                    lcl_dm_fieldid, _
                                                    sNumber, _
                                                    sPrefix, _
                                                    sAddress, _
                                                    sSuffix, _
                                                    sDirection, _
                                                    sSortStreetName, _
                                                    sCity, _
                                                    sStatus, _
                                                    sZip, _
                                                    sValidStreet, _
                                                    sLatitude, _
                                                    sLongitude, _
                                                    lcl_dmid

                       elseif instr(lcl_fieldtype,"LATITUDE") > 0 then
                        		sSQLlat = "UPDATE egov_dm_data SET "
                          sSQLlat = sSQLlat & "latitude = "    & lcl_fieldvalue
                          sSQLlat = sSQLlat & " WHERE dmid = " & lcl_dmid
'dtb_debug(sSQLlat)
                          RunSQLStatement sSQLlat

                       elseif instr(lcl_fieldtype,"LONGITUDE") > 0 then

                        		sSQLlng = "UPDATE egov_dm_data SET "
                          sSQLlng = sSQLlng & "longitude = "   & lcl_fieldvalue
                          sSQLlng = sSQLlng & " WHERE dmid = " & lcl_dmid
'dtb_debug(sSQLlng)
                          RunSQLStatement sSQLlng

                       end if

                    elseif instr(lcl_fieldtype,"WEBSITE") > 0 OR instr(lcl_fieldtype,"EMAIL") > 0 then
                       if oGetImportRecords("dbcolumn_name") <> "" then
                          lcl_fieldvalue = "[" & oGetImportRecords("dbcolumn_name") & "]<>"
                       end if
                    else
                       lcl_fieldvalue = oGetImportRecords("dbcolumn_name")
                    end if

                    lcl_fieldvalue = formatFieldforInsertUpdate(lcl_fieldvalue)

                    lcl_dm_valueid = getDMValueID(lcl_orgid, _
                                                  oGetImportRecords("dm_typeid"), _
                                                  oGetImportRecords("dmid"), _
                                                  lcl_dm_sectionid, _
                                                  lcl_dm_fieldid)

'dtb_debug("dm_sectionid: [" & lcl_dm_sectionid & "] - dm_fieldid: [" & lcl_dm_fieldid & "] - fieldvalue: [" & lcl_fieldvalue & "] - dm_valueid: [" & lcl_dm_valueid & "]")

                    if lcl_dm_valueid > 0 then
'override existing values here!!!
                       sSQL = "UPDATE egov_dm_values SET "
                       sSQL = sSQL & "fieldvalue = "  & lcl_fieldvalue & ", "
                       sSQL = sSQL & "dm_importid = " & lcl_dm_importid
                       sSQL = sSQL & " WHERE dm_valueid = " & lcl_dm_valueid

                       set oUpdateDMValue = Server.CreateObject("ADODB.Recordset")
                       oUpdateDMValue.Open sSQL, Application("DSN"), 3, 1

                    else
                       sSQL = "INSERT INTO egov_dm_values ("
                       sSQL = sSQL & "orgid, "
                       sSQL = sSQL & "dm_typeid, "
                       sSQL = sSQL & "dmid, "
                       sSQL = sSQL & "dm_sectionid, "
                       sSQL = sSQL & "dm_fieldid, "
                       sSQL = sSQL & "fieldvalue, "
                       sSQL = sSQL & "dm_importid "
                       sSQL = sSQL & ") VALUES ("
                       sSQL = sSQL & lcl_orgid        & ", "
                       sSQL = sSQL & lcl_dm_typeid    & ", "
                       sSQL = sSQL & lcl_dmid         & ", "
                       sSQL = sSQL & lcl_dm_sectionid & ", "
                       sSQL = sSQL & lcl_dm_fieldid   & ", "
                       sSQL = sSQL & lcl_fieldvalue   & ", "
                       sSQL = sSQL & lcl_dm_importid
                       sSQL = sSQL & ")"

                       lcl_dm_valueid = RunIdentityInsert(sSQL)

                    end if

                    oGetImportRecords.movenext
                 loop
              end if

              oGetImportRecords.close
              set oGetImportRecords = nothing
             'END: Cycle through all of the records in the import table. ---------

              response.write "&nbsp;&nbsp;" & lcl_dbcolumn_name & " <span class=""redText"">[Complete]</span>"

           else
              response.write ""
           end if

       '-----------------------------------------------------------------------
        elseif lcl_action = "COMPLETE_IMPORT" then
           lcl_current_date = ConvertDateTimetoTimeZone()
           lcl_totalfields  = 0

           if request("totalfields") <> "" then
              lcl_totalfields = request("totalfields")

              if not isnumeric(lcl_totalfields) then
                 lcl_error_msg = "INVALID VALUE: Non-numeric value in 'totalfields'"
              else
                 lcl_totalfields = clng(lcl_totalfields)
              end if
           end if

           sSQLt = "SELECT count(dm_importid) as total_dm_imported "
           sSQLt = sSQLt & " FROM egov_dm_import_data "
           sSQLt = sSQLt & " WHERE dm_importid = " & lcl_dm_importid

           set oTotalDMImported = Server.CreateObject("ADODB.Recordset")
           oTotalDMImported.Open sSQLt, Application("DSN"), 3, 1

           if not oTotalDMImported.eof then
              lcl_totalfields = oTotalDMImported("total_dm_imported")
           end if

          'Create the import record
           sSQLu = "UPDATE egov_dm_import SET "
           sSQLu = sSQLu & " status = 'IMPORT COMPLETED', "
           sSQLu = sSQLu & " importenddate = '" & lcl_current_date & "', "
           sSQLu = sSQLu & " total_imported = " & lcl_totalfields
           sSQLu = sSQLu & " WHERE dm_importid = " & lcl_dm_importid

           set oCompleteImport = Server.CreateObject("ADODB.Recordset")
           oCompleteImport.Open sSQLu, Application("DSN"), 3, 1

           set oTotalDMImported = nothing
           set oCompleteImport  = nothing

           response.write "complete"

       '-----------------------------------------------------------------------
        elseif lcl_action = "GET_NONVALID_ADDRESSES" then
           sSQLnv = "SELECT "
           sSQLnv = sSQLnv & " dmid, "
           sSQLnv = sSQLnv & " dm_typeid, "
           sSQLnv = sSQLnv & " streetaddress, "
           sSQLnv = sSQLnv & " sortstreetname, "
           sSQLnv = sSQLnv & " city, "
           sSQLnv = sSQLnv & " state, "
           sSQLnv = sSQLnv & " latitude, "
           sSQLnv = sSQLnv & " longitude, "
           sSQLnv = sSQLnv & " dm_importid "
           sSQLnv = sSQLnv & " FROM egov_dm_data "
           'sSQLnv = sSQLnv & " WHERE dm_importid = " & lcl_dm_importid
           sSQLnv = sSQLnv & " WHERE validstreet = 'N' "
           sSQLnv = sSQLnv & " AND (latitude = 0 OR latitude is null OR longitude = 0 OR longitude is null) "

           if lcl_includeOnlyImportedNVA = "Y" then
              sSQLnv = sSQLnv & " AND dm_importid <> '' "
              sSQLnv = sSQLnv & " AND dm_importid IS NOT NULL "
           end if

           if lcl_sc_dm_importid <> "" then
              sSQLnv = sSQLnv & " AND dm_importid = " & lcl_sc_dm_importid
           end if

           sSQLnv = sSQLnv & " ORDER BY sortstreetname, streetaddress "

           set oGetNonValidAddresses = Server.CreateObject("ADODB.Recordset")
           oGetNonValidAddresses.Open sSQLnv, Application("DSN"), 3, 1

           if not oGetNonValidAddresses.eof then
              do while not oGetNonValidAddresses.eof
                 lcl_linecount  = lcl_linecount + 1
                 lcl_bgcolor    = changeBGColor(lcl_bgcolor,"#ffffff","#eeeeee")
                 lcl_snumber    = ""
                 lcl_sprefix    = ""
                 lcl_saddress   = oGetNonValidAddresses("streetaddress")
                 lcl_ssuffix    = ""
                 lcl_sdirection = ""
                 lcl_dm_address = buildStreetAddress(lcl_snumber, lcl_sprefix, lcl_saddress, lcl_ssuffix, lcl_sdirection)

                 lcl_address_row = lcl_address_row & "  <tr class=""nonValidAddressRow"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td style=""white-space: nowrap"">" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""hidden"" name=""nvaddresses_dmid"    & lcl_linecount & """ id=""nvaddresses_dmid"    & lcl_linecount & """ value=""" & oGetNonValidAddresses("dmid") & """ size=""10"" maxlength=""100"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""hidden"" name=""nvaddresses_address" & lcl_linecount & """ id=""nvaddresses_address" & lcl_linecount & """ value=""" & lcl_dm_address                & """ size=""10"" maxlength=""100"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & lcl_linecount & ". " & lcl_dm_address
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""text"" name=""nvaddresses_city" & lcl_linecount & """ id=""nvaddresses_city" & lcl_linecount & """ value=""" & oGetNonValidAddresses("city") & """ size=""20"" maxlength=""50"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""text"" name=""nvaddresses_state" & lcl_linecount & """ id=""nvaddresses_state" & lcl_linecount & """ value=""" & oGetNonValidAddresses("state") & """ size=""3"" maxlength=""20"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""text"" name=""nvaddresses_latitude" & lcl_linecount & """ id=""nvaddresses_latitude" & lcl_linecount & """ value=""" & oGetNonValidAddresses("latitude") & """ size=""20"" maxlength=""100"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""text"" name=""nvaddresses_longitude" & lcl_linecount & """ id=""nvaddresses_longitude" & lcl_linecount & """ value=""" & oGetNonValidAddresses("longitude") & """ size=""20"" maxlength=""100"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td align=""center"">" & vbcrlf
                 lcl_address_row = lcl_address_row &             oGetNonValidAddresses("dm_importid") & vbcrlf
                 lcl_address_row = lcl_address_row & "           <input type=""hidden"" name=""nvaddresses_dm_importid" & lcl_linecount & """ id=""nvaddresses_dm_importid" & lcl_linecount & """ value=""" & oGetNonValidAddresses("dm_importid") & """ size=""10"" maxlength=""100"" />" & vbcrlf
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       <td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "           <span id=""displayStatus" & lcl_linecount & """></span>" & vbcrlf
                 lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
                 lcl_address_row = lcl_address_row & "   </tr>" & vbcrlf

                 oGetNonValidAddresses.movenext
              loop
           end if

           lcl_address_row = lcl_address_row & "  <tr  class=""nonValidAddressTotalRow""bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
           lcl_address_row = lcl_address_row & "       <td align=""right"" colspan=""7"">" & vbcrlf
           lcl_address_row = lcl_address_row & "          <input type=""hidden"" name=""nvaddresses_total"" id=""nvaddresses_total"" value=""" & lcl_linecount & """ size=""5"" maxlength=""100"" />" & vbcrlf
           lcl_address_row = lcl_address_row & "          <input type=""hidden"" name=""nvaddresses_linenum"" id=""nvaddresses_linenum"" value=""0"" size=""5"" maxlength=""100"" />" & vbcrlf
           lcl_address_row = lcl_address_row & "          <strong>Total Non-Valid Addresses: </strong>[" & lcl_linecount & "]" & vbcrlf
           lcl_address_row = lcl_address_row & "       </td>" & vbcrlf
           lcl_address_row = lcl_address_row & "   </tr>" & vbcrlf

           set oGetNonValidAddresses = nothing

           response.write lcl_address_row

       '-----------------------------------------------------------------------
        elseif lcl_action = "UPDATE_LAT_LONG" then
           if request("dmid") <> "" then
              lcl_dmid = request("dmid")

              if not isnumeric(lcl_dmid) then
                 lcl_error_msg = "INVALID VALUE: Non-numeric value in 'DMID'"
              else
                 lcl_dmid = clng(lcl_dmid)
              end if
           end if

           if request("linecount") <> "" then
              lcl_linecount = request("linecount")

              if not isnumeric(lcl_linecount) then
                 lcl_error_msg = "INVALID VALUE: Non-numeric value in 'LineCount'"
              else
                 lcl_linecount = clng(lcl_linecount)
              end if
           end if

           if request("sc_latitude") <> "" then
              if not containsApostrophe(request("sc_latitude")) then
                 lcl_sc_latitude           = split(request("sc_latitude"), "_")
                 lcl_latitude_dm_typeid    = replace(lcl_sc_latitude(0),"dmtypeid","")
                 lcl_latitude_dm_sectionid = replace(lcl_sc_latitude(1),"dmsectionid","")
                 lcl_latitude_dm_fieldid   = replace(lcl_sc_latitude(2),"dmfieldid","")
              end if
           end if

           if request("sc_longitude") <> "" then
              if not containsApostrophe(request("sc_longitude")) then
                 lcl_sc_longitude           = split(request("sc_longitude"), "_")
                 lcl_longitude_dm_typeid    = replace(lcl_sc_longitude(0),"dmtypeid","")
                 lcl_longitude_dm_sectionid = replace(lcl_sc_longitude(1),"dmsectionid","")
                 lcl_longitude_dm_fieldid   = replace(lcl_sc_longitude(2),"dmfieldid","")
              end if
           end if

           if request("nvaddresses_latitude") <> "" then
              if not containsApostrophe(request("nvaddresses_latitude")) then
                 lcl_nvaddresses_latitude = request("nvaddresses_latitude")
              end if
           end if

           if request("nvaddresses_longitude") <> "" then
              if not containsApostrophe(request("nvaddresses_longitude")) then
                 lcl_nvaddresses_longitude = request("nvaddresses_longitude")
              end if
           end if

         		sSQLnvu = "UPDATE egov_dm_data SET "
    		     sSQLnvu = sSQLnvu & "latitude = "  & lcl_nvaddresses_latitude  & ", "
    		     sSQLnvu = sSQLnvu & "longitude = " & lcl_nvaddresses_longitude
    		     sSQLnvu = sSQLnvu & " WHERE dmid = " & lcl_dmid

    		   		set oUpdateDMLatLng = Server.CreateObject("ADODB.Recordset")
    		    	oUpdateDMLatLng.Open sSQLnvu, Application("DSN"), 3, 1

          'Now we need to update the egov_dm_values for the latitude and longitude
'dtb_debug("here")
'dtb_debug("latitude: userid: [" & lcl_userid & "] - orgid: [" & lcl_orgid & "] - dm_typeid: [" & lcl_latitude_dm_typeid & "] - dmid: [" & lcl_dmid & "] - dm_sectionid: [" & lcl_latitude_dm_sectionid & "] - dm_fieldid: [" & lcl_latitude_dm_fieldid & "] - dm_valueid: [" & lcl_dm_valueid & "] - nvaddresses_latitude: [" & lcl_nvaddresses_latitude & "] - mp_valueid: [" & lcl_mp_valueid & "] - dm_importid: [" & lcl_dm_importid & "]")
           maintainDMValues lcl_userid, _
                            lcl_orgid, _
                            lcl_latitude_dm_typeid, _
                            lcl_dmid, _
                            lcl_latitude_dm_sectionid, _
                            lcl_latitude_dm_fieldid, _
                            lcl_dm_valueid, _
                            lcl_nvaddresses_latitude, _
                            lcl_mp_valueid, _
                            lcl_dm_importid
'dtb_debug("latitude: userid: [" & lcl_userid & "] - orgid: [" & lcl_orgid & "] - dm_typeid: [" & lcl_longitude_dm_typeid & "] - dmid: [" & lcl_dmid & "] - dm_sectionid: [" & lcl_longitude_dm_sectionid & "] - dm_fieldid: [" & lcl_longitude_dm_fieldid & "] - dm_valueid: [" & lcl_dm_valueid & "] - nvaddresses_latitude: [" & lcl_nvaddresses_longitude & "] - mp_valueid: [" & lcl_mp_valueid & "] - dm_importid: [" & lcl_dm_importid & "]")
           maintainDMValues lcl_userid, _
                            lcl_orgid, _
                            lcl_longitude_dm_typeid, _
                            lcl_dmid, _
                            lcl_longitude_dm_sectionid, _
                            lcl_longitude_dm_fieldid, _
                            lcl_dm_valueid, _
                            lcl_nvaddresses_longitude, _
                            lcl_mp_valueid, _
                            lcl_dm_importid

           response.write lcl_linecount

       '-----------------------------------------------------------------------
        elseif lcl_action = "CANCEL_IMPORT" then
           if lcl_dm_importid <> "" then
              sSQL1 = "DELETE FROM egov_dm_values WHERE dm_importid = "              & lcl_dm_importid
              sSQL2 = "DELETE FROM egov_dm_data WHERE dm_importid = "                & lcl_dm_importid
              sSQL3 = "DELETE FROM egov_dmdata_to_dmcategories WHERE dm_importid = " & lcl_dm_importid
              sSQL4 = "DELETE FROM egov_dm_categories WHERE dm_importid = "          & lcl_dm_importid
              sSQL5 = "DELETE FROM egov_dm_import WHERE dm_importid = "              & lcl_dm_importid

              sSQL6 = "UPDATE egov_dm_import_data SET "
              sSQL6 = sSQL6 & " dm_importid = 0, "
              sSQL6 = sSQL6 & " orgid = NULL, "
              sSQL6 = sSQL6 & " dm_typeid = NULL, "
              sSQL6 = sSQL6 & " dmid = NULL, "
              sSQL6 = sSQL6 & " categoryid = NULL, "
              sSQL6 = sSQL6 & " subcategoryid = NULL"

              RunSQLStatement sSQL1
              RunSQLStatement sSQL2
              RunSQLStatement sSQL3
              RunSQLStatement sSQL4
              RunSQLStatement sSQL5
              RunSQLStatement sSQL6
           end if

           response.write "import cancelled"

        end if
       'END: Determine which action to take in the import process -------------
     end if
  end if

'------------------------------------------------------------------------------
sub getCategoryID(ByVal iUserID, ByVal iOrgID, ByVal iDMTypeID, ByVal iDMImportID, ByVal iCategory, _
                       ByRef lcl_categoryid, ByRef lcl_category_created)

  dim sSQL, sUserID, sOrgID, sDMTypeID, sDMImportID, sCategory
  dim lcl_current_date, lcl_parent_categoryid, lcl_isActive, lcl_isApproved, lcl_mappointcolor

  lcl_current_date      = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  sUserID               = 0
  sOrgID                = 0
  sDMTypeID             = 0
  sDMImportID           = 0
  sCategory             = ""
  lcl_isActive          = 1
  lcl_parent_categoryid = 0
  lcl_isApproved        = 1
  lcl_mappointcolor     = "NULL"

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iDMImportID <> "" then
     sDMImportID = clng(iDMImportID)
  end if

  if iCategory <> "" then
     sCategory = iCategory
     sCategory = dbsafe(sCategory)
     sCategory = "'" & sCategory & "'"
  end if

  if sOrgID > 0 AND sDMTypeID > 0 AND sCategory <> "" then
     sSQLu = "SELECT categoryid "
     sSQLu = sSQLu & " FROM egov_dm_categories "
     sSQLu = sSQLu & " WHERE orgid = " & sOrgID
     sSQLu = sSQLu & " AND dm_typeid = " & sDMTypeID
     sSQLu = sSQLu & " AND upper(categoryname) = " & sCategory
     sSQLu = sSQLu & " AND isActive = 1 "
     sSQLu = sSQLu & " AND parent_categoryid = 0 "

     set oGetDMCatID = Server.CreateObject("ADODB.Recordset")
     oGetDMCatID.Open sSQLu, Application("DSN"), 3, 1

     if not oGetDMCatID.eof then
        do while not oGetDMCatID.eof

           lcl_categoryid       = oGetDMCatID("categoryid")
           lcl_category_created = 0

           oGetDMCatID.movenext
        loop
     else
        sCreatedByID    = sUserID
        sCreatedByDate  = lcl_current_date
        sApprovedByID   = sUserID
        sApprovedByDate = lcl_current_date

       'Since we are creating the category from the admin-side we want to automatically
       '"approve" the category when it's created.
        lcl_isApproved = 1

     		'Insert the new Category
   	   	sSQLi = "INSERT INTO egov_dm_categories ("
        sSQLi = sSQLi & "categoryname, "
        sSQLi = sSQLi & "orgid, "
        sSQLi = sSQLi & "dm_typeid, "
        sSQLi = sSQLi & "isActive, "
        sSQLi = sSQLi & "createdbyid, "
        sSQLi = sSQLi & "createdbydate, "
        sSQLi = sSQLi & "lastmodifiedbyid, "
        sSQLi = sSQLi & "lastmodifiedbydate, "
        sSQLi = sSQLi & "parent_categoryid, "
        sSQLi = sSQLi & "isApproved, "
        sSQLi = sSQLi & "approvedeniedbyid, "
        sSQLi = sSQLi & "approvedeniedbydate, "
        sSQLi = sSQLi & "mappointcolor, "
        sSQLi = sSQLi & "dm_importid"
        sSQLi = sSQLi & ") VALUES ("
        sSQLi = sSQLi & sCategory             & ", "
        sSQLi = sSQLi & sOrgID                & ", "
        sSQLi = sSQLi & sDMTypeID             & ", "
        sSQLi = sSQLi & lcl_isActive          & ", "
        sSQLi = sSQLi & sCreatedByID          & ", "
        sSQLi = sSQLi & sCreatedByDate        & ", "
        sSQLi = sSQLi & "NULL,NULL"           & ", "
        sSQLi = sSQLi & lcl_parent_categoryid & ", "
        sSQLi = sSQLi & lcl_isApproved        & ", "
        sSQLi = sSQLi & sApprovedByID         & ", "
        sSQLi = sSQLi & sApprovedByDate       & ", "
        sSQLi = sSQLi & lcl_mappointcolor     & ", "
        sSQLi = sSQLi & sDMImportID
        sSQLi = sSQLi & ")"

     		'Get the categoryid
    	  	lcl_categoryid = RunIdentityInsert(sSQLi)

        lcl_category_created = 1

     end if

     set oGetDMCatID = nothing

  end if

end sub

'------------------------------------------------------------------------------
 sub getSubCategoryID(ByVal iUserID, ByVal iOrgID, ByVal iDMTypeID, ByVal iDMImportID, _
                      ByVal iSubCategory, ByVal iCategoryID, _
                      ByRef lcl_subcategoryid, ByRef lcl_subcategory_created)

  dim sSQL, sUserID, sOrgID, sDMTypeID, sDMImportID, sSubCategory, sCategoryID
  dim lcl_current_date, lcl_parent_categoryid, lcl_isActive, lcl_isApproved, lcl_mappointcolor

  lcl_current_date      = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  sUserID               = 0
  sOrgID                = 0
  sDMTypeID             = 0
  sDMImportID           = 0
  sSubCategory          = ""
  sCategoryID           = ""
  lcl_isActive          = 1
  lcl_parent_categoryid = 0
  lcl_isApproved        = 1
  lcl_mappointcolor     = "NULL"

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  if iDMImportID <> "" then
     sDMImportID = clng(iDMImportID)
  end if

  if iSubCategory <> "" then
     sSubCategory = iSubCategory
     sSubCategory = dbsafe(sSubCategory)
     sSubCategory = "'" & sSubCategory & "'"
  end if

  if iCategoryID <> "" then
     sCategoryID = clng(iCategoryID)
  end if

  if sOrgID > 0 AND sDMTypeID > 0 AND sSubCategory <> "" AND sCategoryID <> "" then
     sSQLu = "SELECT categoryid "
     sSQLu = sSQLu & " FROM egov_dm_categories "
     sSQLu = sSQLu & " WHERE orgid = " & sOrgID
     sSQLu = sSQLu & " AND dm_typeid = " & sDMTypeID
     sSQLu = sSQLu & " AND upper(categoryname) = " & sSubCategory
     sSQLu = sSQLu & " AND isActive = 1 "

     if sCategoryID <> "" then
        sSQLu = sSQLu & " AND parent_categoryid = " & sCategoryID
     end if

     set oGetDMSubCatID = Server.CreateObject("ADODB.Recordset")
     oGetDMSubCatID.Open sSQLu, Application("DSN"), 3, 1

     if not oGetDMSubCatID.eof then

        do while not oGetDMSubCatID.eof

           lcl_subcategoryid       = oGetDMSubCatID("categoryid")
           lcl_subcategory_created = 0

           oGetDMSubCatID.movenext
        loop
     else
        sCreatedByID          = sUserID
        sCreatedByDate        = lcl_current_date
        sApprovedByID         = sUserID
        sApprovedByDate       = lcl_current_date
        lcl_parent_categoryid = sCategoryID

       'Since we are creating the category from the admin-side we want to automatically
       '"approve" the category when it's created.
        lcl_isApproved = 1

     		'Insert the new Category
   	   	sSQLi = "INSERT INTO egov_dm_categories ("
        sSQLi = sSQLi & "categoryname, "
        sSQLi = sSQLi & "orgid, "
        sSQLi = sSQLi & "dm_typeid, "
        sSQLi = sSQLi & "isActive, "
        sSQLi = sSQLi & "createdbyid, "
        sSQLi = sSQLi & "createdbydate, "
        sSQLi = sSQLi & "lastmodifiedbyid, "
        sSQLi = sSQLi & "lastmodifiedbydate, "
        sSQLi = sSQLi & "parent_categoryid, "
        sSQLi = sSQLi & "isApproved, "
        sSQLi = sSQLi & "approvedeniedbyid, "
        sSQLi = sSQLi & "approvedeniedbydate, "
        sSQLi = sSQLi & "mappointcolor, "
        sSQLi = sSQLi & "dm_importid"
        sSQLi = sSQLi & ") VALUES ("
        sSQLi = sSQLi & sSubCategory          & ", "
        sSQLi = sSQLi & sOrgID                & ", "
        sSQLi = sSQLi & sDMTypeID             & ", "
        sSQLi = sSQLi & lcl_isActive          & ", "
        sSQLi = sSQLi & sCreatedByID          & ", "
        sSQLi = sSQLi & sCreatedByDate        & ", "
        sSQLi = sSQLi & "NULL,NULL"           & ", "
        sSQLi = sSQLi & lcl_parent_categoryid & ", "
        sSQLi = sSQLi & lcl_isApproved        & ", "
        sSQLi = sSQLi & sApprovedByID         & ", "
        sSQLi = sSQLi & sApprovedByDate       & ", "
        sSQLi = sSQLi & lcl_mappointcolor     & ", "
        sSQLi = sSQLi & sDMImportID
        sSQLi = sSQLi & ")"

     		'Get the categoryid
    	  	lcl_subcategoryid = RunIdentityInsert(sSQLi)

        lcl_subcategory_created = 1

     end if

     set oGetDMSubCatID = nothing

  end if

end sub

'------------------------------------------------------------------------------
sub updateDMImport(iUpdateColumn, iUpdateValue, iSearchColumn, iSearchValue)

  dim sUpdateColumn, sUpdateValue, sSearchColumn, sSearchValue

  sUpdateColumn = ""
  sUpdateValue  = ""
  sSearchColumn = ""
  sSearchValue  = ""

  if iUpdateColumn <> "" then
     sUpdateColumn = iUpdateColumn
     sUpdateColumn = dbsafe(sUpdateColumn)
  end if

  if iUpdateValue <> "" then
     sUpdateValue = ucase(iUpdateValue)
     sUpdateValue = dbsafe(sUpdateValue)
     sUpdateValue = "'" & sUpdateValue & "'"
  end if

  if iSearchColumn <> "" then
     if not containsApostrophe(iSearchColumn) then
        sSearchColumn = iSearchColumn
        sSearchColumn = "upper(" & sSearchColumn & ")"
     end if
  end if

  if iSearchValue <> "" then
     sSearchValue = ucase(iSearchValue)
     sSearchValue = dbsafe(sSearchValue)
     sSearchValue = "'" & sSearchValue & "'"
  end if

  if sUpdateColumn <> "" AND sSearchColumn <> "" AND sSearchValue <> "" then
     sSQLimport = "UPDATE egov_dm_import_data SET "
     sSQLimport = sSQLimport & sUpdateColumn & " = " & sUpdateValue
     sSQLimport = sSQLimport & " WHERE " & sSearchColumn & " = " & sSearchValue

     set oUpdateDMImport = Server.CreateObject("ADODB.Recordset")
     oUpdateDMImport.Open sSQLimport, Application("DSN"), 3, 1

     set oUpdateDMImport = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(iValue)

  sValue = ""

  if iValue <> "" then
     sValue =  dbsafe(iValue)
  end if

  sValue = "'" & sValue & "'"

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES (" & sValue & ") "

 	lcl_new_id = RunIdentityInsert(sSQL)


end sub
%>