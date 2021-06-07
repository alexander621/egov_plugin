<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  lcl_orghasfeature_issue_location     = orghasfeature("issue location")
  lcl_orghasfeature_large_address_list = orghasfeature("large address list")

  lcl_dmid        = 0
  lcl_dm_typeid   = 0
  lcl_sectionid   = 0
  lcl_totalfields = 0
  lcl_orgid       = request("orgid")
  lcl_feature     = "datamgr_maint"
  lcl_featurename = getFeatureName(lcl_feature)

  if request("f") <> "" then
     if not containsApostrophe(request("f")) then
        lcl_feature     = request("f")
        lcl_featurename = getFeatureName(lcl_feature)
     end if
  end if

  if not userhaspermission(session("userid"),lcl_feature) then
    	response.redirect sLevel & "permissiondenied.asp?f=" & lcl_feature
  end if

  if request("dmid") <> "" then
     lcl_dmid = request("dmid")
  end if

  if request("dm_typeid") <> "" then
     lcl_dm_typeid = request("dm_typeid")
  end if

  if request("sectionid") <> "" then
     lcl_sectionid = request("sectionid")
  end if

  if request("totalfields") <> "" then
     lcl_totalfields = request("totalfields")
  end if

 'BEGIN: Cycle through all of the fields in the section. ----------------------
 'We need to perform an initial check to see if a record exists for each field on the 
 '   egov_datamgr_values table.  This table is where we store the value for each field.
 'If there is NOT a "dm_valueid" then we need to INSERT a record.
 'If there IS a "dm_valueid" then we need to update the record.
  if lcl_totalfields > 0 then
     for v = 1 to lcl_totalfields

        lcl_dm_valueid   = 0
        lcl_dm_sectionid = 0
        lcl_dm_fieldid   = 0
        lcl_fieldtype    = ""
        lcl_fieldvalue   = ""

        if request("dm_valueid" & v) <> "" then
           lcl_dm_valueid = request("dm_valueid" & v)
        end if

        if request("dm_sectionid" & v) <> "" then
           lcl_dm_sectionid = request("dm_sectionid" & v)
        end if

        if request("dm_fieldid" & v) <> "" then
           lcl_dm_fieldid = request("dm_fieldid" & v)
        end if

        if request("fieldtype" & v) <> "" then
           lcl_fieldtype = ucase(request("fieldtype" & v))
        end if

'        if request("fieldvalue" & v) <> "" then
'           lcl_fieldvalue = request("fieldvalue" & v)
'           lcl_fieldvalue = "'" & dbsafe(lcl_fieldvalue) & "'"
'        end if

       'BEGIN: Check to see if the address is valid or custom --------------------
        if lcl_fieldtype = "ADDRESS" OR lcl_fieldtype = "LATITUDE" OR lcl_fieldtype = "LONGITUDE" then
           if lcl_orghasfeature_issue_location then
              lcl_address = ""

              if trim(request("ques_issue2")) <> "" then
                 sAddress        = request("ques_issue2")
                 lcl_dmv_address = sAddress
                 sLatitude       = request("latitude")
                 sLongitude      = request("longitude")
              else
                 lcl_streetnumber  = ""
                 lcl_streetaddress = ""
                 lcl_latitude      = ""
                 lcl_longitude     = ""

                 if request("residentstreetnumber") <> "" then
                    lcl_address = request("residentstreetnumber")
                 end if

                 if request("streetaddress") <> "" then
                    if lcl_address <> "" then
                       lcl_address = lcl_address & " " & request("streetaddress")
                    else
                       lcl_address = request("streetaddress")
                    end if
                 end if

                 breakOutAddress lcl_address, _
                                 sStreetNumber,_
                                 sStreetName

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

                 lcl_dmv_address = buildStreetAddress(sNumber, _
                                                      sPrefix, _
                                                      sAddress, _
                                                      sSuffix, _
                                                      sDirection)
              end if
           else
              if lcl_fieldtype = "ADDRESS" then
                 sNumber    = ""
                 sAddress   = request("ques_issue2")
                 sPrefix    = ""
                 sSuffix    = ""
                 sDirection = ""

                 lcl_dmv_address = sAddress
              elseif lcl_fieldtype = "LATITUDE" then
                 sLatitude = request("latitude")
              elseif lcl_fieldtype = "LONGITUDE" then
                 sLongitude = request("longitude")
              end if
           end if

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
        'elseif lcl_fieldtype = "LATITUDE" then
        '   sLatitude = request("latitude")
        'elseif lcl_fieldtype = "LONGITUDE" then
        '   sLongitude = request("longitude")
        end if

       'Determine if we pull the value from the screen or if we generate it.
        if lcl_fieldtype = "ADDRESS" then
           lcl_fieldvalue = lcl_dmv_address
        elseif lcl_fieldtype = "LATITUDE" then
           lcl_fieldvalue = sLatitude
        elseif lcl_fieldtype = "LONGITUDE" then
           lcl_fieldvalue = sLongitude
        else
           lcl_fieldvalue = request("dm_fieldvalue" & v)
        end if

       'Make one last check for the Latitude/Longitude fields.
       'If they are NULL then default the value to 0.00
        if lcl_fieldtype = "LATITUDE" OR lcl_fieldtype = "LONGITUDE" then
           if lcl_fieldvalue = "" then
              lcl_fieldvalue = "0.00"
           end if
        end if

        lcl_fieldvalue = formatFieldforInsertUpdate(lcl_fieldvalue)

        if lcl_dm_valueid > 0 then
           sSQL = "UPDATE egov_dm_values SET "
           sSQL = sSQL & "fieldvalue = " & lcl_fieldvalue
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
           sSQL = sSQL & "fieldvalue "
           sSQL = sSQL & ") VALUES ("
           sSQL = sSQL & lcl_orgid           & ", "
           sSQL = sSQL & lcl_dm_typeid & ", "
           sSQL = sSQL & lcl_dmid      & ", "
           sSQL = sSQL & lcl_dm_sectionid    & ", "
           sSQL = sSQL & lcl_dm_fieldid      & ", "
           sSQL = sSQL & lcl_fieldvalue
           sSQL = sSQL & ")"

           lcl_dm_valueid = RunIdentityInsert(sSQL)

        end if

       'Determine if this is the "ADDRESS" section.
       'If "yes" then we need to handle updating the address fields (i.e. breakout of address, latitude, longitude, etc)
       'NOTE: The values entered for this section are stored on egov_dm_data.  These values on this table are the "core"
       'values to update.  The values for this section/fields are also stored on egov_dm_values.  Storing them 
       'on egov_dm_values simply let's us search and display the value easier.
        if lcl_fieldtype = "ADDRESS" then
           if lcl_orghasfeature_issue_location then
              lcl_address    = ""
              lcl_latitude   = ""
              lcl_longitude  = ""

             	if trim(request("ques_issue2")) <> "" then
                 lcl_address   = request("ques_issue2")
                 lcl_latitude  = request("latitude")
                 lcl_longitude = request("longitude")
              else
                 lcl_streetnumber  = ""
                 lcl_streetaddress = ""
                 lcl_latitude      = ""
                 lcl_longitude     = ""

                 if lcl_orghasfeature_large_address_list then
                    if request("residentstreetnumber") <> "" then
                       lcl_address = request("residentstreetnumber")
                    end if
                 end if

                 if request("streetaddress") <> "" then
                    if lcl_address <> "" then
                       lcl_address = lcl_address & " " & request("streetaddress")
                    else
                       lcl_address = request("streetaddress")
                    end if
                 end if

                 breakOutAddress lcl_address, sStreetNumber, sStreetName

                 getAddressInfoNew lcl_orghasfeature_large_address_list, lcl_orgid, sStreetNumber, sStreetName, _
                                   sNumber, sPrefix, sAddress, sSuffix, sDirection, sLatitude, sLongitude, sCity, sState, sZip, _
                                   sCounty, sParcelID, sListedOwner, sLegalDescription, sResidentType, sRegisteredUserID, sValidStreet

                 'lcl_address = buildStreetAddress(sNumber, sPrefix, sAddress, sSuffix, sDirection)

                 if sValidStreet = "Y" then
                    lcl_latitude  = sLatitude
                    lcl_longitude = sLongitude
                 else
                    lcl_latitude  = request("latitude")
                    lcl_longitude = request("longitude")
                 end if
              end if
           else
              sNumber       = ""
              sAddress      = request("ques_issue2")
              sPrefix       = ""
              sSuffix       = ""
              sDirection    = ""
              lcl_latitude  = request("latitude")
              lcl_longitude = request("longitude")
           end if

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

           maintainDMSection_address session("userid"), lcl_orgid, lcl_dmid, lcl_dm_typeid, lcl_dm_sectionid, _
                                     lcl_dm_fieldid, sNumber, sPrefix, sAddress, sSuffix, sDirection, sSortStreetName, _
                                     sCity, sStatus, sZip, sValidStreet, lcl_latitude, lcl_longitude, lcl_dmid
        end if

     next
  end if
 'END: Cycle through all of the fields in the section. ------------------------

  response.redirect "maintain_dmt_section.asp?dmid=" & lcl_dmid & "&sectionid=" & lcl_sectionid & "&f=" & lcl_feature & "&success=SU"
%>