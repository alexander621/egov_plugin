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
  dim lcl_dm_sectionid_current, lcl_dm_sectionid_new
  dim lcl_dm_fieldid_current, lcl_dm_fieldid_new
  dim lcl_userid, lcl_orgid, lcl_mappoint_typeid, lcl_dm_typeid
  dim lcl_mp_fieldid, lcl_transfer_field_data, lcl_dm_sectionid, lcl_dm_fieldid
  dim i, lcl_error_msg, lcl_return_msg, lcl_totalfields, lcl_overrideValues

  lcl_userid              = ""
  lcl_orgid               = ""
  lcl_mappoint_typeid     = ""
  lcl_dm_typeid           = ""
  lcl_mp_fieldid          = ""
  lcl_transfer_field_data = ""
  lcl_dm_sectionid        = ""
  lcl_dm_fieldid          = ""
  lcl_totalfields         = 0
  lcl_action              = ""
  lcl_overrideValues      = "N"
  lcl_error_msg           = ""
  lcl_return_msg          = ""

  if request("userid") <> "" then
     lcl_userid = request("userid")

     if not isnumeric(lcl_userid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'userid'"
     else
        lcl_userid = clng(lcl_userid)
     end if
  end if

  if request("orgid") <> "" then
     lcl_orgid = request("orgid")

     if not isnumeric(lcl_orgid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'orgid'"
     else
        lcl_orgid = clng(lcl_orgid)
     end if
  end if

  if request("totalfields") <> "" then
     lcl_totalfields = request("totalfields")

     if not isnumeric(lcl_totalfields) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'totalfields'"
     else
        lcl_totalfields = clng(lcl_totalfields)
     end if
  end if

  if request("mappoint_typeid") <> "" then
     lcl_mappoint_typeid = request("mappoint_typeid")

     if not isnumeric(lcl_mappoint_typeid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'mappoint_typeid'"
     else
        lcl_mappoint_typeid = clng(lcl_mappoint_typeid)
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

  if request("mp_fieldid") <> "" then
     lcl_mp_fieldid = request("mp_fieldid")

     if not isnumeric(lcl_mp_fieldid) then
        lcl_error_msg = "INVALID VALUE: Non-numeric value in 'mp_fieldid'"
     else
        lcl_mp_fieldid = clng(lcl_mp_fieldid)
     end if
  end if

  if request("transfer_field_data") <> "" then
     if not containsApostrophe(request("transfer_field_data")) then
        lcl_transfer_field_data = split(request("transfer_field_data"), "_")
        lcl_dm_sectionid        = replace(lcl_transfer_field_data(0),"dmsectionid","")
        lcl_dm_fieldid          = replace(lcl_transfer_field_data(1),"dmfieldid","")
     end if
  end if

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

 'Check for an error message.
  if lcl_error_msg <> "" then
     response.write lcl_error_msg
  else
     if lcl_action <> "" then

       'BEGIN: Build the query to retrieve all mappointids to be imported -----
        sSQL = "select distinct "
        sSQL = sSQL & "mpv.mappointid, "
        sSQL = sSQL & "mp.dmid, "

        if lcl_action = "IMPORT_MP_VALUES" then
           sSQL = sSQL & "mpv.mp_valueid, "
           sSQL = sSQL & "mptf.fieldname, "
           sSQL = sSQL & "mpv.fieldvalue, "
           sSQL = sSQL & "mp.dmid "
           'sSQL = sSQL & "mptf.mp_fieldid, "
           'sSQL = sSQL & "mptf.fieldtype, "
           'sSQL = sSQL & "mptf.displayInResults, "
           'sSQL = sSQL & "mptf.displayInInfoPage, "
           'sSQL = sSQL & "mptf.resultsOrder, "
           'sSQL = sSQL & "mptf.inPublicSearch, "
        else
           sSQL = sSQL & "mp.createdbyid, "
           sSQL = sSQL & "mp.createdbydate, "
           sSQL = sSQL & "mp.lastmodifiedbyid, "
           sSQL = sSQL & "mp.lastmodifiedbydate, "
           sSQL = sSQL & "mp.isActive, "
           sSQL = sSQL & "mp.streetnumber, "
           sSQL = sSQL & "mp.streetprefix, "
           sSQL = sSQL & "mp.streetaddress, "
           sSQL = sSQL & "mp.streetsuffix, "
           sSQL = sSQL & "mp.streetdirection, "
           sSQL = sSQL & "mp.sortstreetname, "
           sSQL = sSQL & "mp.city, "
           sSQL = sSQL & "mp.state, "
           sSQL = sSQL & "mp.zip, "
           sSQL = sSQL & "mp.validstreet, "
           sSQL = sSQL & "mp.latitude, "
           sSQL = sSQL & "mp.longitude, "
           sSQL = sSQL & "mp.mappointcolor "
        end if

        sSQL = sSQL & "from egov_mappoints_values mpv "
        sSQL = sSQL & "     inner join egov_mappoints mp ON mp.mappointid = mpv.mappointid "
        sSQL = sSQL & "     inner join egov_mappoints_types_fields mptf ON mptf.mp_fieldid = mpv.mp_fieldid "
        sSQL = sSQL & "where mptf.mappoint_typeid = mpv.mappoint_typeid "
        sSQL = sSQL & "and mptf.mappoint_typeid = " & lcl_mappoint_typeid
        sSQL = sSQL & "and mptf.orgid = " & lcl_orgid

        if lcl_action = "IMPORT_MP_VALUES" then
           sSQL = sSQL & "and mptf.mp_fieldid = " & lcl_mp_fieldid
           sSQL = sSQL & "order by mpv.mappointid, mpv.mp_valueid "
        else
           sSQL = sSQL & "order by mpv.mappointid "
        end if
       'END: Build the query to retrieve all mappointids to be imported -------

       'BEGIN: Determine which action to take in the import process -----------
        if lcl_action = "CREATE_DMID" then
           i = 0
           lcl_totalRowsSelected = 0

           if lcl_totalfields > 0 then
             'First: Cycle through each MapPoint and create a DMID
            		set oGetTransferMPInfo = Server.CreateObject("ADODB.Recordset")
             	oGetTransferMPInfo.Open sSQL, Application("DSN"), 3, 1

              if not oGetTransferMPInfo.eof then
                 lcl_rowCount            = 0
                 lcl_previous_mappointid = 0
                 lcl_totalDMIDsCreated   = 0
                 lcl_totalDMIDsExisting  = 0

                 do while not oGetTransferMPInfo.eof
                    lcl_rowCount           = lcl_rowCount + 1
                    lcl_mappointid         = "NULL"
                    lcl_dm_exists_for_mpid = ""

                    lcl_isCreatedByAdmin   = 1
                    lcl_isApproved         = 1
                    lcl_approvedeniedbyid  = lcl_userid
                    lcl_approvedeniedbydate = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
                    lcl_createdbyid        = "NULL"
                    lcl_createdbydate      = "NULL"
                    lcl_lastmodifiedbyid   = "NULL"
                    lcl_lastmodifiedbydate = "NULL"
                    lcl_isActive           = "0"
                    lcl_streetnumber       = "NULL"
                    lcl_streetprefix       = "NULL"
                    lcl_streetaddress      = "NULL"
                    lcl_streetsuffix       = "NULL"
                    lcl_streetdirection    = "NULL"
                    lcl_sortstreetname     = "NULL"
                    lcl_city               = "NULL"
                    lcl_state              = "NULL"
                    lcl_zip                = "NULL"
                    lcl_validstreet        = "NULL"
                    lcl_latitude           = "NULL"
                    lcl_longitude          = "NULL"
                    lcl_mappointcolor      = "NULL"

                    if oGetTransferMPInfo("createdbyid") <> "" then
                       lcl_createdbyid = oGetTransferMPInfo("createdbyid")
                    end if

                    if oGetTransferMPInfo("createdbydate") <> "" then
                       lcl_createdbydate = oGetTransferMPInfo("createdbydate")
                       lcl_createdbydate = dbsafe(lcl_createdbydate)
                       lcl_createdbydate = "'" & lcl_createdbydate & "'"
                    end if

                    if oGetTransferMPInfo("lastmodifiedbyid") <> "" then
                       lcl_lastmodifiedbyid = oGetTransferMPInfo("lastmodifiedbyid")
                    end if

                    if oGetTransferMPInfo("lastmodifiedbydate") <> "" then
                       lcl_lastmodifiedbydate = oGetTransferMPInfo("lastmodifiedbydate")
                       lcl_lastmodifiedbydate = dbsafe(lcl_lastmodifiedbydate)
                       lcl_lastmodifiedbydate = "'" & lcl_lastmodifiedbydate & "'"
                    end if

                    if oGetTransferMPInfo("isActive") then
                       lcl_isActive = "1"
                    end if

                    if oGetTransferMPInfo("streetnumber") <> "" then
                       lcl_streetnumber = oGetTransferMPInfo("streetnumber")
                       lcl_streetnumber = dbsafe(lcl_streetnumber)
                       lcl_streetnumber = "'" & lcl_streetnumber & "'"
                    end if

                    if oGetTransferMPInfo("streetprefix") <> "" then
                       lcl_streetprefix = oGetTransferMPInfo("streetprefix")
                       lcl_streetprefix = dbsafe(lcl_streetprefix)
                       lcl_streetprefix = "'" & lcl_streetprefix & "'"
                    end if

                    if oGetTransferMPInfo("streetaddress") <> "" then
                       lcl_streetaddress = oGetTransferMPInfo("streetaddress")
                       lcl_streetaddress = dbsafe(lcl_streetaddress)
                       lcl_streetaddress = "'" & lcl_streetaddress & "'"
                    end if

                    if oGetTransferMPInfo("streetsuffix") <> "" then
                       lcl_streetsuffix = oGetTransferMPInfo("streetsuffix")
                       lcl_streetsuffix = dbsafe(lcl_streetsuffix)
                       lcl_streetsuffix = "'" & lcl_streetsuffix & "'"
                    end if

                    if oGetTransferMPInfo("streetdirection") <> "" then
                       lcl_streetdirection = oGetTransferMPInfo("streetdirection")
                       lcl_streetdirection = dbsafe(lcl_streetdirection)
                       lcl_streetdirection = "'" & lcl_streetdirection & "'"
                    end if

                    if oGetTransferMPInfo("sortstreetname") <> "" then
                       lcl_sortstreetname = oGetTransferMPInfo("sortstreetname")
                       lcl_sortstreetname = dbsafe(lcl_sortstreetname)
                       lcl_sortstreetname = "'" & lcl_sortstreetname & "'"
                    end if

                    if oGetTransferMPInfo("city") <> "" then
                       lcl_city = oGetTransferMPInfo("city")
                       lcl_city = dbsafe(lcl_city)
                       lcl_city = "'" & lcl_city & "'"
                    end if

                    if oGetTransferMPInfo("state") <> "" then
                       lcl_state = oGetTransferMPInfo("state")
                       lcl_state = dbsafe(lcl_state)
                       lcl_state = "'" & lcl_state & "'"
                    end if

                    if oGetTransferMPInfo("zip") <> "" then
                       lcl_zip = oGetTransferMPInfo("zip")
                       lcl_zip = dbsafe(lcl_zip)
                       lcl_zip = "'" & lcl_zip & "'"
                    end if

                    if oGetTransferMPInfo("validstreet") <> "" then
                       lcl_validstreet = oGetTransferMPInfo("validstreet")
                       lcl_validstreet = dbsafe(lcl_validstreet)
                       lcl_validstreet = "'" & lcl_validstreet & "'"
                    end if

                    if oGetTransferMPInfo("latitude") <> "" then
                       lcl_latitude = oGetTransferMPInfo("latitude")
                    end if

                    if oGetTransferMPInfo("longitude") <> "" then
                       lcl_longitude = oGetTransferMPInfo("longitude")
                    end if

                    if oGetTransferMPInfo("mappointcolor") <> "" then
                       lcl_mappointcolor = oGetTransferMPInfo("mappointcolor")
                       lcl_mappointcolor = dbsafe(lcl_mappointcolor)
                       lcl_mappointcolor = "'" & lcl_mappointcolor & "'"
                    end if

                   'Check if the MapPointID has already been imported.
                    if oGetTransferMPInfo("mappointid") <> "" then
                       lcl_mappointid         = oGetTransferMPInfo("mappointid")
                       lcl_dm_exists_for_mpid = getDMID_by_mappointid(lcl_mappointid)
                    end if

                    if lcl_dm_exists_for_mpid <> "" then
                      'This MapPointID has already been imported and has a DMID.
                       lcl_dmid               = lcl_dm_exists_for_mpid
                       lcl_totalDMIDsExisting = lcl_totalDMIDsExisting + 1

                       if lcl_overrideValues = "Y" then
                          sSQLu = "UPDATE egov_dm_data SET "
                          sSQLu = sSQLu & "isActive = "            & lcl_isActive          & ", "
                          sSQLu = sSQLu & "streetnumber = "        & lcl_streetnumber      & ", "
                          sSQLu = sSQLu & "streetprefix = "        & lcl_streetprefix      & ", "
                          sSQLu = sSQLu & "streetaddress = "       & lcl_streetaddress     & ", "
                          sSQLu = sSQLu & "streetsuffix = "        & lcl_streetsuffix      & ", "
                          sSQLu = sSQLu & "streetdirection = "     & lcl_streetdirection   & ", "
                          sSQLu = sSQLu & "sortstreetname = "      & lcl_sortstreetname    & ", "
                          sSQLu = sSQLu & "city = "                & lcl_city              & ", "
                          sSQLu = sSQLu & "state = "               & lcl_state             & ", "
                          sSQLu = sSQLu & "zip = "                 & lcl_zip               & ", "
                          sSQLu = sSQLu & "validstreet = "         & lcl_validstreet       & ", "
                          sSQLu = sSQLu & "latitude = "            & lcl_latitude          & ", "
                          sSQLu = sSQLu & "longitude = "           & lcl_longitude         & ", "
                          sSQLu = sSQLu & "mappointcolor = "       & lcl_mappointcolor     & ", "
                          sSQLu = sSQLu & "isApproved = "          & lcl_isApproved        & ", "
                          sSQLu = sSQLu & "approvedeniedbyid = "   & lcl_approvedeniedbyid & ", "
                          sSQLu = sSQLu & "approvedeniedbydate = " & lcl_approvedeniedbydate
                          sSQLu = sSQLu & " WHERE dmid = " & lcl_dmid

                          set oUpdateDMData = Server.CreateObject("ADODB.Recordset")
                          oUpdateDMData.Open sSQLu, Application("DSN"), 3, 1

                          set oUpdateDMData = nothing

                       end if

                    else
                      'If the MapPointID does NOT have a corresponding DMID then create a DMID.
                       lcl_dmid = ""

                       sSQLi = "INSERT INTO egov_dm_data ("
                       sSQLi = sSQLi & "dm_typeid, "
                       sSQLi = sSQLi & "orgid, "
                       sSQLi = sSQLi & "isCreatedByAdmin, "
                       sSQLi = sSQLi & "createdbyid, "
                       sSQLi = sSQLi & "createdbydate, "
                       sSQLi = sSQLi & "lastmodifiedbyid, "
                       sSQLi = sSQLi & "lastmodifiedbydate, "
                       sSQLi = sSQLi & "isActive, "
                       sSQLi = sSQLi & "streetnumber, "
                       sSQLi = sSQLi & "streetprefix, "
                       sSQLi = sSQLi & "streetaddress, "
                       sSQLi = sSQLi & "streetsuffix, "
                       sSQLi = sSQLi & "streetdirection, "
                       sSQLi = sSQLi & "sortstreetname, "
                       sSQLi = sSQLi & "city, "
                       sSQLi = sSQLi & "state, "
                       sSQLi = sSQLi & "zip, "
                       sSQLi = sSQLi & "validstreet, "
                       sSQLi = sSQLi & "latitude, "
                       sSQLi = sSQLi & "longitude, "
                       sSQLi = sSQLi & "mappointcolor, "
                       sSQLi = sSQLi & "isApproved, "
                       sSQLi = sSQLi & "approvedeniedbyid, "
                       sSQLi = sSQLi & "approvedeniedbydate, "
                       sSQLi = sSQLi & "mappointid "
                       sSQLi = sSQLi & ") VALUES ("
                       sSQLi = sSQLi & lcl_dm_typeid           & ", "
                       sSQLi = sSQLi & lcl_orgid               & ", "
                       sSQLi = sSQLi & lcl_isCreatedByAdmin    & ", "
                       sSQLi = sSQLi & lcl_createdbyid         & ", "
                       sSQLi = sSQLi & lcl_createdbydate       & ", "
                       sSQLi = sSQLi & lcl_lastmodifiedbyid    & ", "
                       sSQLi = sSQLi & lcl_lastmodifiedbydate  & ", "
                       sSQLi = sSQLi & lcl_isActive            & ", "
                       sSQLi = sSQLi & lcl_streetnumber        & ", "
                       sSQLi = sSQLi & lcl_streetprefix        & ", "
                       sSQLi = sSQLi & lcl_streetaddress       & ", "
                       sSQLi = sSQLi & lcl_streetsuffix        & ", "
                       sSQLi = sSQLi & lcl_streetdirection     & ", "
                       sSQLi = sSQLi & lcl_sortstreetname      & ", "
                       sSQLi = sSQLi & lcl_city                & ", "
                       sSQLi = sSQLi & lcl_state               & ", "
                       sSQLi = sSQLi & lcl_zip                 & ", "
                       sSQLi = sSQLi & lcl_validstreet         & ", "
                       sSQLi = sSQLi & lcl_latitude            & ", "
                       sSQLi = sSQLi & lcl_longitude           & ", "
                       sSQLi = sSQLi & lcl_mappointcolor       & ", "
                       sSQLi = sSQLi & lcl_isApproved          & ", "
                       sSQLi = sSQLi & lcl_approvedeniedbyid   & ", "
                       sSQLi = sSQLi & lcl_approvedeniedbydate & ", "
                       sSQLi = sSQLi & lcl_mappointid
                       sSQLi = sSQLi & ")"

                       lcl_dmid              = RunInsertStatement(sSQLi)
                       lcl_totalDMIDsCreated = lcl_totalDMIDsCreated + 1

                      'Update egov_mappoints. Set the new dmid for the current mappointid
                       sSQLu = "UPDATE egov_mappoints SET "
                       sSQLu = sSQLu & " dmid = " & lcl_dmid
                       sSQLu = sSQLu & " WHERE mappointid = " & lcl_mappointid

                       set oUpdateMP = Server.CreateObject("ADODB.Recordset")
                       oUpdateMP.Open sSQLu, Application("DSN"), 3, 1

                       set oUpdateMP = nothing
                    end if

                    lcl_previous_mappointid = oGetTransferMPInfo("mappointid")

                    oGetTransferMPInfo.movenext
                 loop
              end if

              oGetTransferMPInfo.close
              set oGetTransferMPInfo = nothing

             'Set up the return data
              lcl_return_msg = "- Total MapPoints to import: <span class=""redText"">[" & lcl_rowCount & "]</span>"

              if lcl_totalDMIDsCreated > 0 then
                 lcl_return_msg = lcl_return_msg & " - Total DataMgr IDs created: <span class=""redText"">[" & lcl_totalDMIDsCreated & "]</span>"
              end if

              if lcl_totalDMIDsExisting > 0 then
                 lcl_return_msg = lcl_return_msg & " - Total DataMgr IDs already existing: <span class=""redText"">[" & lcl_totalDMIDsExisting & "]</span>"
              end if

              response.write lcl_return_msg

           end if

        elseif lcl_action = "IMPORT_MP_VALUES" then

           set oGetTransferMPValues = Server.CreateObject("ADODB.Recordset")
           oGetTransferMPValues.Open sSQL, Application("DSN"), 3, 1

           if not oGetTransferMPValues.eof then

              lcl_MPValueCount            = 0
              lcl_totalDMValueIDsCreated  = 0
              lcl_totalDMValueIDsExisting = 0

              do while not oGetTransferMPValues.eof
                 lcl_MPValueCount  = lcl_MPValueCount + 1
                 lcl_dmid          = "NULL"
                 lcl_mp_valueid    = "NULL"
                 lcl_fieldvalue    = "NULL"
                 lcl_fieldname     = ""
                 lcl_dm_valueid    = ""
                 lcl_results_label = oGetTransferMPValues("fieldname")

                 if oGetTransferMPValues("dmid") <> "" then
                    lcl_dmid = oGetTransferMPValues("dmid")
                 end if

                 if oGetTransferMPValues("mp_valueid") <> "" then
                    lcl_mp_valueid = oGetTransferMPValues("mp_valueid")
                    lcl_dm_valueid = getDMValueID_by_MPValueID(lcl_mp_valueid)
                 end if

                 if oGetTransferMPValues("fieldname") <> "" then
                    lcl_fieldname = ucase(oGetTransferMPValues("fieldname"))
                 end if

                 if oGetTransferMPValues("fieldvalue") <> "" then
                    lcl_fieldvalue = oGetTransferMPValues("fieldvalue")

                    if instr(lcl_fieldname,"WEBSITE") > 0 OR instr(lcl_fieldname,"EMAIL") > 0 then
                       lcl_fieldvalue = "[" & lcl_fieldvalue & "]<>"
                    end if

                    lcl_fieldvalue = dbsafe(lcl_fieldvalue)
                    lcl_fieldvalue = "'" & lcl_fieldvalue & "'"
                 end if

                'Check to see if a dm_valueid exists for the mp_valueid.  
                'If "yes" then override the fieldvalue.  If "no" then create a new mp_valueid.
                 if lcl_dm_valueid <> "" then
                    if lcl_overrideValues = "Y" then
                       sSQLu = "UPDATE egov_dm_values SET "
                       sSQLu = sSQLu & " fieldvalue = " & lcl_fieldvalue
                       sSQLu = sSQLu & " WHERE dm_valueid = " & lcl_dm_valueid

                       set oUpdateMPValues = Server.CreateObject("ADODB.Recordset")
                       oUpdateMPValues.Open sSQLu, Application("DSN"), 3, 1

                       set oUpdateMPValues = nothing
                    end if

                    lcl_totalDMValueIDsExisting = lcl_totalDMValueIDsExisting + 1

                 else
                    sSQLi = "INSERT INTO egov_dm_values ("
                    sSQLi = sSQLi & "orgid, "
                    sSQLi = sSQLi & "dm_typeid, "
                    sSQLi = sSQLi & "dmid, "
                    sSQLi = sSQLi & "dm_sectionid, "
                    sSQLi = sSQLi & "dm_fieldid, "
                    sSQLi = sSQLi & "fieldvalue, "
                    sSQLi = sSQLi & "mp_valueid "
                    sSQLi = sSQLi & ") VALUES ("
                    sSQLi = sSQLi & lcl_orgid        & ", "
                    sSQLi = sSQLi & lcl_dm_typeid    & ", "
                    sSQLi = sSQLi & lcl_dmid         & ", "
                    sSQLi = sSQLi & lcl_dm_sectionid & ", "
                    sSQLi = sSQLi & lcl_dm_fieldid   & ", "
                    sSQLi = sSQLi & lcl_fieldvalue   & ", "
                    sSQLi = sSQLi & lcl_mp_valueid
                    sSQLi = sSQLi & ") "

                    lcl_dm_valueid             = RunInsertStatement(sSQLi)
                    lcl_totalDMValueIDsCreated = lcl_totalDMValueIDsCreated + 1

                   'Update egov_mappoints_values. Set the new dm_valueid for the current mp_valueid
                    sSQLu = "UPDATE egov_mappoints_values SET "
                    sSQLu = sSQLu & " dm_valueid = " & lcl_dm_valueid
                    sSQLu = sSQLu & " WHERE mp_valueid = " & lcl_mp_valueid

                    set oUpdateMPValues = Server.CreateObject("ADODB.Recordset")
                    oUpdateMPValues.Open sSQLu, Application("DSN"), 3, 1

                    set oUpdateMPValues = nothing

                 end if

                 oGetTransferMPValues.movenext
              loop

             'Set up the return data
              lcl_return_msg = lcl_return_msg & "<p>"
              lcl_return_msg = lcl_return_msg & "<table border=""0"">"
              lcl_return_msg = lcl_return_msg &   "<tr>"
              lcl_return_msg = lcl_return_msg &       "<td>- </td>"
              lcl_return_msg = lcl_return_msg &       "<td>Total MapPoint Values for ""<span class=""redText"">" & lcl_results_label & "</span>"" to import: <span class=""redText"">[" & lcl_MPValueCount & "]</span></td>"
              lcl_return_msg = lcl_return_msg &   "</tr>"

              if lcl_totalDMValueIDsCreated > 0 then
                 lcl_return_msg = lcl_return_msg &   "<tr>"
                 lcl_return_msg = lcl_return_msg &       "<td>&nbsp;</td>"
                 lcl_return_msg = lcl_return_msg &       "<td>Total DataMgr values created: <span class=""redText"">[" & lcl_totalDMValueIDsCreated & "]</span></td>"
                 lcl_return_msg = lcl_return_msg &   "</tr>"
              end if

              if lcl_totalDMValueIDsExisting > 0 then
                 lcl_return_msg = lcl_return_msg &   "<tr>"
                 lcl_return_msg = lcl_return_msg &       "<td>&nbsp;</td>"
                 lcl_return_msg = lcl_return_msg &       "<td>Total DataMgr values already existing"

                 if lcl_overrideValues = "Y" then
                    lcl_return_msg = lcl_return_msg & " (existing data overridden with import value)"
                 end if

                 lcl_return_msg = lcl_return_msg & ": <span class=""redText"">[" & lcl_totalDMValueIDsExisting & "]</span></td>"
                 lcl_return_msg = lcl_return_msg &   "</tr>"
              end if

              lcl_return_msg = lcl_return_msg & "</table>"
              lcl_return_msg = lcl_return_msg & "</p>"

              response.write lcl_return_msg

           end if

           oGetTransferMPValues.close
           set oGetTransferMPValues = nothing

        end if
       'END: Determine which action to take in the import process -------------

     end if
  end if
%>