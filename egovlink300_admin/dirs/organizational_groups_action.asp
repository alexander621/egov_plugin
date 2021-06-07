<!-- #include file="../includes/common.asp" //-->
<!-- #include file="organizational_groups_global_functions.asp" //-->
<%
  dim sUserAction

  sUserAction         = ""
  lcl_success         = ""
  lcl_screen_mode_url = ""
  lcl_org_group_id    = ""
  lcl_orgid           = ""

  if request("user_action") <> "" then
     if not containsApostrophe(request("user_action")) then
        sUserAction = ucase(request("user_action"))
     end if
  end if

  if request("org_group_id") <> "" then
     lcl_org_group_id = clng(request("org_group_id"))
  end if

  if request("orgid") <> "" then
     lcl_orgid = clng(request("orgid"))
  end if

  if sUserAction <> "" then
     lcl_org_name            = "NULL"
     lcl_parent_org_group_id = ""
     lcl_org_level           = ""
     lcl_address             = "NULL"
     lcl_address2            = "NULL"
     lcl_city                = "NULL"
     lcl_state               = "NULL"
     lcl_zip                 = "NULL"
     lcl_phone_number        = "NULL"
     lcl_phone_number_ext    = "NULL"
     lcl_fax_number          = "NULL"
     lcl_email               = "NULL"
     lcl_active_flag         = "'N'"

     if request("org_name") <> "" then
        lcl_org_name = dbsafe(request("org_name"))
        lcl_org_name = "'" & lcl_org_name & "'"
     end if

     if request("parent_org_group_id") <> "" then
        lcl_parent_org_group_id = clng(request("parent_org_group_id"))
     end if

     if request("org_level") <> "" then
        lcl_org_level = clng(request("org_level"))
     end if

     if request("address") <> "" then
        lcl_address = dbsafe(request("address"))
        lcl_address = "'" & lcl_address & "'"
     end if

     if request("address2") <> "" then
        lcl_address2 = dbsafe(request("address2"))
        lcl_address2 = "'" & lcl_address2 & "'"
     end if

     if request("city") <> "" then
        lcl_city = dbsafe(request("city"))
        lcl_city = "'" & lcl_city & "'"
     end if

     if request("state") <> "" then
        lcl_state = dbsafe(request("state"))
        lcl_state = "'" & lcl_state & "'"
     end if

     if request("zip") <> "" then
        lcl_zip = dbsafe(request("zip"))
        lcl_zip = "'" & lcl_zip & "'"
     end if

     if request("phone_number") <> "" then
        lcl_phone_number = dbsafe(request("phone_number"))
        lcl_phone_number = "'" & lcl_phone_number & "'"
     end if

     if request("phone_number_ext") <> "" then
        lcl_phone_number_ext = dbsafe(request("phone_number_ext"))
        lcl_phone_number_ext = "'" & lcl_phone_number_ext & "'"
     end if

     if request("fax_number") <> "" then
        lcl_fax_number = dbsafe(request("fax_number"))
        lcl_fax_number = "'" & lcl_fax_number & "'"
     end if

     if request("email") <> "" then
        lcl_email = dbsafe(request("email"))
        lcl_email = "'" & lcl_email & "'"
     end if

     if request("active_flag") <> "" then
        lcl_active_flag = dbsafe(request("active_flag"))
        lcl_active_flag = "'" & lcl_active_flag & "'"
     end if

    'To set up the ORG_LEVEL we first have to get the ORG_LEVEL of the parent, if available, and
    'then add (1) to it.  If the parent_org_group_id is NULL then set the ORG_LEVEL to (1).
     lcl_org_level = getOrgLevelByParentGroupID(lcl_parent_org_group_id)

     if lcl_parent_org_group_id = "" then
        lcl_parent_org_group_id = 0
     end if

     if lcl_orgid = "" then
        lcl_orgid = session("orgid")
     end if
  end if

 '-----------------------------------------------------------------------------
  if sUserAction = "U" then
     sSQLu = "UPDATE egov_staff_directory_groups SET "
     sSQLu = sSQLu & " org_name = "            & lcl_org_name            & ", "
     sSQLu = sSQLu & " parent_org_group_id = " & lcl_parent_org_group_id & ", "
     sSQLu = sSQLu & " org_level = "           & lcl_org_level           & ", "
     sSQLu = sSQLu & " orgid = "               & lcl_orgid               & ", "
     sSQLu = sSQLu & " address = "             & lcl_address             & ", "
     sSQLu = sSQLu & " address2 = "            & lcl_address2            & ", "
     sSQLu = sSQLu & " city = "                & lcl_city                & ", "
     sSQLu = sSQLu & " state = "               & lcl_state               & ", "
     sSQLu = sSQLu & " zip = "                 & lcl_zip                 & ", "
     sSQLu = sSQLu & " phone_number = "        & lcl_phone_number        & ", "
     sSQLu = sSQLu & " phone_number_ext = "    & lcl_phone_number_ext    & ", "
     sSQLu = sSQLu & " fax_number = "          & lcl_fax_number          & ", "
     sSQLu = sSQLu & " email = "               & lcl_email               & ", "
     sSQLu = sSQLu & " active_flag = "         & lcl_active_flag
     sSQLu = sSQLu & " WHERE org_group_id = "  & lcl_org_group_id

     set oUpdateStaffDirectoryGroup = Server.CreateObject("ADODB.Recordset")
     oUpdateStaffDirectoryGroup.Open sSQLu, Application("DSN"), 3, 1

     set oUpdateStaffDirectoryGroup = nothing

    'Now check to see if this org_group_id is a parent_org_group_id.
    'If so then update the org_level for all of its sub-org groups as well.
     if check_for_sub_org_groups(lcl_org_group_id) = "Y" then
        update_sub_org_levels lcl_org_group_id, _
                              lcl_org_level
     end if

     lcl_success = "SU"

'------------------------------------------------------------------------------
  elseif sUserAction = "I" OR sUserAction = "AA" then

   'Set up the columns to be inserted and their related values.
    lcl_insert_columns = "org_name, parent_org_group_id, org_level, orgid"
    lcl_insert_values  = lcl_org_name & ", " & lcl_parent_org_group_id & ", " & lcl_org_level & ", " & lcl_orgid

   '-- Address -----------------------------------------------------------------
    if lcl_address <> "" then
       lcl_insert_columns = lcl_insert_columns & ", address"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_address
    end if
   '-- Address 2 ---------------------------------------------------------------
    if lcl_address2 <> "" then
       lcl_insert_columns = lcl_insert_columns & ", address2"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_address2
    end if
   '-- City --------------------------------------------------------------------
    if lcl_city <> "" then
       lcl_insert_columns = lcl_insert_columns & ", city"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_city
    end if
   '-- State -------------------------------------------------------------------
    if lcl_state <> "" then
       lcl_insert_columns = lcl_insert_columns & ", state"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_state
    end if
   '-- Zip ---------------------------------------------------------------------
    if lcl_zip <> "" then
       lcl_insert_columns = lcl_insert_columns & ", zip"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_zip
    end if
   '-- Phone Number ------------------------------------------------------------
    if lcl_phone_number <> "" then
       lcl_insert_columns = lcl_insert_columns & ", phone_number"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_phone_number
    end if
   '-- Phone Number Ext --------------------------------------------------------
    if lcl_phone_number_ext <> "" then
       lcl_insert_columns = lcl_insert_columns & ", phone_number_ext"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_phone_number_ext
    end if
   '-- Fax Number --------------------------------------------------------------
    if lcl_fax_number <> "" then
       lcl_insert_columns = lcl_insert_columns & ", fax_number"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_fax_number
    end if
   '-- Email -------------------------------------------------------------------
    if lcl_email <> "" then
       lcl_insert_columns = lcl_insert_columns & ", email"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_email
    end if
   '-- Active Flag -------------------------------------------------------------
    if lcl_active_flag <> "" then
       lcl_insert_columns = lcl_insert_columns & ", active_flag"
       lcl_insert_values  = lcl_insert_values  & ", " & lcl_active_flag
    end if
   '----------------------------------------------------------------------------

    sSQLi = "INSERT INTO egov_staff_directory_groups (" & lcl_insert_columns & ") VALUES (" & lcl_insert_values & ")"

    lcl_org_group_id = RunInsertStatement(sSQLi)
    lcl_success      = "SI"

   'If the user has selected to "Add Another" then handle this as an insert
   'But return to the add/edit screen in "INSERT" mode
    if sUserAction = "AA" then
       lcl_org_group_id    = ""
       lcl_screen_mode_url = "&screen_mode=ADD"
    end if

  elseif sUserAction = "D" then

    'Check to see if this org has any sub_org_groups.
     if check_for_sub_org_groups(lcl_org_group_id) = "Y" then
        sSQL1 = "SELECT org_group_id "
        sSQL1 = sSQL1 & " FROM egov_staff_directory_groups "
        sSQL1 = sSQL1 & " WHERE orgid = " & lcl_orgid
        sSQL1 = sSQL1 & " AND parent_org_group_id = " & lcl_org_group_id

        set oCheckForSubGroups = Server.CreateObject("ADODB.Recordset")
        oCheckForSubGroups.Open sSQL1, Application("DSN"), 3, 1

        lcl_delete_org_group_ids = lcl_org_group_id

       'If sub_org_groups exist then build a comma seperated list that is to be used in the delete.
        if not oCheckForSubGroups.eof then
           do while not oCheckForSubGroups.eof
             lcl_delete_org_group_ids = lcl_delete_org_group_ids & ", " & oCheckForSubGroups("org_group_id")

            'Then check each sub_org_group for any sub_org_groups it may have and add those to the delete list
             if check_for_sub_org_groups(oCheckForSubGroups("org_group_id")) = "Y" then
                lcl_sub_orgs = get_sub_org_groups(oCheckForSubGroups("org_group_id"))

                if lcl_sub_orgs <> "" then
                   lcl_delete_org_group_ids = lcl_delete_org_group_ids & ", " & lcl_sub_orgs
                end if
             end if

             oCheckForSubGroups.movenext
           loop
        end if

        oCheckForSubGroups.close
        set oCheckForSubGroups = nothing
     else
        lcl_delete_org_group_ids = lcl_org_group_id
     end if

     sSQLd = "DELETE FROM egov_staff_directory_groups WHERE org_group_id IN (" & lcl_delete_org_group_ids & ")"
     set oDeleteOrgGroupID = Server.CreateObject("ADODB.Recordset")
     oDeleteOrgGroupID.Open sSQLd, Application("DSN"), 3, 1

     set oDeleteOrgGroupID = nothing

     response.redirect "organizational_groups_list.asp?success=D"

  end if

  if sUserAction <> "D" then
     response.redirect "organizational_groups_maint.asp?org_group_id=" & lcl_org_group_id & "&success=" & lcl_success & lcl_screen_mode_url
  end if

'------------------------------------------------------------------------------
function getOrgLevelByParentGroupID(iParentOrgGroupID)
  dim sParentOrgGroupID, sSQL, lcl_return

  sParentOrgGroupID = ""
  sSQL              = ""
  lcl_return        = 1

  if iParentOrgGroupID <> "" then
     sParentOrgGroupID = clng(iParentOrgGroupID)
  end if

  if sParentOrgGroupID <> "" then
     sSQL = "SELECT org_level "
     sSQL = sSQL & " FROM egov_staff_directory_groups "
     sSQL = sSQL & " WHERE org_group_id = " & sParentOrgGroupID

     set oGetOrgLevel = Server.CreateObject("ADODB.Recordset")
     oGetOrgLevel.Open sSQL, Application("DSN"), 3, 1

     if not oGetOrgLevel.eof then
        lcl_return = oGetOrgLevel("org_level") + 1
     end if

  			oGetOrgLevel.close
    	set oGetOrgLevel = nothing
  else
     lcl_return = 1
  end if

  getOrgLevelByParentGroupID = lcl_return

end function

'-----------------------------------------------------------------------------
'This function retrieves a list of sub_org_groups for the parent_org_group_id 
'specified that sets up the delete_org function
'-----------------------------------------------------------------------------
function get_sub_org_groups(p_org_group_id)
  if check_for_sub_org_groups(p_org_group_id) = "Y" then
     sSQL = "SELECT org_group_id FROM egov_staff_directory_groups "
     sSQL = sSQL & " WHERE parent_org_group_id = " & p_org_group_id

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSQL, Application("DSN"), 3, 1

     if not rs.eof then
        while not rs.eof
          if lcl_delete_org_group_ids <> "" then
             lcl_delete_org_group_ids = lcl_delete_org_group_ids & ", " & rs("org_group_id")
          else
             lcl_delete_org_group_ids = rs("org_group_id")
          end if

          if check_for_sub_org_groups(rs("org_group_id")) = "Y" then
             lcl_sub_orgs = get_sub_org_groups(rs("org_group_id"))
          end if

          if lcl_sub_orgs <> "" then
             lcl_delete_org_group_ids = lcl_delete_org_group_ids & ", " & lcl_sub_orgs
          end if             

          rs.movenext
        wend
     else
        lcl_delete_org_group_ids = ""
     end if

     rs.close
     set rs = nothing

  else
     lcl_delete_org_group_ids = ""
  end if

  get_sub_org_groups = lcl_delete_org_group_ids

end function
%>