<!-- #include file="../includes/common.asp" //-->
<%
  dim lcl_action, lcl_success, lcl_orgid, lcl_total_substatuses, lcl_action_linenumber
  dim lcl_new_statusid, sNewSubStatus, sNewParentStatus, lcl_max_displayorder
  dim lcl_delete_statusid
  dim lcl_move_statusid, lcl_move_parentstatus

  lcl_action            = ""
  lcl_success           = ""
  lcl_orgid             = 0
  lcl_action_linenumber = 0

  if request("action") <> "" then
     if not containsApostrophe(request("action")) then
        lcl_action = ucase(request("action"))
     end if
  end if

  if request("orgid") <> "" then
     if not containsApostrophe(request("orgid")) then
        lcl_orgid = clng(request("orgid"))
     end if
  end if

  if request("action_linenumber") <> "" then
     if not containsApostrophe(request("action_linenumber")) then
        lcl_action_linenumber = clng(request("action_linenumber"))
     end if
  end if

  if lcl_action = "ADD" then
     lcl_max_displayorder = 1

     if request("newSubStatus") <> "" then
        sNewSubStatus = request("newSubStatus")
        sNewSubStatus = dbsafe(sNewSubStatus)
        sNewSubStatus = "'" & sNewSubStatus & "'"
     end if

     if request("newParentStatus") <> "" then
        sNewParentStatus = request("newParentStatus")

        lcl_max_displayorder = getMaxDisplayOrder(lcl_orgid, sNewParentStatus)

        sNewParentStatus = dbsafe(sNewParentStatus)
        sNewParentStatus = "'" & sNewParentStatus & "'"
     end if

     sSQLi = "INSERT INTO egov_actionline_requests_statuses ("
     sSQLi = sSQLi & "status_name, "
     sSQLi = sSQLi & "orgid, "
     sSQLi = sSQLi & "parent_status, "
     sSQLi = sSQLi & "display_order, "
     sSQLi = sSQLi & "active_flag"
     sSQLi = sSQLi & ") VALUES ("
     sSQLi = sSQLi & sNewSubStatus        & ", "
     sSQLi = sSQLi & lcl_orgid            & ", "
     sSQLi = sSQLi & sNewParentStatus     & ", "
     sSQLi = sSQLi & lcl_max_displayorder & ", "
     sSQLi = sSQLi & "'Y'"
     sSQLi = sSQLi & ") "

     lcl_new_statusid = RunInsertStatement(sSQLi)
     lcl_success      = "SA"

  elseif lcl_action = "EDIT" then

     lcl_total_substatuses = 0
     lcl_edit_statusid     = 0
     lcl_edit_substatus    = ""
     lcl_edit_active       = ""
     lcl_edit_parentstatus = ""

     if request("total_substatuses") <> "" then
        if not containsApostrophe("total_substatuses") then
           lcl_total_substatuses = clng(request("total_substatuses"))
        end if
     end if

     for e = 1 to lcl_total_substatuses
        if request("editActionStatusID_" & e) <> "" then
           if not containsApostrophe(request("editActionStatusID_" & e)) then
              lcl_edit_statusid = clng(request("editActionStatusID_" & e))
           end if
        end if

        if request("editSubStatus_" & e) <> "" then
           lcl_edit_substatus = request("editSubStatus_" & e)
           lcl_edit_substatus = dbsafe(lcl_edit_substatus)
           lcl_edit_substatus = "'" & lcl_edit_substatus & "'"
        end if

        if request("editActive_" & e) <> "" then
           lcl_edit_active = request("editActive_" & e)
           lcl_edit_active = dbsafe(lcl_edit_active)
           lcl_edit_active = "'" & lcl_edit_active & "'"
        end if

        if request("editParentStatus_" & e) <> "" then
           lcl_edit_parentstatus = request("editParentStatus_" & e)
           lcl_edit_parentstatus = dbsafe(lcl_edit_parentstatus)
           lcl_edit_parentstatus = "'" & lcl_edit_parentstatus & "'"
        end if

        sSQLu = "UPDATE egov_actionline_requests_statuses SET "
        sSQLu = sSQLu & "status_name = "   & lcl_edit_substatus & ", "
        sSQLu = sSQLu & "parent_status = " & lcl_edit_parentstatus & ", "
        sSQLu = sSQLu & "active_flag = "   & lcl_edit_active
        sSQLu = sSQLu & " WHERE action_status_id = " & lcl_edit_statusid

        set oUpdateSubStatuses = Server.CreateObject("ADODB.Recordset")
        oUpdateSubStatuses.Open sSQLu, Application("DSN"), 3, 1

        set oUpdateSubStatuses = nothing
     next

     lcl_success = "SU"

  elseif lcl_action = "DELETE" then
     lcl_delete_statusid = 0

     if lcl_action_linenumber > 0 then
        if request("editActionStatusID_" & lcl_action_linenumber) <> "" then
           if not containsApostrophe(request("editActionStatusID_" & lcl_action_linenumber)) then
              lcl_delete_statusid = clng(request("editActionStatusID_" & lcl_action_linenumber))
           end if
        end if
     end if

     if lcl_delete_statusid > 0 then
        sSQLd = "DELETE FROM egov_actionline_requests_statuses WHERE action_status_id = " & lcl_delete_statusid

        set oDeleteSubStatus = Server.CreateObject("ADODB.Recordset")
        oDeleteSubStatus.Open sSQLd, Application("DSN"), 3, 1

        set oDeleteSubStatus = nothing

        lcl_success = "SD"
     end if

  elseif lcl_action = "MOVEUP" OR lcl_action = "MOVEDOWN" then
     lcl_move_statusid        = 0
     lcl_move_parentstatus    = ""
     lcl_current_status_order = 0

     if lcl_action_linenumber > 0 then
        if request("editActionStatusID_" & lcl_action_linenumber) <> "" then
           if not containsApostrophe(request("editActionStatusID_" & lcl_action_linenumber)) then
              lcl_move_statusid = clng(request("editActionStatusID_" & lcl_action_linenumber))
           end if
        end if

        if request("editParentStatus_" & lcl_action_linenumber) <> "" then
           lcl_move_parentstatus = request("editParentStatus_" & lcl_action_linenumber)
           lcl_move_parentstatus = dbsafe(lcl_move_parentstatus)
           lcl_move_parentstatus = "'" & lcl_move_parentstatus & "'"
        end if

     end if

    'Get the display_order of the "current" action_status_id
     sSQLp = "SELECT display_order AS current_status_order "
     sSQLp = sSQLp & " FROM egov_actionline_requests_statuses "
     sSQLp = sSQLp & " WHERE action_status_id = " & lcl_move_statusid
     sSQLp = sSQLp & " AND active_flag = 'Y' "

     set oCSO = Server.CreateObject("ADODB.Recordset")
     oCSO.Open sSQLp, Application("DSN"), 3, 1

     lcl_current_status_order = oCSO("current_status_order")

    'Get the action_status_id of the next/previous sub-status
     lcl_select_maxmin      = "min"
     lcl_where_displayorder = " > "

     if lcl_action = "MOVEUP" then
        lcl_select_maxmin      = "max"
        lcl_where_displayorder = " < "
     end if

     sSQLp2 = "SELECT ISNULL(" & lcl_select_maxmin & "(display_order)," & lcl_current_status_order & ") AS np_status_order "
     sSQLp2 = sSQLp2 & " FROM egov_actionline_requests_statuses "
     sSQLp2 = sSQLp2 & " WHERE active_flag = 'Y' "
     sSQLp2 = sSQLp2 & " AND orgid = "         & lcl_orgid
     sSQLp2 = sSQLp2 & " AND parent_status = " & lcl_move_parentstatus
     sSQLp2 = sSQLp2 & " AND display_order" & lcl_where_displayorder & lcl_current_status_order

     set oCSO2 = Server.CreateObject("ADODB.Recordset")
     oCSO2.Open sSQLp2, Application("DSN"), 3, 1

     if not oCSO2.eof then
        lcl_np_status_order = oCSO2("np_status_order")

       'Set the display_order to the "current" status display_order for the new sub-status
        sSQLp3 = "UPDATE egov_actionline_requests_statuses SET "
        sSQLp3 = sSQLp3 & " display_order = " & lcl_current_status_order
      	 sSQLp3 = sSQLp3 & " WHERE active_flag = 'Y' "
      	 sSQLp3 = sSQLp3 & " AND orgid = "         & lcl_orgid
      	 sSQLp3 = sSQLp3 & " AND parent_status = " & lcl_move_parentstatus
      	 sSQLp3 = sSQLp3 & " AND display_order = " & lcl_np_status_order

      	 set oCSO3 = Server.CreateObject("ADODB.Recordset")
        oCSO3.Open sSQLp3, Application("DSN"), 3, 1

       'Set the display_order to the "next/previous" status display_order for the current sub-status
        sSQLp4 = "UPDATE egov_actionline_requests_statuses SET "
        sSQLp4 = sSQLp4 & " display_order = " & lcl_np_status_order
        sSQLp4 = sSQLp4 & " WHERE action_status_id = " & lcl_move_statusid

        set oCSO4 = Server.CreateObject("ADODB.Recordset")
        oCSO4.Open sSQLp4, Application("DSN"), 3, 1

        set oCSO3 = nothing
        set oCSO4 = nothing

     end if

     set oCSO  = nothing
     set oCSO2 = nothing

     lcl_success = "SU"
  end if

  response.redirect "action_substatus.asp?success=" & lcl_success

'------------------------------------------------------------------------------
function getMaxDisplayOrder(iOrgID, iParentStatus)

  dim sOrgID, sParentStatus, lcl_return

  sOrgID        = 0
  sParentStatus = ""
  lcl_return    = 1

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iParentStatus <> "" then
     sParentStatus = ucase(iParentStatus)
     sParentStatus = dbsafe(sParentStatus)
     sParentStatus = "'" & sParentStatus & "'"
  end if

  sSQLdo = "SELECT isnull(max(display_order),0)+1 AS max_display_order "
  sSQLdo = sSQLdo & " FROM egov_actionline_requests_statuses "
  sSQLdo = sSQLdo & " WHERE orgid = " & sOrgID
  sSQLdo = sSQLdo & " AND upper(parent_status) = " & sParentStatus

 	set oGetMaxDisplayOrder = Server.CreateObject("ADODB.Recordset")
 	oGetMaxDisplayOrder.Open sSQLdo, Application("DSN"), 3, 1

  if not oGetMaxDisplayOrder.eof then
     lcl_return = oGetMaxDisplayOrder("max_display_order")
  end if

  oGetMaxDisplayOrder.close
  set oGetMaxDisplayOrder = nothing

  getMaxDisplayOrder = lcl_return

end function
%>