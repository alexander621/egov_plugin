<!-- #include file="../includes/common.asp" //-->
<%
  dim lcl_action, lcl_orgid, lcl_success, lcl_total_codesections
  dim lcl_newActionCodeID, lcl_newCode, lcl_newDescription, lcl_newActive
  dim lcl_editActionCodeID, lcl_editCode, lcl_editDescription, lcl_editActive
  dim lcl_deleteActionCodeID

  lcl_action             = ""
  lcl_success            = ""
  lcl_orgid              = 0
  lcl_total_codesections = 0

  if request("p_action") <> "" then
     if not containsApostrophe(request("p_action")) then
        lcl_action = ucase(request("p_action"))
     end if
  end if

  if request("orgid") <> "" then
     if not containsApostrophe(request("orgid")) then
        lcl_orgid = clng(request("orgid"))
     end if
  end if

  if lcl_action = "ADD" then
     lcl_newActionCodeID = 0
     lcl_newCode         = ""
     lcl_newDescription  = "NULL"
     lcl_newActive       = "'Y'"

     if request("newCode") <> "" then
        lcl_newCode = request("newCode")
        lcl_newCode = dbsafe(lcl_newCode)
        lcl_newCode = "'" & lcl_newCode & "'"
     end if

     if request("newDescription") <> "" then
        lcl_newDescription = request("newDescription")
        lcl_newDescription = dbsafe(lcl_newDescription)
        lcl_newDescription = "'" & lcl_newDescription & "'"
     end if

     if lcl_orgid > 0 AND lcl_newCode <> "" then
        sSQLi = "INSERT INTO egov_actionline_code_sections ("
        sSQLi = sSQLi & "code_name, "
        sSQLi = sSQLi & "description, "
        sSQLi = sSQLi & "active_flag, "
        sSQLi = sSQLi & "orgid"
        sSQLi = sSQLi & ") VALUES ("
        sSQLi = sSQLi & lcl_newCode        & ", "
        sSQLi = sSQLi & lcl_newDescription & ", "
        sSQLi = sSQLi & lcl_newActive      & ", "
        sSQLi = sSQLi & lcl_orgid
        sSQLi = sSQLi & ") "

        lcl_newActionCodeID = RunInsertStatement(sSQLi)
        lcl_success         = "SA"
     end if

  elseif lcl_action = "EDIT" then
     if request("total_codesections") <> "" then
        if not containsApostrophe(request("total_codesections")) then
           lcl_total_codesections = clng(request("total_codesections"))
        end if
     end if

     if lcl_total_codesections > 0 then
        for e = 1 to lcl_total_codesections
           lcl_editActionCodeID = 0
           lcl_editCode         = ""
           lcl_editDescription  = "NULL"
           lcl_editActive       = "'Y'"

           if request("editActionCodeID_" & e) <> "" then
              if not containsApostrophe(request("editActionCodeID_" & e)) then
                 lcl_editActionCodeID = clng(request("editActionCodeID_" & e))
              end if
           end if

           if request("editCode_" & e) <> "" then
              lcl_editCode = request("editCode_" & e)
              lcl_editCode = dbsafe(lcl_editCode)
              lcl_editCode = "'" & lcl_editCode & "'"
           end if

           if request("editDescription_" & e) <> "" then
              lcl_editDescription = request("editDescription_" & e)
              lcl_editDescription = dbsafe(lcl_editDescription)
              lcl_editDescription = "'" & lcl_editDescription & "'"
           end if

           if request("editActive_" & e) <> "" then
              lcl_editActive = request("editActive_" & e)
              lcl_editActive = ucase(lcl_editActive)
              lcl_editActive = dbsafe(lcl_editActive)
              lcl_editActive = "'" & lcl_editActive & "'"
           end if

           if lcl_editCode <> "" then
              sSQLu = "UPDATE egov_actionline_code_sections SET "
              sSQLu = sSQLu & " code_name = "   & lcl_editCode        & ", "
              sSQLu = sSQLu & " description = " & lcl_editDescription & ", "
              sSQLu = sSQLu & " active_flag = " & lcl_editActive
              sSQLu = sSQLu & " WHERE action_code_id = " & lcl_editActionCodeID

          		  set oUpdateCS = Server.CreateObject("ADODB.Recordset")
              oUpdateCS.Open sSQLu, Application("DSN"), 3, 1
           end if
        next

        lcl_success = "SU"

     end if
  elseif lcl_action = "DELETE" then
     lcl_deleteActionCodeID = 0

     if request("deleteActionCodeID") <> "" then
        if not containsApostrophe(request("deleteActionCodeID")) then
           lcl_deleteActionCodeID = clng(request("deleteActionCodeID"))
        end if
     end if

     if lcl_deleteActionCodeID > 0 then
        sSQLd = "DELETE FROM egov_actionline_code_sections "
        sSQLd = sSQLd & " WHERE action_code_id = " & lcl_deleteActionCodeID

        set oDeleteCS = Server.CreateObject("ADODB.Recordset")
        oDeleteCS.Open sSQLd, Application("DSN"), 3, 1

        set oDeleteCS = nothing

        lcl_success = "SD"
     end if

  end if

  response.redirect "action_code_sections.asp?success=" & lcl_success
%>