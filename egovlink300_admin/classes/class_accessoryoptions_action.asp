<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
  lcl_orgid         = 0
  lcl_classid       = 0
  lcl_atype         = ""
  lcl_total_options = 0

  if request("orgid") <> "" then
     lcl_orgid = clng(request("orgid"))
  end if

  if request("classid") <> "" then
     lcl_classid = clng(request("classid"))
  end if

  if request("atype") <> "" then
     if not containsApostrophe(request("atype")) then
        lcl_atype = request("atype")
        lcl_atype = ucase(lcl_atype)
     end if
  end if

  if request("total_options") <> "" then
     lcl_total_options = clng(request("total_options"))
  end if

  if  lcl_orgid         > 0 _
  AND lcl_classid       > 0 _
  AND lcl_total_options > 0 _
  AND lcl_atype <> "" then

    'Setup the "accessorytype" for database use
     lcl_accessorytype = lcl_atype
     lcl_accessorytype = dbsafe(lcl_accessorytype)
     lcl_accessorytype = "'" & lcl_accessorytype & "'"

    'Remove all of the assignments for the org and class/event
     sSQL = "DELETE FROM egov_class_teamroster_accessories_to_class "
     sSQL = sSQL & " WHERE orgid = " & lcl_orgid
     sSQL = sSQL & " AND classid = " & lcl_classid
     sSQL = sSQL & " AND accessoryid IN (select accessoryid "
     sSQL = sSQL &                     " from egov_class_teamroster_accessories "
     sSQL = sSQL &                     " where UPPER(accessorytype) = " & lcl_accessorytype
     sSQL = sSQL &                     " and orgid = " & lcl_orgid
     sSQL = sSQL &                     ")"

    	set oDelAccessories = Server.CreateObject("ADODB.Recordset")
    	oDelAccessories.Open sSQL, Application("DSN"), 3, 1

     set oDelAccessories = nothing

    'Assign all values to org for the class/event
     lcl_total_options = 0

     if request("total_options") <> "" then
        lcl_total_options = request("total_options")
     end if

     if lcl_total_options > 0 then
        for i = 1 to lcl_total_options

           sSQLu = ""
           sSQLi = ""
           sSQLa = ""

          'Determine which action to take: insert/update/delete
           if request.form("delete_" & i) <> "" then
              lcl_delete_accessoryid = request.form("delete_" & i)

              sSQLd = "DELETE FROM egov_class_teamroster_accessories "
              sSQLd = sSQLd & " WHERE accessoryid = " & lcl_delete_accessoryid

             	set oDeleteAccessory = Server.CreateObject("ADODB.Recordset")
             	oDeleteAccessory.Open sSQLd, Application("DSN"), 3, 1

              set oDeleteAccessory = nothing
           else
              lcl_assign_accessoryid = 0
              lcl_accessoryid        = 0
              lcl_accessoryname      = ""
              lcl_accessoryvalue     = ""
              lcl_displayorder       = 1

              if trim(request.form("assign_accessoryid_" & i)) <> "" then
                 lcl_assign_accessoryid = request.form("assign_accessoryid_" & i)
                 lcl_assign_accessoryid = clng(lcl_assign_accessoryid)
              end if

              if trim(request.form("accessoryid_" & i)) <> "" then
                 lcl_accessoryid = request.form("accessoryid_" & i)
                 lcl_accessoryid = clng(lcl_accessoryid)
              end if

              if trim(request.form("accessoryname_" & i)) <> "" then
                 lcl_accessoryname = request.form("accessoryname_"  & i)
                 lcl_accessoryname = trim(lcl_accessoryname)
              end if

              if trim(request.form("accessoryvalue_" & i)) <> "" then
                 lcl_accessoryvalue = request.form("accessoryvalue_"  & i)
                 lcl_accessoryvalue = trim(lcl_accessoryvalue)
              end if

              if trim(request.form("displayorder_" & i)) <> "" then
                 lcl_displayorder = request.form("displayorder_"  & i)
                 lcl_displayorder = clng(lcl_displayorder)
              end if

              if lcl_accessoryname <> "" AND lcl_accessoryvalue <> "" AND lcl_displayorder <> "" then
                 if lcl_accessoryname <> "" then
                    lcl_accessoryname = dbsafe(lcl_accessoryname)
                    lcl_accessoryname = "'" & lcl_accessoryname & "'"
                 end if

                 if lcl_accessoryvalue <> "" then
                    lcl_accessoryvalue = dbsafe(lcl_accessoryvalue)
                    lcl_accessoryvalue = "'" & lcl_accessoryvalue & "'"
                 end if

                 if lcl_accessoryid > 0 then
                   'BEGIN: Update accessory --------------------------------------
                    sSQLu = "UPDATE egov_class_teamroster_accessories SET "
                    sSQLu = sSQLu & " accessoryname = "  & lcl_accessoryname  & ", "
                    sSQLu = sSQLu & " accessoryvalue = " & lcl_accessoryvalue & ", "
                    sSQLu = sSQLu & " displayorder = "   & lcl_displayorder
                    sSQLu = sSQLu & " WHERE accessoryid = " & lcl_accessoryid

                   	set oUpdateAccessory = Server.CreateObject("ADODB.Recordset")
                   	oUpdateAccessory.Open sSQLu, Application("DSN"), 3, 1

                    set oUpdateAccessory = nothing
                   'END: Update accessory ----------------------------------------
                 else
                   'BEGIN: Insert accessory --------------------------------------
                    sSQLi = "INSERT INTO egov_class_teamroster_accessories ("
                    sSQLi = sSQLi & "orgid, "
                    sSQLi = sSQLi & "accessorytype, "
                    sSQLi = sSQLi & "accessoryname, "
                    sSQLi = sSQLi & "accessoryvalue, "
                    sSQLi = sSQLi & "displayorder "
                    sSQLi = sSQLi & ") VALUES ( "
                    sSQLi = sSQLi & lcl_orgid          & ", "
                    sSQLi = sSQLi & lcl_accessorytype  & ", "
                    sSQLi = sSQLi & lcl_accessoryname  & ", "
                    sSQLi = sSQLi & lcl_accessoryvalue & ", "
                    sSQLi = sSQLi & lcl_displayorder
                    sSQLi = sSQLi & ") "

                    lcl_accessoryid = RunInsertStatement(sSQLi)

                   'Determine if the user wants this new accessory assigned
                    if lcl_assign_accessoryid > 0 then
                       lcl_assign_accessoryid = clng(lcl_accessoryid)
                    end if
                   'END: Insert accessory ----------------------------------------
                 end if

                'BEGIN: Assign any/all of the options that have been "checked" ---
                 if lcl_assign_accessoryid > 0 then
                    sSQLa = "INSERT INTO egov_class_teamroster_accessories_to_class ("
                    sSQLa = sSQLa & "orgid, "
                    sSQLa = sSQLa & "classid, "
                    sSQLa = sSQLa & "accessoryid, "
                    sSQLa = sSQLa & "displayorder"
                    sSQLa = sSQLa & ") VALUES ("
                    sSQLa = sSQLa & lcl_orgid               & ", "
                    sSQLa = sSQLa & lcl_classid             & ", "
                    sSQLa = sSQLa & lcl_assign_accessoryid  & ", "
                    sSQLa = sSQLa & lcl_displayorder
                    sSQLa = sSQLa & ") "

                   	set oAssignAccessory = Server.CreateObject("ADODB.Recordset")
                   	oAssignAccessory.Open sSQLa, Application("DSN"), 3, 1

                    set oAssignAccessory = nothing
                 end if
                'END: Assign any/all of the options that have been "checked" -----
              end if
           end if
        next

        if i > 0 then
           lcl_success = "SU"
        end if
     end if
  end if

  lcl_url_parameters = ""
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "orgid",   lcl_orgid)
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "classid", lcl_classid)
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "atype",   lcl_atype)
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", lcl_success)

  response.redirect "class_accessoryoptions_list.asp" & lcl_url_parameters
%>