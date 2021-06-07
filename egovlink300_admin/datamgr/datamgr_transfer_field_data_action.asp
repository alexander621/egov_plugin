<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
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
  dim lcl_transfer_field_data, lcl_totalfields
  dim i

  lcl_totalfields = 0

  if request("totalfields") <> "" then
     lcl_totalfields = request("totalfields")

     if not isnumeric(lcl_totalfields) then
        response.redirect "datamgr_transfer_field_data.asp"
     end if
  end if

  lcl_totalfields = clng(lcl_totalfields)

 'Retrieve the search options
  lcl_sc_dm_typeid = ""

  if request("sc_dm_typeid") <> "" then
     lcl_sc_dm_typeid = request("sc_dm_typeid")
  end if

 'Build return parameters
  lcl_url_parameters = ""
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_dm_typeid", lcl_sc_dm_typeid)

 'Determine which row has been selected.
  i = 0

  if lcl_totalfields > 0 then
     for i = 1 to lcl_totalfields
         lcl_transfer_field_data = ""

         if request("transfer_field_data" & i) <> "" then
            lcl_dm_sectionid_current = request("dm_sectionid" & i)
            lcl_dm_fieldid_current   = request("dm_fieldid" & i)

            lcl_transfer_field_data  = split(request("transfer_field_data" & i), "_")
            lcl_dm_sectionid_new     = replace(lcl_transfer_field_data(0),"dmsectionid","")
            lcl_dm_fieldid_new       = replace(lcl_transfer_field_data(1),"dmfieldid","")

            sSQL = "UPDATE egov_dm_values SET "
            sSQL = sSQL & " dm_sectionid = " & lcl_dm_sectionid_new & ", "
            sSQL = sSQL & " dm_fieldid = "   & lcl_dm_fieldid_new
            sSQL = sSQL & " WHERE dm_sectionid = " & lcl_dm_sectionid_current
            sSQL = sSQL & " AND dm_fieldid = "     & lcl_dm_fieldid_current

          		set oTransferDataValues = Server.CreateObject("ADODB.Recordset")
           	oTransferDataValues.Open sSQL, Application("DSN"), 3, 1

            set oTransferDataValues = nothing

         end if
     next
  end if


  lcl_redirect_url = "datamgr_transfer_field_data.asp"
  lcl_redirect_url = lcl_redirect_url & lcl_url_parameters

  response.redirect lcl_redirect_url
%>