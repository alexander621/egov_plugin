<!-- #include file="../includes/common.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
  if request("user_action") = "" then
     response.redirect "datamgr_list.asp"
  end if

  if request("f") <> "" then
     lcl_feature = request("f")
  else
     lcl_feature = ""
  end if

 'Check for org features
  lcl_orghasfeature_large_address_list = orghasfeature("large address list")

 'Setup variables
  lcl_useraction    = ""
  lcl_dmid          = 0
  lcl_dm_typeid     = 0
  lcl_mappointcolor = ""
  lcl_orgid         = request("orgid")
  lcl_categoryid    = 0
'  lcl_street_number   = request("residentstreetnumber")
'  lcl_street_address  = request("streetaddress")
'  sNumber             = ""
'  sPrefix             = ""
'  sAddress            = ""
'  sSuffix             = ""
'  sDirection          = ""
'  sValidStreet        = request("validstreet")
'  sCity               = request("city")
'  sState              = request("state")
'  sZip                = request("zip")
'  sLatitude           = 0.00
'  sLongitude          = 0.00
  'sCounty             = ""
  'sParcelID           = ""
  'sListedOwner        = ""
  'sLegalDescription   = ""
  'sResidentType       = ""
  'sRegisteredUserID   = 0
'  sSortStreetName     = ""
  lcl_isActive      = 1
  'sStatusID           = 0
  lcl_userid          = session("userid")
  lcl_current_date  = "'" & dbsafe(ConvertDateTimetoTimeZone()) & "'"
  lcl_redirect_url  = "datamgr_list.asp"

  'oSave("streetunit")       = request("streetunit")
  'oSave("county")           = request("county")
  'oSave("parcelidnumber")   = request("parcelidnumber")
  'oSave("listedowner")      = request("listedowner")
  'oSave("residenttype")     = request("residenttype")
  'oSave("legaldescription") = request("legaldescription")

  'if request("registereduserid") = "" then
  '   oSave("registereduserid") = 0
  'else
  '   oSave("registereduserid") = request("registereduserid")
  'end if

  if request("user_action") <> "" then
     lcl_useraction = UCASE(request("user_action"))
  end if

  if request("dmid") <> "" then
     lcl_dmid = request("dmid")
  end if

  if request("dm_typeid") <> "" then
     lcl_dm_typeid = request("dm_typeid")
  end if

  if request("categoryid") <> "" then
     lcl_categoryid = request("categoryid")
  end if

'  if request("mappointcolor") <> "" then
'     lcl_mappointcolor = request("mappointcolor")
'  else
'     lcl_mappointcolor = getMapPointTypePointColor(lcl_dm_typeid)
'  end if

 'Retrieve the search options
  lcl_sc_dm_typeid = ""

  if request("sc_dm_typeid") <> "" then
     lcl_sc_dm_typeid = request("sc_dm_typeid")
  end if

 'Build return parameters
  lcl_url_parameters = ""
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_dm_typeid", lcl_sc_dm_typeid)

  if lcl_feature <> "datamgr_maint" then
     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
  end if

 'Execute the user's action
  if lcl_useraction = "DELETE" then
    'Clean up all tables BEFORE actually deleting the DM Data record
    '1. delete all "field values" associated to this DMID
    '2. delete all sub-category assignments
    '3. delete the DM Data record
     lcl_subcategoryid = ""

     deleteDMValue "dmid", lcl_dmid
     deleteSubCategoryAssignments lcl_dmid, lcl_subcategoryid
     deleteDMOwners lcl_dmid

     sSQL = "DELETE FROM egov_dm_data WHERE dmid = " & lcl_dmid

   		set oDeleteDMData = Server.CreateObject("ADODB.Recordset")
    	oDeleteDMData.Open sSQL, Application("DSN"), 3, 1

     set oDeleteDMData = nothing

     lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SD")
     lcl_redirect_url   = "datamgr_list.asp" & lcl_url_parameters
  else
     if request("isActive") = "Y" then
        lcl_isActive = 1
     else
        lcl_isActive = 0
     end if

     if lcl_useraction = "UPDATE" then

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dmid", lcl_dmid)
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SU")
        lcl_redirect_url   = "datamgr_maint.asp" & lcl_url_parameters

    '---------------------------------------------------------------------------
     else  'New DM Data
    '---------------------------------------------------------------------------

        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "SA")

     end if

     lcl_dm_sectionid = ""
     lcl_dm_fieldid   = ""
     lcl_dm_importid  = ""

     maintainDMData lcl_userid, _
                    lcl_orgid, _
                    lcl_dmid, _
                    lcl_dm_typeid, _
                    lcl_dm_sectionid, _
                    lcl_dm_fieldid, _
                    lcl_isActive, _
                    lcl_categoryid, _
                    lcl_dm_importid, _
                    lcl_dmid

     if lcl_useraction = "ADD" then
        lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dmid", lcl_dmid)
        lcl_redirect_url   = "datamgr_maint.asp" & lcl_url_parameters
     end if


  end if

  response.redirect lcl_redirect_url
%>