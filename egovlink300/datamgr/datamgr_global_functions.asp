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
     elseif iSuccess = "CTE" then
        lcl_return = "User successfully changed to EDITOR"
     elseif iSuccess = "CTO" then
        lcl_return = "User successfully changed to OWNER"
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function getDMTypeByFeature(p_orgid, p_feature)

  lcl_return = ""

  if p_feature <> "" then
     sSQL = "SELECT dm_typeid "
     sSQL = sSQL & " FROM egov_dm_types "
     sSQL = sSQL & " WHERE UPPER(feature_public) = '" & UCASE(p_feature) & "' "
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
function getDMTypes(p_orgid)

  lcl_return = ""

  if p_orgid <> "" then
     sSQL = "SELECT distinct dmt.dm_typeid "
     sSQL = sSQL & " FROM egov_dm_types dmt, egov_dm_data dmd "
     sSQL = sSQL & " WHERE dmt.dm_typeid = dmd.dm_typeid "
     sSQL = sSQL & " AND dmd.orgid = " & p_orgid
     sSQL = sSQL & " AND dmt.isActive = 1 "
     sSQL = sSQL & " AND dmd.isActive = 1 "
     'sSQL = sSQL & " AND dmd.latitude is not null "
     'sSQL = sSQL & " AND dmd.latitude <> 0.00 "
     'sSQL = sSQL & " AND dmd.longitude is not null "
     'sSQL = sSQL & " AND dmd.longitude <> 0.00 "

     set oGetDMTypes = Server.CreateObject("ADODB.Recordset")
     oGetDMTypes.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTypes.eof then
        do while not oGetDMTypes.eof

           if lcl_return <> "" then
              lcl_return = lcl_return & "," & oGetDMTypes("dm_typeid")
           else
              lcl_return = oGetDMTypes("dm_typeid")
           end if

           oGetDMTypes.movenext
        loop
     end if

     oGetDMTypes.close
     set oGetDMTypes = nothing

  end if

  getDMTypes = lcl_return

end function

'------------------------------------------------------------------------------
sub getDMTypeInfo(ByVal iDMTypeID, _
                  ByVal p_orgid, _
                  ByVal iFeature, _
                  ByRef lcl_total_dmtypes, _
                  ByRef lcl_dm_typeid, _
                  ByRef lcl_description, _
                  ByRef lcl_mappointcolor, _
                  ByRef lcl_displayMap, _
                  ByRef lcl_enableOwnerMaint, _
                  ByRef lcl_useAdvancedSearch, _
                  ByRef lcl_dmt_latitude, _
                  ByRef lcl_dmt_longitude, _
                  ByRef lcl_defaultzoomlevel, _
                  ByRef lcl_googleMapType, _
                  ByRef lcl_googleMapMarker, _
                  ByRef lcl_accountInfoSectionID, _
                  ByRef lcl_defaultCategoryID, _
                  ByRef lcl_includeBlankCategoryOption, _
                  ByRef lcl_intro_message)

  lcl_dm_typeid                  = 0
  lcl_feature                    = ""
  lcl_description                = ""
  lcl_mappointcolor              = "green"
  lcl_total_dmtypes              = 0
  lcl_displayMap                 = true
  lcl_enableOwnerMaint           = false
  lcl_useAdvancedSearch          = false
  lcl_dmt_latitude               = ""
  lcl_dmt_longitude              = ""
  lcl_defaultzoomlevel           = "13"
  lcl_googleMapType              = "ROADMAP"
  lcl_googleMapMarker            = "GOOGLE"
  lcl_accountInfoSectionID       = 0
  lcl_defaultCategoryID          = ""
  lcl_includeBlankCategoryOption = 0
  lcl_intro_message              = ""

  if iFeature <> "" then
     lcl_feature = iFeature
  end if

 'Get the DM TypeID
  if iDMTypeID <> "" then
     lcl_dm_typeid = CLng(iDMTypeID)
  else
    'If a feature has been passed in and the DM Type could not be found then it means that the 
    'feature has not yet been assigned to the DM Type.  Therefore, we should not show the DM Data records.

     if lcl_feature = "" then
       'Check to see if org has only one DM Type.
       'If "yes" then show the DM Data records for that DM Type
       'If "no" then grab the first one in the list (ordered by description)
        lcl_dmtypes = getDMTypes(p_orgid)

        if lcl_dmtypes <> "" then
           'sSQL = "SELECT DISTINCT dm_typeid "
           'sSQL = sSQL & " FROM egov_dm_types "
           'sSQL = sSQL & " WHERE dm_typeid IN (" & lcl_dmtypes & ") "

           sSQL = "SELECT distinct dmt.dm_typeid "
           sSQL = sSQL & " FROM egov_dm_types dmt, "
           sSQL = sSQL &      " egov_dm_data dmd "
           sSQL = sSQL & " WHERE dmt.dm_typeid = dmd.dm_typeid "
           sSQL = sSQL & " AND dmd.orgid = " & p_orgid
           sSQL = sSQL & " AND dmt.dm_typeid IN (" & lcl_dmtypes & ") "
           sSQL = sSQL & " AND dmt.isActive = 1 "
           sSQL = sSQL & " AND dmd.isActive = 1 "
           'sSQL = sSQL & " AND dmd.latitude is not null "
           'sSQL = sSQL & " AND dmd.latitude <> 0.00 "
           'sSQL = sSQL & " AND dmd.longitude is not null "
           'sSQL = sSQL & " AND dmd.longitude <> 0.00 "

          	set oGetDefaultDMTypeID = Server.CreateObject("ADODB.Recordset")
           oGetDefaultDMTypeID.Open sSQL, Application("DSN"), 3, 1

           if not oGetDefaultDMTypeID.eof then
              lcl_dm_typeid = oGetDefaultDMTypeID("dm_typeid")
           end if

           oGetDefaultDMTypeID.close
           set oGetDefaultDMTypeID = nothing
        end if
     end if
  end if

 'Get the DM Type Info
  if lcl_dm_typeid <> "" then
     sSQL = "SELECT dm_typeid, "
     sSQL = sSQL & " description, "
     sSQL = sSQL & " isnull(mappointcolor, 'green') as mappointcolor, "
     sSQL = sSQL & " displayMap, "
     sSQL = sSQL & " enableOwnerMaint, "
     sSQL = sSQL & " useAdvancedSearch, "
     sSQL = sSQL & " latitude, "
     sSQL = sSQL & " longitude, "
     sSQL = sSQL & " defaultzoomlevel, "
     sSQL = sSQL & " googleMapType, "
     sSQL = sSQL & " googleMapMarker, "
     sSQL = sSQL & " accountInfoSectionID, "
     sSQL = sSQL & " defaultcategoryid, "
     sSQL = sSQL & " includeBlankCategoryOption, "
     sSQL = sSQL & " intro_message "
     sSQL = sSQL & " FROM egov_dm_types "
     sSQL = sSQL & " WHERE dm_typeid = " & lcl_dm_typeid

    	set oGetDMTypeInfo = Server.CreateObject("ADODB.Recordset")
     oGetDMTypeInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTypeInfo.eof then
        lcl_dm_typeid                  = oGetDMTypeInfo("dm_typeid")
        lcl_description                = oGetDMTypeInfo("description")
        lcl_mappointcolor              = oGetDMTypeInfo("mappointcolor")
        lcl_displayMap                 = oGetDMTypeInfo("displayMap")
        lcl_enableOwnerMaint           = oGetDMTypeInfo("enableOwnerMaint")
        lcl_useAdvancedSearch          = oGetDMTypeInfo("useAdvancedSearch")
        lcl_dmt_latitude               = oGetDMTypeInfo("latitude")
        lcl_dmt_longitude              = oGetDMTypeInfo("longitude")
        lcl_defaultzoomlevel           = oGetDMTypeInfo("defaultzoomlevel")
        lcl_googleMapType              = oGetDMTypeInfo("googleMapType")
        lcl_googleMapMarker            = oGetDMTypeInfo("googleMapMarker")
        lcl_accountInfoSectionID       = oGetDMTypeInfo("accountInfoSectionID")
        lcl_defaultCategoryID          = oGetDMTypeInfo("defaultcategoryid")
        lcl_includeBlankCategoryOption = oGetDMTypeInfo("includeBlankCategoryOption")
	if oGetDMTypeInfo("intro_message") <> "" and not isnull(oGetDMTypeInfo("intro_message")) then
        	lcl_intro_message              = replace(oGetDMTypeInfo("intro_message"),"http://www.egovlink.com","https://www.egovlink.com")
	end if
     end if

     oGetDMTypeInfo.close
     set oGetDMTypeInfo = nothing

    'Find the total available, active, DM Types for the org
     lcl_total_dmtypes = getTotalDMTypes(p_orgid)
  end if

end sub

'------------------------------------------------------------------------------
function getDMTypes(p_orgid)

  lcl_return = ""

  if p_orgid <> "" then
     sSQL = "SELECT distinct dmt.dm_typeid "
     sSQL = sSQL & " FROM egov_dm_types dmt, egov_dm_data dmd "
     sSQL = sSQL & " WHERE dmt.dm_typeid = dmd.dm_typeid "
     sSQL = sSQL & " AND dmd.orgid = " & p_orgid
     sSQL = sSQL & " AND dmt.isActive = 1 "
     sSQL = sSQL & " AND dmd.isActive = 1 "
     'sSQL = sSQL & " AND dmd.latitude is not null "
     'sSQL = sSQL & " AND dmd.latitude <> 0.00 "
     'sSQL = sSQL & " AND dmd.longitude is not null "
     'sSQL = sSQL & " AND dmd.longitude <> 0.00 "

     set oGetDMTypes = Server.CreateObject("ADODB.Recordset")
     oGetDMTypes.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTypes.eof then
        do while not oGetDMTypes.eof

           if lcl_return <> "" then
              lcl_return = lcl_return & "," & oGetDMTypes("dm_typeid")
           else
              lcl_return = oGetDMTypes("dm_typeid")
           end if

           oGetDMTypes.movenext
        loop
     end if

     oGetDMTypes.close
     set oGetDMTypes = nothing

  end if

  getDMTypes = lcl_return

end function

'------------------------------------------------------------------------------
function getTotalDMTypes(p_orgid)
  lcl_return = 0

  if p_orgid <> "" then
     sSQL = "SELECT count(distinct dmt.dm_typeid) as total_dmtypes "
     sSQL = sSQL & " FROM egov_dm_types dmt, egov_dm_data dmd "
     sSQL = sSQL & " WHERE dmt.dm_typeid = dmd.dm_typeid "
     sSQL = sSQL & " AND dmd.orgid = " & p_orgid
     sSQL = sSQL & " AND dmt.isActive = 1 "
     sSQL = sSQL & " AND dmd.isActive = 1 "
     'sSQL = sSQL & " AND dmd.latitude is not null "
     'sSQL = sSQL & " AND dmd.latitude <> 0.00 "
     'sSQL = sSQL & " AND dmd.longitude is not null "
     'sSQL = sSQL & " AND dmd.longitude <> 0.00 "

     set oGetTotalDMTypes = Server.CreateObject("ADODB.Recordset")
     oGetTotalDMTypes.Open sSQL, Application("DSN"), 3, 1

     if not oGetTotalDMTypes.eof then
        lcl_return = oGetTotalDMTypes("total_dmtypes")
     end if

     oGetTotalDMTypes.close
     set oGetTotalDMTypes = nothing

  end if

  getTotalDMTypes = lcl_return

end function

'------------------------------------------------------------------------------
sub GetCityPoint(ByVal p_orgid, ByRef sCityLat, ByRef sCityLng)
    sCityLat  = ""
    sCityLng  = ""

   'Get the point to center the map
    sSQL = "SELECT latitude, longitude "
    sSQL = sSQL & " FROM organizations "
    sSQL = sSQL & " WHERE orgid = " & p_orgid

    set oCityPoint = Server.CreateObject("ADODB.Recordset")
    oCityPoint.Open sSQL, Application("DSN"), 3, 1

    if not oCityPoint.eof then
       sCityLat = oCityPoint("latitude")
       sCityLng = oCityPoint("longitude")
    end if

    oCityPoint.close
    set oCityPoint = nothing

end sub

'------------------------------------------------------------------------------
sub displayDMCategories(iOrgID, iFeature)

  sSQL = "select categoryid, "
  sSQL = sSQL & " categoryname, "
  sSQL = sSQL & " orgid, "
  sSQL = sSQL & " isActive, "
  sSQL = sSQL & " createdbyid, "
  sSQL = sSQL & " createdbydate, "
  sSQL = sSQL & " lastmodifiedbyid, "
  sSQL = sSQL & " lastmodifiedbydate, "
  sSQL = sSQL & " parent_categoryid "
  sSQL = sSQL & " from egov_dm_categories "
  sSQL = sSQL & " where categoryid in (select distinct categoryid "
  sSQL = sSQL &                      " from egov_dm_data "
  sSQL = sSQL &                      " where orgid = " & iOrgID
  sSQL = sSQL &                      " and isActive = 1) "
  sSQL = sSQL & " and isActive = 1 "

  set oDisplayDMCategories = Server.CreateObject("ADODB.Recordset")
  oDisplayDMCategories.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayDMCategories.eof then

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
     response.write "  <tr>" & vbcrlf

     do while not oDisplayDMCategories.eof

        lcl_category_url = "datamgr.asp"
        lcl_category_url = lcl_category_url & "?f=" & iFeature
        lcl_category_url = lcl_category_url & "&cid=" & oDisplayDMCategories("categoryid")

        response.write "      <td><a href=""" & lcl_category_url & """>"& oDisplayDMCategories("categoryname") & "</a></td>" & vbcrlf

        oDisplayDMCategories.movenext
     loop

     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

  end if

  oDisplayDMCategories.close
  set oDisplayDMCategories = nothing

end sub

'------------------------------------------------------------------------------
function getOriginalLayoutID()

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
function checkLayoutExists(iLayoutID)
  lcl_return = False

  if iLayoutID <> "" then
     sSQL = "SELECT 'Y' AS lcl_exists "
     sSQL = sSQL & " FROM egov_dm_layouts "
     sSQL = sSQL & " WHERE layoutid = " & iLayoutID

     set oCheckLayoutExists = Server.CreateObject("ADODB.Recordset")
     oCheckLayoutExists.Open sSQL, Application("DSN"), 3, 1

     if not oCheckLayoutExists.eof then
        if oCheckLayoutExists("lcl_exists") = "Y" then
           lcl_return = True
        end if
     end if

     oCheckLayoutExists.close
     set oCheckLayoutExists = nothing
  end if

  checkLayoutExists = lcl_return

end function

'------------------------------------------------------------------------------
sub getLayoutInfo(ByVal iLayoutID, ByRef lcl_layoutname, ByRef lcl_isOriginalLayout, _
                  ByRef lcl_useLayoutSections, ByRef lcl_totalcolumns, ByRef lcl_columnwidth_left, _
                  ByRef lcl_columnwidth_middle, ByRef lcl_columnwidth_right)

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
  lcl_current_date         = "'" & dbsafe(ConvertDateTimetoTimeZone(iOrgID)) & "'"
  lcl_subcategoryid        = 0
  lcl_subcategoryname      = iSubCategoryName
  lcl_subisActive          = 0
  lcl_subisApproved        = 0
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
        'lcl_subisApproved    = 1
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
        lcl_subcategory_exists = checkSubCategoryExistsByCategoryName(lcl_subcategoryname, lcl_dm_typeid)

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
           sSQLs = sSQLs & "parent_categoryid "
           'sSQLs = sSQLs & "isApproved, "
           'sSQLs = sSQLs & "approvedeniedbyid, "
           'sSQLs = sSQLs & "approvedeniedbydate "
           'sSQLs = sSQLs & "mappointcolor"
           sSQLs = sSQLs & ") VALUES ("
           sSQLs = sSQLs & lcl_subcategoryname   & ", "
           sSQLs = sSQLs & lcl_orgid             & ", "
           sSQLs = sSQLs & lcl_dm_typeid         & ", "
           sSQLs = sSQLs & lcl_subisActive       & ", "
           sSQLs = sSQLs & sCreatedByID          & ", "
           sSQLs = sSQLs & sCreatedByDate        & ", "
           sSQLs = sSQLs & "NULL,NULL"           & ", "
           sSQLs = sSQLs & lcl_parentcategoryid
           'sSQLs = sSQLs & lcl_subisApproved     & ", "
           'sSQLs = sSQLs & sApprovedByID         & ", "
           'sSQLs = sSQLs & sApprovedByDate
           'sSQLs = sSQLs & lcl_mappointcolor
           sSQLs = sSQLs & ")"

        		'Get the Sub-CategoryID
 	        	'lcl_subcategoryid = RunIdentityInsert(sSQLs)
 	        	lcl_subcategoryid = RunIdentityInsertStatement(sSQLs)
        end if
     end if

  end if

 'Set up the return sub_categoryid
  lcl_return = lcl_subcategoryid

  maintainSubCategory = lcl_return

end function

'------------------------------------------------------------------------------
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
'sub maintainMapPointValues(iUserID, iOrgID, iDMTypeID, iDMID, iDMSectionID, iDMFieldID, iDMValueID, iFieldValue)
sub maintainDMValues(iUserID, iOrgID, iDMTypeID, iDMID, iDMSectionID, iDMFieldID, iDMValueID, iFieldValue)

  dim lcl_userid, lcl_orgid, lcl_dm_typeid, lcl_dm_sectionid
  dim lcl_dm_fieldid, lcl_dm_valueid, lcl_fieldvalue

  lcl_userid       = iUserID
  lcl_orgid        = iOrgID
  lcl_dm_typeid    = iDMTypeID
  lcl_dmid         = iDMID
  lcl_dm_sectionid = iDMSectionID
  lcl_dm_fieldid   = iDMFieldID
  lcl_dm_valueid   = iDMValueID
  lcl_fieldvalue   = formatFieldforInsertUpdate(iFieldValue)

 'If a dm_valueid exists then update the DM Data Value.  Otherwise, insert it.
  lcl_dm_valueid = getDMValueID(lcl_dm_valueid, lcl_dm_typeid, lcl_dmid, lcl_dm_sectionid, lcl_dm_fieldid)

  if lcl_dm_valueid = "" then
     lcl_dm_valueid = 0
  end if

  if lcl_dm_valueid > 0 then
     sSQL = "UPDATE egov_dm_values SET "
     sSQL = sSQL & " fieldvalue = "       & lcl_fieldvalue
     sSQL = sSQL & " WHERE dm_valueid = " & lcl_dm_valueid

     set oMaintainDMValues = Server.CreateObject("ADODB.Recordset")
     oMaintainDMValues.Open sSQL, Application("DSN"), 3, 1

     set oMaintainDMValues = nothing

  else
    'First check to see if a egov_dm_data record exists for the "dm_typeid, dm_sectionid, and dm_fieldid".
    'If "yes" then get use the dmid and create a record on egov_dm_values for it.
    'If "no" then create a new dmid and then use that id to create a record for it on egov_dm_values.
     'maintainDMData lcl_userid, lcl_orgid, lcl_dm_typeid, lcl_dm_sectionid, lcl_dm_fieldid, lcl_dmid

     sSQL = "INSERT INTO egov_dm_values ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "dm_typeid, "
     sSQL = sSQL & "dmid, "
     sSQL = sSQL & "dm_sectionid, "
     sSQL = sSQL & "dm_fieldid, "
     sSQL = sSQL & "fieldvalue "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & lcl_orgid            & ", "
     sSQL = sSQL & lcl_dm_typeid  & ", "
     sSQL = sSQL & lcl_dmid       & ", "
     sSQL = sSQL & lcl_dm_sectionid     & ", "
     sSQL = sSQL & lcl_dm_fieldid       & ", "
     sSQL = sSQL & lcl_fieldvalue
     sSQL = sSQL & ")"

  		'Get the DMID
 	  	'lcl_dm_valueid = RunIdentityInsert(sSQL)
 	  	lcl_dm_valueid = RunIdentityInsertStatement(sSQL)
  end if
end sub

'------------------------------------------------------------------------------
sub maintainDMData(ByVal iUserID, ByVal iOrgID, ByVal iDMID, ByVal iDMTypeID, ByVal iDMSectionID, _
                   ByVal iDMFieldID, ByVal iIsActive, ByVal iCategoryID, ByRef lcl_dmid)

  dim sUserID, sDMID, sDMTypeID, sDMSectionID, sDMFieldID
  dim sCategoryID, sIsActive, lcl_current_date, sIsCreatedByAdmin, sIsLastUpdatedByAdmin
  dim sSQLi, sSQLu, oUpdateDMData

  sUserID                = 0
  sDMID                  = 0
  sDMTypeID              = ""
  sDMSectionID           = ""
  sDMFieldID             = ""
  sCategoryID            = ""
  sIsActive              = ""
  sIsCreatedByAdmin      = 0
  sIsLastModifiedByAdmin = 0
  lcl_dmid               = 0
  lcl_current_date       = "'" & dbsafe(ConvertDateTimetoTimeZone(iOrgID)) & "'"

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

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

  if iIsActive <> "" then
     sIsActive = iIsActive
  end if

 'BEGIN: Update ---------------------------------------------------------------
  if sDMID > 0 then

   		sSQLu = "UPDATE egov_dm_data SET "
     sSQLu = sSQLu & "isLastModifiedByAdmin = " & sIsLastModifiedByAdmin & ", "
     sSQLu = sSQLu & "lastmodifiedbyid = "      & sUserID                & ", "
     sSQLu = sSQLu & "lastmodifiedbydate = "    & lcl_current_date       & ", "
     sSQLu = sSQLu & "dm_typeid = "             & sDMTypeID              & ", "
     sSQLu = sSQLu & "categoryid = "            & sCategoryID            & ", "
     sSQLu = sSQLu & "isActive = "              & sIsActive
     sSQLu = sSQLu & " WHERE dmid = " & sDMID

   		set oUpdateDMData = Server.CreateObject("ADODB.Recordset")
    	oUpdateDMData.Open sSQLu, Application("DSN"), 3, 1

     set oUpdateDMData = nothing

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

     'if sIsActive = "" then
     '   sIsActive = 1
     'end if

     sIsActive             = 0
     sCreatedByID          = sUserID
     sCreatedByDate        = lcl_current_date
     'sApprovedDeniedByID   = sUserID
     'sApprovedDeniedByDate = lcl_current_date

    'Insert the new DM Data
     sSQLi = "INSERT INTO egov_dm_data ("
     sSQLi = sSQLi & "dm_typeid, "
     sSQLi = sSQLi & "orgid, "
     sSQLi = sSQLi & "categoryid, "
     sSQLi = sSQLi & "isCreatedByAdmin, "
     sSQLi = sSQLi & "createdbyid, "
     sSQLi = sSQLi & "createdbydate, "
     'sSQLi = sSQLi & "isApproved, "
     'sSQLi = sSQLi & "approvedeniedbyid, "
     'sSQLi = sSQLi & "approvedeniedbydate, "
     'sSQLi = sSQLi & "lastmodifiedbyid, "
     'sSQLi = sSQLi & "lastmodifiedbydate, "
     sSQLi = sSQLi & "isActive "
     sSQLi = sSQLi & ") VALUES ("
     sSQLi = sSQLi & sDMTypeID             & ", "
     sSQLi = sSQLi & iOrgID                & ", "
     sSQLi = sSQLi & sCategoryID           & ", "
     sSQLi = sSQLi & sIsCreatedByAdmin     & ", "
     sSQLi = sSQLi & sCreatedByID          & ", "
     sSQLi = sSQLi & sCreatedByDate        & ", "
     'sSQLi = sSQLi & "1, "
     'sSQLi = sSQLi & sApprovedDeniedByID   & ", "
     'sSQLi = sSQLi & sApprovedDeniedByDate & ", "
     'sSQLi = sSQLi & "NULL,NULL"           & ", "
     sSQLi = sSQLi & sIsActive
     sSQLi = sSQLi & ")"

    'Get the DMID
     'lcl_dmid = RunIdentityInsert(sSQLi)
     lcl_dmid = RunIdentityInsertStatement(sSQLi)

    'Make the user the unapproved owner of this DM Data record
     lcl_owner_type              = "OWNER"
     lcl_isApprovedDeniedByAdmin = false
     lcl_isApproved              = true

     insertOwnerEditor iOrgID, lcl_dmid, sUserID, lcl_owner_type, lcl_isApprovedDeniedByAdmin, lcl_isApproved

  end if

end sub

'------------------------------------------------------------------------------
sub insertOwnerEditor(iOrgID, iDMID, iUserID, iOwnerType, iIsApprovedDeniedByAdmin, iIsApproved)

  dim sOrgID, sDMID, sUserID, sOwnerType, sIsApprovedDeniedByAdmin, sIsApproved, lcl_current_date
  dim sSQL, lcl_dm_ownerid

  sOrgID                   = 0
  sDMID                    = 0
  sUserID                  = 0
  sOwnerType               = "EDITOR"
  sIsApprovedDeniedByAdmin = 0
  sIsApproved              = 0
  lcl_dm_ownerid           = 0
  lcl_current_date         = "NULL"

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  if iOwnerType <> "" then
     sOwnerType = ucase(iOwnerType)
     sOwnerType = "'" & dbsafe(sOwnerType) & "'"
  end if

  if iIsApprovedDeniedByAdmin <> "" then
     if iIsApprovedDeniedByAdmin then
        sIsApprovedDeniedByAdmin = 1
     end if
  end if

  if iIsApproved <> "" then
     if iIsApproved then
        sIsApproved = 1
     else
        sIsApproved = 0
     end if

     lcl_current_date = "'" & dbsafe(ConvertDateTimetoTimeZone(sOrgID)) & "'"
  end if

  sSQL = "INSERT INTO egov_dm_owners ("
  sSQL = sSQL & "orgid, "
  sSQL = sSQL & "dmid, "
  sSQL = sSQL & "userid, "
  sSQL = sSQL & "ownertype, "
  sSQL = sSQL & "isApprovedDeniedByAdmin, "
  sSQL = sSQL & "isApproved, "
  sSQL = sSQL & "approvedeniedbyid, "
  sSQL = sSQL & "approvedeniedbydate "
  sSQL = sSQL & ") VALUES ("
  sSQL = sSQL & sOrgID                   & ", "
  sSQL = sSQL & sDMID                    & ", "
  sSQL = sSQL & sUserID                  & ", "
  sSQL = sSQL & sOwnerType               & ", "
  sSQL = sSQL & sIsApprovedDeniedByAdmin & ", "
  sSQL = sSQL & sIsApproved              & ", "
  sSQL = sSQL & sUserID                  & ", "
  sSQL = sSQL & lcl_current_date
  sSQL = sSQL & ") "

  lcl_dm_ownerid = RunIdentityInsertStatement(sSQL)

end sub

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
function checkSubCategoryExistsByCategoryName(iSubCategoryName, iDM_TypeID)

  dim lcl_return, sSubCategoryName, sDM_TypeID

  lcl_return       = false
  sSubCategoryName = ""
  sDM_TypeID       = 0

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

  if sDM_TypeID > 0 AND sSubCategoryName <> "" then
     sSQL = "SELECT distinct 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_dm_categories "
     sSQL = sSQL & " WHERE dm_typeid = " & sDM_TypeID
     sSQL = sSQL & " AND UPPER(categoryname) = " & sSubCategoryName

     set oCheckSCExistsByCategoryName = Server.CreateObject("ADODB.Recordset")
     oCheckSCExistsByCategoryName.Open sSQL, Application("DSN"), 3, 1

     if not oCheckSCExistsByCategoryName.eof then
        lcl_return = true
     end if

     oCheckSCExistsByCategoryName.close
     set oCheckSCExistsByCategoryName = nothing

  end if

  checkSubCategoryExistsByCategoryName = lcl_return

end function

'------------------------------------------------------------------------------
function checkDMOwnerExists(iDMID)
  dim lcl_return, sSQL, sDMID

  lcl_return = false

  if iDMID <> "" then
     sDMID = clng(iDMID)

     sSQL = "SELECT distinct 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_dm_owners "
     sSQL = sSQL & " WHERE dmid = " & sDMID
     sSQL = sSQL & " AND ownertype = 'OWNER' "
     sSQL = sSQL & " AND isApproved = 1 "

     set oCheckDMOwnerExists = Server.CreateObject("ADODB.Recordset")
     oCheckDMOwnerExists.Open sSQL, Application("DSN"), 3, 1

     if not oCheckDMOwnerExists.eof then
        if oCheckDMOwnerExists("lcl_exists") = "Y" then
           lcl_return = true
        end if
     end if

  end if

  checkDMOwnerExists = lcl_return

end function

'------------------------------------------------------------------------------
function checkDMIDExists(iOrgID, iDMID)

  dim sSQL, lcl_return, sOrgID, sDMID

  lcl_return = false
  sOrgID     = 0
  sDMID      = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  sSQL = "SELECT distinct dmid "
  sSQL = sSQL & " FROM egov_dm_data "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND dmid = " & sDMID

  set oCheckDMIDExists = Server.CreateObject("ADODB.Recordset")
  oCheckDMIDExists.Open sSQL, Application("DSN"), 3, 1

  if not oCheckDMIDExists.eof then
     lcl_return = true
  end if

  oCheckDMIDExists.close
  set oCheckDMIDExists = nothing

  checkDMIDExists = lcl_return

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
function getOwnerName(iOwnerID)

  dim lcl_return, lcl_ownerid

  lcl_return    = ""
  lcl_ownerid   = 0
  lcl_ownername = ""

  if iOwnerID <> "" then
     lcl_ownerid = clng(iOwnerID)

     sSQL = "SELECT userfname, "
     sSQL = sSQL & " userlname "
     sSQL = sSQL & " FROM egov_users "
     sSQL = sSQL & " WHERE userid = " & lcl_ownerid

     set oGetOwnerName = Server.CreateObject("ADODB.Recordset")
     oGetOwnerName.Open sSQL, Application("DSN"), 3, 1

     if not oGetOwnerName.eof then
        if oGetOwnerName("userfname") <> "" then
           if lcl_ownername <> "" then
              lcl_ownername = lcl_ownername & oGetOwnerName("userfname")
           else
              lcl_ownername = oGetOwnerName("userfname")
           end if
        end if

        if oGetOwnerName("userlname") <> "" then
           if lcl_ownername <> "" then
              lcl_ownername = lcl_ownername & " " & oGetOwnerName("userlname")
           else
              lcl_ownername = oGetOwnerName("userlname")
           end if
        end if
     end if

     if lcl_ownername <> "" then
        lcl_return = lcl_ownername
     end if

     oGetOwnerName.close
     set oGetOwnerName = nothing

  end if

  getOwnerName = lcl_return

end function

'------------------------------------------------------------------------------
sub addSubCategoryAssignment(iOrgID, iDMTypeID, iDMID, iSubCategoryID)

  dim sOrgID, sDMTypeID, sDMID, sSubCategoryID

  sOrgID         = 0
  sDMTypeID      = 0
  sDMID          = 0
  sSubCategoryID = ""

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

  if sDMTypeID > 0 AND sDMID > 0 AND sSubCategoryID <> "" then  
     sSQL = "INSERT INTO egov_dmdata_to_dmcategories ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "dm_typeid, "
     sSQL = sSQL & "dmid, "
     sSQL = sSQL & "categoryid"
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & sOrgID    & ", "
     sSQL = sSQL & sDMTypeID & ", "
     sSQL = sSQL & sDMID     & ", "
     sSQL = sSQL & sSubCategoryID
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
'function RunIdentityInsert( sInsertStatement )
'	 dim sSQLidentity, iReturnValue, oIdentityInsert

'	 iReturnValue = 0

	'Insert new row into database and get rowid
' 	sSQLidentity = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

' 	set oIdentityInsert = Server.CreateObject("ADODB.Recordset")
'	 oIdentityInsert.Open sSQLidentity, Application("DSN"), 3, 3

' 	iReturnValue = oIdentityInsert("ROWID")

' 	oIdentityInsert.close
'	 set oIdentityInsert = nothing

' 	RunIdentityInsert = iReturnValue

'end function

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

'dtb_debug("lcl_streetaddress: [" & lcl_streetaddress & "] - lcl_compare_address: [" & lcl_compare_address & "] - [" & lcl_selected_address & "]")

       response.write "<option value=""" & lcl_streetaddress & """" & lcl_selected_address & ">" & lcl_streetaddress & "</option>" & vbcrlf

    			oAddressList.MoveNext
    loop

    response.write "</select>&nbsp;" & vbcrlf
    response.write "<input type=""button"" id=""validateAddress"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
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
function getColumnNumber(iTotalColumns, iColumnLocation)

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
  lcl_current_date = "'" & dbsafe(ConvertDateTimetoTimeZone(iOrgID)) & "'"

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

   		sSQL = "UPDATE egov_dm_data SET "
     sSQL = sSQL & "lastmodifiedbyid = "   & sUserID          & ", "
     sSQL = sSQL & "lastmodifiedbydate = " & lcl_current_date & ", "
     sSQL = sSQL & "dm_typeid = "          & sDMTypeID        & ", "
     sSQL = sSQL & "streetnumber = "       & sNumber          & ", "
     sSQL = sSQL & "streetprefix = "       & sPrefix          & ", "
     sSQL = sSQL & "streetaddress = "      & sAddress         & ", "
     sSQL = sSQL & "streetsuffix = "       & sSuffix          & ", "
     sSQL = sSQL & "streetdirection = "    & sDirection       & ", "
     sSQL = sSQL & "sortstreetname = "     & sSortStreetName  & ", "
     sSQL = sSQL & "city = "               & sCity            & ", "
     sSQL = sSQL & "state = "              & sState           & ", "
     sSQL = sSQL & "zip = "                & sZip             & ", "
     sSQL = sSQL & "validstreet = "        & sValidStreet     & ", "
     sSQL = sSQL & "latitude = "           & sLatitude        & ", "
     sSQL = sSQL & "longitude = "          & sLongitude
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

     sCreatedByID   = sUserID
     sCreatedByDate = lcl_current_date

    'Insert the new Map-Point
     sSQL = "INSERT INTO egov_dm_data ("
     sSQL = sSQL & "dm_typeid, "
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "createdbyid, "
     sSQL = sSQL & "createdbydate, "
     sSQL = sSQL & "lastmodifiedbyid, "
     sSQL = sSQL & "lastmodifiedbydate, "
     sSQL = sSQL & "streetnumber, "
     sSQL = sSQL & "streetprefix, "
     sSQL = sSQL & "streetaddress, "
     sSQL = sSQL & "streetsuffix, "
     sSQL = sSQL & "streetdirection, "
     sSQL = sSQL & "sortstreetname, "
     sSQL = sSQL & "city, "
     sSQL = sSQL & "state, "
     sSQL = sSQL & "zip, "
     sSQL = sSQL & "validstreet, "
     sSQL = sSQL & "latitude, "
     sSQL = sSQL & "longitude "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & sDMTypeID  & ", "
     sSQL = sSQL & iOrgID           & ", "
     sSQL = sSQL & sCreatedByID     & ", "
     sSQL = sSQL & sCreatedByDate   & ", "
     sSQL = sSQL & "NULL,NULL"      & ", "
     sSQL = sSQL & sStreetNumber    & ", "
     sSQL = sSQL & sStreetPrefix    & ", "
     sSQL = sSQL & sStreetAddress   & ", "
     sSQL = sSQL & sStreetSuffix    & ", "
     sSQL = sSQL & sStreetDirection & ", "
     sSQL = sSQL & sSortStreetName  & ", "
     sSQL = sSQL & sCity            & ", "
     sSQL = sSQL & sState           & ", "
     sSQL = sSQL & sZip             & ", "
     sSQL = sSQL & sValidStreet     & ", "
     sSQL = sSQL & sLatitude        & ", "
     sSQL = sSQL & sLongitude
     sSQL = sSQL & ")"

    'Get the DMID
     'lcl_dmid = RunIdentityInsert(sSQL)
     lcl_dmid = RunIdentityInsertStatement(sSQL)
  end if

end sub

'------------------------------------------------------------------------------
function getFieldValue_by_DMValueID(iDMValueID)
  dim lcl_return, sDMValueID

  lcl_return = ""
  sDMValueID = 0

  if iDMValueID <> "" then
     sDMValueID = clng(iDMValueID)
  end if

  sSQL = "SELECT fieldvalue "
  sSQL = sSQL & " FROM egov_dm_values "
  sSQL = sSQL & " WHERE dm_valueid = " & sDMValueID

  set oGetDMFieldValue = Server.CreateObject("ADODB.Recordset")
  oGetDMFieldValue.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMFieldValue.eof then
     lcl_return = oGetDMFieldValue("fieldvalue")
  end if

  oGetDMFieldValue.close
  set oGetDMFieldValue = nothing

  getFieldValue_by_DMValueID = lcl_return

end function

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
sub dtb_debug(iValue)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(iValue,"'","''") & "') "

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub

'------------------------------------------------------------------------------
function dbsafe(iValue)

  lcl_return = ""

  if iValue <> "" then
     lcl_return = replace(iValue,"'","''")
  end if

  dbsafe = lcl_return


end function
%>
