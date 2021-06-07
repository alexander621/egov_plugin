<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr.asp
' AUTHOR:   David Boyer
' CREATED:  04/14/2011
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the custom "Data Manager" data list.
'
' MODIFICATION HISTORY
' 1.0  04/14/11	 David Boyer - Initial Version
' 1.1  11/19/13  Terry Foster - CLng Bug Fix
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 Dim oDataMgr
     twfX = 2
     dim arrHeaders()

 set oDataMgr = New classOrganization

 'lcl_permission_check = "datamgr"
 lcl_feature = ""
 lcl_dmt     = ""

'Determine what type of "dm data" are to be displayed
'If there is only a single DM Type then show those DM Data records.
'If there are mulitple then grab the first available
'Check to see if the "dm data" record exists
'NOTE: request("d") = dm_typeid
 if request("d") <> "" then
    if not containsApostrophe(request("d")) then
       lcl_dmt = request("d")
    end if
 else

   'Determine if we are accessing a specific DM Data record via a feature (link).
   'For example, features may be created for specific DM Types such as "Available Properties".
   'Those features may be turned on for the public so that there will be a link in the menu bar, welcome page, and footer.
   'Those links will access this page.  The problem we run into is that we don't know which "Available Properties" DM Data to
   'retrieve since it's a different ID for each org.  Passing in the feature name and having it associated to the DM Type
   'will let us find the correct DM Type.
   'NOTE: request("f") = feature name of the dm type to be displayed.
    if request("f") <> "" then
       if not containsApostrophe(request("f")) then
          lcl_feature = Track_DBSafe(request("f"))
          lcl_dmt     = getDMTypeByFeature(iorgid, lcl_feature)
       end if
    end if
 end if

'If both a DM TypeID or Feature have NOT been passed in then try and find a DM TypeID to default to.
'If one cannot be found after this check then the DataMgr feature has not been set up correctly.  
'More than likely the feature simply needs to be turned on.
 if lcl_dmt = "" then
   'Check to see if org has only one DM Type.
   'If "yes" then show the DM Data for that DM Type
   'If "no" then grab the first one in the list (ordered by description)
    lcl_dmtypes = getDMTypes(iorgid)

    if lcl_dmtypes <> "" then
       sSQL = "SELECT distinct dmt.dm_typeid "
       sSQL = sSQL & " FROM egov_dm_types dmt, egov_dm_data dmd "
       sSQL = sSQL & " WHERE dmt.dm_typeid = dmd.dm_typeid "
       sSQL = sSQL & " AND dmd.orgid = " & iorgid
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
          lcl_dmt = oGetDefaultDMTypeID("dm_typeid")
       end if

       oGetDefaultDMTypeID.close
       set oGetDefaultDMTypeID = nothing
    end if

    if lcl_dmt = "" then
       lcl_dmt = 0
    end if
 end if

'Retrieve all of the DM TypeID info
 getDMTypeInfo lcl_dmt, _
               iorgid, _
               lcl_feature, _
               lcl_total_dmtypes, _
               lcl_dm_typeid, _
               lcl_description, _
               lcl_mappointcolor, _
               lcl_displayMap, _
               lcl_enableOwnerMaint, _
               lcl_useAdvancedSearch, _
               lcl_dmt_latitude, _
               lcl_dmt_longitude, _
               lcl_defaultzoomlevel, _
               lcl_googleMapType, _
               lcl_googleMapMarker, _
               lcl_accountInfoSectionID, _
               lcl_defaultCategoryID, _
               lcl_includeBlankCategoryOption, _
               lcl_intro_message

'We only need to get the map info if we are displaying the map
 if lcl_displayMap then

   'Get the city's map "center point"
    GetCityPoint iorgid, _
                 sCityLat, _
                 sCityLng

   'Set the Latitude, Longitude, and Zoom Level
    sLat             = "0.00"
    sLng             = "0.00"
    sZoom            = "13"
    sGoogleMapType   = "ROADMAP"
    sGoogleMapMarker = "GOOGLE"

    if lcl_dmt_latitude = "" or isnull(lcl_dmt_latitude) then
       if sCityLat = "" then
          sLat = sCityLat
       end if
    else
       sLat = lcl_dmt_latitude
    end if

    if lcl_dmt_longitude = "" or isnull(lcl_dmt_longitude) then
       if sCityLng = "" then
          sLng = sCityLng
       end if
    else
       sLng = lcl_dmt_longitude
    end if

    if lcl_defaultzoomlevel <> "" then
       sZoom = lcl_defaultzoomlevel
    end if

    if lcl_googleMapType <> "" then
       sGoogleMapType = ucase(lcl_googleMapType)
    end if

    if lcl_googleMapMarker <> "" then
       sGoogleMapMarker = ucase(lcl_googleMapMarker)
    end if
 else
    sLat             = ""
    sLng             = ""
    sZoom            = ""
    sGoogleMapType   = ""
    sGoogleMapMarker = ""
 end if

'BEGIN: Check for search criteria ---------------------------------------------
 dim lcl_total_subcategories

 lcl_searchvalue          = ""
 lcl_sc_searchfield_0     = ""
 lcl_sc_dm_categoryid     = 0
 lcl_sc_subcategory_ids   = ""
 lcl_sc_subcategory_names = ""
 lcl_total_subcategories  = 0

 if request("sc_searchvalue") <> "" then
    lcl_sc_searchfield_0 = request("sc_searchfield_0")

    lcl_searchvalue = lcl_sc_searchfield_0
    lcl_searchvalue = ucase(lcl_searchvalue)
    lcl_searchvalue = dbsafe(lcl_searchvalue)
    lcl_searchvalue = "'%" & lcl_searchvalue & "%'"
 end if

 if request("sc_dm_categoryid") <> "" and isnumeric(replace(request("sc_dm_categoryid"),"'","")) then
    lcl_sc_dm_categoryid = replace(request("sc_dm_categoryid"),"'","")
    lcl_sc_dm_categoryid = clng(lcl_sc_dm_categoryid)
 else
    'if not lcl_includeBlankCategoryOption AND lcl_defaultCategoryID <> "" then
    if lcl_defaultCategoryID <> "" then
       if lcl_defaultCategoryID > 0 then
          lcl_sc_dm_categoryid = lcl_defaultCategoryID
       end if
    end if
 end if

 if request("total_subcategories") <> "" then
    lcl_total_subcategories = request("total_subcategories")
    lcl_total_subcategories = clng(lcl_total_subcategories)
 end if

 if lcl_total_subcategories > 0 then
    i = 0
    for i = 1 to lcl_total_subcategories
       if request("subcategoryid" & i) <> "" then
          if lcl_sc_subcategory_ids <> "" then
             lcl_sc_subcategory_ids = lcl_sc_subcategory_ids & "," & request("subcategoryid" & i)
          else
             lcl_sc_subcategory_ids = request("subcategoryid" & i)
          end if

          if lcl_sc_subcategory_names <> "" then
             lcl_sc_subcategory_names = lcl_sc_subcategory_names & "," & request("subcategoryname" & i)
          else
             lcl_sc_subcategory_names = request("subcategoryname" & i)
          end if

         'Check to see if we need to repopulate the selected list.
         'We do this because if the "Advanced..." link was NOT clicked and sub-categories were not selected
         '  but criteria was changed in the search like the "search" input box (i.e. searching on "Chinese" to "Italian")
         '  then we want to continue to select any/all sub-categories that the user may have already selected.
         '*** NOTE: we will perform this check for the sub-categorynames as well
          if lcl_sc_subcategory_ids = "" then
             if request("subcategoryids") <> "" then
                if not containsApostrophe(request("subcategoryids")) then
                   lcl_sc_subcategory_ids = request("subcategoryids")
                end if
             end if
          end if

         'NOTE: NOT performing any validation check on this field as we are ONLY using it for display purposes.
          if lcl_sc_subcategory_names = "" then
             if request("subcategorynames") <> "" then
                lcl_sc_subcategory_names = request("subcategorynames")
             end if
          end if
       end if
    next
 end if
'END: Check for search criteria -----------------------------------------------

'BEGIN: Build the query to be used within the mapping functions ---------------
 lcl_query = " SELECT dmd.dmid, "
 lcl_query = lcl_query & "dmd.dm_typeid, "
 lcl_query = lcl_query & "dmtf.dm_fieldid, "
 lcl_query = lcl_query & "dmsf.fieldname, "
 lcl_query = lcl_query & "dmsf.fieldtype, "
 lcl_query = lcl_query & "dmv.fieldvalue, "
 lcl_query = lcl_query & "dmv.dm_valueid, "
 lcl_query = lcl_query & "dmtf.displayFieldName, "
 lcl_query = lcl_query & "dmtf.displayInResults, "
 lcl_query = lcl_query & "dmtf.isSidebarLink, "
 lcl_query = lcl_query & "dmtf.resultsOrder, "
 lcl_query = lcl_query & "dmd.streetnumber, "
 lcl_query = lcl_query & "dmd.streetprefix, "
 lcl_query = lcl_query & "dmd.streetaddress, "
 lcl_query = lcl_query & "dmd.streetsuffix, "
 lcl_query = lcl_query & "dmd.streetdirection, "
 lcl_query = lcl_query & "dmd.latitude, "
 lcl_query = lcl_query & "dmd.longitude, "
 lcl_query = lcl_query & "ISNULL(ISNULL(dmd.mappointcolor, dmc.mappointcolor), 'green') AS mappointcolor, "
 lcl_query = lcl_query & "dmc.categoryname, "
 lcl_query = lcl_query & "dmt.displayMap, "
 lcl_query = lcl_query & "dmt.useAdvancedSearch "
 lcl_query = lcl_query & " FROM egov_dm_data AS dmd "
 lcl_query = lcl_query &      " INNER JOIN egov_dm_types AS dmt ON dmt.dm_typeid = dmd.dm_typeid "
 lcl_query = lcl_query &      " INNER JOIN egov_dm_types_fields AS dmtf ON dmtf.dm_typeid = dmt.dm_typeid "
 lcl_query = lcl_query &      " INNER JOIN egov_dm_types_sections dmts ON dmts.dm_sectionid = dmtf.dm_sectionid "
 lcl_query = lcl_query &      " INNER JOIN egov_dm_sections_fields AS dmsf ON dmsf.section_fieldid = dmtf.section_fieldid "
 lcl_query = lcl_query &      " INNER JOIN egov_dm_categories AS dmc ON dmc.categoryid = dmd.categoryid "
 lcl_query = lcl_query &      " LEFT OUTER JOIN egov_dm_values AS dmv "
 lcl_query = lcl_query &                 " ON dmv.dm_fieldid = dmtf.dm_fieldid "
 lcl_query = lcl_query &                 " AND dmv.dmid = dmd.dmid "
 lcl_query = lcl_query &                 " AND dmv.dm_typeid = dmd.dm_typeid "
 lcl_query = lcl_query & " WHERE dmd.orgid = " & iOrgID
 lcl_query = lcl_query & " AND dmd.dm_typeid = " & lcl_dm_typeid
 lcl_query = lcl_query & " AND dmt.isActive = 1 "
 lcl_query = lcl_query & " AND dmts.isActive = 1 "
 lcl_query = lcl_query & " AND dmd.isActive = 1 "
 lcl_query = lcl_query & " AND dmsf.isActive = 1 "
 lcl_query = lcl_query & " AND dmd.isApproved = 1 "
 lcl_query = lcl_query & " AND dmtf.displayInResults = 1 "

 if lcl_displayMap then
    lcl_query = lcl_query & " AND dmd.latitude IS NOT NULL "
    lcl_query = lcl_query & " AND dmd.latitude <> 0.00 "
    lcl_query = lcl_query & " AND dmd.longitude IS NOT NULL "
    lcl_query = lcl_query & " AND dmd.longitude <> 0.00 "
 end if

'Get the search criteria fields - categories
 if lcl_sc_dm_categoryid > 0 then
    lcl_query = lcl_query & " AND dmd.categoryid = " & lcl_sc_dm_categoryid
 end if

 if lcl_sc_subcategory_ids <> "" then
    lcl_query = lcl_query & " AND dmd.dmid IN (select distinct dmid "
    lcl_query = lcl_query &                  " from egov_dmdata_to_dmcategories "
    lcl_query = lcl_query &                  " where orgid = " & iOrgID
    lcl_query = lcl_query &                  " and dm_typeid = " & lcl_dm_typeid
    lcl_query = lcl_query &                  " and categoryid in (" & lcl_sc_subcategory_ids & ") "
    lcl_query = lcl_query &                  ") "
 end if

'Get the search criteria fields - dynamic fields
 sSQL = "SELECT dm_fieldid, "
 sSQL = sSQL & " dm_typeid "
 'sSQL = sSQL & " fieldname, "
 'sSQL = sSQL & " fieldtype "
 sSQL = sSQL & " FROM egov_dm_types_fields "
 sSQL = sSQL & " WHERE dm_typeid = " & lcl_dm_typeid
 sSQL = sSQL & " AND inPublicSearch = 1 "
 sSQL = sSQL & " AND dm_sectionid IN (select distinct dm_sectionid "
 sSQL = sSQL &                      " from egov_dm_types_sections "
 sSQL = sSQL &                      " where orgid = " & iOrgID
 sSQL = sSQL &                      " and dm_typeid = " & lcl_dm_typeid
 sSQL = sSQL &                      " and isActive = 1 "
 sSQL = sSQL &                      ") "

 set oGetSearchCriteria = Server.CreateObject("ADODB.Recordset")
 oGetSearchCriteria.Open sSQL, Application("DSN"), 3, 1

 if not oGetSearchCriteria.eof then
    lcl_line_count_search = 0

    if NOT lcl_useAdvancedSearch then
       lcl_sc_searchfield = request("sc_searchfield_0")
    end if

    do while not oGetSearchCriteria.eof
       lcl_line_count_search = lcl_line_count_search + 1

      'Determine which search criteria layout to display.
      'if "lcl_useAdvancedSearch" = TRUE then show ALL of the fields selected to display in the search as searchable fields.
      'if "lcl_useAdvancedSearch" = FALSE then show only a single textbox and use that value to search all of the fields selected as searchable fields.
       if lcl_useAdvancedSearch then
          lcl_sc_searchfield = request("sc_searchfield_" & oGetSearchCriteria("dm_fieldid"))
       end if

       'if request("sc_searchfield_" & oGetSearchCriteria("mp_fieldid")) <> "" then
       '   lcl_fieldvalue = UCASE(request("sc_searchfield_" & oGetSearchCriteria("mp_fieldid")))
       if lcl_sc_searchfield <> "" then
          lcl_fieldvalue = UCASE(lcl_sc_searchfield)
          lcl_fieldvalue = dbsafe(lcl_fieldvalue)

          if NOT lcl_useAdvancedSearch then
             if lcl_line_count_search = 1 then
                lcl_query = lcl_query & " AND ("
             else
                lcl_query = lcl_query & " OR "
             end if
          else
             lcl_query = lcl_query & " AND "
          end if

          lcl_query = lcl_query &      " dmd.dmid in ("
          lcl_query = lcl_query &      " select distinct dmv" & oGetSearchCriteria("dm_fieldid") & ".dmid "
          lcl_query = lcl_query &      " from egov_dm_values dmv" & oGetSearchCriteria("dm_fieldid")
          lcl_query = lcl_query &      " where UPPER(dmv" & oGetSearchCriteria("dm_fieldid")& ".fieldvalue) LIKE ('%" & lcl_fieldvalue & "%') "
          lcl_query = lcl_query &      " AND dmv" & oGetSearchCriteria("dm_fieldid") & ".dm_fieldid = " & oGetSearchCriteria("dm_fieldid")
          lcl_query = lcl_query &      ") "

          'lcl_query = lcl_query & " AND ("
          'lcl_query = lcl_query &      " mptf.mp_fieldid = " & oGetSearchCriteria("mp_fieldid")
          'lcl_query = lcl_query &      " AND UPPER(mpv.fieldvalue) LIKE ('%" & lcl_fieldvalue & "%') "
          'lcl_query = lcl_query &      ") "
       else
          lcl_fieldvalue = ""
       end if

       oGetSearchCriteria.movenext
    loop

    if NOT lcl_useAdvancedSearch then
       if lcl_sc_searchfield <> "" then
          lcl_query = lcl_query & ")"
       end if
    end if

 end if

' lcl_query = lcl_query & " ORDER BY dmtf.resultsorder "

'response.write lcl_query & "<br />"
 oGetSearchCriteria.close
 set oGetSearchCriteria = nothing


' if lcl_searchvalue <> "" then
'    lcl_query = lcl_query & " AND upper(dmv.fieldvalue) LIKE (" & lcl_searchvalue & ") "
' end if
'response.write lcl_query & "<br />"
'END: Build the query to be used within the mapping functions -----------------

'BEGIN: Build the query to be used to get the column headers ------------------
 lcl_query_columnheaders = " SELECT distinct dmsf.section_fieldid, "
 lcl_query_columnheaders = lcl_query_columnheaders & " dmsf.fieldname, "
 lcl_query_columnheaders = lcl_query_columnheaders & " dmtf.resultsOrder "
 lcl_query_columnheaders = lcl_query_columnheaders & " FROM egov_dm_data AS dmd "
 lcl_query_columnheaders = lcl_query_columnheaders &      " INNER JOIN egov_dm_types AS dmt "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " ON dmt.dm_typeid = dmd.dm_typeid "
 lcl_query_columnheaders = lcl_query_columnheaders &      " INNER JOIN egov_dm_types_fields AS dmtf "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " ON dmtf.dm_typeid = dmt.dm_typeid "
 lcl_query_columnheaders = lcl_query_columnheaders &      " LEFT OUTER JOIN egov_dm_values AS dmv "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " ON dmv.dm_fieldid = dmtf.dm_fieldid "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " AND dmv.dmid = dmd.dmid "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " AND dmv.dm_typeid = dmd.dm_typeid "
 lcl_query_columnheaders = lcl_query_columnheaders &      " INNER JOIN egov_dm_types_sections dmts "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " ON dmts.dm_sectionid = dmtf.dm_sectionid "
 lcl_query_columnheaders = lcl_query_columnheaders &      " INNER JOIN egov_dm_sections_fields AS dmsf "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " ON dmsf.section_fieldid = dmtf.section_fieldid "
 lcl_query_columnheaders = lcl_query_columnheaders & " WHERE dmd.orgid = " & iOrgID
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.dm_typeid = " & lcl_dm_typeid
' lcl_query_columnheaders = lcl_query_columnheaders & " AND dmv.dmid > 0 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmt.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmts.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmsf.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.isApproved = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmtf.displayInResults = 1 "

 'if lcl_displayMap then
 '   lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.latitude IS NOT NULL "
 '   lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.latitude <> 0.00 "
 '   lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.longitude IS NOT NULL "
 '   lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.longitude <> 0.00 "
 'end if

' if lcl_sc_dm_categoryid > 0 then
'    lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.categoryid = " & lcl_sc_dm_categoryid
' end if

' if lcl_sc_subcategory_ids <> "" then
'    lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.dmid IN (select distinct dmid "
'    lcl_query_columnheaders = lcl_query_columnheaders &                  " from egov_dmdata_to_dmcategories "
'    lcl_query_columnheaders = lcl_query_columnheaders &                  " where orgid = " & iOrgID
'    lcl_query_columnheaders = lcl_query_columnheaders &                  " and dm_typeid = " & lcl_dm_typeid
'    lcl_query_columnheaders = lcl_query_columnheaders &                  " and categoryid in (" & lcl_sc_subcategory_ids & ") "
'    lcl_query_columnheaders = lcl_query_columnheaders &                  ") "
' end if

' if lcl_searchvalue <> "" then
'    lcl_query_columnheaders = lcl_query_columnheaders & " AND upper(dmv.fieldvalue) LIKE (" & lcl_searchvalue & ") "
' end if

 lcl_query_columnheaders = lcl_query_columnheaders & " ORDER BY dmtf.resultsOrder "
'END: Build the query to be used to get the column headers --------------------

'response.write lcl_query_columnheaders
'response.write lcl_query

'Check for org "edit displays"
 lcl_orghasdisplay_datamgr_intro = OrgHasDisplay(iorgid,"datamgr_intro")

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

'Set up the BODY "onload" and "onunload"
 if lcl_displayMap then
    lcl_onload   = "initialize();"
    'lcl_onunload = "GUnload();"
 end if

'Check for cookies
 lcl_cookie_userid = ""

 if request.cookies("userid") <> "" then
    lcl_cookie_userid = request.cookies("userid")
 end if

'Set up page redirect
 session("RedirectPage") = request.servervariables("script_name") & "?" & request.querystring()

'Build return parameters
 'lcl_url_parameters = ""
 'lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "d", lcl_dmt)
 'lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
%>
<html>
<head>

 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
<!--  <link rel="stylesheet" type="text/css" href="mapstyle.css" /> -->
<!--  <link rel="stylesheet" type="text/css" href="layout_styles.css" /> -->

 	<script type="text/javascript" src="../scripts/modules.js"></script>
 	<script type="text/javascript" src="../scripts/easyform.js"></script>
  <script type="text/javascript" src="../scripts/ajaxLib.js"></script>
  <script type="text/javascript" src="../scripts/setfocus.js"></script>
  <script type="text/javascript" src="../scripts/removespaces.js"></script>
 	<script type="text/javascript" src="../scripts/column_sorting.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

<%
'BEGIN: Google Maps javascript ------------------------------------------------
'Build the Google Maps javascript ONLY if we are displaying the map
 if lcl_displayMap then
sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"
    response.write "  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, user-scalable=no"" />" & vbcrlf
    response.write "  <script type=""text/javascript"" src=""https://maps.google.com/maps/api/js?sensor=false&key=" & sGoogleMapAPIKey & """></script>" & vbcrlf
%>
<script type="text/javascript">
  //Set up MapPoints and Info Windows
<% createMapPoints lcl_feature, lcl_query %>

  var markers    = [];
  var infowindow = [];
  var map;

  function initialize() {
    var myLatlng = new google.maps.LatLng(<%=sLat%>, <%=sLng%>);
    var myOptions = {
       mapTypeId: google.maps.MapTypeId.<%=sGoogleMapType%>,  //maptypes: ROADMAP, SATELLITE, HYBRID, TERRAIN
       zoom:      <%=sZoom%>,
       center:    myLatlng
    }

    map = new google.maps.Map(document.getElementById("map_canvas"), myOptions);

    //Create the mappoints
    for (var i=0; i < mappoints.length; i++) {
       //setTimeout(function(iResult) {
       addMarker(i);
       //}, i * 500);
    }

    //Cycle through the sidebar array to build the html for the sidebar section
    var lcl_sidebar_html = '';

    for (var i=0; i < sidebar_links.length; i++) {
       lcl_sidebar_html += sidebar_links[i];
    }

    if(sidebar_links.length > 0) {

       lcl_sidebar_html = '<table id="sidebar_table">' + lcl_sidebar_html + '</table>';

       $('#sidebar_links').html(function() {
         return lcl_sidebar_html;
       });
    } else {
       $('#sidebar').css('display','none');
       $('#map_canvas').css('width','99%');
    }
  }

  function addMarker(iRowCount) {
    var lcl_pointcolor         = '';
    var lcl_markernum          = iRowCount + 1;
    var lcl_marker             = '';
    var lcl_marker_url         = 'http://gmaps-samples.googlecode.com/svn/trunk/markers/';
    var lcl_googleMapMarker    = '<%=sGoogleMapMarker%>';
    var lcl_marker_numberlimit = 99;

    lcl_pointcolor = mappointcolors[iRowCount];

    if(lcl_pointcolor == '') {
       lcl_pointcolor = 'green';
    }

    if(lcl_googleMapMarker == 'CUSTOMMARKER1') {
       lcl_marker_numberlimit = 600;
       lcl_marker_url         = 'mappoint_markers/custommarker1/';
    }

    if(lcl_markernum > lcl_marker_numberlimit) {
       lcl_marker = 'blank';
    } else {
       lcl_marker = 'marker' + lcl_markernum;
    }

    var image = lcl_marker_url + lcl_pointcolor + '/' + lcl_marker + '.png';

    var pinColor = "FE7569";
    if (lcl_pointcolor == "blue")
    {
	    pinColor = "839afa";
    }
    if (lcl_pointcolor == "green")
    {
	    pinColor = "92e415";
    }
    if (lcl_pointcolor == "pink")
    {
	    pinColor = "FFC0CB";
    }
    if (lcl_pointcolor == "orange")
    {
	    pinColor = "FFA500";
    }

    //console.log(iRowCount+1 + "|" + pinColor);

    var pinImage = new google.maps.MarkerImage("https://chart.apis.google.com/chart?chst=d_map_pin_letter&chld=" + (iRowCount+1) + "|" + pinColor,
        new google.maps.Size(21, 34),
        new google.maps.Point(0,0),
        new google.maps.Point(10, 34));

	markers.push( new google.maps.Marker({
                position: mappoints[iRowCount],
                map: map,
                icon: pinImage,
       		animation: google.maps.Animation.DROP
            }));
/*
    markers.push(new google.maps.Marker({
       position:  mappoints[iRowCount],
       map:       map,
       draggable: false,
       animation: google.maps.Animation.DROP,
       icon:      image
       //title:     "here" + iRowCount
    }));
    */

    openInfoWindow(iRowCount);

    return iRowCount;
  }

  function openInfoWindow(iRowCount) {
    infowindow.push(new google.maps.InfoWindow({
       size:     new google.maps.Size(50,50),
       position: mappoints[iRowCount],
       content:  bubbleInfo[iRowCount]
    }));

    google.maps.event.addListener(markers[iRowCount], 'click', function() {
       infowindow[iRowCount].open(map,markers[iRowCount]);
    });
  }

		//This function picks up the click and opens the corresponding info window
		function sidebarClick(i) {
//			 markers[i].show();

    google.maps.event.trigger(markers[i], 'click', function() {
       infowindow[i].open(map,markers[i]);
    });

		}
</script>
<%
 end if
'END: Google Maps javascript --------------------------------------------------
%>

<script type="text/javascript">
$(document).ready(function() {
  if($('#advanced_search').is(':visible')) {

     //Advanced Search: Click
     $('#advanced_search').click(function() {
       if($('#advanced_searchoptions').is(':visible')) {
          $('#advanced_searchoptions').hide('slow');
       } else {
          //$('#advanced_searchoptions').show('slow');
          displayAdvancedSearchOptions();
       }
     });

     //Advanced Search: Category - Change
     if($('#sc_dm_categoryid').is(':visible')) {
        $('#sc_dm_categoryid').change(function() {
          if($('#advanced_searchoptions').is(':visible')) {
             displayAdvancedSearchOptions();
          }
        });
     }
  }

  $('#addDMDataButton').click(function() {
    var lcl_url = '';

    lcl_url += 'datamgr_maint.asp';
    lcl_url += '?f=<%=lcl_feature%>';

    location.href = lcl_url;
  });

<% if lcl_enableOwnerMaint then %>
  $('#myDataMgrButton').click(function() {
    var lcl_url = '';

    lcl_url += 'mydatamgr.asp';
    lcl_url += '?f=<%=lcl_feature%>';

    location.href = lcl_url;
  });
<% end if %>
});

function openDMInfo(p_ID) {

  lcl_feature = '<%=lcl_feature%>';

  var lcl_dm_url;
  lcl_dm_url  = "datamgr_info.asp";
  lcl_dm_url += "?dm=" + p_ID;

  if(lcl_feature != "") {
     lcl_dm_url += "&f=" + lcl_feature;
  }

  location.href = lcl_dm_url;
}

var sorter = new TINY.table.sorter("sorter");

function listOrderInit() {
  sorter.head      = "head";
  sorter.asc       = "asc";
  sorter.desc      = "desc";
  sorter.even      = "evenrow";
  sorter.odd       = "oddrow";
  sorter.evensel   = "evenselected";
  sorter.oddsel    = "oddselected";
  //sorter.paginate  = true;
  //sorter.currentid = "currentpage";
  //sorter.limitid   = "pagelimit";
  sorter.init("mappoints",0);
}

function displayAdvancedSearchOptions() {
  if($('#advanced_searchoptions').not(':visible')) {
     $('#advanced_searchoptions').show('slow');
     $('#subcategorynames_display').html('');
  }

  var lcl_categoryid = $('#sc_dm_categoryid').val();

  //alert('build_subcategory_list.asp?userid=<%=iuserid%>&orgid=<%=iorgid%>&dm_typeid=<%=lcl_dmt%>&categoryid=' + lcl_categoryid + '&useraction=SEARCH&isAjax=Y');
  $.post('build_subcategory_list.asp', {
     userid:           '<%=iuserid%>',
     orgid:            '<%=iorgid%>',
     dm_typeid:        '<%=lcl_dmt%>',
     categoryid:       lcl_categoryid,
     subcategoryids:   '<%=lcl_sc_subcategory_ids%>',
     useraction:       'SEARCH',
     isAjax:           'Y'
  }, function(result) {
     $('#subCategoryList').html(result);
  });
}

function openLogin() {
  editURL = '../user_login.asp';

  location.href = editURL;
}
</script>

<style type="text/css">
  html { height: 100% }
  body { height: 100%; margin: 0px; padding: 0px; }
/*
  #test_text {
     border:   1pt solid #ff0000;
     height:   20px;
     overflow: hidden
  }
*/
</style>
<!--#include file="../include_top.asp"-->
</head>
<%
  response.write "<p>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""max-width:800px;"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font class=""pagetitle"">" & lcl_description & " Map</font>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"">" & vbcrlf
  response.write "          &nbsp;" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</p>" & vbcrlf

  RegisteredUserDisplay("../")

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"" class=""datamgr"">" & vbcrlf

 'BEGIN: Intro Message --------------------------------------------------------
  if lcl_enableOwnerMaint then
     response.write "<div align=""right"">" & vbcrlf
     response.write "  <input type=""button"" name=""addDMDataButton"" id=""addDMDataButton"" class=""button"" value=""Add " & lcl_description & """ />" & vbcrlf
     response.write "  <input type=""button"" name=""myDataMgrButton"" id=""myDataMgrButton"" class=""button"" value=""My "  & lcl_description & """ />" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  if lcl_intro_message <> "" then
     response.write "  <div id=""intro_message"">" & lcl_intro_message & "</div>" & vbcrlf
  end if

'  response.write "  <div id=""categories"">" & vbcrlf
'                      displayDMCategories iorgid, lcl_feature
'  response.write "  </div>" & vbcrlf
 'END: Intro Message ----------------------------------------------------------

 'BEGIN: Search Criteria ------------------------------------------------------
  response.write "  <div id=""search_criteria"">" & vbcrlf
  response.write "    <form name=""datamgr_searchoptions"" id=""datamgr_searchoptions"" method=""post"" action=""datamgr.asp"">" & vbcrlf
  response.write "      <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""20"" maxlength=""500"" />" & vbcrlf
  response.write "    <fieldset>" & vbcrlf
  response.write "      <legend>Search Options&nbsp;</legend>" & vbcrlf
  response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "        <tr valign=""top"">" & vbcrlf
  response.write "            <td>" & vbcrlf

  if lcl_line_count_search > 0 then
     response.write "                Search:" & vbcrlf
     response.write "                <input type=""text"" name=""sc_searchfield_0"" id=""sc_searchfield_0"" size=""30"" maxlength=""30"" value=""" & lcl_sc_searchfield & """ />" & vbcrlf
     response.write "                &nbsp;&nbsp;" & vbcrlf
  end if

'  response.write "                Category:" & vbcrlf
'  response.write "                <select name=""sc_dm_categoryid"" id=""sc_dm_categoryid"">" & vbcrlf
                                    lcl_categorytype = "PARENT"
                                    lcl_show_totals  = true

                                    displayDMCategoryOptions lcl_categorytype, _
                                                             iorgid, _
                                                             lcl_dm_typeid, _
                                                             lcl_sc_dm_categoryid, _
                                                             lcl_includeBlankCategoryOption, _
                                                             lcl_show_totals, _
                                                             lcl_displayMap
'  response.write "                </select>" & vbcrlf
'  response.write "                &nbsp;<span id=""advanced_search"">Advanced...</span>" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "        </tr>" & vbcrlf
  response.write "        <tr>" & vbcrlf
  response.write "            <td>" & vbcrlf
  response.write "                <input type=""hidden"" name=""subcategoryids"" id=""subcategoryids"" value=""" & lcl_sc_subcategory_ids & """ />" & vbcrlf
  response.write "                <input type=""hidden"" name=""subcategorynames"" id=""subcategorynames"" value=""" & lcl_sc_subcategory_names & """ />" & vbcrlf
  response.write "                Sub-Categories: <span id=""subcategorynames_display"">" & lcl_sc_subcategory_names & "</span>" & vbcrlf
  response.write "                <div id=""advanced_searchoptions"">" & vbcrlf
  response.write "                   <span id=""subCategoryList""></span>" & vbcrlf
  response.write "                </div>" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "        </tr>" & vbcrlf
  response.write "        <tr>" & vbcrlf
  response.write "            <td>" & vbcrlf
  response.write "                <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "        </tr>" & vbcrlf
  response.write "      </table>" & vbcrlf
  response.write "    </fieldset>" & vbcrlf
  response.write "    </form>" & vbcrlf
  response.write "  </div>" & vbcrlf
 'END: Search Criteria --------------------------------------------------------

 'BEGIN: Google Map and Side Bar ----------------------------------------------
  if lcl_displayMap then
     response.write "<div id=""map_canvas"">&nbsp;</div>" & vbcrlf
     response.write "<div id=""sidebar"">" & vbcrlf

     if lcl_description <> "" then
        response.write "  <span id=""sidebar_description"">" & lcl_description & "</span><br />" & vbcrlf
     end if

     response.write "  <div id=""sidebar_links""></div>" & vbcrlf
     response.write "</div>" & vbcrlf
  end if
 'END: Google Map and Side Bar ------------------------------------------------

 'BEGIN: Results List ---------------------------------------------------------
  response.write "  <div id=""list_results"">" & vbcrlf
  response.write "    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td><strong>" & lcl_description & "</strong></td>" & vbcrlf
  response.write "          <td align=""right"">&nbsp;</td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf
                      displayDMList lcl_feature, _
                                    lcl_query, _
                                    lcl_query_columnheaders, _
                                    lcl_displayMap
  response.write "  </div>" & vbcrlf
 'END: Results List -----------------------------------------------------------

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub createMapPoints(iFeature, iSQL)
 dim lcl_bgcolor

 'Cycle through all of the values and build:
 '  1. the MapPoints Array to display each map-point on the map.
 '  2. build each "info window (bubble)" for each map-point.
 '  3. build the sidebar section for each map-point.
  iRowCount            = 0
  iArrayCount          = 0
  lcl_previous_dmid    = 0
  lcl_list_mappoints   = ""
  lcl_list_mpcolors    = ""
  lcl_list_sidebar     = ""
  lcl_list_infowindows = ""
  lcl_infowindow_url   = ""
  lcl_query_mappoints  = iSQL
  lcl_bgcolor          = "#efefef"

  lcl_query_mappoints = lcl_query_mappoints & " ORDER BY dmd.dmid, dmtf.resultsOrder "
'dtb_debug(lcl_query_mappoints)
 	set oPoints = Server.CreateObject("ADODB.Recordset")
	on error resume next
 	oPoints.Open lcl_query_mappoints, Application("DSN"), 3, 1
	errnum = err.number
	on error goto 0

	if errnum <> 0 then
		oPoints.Open "SELECT top 1 userid FROM users WHERE 1=2", Application("DSN"), 3, 1
		response.write "<!--" & lcl_query_mappoints & "-->"
	end if

  if not oPoints.eof then
    	do while not oPoints.eof
        iRowCount   = iRowCount + 1
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#efefef","#ffffff")

       'If the previous DMID <> the current DMID then close out the loop and reset the row count
        if lcl_previous_dmid <> oPoints("dmid") then
           if iRowCount > 1 then
              iRowCount          = 1
              lcl_fieldvalue     = ""
              lcl_list_mappoints = lcl_list_mappoints & ", "  & vbcrlf
              lcl_list_mpcolors  = lcl_list_mpcolors  & ", "  & vbcrlf

              if oPoints("isSidebarLink") then
                 lcl_list_sidebar = lcl_list_sidebar & "', " & vbcrlf
              end if

              lcl_infowindow_url = "datamgr_info.asp?f=" & iFeature
              lcl_infowindow_url = lcl_infowindow_url & "&dm=" & lcl_previous_dmid
              'lcl_infowindow_url = "http://www.yahoo.com"

              lcl_list_infowindows = lcl_list_infowindows & "<a href=""" & lcl_infowindow_url & """>[more details...]</a>"
              lcl_list_infowindows = lcl_list_infowindows & "</div>', " & vbcrlf
           end if

           if iRowCount = 1 then
              lcl_list_infowindows = lcl_list_infowindows & "'<div>"

              if oPoints("isSidebarLink") then
                 if lcl_list_sidebar <> "" then
                    lcl_list_sidebar = lcl_list_sidebar & "'"
                 end if
              end if
           end if

          'Build the MapPoint and Sidebar Link
           if oPoints("displayInResults") then
              lcl_list_mappoints = lcl_list_mappoints & "new google.maps.LatLng(" & oPoints("latitude") & ", " & oPoints("longitude") & ")"
              lcl_list_mpcolors  = lcl_list_mpcolors  & "'" & oPoints("mappointcolor") & "'"

             'ONLY build the sidebar if the "isSidebarLink" is TRUE.
             'If "true" then pull the VALUE as the "link text".
             'The "iArrayCount" is basically the "RowCount", but instead of being reset when/if the DMTypeID changes
             '   if continues to increment each individual DM Value.  This count corresponds to the "mappoints" array
             '   used to access the Google Maps features.
             'NOTE: The query ONLY pulls columns/fields that are set to "DisplayInResults".
              'if oPoints("isSidebarLink") then
              '   lcl_sidebar_value = oPoints("fieldvalue")

              '   if lcl_sidebar_value <> "" then
              '      lcl_sidebar_value = replace(lcl_sidebar_value,"'","\'")
              '   else
              '      lcl_sidebar_value = "blah"
              '   end if

              '   lcl_list_sidebar = lcl_list_sidebar & "<tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>"
              '   lcl_list_sidebar = lcl_list_sidebar &     "<td nowrap=""nowrap"">"
              '   lcl_list_sidebar = lcl_list_sidebar &         "<img src=""mappoint_colors/bg_" & oPoints("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />&nbsp;" & (iArrayCount + 1) & "."
              '   lcl_list_sidebar = lcl_list_sidebar &     "</td>"
              '   lcl_list_sidebar = lcl_list_sidebar &     "<td width=""100%"">"
              '   lcl_list_sidebar = lcl_list_sidebar &         "<a href=""javascript:sidebarClick(" & iArrayCount & ")"">" & lcl_sidebar_value & "</a>"
              '   lcl_list_sidebar = lcl_list_sidebar &     "</td>"
              '   lcl_list_sidebar = lcl_list_sidebar & "</tr>"
              'end if
           end if

           iArrayCount = iArrayCount + 1

        end if

       'Only display on the map/results if the admin has set to display.
        if oPoints("displayInResults") then
           lcl_fieldvalue     = oPoints("fieldvalue")
           lcl_fieldname      = ""

          'Need to modify code since we are writing back into a javascript routine (array)
           if lcl_fieldvalue <> "" then

             'Check to see if we show the label or not.
              if oPoints("displayFieldName") then
                 lcl_fieldname  = oPoints("fieldname")

                 if lcl_fieldname <> "" then
                   'Need to modify code since we are writing back into a javascript routine (array)
                    lcl_fieldname = replace(lcl_fieldname,"'","\'")
                    lcl_fieldname = "<strong>" & lcl_fieldname & "</strong>: "

                    lcl_list_infowindows = lcl_list_infowindows & lcl_fieldname
                 end if
              end if

             'Build the "info window"
              lcl_fieldvalue = replace(lcl_fieldvalue,"'","\'")
              lcl_fieldvalue = replace(lcl_fieldvalue,chr(10),"")
              lcl_fieldvalue = replace(lcl_fieldvalue,chr(13),"<br />")

              if instr(oPoints("fieldtype"),"WEBSITE") > 0 OR instr(oPoints("fieldtype"),"EMAIL") > 0 then
                 lcl_fieldvalue = buildURLDisplayValue(lcl_fieldtype, lcl_fieldvalue)
              end if

              lcl_list_infowindows = lcl_list_infowindows & lcl_fieldvalue & "<br />"
           end if

           if oPoints("isSidebarLink") then

              lcl_sidebar_value = oPoints("fieldvalue")

             'ONLY build the sidebar if the "isSidebarLink" is TRUE.
             'If "true" then pull the VALUE as the "link text".
             'The "iArrayCount" is basically the "RowCount", but instead of being reset when/if the DMTypeID changes
             '   if continues to increment each individual DM Value.  This count corresponds to the "mappoints" array
             '   used to access the Google Maps features.
             'NOTE: (1) MUST be subtracted from the iArrayCount because the arraycount has already been incremented 
             '      by the time it gets to this part of the code.  (see line 848-ish)
             'NOTE: The query ONLY pulls columns/fields that are set to "DisplayInResults".
              if lcl_sidebar_value <> "" then
                 lcl_sidebar_value = replace(lcl_sidebar_value,"'","\'")
              else
                 lcl_sidebar_value = "blah"
              end if

              lcl_list_sidebar = lcl_list_sidebar & "<tr valign=""top"">"
              lcl_list_sidebar = lcl_list_sidebar &     "<td style=""background-color:" & lcl_bgcolor & "; white-space:nowrap;"">"
              lcl_list_sidebar = lcl_list_sidebar &         "<img src=""mappoint_colors/bg_" & oPoints("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />&nbsp;" & iArrayCount & "."
              lcl_list_sidebar = lcl_list_sidebar &     "</td>"
              lcl_list_sidebar = lcl_list_sidebar &     "<td style=""background-color:" & lcl_bgcolor & "; width:100%;"">"
              lcl_list_sidebar = lcl_list_sidebar &         "<a href=""javascript:sidebarClick(" & iArrayCount - 1 & ")"">" & lcl_sidebar_value & "</a>"
              lcl_list_sidebar = lcl_list_sidebar &     "</td>"
              lcl_list_sidebar = lcl_list_sidebar & "</tr>"
           end if
        end if

        lcl_previous_dmid = oPoints("dmid")
        oPoints.movenext
     loop

     if iRowCount > 0 then
        lcl_infowindow_url = "datamgr_info.asp?f=" & iFeature
        lcl_infowindow_url = lcl_infowindow_url & "&dm=" & lcl_previous_dmid
        'lcl_infowindow_url = "http://www.yahoo.com"

        lcl_list_infowindows = lcl_list_infowindows & "<a href=""" & lcl_infowindow_url & """>[more details...]</a>"
        lcl_list_infowindows = lcl_list_infowindows & "</div>'"

        if lcl_list_sidebar <> "" then
           lcl_list_sidebar = "'" & lcl_list_sidebar & "'"
        end if
     end if

  end if

  oPoints.close
  set oPoints = nothing

 'Create the MapPoints array
  response.write "var mappoints = [" & vbcrlf
  response.write lcl_list_mappoints & vbcrlf
  response.write "];" & vbcrlf

 'Create the MapPoints Colors array
  response.write "var mappointcolors = [" & vbcrlf
  response.write lcl_list_mpcolors & vbcrlf
  response.write "];" & vbcrlf

 'Create the Info Window array
  response.write "var bubbleInfo = [" & vbcrlf
  response.write lcl_list_infowindows & vbcrlf
  response.write "];" & vbcrlf

 'Build the Sidebar section
  response.write "var sidebar_links = [" & vbcrlf
  response.write lcl_list_sidebar & vbcrlf
  response.write "];"

end sub

'------------------------------------------------------------------------------
sub displayDMList(iFeature, iSQL_resultslist, iSQL_columnheaders, iDisplayMap)
 dim lcl_bgcolor

 'Cycle through all of the values and build the results list
  iRowCount               = 0
  iArrayCount             = 0
  lcl_previous_dmid       = 0
  lcl_previous_dm_typeid  = 0
  lcl_fieldtype           = ""
  lcl_fieldvalue          = ""
  lcl_section_fieldids    = ""
  lcl_query_resultslist   = iSQL_resultslist
  lcl_query_columnheaders = iSQL_columnheaders
  lcl_scripts             = ""
  lcl_bgcolor             = "#ffffff"

  response.write "<table id=""mappoints"" cellpadding=""2"" cellspacing=""0"" border=""0"" class=""mappoints_sortable liquidtable"">" & vbcrlf
  response.write "  <thead>" & vbcrlf

 'BEGIN: Build the row of column headers -----------------------------------
 	set oDMListHeaders = Server.CreateObject("ADODB.Recordset")
 	oDMListHeaders.Open lcl_query_columnheaders, Application("DSN"), 3, 1

  if not oDMListHeaders.eof then
     response.write "  <tr valign=""bottom"">" & vbcrlf

     if iDisplayMap then
        response.write "      <th nowrap=""nowrap""><span>Map #</span></th>" & vbcrlf
     end if
     response.write "      <th class=""nosort"">&nbsp;</th>" & vbcrlf

     response.write "      <th nowrap=""nowrap""><span>Category</span></th>" & vbcrlf


     do while not oDMListHeaders.eof
	     twfX = twfX + 1
	     redim preserve arrHeaders(twfX)
	     arrHeaders(twfX) = oDMListHeaders("fieldname")
	     response.write "<!--TWF " & twfX & " : " & arrHeaders(twfX) & "-->" & vbcrlf
        response.write "      <th><span>" & oDMListHeaders("fieldname") & "</span></th>" & vbcrlf

        if lcl_section_fieldids <> "" then
           lcl_section_fieldids = lcl_section_fieldids & "," & oDMListHeaders("section_fieldid")
        else
           lcl_section_fieldids = oDMListHeaders("section_fieldid")
        end if

        oDMListHeaders.movenext
     loop

     response.write "  </tr>" & vbcrlf

  end if

  oDMListHeaders.close
  set oDMListHeaders = nothing

  response.write "  </thead>" & vbcrlf
 'END: Build the row of column headers -------------------------------------

 'BEGIN: Build results list ---------------------------------------------------
  if lcl_section_fieldids <> "" then
     lcl_query_resultslist = lcl_query_resultslist & " AND dmtf.section_fieldid in (" & lcl_section_fieldids & ") "
  end if

  lcl_query_resultslist = lcl_query_resultslist & " ORDER BY dmd.dmid, dmtf.resultsOrder "

 	set oDMList = Server.CreateObject("ADODB.Recordset")
	on error resume next
 	oDMList.Open lcl_query_resultslist, Application("DSN"), 3, 1
	errnum = err.number
	on error goto 0

	if errnum <> 0 then
		oDMList.Open "SELECT top 1 userid FROM users WHERE 1=2", Application("DSN"), 3, 1
		response.write "<!--" & lcl_query_resultslist & "-->"
	end if

  if not oDMList.eof then
     lcl_scripts = lcl_scripts & "listOrderInit();"

     if iDisplayMap then
        lcl_marker_url         = "http://gmaps-samples.googlecode.com/svn/trunk/markers/"
        lcl_marker_numberlimit = 99

        if sGoogleMapMarker = "CUSTOMMARKER1" then
           lcl_marker_numberlimit = 600
           lcl_marker_url         = "mappoint_markers/custommarker1/"
        end if
     end if

	y = 2
     do while not oDMList.eof
	     y = y + 1

        iRowCount = iRowCount + 1
        lcl_return_url = "datamgr_info.asp"

       'If the previous DMID <> the current DMID then close out the loop and reset the row count
        if lcl_previous_dmid <> oDMList("dmid") then
		y = 3
           if iRowCount > 1 then
              iRowCount   = 1
              iArrayCount = iArrayCount + 1

              lcl_return_url = setupUrlParameters(lcl_return_url, "f",  iFeature)
              lcl_return_url = setupUrlParameters(lcl_return_url, "dm", lcl_previous_dmid)

              response.write "  </tr>" & vbcrlf
           end if

           if iRowCount = 1 then
              lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
              lcl_onclick = " onclick=""openDMInfo(" & oDMList("dmid") & ");"""

              response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
              'response.write "      <td align=""center"">" & (iArrayCount + 1) & ".</td>" & vbcrlf

              if iDisplayMap then
                 if iArrayCount + 1 > lcl_marker_numberlimit then
                    lcl_marker = "blank"
                 else
                    lcl_marker = "marker" & iArrayCount + 1
                 end if

                 'lcl_map_img = lcl_marker_url & oDMList("mappointcolor") & "/" & lcl_marker & ".png"

                 'response.write "      <td class=""repeatheaders"">Map #</td>" & vbcrlf
                 response.write "      <td align=""center"" onclick=""sidebarClick(" & iArrayCount & ")"">" & vbcrlf
                 'response.write "          <img src=""" & lcl_map_img & """ valign=""middle"" style=""cursor:pointer"" onclick=""sidebarClick(" & iArrayCount & ")"" />" & vbcrlf
                 response.write "          <div style=""cursor:pointer"" onclick=""sidebarClick(" & iArrayCount & ")"" ><span class=""repeatheaders"">Map #:</span>" & (iArrayCount+1) &  "</div>" & vbcrlf
		 'response.write  iRowCount
                 response.write "      </td>" & vbcrlf
              end if
           response.write "      <td align=""center"" onclick=""myclick(" & iRowCount-1 & ");"">" & vbcrlf
           response.write "          <span class=""repeatheaders"">Category Color:</span><img src=""mappoint_colors/bg_" & oDMList("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />" & vbcrlf
           response.write "      </td>" & vbcrlf

              response.write "      <td nowrap=""nowrap""" & lcl_onclick & ">" & vbcrlf
              'response.write "          <img src=""mappoint_colors/bg_" & oDMList("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />" & vbcrlf
              response.write "          <span class=""repeatheaders"">Category:</span>&nbsp;" & oDMList("categoryname") & vbcrlf
              response.write "      </td>" & vbcrlf
           end if
        end if

       'Display the fieldvalue(s) that have been selected to appear in the results list
        lcl_fieldtype  = oDMList("fieldtype")
        lcl_fieldvalue = oDMList("fieldvalue")

       'Need to modify code since we are writing back into a javascript routine (array)
        if lcl_fieldvalue <> "" then
           if instr(lcl_fieldtype,"WEBSITE") > 0 OR instr(lcl_fieldtype,"EMAIL") > 0 then
              lcl_fieldvalue = buildURLDisplayValue(lcl_fieldtype, lcl_fieldvalue)
           else
              lcl_fieldvalue = replace(lcl_fieldvalue,chr(10),"")
              lcl_fieldvalue = replace(lcl_fieldvalue,chr(13),"<br />")
           end if
        else
           lcl_fieldvalue = "&nbsp;"
        end if

	strHeader = ""
	if y <= UBOUND(arrHeaders) then 	strHeader = arrHeaders(y)
        response.write "      <td" & lcl_onclick & "><span class=""repeatheaders blockHeader"">" & strHeader & ":</span>" & lcl_fieldvalue & "</td>" & vbcrlf

        lcl_previous_dmid      = oDMList("dmid")
        lcl_previous_dm_typeid = oDMList("dm_typeid")

        oDMList.movenext
     loop

     response.write "  </tr>" & vbcrlf

  end if

  response.write "</table>" & vbcrlf

  oDMList.close
  set oDMList = nothing

  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if

 'END: Build results list -----------------------------------------------------

end sub

'------------------------------------------------------------------------------
sub displayDMCategoryOptions(iCategoryType, iOrgID, iDMTypeID, iSC_DMCategoryID, iIncludeBlankValue, iShowTotals, iDisplayMap)

  lcl_ctype               = ""   '"PARENT", "SUBCATEGORY", or "" <-- used to get both types of categories
  lcl_includeBlankValue   = 0
  sOrgID                  = 0
  sDMTypeID               = 0
  sSC_DMCategoryid        =    ""
  sShowTotals             = false
  sDisplayMap             = true
  lcl_subcategories_exist = 0
  lcl_scripts             = ""

  if iCategoryType <> "" then
     lcl_ctype = iCategoryType
     lcl_ctype = ucase(lcl_ctype)
  end if

  if iOrgID <> "" then
     sOrgID = iOrgID
     sOrgID = clng(sOrgID)
  end if

  if iDMTypeID <> "" then
     sDMTypeID = iDMTypeID
     sDMTypeID = clng(sDMTypeID)
  end if

  if iSC_DMCategoryID <> "" then
     sSC_DMCategoryID = iSC_DMCategoryID
     sSC_DMCategoryID = clng(sSC_DMCategoryID)
  end if

  if iShowTotals <> "" then
     sShowTotals = iShowTotals
  end if

  if iDisplayMap <> "" then
     sDisplayMap = iDisplayMap
  end if

  if iIncludeBlankValue <> "" then

    'Just a validation check before assigning the value passed in 
    'to ensure that the value passed in is a true/false value
     if iIncludeBlankValue OR not iIncludeBlankValue then
        lcl_includeBlankValue = iIncludeBlankValue
     end if
  end if

  sSQL = "SELECT c.categoryid, "
  sSQL = sSQL & " c.categoryname, "
  sSQL = sSQL & " c.mappointcolor, "
  sSQL = sSQL & " (select count(sc.categoryid) "
  sSQL = sSQL &  " from egov_dm_categories sc "
  sSQL = sSQL &       " inner join egov_dmdata_to_dmcategories dtc "
  sSQL = sSQL &                  " on sc.categoryid = dtc.categoryid "
  sSQL = sSQL &                  " and dtc.dm_typeid = " & sDMTypeID
  sSQL = sSQL &  " and sc.parent_categoryid = c.categoryid "
  sSQL = sSQL &  " and sc.isActive = 1 "
  sSQL = sSQL &  " and sc.isApproved = 1 "
  sSQL = sSQL &  " and sc.orgid = " & sOrgID
  sSQL = sSQL &  " and sc.dm_typeid = " & sDMTypeID
  sSQL = sSQL &  ") as total_subcategories "
  sSQL = sSQL & " FROM egov_dm_categories c "
  sSQL = sSQL & " WHERE c.isActive = 1 "
  sSQL = sSQL & " AND c.isApproved = 1 "
  sSQL = sSQL & " AND c.orgid = "     & sOrgID
  sSQL = sSQL & " AND c.dm_typeid = " & sDMTypeID

  if lcl_ctype = "PARENT" then
     sSQL = sSQL & " AND c.parent_categoryid = 0 "
  end if

  sSQL = sSQL & " ORDER BY upper(c.categoryname) "

 	set oDMCategoryOptions = Server.CreateObject("ADODB.Recordset")
 	oDMCategoryOptions.Open sSQL, Application("DSN"), 3, 1

  if not oDMCategoryOptions.eof then

     response.write "Category:" & vbcrlf
     response.write "<select name=""sc_dm_categoryid"" id=""sc_dm_categoryid"">" & vbcrlf

     if lcl_includeBlankValue then
        response.write "  <option value=""0"">[All Categories]</option>" & vbcrlf
     end if

     do while not oDMCategoryOptions.eof

        lcl_selected_category   = ""
        lcl_option_text         = ""
        lcl_total_assignments   = 0
        lcl_subcategories_exist = lcl_subcategories_exist + oDMCategoryOptions("total_subcategories")

        if oDMCategoryOptions("categoryid") = sSC_DMCategoryID then
           lcl_selected_category = " selected=""selected"""
        end if

        lcl_option_text = oDMCategoryOptions("categoryname")

       'Count total assignments if showing totals in option
        if sShowTotals then
           lcl_total_assignments = getTotalCategoryAssignments(sOrgID, sDMTypeID, oDMCategoryOptions("categoryid"), sDisplayMap)

           if lcl_total_assignments > 0 then
              lcl_option_text = lcl_option_text & "&nbsp;(" & lcl_total_assignments & ")"
           end if
        end if

        response.write "  <option value=""" & oDMCategoryOptions("categoryid") & """" & lcl_selected_category & ">" & lcl_option_text & "</option>" & vbcrlf

        oDMCategoryOptions.movenext
     loop

     response.write "</select>" & vbcrlf
'     response.write "<input type=""text"" name=""total_subcategories"" id=""total_subcategories"" value=""" & lcl_subcategories_exist & """ size=""3"" maxlength=""10"" />" & vbcrlf

     if lcl_subcategories_exist > 0 then
        response.write "&nbsp;<span id=""advanced_search"">Advanced...</span>" & vbcrlf
     end if

  end if

  oDMCategoryOptions.close
  set oDMCategoryOptions = nothing

end sub

'------------------------------------------------------------------------------
function getTotalCategoryAssignments(p_orgid, p_dm_typeid, p_categoryid, p_displaymap)
  lcl_return = 0

  sSQL = " SELECT COUNT(categoryid) as total_assignments "
  sSQL = sSQL & " FROM egov_dm_data "
  sSQL = sSQL & " WHERE orgid = " & p_orgid
  sSQL = sSQL & " AND dm_typeid = " & p_dm_typeid
  sSQL = sSQL & " AND categoryid IN (" & p_categoryid & ") "
  sSQL = sSQL & " AND isActive = 1 "
  sSQL = sSQL & " AND isApproved = 1 "

  'if p_displaymap then
  '   sSQL = sSQL & " AND latitude IS NOT NULL "
  '   sSQL = sSQL & " AND latitude <> 0.00 "
  '   sSQL = sSQL & " AND longitude IS NOT NULL "
  '   sSQL = sSQL & " AND longitude <> 0.00 "
  'end if

 	set oGetTotalCatAssignments = Server.CreateObject("ADODB.Recordset")
 	oGetTotalCatAssignments.Open sSQL, Application("DSN"), 3, 1

  if not oGetTotalCatAssignments.eof then
     lcl_return = oGetTotalCatAssignments("total_assignments")
  end if

  oGetTotalCatAssignments.close
  set oGetTotalCatAssignments = nothing

  getTotalCategoryAssignments = lcl_return

end function
%>
