<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mydatamgr.asp
' AUTHOR: David Boyer
' CREATED: 08/29/2011
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  user can view any/all data manager records that they are an "Owner/Editor" of.
'
' MODIFICATION HISTORY
' 1.0	 08/29/2011	David Boyer	- Initial Version
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

 set oDataMgr = New classOrganization

 'lcl_permission_check = "datamgr"
 lcl_feature        = ""
 lcl_dmt            = ""
 lcl_url_parameters = ""

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
          lcl_feature = request("f")
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

 if lcl_feature <> "datamgr_maint" then
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
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

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

'Get the userid if available
 lcl_cookie_userid       = ""
 session("RedirectPage") = request.servervariables("script_name") & "?" & request.querystring()

	if request.cookies("userid") <> "" and request.cookies("userid") <> "-1" then
		  lcl_cookie_userid = request.cookies("userid")
 else
    response.redirect "../user_login.asp"
 end if

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
 lcl_query_columnheaders = lcl_query_columnheaders &      " INNER JOIN egov_dm_sections AS dms "
 lcl_query_columnheaders = lcl_query_columnheaders &                   " ON dms.sectionid = dmts.sectionid "
 lcl_query_columnheaders = lcl_query_columnheaders & " WHERE dmd.orgid = " & iOrgID
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.dm_typeid = " & lcl_dm_typeid
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmt.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmts.isActive = 1 "
' lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmsf.isActive = 1 "
 'lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.isApproved = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dmtf.displayInResults = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dms.isActive = 1 "
 lcl_query_columnheaders = lcl_query_columnheaders & " AND dms.isAccountInfoSection = 1 "

 if lcl_displayMap then
    lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.latitude IS NOT NULL "
    lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.latitude <> 0.00 "
    lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.longitude IS NOT NULL "
    lcl_query_columnheaders = lcl_query_columnheaders & " AND dmd.longitude <> 0.00 "
 end if

 lcl_query_columnheaders = lcl_query_columnheaders & " ORDER BY dmtf.resultsOrder "
'END: Build the query to be used to get the column headers --------------------

'BEGIN: Build the query to be used within the mapping functions ---------------
 lcl_query_resultslist = "SELECT dmd.dmid, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.dm_typeid, "
 lcl_query_resultslist = lcl_query_resultslist & "dmtf.dm_fieldid, "
 lcl_query_resultslist = lcl_query_resultslist & "dmsf.fieldname, "
 lcl_query_resultslist = lcl_query_resultslist & "dmsf.fieldtype, "
 lcl_query_resultslist = lcl_query_resultslist & "dmv.fieldvalue, "
 lcl_query_resultslist = lcl_query_resultslist & "dmv.dm_valueid, "
 lcl_query_resultslist = lcl_query_resultslist & "dmtf.displayFieldName, "
 lcl_query_resultslist = lcl_query_resultslist & "dmtf.displayInResults, "
 lcl_query_resultslist = lcl_query_resultslist & "dmtf.isSidebarLink, "
 lcl_query_resultslist = lcl_query_resultslist & "dmtf.resultsOrder, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.streetnumber, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.streetprefix, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.streetaddress, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.streetsuffix, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.streetdirection, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.latitude, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.longitude, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.isCreatedByAdmin, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.createdbyid, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.createdbydate, "
 lcl_query_resultslist = lcl_query_resultslist & "ISNULL(ISNULL(dmd.mappointcolor, dmc.mappointcolor), 'green') AS mappointcolor, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.isApproved as isApproved_dmdata, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.approvedeniedbyid as approvedeniedbyid_dmdata, "
 lcl_query_resultslist = lcl_query_resultslist & "'By Admin' AS approvedeniedbyname_dmdata_admin, "
 lcl_query_resultslist = lcl_query_resultslist & "dmd.approvedeniedbydate as approvedeniedbydate_dmdata, "
 lcl_query_resultslist = lcl_query_resultslist & "dmc.categoryname, "
 lcl_query_resultslist = lcl_query_resultslist & "dmt.displayMap, "
 lcl_query_resultslist = lcl_query_resultslist & "dmt.useAdvancedSearch, "
 lcl_query_resultslist = lcl_query_resultslist & "dmo.ownertype, "
 lcl_query_resultslist = lcl_query_resultslist & "dmo.isApprovedDeniedByAdmin as isApprovedDeniedByAdmin_owner, "
 lcl_query_resultslist = lcl_query_resultslist & "dmo.isApproved as isApproved_owner, "
 lcl_query_resultslist = lcl_query_resultslist & "dmo.approvedeniedbydate as approvedeniedbydate_owner, "
 lcl_query_resultslist = lcl_query_resultslist & "u2.userfname + ' ' + u2.userlname AS approvedeniedbyname_owner_citizen, "
 lcl_query_resultslist = lcl_query_resultslist & "'By Admin' AS approvedeniedbyname_owner_admin "
 lcl_query_resultslist = lcl_query_resultslist & " FROM egov_dm_data AS dmd "
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_types AS dmt "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmt.dm_typeid = dmd.dm_typeid "
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_types_fields AS dmtf "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmt.dm_typeid = dmtf.dm_typeid "
 lcl_query_resultslist = lcl_query_resultslist &      " LEFT OUTER JOIN egov_dm_values AS dmv "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmv.dm_fieldid = dmtf.dm_fieldid "
 lcl_query_resultslist = lcl_query_resultslist &                 " AND dmv.dmid = dmd.dmid "
 lcl_query_resultslist = lcl_query_resultslist &                 " AND dmv.dm_typeid = dmd.dm_typeid "
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_types_sections dmts "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmts.dm_sectionid = dmtf.dm_sectionid "
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_sections_fields AS dmsf "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmsf.section_fieldid = dmtf.section_fieldid "
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_categories AS dmc "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmc.categoryid = dmd.categoryid "
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_owners AS dmo "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dmo.dmid = dmd.dmid "
 lcl_query_resultslist = lcl_query_resultslist &                 " AND dmo.orgid = " & iOrgID
 lcl_query_resultslist = lcl_query_resultslist &                 " AND dmo.userid = " & lcl_cookie_userid
 lcl_query_resultslist = lcl_query_resultslist &      " INNER JOIN egov_dm_sections AS dms "
 lcl_query_resultslist = lcl_query_resultslist &                 " ON dms.sectionid = dmts.sectionid "
' lcl_query_resultslist = lcl_query_resultslist &                 " AND (dmo.isApproved = 1 "
' lcl_query_resultslist = lcl_query_resultslist &                 "  OR  dmo.isApproved = 0 AND (dmo.approvedeniedbydate = '' or dmo.approvedeniedbydate is null)) "
 lcl_query_resultslist = lcl_query_resultslist &      " LEFT OUTER JOIN egov_users u2 ON u2.userid = dmo.approvedeniedbyid "
 lcl_query_resultslist = lcl_query_resultslist & " WHERE dmd.orgid = " & iOrgID
 lcl_query_resultslist = lcl_query_resultslist & " AND dmd.dm_typeid = " & lcl_dm_typeid
 lcl_query_resultslist = lcl_query_resultslist & " AND dmt.isActive = 1 "
 lcl_query_resultslist = lcl_query_resultslist & " AND dmts.isActive = 1 "
' lcl_query_resultslist = lcl_query_resultslist & " AND dmd.isActive = 1 "
 lcl_query_resultslist = lcl_query_resultslist & " AND dmsf.isActive = 1 "
' lcl_query_resultslist = lcl_query_resultslist & " AND dmd.isApproved = 1 "
 lcl_query_resultslist = lcl_query_resultslist & " AND dmtf.displayInResults = 1 "
 lcl_query_resultslist = lcl_query_resultslist & " AND dms.isActive = 1 "
 lcl_query_resultslist = lcl_query_resultslist & " AND dms.isAccountInfoSection = 1 "

' if lcl_displayMap then
'    lcl_query_resultslist = lcl_query_resultslist & " AND dmd.latitude IS NOT NULL "
'    lcl_query_resultslist = lcl_query_resultslist & " AND dmd.latitude <> 0.00 "
'    lcl_query_resultslist = lcl_query_resultslist & " AND dmd.longitude IS NOT NULL "
'    lcl_query_resultslist = lcl_query_resultslist & " AND dmd.longitude <> 0.00 "
' end if
'END: Build the query to be used within the mapping functions -----------------
%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	 <title>E-Gov Services <%=sOrgName%></title>

	 <link rel="stylesheet" type="text/css" href="../css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
	 <link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

	 <script type="text/javascript" src="../scripts/modules.js"></script>
	 <script type="text/javascript" src="../scripts/easyform.js"></script>  
 	<script type="text/javascript" src="../scripts/column_sorting.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

<script type="text/javascript">
<!--
$(document).ready(function(){
  $('#returnButton').click(function(){
    location.href='datamgr.asp<%=lcl_url_parameters%>';
  });
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

function openWin2(url, name) {
			popupWin = window.open(url, name,"resizable,width=500,height=450");
}
	//-->
	</script>
</head>

<!--#include file="../include_top.asp"-->
<%
  response.write "<font class=""pagetitle"">My " & lcl_description & "</font><br />" & vbcrlf

  RegisteredUserDisplay( sLevel )

  response.write "<div id=""content"">" & vbcrlf
  response.write " 	<div id=""centercontent"">" & vbcrlf

  response.write "<div style=""margin-bottom:5px;"">" & vbcrlf
  response.write "  <input type=""button"" name=""returnButton"" id=""returnButton"" class=""button"" value=""Return to " & lcl_description & """ />" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""mappoints_sortable liquidtable"" style=""max-width:800px"">" & vbcrlf
  response.write "  <thead>" & vbcrlf

 'BEGIN: Build the row of column headers -----------------------------------
 	set oDMListHeaders = Server.CreateObject("ADODB.Recordset")
 	oDMListHeaders.Open lcl_query_columnheaders, Application("DSN"), 3, 1

dim arrHeaders()
x = 2
  if not oDMListHeaders.eof then

     response.write "  <tr valign=""bottom"">" & vbcrlf
     'response.write "      <th nowrap=""nowrap""><span>Map #</span></th>" & vbcrlf
     response.write "      <th nowrap=""nowrap""><span>Category</span></th>" & vbcrlf

     do while not oDMListHeaders.eof
	     x = x + 1
	     redim preserve arrHeaders(x)
	     arrHeaders(x) = oDMListHeaders("fieldname")
        response.write "      <th><span>" & oDMListHeaders("fieldname") & "</span></th>" & vbcrlf

        if lcl_section_fieldids <> "" then
           lcl_section_fieldids = lcl_section_fieldids & "," & oDMListHeaders("section_fieldid")
        else
           lcl_section_fieldids = oDMListHeaders("section_fieldid")
        end if

        oDMListHeaders.movenext
     loop

     response.write "      <th nowrap=""nowrap""><span>Status</span></th>" & vbcrlf
     response.write "      <th nowrap=""nowrap""><span>Owner Status</span></th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"">&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

  end if

  oDMListHeaders.close
  set oDMListHeaders = nothing

  response.write "  </thead>" & vbcrlf
 'END: Build the row of column headers ----------------------------------------

 'BEGIN: Build results list ---------------------------------------------------
  lcl_previous_dmid             = 0
  lcl_previous_isCreatedByAdmin = false
  lcl_previous_createdbyid      = 0
  lcl_previous_createdbydate    = ""

  lcl_previous_isApprovedDeniedByAdmin_owner      = false
  lcl_previous_approvedeniedbyname_owner_admin    = ""
  lcl_previous_approvedeniedbyname_owner_citizen  = ""
  lcl_previous_approvedeniedbydate_owner          = ""
  lcl_previous_isApproved_owner                   = false
  lcl_previous_ownertype_dmdata                   = ""

  lcl_previous_isApprovedDeniedByAdmin_dmdata     = ""
  lcl_previous_approvedeniedbyname_dmdata_admin   = ""
  lcl_previous_approvedeniedbyname_dmdata_citizen = ""
  lcl_previous_approvedeniedbydata_dmdata         = ""
  lcl_previous_isApproved_dmdata                  = false
  lcl_previous_ownertype_owner                    = ""

  lcl_show_editbutton                             = 0
  iRowCount                                       = 0
  iColumnCount                                    = 1

  if lcl_section_fieldids <> "" then
     lcl_query_resultslist = lcl_query_resultslist & " AND dmtf.section_fieldid in (" & lcl_section_fieldids & ") "
  end if

  lcl_query_resultslist = lcl_query_resultslist & " ORDER BY dmd.dmid, dmtf.resultsOrder "

 	set oDMList = Server.CreateObject("ADODB.Recordset")
 	oDMList.Open lcl_query_resultslist, Application("DSN"), 3, 1

y = 2
  if not oDMList.eof then
     do while not oDMList.eof
	     y = y + 1

        lcl_return_url = "datamgr_info.asp"
        lcl_return_url = setupUrlParameters(lcl_return_url, "f",  iFeature)

       'If the previous DMID <> the current DMID then:
        if lcl_previous_dmid <> oDMList("dmid") then
		y = 3
           iRowCount = iRowCount + 1

           if iRowCount > 0 AND iRowCount <> 1 then
              iColumnCount   = 1
              iArrayCount    = iArrayCount + 1
              lcl_return_url = setupUrlParameters(lcl_return_url, "dm", lcl_previous_dmid)

              'response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approved_denied_status_dmdata & "<br />" & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
              'response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approved_denied_status_owner & "<br />" & lcl_approvedenied_info_owner & "</td>" & vbcrlf
 '             response.write "      <td nowrap=""nowrap"" align=""center"">[" & lcl_cookie_userid & "] - [" & lcl_previous_createdbyid & "] - [" & oDMList("dmid") & "] - [" & lcl_show_editbutton & "] " & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
 '             response.write "      <br />[" & lcl_cookie_userid & "] - [" & lcl_previous_createdbyid & "] - [" & oDMList("dmid") & "] - [" & lcl_show_editbutton & "] " & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
              response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
              response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approvedenied_info_owner  & "</td>" & vbcrlf
              response.write "      <td nowrap=""nowrap"" align=""center"">" & vbcrlf

              'if lcl_isApprovedDeniedByAdmin_owner then
              if lcl_show_editbutton > 0 then
                 response.write "          <input type=""button"" name=""editDMDataButton" & iArrayCount + 1 & """ id=""editDMDataButton" & iArrayCount + 1 & """ value=""Edit"" class=""button"" onclick=""openDMInfo('" & lcl_previous_dmid & "');"" />" & vbcrlf
              end if

              response.write "      </td>" & vbcrlf
              response.write "  </tr>" & vbcrlf

              lcl_previous_isCreatedByAdmin                   = false
              lcl_previous_createdbyid                        = 0
              lcl_previous_createdbydate                      = ""

              lcl_previous_isApprovedDeniedByAdmin_dmdata     = ""
              lcl_previous_approvedeniedbyname_dmdata_admin   = ""
              lcl_previous_approvedeniedbyname_dmdata_citizen = ""
              lcl_previous_approvedeniedbydata_dmdata         = ""
              lcl_previous_isApproved_dmdata                  = false

              lcl_previous_isApprovedDeniedByAdmin_owner      = false
              lcl_previous_approvedeniedbyname_owner_admin    = ""
              lcl_previous_approvedeniedbyname_owner_citizen  = ""
              lcl_previous_approvedeniedbydate_owner          = ""
              lcl_previous_isApproved_owner                   = false
              lcl_previous_ownertype_owner                    = ""

              lcl_show_editbutton                             = 0
           end if

          'If this is the first column then we want to show the category
           if iColumnCount = 1 then
              lcl_bgcolor              = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
              lcl_display_categoryname = "&nbsp;"

              if oDMList("categoryname") <> "" then
                 lcl_display_categoryname = oDMList("categoryname")
              end if

              response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
              response.write "      <td class=""repeatheaders"">Category</td>" & vbcrlf
              response.write "      <td nowrap=""nowrap"">" & vbcrlf
'              response.write "          iRowCount: [" & iRowCount & "] - iColumnCount: [" & iColumnCount & "] - [" & oDMList("dmid") & " - " & lcl_previous_dmid & "]&nbsp;" & oDMList("categoryname") & vbcrlf
'              response.write "          iRowCount: [" & iRowCount & "] - iColumnCount: [" & iColumnCount & "] " & oDMList("categoryname") & vbcrlf
              response.write "          " & lcl_display_categoryname & vbcrlf
              response.write "      </td>" & vbcrlf
          '--------------------------------------------------------------------
'           else  'If not then we want to close out the loop and reset the row count
          '--------------------------------------------------------------------
'              iColumnCount   = 0

           end if
        else
           iColumnCount = iColumnCount + 1
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

        'response.write "      <td>iRowCount: [" & iRowCount & "] - iColumnCount: [" & iColumnCount & "] " & lcl_fieldvalue & "</td>" & vbcrlf
        response.write "      <td class=""repeatheaders"">" & arrHeaders(y) & "</td>" & vbcrlf
        response.write "      <td>" & lcl_fieldvalue & "</td>" & vbcrlf

       'We need to set these variables at the end of the loop because as the loop cycles we check in the
       'beginning of the loop for the end of the row.  If we are at the end of the row we cannot go back and
       'grab the previous values in the loop as we've cycled and are on the newest row.  Therefore, we set
       'these variables to refer to.
        lcl_previous_dmid                               = oDMList("dmid")
        lcl_previous_dm_typeid                          = oDMList("dm_typeid")

        lcl_previous_isCreatedByAdmin                   = oDMList("isCreatedByAdmin")
        lcl_previous_createdbyid                        = oDMList("createdbyid")
        lcl_previous_createdbydate                      = oDMList("createdbydate")

        lcl_previous_approvedeniedbyname_dmdata_admin   = oDMList("approvedeniedbyname_dmdata_admin")
        lcl_previous_approvedeniedbydate_dmdata         = oDMList("approvedeniedbydate_dmdata")
        lcl_previous_isApproved_dmdata                  = oDMList("isApproved_dmdata")

        lcl_previous_isApprovedDeniedByAdmin_owner      = oDMList("isApprovedDeniedByAdmin_owner")
        lcl_previous_approvedeniedbyname_owner_admin    = oDMList("approvedeniedbyname_owner_admin")
        lcl_previous_approvedeniedbyname_owner_citizen  = oDMList("approvedeniedbyname_owner_citizen")
        lcl_previous_approvedeniedbydate_owner          = oDMList("approvedeniedbydate_owner")
        lcl_previous_isApproved_owner                   = oDMList("isApproved_owner")
        lcl_previous_ownertype_owner                    = oDMList("ownertype")

       'Set up Approve/Deny Information
        lcl_approvedenied_info_dmdata = setupApprovedDeniedInfo(lcl_previous_isApprovedDeniedByAdmin_dmdata, _
                                                                lcl_previous_approvedeniedbyname_dmdata_admin, _
                                                                lcl_previous_approvedeniedbyname_dmdata_citizen, _
                                                                lcl_previous_approvedeniedbydate_dmdata, _
                                                                lcl_previous_isApproved_dmdata, _
                                                                lcl_previous_ownertype_dmdata)

        lcl_approvedenied_info_owner = setupApprovedDeniedInfo(lcl_previous_isApprovedDeniedByAdmin_owner, _
                                                               lcl_previous_approvedeniedbyname_owner_admin, _
                                                               lcl_previous_approvedeniedbyname_owner_citizen, _
                                                               lcl_previous_approvedeniedbydate_owner, _
                                                               lcl_previous_isApproved_owner, _
                                                               lcl_previous_ownertype_owner)

       'Determine if the "edit" button is displayed
        lcl_show_editbutton = showHideEditButton(lcl_previous_isApproved_dmdata, _
                                                 lcl_previous_isApprovedDeniedByAdmin_dmdata, _
                                                 lcl_previous_approvedeniedbydate_dmdata, _
                                                 lcl_previous_isApproved_owner, _
                                                 lcl_previous_isApprovedDeniedByAdmin_owner, _
                                                 lcl_previous_approvedeniedbydate_dmdata, _
                                                 lcl_previous_isCreatedByAdmin, _
                                                 lcl_previous_createdbyid, _
                                                 lcl_cookie_userid)

        oDMList.movenext
     loop

     if iRowCount > 0 then
        lcl_return_url = setupUrlParameters(lcl_return_url, "dm", lcl_previous_dmid)

        'response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approved_denied_status_dmdata & "<br />" & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
        'response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approved_denied_status_owner & "<br />" & lcl_approvedenied_info_owner & "</td>" & vbcrlf
        'response.write "      <td nowrap=""nowrap"" align=""center"">[" & iColumnCount + 1 & "] - [bottom - " & lcl_show_editbutton & "] " & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
              response.write "      <td class=""repeatheaders"">Status</td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approvedenied_info_dmdata & "</td>" & vbcrlf
              response.write "      <td class=""repeatheaders"">Owner Status</td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"" align=""center"">" & lcl_approvedenied_info_owner  & "</td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"" align=""center"">" & vbcrlf

'dtb_debug("[" & lcl_previous_isApproved_dmdata & "] - " & _
'          "[" & lcl_previous_isApprovedDeniedByAdmin_dmdata & "] - " & _
'          "[" & lcl_previous_approvedeniedbydate_dmdata & "] - " & _
'          "[" & lcl_previous_isApproved_owner & "] - " & _
'          "[" & lcl_previous_isApprovedDeniedByAdmin_owner & "] - " & _
'          "[" & lcl_previous_approvedeniedbydate_dmdata & "] - " & _
'          "[" & lcl_previous_isCreatedByAdmin & "] - " & _
'          "[" & lcl_previous_createdbyid & "] - " & _
'          "[" & lcl_cookie_userid & "]")

        if lcl_show_editbutton > 0 then
           response.write "          <input type=""button"" name=""editDMDataButton" & iArrayCount + 1 & """ id=""editDMDataButton" & iArrayCount + 1 & """ value=""Edit"" class=""button"" onclick=""openDMInfo('" & lcl_previous_dmid & "');"" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if
  end if

  response.write "</table>" & vbcrlf

  oDMList.close
  set oDMList = nothing
 'END: Build results list -----------------------------------------------------

  response.write " 	</div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#include file="../include_bottom.asp"-->
<%
'------------------------------------------------------------------------------
function setupApprovedDeniedInfo(iIsApprovedDeniedByAdmin, iAdminName, iCitizenName, _
                                 iApproveDeniedDate, iIsApproved, iOwnerType)

  dim lcl_return, lcl_display_info, lcl_display_statusinfo

  lcl_return             = ""
  lcl_display_info       = ""
  lcl_display_statusinfo = ""

 'Determine if this owner/editor has been approved by a citizen user or an admin user
  if iIsApprovedDeniedByAdmin <> "" then
     if iIsApprovedDeniedByAdmin then
        lcl_display_info = formatAdminActionsInfo(iAdminName, iApproveDeniedDate)
     else
        lcl_display_info = formatAdminActionsInfo(iCitizenName, iApproveDeniedDate)
     end if
  else
     lcl_display_info = formatAdminActionsInfo(iAdminName, iApproveDeniedDate)
  end if

  'if lcl_display_info <> "" AND iApproveDeniedDate <> "" then
  if iApproveDeniedDate <> "" then
     if iIsApproved then
        lcl_display_statusinfo = "APPROVED"
     else
        lcl_display_statusinfo = "DENIED"
     end if

     if iOwnerType <> "" then
        lcl_display_statusinfo = iOwnerType & " - " & lcl_display_statusinfo
     end if

  else
     if iOwnerType <> "" then
        lcl_display_statusinfo = iOwnerType & "<br />"
     end if

     lcl_display_statusinfo = lcl_display_statusinfo & "WAITING FOR<br />APPROVAL"
  end if

  if lcl_display_statusinfo <> "" then
     lcl_display_statusinfo = "<span class=""redText"">" & lcl_display_statusinfo & "</span>"
  end if

  lcl_return = lcl_display_statusinfo

  if iApproveDeniedDate <> "" then
     lcl_return = lcl_return & "<br />" & lcl_display_info
  end if

  setupApprovedDeniedInfo = lcl_return

end function

'------------------------------------------------------------------------------
function showHideEditButton(iIsApproved_dmdata, iIsApprovedDeniedByAdmin_dmdata, iApprovedDeniedByDate_dmdata, _
                            iIsApproved_owner,  iIsApprovedDeniedByAdmin_owner,  iApprovedDeniedByDate_owner, _
                            iIsCreatedByAdmin, iCreatedById, iUserID)

  dim lcl_return, lcl_showbutton, sIsCreatedByAdmin, sCreatedById, sUserID
  dim sIsApproved_dmdata, sIsApprovedDeniedByDadmin_dmdata, sApprovedDeniedByDate_dmdata
  dim sIsApproved_owner,  sIsApprovedDeniedByDadmin_owner,  sApprovedDeniedByDate_owner

  lcl_return     = 0
  lcl_showbutton = 0

  sIsApproved_dmdata              = 0
  sIsApprovedDeniedByadmin_dmdata = 0
  sApprovedDeniedByDate_dmdata    = ""
  sIsApproved_owner               = 0
  sIsApprovedDeniedByadmin_owner  = 0
  sApprovedDeniedByDate_owner     = ""
  sIsCreatedByAdmin               = 0
  sCreatedByID                    = 0
  sUserID                         = 0

  if iIsApproved_dmdata <> "" then
     sIsApproved_dmdata = iIsApproved_dmdata
  end if

  if iIsApprovedDeniedByAdmin_dmdata <> "" then
     sIsApprovedDeniedByAdmin_dmdata = iIsApprovedDeniedByAdmin_dmdata
  end if

  if iApprovedDeniedByDate_dmdata <> "" then
     sApprovedDeniedByDate_dmdata = iApprovedDeniedByDate_dmdata
  end if

  if iIsApproved_owner <> "" then
     sIsApproved_owner = iIsApproved_owner
  end if

  if iIsApprovedDeniedByAdmin_owner <> "" then
     sIsApprovedDeniedByAdmin_owner = iIsApprovedDeniedByAdmin_owner
  end if

  if iApprovedDeniedByDate_owner <> "" then
     sApprovedDeniedByDate_owner = iApprovedDeniedByDate_owner
  end if

  if iIsCreatedByAdmin <> "" then
     sIsCreatedByAdmin = iIsCreatedByAdmin
  end if

  if iCreatedById <> "" then
     sCreatedByID = iCreatedById
  end if

  if iUserID <> "" then
     sUserID = iUserID
  end if

 'Work though all conditions to determine if the "edit" button is displayed or not
 'Increment the "return" value each time a non-denied condition passes.
 'SHOW BUTTON:
 ' 1. DM Data Status = Waiting AND Owner Status = Waiting  AND User =  DM Data Creator
 ' 2. DM Data Status = Waiting AND Owner Status = Approved AND User =  DM Data Creator
 ' 3. DM Data Status = Waiting AND Owner Status = Approved AND User <> DM Data Creator
 ' 4. DM Data Status = Approved AND Owner Status = Waiting  AND  User = DM Data Creator
 ' 5. DM Data Status = Approved AND Owner Status = Approved AND (User = DM Data Creator OR User <> DM Data Creator)

dtb_debug("sIsApproved_dmdata: [" & sIsApproved_dmdata & "] - sApprovedDeniedByDate_owner: [" & sApprovedDeniedByDate_owner & "] - sUserID: [" & sUserID & "] - sCreatedByID: [" & sCreatedByID & "] - sIsApproved_owner: [" & sIsApproved_owner & "] - sApprovedDeniedByDate_dmdata: [" & sApprovedDeniedByDate_dmdata & "]")

  if sIsApproved_dmdata then
     if sApprovedDeniedByDate_owner = "" then
        if sUserID = sCreatedByID then
           lcl_showbutton = lcl_showbutton + 1
        end if
     else
        if sIsApproved_owner then
           lcl_showbutton = lcl_showbutton + 1
        end if
     end if
  else
     if sApprovedDeniedByDate_dmdata = "" then
        if sIsApproved_owner then
           lcl_showbutton = lcl_showbutton + 1

           'if sUserID = sCreatedByID then
           '   lcl_showbutton = lcl_showbutton + 1
           'else
           '   lcl_showbutton = lcl_showbutton + 1
           'end if
        else
           if sApprovedDeniedByDate_owner = "" then
              if sUserID = sCreatedByID then
                 lcl_showbutton = lcl_showbutton + 1
              end if
           end if
        end if
     end if
  end if

dtb_debug("lcl_showbutton: [" & lcl_showbutton & "]")

  showHideEditButton = lcl_showbutton

end function
%>
