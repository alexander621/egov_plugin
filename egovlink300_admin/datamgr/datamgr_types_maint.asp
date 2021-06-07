<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_types_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DM Type
'
' MODIFICATION HISTORY
' 1.0 03/05/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'sSQL = "SELECT orgid, latitude, longitude, defaultzoomlevel "
'sSQL = sSQL & " FROM organizations "
'sSQL = sSQL & " WHERE defaultzoomlevel is not null or defaultzoomlevel <> '' "

'set rs1 = Server.CreateObject("ADODB.Recordset")
'rs1.Open sSQL, Application("DSN"), 3, 1

'if not rs1.eof then
'   do while not rs1.eof

'      sSQL = "UPDATE egov_dm_types SET "
'      sSQL = sSQL & " latitude = "  & rs1("latitude")  & ", "
'      sSQL = sSQL & " longitude = " & rs1("longitude") & ", "
'      sSQL = sSQL & " defaultzoomlevel = '" & rs1("defaultzoomlevel") & "' "
'      sSQL = sSQL & " WHERE orgid = " & rs1("orgid")

'      set rs2 = Server.CreateObject("ADODB.Recordset")
'      rs2.Open sSQL, Application("DSN"), 3, 1

'      rs1.movenext
'   loop

'   rs2.close
'   set rs2 = nothing

'end if

'rs1.close
'set rs1 = nothing

'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel               = "../"  'Override of value from common.asp
 lcl_onload           = ""
 lcl_isRootAdmin      = False
 lcl_isTemplate       = False
 lcl_isTemplate_url   = ""
 lcl_isTemplate_title = ""

'Determine if there is a specific feature associated to the DM Types
'  "Is Limited": means that the admin user is maintaining a specific feature instead of the root admin viewing ALL DM Types
'      (i.e. "Is Limited" will be true when an admin clicks on "Maintain Available Properties (fields)", but NOT 
'            when a root admin clicks on "Maintain DM Types" and then selects a specific DM Type to edit.
 lcl_dm_typeid    = ""
 lcl_feature      = "datamgr_types_maint"
 lcl_isLimited    = False
 lcl_pagetitle    = "DM Types"
 lcl_sectiontitle = lcl_pagetitle & ": Fields"

 if request("f") <> "" AND request("f") <> "datamgr_types_maint" then
    lcl_feature      = request("f")
    lcl_isLimited    = True
    lcl_pagetitle    = getFeatureName(lcl_feature)
    lcl_sectiontitle = lcl_pagetitle

   'Retrieve the DM_TypeID
    if request("dm_typeid") <> "" then
       lcl_dm_typeid = request("dm_typeid")
    else
       lcl_dm_typeid = getDMTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

       if lcl_dm_typeid = 0 then
         	response.redirect sLevel & "permissiondenied.asp"
       end if
    end if
 else
    if request("dm_typeid") <> "" then
       lcl_dm_typeid = request("dm_typeid")
    end if
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Determine if this is a template
 if request("t") = "Y" then
    lcl_isTemplate       = True
    lcl_isTemplate_url   = "&t=Y"
    lcl_isTemplate_title = " Template"
 end if

'Retrieve the dm_typeid to be maintain.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 'if request("dm_typeid") <> "" then
 if lcl_dm_typeid <> "" then
    'lcl_dm_typeid = request("dm_typeid")

    if isnumeric(lcl_dm_typeid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "datamgr_types_list.asp"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_sendToLabel = "Create"
    lcl_dm_typeid   = 0
 end if

'Set up local variables
 lcl_description                  = ""
 lcl_isActive                     = 1
 lcl_createdbyid                  = 0
 lcl_createdbydate                = ""
 lcl_createdbyname                = ""
 lcl_lastmodifiedbyid             = 0
 lcl_lastmodifiedbydate           = ""
 lcl_lastmodifiedbyname           = ""
 lcl_mappointcolor                = "green"
 lcl_feature_public               = ""
 lcl_feature_maintain             = ""
 lcl_feature_maintain_fields      = ""
 lcl_feature_owners               = ""
 lcl_assignedto                   = 0
 lcl_displayMap                   = 1
 lcl_latitude                     = ""
 lcl_longitude                    = ""
 sLat                             = ""
 sLng                             = ""
 lcl_defaultZoomLevel             = ""
 lcl_googleMapType                = "ROADMAP"
 lcl_layoutid                     = 0
 lcl_accountInfoSectionID         = 0
 lcl_defaultcategoryid            = 0
 lcl_includeBlankCategoryOption   = 1
 lcl_intro_message                = ""
 lcl_checked_isActive             = " checked=""checked"""
 lcl_checked_enableOwnerMaint     = ""
 lcl_checked_displayMap           = " checked=""checked"""
 lcl_checked_useAdvancedSearch    = ""
 lcl_checked_includeBlankCategory = " checked=""checked"""

'Get the Latitude and Longitude for the org
 GetCityPoint session("orgid"), sLat, sLng

 if lcl_screen_mode = "EDIT" then

   'Retrieve all of the data for the DM Type
    sSQL = "SELECT t.dm_typeid, "
    sSQL = sSQL & " t.description, "
    sSQL = sSQL & " t.isActive, "
    sSQL = sSQL & " t.enableOwnerMaint, "
    sSQL = sSQL & " t.createdbyid, "
    sSQL = sSQL & " t.createdbydate, "
    sSQL = sSQL & " t.lastmodifiedbyid, "
    sSQL = sSQL & " t.lastmodifiedbydate, "
    sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
    sSQL = sSQL & " t.mappointcolor, "
    sSQL = sSQL & " t.feature_public, "
    sSQL = sSQL & " t.feature_maintain, "
    sSQL = sSQL & " t.feature_maintain_fields, "
    sSQL = sSQL & " t.feature_owners, "
    sSQL = sSQL & " t.assignedto, "
    sSQL = sSQL & " t.displayMap, "
    sSQL = sSQL & " t.useAdvancedSearch, "
    sSQL = sSQL & " t.latitude, "
    sSQL = sSQL & " t.longitude, "
    sSQL = sSQL & " t.defaultzoomlevel, "
    sSQL = sSQL & " t.googleMapType, "
    sSQL = sSQL & " isnull(t.googleMapMarker, 'GOOGLE') as googleMapMarker, "
    sSQL = ssQL & " t.layoutid, "
    sSQL = sSQL & " t.accountInfoSectionID, "
    sSQL = sSQL & " t.defaultcategoryid, "
    sSQL = sSQL & " t.includeBlankCategoryOption, "
    sSQL = sSQL & " t.intro_message "
    sSQL = sSQL & " FROM egov_dm_types t "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON t.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON t.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE t.dm_typeid = " & lcl_dm_typeid

    set oDMTypes = Server.CreateObject("ADODB.Recordset")
    oDMTypes.Open sSQL, Application("DSN"), 3, 1

    if not oDMTypes.eof then
       lcl_description                = oDMTypes("description")
       lcl_isActive                   = oDMTypes("isActive")
       lcl_enableOwnerMaint           = oDMTypes("enableOwnerMaint")
       lcl_createdbyid                = oDMTypes("createdbyid")
       lcl_createdbydate              = oDMTypes("createdbydate")
       lcl_createdbyname              = oDMTypes("createdbyname")
       lcl_lastmodifiedbyid           = oDMTypes("lastmodifiedbyid")
       lcl_lastmodifiedbydate         = oDMTypes("lastmodifiedbydate")
       lcl_lastmodifiedbyname         = oDMTypes("lastmodifiedbyname")
       lcl_mappointcolor              = oDMTypes("mappointcolor")
       lcl_feature_public             = oDMTypes("feature_public")
       lcl_feature_maintain           = oDMTypes("feature_maintain")
       lcl_feature_maintain_fields    = oDMTypes("feature_maintain_fields")
       lcl_feature_owners             = oDMTypes("feature_owners")
       lcl_assignedto                 = oDMTypes("assignedto")
       lcl_displayMap                 = oDMTypes("displayMap")
       lcl_useAdvancedSearch          = oDMTypes("useAdvancedSearch")
       lcl_latitude                   = oDMTypes("latitude")
       lcl_longitude                  = oDMTypes("longitude")
       lcl_defaultZoomLevel           = oDMTypes("defaultzoomlevel")
       lcl_googleMapType              = oDMTypes("googleMapType")
       lcl_googleMapMarker            = oDMTypes("googleMapMarker")
       lcl_layoutid                   = oDMTypes("layoutid")
       lcl_accountInfoSectionID       = oDMTypes("accountInfoSectionID")
       lcl_defaultcategoryid          = oDMTypes("defaultcategoryid")
       lcl_includeBlankCategoryOption = oDMTypes("includeBlankCategoryOption")
       lcl_intro_message              = oDMTypes("intro_message")

      'If the DM Type is NOT "active" then do NOT "check" the checkbox
       if not lcl_isActive then
          lcl_checked_isActive = ""
       end if

       if lcl_enableOwnerMaint then
          lcl_checked_enableOwnerMaint = " checked=""checked"""
       end if

       if not lcl_displayMap then
          lcl_checked_displayMap = ""
       end if

       if lcl_useAdvancedSearch then
          lcl_checked_useAdvancedSearch = " checked=""checked"""
       end if

       if not lcl_includeBlankCategoryOption then
          lcl_checked_includeBlankCategory = ""
       end if

       if lcl_googleMapType = "" OR isnull(lcl_googleMapType) then
          lcl_googleMapType = "ROADMAP"
       else
          lcl_googleMapType = ucase(lcl_googleMapType)
       end if

      'Determine which Google Map Type is selected.
       lcl_selected_googleMapType_roadmap   = ""
       lcl_selected_googleMapType_satellite = ""
       lcl_selected_googleMapType_hybrid    = ""
       lcl_selected_googleMapType_terrain   = ""

       if lcl_googleMapType = "SATELLITE" then
          lcl_selected_googleMapType_satellite = " selected=""selected"""
       elseif lcl_googleMapType = "HYBRID" then
          lcl_selected_googleMapType_hybrid    = " selected=""selected"""
       elseif lcl_googleMapType = "TERRAIN" then
          lcl_selected_googleMapType_terrain   = " selected=""selected"""
       else
          lcl_selected_googleMapType_roadmap   = " selected=""selected"""
       end if

      'Determine which Google Map Marker is selected.
       lcl_selected_googleMapMarker_google        = ""
       lcl_selected_googleMapMarker_custommarker1 = ""

       if lcl_googleMapMarker = "CUSTOMMARKER1" then
          lcl_selected_googleMapMarker_custommarker1 = " selected=""selected"""
       else
          lcl_selected_googleMapMarker_google        = " selected=""selected"""
       end if

    else
       response.redirect("datamgr_types_list.asp?success=NE")
    end if

    oDMTypes.close
    set oDMTypes = nothing

 else
    lcl_latitude  = sLat
    lcl_longitude = sLng

    if lcl_latitude = "" then
       lcl_latitude = 0.00
    end if

    if lcl_longitude = "" then
       lcl_longitude = 0.00
    end if

   'Zoom Levels are from 0 to 21+ (with 0 meaning "max zoomed OUT" or "entire world view")
   'Depending on the area in the map determines what the max zoom IN can be.
   'One area may only be able to zoom in to "14" will others may allow you to zoom in to "20" or more.
   'We default to "13"
    if lcl_defaultZoomLevel = "" OR isnull(lcl_defaultZoomLevel) then
 		   	lcl_defaultZoomLevel = "13"
    end if
 end if

'Check for org features
' lcl_orghasfeature_rssfeeds_mayorsblog = orghasfeature("rssfeeds_mayorsblog")

'Check for user permissions
' lcl_userhaspermission_rssfeeds_mayorsblog = userhaspermission(session("userid"),"rssfeeds_mayorsblog")

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

'Check for associated DM Date to determine if this DM Type can/cannot be deleted.
'*** NOTE: if this IS a template then allow then bypass the check.
 if lcl_isTemplate then
    lcl_canDelete = true
 else
    lcl_canDelete = False
 end if

'Determine if the Layout exists and if not then get the "Original" Layout
 getLayoutInfo lcl_layoutid, _
               lcl_layoutname, _
               lcl_isOriginalLayout, _
               lcl_useLayoutSections, _
               lcl_totalcolumns, _
               lcl_columnwidth_left, _
               lcl_columnwidth_middle, _
               lcl_columnwidth_right

 if lcl_isRootAdmin AND not lcl_isTemplate then
    'lcl_canDelete = True
    lcl_canDelete = checkForDMDataByDMTypeID(lcl_dm_typeid)

    lcl_onload = lcl_onload & "displayFeature('feature_public','displayfeature_public');"
    lcl_onload = lcl_onload & "displayFeature('feature_maintain','displayfeature_maintain');"
    lcl_onload = lcl_onload & "displayFeature('feature_maintain_fields','displayfeature_maintain_fields');"
    lcl_onload = lcl_onload & "displayFeature('feature_owners','displayfeature_owners');"
 end if

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"
 lcl_onload  = lcl_onload & "enableDisableMapSetupFields();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_pagetitle%> - <%=lcl_screen_mode%><%=lcl_isTemplate_title%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css" rel="stylesheet" type="text/css"/>

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script language="javascript" src="../scripts/datamgr_fields_addrow.js"></script>

<% '  <script type="text/javascript" src="../scripts/jquery-1.4.4.min.js"></script> %>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>
  <script type="text/javascript" src="../scripts/jquery-ui-1.8.4.custom.min.js"></script>

<style type="text/css">
  .hidden            { display: none; }
  .requiredFieldsMsg { color: #800000; }

  .helpOption     { cursor: pointer }
  .helpOptionText {
     background-color:      #a80000;
     font-size:             12px;
     color:                 #ffffff;
     padding:               5px 5px;
     margin:                5px 5px;
     border:                1pt solid #000000;
     -webkit-border-radius: 5px;
     -moz-border-radius:    5px;
  }

  .textarea_intromessage {
     width:  550px;
     height: 100px;
  }

  .fieldset {
     border: 1pt solid #808080;
     -webkit-border-radius: 5px;
     -moz-border-radius:    5px;
  }

  .placeHolderHighlight {
     background-color: red;
     height:      1.5em;
     line-height: 1.2em;
  }

  .dragDropArrows {
     cursor: pointer;
  }
</style>

<script type="text/javascript">
$(document).ready(function(){
//Return a helper with preserved width of cells
var fixHelper = function(e, ui) {
  	ui.children().each(function() {
		    $(this).width($(this).width());
	  });

	  return ui;
};

  $('#addFieldTBL tbody').sortable({
     connectWith: '#addFieldTBL tbody',
     helper:      fixHelper,
     revert:      true,
     cursor:      'move',
     stop: function(event, ui) {

        $el = $(ui.item);
        $el.find('tr').click();
        $el.effect('highlight',{},2000);

        //update the Display Order
	      	$('#addFieldTBL tbody').each(function(){
          var itemorder       = $(this).sortable('toArray');
          var lcl_total_items = itemorder.length;
          var lcl_rowID       = '';

          for(var i = 0; i < lcl_total_items; i++) {
             lcl_rowID = itemorder[i];
             lcl_rowID = lcl_rowID.replace('addFieldRow','');

             $('#resultsOrder' + lcl_rowID).val(i+1);
          }
        });
     }
  })
  .disableSelection();

  //Miscellaneous Setup
  $('#requiredFieldsMsg').addClass('requiredFieldsMsg');

  if($('#displayMap').prop('checked')) {
     $('#sidebarlink_columnheader').html('Sidebar<br />Link (key)');
  } else {
     $('#sidebarlink_columnheader').html('Record<br />Key');
  }

  //-- "Display Map on Public Page" Checkbox: onClick -------------------------
  $('#displayMap').click(function() {
    //if($('#displayMap').attr('checked')) {
    if($('#displayMap').prop('checked')) {
       $('#requiredFieldsMsg').toggleClass('requiredFieldsMsg');
       $('#requiredFieldsMsg').show('slow');

       $('#sidebarlink_columnheader').html('Sidebar<br />Link (key)');

       //$('#mappointcolor').attr('disabled',false);
       //$('#latitude').attr('disabled',false);
       //$('#longitude').attr('disabled',false);
       //$('#getLatLongButton').attr('disabled',false);
       //$('#defaultzoomlevel').attr('disabled',false);

       $('#latitude').prop('disabled',false);
       $('#longitude').prop('disabled',false);
       $('#getLatLongButton').prop('disabled',false);
       $('#defaultzoomlevel').prop('disabled',false);
    } else {
       $('#requiredFieldsMsg').toggleClass('requiredFieldsMsg');
       $('#requiredFieldsMsg').hide('slow');

       $('#sidebarlink_columnheader').html('Record<br />Key');

       //$('#mappointcolor').attr('disabled',true);
       //$('#latitude').attr('disabled',true);
       //$('#longitude').attr('disabled',true);
       //$('#getLatLongButton').attr('disabled',true);
       //$('#defaultzoomlevel').attr('disabled',true);

       $('#latitude').prop('disabled',true);
       $('#longitude').prop('disabled',true);
       $('#getLatLongButton').prop('disabled',true);
       $('#defaultzoomlevel').prop('disabled',true);
    }
  });

  //Cycle through all fields and see if one has been "checked" to be the "isSidebarLink".
  //If "yes" then disable all of the other "isSidebarLink" checkboxes.
  function countChecked(iID) {
    var lcl_total_elements = $(iID).length;
    var lcl_total_checked  = 0;

    $(iID).each(function() {
      //if ($(this).attr('checked')) {
      if ($(this).prop('checked')) {
          lcl_total_checked += 1;
      }
    });

    return lcl_total_checked;
  }

  function enableDisableCheckboxes(iFieldID) {
    var lcl_total = countChecked(iFieldID);

    if(lcl_total > 0) {
       $(iFieldID).each(function() {
         //if (! $(this).attr('checked')) {
         //    $(this).attr('disabled','disabled');
         //}
         if (! $(this).prop('checked')) {
             $(this).prop('disabled','disabled');
         }
       });
    } else {
       $(iFieldID).each(function() {
          //$(this).attr('disabled','');
          $(this).prop('disabled','');
       });
    }
  }

  //-- Sidebar Link Checkboxes: onClick ---------------------------------------
  $('.isSidebarLinkCheckbox').click(function() {
    enableDisableCheckboxes('.isSidebarLinkCheckbox');
  });

  //-- This call sets up the checkboxes when the page is loaded/refreshed -----
  enableDisableCheckboxes('.isSidebarLinkCheckbox');
<%
  if lcl_isRootAdmin AND not lcl_isTemplate AND lcl_screen_mode = "ADD" then
    '-- Template: onChange ----------------------------------------------------
     response.write "  $('#DMTemplateID').change(function() {" & vbcrlf
     response.write "    enableDisableTemplateLayoutFields();" & vbcrlf
     response.write "  });" & vbcrlf

    '-- Account Info Section: onChange ----------------------------------------
     response.write "  $('#accountInfoSectionID').change(function() {" & vbcrlf
     response.write "    enableDisableTemplateLayoutFields();" & vbcrlf
     response.write "  });" & vbcrlf

    '-- Layout: onChange ------------------------------------------------------
     response.write "  $('#layoutid').change(function() {" & vbcrlf
     response.write "    enableDisableTemplateLayoutFields();" & vbcrlf
     response.write "  });" & vbcrlf
  end if
%>

  //Set up "help" definitions
  $('#helpFeature_public_text').hide();
  $('#helpFeature_maintain_text').hide();
  $('#helpFeature_maintain_fields_text').hide();
  $('#helpFeature_owners_text').hide();
  $('#helpFeature_defaultCategory_text').hide();

  $('#helpFeature_public').click(function() {
     $('#helpFeature_public_text').toggle('slow');
  });

  $('#helpFeature_maintain').click(function() {
     $('#helpFeature_maintain_text').toggle('slow');
  });

  $('#helpFeature_maintain_fields').click(function() {
     $('#helpFeature_maintain_fields_text').toggle('slow');
  });

  $('#helpFeature_owners').click(function() {
     $('#helpFeature_owners_text').toggle('slow');
  });

  $('#helpFeature_defaultCategory').click(function() {
     $('#helpFeature_defaultCategory_text').toggle('slow');
  });

  //Enable/Disable "owner" fields
  $('#enableOwnerMaint').attr('disabled',true);
  $('#assignedto').attr('disabled',true);

  if($('#feature_owners').val() != '') {
     $('#enableOwnerMaint').attr('disabled',false);
     $('#assignedto').attr('disabled',false);
  }

  $('#feature_owners').change(function() {
    if($('#feature_owners').val() != '') {
       $('#enableOwnerMaint').attr('disabled',false);
       $('#assignedto').attr('disabled',false);
    } else {
       $('#enableOwnerMaint').val('');
       $('#assignedto').val('');

       $('#enableOwnerMaint').attr('disabled',true);
       $('#assignedto').attr('disabled',true);
    }
  });

  $('#googleMapMarker').change(function() {
     showGoogleMarker();
  });

  showGoogleMarker();
});

<%
  if lcl_isRootAdmin AND not lcl_isTemplate AND lcl_screen_mode = "ADD" then
     lcl_onload = "enableDisableTemplateLayoutFields();"

    'BEGIN: Enable/Disable Template/Layout Fields -----------------------------
     response.write "  function enableDisableTemplateLayoutFields() {" & vbcrlf
     'response.write "    $('#DMTemplateID').attr('disabled',false);" & vbcrlf
     'response.write "    $('#accountInfoSectionID').attr('disabled',false);" & vbcrlf
     'response.write "    $('#layoutid').attr('disabled',false);" & vbcrlf

     response.write "    $('#DMTemplateID').prop('disabled',false);" & vbcrlf
     response.write "    $('#accountInfoSectionID').prop('disabled',false);" & vbcrlf
     response.write "    $('#layoutid').prop('disabled',false);" & vbcrlf

     'response.write "    if($('#DMTemplateID').val() != '' || $('#accountInfoSectionID').attr('selectedIndex') > 0 || $('#layoutid').attr('selectedIndex') > 0) {" & vbcrlf
     response.write "    if($('#DMTemplateID').val() != '' || $('#accountInfoSectionID').prop('selectedIndex') > 0 || $('#layoutid').prop('selectedIndex') > 0) {" & vbcrlf
     response.write "       if($('#DMTemplateID').val() != '') {" & vbcrlf
     'response.write "          $('#accountInfoSectionID').attr('disabled',true);" & vbcrlf
     'response.write "          $('#layoutid').attr('disabled',true);" & vbcrlf
     response.write "          $('#accountInfoSectionID').prop('disabled',true);" & vbcrlf
     response.write "          $('#layoutid').prop('disabled',true);" & vbcrlf
     response.write "       } else {" & vbcrlf
     'response.write "          if($('#accountInfoSectionID').attr('selectedIndex') > 0 || $('#layoutid').attr('selectedIndex') > 0) {" & vbcrlf
     'response.write "             $('#DMTemplateID').attr('disabled',true);" & vbcrlf
     'response.write "          }" & vbcrlf
     response.write "          if($('#accountInfoSectionID').prop('selectedIndex') > 0 || $('#layoutid').prop('selectedIndex') > 0) {" & vbcrlf
     response.write "             $('#DMTemplateID').prop('disabled',true);" & vbcrlf
     response.write "          }" & vbcrlf
     response.write "       }" & vbcrlf
     response.write "    }" & vbcrlf
     response.write "  }" & vbcrlf
    'END: Enable/Disable Template/Layout Fields -------------------------------
  end if
%>

function showGoogleMarker() {
   $('#googleMapMarkerImg').html();
   var lcl_googleMapMarker = $('#googleMapMarker option:selected').val();
   var lcl_img_filename    = '';
   var lcl_img_display     = '';

   if(lcl_googleMapMarker == 'CUSTOMMARKER1') {
      lcl_img_filename = '<%=Application("MAP_URL")%>datamgr/mappoint_markers/custommarker1/red/marker99.png';
   } else {
      lcl_img_filename = 'http://gmaps-samples.googlecode.com/svn/trunk/markers/green/marker99.png';
   }

   lcl_img_display = '&nbsp;&nbsp;&nbsp;<img src="' + lcl_img_filename + '" align="top" />';

   $('#googleMapMarkerImg').html(lcl_img_display);
}

function maintainLayout() {
  w = 1000;
  h = 700;
  l = (screen.availWidth/2)-(w/2);
  t = (screen.availHeight/2)-(h/2);

  pickerURL  = "datamgr_types_layout_maint.asp";
  pickerURL += "?dm_typeid=" + document.getElementById("dm_typeid").value;
  pickerURL += "&layoutid="  + document.getElementById("layoutid").value;

  if('<%=lcl_feature%>' != 'datamgr_types_maint') {
     pickerURL += "&f=<%=lcl_feature%>";
  }

  eval('window.open("' + pickerURL + '", "_mpt_layout", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');  
//  window.open(pickerURL, "_mpt_layout", "width=" + w + ",height=" + h + ",left=" + l + ",top=" + t + ",toolbar=0,statusbar=0,scrollbars=1,menubar=0");
}

function maintainCategories() {
  w = 1000;
  h = 700;
  l = (screen.availWidth/2)-(w/2);
  t = (screen.availHeight/2)-(h/2);

  pickerURL  = "datamgr_categories_list.asp";
  pickerURL += "?dm_typeid="  + document.getElementById("dm_typeid").value;

  if('<%=lcl_feature%>' != 'datamgr_types_maint') {
     pickerURL += "&f=<%=lcl_feature%>";
  }

  eval('window.open("' + pickerURL + '", "_mpt_layout", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');  
//  window.open(pickerURL, "_mpt_layout", "width=" + w + ",height=" + h + ",left=" + l + ",top=" + t + ",toolbar=0,statusbar=0,scrollbars=1,menubar=0");
}

var control_field = "";

function enableDisableMapSetupFields() {
  lcl_isFieldChecked = false;

  if(document.getElementById("displayMap")) {
     lcl_isFieldChecked = document.getElementById("displayMap").checked;
  }

  document.getElementById("requiredFieldsMsg").style.display     = "none";
  //document.getElementById("mappointcolor").disabled              = true;
  document.getElementById("latitude").disabled                   = true;
  document.getElementById("longitude").disabled                  = true;
  document.getElementById("defaultzoomlevel").disabled = true;

  if(lcl_isFieldChecked) {
     document.getElementById("requiredFieldsMsg").style.display     = "inline";
     //document.getElementById("mappointcolor").disabled              = false;
     document.getElementById("latitude").disabled                   = false;
     document.getElementById("longitude").disabled                  = false;
     document.getElementById("defaultzoomlevel").disabled = false;

     if(document.getElementById("defaultzoomlevel").value == "") {
        document.getElementById("defaultzoomlevel").value = "13";
     }
  }
}

function getOrgLatLong() {
  var lcl_replace_latlng = true;

  if(document.getElementById("latitude").value != "" || document.getElementById("longitude").value != "") {
     var latlng = confirm('Any values entered into the Latitude or Longitude will be overridden.\nAre you sure you want to continue?');
     if(!latlng) {
        lcl_replace_latlng = false;
     }
  }

  if(lcl_replace_latlng) {
     document.getElementById("latitude").value  = "<%=sLat%>";
     document.getElementById("longitude").value = "<%=sLng%>";
  }
}

function confirmDelete() {
  //var r = confirm('Are you sure you want to delete the "' + document.getElementById("title").value + '" blog entry?  \r NOTE: Any/All comments will be deleted as well.');
  var r = confirm('Are you sure you want to delete: "' + document.getElementById("description").value + '"');
  if (r==true) {
      location.href="datamgr_types_action.asp?user_action=DELETE&dm_typeid=<%=lcl_dm_typeid & lcl_isTemplate_url%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count    = 0;
  var lcl_isFieldChecked = false;

  //---------------------------------------------------------------------------
  //Check the DM Types Fields
  //---------------------------------------------------------------------------
<% if lcl_screen_mode <> "ADD" then %>
  if(document.getElementById("displayMap")) {
     lcl_isFieldChecked = document.getElementById("displayMap").checked;
  }

  if(lcl_isFieldChecked) {
     if(document.getElementById("defaultzoomlevel").value=="") {
        //inlineMsg(document.getElementById("defaultzoomlevel").id,'<strong>Required Field Missing: </strong> Zoom Level',10,'defaultzoomlevel');
        //lcl_focus       = document.getElementById("defaultzoomlevel");
        //lcl_false_count = lcl_false_count + 1;
        document.getElementById("defaultzoomlevel").value = "13";
     }else{
    				var rege = /^\d+$/;
				    var Ok = rege.exec(document.getElementById("defaultzoomlevel").value);

    				if ( ! Ok ) {
           inlineMsg(document.getElementById("defaultzoomlevel").id,'<strong>Invalid Value: </strong> Zoom Level must be numeric.',10,'defaultzoomlevel');
           lcl_focus       = document.getElementById("defaultzoomlevel");
           lcl_false_count = lcl_false_count + 1;
    			} else {
           clearMsg("defaultzoomlevel");
       }
     }

     if(document.getElementById("longitude").value=="") {
        inlineMsg(document.getElementById("longitude").id,'<strong>Required Field Missing: </strong> Longitude',10,'longitude');
        lcl_focus       = document.getElementById("longitude");
        lcl_false_count = lcl_false_count + 1;
     }else{
    				//var rege = /^\d+$/;
        var rege = /^-?\d+(\.\d+)?$/;
				    var Ok = rege.exec(document.getElementById("longitude").value);

    				if ( ! Ok ) {
           inlineMsg(document.getElementById("longitude").id,'<strong>Invalid Value: </strong> Longitude must be numeric.<br /><span style="color:#800000;">(i.e. 30.44111 or -85.4744111)</span>',10,'longitude');
           lcl_focus       = document.getElementById("longitude");
           lcl_false_count = lcl_false_count + 1;
    			} else {
           clearMsg("longitude");
       }
     }

     if(document.getElementById("latitude").value=="") {
        inlineMsg(document.getElementById("latitude").id,'<strong>Required Field Missing: </strong> Latitude',10,'latitude');
        lcl_focus       = document.getElementById("latitude");
        lcl_false_count = lcl_false_count + 1;
     }else{
    				//var rege = /^\d+$/;
        var rege = /^-?\d+(\.\d+)?$/;
				    var Ok = rege.exec(document.getElementById("latitude").value);

    				if ( ! Ok ) {
           inlineMsg(document.getElementById("latitude").id,'<strong>Invalid Value: </strong> Latitude must be numeric.<br /><span style="color:#800000;">(i.e. 30.44111 or -85.4744111)</span>',10,'latitude');
           lcl_focus       = document.getElementById("latitude");
           lcl_false_count = lcl_false_count + 1;
    			} else {
           clearMsg("latitude");
       }
     }
  }

  if(document.getElementById("includeBlankCategoryOption").checked == false && document.getElementById("defaultCategoryID").selectedIndex == 0) {
     inlineMsg(document.getElementById("defaultCategoryID").id,'<strong>Invalid Value: </strong> Default Search Category cannot be "blank" if the "<strong>Include blank option in public-side dropdown list</strong>" checkbox has NOT been selected.',10,'defaultCategoryID');
     lcl_focus       = document.getElementById("defaultCategoryID");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("includeBlankCategoryOption");
  }

  if(document.getElementById("description").value=="") {
     inlineMsg(document.getElementById("description").id,'<strong>Required Field Missing: </strong> Description',10,'description');
     lcl_focus       = document.getElementById("description");
     lcl_false_count = lcl_false_count + 1;
  }else{
     clearMsg("description");
  }

<% end if %>
  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("datamgr_types_maint").submit();
     return true;
  }
}

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
  w = 600;
  h = 400;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";
  lcl_folderStart     = 0;

  //Determine which options will be displayed
  if((p_displayDocuments=="")||(p_displayDocuments==undefined)) {
      lcl_displayDocuments = "";
  }else{
      lcl_displayDocuments = "&displayDocuments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayActionLine=="")||(p_displayActionLine==undefined)) {
      lcl_displayActionLine = "";
  }else{
      lcl_displayActionLine = "&displayActionLine=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayPayments=="")||(p_displayPayments==undefined)) {
      lcl_displayPayments = "";
  }else{
      lcl_displayPayments = "&displayPayments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayURL=="")||(p_displayURL==undefined)) {
      lcl_displayURL = "";
  }else{
      lcl_displayURL = "&displayURL=Y";
  }

  if(lcl_folderStart > 0) {
     //lcl_showFolderStart = "&folderStart=unpublished_documents";
     lcl_showFolderStart = "&folderStart=CITY_ROOT";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
		    var caretPos = textEl.caretPos;
  			 caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? text + ' ' : text;
  } else {
   			textEl.value = textEl.value + text;
	 }
}

//function deleteField(iRowID) {
//  document.getElementById("deleteField" + iRowID).value = "Y";
//  document.getElementById("addFieldRow" + iRowID).style.display = "none";
//}

function getTemplateFields(iFieldID) {
<%
  lcl_js_rootadmin     = "N"
  lcl_js_islimited     = "N"
  lcl_js_istemplate    = "N"
  lcl_js_isdisplayonly = "Y"

  if lcl_isRootAdmin then
     lcl_js_rootadmin = "Y"
  end if

  if lcl_isLimited then
     lcl_js_islimited = "Y"
  end if

  if lcl_isTemplate then
     lcl_js_istemplate = "Y"
  end if
%>
  var lcl_isRootAdmin    = "<%=lcl_js_rootadmin%>";
  var lcl_isLimited      = "<%=lcl_js_islimited%>";
  var lcl_isTemplate     = "<%=lcl_js_istemplate%>";
  var lcl_isDisplayOnly  = "<%=lcl_js_isdisplayonly%>";
  var lcl_mptid          = document.getElementById(iFieldID).value;
  var lcl_iframe_url     = ""
  var lcl_iframe_width   = "0";
  var lcl_iframe_height  = "0";

  if(lcl_mptid != "") {
     lcl_iframe_url  = 'getDMTypeTemplateFields.asp';
     lcl_iframe_url += '?dm_typeid='     + encodeURIComponent(lcl_mptid);
     lcl_iframe_url += '&isRootAdmin='   + encodeURIComponent(lcl_isRootAdmin);
     lcl_iframe_url += '&isLimited='     + encodeURIComponent(lcl_isLimited);
     lcl_iframe_url += '&isTemplate='    + encodeURIComponent(lcl_isTemplate);
     lcl_iframe_url += '&isDisplayOnly=' + encodeURIComponent(lcl_isDisplayOnly);

     lcl_iframe_width  = "760";
     lcl_iframe_height = "300";
  }

  document.getElementById("previewDMTemplateFields").width  = lcl_iframe_width;
  document.getElementById("previewDMTemplateFields").height = lcl_iframe_height;
  document.getElementById("previewDMTemplateFields").src    = lcl_iframe_url;
}

function displayTemplateFields(p_code) {
  document.getElementById("previewDMTemplateFields").innerHTML = p_code;
}

function displayFeature(p_fieldid, p_displayid) {
  document.getElementById(p_displayid).innerHTML='[' + document.getElementById(p_fieldid).value + ']'
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
</script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">
<% ShowHeader sLevel %>
<!-- #include file="../menu/menu.asp" -->
<%
  response.write "<form name=""datamgr_types_maint"" id=""datamgr_types_maint"" method=""post"" action=""datamgr_types_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""user_action"" id=""user_action"" size=""4"" maxlength=""20"" value="""" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" size=""5"" maxlength=""5"" value="""     & lcl_dm_typeid    & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""screen_mode"" id=""screen_mode"" size=""4"" maxlength=""4"" value=""" & lcl_screen_mode  & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" size=""4"" maxlength=""10"" value="""            & session("orgid") & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" size=""10"" maxlength=""50"" value="""                   & lcl_feature      & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""t"" id=""t"" size=""5"" maxlength=""5"" value="""                     & request("t")     & """ />" & vbcrlf

  if lcl_screen_mode = "EDIT" then
     response.write "  <input type=""hidden"" name=""layoutid"" id=""layoutid"" value=""" & lcl_layoutid & """ size=""5"" maxlength=""10"" />" & vbcrlf
  end if

  if not lcl_isRootAdmin then
     response.write "  <input type=""hidden"" name=""feature_public"" id=""feature_public"" value="""                   & lcl_feature_public          & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain"" id=""feature_maintain"" value="""               & lcl_feature_maintain        & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain_fields"" id=""feature_maintain_fields"" value=""" & lcl_feature_maintain_fields & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_owners"" id=""feature_owners"" value="""                   & lcl_feature_owners          & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""accountInfoSectionID"" id=""accountInfoSectionID"" value="""       & lcl_accountInfoSectionID    & """ />" & vbcrlf
     response.write "  <textarea name=""intro_message"" id=""intro_message"" style=""display:none"">"                   & lcl_intro_message           & "</textarea>" & vbcrlf
  end if

  if lcl_isTemplate then
     response.write "  <input type=""hidden"" name=""isTemplate"" id=""isTemplate"" value=""Y"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""useAdvancedSearch"" id=""useAdvancedSearch"" value=""Y"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""enableOwnerMaint"" id=""enableOwnerMaint"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""displayMap"" id=""displayMap"" value=""Y"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""mappointcolor"" id=""mappointcolor"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""latitude"" id=""latitude"" value="""" size=""15"" maxlength=""10"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""longitude"" id=""longitude"" value="""" size=""15"" maxlength=""10"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""defaultzoomlevel"" id=""defaultzoomlevel"" value="""" size=""10"" maxlength=""10"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_public"" id=""feature_public"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain"" id=""feature_maintain"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain_fields"" id=""feature_maintain_fields"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_owners"" id=""feature_owners"" value="""" />" & vbcrlf
  else
     response.write "  <input type=""hidden"" name=""isTemplate"" id=""isTemplate"" value=""N"" />" & vbcrlf

     if not lcl_isRootAdmin then
        lcl_enableOwnerMaint_value = ""

        if lcl_enableOwnerMaint then
           lcl_enableOwnerMaint_value = "Y"
        end if

        response.write "  <input type=""hidden"" name=""enableOwnerMaint"" id=""enableOwnerMaint"" value=""" & lcl_enableOwnerMaint_value & """ />" & vbcrlf
        response.write "  <input type=""hidden"" name=""assignedto"" id=""assignedto"" value=""" & lcl_assignedto & """ />" & vbcrlf
     end if
  end if

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_pagetitle & ": " & lcl_screen_mode & lcl_isTemplate_title & "</strong></font><br />" & vbcrlf

  if not lcl_isLimited then
     response.write "<input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='datamgr_types_list.asp" & replace(lcl_isTemplate_url,"&","?") & "';"" />" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf

                            displayButtons "TOP", _
                                           lcl_screen_mode, _
                                           lcl_canDelete, _
                                           lcl_isLimited, _
                                           lcl_isTemplate

  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""left"" colspan=""2"">" & lcl_pagetitle & lcl_isTemplate_title & "</th>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">Description:</td>" & vbcrlf
  response.write "                <td width=""100%""><input type=""text"" name=""description"" id=""description"" value=""" & lcl_description & """ size=""50"" maxlength=""500"" onchange=""clearMsg('description');"" /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td>&nbsp;</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""isActive"" id=""isActive"" value=""Y""" & lcl_checked_isActive & "/>&nbsp;Active</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  if not lcl_isTemplate then
     response.write "            <tr>" & vbcrlf
     response.write "                <td>&nbsp;</td>" & vbcrlf
     response.write "                <td><input type=""checkbox"" name=""useAdvancedSearch"" id=""useAdvancedSearch"" value=""Y""" & lcl_checked_useAdvancedSearch & "/>&nbsp;Use Advanced Search (public-side)</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""2"">" & vbcrlf
     response.write "                    <p>" & vbcrlf
     response.write "                    <fieldset class=""fieldset"">" & vbcrlf
     response.write "                      <legend style=""color:#000080"">Map Setup&nbsp;</legend>" & vbcrlf
     response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
     response.write "                        <tr valign=""top"">" & vbcrlf
     response.write "                            <td colspan=""3"" style=""padding-bottom:5px;"">" & vbcrlf
     'response.write "                                <input type=""checkbox"" name=""displayMap"" id=""displayMap"" value=""Y""" & lcl_checked_displayMap & " onclick=""enableDisableMapSetupFields();"" />" & vbcrlf
     response.write "                                <input type=""checkbox"" name=""displayMap"" id=""displayMap"" value=""Y""" & lcl_checked_displayMap & " />" & vbcrlf
     response.write "                                Display Map on public page" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">" & vbcrlf
     response.write "                                <div id=""requiredFieldsMsg"">" & vbcrlf

     if lcl_isRootAdmin then
        response.write "                                  <strong>*** REMINDER ***</strong><br />" & vbcrlf
        response.write "                                   3 Field Types are required when this value is checked:<br />" & vbcrlf
        response.write "                                  <strong>ADDRESS, LATITUDE, and LONGITUDE</strong>.  Without these<br />" & vbcrlf
        response.write "                                  values the map will NOT function properly." & vbcrlf
     end if

     response.write "                                </div>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     'response.write "                        <tr>" & vbcrlf
     'response.write "                            <td nowrap=""nowrap"">Map Point Color:</td>" & vbcrlf
     'response.write "                            <td colspan=""3"">" & vbcrlf
     'response.write "                                <select name=""mappointcolor"" id=""mappointcolor"">" & vbcrlf
     '                                                  displayMapPointColors lcl_mappointcolor
     'response.write "                                </select>" & vbcrlf
     'response.write "                            </td>" & vbcrlf
     'response.write "                        </tr>" & vbcrlf
     response.write "                        <tr>" & vbcrlf
     response.write "                            <td colspan=""4"" style=""padding-top:10px; color:#800000"">*** Latitude and Longitude are used to ""center"" the map displayed on the public-side ***</td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                        <tr>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">Latitude:</td>" & vbcrlf
     response.write "                            <td><input type=""text"" name=""latitude"" id=""latitude"" value=""" & lcl_latitude & """ size=""15"" maxlength=""10"" onchange=""clearMsg('latitude');"" /></td>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">&nbsp;Longitude: <input type=""text"" name=""longitude"" id=""longitude"" value=""" & lcl_longitude & """ size=""15"" maxlength=""10"" onchange=""clearMsg('longitude');"" /></td>" & vbcrlf
     response.write "                            <td><input type=""button"" name=""getLatLongButton"" id=""getLatLongButton"" value=""Get Org Latitude/Longitude"" class=""button"" onclick=""getOrgLatLong();"" /></td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                        <tr>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">Zoom Level:</td>" & vbcrlf
     response.write "                            <td colspan=""3"">" & vbcrlf
     response.write "			                      			    <input type=""text"" name=""defaultzoomlevel"" id=""defaultzoomlevel"" value=""" & lcl_defaultzoomlevel & """ size=""10"" maxlength=""10"" onchange=""clearMsg('defaultzoomlevel');"" />" & vbcrlf
     response.write "			                      			    <span style=""color:#800000"">Zoom Levels: 0 to 21+ || Max Zoom OUT: 0 || Default Zoom: 13</span>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                        <tr>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">Google Map Type:</td>" & vbcrlf
     response.write "                            <td colspan=""3"">" & vbcrlf
     response.write "			                     <select name=""googleMapType"" id=""googleMapType"">" & vbcrlf
     response.write "                                  <option value=""ROADMAP"""   & lcl_selected_googleMapType_roadmap   & ">Roadmap</option>" & vbcrlf
     response.write "                                  <option value=""SATELLITE""" & lcl_selected_googleMapType_satellite & ">Satellite</option>" & vbcrlf
     response.write "                                  <option value=""HYBRID"""    & lcl_selected_googleMapType_hybrid    & ">Hybrid</option>" & vbcrlf
     response.write "                                  <option value=""TERRAIN"""   & lcl_selected_googleMapType_terrain   & ">Terrain</option>" & vbcrlf
     response.write "			                     </select>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                        <tr valign=""top"">" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">Google Map Marker:</td>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"" colspan=""3"">" & vbcrlf
     response.write "			                     <select name=""googleMapMarker"" id=""googleMapMarker"">" & vbcrlf
     response.write "                                  <option value=""GOOGLE"""        & lcl_selected_googleMapMarker_google        & ">Google (Default)</option>" & vbcrlf
     response.write "                                  <option value=""CUSTOMMARKER1""" & lcl_selected_googleMapMarker_custommarker1 & ">Custom Marker 1</option>" & vbcrlf
     response.write "			                     </select>" & vbcrlf
     response.write "                                <div id=""googleMapMarkerImg"" style=""display:inline;""></div>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                      </table>" & vbcrlf
     response.write "                    </fieldset>" & vbcrlf
     response.write "                    </p>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if

 'BEGIN: Root Admin Fields ----------------------------------------------------
  if lcl_isRootAdmin then
     if not lcl_isTemplate then
        response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td nowrap=""nowrap"">Feature<br />(public-side):</td>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <select name=""feature_public"" id=""feature_public"" onchange=""displayFeature('feature_public','displayfeature_public');"">" & vbcrlf
        response.write "                      <option value=""""></option>" & vbcrlf
                                              showFeatureOptions lcl_feature_public
        response.write "                    </select>" & vbcrlf
        response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_public"" id=""helpFeature_public"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
        response.write "                    <div name=""helpFeature_public_text"" id=""helpFeature_public_text"" class=""helpOptionText"">" & vbcrlf
        response.write "                      <p><strong>E-GOV TIP:</strong><br />This option is used to connect the DataManager public displayed feature (home page option - i.e. Available Properties feature) to this DM Type.</p>" & vbcrlf
        response.write "                    </div>" & vbcrlf
        response.write "                    <span id=""displayfeature_public"" style=""color:#800000;""></span>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td nowrap=""nowrap"">Feature<br />(maintenance):</td>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <select name=""feature_maintain"" id=""feature_maintain"" onchange=""displayFeature('feature_maintain','displayfeature_maintain');"">" & vbcrlf
        response.write "                      <option value=""""></option>" & vbcrlf
                                              showFeatureOptions lcl_feature_maintain
        response.write "                    </select>" & vbcrlf
        response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_maintain"" id=""helpFeature_maintain"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
        response.write "                    <div name=""helpFeature_maintain_text"" id=""helpFeature_maintain_text"" class=""helpOptionText"">" & vbcrlf
        response.write "                      <p><strong>E-GOV TIP:</strong><br />This option is used to determine which feature/permission to use when accessing the maintenance screen.</p>" & vbcrlf
        response.write "                    </div>" & vbcrlf
        response.write "                    <span id=""displayfeature_maintain"" style=""color:#800000;""></span>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td nowrap=""nowrap"">Feature<br />(maintenance - fields):</td>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <select name=""feature_maintain_fields"" id=""feature_maintain_fields"" onchange=""displayFeature('feature_maintain_fields','displayfeature_maintain_fields');"">" & vbcrlf
        response.write "                      <option value=""""></option>" & vbcrlf
                                              showFeatureOptions lcl_feature_maintain_fields
        response.write "                    </select>" & vbcrlf
        response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_maintain_fields"" id=""helpFeature_maintain_fields"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
        response.write "                    <div name=""helpFeature_maintain_fields_text"" id=""helpFeature_maintain_fields_text"" class=""helpOptionText"">" & vbcrlf
        response.write "                      <p><strong>E-GOV TIP:</strong><br />This option is used to determine which feature/permission to use when accessing the DM Types maintenance screen for admins.  It will limit the screen to this specific DM Type from the navigation menu.</p>" & vbcrlf
        response.write "                    </div>" & vbcrlf
        response.write "                    <span id=""displayfeature_maintain_fields"" style=""color:#800000;""></span>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td nowrap=""nowrap"">Feature<br />(maintenance - owners):</td>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <select name=""feature_owners"" id=""feature_owners"" onchange=""displayFeature('feature_owners','displayfeature_owners');"">" & vbcrlf
        response.write "                      <option value=""""></option>" & vbcrlf
                                              showFeatureOptions lcl_feature_owners
        response.write "                    </select>" & vbcrlf
        response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_owners"" id=""helpFeature_owners"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
        response.write "                    <div name=""helpFeature_owners_text"" id=""helpFeature_owners_text"" class=""helpOptionText"">" & vbcrlf
        response.write "                      <p><strong>E-GOV TIP:</strong><br />This option is used to determine which feature/permission to use when accessing the DM Owners/Editors maintenance screen for admins.  It will limit the screen to this specific DM Type from the navigation menu.</p>" & vbcrlf
        response.write "                    </div>" & vbcrlf
        response.write "                    <span id=""displayfeature_owners"" style=""color:#800000;""></span>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr>" & vbcrlf
        response.write "                <td>&nbsp;</td>" & vbcrlf
        response.write "                <td><input type=""checkbox"" name=""enableOwnerMaint"" id=""enableOwnerMaint"" value=""Y""" & lcl_checked_enableOwnerMaint & "/>&nbsp;Enable Owner Maintenance (public-side)</td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr>" & vbcrlf
        response.write "                <td>Assigned To:</td>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <select name=""assignedto"" id=""assignedto"">" & vbcrlf
                                              displayAssignEmails session("orgid"), lcl_assignedto
        response.write "                    </select>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf

        if lcl_screen_mode = "ADD" then
           response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
           response.write "            <tr valign=""top"">" & vbcrlf
           response.write "                <td colspan=""2"">" & vbcrlf
           response.write "                    <fieldset class=""fieldset"">" & vbcrlf
           response.write "                      <legend>Template Setup&nbsp;</legend>" & vbcrlf
           response.write "                      <table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
           response.write "                        <tr>" & vbcrlf
           response.write "                            <td>" & vbcrlf
           response.write "                                <div style=""color:#800000; margin-bottom:5px;"">" & vbcrlf
           response.write "                                  *** NOTE: Any changes made to these template fields will NOT be recognized " & vbcrlf
           response.write "                                  when creating a new DM Type.<br />" & vbcrlf
           response.write "                                  All changes to template fields need to be made to the template itself " & vbcrlf
           response.write "                                  before creating a DM Type." & vbcrlf
           response.write "                                </div>" & vbcrlf
           response.write "                                Template: <select name=""DMTemplateID"" id=""DMTemplateID"">" & vbcrlf
                                                             displayTemplateOptions
           response.write "                                </p>" & vbcrlf
           'response.write "                                <input type=""button"" name=""previewButton"" id=""previewButton"" value=""Preview Template"" class=""button"" onclick=""getTemplateFields('DMTemplateID')"" />" & vbcrlf
           response.write "                            </td>" & vbcrlf
           response.write "                        </tr>" & vbcrlf
           response.write "                        <tr>" & vbcrlf
           response.write "                            <td>" & vbcrlf
           response.write "                                <iframe name=""previewDMTemplateFields"" id=""previewDMTemplateFields"" frameborder=""0"" src="""" width=""0"" height=""0""></iframe>" & vbcrlf
           response.write "                            </td>" & vbcrlf
           response.write "                        </tr>" & vbcrlf
           response.write "                      </table>" & vbcrlf
           response.write "                    </fieldset>" & vbcrlf
           response.write "                </td>" & vbcrlf
           response.write "            </tr>" & vbcrlf
        end if

        response.write "            <tr><td colspan=""2""><hr style=""color:#efefef"" /></td></tr>" & vbcrlf
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td nowrap=""nowrap"" style=""height:15px"">Intro Message:<br />(public-side)</td>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <textarea name=""intro_message"" id=""intro_message"" class=""textarea_intromessage"">" & lcl_intro_message & "</textarea>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
     end if

     response.write "            <tr>" & vbcrlf
     response.write "                <td nowrap=""nowrap"">Account Info Section:</td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""accountInfoSectionID"" id=""accountInfoSectionID"">" & vbcrlf
     response.write "                      <option value=""""></option>" & vbcrlf
                                           displayAccountInfoOptions lcl_accountInfoSectionID
     response.write "                    </select>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td nowrap=""nowrap"">Layout:</td>" & vbcrlf
     response.write "                <td>" & vbcrlf

    'When adding only show the layout options available.
    'When editing show the layout selected and the "Maintain Layout" button.
     if lcl_screen_mode = "ADD" then
        response.write "                    <select name=""layoutid"" id=""layoutid"">" & vbcrlf
        response.write "                      <option value=""""></option>" & vbcrlf
                                              displayLayoutOptions lcl_layoutid
        response.write "                    </select>" & vbcrlf
     else
       'Determine what browser is being used.
       'This check is ONLY for IE9.  The drag-n-drop does not currently work with IE9 (jQuery).
       'We will disable it for now and show a message is IE9 is being used.
        'lcl_browser = ucase(request.serverVariables("HTTP_USER_AGENT"))
        'ie9         = instr(lcl_browser,"MSIE 9")

        response.write                      lcl_layoutname & "&nbsp;" & vbcrlf

        'if ie9 > 0 then
        '   response.write "[ie9 message here!]" & vbcrlf
        'else
           response.write "                    <input type=""button"" name=""maintainLayoutButton"" id=""maintainLayoutButton"" value=""Maintain Layout"" class=""button"" onclick=""maintainLayout()"" />" & vbcrlf
        'end if
     end if

     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Root Admin Fields ------------------------------------------------------

  if lcl_screen_mode = "EDIT" then
     response.write "            <tr>" & vbcrlf
     response.write "                <td nowrap=""nowrap"">Default Search Category:<br />(public-side search)</td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    <select name=""defaultCategoryID"" id=""defaultCategoryID"" onchange=""clearMsg('defaultCategoryID');"">" & vbcrlf
     response.write "                      <option value=""0""></option>" & vbcrlf
                                           lcl_parent_categoryid = 0

                                           displayDMTCategories session("orgid"), _
                                                                lcl_dm_typeid, _
                                                                lcl_parent_categoryid, _
                                                                lcl_defaultcategoryid
     response.write "                    </select>" & vbcrlf
     response.write "                    <input type=""button"" name=""maintainDMTypeCategories"" id=""maintainDMTypeCategories"" class=""button"" value=""Maintain Categories"" onclick=""maintainCategories()"" />" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td>&nbsp;</td>" & vbcrlf
     response.write "                <td width=""100%"">" & vbcrlf
     response.write "                    <input type=""checkbox"" name=""includeBlankCategoryOption"" id=""includeBlankCategoryOption"" value=""Y""" & lcl_checked_includeBlankCategory & " onclick=""clearMsg('defaultCategoryID');"" /> Include ""blank"" option in public-side dropdown list" & vbcrlf
     response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_defaultCategory"" id=""helpFeature_defaultCategory"" class=""helpOption"" alt=""Click for more info"" />" & vbcrlf
     response.write "                    <div name=""helpFeature_defaultCategory_text"" id=""helpFeature_defaultCategory_text"" class=""helpOptionText"">" & vbcrlf
     response.write "                      <p><strong>E-GOV TIP:</strong><br />Check this option if you wish for there to be a blank category option available in the dropdown list in the search on the public-side.</p>" & vbcrlf
     response.write "                    </div>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  end if

  if lcl_displayCreatedByInfo <> "" then
     response.write "            <tr>" & vbcrlf
     response.write "                <td nowrap=""nowrap"" style=""height:15px"">Created By:</td>" & vbcrlf
     response.write "                <td style=""color:#800000"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if

  if lcl_displayLastModifiedByInfo <> "" then
     response.write "            <tr>" & vbcrlf
     response.write "                <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
     response.write "                <td style=""color:#800000"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if

  response.write "          </table>" & vbcrlf
  response.write "          </p>" & vbcrlf

 'Build the Layout
  response.write "<p>" & vbcrlf

 'Retrieve any/all fields related to this DM Type
 'ONLY show these field if "screen mode" = "EDIT"
  if lcl_screen_mode = "EDIT" then
     'if not lcl_isOriginalLayout then
     '   response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""0"" size=""3"" maxlength=""100"" />" & vbcrlf
     '   buildMapPointLayout lcl_layoutid, lcl_dm_typeid
     'else
        'displayMPTypesFields session("orgid"), lcl_dm_typeid, lcl_isRootAdmin, lcl_isLimited, False
        lcl_sectionid     = 0
        lcl_isDisplayOnly = False

        displayDMTSectionFields session("orgid"), _
                                lcl_dm_typeid, _
                                lcl_sectionid, _
                                lcl_isRootAdmin, _
                                lcl_isLimited, _
                                lcl_isDisplayOnly
     'end if
  end if

 'Display the bottom row of buttons
  displayButtons "BOTTOM", _
                 lcl_screen_mode, _
                 lcl_canDelete, _
                 lcl_isLimited, _
                 lcl_isTemplate

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  response.write "</p>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf

 'Determine if there are any scripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if
%>

<!--#include file="../admin_footer.asp"-->
<%
response.write "</body>" & vbcrlf
response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode, iCanDelete, iIsLimited, iIsTemplate)

  if iTopBottom <> "" then
     iTopBottom = UCASE(iTopBottom)
  else
     iTopBottom = "TOP"
  end if

  if iTopBottom = "BOTTOM" then
     lcl_style_div = "margin-top:5px;"
  else
     lcl_style_div = "margin-bottom:5px;"
  end if

  'lcl_return_parameters = "?sc_org_name=" & session("sc_org_name") & "&sc_show_members=" & session("sc_show_members")
  lcl_return_parameters = ""

  if iIsTemplate then
     if lcl_return_parameters = "" then
        lcl_return_parameters = "?t=Y"
     else
        lcl_return_parameters = lcl_return_parameters & "&t=Y"
     end if
  end if

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf

  if not iIsLimited then
     response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='datamgr_types_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf
  end if

  if lcl_screen_mode = "ADD" then
     'response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" onclick=""return validateFields('ADDANOTHER');"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     if iCanDelete AND not iIsLimited then
        response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     end if

     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub showFeatureOptions(iFeature)

  sSQL = "SELECT feature, featurename "
  sSQL = sSQL & " FROM egov_organization_features "
  sSQL = sSQL & " ORDER BY featurename "

  set oFeatureOptions = Server.CreateObject("ADODB.Recordset")
  oFeatureOptions.Open sSQL, Application("DSN"), 3, 1

  if not oFeatureOptions.eof then
     do while not oFeatureOptions.eof

        if UCASE(iFeature) = UCASE(oFeatureOptions("feature")) then
           lcl_selected_feature = " selected=""selected"""
        else
           lcl_selected_feature = ""
        end if

        response.write "  <option value=""" & oFeatureOptions("feature") & """" & lcl_selected_feature & ">" & oFeatureOptions("featurename") & "</option>" & vbcrlf

        oFeatureOptions.movenext
     loop
  end if

  oFeatureOptions.close
  set oFeatureOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayAccountInfoOptions(iAccountInfoSectionID)

  sSQL = "SELECT dms.sectionid, "
  sSQL = sSQL & " dms.sectionname, "
  sSQL = sSQL & " o.orgcity "
  sSQL = sSQL & " FROM egov_dm_sections dms "
  sSQL = sSQL &      " LEFT OUTER JOIN organizations o ON o.orgid = dms.section_orgid "
  sSQL = sSQL & " WHERE dms.isAccountInfoSection = 1 "
  sSQL = sSQL & " AND dms.isActive = 1 "
  sSQL = sSQL & " ORDER BY UPPER(dms.sectionname) "

  set oAccountInfoOptions = Server.CreateObject("ADODB.Recordset")
  oAccountInfoOptions.Open sSQL, Application("DSN"), 3, 1

  if not oAccountInfoOptions.eof then
     do while not oAccountInfoOptions.eof

        lcl_sectionname        = oAccountInfoOptions("sectionname")
        lcl_selected_sectionid = ""

        if lcl_sectionname <> "" then
           lcl_sectionname = lcl_sectionname & " [" & oAccountInfoOptions("orgcity") & "]"
        end if

        if clng(iAccountInfoSectionID) = oAccountInfoOptions("sectionid") then
           lcl_selected_sectionid = " selected=""selected"""
        end if

        response.write "  <option value=""" & oAccountInfoOptions("sectionid") & """" & lcl_selected_sectionid & ">" & lcl_sectionname & "</option>" & vbcrlf

        oAccountInfoOptions.movenext
     loop
  end if

  oAccountInfoOptions.close
  set oAccountInfoOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayDMTSectionFields(iOrgID, iDMTypeID, iSectionID, iIsRootAdmin, iIsLimited, iIsDisplayOnly)

 'Determine if we are editing fields for a MapPointType or a Section
  lcl_dmtypeid  = iDMTypeID
  lcl_sectionid = iSectionID
  sSQL          = ""

  'if not iIsDisplayOnly then
  '   response.write "<div style=""margin-top:20px; margin-bottom:5px;"">" & vbcrlf
  '   response.write "  <strong>" & lcl_sectiontitle & "</strong><br />" & vbcrlf
  '   response.write "  <input type=""button"" name=""addMPTField"" id=""addMPTField"" value=""Add Field"" class=""button"" onclick=""addFieldRow('" & iIsRootAdmin & "', '" & iIsLimited & "', 'addFieldTBL','totalFields','addFieldRow', '" & lcl_edit_type & "');"" />" & vbcrlf
  '   response.write "</div>" & vbcrlf
  'end if

  response.write "<table id=""addFieldTBL"" border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "<thead>" & vbcrlf
  response.write "  <tr id=""addFieldRow0"">" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "      <th align=""left"">Field Name</th>" & vbcrlf
  response.write "      <th>Display In<br />Results</th>" & vbcrlf
  response.write "      <th>Display On<br />Info Page</th>" & vbcrlf
  response.write "      <th>Display<br />Order</th>" & vbcrlf
  response.write "      <th>In Public<br />Search</th>" & vbcrlf
  response.write "      <th>Display<br />Label</th>" & vbcrlf

  if lcl_isRootAdmin then
     response.write "      <th id=""sidebarlink_columnheader"">Sidebar<br />Link (key)</th>" & vbcrlf
  end if

  response.write "      <th align=""left"">Section Name</th>" & vbcrlf

  'if lcl_isRootAdmin then
  '   response.write "      <th align=""left"">Transfer Data To...</th>" & vbcrlf
  'end if

  'response.write "      <th>Remove</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</thead>" & vbcrlf
  response.write "<tbody>" & vbcrlf

  iRowCount   = 0
  lcl_bgcolor = "#ffffff"

 'Retrieve all of the MapPoint Type Fields that:
 '  1. have an "ACTIVE" MapPoint Type Section Field
 '  2. have an "ACTIVE" MapPoint Section
 '  3. have an "ACTIVE" MapPoint Section Field
  sSQL = "SELECT dmtf.dm_fieldid, "
  sSQL = sSQL & " dmtf.dm_typeid, "
  sSQL = sSQL & " dmtf.orgid, "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQL = sSQL & " dmtf.section_fieldid, "
  sSQL = sSQL & " dmtf.displayInResults, "
  sSQL = sSQL & " dmtf.displayInInfoPage, "
  sSQL = sSQL & " dmtf.resultsOrder, "
  sSQL = sSQL & " dmtf.inPublicSearch, "
  sSQL = sSQL & " dmtf.displayFieldName, "
  sSQL = sSQL & " dmtf.isSidebarLink, "
  sSQL = sSQL & " sf.fieldname, "
  sSQL = sSQL & " sf.fieldtype, "
  sSQL = sSQL & " dms.sectionname "
  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &            " ON dmtf.dm_sectionid = dmts.dm_sectionid "
  sSQL = sSQL &            " AND dmts.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
  sSQL = sSQL &            " ON dms.sectionid = dmts.sectionid "
  sSQL = sSQL &            " AND dms.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields sf "
  sSQL = sSQL &            " ON dmtf.section_fieldid = sf.section_fieldid "
  sSQL = sSQL &            " AND sf.isActive = 1 "
  sSQL = sSQL & " WHERE dmtf.dm_typeid = " & lcl_dmtypeid
  sSQL = sSQL & " AND dmtf.orgid = " & iOrgID
  sSQL = sSQL & " AND sf.fieldtype <> 'CATEGORIES_FIELD' "
  sSQL = sSQL & " AND sf.fieldtype <> 'SUBCATEGORIES_FIELD' "
  sSQL = sSQL & " ORDER BY dmtf.resultsOrder, dmtf.dm_fieldid "

  set oDMTSectionFields = Server.CreateObject("ADODB.Recordset")
  oDMTSectionFields.Open sSQL, Application("DSN"), 3, 1

  if not oDMTSectionFields.eof then
     do while not oDMTSectionFields.eof

        iRowCount        = iRowCount + 1
        lcl_bgcolor      = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_resultsOrder = oDMTSectionFields("resultsOrder")

        lcl_checked_displayInResults  = isCheckboxChecked(oDMTSectionFields("displayInResults"))
        lcl_checked_displayInInfoPage = isCheckboxChecked(oDMTSectionFields("displayInInfoPage"))
        lcl_checked_inPublicSearch    = isCheckboxChecked(oDMTSectionFields("inPublicSearch"))
        lcl_checked_displayFieldName  = isCheckboxChecked(oDMTSectionFields("displayFieldName"))
        lcl_checked_isSidebarLink     = isCheckboxChecked(oDMTSectionFields("isSidebarLink"))

        response.write "  <tr id=""addFieldRow" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ align=""center"">" & vbcrlf
        response.write "      <td align=""left"">" & vbcrlf
        response.write "          <img src=""arrow.png"" width=""12"" height=""12"" title=""click to drag and reorder"" class=""dragDropArrows"" /> " & iRowCount & "." & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td align=""left"">" & vbrlf
        response.write "          <input type=""hidden"" name=""dm_fieldid"      & iRowCount & """ id=""dm_fieldid"      & iRowCount & """ value=""" & oDMTSectionFields("dm_fieldid")      & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_sectionid"    & iRowCount & """ id=""dm_sectionid"    & iRowCount & """ value=""" & oDMTSectionFields("dm_sectionid")    & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""section_fieldid" & iRowCount & """ id=""section_fieldid" & iRowCount & """ value=""" & oDMTSectionFields("section_fieldid") & """ />" & vbcrlf
        response.write            oDMTSectionFields("fieldname") & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInResults" & iRowCount & """ id=""displayInResults" & iRowCount & """ value=""1""" & lcl_checked_displayInResults & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInInfoPage" & iRowCount & """ id=""displayInInfoPage" & iRowCount & """ value=""1""" & lcl_checked_displayInInfoPage & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        'response.write "          <input type=""text"" name=""resultsOrder" & iRowCount & """ id=""resultsOrder" & iRowCount & """ value=""" & lcl_resultsOrder & """ size=""3"" maxlength=""5"" />" & vbcrlf
        response.write "          <input type=""text"" name=""resultsOrder" & iRowCount & """ id=""resultsOrder" & iRowCount & """ value=""" & iRowCount & """ size=""3"" maxlength=""5"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""inPublicSearch" & iRowCount & """ id=""inPublicSearch" & iRowCount & """ value=""1""" & lcl_checked_inPublicSearch & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td><input type=""checkbox"" name=""displayFieldName" & iRowCount & """ id=""displayFieldName" & iRowCount & """ value=""1""" & lcl_checked_displayFieldName & " /></td>" & vbcrlf

        if lcl_isRootAdmin then
           response.write "      <td><input type=""checkbox"" name=""isSidebarLink" & iRowCount & """ id=""isSidebarLink" & iRowCount & """ value=""1""" & lcl_checked_isSidebarLink & " class=""isSidebarLinkCheckbox"" /></td>" & vbcrlf
        end if

        response.write "      <td align=""left"">" & oDMTSectionFields("sectionname") & "</td>" & vbcrlf
        'response.write "      <td>" & vbcrlf
        'response.write "          <input type=""checkbox"" name=""deleteField" & iRowCount & """ id=""deleteField" & iRowCount & """ value=""Y"" />" & vbcrlf
        'response.write "      </td>" & vbcrlf

        'if lcl_isRootAdmin then
        '   response.write "      <td align=""left"">" & vbcrlf
        '   response.write "          <select name=""transferData" & iRowCount & """ id=""transferData" & iRowCount & """>" & vbcrlf
        '   response.write "            <option value=""""></option>" & vbcrlf
        '                               displayTransferFieldOptions iOrgID, lcl_dmtypeid
        '   response.write "          </select>" & vbcrlf
        '   response.write "      </td>" & vbcrlf
        'end if

        response.write "  </tr>" & vbcrlf

        oDMTSectionFields.movenext
     loop
  end if

  oDMTSectionFields.close
  set oDMTSectionFields = nothing

  response.write "</tbody>" & vbcrlf

  response.write "</table>" & vbcrlf
  response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""" & iRowCount & """ size=""3"" maxlength=""100"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
'sub displayTransferFieldOptions(p_orgid, p_dmtypeid)

'  sSQL = "SELECT dmtf.dm_fieldid, "
'  sSQL = sSQL & " dmtf.dm_typeid, "
'  sSQL = sSQL & " dmtf.orgid, "
'  sSQL = sSQL & " dmtf.dm_sectionid, "
'  sSQL = sSQL & " dmtf.section_fieldid, "
'  sSQL = sSQL & " dmtf.displayInResults, "
'  sSQL = sSQL & " dmtf.displayInInfoPage, "
'  sSQL = sSQL & " dmtf.resultsOrder, "
'  sSQL = sSQL & " dmtf.inPublicSearch, "
'  sSQL = sSQL & " dmtf.displayFieldName, "
'  sSQL = sSQL & " dmtf.isSidebarLink, "
'  sSQL = sSQL & " sf.fieldname, "
'  sSQL = sSQL & " dms.sectionname "
'  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
'  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
'  sSQL = sSQL &            " ON dmtf.dm_sectionid = dmts.dm_sectionid "
'  sSQL = sSQL &            " AND dmts.isActive = 1 "
'  sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
'  sSQL = sSQL &            " ON dms.sectionid = dmts.sectionid "
'  sSQL = sSQL &            " AND dms.isActive = 1 "
'  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields sf "
'  sSQL = sSQL &            " ON dmtf.section_fieldid = sf.section_fieldid "
'  sSQL = sSQL &            " AND sf.isActive = 1 "
'  sSQL = sSQL & " WHERE dmtf.dm_typeid = " & p_dmtypeid
'  sSQL = sSQL & " AND dmtf.orgid = " & p_orgid
'  sSQL = sSQL & " ORDER BY dmtf.resultsOrder, dmtf.dm_fieldid "

'  set oTransferFieldOptions = Server.CreateObject("ADODB.Recordset")
'  oTransferFieldOptions.Open sSQL, Application("DSN"), 3, 1

'  if not oTransferFieldOptions.eof then
'     do while not oTransferFieldOptions.eof

'        response.write "  <option value=""0"">" & oTransferFieldOptions("sectionname") & ": " & oTransferFieldOptions("fieldname") & "</option>" & vbcrlf

'        oTransferFieldOptions.movenext
'     loop
'  end if

'  oTransferFieldOptions.close
'  set oTransferFieldOptions = nothing

'end sub

'-----------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = p_value
  end if

  dbsafe = lcl_value

end function

'------------------------------------------------------------------------------
sub displayAssignEmails(iOrgID, iUserID)

  dim sSQL, sOrgID, sUserID

  sSQL    = ""
  sOrgID  = 0
  sUserID = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iUserID <> "" then
     sUserID = clng(sUserID)
  end if

		sSQL = "SELECT userid, "
  sSQL = sSQL & " email, "
  sSQL = sSQL & " firstname, "
  sSQL = sSQL & " lastname "
		sSQL = sSQL & " FROM users "
		sSQL = sSQL & " WHERE orgid = " & sOrgID
		sSQL = sSQL & " AND (IsRootAdmin is null or IsRootAdmin = 0) "
		sSQL = sSQL & " AND email <> '' "
		sSQL = sSQL & " ORDER BY lastname, firstname"

		set oUsers = Server.CreateObject("ADODB.Recordset")
		oUsers.Open sSQL, Application("DSN"), 1, 3
		
		if not oUsers.eof then

  			response.write "  <option value=""0"">Please Select...</option>" & vbcrlf

  			do while not oUsers.eof

    				if iUserID = oUsers("userID") then
           lcl_selected = " selected=""selected"""
        else
           lcl_selected = ""
        end if

  						response.write "  <option value=""" & oUsers("userID") & """" & lcl_selected & ">" & replace(oUsers("FirstName"),"'","\'") & " " & replace(oUsers("LastName"),"'","\'") & "</option>" & vbcrlf

    				oUsers.movenext
 	  	loop

  end if

  oUsers.close
  set oUsers = nothing

end sub
%>