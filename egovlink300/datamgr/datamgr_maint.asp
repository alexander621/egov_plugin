<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!-- #include file="datamgr_build_sections_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an "owner/editor" to modify their DM Data info
'
' MODIFICATION HISTORY
' 1.0 06/13/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 dim sSQL

'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

'Retrieve the dmid to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("dmid") <> "" then
    lcl_dmid = request("dmid")

    if isnumeric(lcl_dmid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "datamgr.asp"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_sendToLabel = "Create"
    lcl_dmid  = 0
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature         = "datamgr_maint"
 lcl_featurename     = "DM Data"
 lcl_dm_typeid       = 0
 lcl_layoutid        = 0
 lcl_mappointcolor   = ""
 lcl_canSelectDMType = True

 if request("f") <> "" then
    lcl_feature     = request("f")
    lcl_featurename = getFeatureName(lcl_feature)
    lcl_dm_typeid   = getDMTypeByFeature(iorgid, lcl_feature)

    if lcl_screen_mode = "ADD" then
       lcl_onload          = "validateFields('ADD');"
       lcl_canSelectDMType = False
    end if
 end if

 lcl_featureNameLabel = lcl_featurename
 lcl_featureNameLabel = replace(lcl_featureNameLabel,"Maintain ","")

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

'Check for org features
 lcl_orghasfeature_feature            = orghasfeature(iorgid, lcl_feature)
 lcl_orghasfeature_feature_maintain   = orghasfeature(iorgid, lcl_feature)
 lcl_orghasfeature_issue_location     = orghasfeature(iorgid, "issue location")
 lcl_orghasfeature_large_address_list = orghasfeature(iorgid, "large address list")

'Check for a screen message
 lcl_success = request("success")
 'lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'Set up local variables
 lcl_orgid                 = iorgid
 lcl_cookie_userid         = ""
 lcl_categoryid            = 0
 lcl_isCreatedByAdmin      = false
 lcl_createdbyid           = 0
 lcl_createdbydate         = ""
 lcl_isLastModifiedByAdmin = false
 lcl_lastmodifiedbyid      = 0
 lcl_lastmodifiedbydate    = ""
 lcl_approvedeniedbyid     = 0
 lcl_approvedeniedbydate   = ""
 lcl_approvedeniedbyname   = ""
 lcl_isApproved_dmdata     = ""
 lcl_isActive              = 1
 lcl_isActive_value        = "N"
 lcl_checked_isactive      = " checked=""checked"""
 lcl_enableOwnerMaint      = 0
 lcl_parent_categoryid     = 0
 lcl_dm_latitude           = ""
 lcl_dm_longitude          = ""
 lcl_description           = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the DM Data
    sSQL = "SELECT dm.dmid, "
    sSQL = sSQL & " dm.dm_typeid, "
    sSQL = sSQL & " dmt.description, "
    sSQL = sSQL & " dmt.enableOwnerMaint, "
    sSQL = sSQL & " dm.orgid, "
    sSQL = sSQL & " dm.categoryid, "
    sSQL = sSQL & " dm.isCreatedByAdmin, "
    sSQL = sSQL & " dm.createdbyid, "
    sSQL = sSQL & " dm.createdbydate, "
    sSQL = sSQL & " dm.isActive, "
    sSQL = sSQL & " dm.isApproved, "
    sSQL = sSQL & " dm.isLastModifiedByAdmin, "
    sSQL = sSQL & " dm.lastmodifiedbyid, "
    sSQL = sSQL & " dm.lastmodifiedbydate, "
    sSQL = sSQL & " dm.approvedeniedbyid, "
    sSQL = sSQL & " dm.approvedeniedbydate, "
    sSQL = sSQL & " u.userfname + ' ' + u.userlname AS createdbyname_citizen, "
    sSQL = sSQL & " u2.userfname + ' ' + u2.userlname AS lastmodifiedbyname_citizen, "
    sSQL = sSQL & " 'Created By Admin' AS createdbyname_admin, "
    sSQL = sSQL & " 'Last Modified By Admin' AS lastmodifiedbyname_admin, "
    'sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    'sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
    sSQL = sSQL & " u3.firstname + ' ' + u3.lastname AS approvedeniedbyname, "
    sSQL = sSQL & " dm.latitude, "
    sSQL = sSQL & " dm.longitude "
    sSQL = sSQL & " FROM egov_dm_data dm "
    sSQL = sSQL &      " LEFT OUTER JOIN egov_users u ON u.userid = dm.createdbyid AND u.orgid = " & iorgid
    sSQL = sSQL &      " LEFT OUTER JOIN egov_users u2 ON u2.userid = dm.lastmodifiedbyid AND u.orgid = " & iorgid
    'sSQL = sSQL &      " LEFT OUTER JOIN users u ON dm.createdbyid = u.userid AND u.orgid = " & iorgid
    'sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dm.lastmodifiedbyid = u2.userid AND u2.orgid = " & iorgid
    sSQL = sSQL &      " LEFT OUTER JOIN users u3 ON u3.userid = dm.approvedeniedbyid AND u3.orgid = " & iorgid
    sSQL = sSQL &      " INNER JOIN egov_dm_types dmt ON dmt.dm_typeid = dm.dm_typeid "
    sSQL = sSQL & " WHERE dm.dmid = " & lcl_dmid

    set oDMData = Server.CreateObject("ADODB.Recordset")
    oDMData.Open sSQL, Application("DSN"), 3, 1

    if not oDMData.eof then
       lcl_dmid                = oDMData("dmid")
       lcl_dm_typeid           = oDMData("dm_typeid")
       lcl_description         = oDMData("description")
       lcl_orgid               = oDMData("orgid")
       lcl_categoryid          = oDMData("categoryid")
       lcl_createdbyid         = oDMData("createdbyid")
       lcl_createdbydate       = oDMData("createdbydate")
       lcl_lastmodifiedbyid    = oDMData("lastmodifiedbyid")
       lcl_lastmodifiedbydate  = oDMData("lastmodifiedbydate")
       lcl_approvedeniedbyid   = oDMData("approvedeniedbyid")
       lcl_approvedeniedbydate = oDMData("approvedeniedbydate")
       lcl_approvedeniedbyname = oDMData("approvedeniedbyname")
       lcl_isApproved_dmdata   = oDMData("isApproved")
       lcl_isActive            = oDMData("isActive")
       lcl_enableOwnerMaint    = oDMData("enableOwnerMaint")
       lcl_dm_latitude         = oDMData("latitude")
       lcl_dm_longitude        = oDMData("longitude")

      'Determine if the checkbox(es) are checked or not
       if oDMData("isActive") then
          lcl_isActive_value   = "Y"
          lcl_status_display   = "ACTIVE"
          lcl_status_button    = "Inactivate"
       else
          lcl_checked_isactive = ""
          lcl_isActive_value   = "N"
          lcl_status_display   = "INACTIVE"
          lcl_status_button    = "Activate"
       end if

      'Determine if this has been created/last modidified by a citizen user or an admin user
       if oDMData("isCreatedByAdmin") then
          lcl_isCreatedByAdmin = oDMData("isCreatedByAdmin")
       end if

       if oDMData("isLastModifiedByAdmin") then
          lcl_isLastModifiedByAdmin = oDMData("isLastModifiedByAdmin")
       end if

       if lcl_isCreatedByAdmin then
          lcl_displayCreatedByInfo = setupUserMaintLogInfo(oDMData("createdbyname_admin"), oDMData("createdbydate"))
       else
          lcl_displayCreatedByInfo = setupUserMaintLogInfo(oDMData("createdbyname_citizen"), oDMData("createdbydate"))
       end if

       if lcl_isLastModifiedByAdmin then
          lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(oDMData("lastmodifiedbyname_admin"), oDMData("lastmodifiedbydate"))
       else
          lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(oDMData("lastmodifiedbyname_citizen"), oDMData("lastmodifiedbydate"))
       end if

    else

       lcl_dmid_exists = checkDMIDExists(iOrgID, lcl_dmid)

       if not lcl_dmid_exists then
          lcl_add_params = setupUrlParameters(lcl_url_parameters, "success", "NE")
          response.redirect "datamgr.asp" & lcl_add_params
       end if

       if lcl_success <> "" AND lcl_success <> "NE" then
          lcl_add_params = setupUrlParameters(lcl_url_parameters, "success", lcl_success)
       else
          lcl_add_params = setupUrlParameters(lcl_url_parameters, "success", "NE")
       end if
    end if

    oDMData.close
    set oDMData = nothing
 end if

'Format the created/last modified/approved-denied by info
 'lcl_displayCreatedByInfo        = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 'lcl_displayLastModifiedByInfo   = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)
 lcl_displayApprovedDeniedByInfo = setupUserMaintLogInfo(lcl_approvedeniedbyname, lcl_approvedeniedbydate)

 if lcl_isApproved_dmdata <> "" then
    if lcl_isApproved_dmdata then
       lcl_approvedDeniedLabel = "Approved By:"
    else
       lcl_approvedDeniedLabel = "Denied By:"
    end if
 end if

'Get the layoutid
 lcl_layoutid = getDMTLayoutID(lcl_dm_typeid)

'If the "large address" feature is turned on then enable/disable the "Import Address Fields" button
 'if lcl_orghasfeature_large_address_list then
 '   lcl_onload = lcl_onload & "checkAddressButtons();"
 'end if

'Get the userid if available
	if request.cookies("userid") <> "" and request.cookies("userid") <> "-1" then
		  lcl_cookie_userid = request.cookies("userid")
 else
    session("RedirectPage") = request.servervariables("script_name") & "?" & request.querystring()
    response.redirect "../user_login.asp"
 end if

 if lcl_screen_mode = "EDIT" then
    if lcl_dm_latitude <> "" AND lcl_dm_longitude <> "" then
       lcl_onload = lcl_onload & "initialize(" & lcl_dm_latitude & "," & lcl_dm_longitude & ");"
    end if
 end if

 dim lcl_scripts

'Get the owner information
 if lcl_enableOwnerMaint then
    getDMOwnerEditorInfo lcl_dmid, _
                         lcl_cookie_userid, _
                         lcl_ownerid, _
                         lcl_ownertype, _
                         lcl_isOwner, _
                         lcl_isApproved, _
                         lcl_isWaitingApproval

    lcl_ownername = getOwnerName(lcl_ownerid)

   'Verify that the user signed in IS the owner.  If not then redirect them to the main listing page.
    lcl_cannotview_maint = 0

    if lcl_cookie_userid <> "" and lcl_ownerid <> "" then
       if clng(lcl_cookie_userid) <> clng(lcl_ownerid) then
          lcl_cannotview_maint = lcl_cannotview_maint + 1
       else
          if not lcl_isApproved then
             lcl_cannotview_maint = lcl_cannotview_maint + 1
          end if
       end if
    else
       lcl_cannotview_maint = lcl_cannotview_maint + 1
    end if

   'Determine if the user is to be redirectred or not
    if lcl_cannotview_maint > 0 then
       response.redirect "datamgr.asp" & lcl_url_parameters
    end if
 end if
%>
<html>
<head>
  <title>E-Gov Administration Console {DM Data - <%=lcl_screen_mode%>}</title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
<!--  <link rel="stylesheet" type="text/css" href="mapstyle.css" /> -->
<!--  <link rel="stylesheet" type="text/css" href="layout_styles.css" /> -->

 	<script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/easyform.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/setfocus.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

  <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no" />
  <script type="text/javascript" src="https://maps.google.com/maps/api/js?sensor=false"></script>

<script language="javascript">
$(document).ready(function(){
  $('#helpFeature_saveCategory_text').hide();
  $('#helpFeature_addSubCategory_text').hide();

  $('#helpFeature_addSubCategory').click(function() {
     $('#helpFeature_addSubCategory_text').toggle('slow');
  });

  $('#helpFeature_saveCategory').click(function() {
     $('#helpFeature_saveCategory_text').toggle('slow');
  });


  $.fn.enableDisableButton = function(iCompareField1, iCompareField2, iButtonID, iEvent, iAction1, iAction2) { 
    if($(iCompareField1).val() != $(iCompareField2).val()) {
       //$(iButtonID).attr('disabled',false);
       //$(iButtonID).attr(iEvent,iAction1);
       $(iButtonID).prop(iEvent,iAction1);
    } else {
       //$(iButtonID).attr('disabled',true);
       //$(iButtonID).attr(iEvent,iAction2);
       $(iButtonID).prop(iEvent,iAction2);
    }
  }

<% if lcl_screen_mode = "EDIT" then %>

  //BEGIN: Activate/Inactivate Button -----------------------------------------
  $('#activeButton').click(function() {
    var lcl_isActive = '';

    lcl_isActive = $('#isActive').val();

    if(lcl_isActive == 'N') {
       $('#isActive').val('Y');

       $.post('send_dm_email.asp', {
         orgid:  '<%=iOrgID%>',
         userid: '<%=lcl_cookie_userid%>',
         dmid:   '<%=lcl_dmid%>',
         action: 'REQUEST_DM_APPROVAL',
         isAjax: 'Y'
      }, function(result) {
         saveDMChanges();
      });

    } else {
       $('#isActive').val('N');

       saveDMChanges();
    }
  });
  //END: Activate/Inactivate Button -------------------------------------------

  //BEGIN: Change Category Button ---------------------------------------------
  $('#changeCategoryButton').click(function() {
    $('#changeCategoryButton').removeAttr('style','');
    $('#helpFeature_saveCategory_text').hide();
    saveDMChanges();
  });

  $('#categoryid').enableDisableButton('#categoryid', '#original_categoryid', '#changeCategoryButton', 'disabled', false, true);

  //enable/disable "Change Category" button
  $('#categoryid').change(function() {
    clearMsg('subCategoryAddButton');
    $('#helpFeature_saveCategory_text').hide();

    $('#categoryid').enableDisableButton('#categoryid', '#original_categoryid', '#changeCategoryButton', 'disabled', false, true);

    //We need to ensure that the data in the sub-categories section are current with the (parent) category selected in 
    //the "category" dropdown list.  This means that the user may have opened the sub-category list under one (parent) category 
    //and if this function is being executed it means that they are attempting to change that value and we need to make sure 
    //that a sub-category list for one category is NOT available for another category.
    $('#categoryid').enableDisableButton('#categoryid', '#original_categoryid', '#subCategorySelectButton', 'disabled', true, false);
    $('#categoryid').enableDisableButton('#categoryid', '#original_categoryid', '#subCategorySaveButton', 'disabled', true, true);
    $('#sub_sc_categoryname').val('');
    $('#subCategoryDIV').hide('slow');

    if(! $('#changeCategoryButton').prop('disabled')) {
       $('#changeCategoryButton').css({'height':'16pt','background-color':'#A80000','color':'#ffffff','font-weight':'bold','background-color':'#ff0000','cursor':'pointer'});
    } else {
       $('#changeCategoryButton').removeAttr('style','');
    }

  });

<% else %>
  $('#changeCategoryButton').hide();
<% end if %>
  //END: Change Category Button -----------------------------------------------

  //BEGIN: Owners/Editors Button ----------------------------------------------
  //$('#ownersEditorsDIV').hide('fast');

  //Owners/Editors - Maintain Button: Click
  //$('#maintainDMOwners').click(function() {
  //   $('#ownersEditorsDIV').toggle('slow');
  //});

  //BEGIN: Sub-Categories Buttons ---------------------------------------------
  $('#subCategoryDIV').hide('fast');
  //$('#subCategorySaveButton').attr('disabled','disabled');
  $('#subCategorySaveButton').prop('disabled','disabled');

  //Sub-Categories - Select Button: Click
  $('#subCategorySelectButton').click(function() {
     $('#subCategoryDIV').show('slow');
     $('#subCategorySelectButton').prop('disabled','disabled');
     $('#subCategorySaveButton').prop('disabled','');
     $('#sub_sc_categoryname').prop('disabled','');
     $('#sub_searchButton').prop('disabled','');

     //Build the sub-category list
     var lcl_categoryid = $('#categoryid').val();
     var lcl_dm_typeid  = $('#dm_typeid').val();
     var lcl_dmid       = $('#dmid').val();

//alert('left off here - datamgr_maint.asp line 377 calling build_subcategory_list.asp.  need to do a UNION to pull all approved subcategories and all those subcategories that were created for this DMID, but not yet approved.');

     $.post('build_subcategory_list.asp', {
        userid:           '<%=lcl_cookie_userid%>',
        orgid:            '<%=iorgid%>',
        dm_typeid:        lcl_dm_typeid,
        dmid:             lcl_dmid,
        categoryid:       lcl_categoryid,
        sub_categoryid:   '',
        sub_categoryname: '',
        useraction:       'DISPLAY',
        isAjax:           'Y'
     }, function(result) {
        $('#subCategoryList').html(result);
     });
  });

  //Sub-Categories - Save Button: Click
  $('#subCategorySaveButton').click(function() {
     clearMsg('subCategoryAddButton');
     $('#helpFeature_addSubCategory_text').hide();
     $('#subCategoryDIV').hide('slow');
     $('#subCategorySelectButton').prop('disabled','');
     $('#subCategorySaveButton').prop('disabled','disabled');

     var lcl_total_subcategories     = document.getElementById('total_subcategories').value;
     var lcl_total_subcategories_new = document.getElementById('total_subcategories_new').value;
     var i = 0;
     var n = 0;

     //BEGIN: Loop through existing/new sub-categories ------------------------
     var lcl_sub_categoryids_assigned   = '';
     var lcl_sub_categoryids_unassigned = '';

     //Determine which sub-categories are to be assigned/unassigned
     //-- Existing Sub-Categories ---------------------------------------------
     for(i = 1; i <= lcl_total_subcategories; i++) {
        var lcl_cid            = document.getElementById('subcategoryid' + i);
        var lcl_sub_categoryid = lcl_cid.value;

        if(lcl_cid.checked) {
           if(lcl_sub_categoryids_assigned != '') {
              lcl_sub_categoryids_assigned = lcl_sub_categoryids_assigned + ',' + lcl_sub_categoryid;
           } else {
              lcl_sub_categoryids_assigned = lcl_sub_categoryid;
           }
        } else {
           if(lcl_sub_categoryids_unassigned != '') {
              lcl_sub_categoryids_unassigned = lcl_sub_categoryids_unassigned + ',' + lcl_sub_categoryid;
           } else {
              lcl_sub_categoryids_unassigned = lcl_sub_categoryid;
           }
        }
     }

     //Determine which sub-categories are to be assigned
     //   *** do not need to worry about unassigning them, they are new so they haven't been assigned yet.
     //-- New Sub-Categories --------------------------------------------------
     for(n = 1; n <= lcl_total_subcategories_new; n++) {
        var lcl_cid_new            = document.getElementById('new_subcategoryid' + n);
        var lcl_sub_categoryid_new = lcl_cid_new.value;

       //Determine which sub-categories are to be assigned/unassigned
        if(lcl_cid_new.checked) {
           if(lcl_sub_categoryids_assigned != '') {
              lcl_sub_categoryids_assigned = lcl_sub_categoryids_assigned + ',' + lcl_sub_categoryid_new;
           } else {
              lcl_sub_categoryids_assigned = lcl_sub_categoryid_new;
           }
        //} else {
        //   if(lcl_sub_categoryids_unassigned != '') {
        //      lcl_sub_categoryids_unassigned = lcl_sub_categoryids_unassigned + ',' + lcl_sub_categoryid;
        //   } else {
        //      lcl_sub_categoryids_unassigned = lcl_sub_categoryid;
        //   }
        }
     }

     //Assign the Sub-Categories
     if(lcl_sub_categoryids_assigned != '') {
        maintainSubCategoryAssignments('ASSIGN', lcl_sub_categoryids_assigned)
     }

     //Unassign the Sub-Categories
     if(lcl_sub_categoryids_unassigned != '') {
        maintainSubCategoryAssignments('UNASSIGN', lcl_sub_categoryids_unassigned)
     }
     //END: Loop through existing/new sub-categories --------------------------

     //BEGIN: Check for new sub-categories and notify admin if needed
     if(lcl_total_subcategories_new > 0) {
        $.post('send_dm_email.asp', {
          orgid:  '<%=iOrgID%>',
          userid: '<%=lcl_cookie_userid%>',
          dmid:   '<%=lcl_dmid%>',
          action: 'NEW_SUBCATEGORY',
          isAjax: 'Y'
       }, function(result) {
       });
     }
     //END: Check for new sub-categories and notify admin if needed
  });

  //Sub-Categories - Search Button: Click
  $('#sub_searchButton').click(function() {
    var lcl_searchvalue = $('#sub_sc_categoryname').val();
    var lcl_foundCount  = 0;

    //Hide all of the TDs
    $('.subCategoryCell').each(function() {

      //Get the "id" for the current <TD> in the loop
      //var lcl_columnid = $(this).attr("id");
      var lcl_columnid = $(this).prop("id");

      //Once we have the "columnid" we need only the column number
      var lcl_id = lcl_columnid.replace("subcategorycell","");

      //Get the categoryname so we can perform the search
      var lcl_value   = $('#subcategoryname' + lcl_id).html();
      var lcl_showRow = false;

      //Compare the value in the current row in the loop to the search value
      if(lcl_searchvalue == '') {
         lcl_showRow = true;
      } else {
         lcl_value       = lcl_value.toUpperCase();
         lcl_searchvalue = lcl_searchvalue.toUpperCase();

         if(lcl_value.indexOf(lcl_searchvalue) > -1) {
            lcl_showRow = true;
         }
      }

      //Determine if we are showing/hiding the current row.
      if(lcl_showRow) {
         lcl_foundCount = lcl_foundCount + 1;
         $('#' + lcl_columnid).show("slow");
      } else {
         $('#' + lcl_columnid).hide("slow");
      }
    });
  });
  //END: Sub-Categories Buttons -----------------------------------------------

<% if lcl_enableOwnerMaint then %>
  $('#myDataMgrButton').click(function() {
    var lcl_url = '';

    lcl_url += 'mydatamgr.asp';
    lcl_url += '?f=<%=lcl_feature%>';

    location.href = lcl_url;
  });
<% end if %>
});

function addSubCategory() {
  var lcl_new_value = document.getElementById('subcategory_add').value;

  if(lcl_new_value == '') {
     inlineMsg(document.getElementById('subCategoryAddButton').id,'<strong>Required Field Missing: </strong>Other',10,'subCategoryAddButton');
     document.getElementById('subcategory_add').focus();
     return false;
  } else {
     clearMsg('subCategoryAddButton');

     document.getElementById('sub_sc_categoryname').disabled = true;
     document.getElementById('sub_searchButton').disabled    = true;

     //Check to see if the sub-category already exists on the DM TypeID
     //If "no" then create the sub-category.
     //Once the sub-category is created, add the HTML for the new sub-category option to the screen.
     var lcl_dm_typeid  = document.getElementById("dm_typeid").value;
     var lcl_dmid       = document.getElementById("dmid").value;
     var lcl_categoryid = document.getElementById("categoryid").value;

     $.post('build_subcategory_list.asp', {
        userid:           '<%=lcl_cookie_userid%>',
        orgid:            '<%=iorgid%>',
        dm_typeid:        lcl_dm_typeid,
        dmid:             lcl_dmid,
        categoryid:       lcl_categoryid,
        sub_categoryid:   '0',
        sub_categoryname: lcl_new_value,
        useraction:       'ADD',
        isAjax:           'Y'
     }, function(result) {
        var totalvalues = Number(document.getElementById("total_subcategories_new").value);

        if(result == "already exists") {
           inlineMsg(document.getElementById('subCategoryAddButton').id,'<strong>Duplicate Value: </strong> "' + lcl_new_value + '" already exists on a category.',10,'subCategoryAddButton');

           document.getElementById('sub_sc_categoryname').disabled = false;
           document.getElementById('sub_searchButton').disabled    = false;
        } else {
           //Set up the new row if it does NOT already exist
           //if(result != "already exists") {
           var mytbl              = document.getElementById('subCategoriesList');
           mytbl                  = mytbl.insertRow(0);
           mytbl.style.background = '#efefef';

           //Increase the total rows by one.  This is index for the new row.
           totalvalues = totalvalues + 1;

           //Build the cell for the new row.
           var a = mytbl.insertCell(0);  //CategoryID and CategoryName

           var lcl_sc_html = '';
           lcl_sc_html += '<input type="checkbox" name="new_subcategoryid' + totalvalues + '" id="new_subcategoryid'   + totalvalues + '" value="' + result + '" checked=""checked"" />';
           lcl_sc_html += '<input type="hidden" name="new_subcategoryname' + totalvalues + '" id="new_subcategoryname' + totalvalues + '" value="' + lcl_new_value + '" size="20" maxlength="100" />';
           lcl_sc_html += '<input type="hidden" name="isNewSubCategory'    + totalvalues + '" id="isNewSubCategory'    + totalvalues + '" value="Y" size="1" maxlength="1" />';
           lcl_sc_html += '<span class="redText">New Sub-Category: </span>' + lcl_new_value + '&nbsp;';
           lcl_sc_html += '<span class="redText">[waiting for approval]</span>';

           a.id        = 'subcategorycell' + totalvalues;
           a.className = 'subCategoryCell';
           a.colSpan   = 3;
           a.innerHTML = lcl_sc_html;

           //Clean up
           document.getElementById('subcategory_add').value = '';
           document.getElementById('total_subcategories_new').value = totalvalues;
           //}
        }
     });

     //Clean up
     document.getElementById('subcategory_add').focus();
  }
}

function maintainSubCategoryAssignments(iUserAction, iSubCategoryIDs) {
  var lcl_dm_typeid  = document.getElementById("dm_typeid").value;
  var lcl_dmid       = document.getElementById("dmid").value;
  var lcl_categoryid = document.getElementById("categoryid").value;

  $.post('build_subcategory_list.asp', {
     userid:           '<%=lcl_cookie_userid%>',
     orgid:            '<%=iorgid%>',
     dm_typeid:        lcl_dm_typeid,
     dmid:             lcl_dmid,
     categoryid:       lcl_categoryid,
     sub_categoryid:   iSubCategoryIDs,
     sub_categoryname: '',
     useraction:       iUserAction,
     isAjax:           'Y'
//  }, function() {
  });
}

function saveDMChanges() {
  clearMsg();
  $('#user_action').val('UPDATE');
  //$('#datamgr_maint').attr('action','datamgr_action.asp');
  $('#datamgr_maint').prop('action','datamgr_action.asp');
  $('#datamgr_maint').submit();
}

<% if lcl_enableOwnerMaint then %>
function approveDenyOwnerEditor(iOwnerType, iRowID, iDMOwnerID, iAction) {
  var lcl_ownertype  = 'EDITOR';
  var lcl_dm_ownerid = '';
  var lcl_action     = 'DENIED';

  if(iAction != '') {
     lcl_action = iAction;

     if(iOwnerType != '') {
        lcl_ownertype = iOwnerType;
     }

     if(iDMOwnerID != '') {
        lcl_dm_ownerid = iDMOwnerID;
     }

     //Approve/Deny the Owner/Editor
     $.post('approveDenyOwnerEditor.asp', {
        orgid:           '<%=lcl_orgid%>',
        userid:          '<%=lcl_cookie_userid%>',
        dm_ownerid:      lcl_dm_ownerid,
        approval_action: lcl_action,
        isAjax:          'Y'
     }, function(result) {
        var lcl_display_status = '';
        var lcl_display_info   = '';
        var lcl_status_value   = '';
        var lcl_button_nameid  = '';
        var lcl_button_value   = '';
        var lcl_button_action  = '';
        var lcl_button         = '';

        if(result == 'approved') {
           lcl_status_value  = 'APPROVED';
           lcl_button_nameid = lcl_ownertype + '_' + 'denyButton' + iRowID;
           lcl_button_value  = 'Deny';
           lcl_button_action = 'DENIED';
        } else {
           lcl_status_value  = 'DENIED';
           lcl_button_nameid = lcl_ownertype + '_' + 'approveButton' + iRowID;
           lcl_button_value  = 'Approve';
           lcl_button_action = 'APPROVED';
        }

        //Build the approve/deny info
        lcl_display_status  = '<span class="redText">' + lcl_status_value + '</span><br />';

        //Build the approve/deny button
        lcl_button  = "<input ";
        lcl_button +=   "type='button' ";
        lcl_button +=   "class='button' ";
        lcl_button +=   "name='"  + lcl_button_nameid + "' ";
        lcl_button +=   "id='"    + lcl_button_nameid + "' ";
        lcl_button +=   "value='" + lcl_button_value  + "' ";
        lcl_button +=   "onclick=\"approveDenyOwnerEditor('" + lcl_ownertype + "','" + iRowID + "','" + lcl_dm_ownerid + "','" + lcl_button_action + "');\" ";
        lcl_button += "/>";

        $('#' + lcl_ownertype + '_approvedDeniedStatus'  + iRowID).html(lcl_display_status);
        $('#' + lcl_ownertype + '_approvedDeniedInfo'    + iRowID).html(lcl_display_info);
        $('#' + lcl_ownertype + '_approvedDeniedButtons' + iRowID).html(lcl_button);
     });
  }
}

function changeOwnerType(iDMOwnerID, iChangeType) {
  var lcl_dm_ownerid = '0';
  var lcl_changetype = 'EDITOR';

  if(lcl_dm_ownerid != '') {
     lcl_dm_ownerid = iDMOwnerID;
  }

  if(iChangeType != '') {
     lcl_changetype = iChangeType;
  }

  $.post('changeOwnerType.asp', {
     orgid:      '<%=lcl_orgid%>',
     userid:     '<%=lcl_cookie_userid%>',
     dm_ownerid: lcl_dm_ownerid,
     changetype: lcl_changetype,
     isAjax:     'Y'
  }, function(result) {
     if(result != '') {
        location.href = 'datamgr_maint.asp?f=<%=lcl_feature%>&dmid=<%=lcl_dmid%>&success=' + result;
     }
  });
}
<% end if %>

function confirmDelete() {
  //var r = confirm('Are you sure you want to delete the "' + document.getElementById("title").value + '" blog entry?  \r NOTE: Any/All comments will be deleted as well.');
//  var r = confirm('Are you sure you want to delete: "' + document.getElementById("description").value + '"');
  var r = confirm('Are you sure you want to delete this <%=lcl_featureNameLabel%>?');
  if (r==true) {

    <%
      lcl_delete_params = lcl_url_parameters
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "user_action", "DELETE")
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "dmid", lcl_dmid)
    %>
      location.href="datamgr_action.asp<%=lcl_delete_params%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count = 0;

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("datamgr_maint").submit();
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

function openWin(p_url, p_width, p_height) {
  w = 600;
  h = 400;

  if((p_width!="")&&(p_width!=undefined)) {
      w = p_width;
  }

  if((p_height!="")&&(p_height!=undefined)) {
      h = p_height;
  }

  l = (screen.availWidth/2)-(w/2);
  t = (screen.availHeight/2)-(h/2);

  lcl_url = p_url;

  eval('window.open("' + lcl_url + '", "_blank", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1")');
}

function editSection(iSectionID) {
  var lcl_url  = 'maintain_dmt_section.asp';
      lcl_url += '?sectionid=' + iSectionID;
      lcl_url += '&dmid=<%=lcl_dmid%>';
      lcl_url += '&f=<%=lcl_feature%>';

  openWin(lcl_url, 1000, 600);
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

function initialize(iLat, iLng) {
  var myLatLng  = new google.maps.LatLng(iLat, iLng);
  var myOptions = {
     mapTypeId: google.maps.MapTypeId.ROADMAP,  //maptypes: ROADMAP, SATELLITE, HYBRID, TERRAIN
     zoom:      17,
     center:    myLatLng
     //streetViewControl: false
  }

  //Create the mappoint.  This "addMarker" is slightly different than what is on the 
  //datamagr_list.asp as here we are only building a single mappoint and do not need
  //to search the array as we already have the latitude and longitude.
  if(document.getElementById("map_canvas_dot")) {
     var mapDot = new google.maps.Map(document.getElementById("map_canvas_dot"), myOptions);
     addMarker(mapDot,myLatLng);
  }

  //Create the street view version of the map
  if(document.getElementById("map_canvas_streetview")) {
     var panorama;
     var mapStreetView = new google.maps.Map(document.getElementById("map_canvas_streetview"), myOptions);

     addMarker(mapStreetView,myLatLng);

     panorama = mapStreetView.getStreetView();
     panorama.setPosition(myLatLng);
     panorama.setPov({
        heading: 265,
        pitch:   5,
        zoom:    1
     });
     panorama.setVisible(true);
  }
}

function addMarker(iMap,iLatLng) {
  //var image = "[filename goes here]"
  new google.maps.Marker({
     position:  iLatLng,
     map:       iMap,
     draggable: false,
     animation: google.maps.Animation.DROP
     //icon:      image,
     //title:     "here" + iRowCount
  });
}
</script>

<style type="text/css">
  body       { background-color: #efefef; }
  #screenMsg { color: #ff0000; }
</style>

</head>
<!--#include file="../include_top.asp"-->
<%
  RegisteredUserDisplay(sLevel)

  response.write "  <form name=""datamgr_maint"" id=""datamgr_maint"" method=""post"" action=""datamgr_action.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" size=""4"" maxlength=""20""  name=""user_action"" id=""user_action"" value="""" />" & vbcrlf
  response.write "    <input type=""hidden"" size=""4"" maxlength=""10""  name=""orgid"" id=""orgid"" value="""               & iorgid             & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""4"" maxlength=""10""  name=""screen_mode"" id=""screen_mode"" value="""   & lcl_screen_mode    & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""5"" maxlength=""500"" name=""f"" id=""f"" value="""                       & lcl_feature        & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""5"" maxlength=""10""  name=""dmid"" id=""dmid"" value="""                 & lcl_dmid           & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""3"" maxlength=""20""  name=""dm_typeid"" id=""dm_typeid"" value="""       & lcl_dm_typeid      & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""3"" maxlength=""100"" name=""sc_dm_typeid"" id=""sc_dm_typeid"" value=""" & lcl_sc_dm_typeid   & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""3"" maxlength=""20""  name=""u"" id=""u"" value="""                       & lcl_cookie_userid  & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""1"" maxlength=""10""  name=""isActive"" id=""isActive"" value="""         & lcl_isActive_value & """ />" & vbcrlf
  response.write "    <input type=""hidden"" size=""5"" maxlength=""10""  name=""original_categoryid"" id=""original_categoryid"" value=""" & lcl_categoryid & """ />" & vbcrlf
  response.write "<div id=""centercontent"" class=""datamgr"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""max-width:800px;"" class=""start"">" & vbcrlf

 'Determine if the DM Type is pre-selected or not.  If yes then show the dropdown list.
 'Otherwise, get the "dm_typeid" based off of the feature passed in.
 'NOTE: "lcl_canSelectDMType", if set to True, represent BOTH the display and selecting of the DM Type.
        'This means that in the "EDIT" screen if "lcl_canSelectDMType" is true then the value of the dm_typeid is displayed.
        'If the screen mode is "ADD" and "lcl_canSelectDMType" is true then the dropdown list of DM Types is display.
  if lcl_canSelectDMType then
     response.write "  <tr valign=""bottom"">" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='datamgr.asp" & lcl_url_parameters & "';"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td align=""right"">" & vbcrlf
     response.write "          <span id=""screenMsg""></span><br />" & vbcrlf

     if lcl_enableOwnerMaint then
        response.write "       <input type=""button"" name=""myDataMgrButton"" id=""myDataMgrButton"" class=""button"" value=""My " & lcl_description & """ />" & vbcrlf
     end if

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td colspan=""2"" class=""datamgrContent"">" & vbcrlf
     response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"">" & vbcrlf
'     response.write "            <tr valign=""top"">" & vbcrlf
'     response.write "                <td>" & vbcrlf
'     response.write "                    <input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='datamgr.asp" & lcl_url_parameters & "';"" />" & vbcrlf
'     response.write "                </td>" & vbcrlf
'     response.write "                <td align=""right"">" & vbcrlf
'     response.write "                    <input type=""button"" name=""myDataMgrButton"" id=""myDataMgrButton"" class=""button"" value=""My " & lcl_description & """ />" & vbcrlf
'     response.write "                </td>" & vbcrlf
'     response.write "            </tr>" & vbcrlf

'     response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf

     lcl_dm_welcome_msg = "&nbsp;"

     if lcl_enableOwnerMaint then
        lcl_dm_welcome_msg = "Welcome <span class=""redText"">" & lcl_ownername & "</span>" & vbcrlf
     end if

     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td>" & lcl_dm_welcome_msg & "</td>" & vbcrlf
     response.write "                <td align=""right"">" & vbcrlf
     response.write "                    Status: <span class=""redText"">" & lcl_status_display & "</span>" & vbcrlf
     response.write "                    <input type=""button"" name=""activeButton"" id=""activeButton"" value=""" & lcl_status_button & """ class=""button"" />" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td colspan=""3"">&nbsp;</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "</table>" & vbcrlf
  	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""max-width:800px;"" class=""start"">" & vbcrlf

     if lcl_screen_mode = "EDIT" then
        response.write "            <tr valign=""top"">" & vbcrlf

       'BEGIN: Account Info Section -------------------------------------------
        response.write "                <td class=""respCol"">" & vbcrlf
                                            lcl_dm_sectionid         = ""
                                            lcl_sectionid            = getAccountInfoSectionID(lcl_dm_typeid)
                                            lcl_sectionname          = ""
                                            lcl_sectiontype          = ""
                                            lcl_sectionIsActive      = ""
                                            lcl_isAccountInfoSection = True
                                            lcl_sectionlocation      = ""
                                            lcl_sectionorder         = ""
                                            lcl_totaldraggable_items = 0
                                            lcl_sectionmode          = "EDIT"

                                            displaySection_accountInfo lcl_dmid, lcl_dm_typeid, lcl_sectionid
        response.write "                </td>" & vbcrlf
       'END: Account Info Section ---------------------------------------------

       'BEGIN: Category/Sub-Category Section ----------------------------------
        response.write "                <td nowrap=""nowrap"" class=""respCol"">" & vbcrlf
        response.write "                    <fieldset class=""accountinfo"">" & vbcrlf
        response.write "                      <legend>Category Info&nbsp;</legend>" & vbcrlf
        response.write "                      <p>" & vbcrlf
        response.write "                        Category:" & vbcrlf
        response.write "                        <select name=""categoryid"" id=""categoryid"">" & vbcrlf
                                                  displayDMTCategories iorgid, lcl_dm_typeid, lcl_parent_categoryid, lcl_categoryid
        response.write "                        </select>" & vbcrlf
        response.write "                        <input type=""button"" name=""changeCategoryButton"" id=""changeCategoryButton"" value=""Save Category Change"" />" & vbcrlf
        response.write "                        <img src=""../images/help.jpg"" name=""helpFeature_saveCategory"" id=""helpFeature_saveCategory"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
        response.write "                        <div name=""helpFeature_saveCategory_text"" id=""helpFeature_saveCategory_text"" class=""helpOptionText"">" & vbcrlf
        response.write "                          <strong>E-GOV TIP:</strong><br />Whenever selecting a new category from the dropdown list,<br />be sure to press the ""Save Category Change"" button to save the change." & vbcrlf
        response.write "                        </div>" & vbcrlf
        response.write "                      </p>" & vbcrlf
        response.write "                      <fieldset id=""fieldset_subcategories"">" & vbcrlf
        response.write "                        <legend id=""legend_subcategories"">Sub-Categories&nbsp;</legend>" & vbcrlf
        response.write "                        <input type=""button"" name=""subCategorySelectButton"" id=""subCategorySelectButton"" value=""View/Select Sub-Categories"" />" & vbcrlf
        response.write "                        <input type=""button"" name=""subCategorySaveButton"" id=""subCategorySaveButton"" value=""Save Sub-Category Changes"" />" & vbcrlf
        response.write "                        <div id=""subCategoryDIV"">" & vbcrlf
        response.write "                          <div id=""subcategory_instructions"">" & vbcrlf
        response.write "                          <p>" & vbcrlf
        response.write "                            <strong>Instructions: </strong><span class=""redText"">""Check"" all sub-categories that apply.<br />" & vbcrlf
        response.write "                            Click the ""Save Sub-Category Changes"" button to save your changes.</span>" & vbcrlf
        response.write "                          </p>" & vbcrlf
        response.write "                          </div>" & vbcrlf
        response.write "                          <div id=""subcategory_search"">" & vbcrlf
        response.write "                            <input type=""text"" name=""sub_sc_categoryname"" id=""sub_sc_categoryname"" size=""20"" maxlength=""100"" />" & vbcrlf
        response.write "                            <input type=""button"" name=""sub_searchButton"" id=""sub_searchButton"" value=""Search Sub-Categories"" /><br />" & vbcrlf
        response.write "                          </div>" & vbcrlf
        response.write "                          <span id=""subCategoryList""></span>" & vbcrlf
        response.write "                          <div id=""subCategoryAddRow"">" & vbcrlf
        response.write "                            Other: <input type=""text"" name=""subcategory_add"" id=""subcategory_add"" value="""" size=""20"" maxlength=""100"" onchange=""clearMsg('subCategoryAddButton');"" />" & vbcrlf
        response.write "                            <input type=""button"" name=""subCategoryAddButton"" id=""subCategoryAddButton"" value=""Add"" onclick=""addSubCategory();"" />" & vbcrlf
        response.write "                            <img src=""../images/help_graybg.jpg"" name=""helpFeature_addSubCategory"" id=""helpFeature_addSubCategory"" class=""helpOption"" alt=""Click for more info"" /><br />" & vbcrlf
        response.write "                            <div name=""helpFeature_addSubCategory_text"" id=""helpFeature_addSubCategory_text"" class=""helpOptionText"">" & vbcrlf
        response.write "                              <strong>E-GOV TIP:</strong><br />Clicking on the ""add"" button will add the sub-category but NOT automatically assign it." & vbcrlf
        response.write "                            </div>" & vbcrlf
        response.write "                          </div>" & vbcrlf
        response.write "                        </div>" & vbcrlf
        response.write "                      </fieldset>" & vbcrlf
        response.write "                    </fieldset>" & vbcrlf
        response.write "                </td>" & vbcrlf
       'END: Category/Sub-Category Section ------------------------------------

        response.write "            </tr>" & vbcrlf

       'BEGIN: History --------------------------------------------------------
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td colspan=""2"">" & vbcrlf
        response.write "                    <fieldset id=""fieldset_historylog"">" & vbcrlf
        response.write "                      <legend>History&nbsp;</legend>" & vbcrlf
        response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

       'Created By
        response.write "                        <tr>" & vbcrlf
        response.write "                            <td nowrap=""nowrap"">Created By:</td>" & vbcrlf
        response.write "                            <td class=""redText"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
        response.write "                        </tr>" & vbcrlf

       'Approved/Denied By
        response.write "                        <tr>" & vbcrlf

        if lcl_isApproved_dmdata <> "" then
           response.write "                            <td nowrap=""nowrap"">" & lcl_approvedDeniedLabel & "</td>" & vbcrlf
           response.write "                            <td class=""redText"">" & lcl_displayApprovedDeniedByInfo & "</td>" & vbcrlf
        else
           response.write "                            <td nowrap=""nowrap"">Approval Status:</td>" & vbcrlf
           response.write "                            <td class=""redText"">[WAITING FOR APPROVAL]</td>" & vbcrlf
        end if

        response.write "                        </tr>" & vbcrlf

       'Last Modified By
        if lcl_displayLastModifiedByInfo <> "" then
           response.write "                        <tr>" & vbcrlf
           response.write "                            <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
           response.write "                            <td class=""redText"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
           response.write "                        </tr>" & vbcrlf
        end if

        response.write "                      </table>" & vbcrlf
        response.write "                    </fieldset>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
       'END: History ----------------------------------------------------------

       'BEGIN: Owner/Editor Section (Available ONLY to OWNERS) ----------------
        if lcl_enableOwnerMaint then
           if lcl_ownertype = "OWNER" then
              response.write "            <tr>" & vbcrlf
              response.write "                <td colspan=""2"">" & vbcrlf
                                                  displaySection_ownersInfo lcl_cookie_userid, lcl_orgid, lcl_dmid
              response.write "                </td>" & vbcrlf
              response.write "            </tr>" & vbcrlf
           end if
        end if
       'END: Owner/Editor Section ---------------------------------------------

     else
        response.write "            <input type=""hidden"" name=""mappointcolor"" id=""mappointcolor"" value=""" & lcl_mappointcolor & """ />" & vbcrlf
        response.write "            <tr><td colspan=""2""></td></tr>" & vbcrlf
     end if

     response.write "          </table>" & vbcrlf

    'BEGIN: Build the Layout --------------------------------------------------
     response.write "<fieldset class=""layoutfieldset"">" & vbcrlf
     response.write "  <legend>Current Layout</legend>" & vbcrlf

    'Retrieve any/all fields related to this Map-Point Type
     lcl_displayFieldsetLegend    = False
     lcl_displayFieldsetBorder    = False
     lcl_displayAvailableSections = False
     lcl_section_mode             = "EDIT"

     buildDMLayout lcl_layoutid, lcl_dm_typeid, lcl_dmid, lcl_displayFieldsetLegend, _
                   lcl_displayFieldsetBorder, lcl_displayAvailableSections, lcl_section_mode


    'Retrieve any/all fields related to this DM Type
     'displayDMTypesFields iorgid, lcl_dm_typeid

     response.write "</fieldset>" & vbcrlf
    'END: Build the Layout ----------------------------------------------------

     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  else
'     response.write "<input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" value=""" & lcl_dm_typeid & """ />" & vbcrlf
     response.write "<input type=""hidden"" name=""mappointcolor"" id=""mappointcolor"" value=""" & lcl_mappointcolor & """ />" & vbcrlf
     response.write "<span style=""color:#800000"">Processing...</span>" & vbcrlf
     response.write "<select name=""categoryid"" id=""categoryid"" style=""visibility:hidden"">" & vbcrlf
                       displayDMTCategories iorgid, lcl_dm_typeid, lcl_parent_categoryid, lcl_categoryid
     response.write "</select>" & vbcrlf
  end if

  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf

 'Determine if there are any scripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if
%>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub displaySection_accountInfo(iDMID, iDMTypeID, iSectionID)

 'Get the "account info" section and fields for this DMTypeID
  sSQL = "SELECT "
  sSQL = sSQL & " dmtf.dm_fieldid, "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQL = sSQL & " dmsf.fieldname, "
  sSQL = sSQL & " dmsf.fieldtype, "
  sSQL = sSQL & " dmv.dm_valueid, "
  sSQL = sSQL & " dmv.fieldvalue "
  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_values dmv "
  sSQL = sSQL &                      " ON dmv.dm_fieldid = dmtf.dm_fieldid "
  sSQL = sSQL &                      " AND dmv.dmid = " & iDMID
  sSQL = sSQL &                      " AND dmv.dm_typeid = " & iDMTypeID
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &                      " ON dmsf.section_fieldid = dmtf.section_fieldid "
  sSQL = sSQL &                      " AND dmsf.sectionid = " & iSectionID
  sSQL = sSQL &                      " AND dmsf.isActive = 1 "
  sSQL = sSQL & " WHERE dmtf.dm_sectionid IN (SELECT dmts.dm_sectionid "
  sSQL = sSQL &                             " FROM egov_dm_types_sections dmts "
  sSQL = sSQL &                             " WHERE dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL &                             " AND dmts.sectionid IN (SELECT dms.sectionid "
  sSQL = sSQL &                                                    " FROM egov_dm_sections dms "
  sSQL = sSQL &                                                    " WHERE dms.isAccountInfoSection = 1 "
  sSQL = sSQL &                                                    " AND dms.isActive = 1 "
  sSQL = sSQL &                                                    " AND dms.sectionid = " & iSectionID
  sSQL = sSQL &                                                    " ) "
  sSQL = sSQL &                            " ) "
  sSQL = sSQL & " ORDER BY dmtf.dm_sectionid, dmsf.displayOrder, dmtf.dm_fieldid "
'response.write sSQL
  set oDisplayAccountInfo = Server.CreateObject("ADODB.Recordset")
  oDisplayAccountInfo.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayAccountInfo.eof then
     iRowCount           = 0
     lcl_field_maxlength = "4000"

     response.write "<fieldset class=""accountinfo"">" & vbcrlf
     response.write "  <legend>Account Info&nbsp;</legend>" & vbcrlf
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf

     do while not oDisplayAccountInfo.eof
        iRowCount      = iRowCount + 1
        lcl_fieldvalue = oDisplayAccountInfo("fieldvalue")
        lcl_fieldname  = oDisplayAccountInfo("fieldname")
        lcl_fieldtype  = oDisplayAccountInfo("fieldtype")

        if instr(lcl_fieldtype,"WEBSITE") > 0 OR instr(lcl_fieldtype,"EMAIL") > 0 then
           lcl_fieldvalue = buildURLDisplayValue(lcl_fieldtype, lcl_fieldvalue)
        end if

        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <input type=""hidden"" name=""dm_valueid"   & iRowCount & """ id=""dm_valueid"   & iRowCount & """ value=""" & oDisplayAccountInfo("dm_valueid")   & """ />" & vbcrlf
        response.write "      <input type=""hidden"" name=""dm_fieldid"   & iRowCount & """ id=""dm_fieldid"   & iRowCount & """ value=""" & oDisplayAccountInfo("dm_fieldid")   & """ />" & vbcrlf
        response.write "      <input type=""hidden"" name=""dm_sectionid" & iRowCount & """ id=""dm_sectionid" & iRowCount & """ value=""" & oDisplayAccountInfo("dm_sectionid") & """ />" & vbcrlf

        if lcl_fieldname <> "" then
           response.write "      <td nowrap=""nowrap"" class=""accountInfo_fieldname"">" & lcl_fieldname & ":</td>" & vbcrlf
           response.write "      <td width=""100%"">" & lcl_fieldvalue & "</td>" & vbcrlf
        else
           response.write "      <td colspan=""2"">" & lcl_fieldvalue & "</td>" & vbcrlf
        end if

        response.write "  </tr>" & vbcrlf

        oDisplayAccountInfo.movenext
     loop

     response.write "</table>" & vbcrlf

     lcl_iMode                      = "ACCOUNTINFO_VIEW"
     lcl_enabledisable_button_label = 1

     displayButtonsSection lcl_iMode, iDMTypeID, iSectionID, lcl_enabledisable_button_label

     response.write "</fieldset>" & vbcrlf
  end if

  oDisplayAccountInfo.close
  set oDisplayAccountInfo = nothing

  'response.write "<input type=""text"" name=""totalfields"" id=""totalfields"" value=""" & iRowCount & """ size=""5"" maxlength=""10"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displaySection_ownersInfo(iCurrentUserID, iOrgID, iDMID)
  dim sCurrentUserID, sOrgID, sDMID
  dim lcl_userfname, lcl_userlname, lcl_userfullname, lcl_useremail

  sCurrentUserID = 0
  sOrgID         = 0
  sDMID          = 0

  if iCurrentUserID <> "" then
     sCurrentUserID = clng(iCurrentUserID)
  end if

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  response.write "<fieldset class=""ownerinfo"">" & vbcrlf
  'response.write "  <legend>Owner/Editor Info&nbsp;</legend>" & vbcrlf
  response.write "  <legend>Owner Info&nbsp;</legend>" & vbcrlf
  response.write "  <div class=""ownerInfoContainer"">" & vbcrlf
  response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
                      displayOwnerEditorInfoRow sCurrentUserID, sOrgID, sDMID, "OWNER"
                      displayOwnerEditorInfoRow sCurrentUserID, sOrgID, sDMID, "EDITOR"
  response.write "  </table>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub displayOwnerEditorInfoRow(iCurrentUserID, iOrgID, iDMID, iOwnerType)

  dim sCurrentUserID, sOrgID, sDMID, sOwnerType, sOwnerEditorLabel, sChangeToButtonLabel
  dim lcl_dm_ownerid, lcl_userfname, lcl_userlname, lcl_userfullname, lcl_useremail, lcl_isOwner
  dim lcl_rowcount, lcl_bgcolor, lcl_previous_ownertype, lcl_isApprovedDeniedByAdmin

  sCurrentUserID         = 0
  sOrgID                 = 0
  sDMID                  = 0
  sOwnerType             = "OWNER"
  sOwnerEditorLabel      = "Owners"
  sChangeToButtonLabel   = "EDITOR"
  lcl_rowcount           = 0
  lcl_bgcolor            = "#ffffff"
  lcl_previous_ownertype = ""

  if iCurrentUserID <> "" then
     sCurrentUserID = clng(iCurrentUserID)
  end if

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  if iOwnerType <> "" then
     sOwnerType = ucase(iOwnerType)
  end if

  if sOwnerType = "EDITOR" then
     sOwnerEditorLabel    = "Editors"
     sChangeToButtonLabel = "OWNER"
  end if

  getDMOwnerEditorInfo sDMID, _
                       sCurrentUserID, _
                       lcl_ownerid, _
                       lcl_ownertype, _
                       lcl_isOwner, _
                       lcl_isApproved, _
                       lcl_isWaitingApproval

  sSQL = "SELECT dmo.dm_ownerid, "
  sSQL = sSQL & " dmo.userid, "
  sSQL = sSQL & " u.userfname, "
  sSQL = sSQL & " u.userlname, "
  sSQL = sSQL & " u.useremail, "
  sSQL = sSQL & " dmo.ownertype, "
  sSQL = sSQL & " dmo.isApproved, "
  sSQL = sSQL & " dmo.isApprovedDeniedByAdmin, "
  sSQL = sSQL & " dmo.approvedeniedbyid, "
  sSQL = sSQL & " dmo.approvedeniedbydate, "
  sSQL = sSQL & " u2.userfname + ' ' + u2.userlname AS approvedeniedbyname_citizen, "
  sSQL = sSQL & " 'Approved By Admin' AS approvedeniedbyname_admin "
  sSQL = sSQL & " FROM egov_dm_owners dmo "
  sSQL = sSQL &      " INNER JOIN egov_users u ON u.userid = dmo.userid "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_users u2 ON u2.userid = dmo.approvedeniedbyid "
  sSQL = sSQL & " WHERE dmo.orgid = " & sOrgID
  sSQL = sSQL & " AND dmo.dmid = " & sDMID
  sSQL = sSQL & " AND upper(dmo.ownertype) = '" & dbsafe(sOwnerType) & "'"
  sSQL = sSQL & " ORDER BY dmo.userid "

  set oDisplayOwnerInfo = Server.CreateObject("ADODB.Recordset")
  oDisplayOwnerInfo.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayOwnerInfo.eof then
     do while not oDisplayOwnerInfo.eof
        lcl_rowcount = lcl_rowcount + 1
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#ffffff","#efefef")

        lcl_dm_ownerid   = ""
        lcl_userfname    = ""
        lcl_userlname    = ""
        lcl_userfullname = ""
        lcl_useremail    = ""

        if oDisplayOwnerInfo("dm_ownerid") <> "" then
           lcl_dm_ownerid = oDisplayOwnerInfo("dm_ownerid")
        end if

        if oDisplayOwnerInfo("userfname") <> "" then
           lcl_userfname = trim(oDisplayOwnerInfo("userfname"))
        end if

        if oDisplayOwnerInfo("userlname") <> "" then
           lcl_userlname = trim(oDisplayOwnerInfo("userlname"))
        end if

        if oDisplayOwnerInfo("useremail") <> "" then
           lcl_useremail = trim(oDisplayOwnerInfo("useremail"))
        end if

       'Build the full username
        if trim(lcl_userfname) <> "" then
           lcl_userfullname = trim(lcl_userfname)
        end if

        if trim(lcl_userlname) <> "" then
           if lcl_userfullname <> "" then
              lcl_userfullname = lcl_userfullname & " " & lcl_userlname
           else
              lcl_userfullname = lcl_userlname
           end if
        end if

       'Set up Approve/Deny Buttons for display
        lcl_isApprovedDeniedByAdmin    = False
        lcl_show_approvedButton        = 1
        lcl_show_deniedButton          = 1
        lcl_display_approvedDeniedInfo = ""
        lcl_approved_denied_status     = ""

        if oDisplayOwnerInfo("isApprovedDeniedByAdmin") then
           lcl_isApprovedDeniedByAdmin = oDisplayOwnerInfo("isApprovedDeniedByAdmin")
        end if

       'Determine if this has been approved by a citizen user or an admin user
        if lcl_isApprovedDeniedByAdmin then
           lcl_approvedenied_info = formatAdminActionsInfo(oDisplayOwnerInfo("approvedeniedbyname_admin"), oDisplayOwnerInfo("approvedeniedbydate"))
        else
           lcl_approvedenied_info = formatAdminActionsInfo(oDisplayOwnerInfo("approvedeniedbyname_citizen"), oDisplayOwnerInfo("approvedeniedbydate"))
        end if

        if lcl_approvedenied_info <> "" then
           if oDisplayOwnerInfo("isApproved") then
              lcl_show_approvedButton    = 0
              lcl_approved_denied_status = "APPROVED"
           else
              lcl_show_deniedButton      = 0
              lcl_approved_denied_status = "DENIED"
           end if
        else
           lcl_approved_denied_status = "WAITING FOR<br />APPROVAL"
           'lcl_waiting_count          = lcl_waiting_count + 1
        end if

        if lcl_previous_ownertype <> oDisplayOwnerInfo("ownertype") then
           lcl_bgcolor = "#eeeeee"

           response.write "  <tr>" & vbcrlf
           response.write "      <td id=""ownereditor_titlerow"">" & sOwnerEditorLabel & "</td>" & vbcrlf
           response.write "      <td id=""ownereditor_titlerow"">Email</td>" & vbcrlf
           response.write "      <td id=""ownereditor_titlerow_center"">Approval Status</td>" & vbcrlf
           response.write "      <td id=""ownereditor_titlerow_center"">Approved/Denied By</td>" & vbcrlf
           'response.write "      <td id=""ownereditor_titlerow"">Change To...</td>" & vbcrlf
           response.write "  </tr>" & vbcrlf
        end if

        response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "      <td nowrap=""nowrap"">" & lcl_userfullname & "</td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"">" & lcl_useremail    & "</td>" & vbcrlf
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""" & oDisplayOwnerInfo("ownertype") & "_approvedDeniedStatus" & lcl_rowcount & """ class=""redText"">" & lcl_approved_denied_status & "</span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf

        if lcl_isOwner then
           response.write "          <span id=""" & oDisplayOwnerInfo("ownertype") & "_approvedDeniedInfo" & lcl_rowcount & """>" & lcl_approvedenied_info & "</span><br />" & vbcrlf

           if sCurrentUserID <> oDisplayOwnerInfo("userid") then
              response.write "          <span id=""" & oDisplayOwnerInfo("ownertype") & "_approvedDeniedButtons" & lcl_rowcount & """>" & vbcrlf

              if lcl_show_approvedButton > 0 then
                 response.write "          <input type=""button"" name=""" & oDisplayOwnerInfo("ownertype") & "_approveButton" & lcl_rowcount & """ id=""" & oDisplayOwnerInfo("ownertype") & "_approveButton" & lcl_rowcount & """ class=""button"" value=""Approve"" onclick=""approveDenyOwnerEditor('" & oDisplayOwnerInfo("ownertype") & "', '" & lcl_rowcount & "', '" & lcl_dm_ownerid & "','APPROVED');"" />" & vbcrlf
              end if

              if lcl_show_deniedButton > 0 then
                 response.write "          <input type=""button"" name=""" & oDisplayOwnerInfo("ownertype") & "_denyButton" & lcl_rowcount & """ id=""" & oDisplayOwnerInfo("ownertype") & "_denyButton" & lcl_rowcount & """ class=""button"" value=""Deny"" onclick=""approveDenyOwnerEditor('" & oDisplayOwnerInfo("ownertype") & "', '" & lcl_rowcount & "', '" & lcl_dm_ownerid & "','DENIED');"" />" & vbcrlf
              end if

              response.write "          </span>" & vbcrlf
              'response.write "      </td>" & vbcrlf

              'response.write "      <td valign=""bottom"">" & vbcrlf
              'response.write "          <input type=""button"" name=""" & oDisplayOwnerInfo("ownertype") & "_changeOwnerTypeButton" & lcl_rowcount & """ id=""" & oDisplayOwnerInfo("ownertype") & "_changeOwnerTypeButton" & lcl_rowcount & """ class=""button"" value=""" & sChangeToButtonLabel & """ onclick=""changeOwnerType('" & oDisplayOwnerInfo("dm_ownerid") & "','" & sChangeToButtonLabel & "');"" />" & vbcrlf

           'else
           '   response.write "&nbsp;</td><td>&nbsp;" & vbcrlf
           end if
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        lcl_previous_ownertype = oDisplayOwnerInfo("ownertype")

        oDisplayOwnerInfo.movenext
     loop
  end if

  oDisplayOwnerInfo.close
  set oDisplayOwnerInfo = nothing

end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode, iFeatureNameLabel, iReturnParameters)

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

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  'response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='datamgr_list.asp" & iReturnParameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  'else
     'response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     'response.write "<input type=""button"" name=""previewButton"" id=""previewButton"" value=""Preview Site"" class=""button"" onclick=""alert('coming soon');"" />" & vbcrlf
     'response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add Another " & iFeatureNameLabel & """ class=""button"" onclick=""location.href='datamgr_maint.asp" & iReturnParameters & "'"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function checkIsOwner(iCurrentUserID, iOrgID, iDMID)

  dim sCurrentUserID, sOrgID, sDMID, lcl_return

  sCurrentUserID = 0
  sOrgID         = 0
  sDMID          = 0
  lcl_return     = false

  if iCurrentUserID <> "" then
     sCurrentUserID = clng(iCurrentUserID)
  end if

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDMID <> "" then
     sDMID = clng(iDMID)
  end if

  sSQL = "SELECT distinct 'Y' as isOwner "
  sSQL = sSQL & " FROM egov_dm_owners "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND dmid = " & sDMID
  sSQL = sSQL & " AND userid = " & sCurrentUserID
  sSQL = sSQL & " AND ownertype = 'OWNER' "

  set oCheckIsOwner = Server.CreateObject("ADODB.Recordset")
  oCheckIsOwner.Open sSQL, Application("DSN"), 3, 1

  if not oCheckIsOwner.eof then
     lcl_return = true
  end if

  oCheckIsOwner.close
  set oCheckIsOwner = nothing

  checkIsOwner = lcl_return

end function

'-----------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = p_value
  end if

  dbsafe = lcl_value

end function
%>
