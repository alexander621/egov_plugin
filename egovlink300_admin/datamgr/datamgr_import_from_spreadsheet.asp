<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_import_from_spreadsheet.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This screen walks through the steps needed to import data from a spreadsheet into the DataMgr tables.
'
' MODIFICATION HISTORY
' 1.0  12/08/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel             = "../"  'Override of value from common.asp
 lcl_isRootAdmin    = False
 lcl_feature        = "datamgr_maint"
 lcl_url_parameters = ""

'Determine if the parent feature is "offline"
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

 if request("f") <> "" then
    lcl_feature = request("f")

   'Build return parameters
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Build page variables
 lcl_featurename = getFeatureName(lcl_feature)
 lcl_pagetitle   = lcl_featurename & ": Import Data from a spreadsheet"

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

 'Check for import options
  lcl_orgid            = session("orgid")
  lcl_dm_typeid        = ""
  lcl_hasCategories    = ""
  lcl_hasSubCategories = ""

  if request("orgid") <> "" then
     lcl_orgid = request("orgid")
     lcl_orgid = clng(lcl_orgid)
  end if

 if request("dm_typeid") <> "" then
    lcl_dm_typeid = request("dm_typeid")
    lcl_dm_typeid = clng(lcl_dm_typeid)
 end if

 if request("hasCategories") <> "" then
    if not containsApostrophe(request("hasCategories")) then
       lcl_hasCategories = request("hasCategories")
    end if
 end if

 if request("hasSubCategories") <> "" then
    if not containsApostrophe(request("hasSubCategories")) then
       lcl_hasSubCategories = request("hasSubCategories")
    end if
 end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

<style type="text/css">
  .instructions {
     color:         #ff0000;
     font-size:     11pt;
     margin-bottom: 10pt;
  }

  .redText {
     color: #ff0000;
  }

  .importoptions_label,
  .columnHeader,
  #startImport,
  #assignOrgDMType_results,
  #setupCategories_results,
  #setupSubCategories_results,
  #setupAddresses_results,
  #dbcolumnrows_results,
  #importdata_results {
     white-space: nowrap;
  }

  .importoptions_dropdown {
     width: 100%;
  }

  .fieldset legend {
     color: #800000;
  }

  #setupColumns {
    margin-bottom: 10px;
  }
</style>

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

<script language="javascript">
<!--

$(document).ready(function(){
  var lcl_importLineNumber;



  //Step 1 - Setup/Start Import
//  $('#cancelImportButton').prop('disabled',true);
  $('#dbcolumn_dm_importdata_id').prop('disabled',true);
  $('#dbcolumn_dm_importid').prop('disabled',true);
  $('#dbcolumn_orgid').prop('disabled',true);
  $('#dbcolumn_dm_typeid').prop('disabled',true);
  $('#dbcolumn_dmid').prop('disabled',true);
  $('#orgid').prop('disabled',true);
  $('#dm_typeid').prop('disabled',true);
  $('#assignOrgDMTButton').prop('disabled',true);

  //Step 2 - Categories and Sub-Categories
  $('#hasCategories').prop('selectedIndex',0);
  $('#hasSubCategories').prop('selectedIndex',1);

  $('#step2').css('display','none');
  $('#hasCategories').prop('disabled',true);
  $('#dbcolumn_categoryid').prop('disabled',false);
  $('#setupCategoriesButton').prop('disabled',true);

  $('#hasSubCategories').prop('disabled',true);
  $('#dbcolumn_subcategoryid').prop('disabled',true);
  $('#setupSubCategoriesButton').prop('disabled',true);

  //Step 3 - Add DB Columns and validate addresses
  $('#step3').css('display','none');
  $('#validateAddresses').prop('disabled',true);
  $('#addDBColumnRowButton').prop('disabled',true);

  $('#dbColumnTable').css('display','none');
  $('#totaltransferfields_label').css('display','none');
  $('#importResults').css('display','none');
  $('#beginImportButton').css('display','none');

  //BEGIN: Start Import Button ------------------------------------------------
  $('#startImportButton').click(function() {
     $('#startImportButton').prop('disabled',true);
     $('#cancelImportButton').prop('disabled',false);

     $('#importResults').show('slow',function(){
        $.post('datamgr_import_from_spreadsheet_action.asp', {
           userid:          '<%=session("userid")%>',
           orgid:           '<%=session("orgid")%>',
           action:          'START_IMPORT',
           isAjax:          'Y'
        }, function(result) {
           if(result.length > 0) {
              if(result.indexOf('INVALID VALUE') < 0) {
                 lcl_importLineNumber = $('#importLineNumber').val();
                 lcl_importLineNumber = Number(lcl_importLineNumber);
                 lcl_importLineNumber = lcl_importLineNumber + 1;

                 lcl_msg = lcl_importLineNumber + '. Import Started (#' + result + ')';

                 $('#importLineNumber').val(lcl_importLineNumber);
                 $('#dm_importid').val(result);
                 $('#startImport').html(lcl_msg);
                 $('#dbcolumn_dm_importdata_id').prop('disabled',false);
                 $('#dbcolumn_dm_importid').prop('disabled',false);
                 $('#dbcolumn_orgid').prop('disabled',false);
                 $('#dbcolumn_dm_typeid').prop('disabled',false);
                 $('#dbcolumn_dmid').prop('disabled',false);
                 //$('#orgid').prop('disabled',false);
                 //$('#dm_typeid').prop('disabled',false);
                 //$('#assignOrgDMTButton').prop('disabled',false);
              }
           }
        });
     });
  });
  //END: Start Import Button --------------------------------------------------

  //BEGIN: Cancel Import Button -----------------------------------------------
  $('#cancelImportButton').click(function() {
     var lcl_orgid;
     var lcl_dm_typeid;
     var lcl_dm_importid;

     lcl_orgid       = $('#orgid').val();
     lcl_dm_typeid   = $('#dm_typeid').val();
     lcl_dm_importid = $('#dm_importid').val();

     $('#cancelImportButton').prop('disabled',true);

     //lcl_cancel_url  = 'datamgr_import_from_spreadsheet_action.asp'
     //lcl_cancel_url += '?userid=<%=session("userid")%>';
     //lcl_cancel_url += '&orgid=' + lcl_orgid;
     //lcl_cancel_url += '&dm_typeid=' + lcl_dm_typeid;
     //lcl_cancel_url += '&dm_importid=' + lcl_dm_importid;
     //lcl_cancel_url += '&action=CANCEL_IMPORT';
     //lcl_cancel_url += '&isAjax=Y';
     //alert(lcl_cancel_url);

     $.post('datamgr_import_from_spreadsheet_action.asp', {
        userid:      '<%=session("userid")%>',
        orgid:       lcl_orgid,
        dm_typeid:   lcl_dm_typeid,
        dm_importid: lcl_dm_importid,
        action:      'CANCEL_IMPORT',
        isAjax:      'Y'
     }, function(result) {
        if(result == 'import cancelled') {
           location.href='datamgr_import_from_spreadsheet.asp<%=lcl_url_parameters%>';
        }
     });
  });
  //END: Cancel Import Button -------------------------------------------------

  //BEGIN: Assign Org/DM Type Button ------------------------------------------
  $('#assignOrgDMTButton').click(function() {
     var lcl_orgid;
     var lcl_dm_typeid;
     var lcl_dm_importid;

     lcl_orgid       = $('#orgid').val();
     lcl_dm_typeid   = $('#dm_typeid').val();
     lcl_dm_importid = $('#dm_importid').val();

     $('#assignOrgDMType_results').html('<br />Processing...');

     $.post('datamgr_import_from_spreadsheet_action.asp', {
        userid:      '<%=session("userid")%>',
        orgid:       lcl_orgid,
        dm_typeid:   lcl_dm_typeid,
        dm_importid: lcl_dm_importid,
        action:      'ASSIGN_ORG_DMTYPE',
        isAjax:      'Y'
     }, function(result) {
        lcl_importLineNumber = $('#importLineNumber').val();
        lcl_importLineNumber = Number(lcl_importLineNumber);
        lcl_importLineNumber = lcl_importLineNumber + 1;
        $('#importLineNumber').val(lcl_importLineNumber);

        $('#assignOrgDMType_results').html('<br />' + lcl_importLineNumber + '. ' + result);

        if(result.indexOf('INVALID VALUE') < 0) {
           $('#step2').show('slow',function(){
              $('#dbcolumn_dm_importdata_id').prop('disabled',true);
              $('#dbcolumn_dm_importid').prop('disabled',true);
              $('#dbcolumn_orgid').prop('disabled',true);
              $('#dbcolumn_dm_typeid').prop('disabled',true);
              $('#dbcolumn_dmid').prop('disabled',true);

              $('#orgid').prop('disabled',true);
              $('#dm_typeid').prop('disabled',true);
              $('#assignOrgDMTButton').prop('disabled',true);

              $('#hasCategories').prop('disabled',false);

              if($('#hasCategories').val() == 'N') {
                 $('#dbcolumn_categoryid').prop('disabled',true);
                 $('#setupCategoriesButton').prop('disabled',false);
              } else {
                 $('#dbcolumn_categoryid').prop('disabled',false);
                 $('#setupCategoriesButton').prop('disabled',true);
              }
           });
        }
     });
  });
  //END: Assign Org/DM Type Button --------------------------------------------

  //BEGIN: "Has Categories" dropdown list check -------------------------------
  $('#hasCategories').change(function() {
    if($('#hasCategories').val() == 'N') {
       document.getElementById('dbcolumn_categoryid').checked = false;
       $('#dbcolumn_categoryid').prop('disabled',true);
       $('#setupCategoriesButton').prop('disabled',false);
    } else {
       $('#dbcolumn_categoryid').prop('disabled',false);
       $('#setupCategoriesButton').prop('disabled',true);
    }
  });
  //END: "Has Categories" dropdown list check ---------------------------------

  //BEGIN: Setup Categories Button --------------------------------------------
  $('#setupCategoriesButton').click(function() {
     var lcl_hasCategories = '';

     lcl_hasCategories = $('#hasCategories').val();

     if(lcl_hasCategories == 'Y') {
        var lcl_orgid;
        var lcl_dm_typeid;
        var lcl_dm_importid;

        lcl_orgid       = $('#orgid').val();
        lcl_dm_typeid   = $('#dm_typeid').val();
        lcl_dm_importid = $('#dm_importid').val();

        $('#setupCategories_results').html('<br />Processing...');

        $.post('datamgr_import_from_spreadsheet_action.asp', {
           userid:      '<%=session("userid")%>',
           orgid:       lcl_orgid,
           dm_typeid:   lcl_dm_typeid,
           dm_importid: lcl_dm_importid,
           action:      'SETUP_CATEGORIES',
           isAjax:      'Y'
        }, function(result) {
           lcl_importLineNumber = $('#importLineNumber').val();
           lcl_importLineNumber = Number(lcl_importLineNumber);
           lcl_importLineNumber = lcl_importLineNumber + 1;
           $('#importLineNumber').val(lcl_importLineNumber);

           if(result.indexOf('COMPLETED') < 0) {
              $('#setupCategories_results').html('<br />' + lcl_importLineNumber + '. ' + result);
           } else {
              $('#setupCategories_results').html('<br />' + lcl_importLineNumber + '. Categories have been setup');
           }

           $('#hasCategories').prop('disabled',true);
           $('#dbcolumn_categoryid').prop('disabled',true);
           $('#setupCategoriesButton').prop('disabled',true);
           $('#hasSubCategories').prop('disabled',false);

           if($('#hasSubCategories').val() == 'N') {
              $('#dbcolumn_subcategoryid').prop('disabled',true);
              $('#setupSubCategoriesButton').prop('disabled',false);
           } else {
              $('#dbcolumn_subcategoryid').prop('disabled',false);
              $('#setupSubCategoriesButton').prop('disabled',true);
           }
        });
     } else {
        $('#hasCategories').prop('disabled',true);
        $('#dbcolumn_categoryid').prop('disabled',true);
        $('#setupCategoriesButton').prop('disabled',true);

        $('#hasSubCategories').prop('disabled',false);

        if($('#hasSubCategories').val() == 'N') {
           $('#dbcolumn_subcategoryid').prop('disabled',true);
           $('#setupSubCategoriesButton').prop('disabled',false);
        } else {
           $('#dbcolumn_subcategoryid').prop('disabled',false);
           $('#setupSubCategoriesButton').prop('disabled',true);
        }
     }
  });
  //END: Setup Categories Button ----------------------------------------------

  //BEGIN: "Has Sub-Categories" dropdown list check ---------------------------
  $('#hasSubCategories').change(function() {
    if($('#hasSubCategories').val() == 'N') {
       document.getElementById('dbcolumn_subcategoryid').checked = false;
       $('#dbcolumn_subcategoryid').prop('disabled',true);
       $('#setupSubCategoriesButton').prop('disabled',false);
    } else {
       $('#dbcolumn_subcategoryid').prop('disabled',false);
       $('#setupSubCategoriesButton').prop('disabled',true);
    }
  });
  //END: "Has Sub-Categories" dropdown list check -----------------------------

  //BEGIN: Setup Sub-Categories Button ----------------------------------------
  $('#setupSubCategoriesButton').click(function() {
     var lcl_hasSubCategories = '';

     lcl_hasSubCategories = $('#hasSubCategories').val();

     if(lcl_hasSubCategories != 'N') {
        var lcl_orgid;
        var lcl_dm_typeid;
        var lcl_dm_importid;

        lcl_orgid       = $('#orgid').val();
        lcl_dm_typeid   = $('#dm_typeid').val();
        lcl_dm_importid = $('#dm_importid').val();

        $('#setupSubCategories_results').html('<br />Processing...');

        $.post('datamgr_import_from_spreadsheet_action.asp', {
           userid:          '<%=session("userid")%>',
           orgid:           lcl_orgid,
           dm_typeid:       lcl_dm_typeid,
           dm_importid:     lcl_dm_importid,
           action:          'SETUP_SUBCATEGORIES',
           subcategorytype: lcl_hasSubCategories,
           isAjax:          'Y'
        }, function(result) {
           lcl_importLineNumber = $('#importLineNumber').val();
           lcl_importLineNumber = Number(lcl_importLineNumber);
           lcl_importLineNumber = lcl_importLineNumber + 1;
           $('#importLineNumber').val(lcl_importLineNumber);

           if(result.indexOf('COMPLETED') < 0) {
              $('#setupSubCategories_results').html('<br />' + lcl_importLineNumber + '. ' + result);
           } else {
              $('#setupSubCategories_results').html('<br />' + lcl_importLineNumber + '. Sub-Categories have been setup');
           }

           $('#step3').show('slow',function(){
              $('#hasSubCategories').prop('disabled',true);
              $('#dbcolumn_subcategoryid').prop('disabled',true);
              $('#setupSubCategoriesButton').prop('disabled',true);

              $('#validateAddresses').prop('disabled',false);
              $('#addDBColumnRowButton').prop('disabled',false);
              $('#dbColumnTable').show('slow',function(){
                 $('#totaltransferfields_label').show('slow');
              });
           });
        });
     } else {
        $('#step3').show('slow',function(){
           $('#hasSubCategories').prop('disabled',true);
           $('#dbcolumn_subcategoryid').prop('disabled',true);
           $('#setupSubCategoriesButton').prop('disabled',true);

           $('#validateAddresses').prop('disabled',false);
           $('#addDBColumnRowButton').prop('disabled',false);
           $('#dbColumnTable').show('slow',function(){
              $('#totaltransferfields_label').show('slow');
           });
        });
     }

  });
  //END: Setup Sub-Categories Button ------------------------------------------

  //BEGIN: Add DB Column Row Button -------------------------------------------
  $('#addDBColumnRowButton').click(function() {
    var lcl_total   = $('#totalfields').val();
    var lcl_bgcolor = $('#dbcolumnrow' + lcl_total).prop('bgcolor');
    var lcl_orgid;
    var lcl_dm_typeid;
    var lcl_transferfield_option = '';

    //Determine what the rowid and total row count is
    var num           = new Number(lcl_total);
    var lcl_new_total = (num + 1);
    var lcl_new_rowid = lcl_new_total.toString();
    var lcl_row_html  = '';

    if(lcl_bgcolor == '#eeeeee') {
       lcl_bgcolor = '#ffffff';
    } else {
       lcl_bgcolor = '#eeeeee';
    }

    lcl_orgid                = $('#orgid').val();
    lcl_dm_typeid            = $('#dm_typeid').val();

    //Build the transfer field DM options
    $.post('datamgr_import_from_spreadsheet_action.asp', {
       userid:          '<%=session("userid")%>',
       orgid:           lcl_orgid,
       dm_typeid:       lcl_dm_typeid,
       action:          'BUILD_DM_TRANSFERFIELD_OPTIONS',
       isAjax:          'Y'
    }, function(result) {

       if(result.length > 0 && result.indexOf('INVALID VALUE') < 0) {

          var lcl_addColumns = $('#addTotalColumns').val();

          if(lcl_addColumns != '') {
             var a;
             for(a = 1; a <= lcl_addColumns; a++) {

                lcl_row_html  = '  <tr id="dbcolumnrow' + a + '" class="dbColumnRow" align="center" bgcolor="' + lcl_bgcolor + '" valign="top">';
                lcl_row_html += '      <td align="left" class="columnHeader">';
                lcl_row_html += '          <input type="text" name="dbcolumn_name' + a + '" id="dbcolumn_name' + a + '" value="" size="30" maxlength="100" onchange="clearMsg(\'dbcolumn_name' + a + '\');" />';
                lcl_row_html += '      </td>';
                lcl_row_html += '      <td class="formlist">';
                lcl_row_html += '          <select name="transfer_field_data' + a + '" id="transfer_field_data' + a + '" class="transferFieldData">';
                lcl_row_html += '            <option value="">&nbsp;</option>';
                lcl_row_html += result;
                lcl_row_html += '          </select>';
                lcl_row_html += '      </td>';
                //lcl_row_html += '      <td align="center">';
                //lcl_row_html += '          <input type="checkbox" name="isAddressField' + a + '" id="isAddressField' + a + '" value="Y" />';
                //lcl_row_html += '      </td>';
                lcl_row_html += '  </tr>';

                //Append the new row to the table and increment the dbcolumn total.
                $('#dbColumnTable').append(lcl_row_html);
                $('#totalfields').val(a);
                $('#totaltransferfields_value').html(a);

                if($('#dbcolumnrows_results').html() == '') {
                   lcl_importLineNumber = $('#importLineNumber').val();
                   lcl_importLineNumber = Number(lcl_importLineNumber);
                   lcl_importLineNumber = lcl_importLineNumber + 1;
                   $('#importLineNumber').val(lcl_importLineNumber);

                   $('#dbcolumnrows_results').html(lcl_importLineNumber + '. Adding DB Columns to import...');
                }
             }

          } else {

             lcl_row_html += '  <tr id="dbcolumnrow' + lcl_new_rowid + '" class="dbColumnRow" align="center" bgcolor="' + lcl_bgcolor + '" valign="top">';
             lcl_row_html += '      <td align="left" class="columnHeader">';
             lcl_row_html += '          <input type="text" name="dbcolumn_name' + lcl_new_rowid + '" id="dbcolumn_name' + lcl_new_rowid + '" value="" size="30" maxlength="100" onchange="clearMsg(\'dbcolumn_name' + lcl_new_rowid + '\');" />';
             lcl_row_html += '      </td>';
             lcl_row_html += '      <td class="formlist">';
             lcl_row_html += '          <select name="transfer_field_data' + lcl_new_rowid + '" id="transfer_field_data' + lcl_new_rowid + '" class="transferFieldData">';
             lcl_row_html += '            <option value="">&nbsp;</option>';
             lcl_row_html += result;
             lcl_row_html += '          </select>';
             lcl_row_html += '      </td>';
             //lcl_row_html += '      <td align="center">';
             //lcl_row_html += '          <input type="checkbox" name="isAddressField' + lcl_new_rowid + '" id="isAddressField' + lcl_new_rowid + '" value="Y" />';
             //lcl_row_html += '      </td>';
             lcl_row_html += '  </tr>';

             //Append the new row to the table and increment the dbcolumn total.
             $('#dbColumnTable').append(lcl_row_html);
             $('#totalfields').val(lcl_new_rowid);
             $('#totaltransferfields_value').html(lcl_new_rowid);

             if($('#dbcolumnrows_results').html() == '') {
                lcl_importLineNumber = $('#importLineNumber').val();
                lcl_importLineNumber = Number(lcl_importLineNumber);
                lcl_importLineNumber = lcl_importLineNumber + 1;
                $('#importLineNumber').val(lcl_importLineNumber);

                $('#dbcolumnrows_results').html(lcl_importLineNumber + '. Adding DB Columns to import...');
             }
          }

          $('#beginImportButton').show('slow');
       }
    });
  });
  //END: Add DB Column Row Button ---------------------------------------------
});

//function enableDisableStep1Button() {
//  $('#step1Button').prop('disabled',true);

//  if($('#assignOrgDMType_orgid').prop('selectedIndex') > 0) {
//     $('#step1Button').prop('disabled',false);
//  }
//}

function beginImport() {
  var lcl_orgid               = $('#orgid').val();
  var lcl_dm_importid         = $('#dm_importid').val();
  var lcl_dm_typeid           = $('#dm_typeid').val();
  var lcl_totalfields         = $('#totalfields').val();
  var lcl_totalfieldsimported = Number(0);
  var lcl_validateAddresses   = $('#validateAddresses').val();

  lcl_importLineNumber = $('#importLineNumber').val();
  //lcl_importLineNumber = Number(lcl_importLineNumber);
  //lcl_importLineNumber = lcl_importLineNumber + 1;
  //$('#importLineNumber').val(lcl_importLineNumber);

  $('#dbcolumnrows_results').html(lcl_importLineNumber + '. Importing spreadsheet data...');

  $('#validateAddresses').prop('disabled',true);
  $('#addDBColumnRowButton').prop('disabled',true);
  $('#beginImportButton').prop('disabled',true);

  //Cycle through DMT Field rows and determine how many have been selected
  if(lcl_totalfields > 0) {
     var i = 1;
     var lcl_dbcolumn_name       = '';
     var lcl_transfer_field_data = '';
     var lcl_results             = '';

     for(i = 1; i <= lcl_totalfields; i++) {
         lcl_dbcolumn_name       = $('#dbcolumn_name'       + i).val();
         lcl_transfer_field_data = $('#transfer_field_data' + i).val();

         $('#dbcolumn_name'       + i).prop('disabled',true);
         $('#transfer_field_data' + i).prop('disabled',true);

         if(lcl_transfer_field_data != '') {
            $.post('datamgr_import_from_spreadsheet_action.asp', {
               userid:              '<%=session("userid")%>',
               orgid:               lcl_orgid,
               dm_importid:         lcl_dm_importid,
               dm_typeid:           lcl_dm_typeid,
               dbcolumn_name:       lcl_dbcolumn_name,
               validateAddresses:   lcl_validateAddresses,
               transfer_field_data: lcl_transfer_field_data,
               action:              'IMPORT_SPREADSHEET_VALUES',
               //overridevalues:      lcl_overrideValues,
               isAjax:              'Y'
            }, function(result) {
               var lcl_importdata          = '';
               var lcl_import_complete_msg = '';

               if(result != '') {
                  if($('#importdata_results').html() == '') {
                     lcl_importdata = result;
                  } else {
                     lcl_importdata  = $('#importdata_results').html();
                     lcl_importdata += '<br />';
                     lcl_importdata += result;
                  }

                  $('#importdata_results').html(lcl_importdata);

                  //Determine if the last dbcolumn has been imported.
                  //If "yes" then move to the next step.
                  lcl_totalfieldsimported = lcl_totalfieldsimported + 1

                  if(lcl_totalfieldsimported == lcl_totalfields) {
                     lcl_import_complete_msg  = $('#importdata_results').html();
                     lcl_import_complete_msg += '<br />';
                     lcl_import_complete_msg += '<span class="redText">Import Complete</span>';
                     $('#importdata_results').html(lcl_import_complete_msg);

                     $.post('datamgr_import_from_spreadsheet_action.asp', {
                        orgid:       lcl_orgid,
                        dm_importid: lcl_dm_importid,
                        action:      'COMPLETE_IMPORT',
                        isAjax:      'Y'
                     });
                     //}, function(result) {
                        //if(result == 'complete') {
                        //   lcl_import_complete_msg  = $('#importdata_results').html();
                        //   lcl_import_complete_msg += '<br />';
                        //   lcl_import_complete_msg += '<span class="redText">Import Complete</span>';
                        //   $('#importdata_results').html(lcl_import_complete_msg);
                        //}
                     //});
                  }
               }
            });
         }
      }
   }
}

function confirmDelete(p_id) {
  lcl_datamgr = document.getElementById("datamgr"+p_id).innerHTML;

 	if (confirm("Are you sure you want to delete '" + lcl_datamgr + "' ?")) { 
  				//DELETE HAS BEEN VERIFIED
		  		location.href='datamgr_action.asp<%=lcl_delete_datamgr%>&dmid='+ p_id;
		}
}

function doCalendar(ToFrom) {
  w = 350;
  h = 250;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}

function openCustomReports(p_report) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("../customreports/customreports.asp?cr=' + p_report + '&dmt=<%=lcl_dm_typeid%>", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,resizable=1,menubar=0,left=' + l + ',top=' + t + '")');
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

function calculateTotalDBColumns() {
  var lcl_total_dbcolumns = Number(0);

  if(document.getElementById('dbcolumn_dm_importdata_id').checked) {
     lcl_total_dbcolumns = lcl_total_dbcolumns + 1;
  }

  if(document.getElementById('dbcolumn_dm_importid').checked) {
     lcl_total_dbcolumns = lcl_total_dbcolumns + 1;
  }

  if(document.getElementById('dbcolumn_orgid').checked) {
     lcl_total_dbcolumns = lcl_total_dbcolumns + 1;
  }

  if(document.getElementById('dbcolumn_dm_typeid').checked) {
     lcl_total_dbcolumns = lcl_total_dbcolumns + 1;
  }

  if(document.getElementById('dbcolumn_dmid').checked) {
     lcl_total_dbcolumns = lcl_total_dbcolumns + 1;
  }

  $('#total_dbcolumns').val(lcl_total_dbcolumns);

  if(lcl_total_dbcolumns == 5) {
     $('#orgid').prop('disabled',false);
     $('#dm_typeid').prop('disabled',false);
     $('#assignOrgDMTButton').prop('disabled',false);
  } else {
     $('#orgid').prop('disabled',true);
     $('#dm_typeid').prop('disabled',true);
     $('#assignOrgDMTButton').prop('disabled',true);
  }

}

function checkboxEnableFields(p_fieldid) {
  var lcl_fieldid = p_fieldid;
  var lcl_field   = document.getElementById(lcl_fieldid);

  if(lcl_field.checked) {
     if(lcl_fieldid == 'dbcolumn_categoryid') {
        $('#setupCategoriesButton').prop('disabled',false);
     } else if(lcl_fieldid == 'dbcolumn_subcategoryid') {
        $('#setupSubCategoriesButton').prop('disabled',false);
     }
  } else {
     if(lcl_fieldid == 'dbcolumn_categoryid') {
        $('#setupCategoriesButton').prop('disabled',true);
     } else if(lcl_fieldid == 'dbcolumn_subcategoryid') {
        $('#setupSubCategoriesButton').prop('disabled',true);
     }
  }
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<p style="display:none">DO NOT forget to assign the sub-categories when creating a new DMID!</p>

<%
 response.write "<form name=""spreadsheetImport"" id=""spreadsheetImport"" method=""post"" action=""datamgr_import_from_spreadsheet_action.asp"">" & vbcrlf
 response.write "  <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
 response.write "  <input type=""text"" name=""dm_importid"" id=""dm_importid"" value="""" size=""5"" maxlength=""10"" />" & vbcrlf
 response.write "  <input type=""hidden"" name=""importLineNumber"" id=""importLineNumber"" value=""0"" maxlength=""10"" />" & vbcrlf
 'response.write "  <input type=""hidden"" name=""action"" id=""action"" value="""" size=""20"" maxlength=""20"" />" & vbcrlf

 response.write "<div id=""content"">" & vbcrlf
 response.write " 	<div id=""centercontent"">" & vbcrlf
 response.write "    <p>" & vbcrlf
 response.write "    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""1000px"">" & vbcrlf
 response.write "      <tr>" & vbcrlf
 response.write "          <td><font size=""+1""><strong>" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
 response.write "          <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;"">&nbsp;</span></td>" & vbcrlf
 response.write "      </tr>" & vbcrlf
 response.write "    </table>" & vbcrlf
 response.write "    </p>" & vbcrlf
 response.write "    <p>" & vbcrlf
 response.write "      <table border=""0"" width=""100%"">" & vbcrlf
 response.write "        <tr>" & vbcrlf
 response.write "            <td>" & vbcrlf
 response.write "                <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
 response.write "                  <tr>" & vbcrlf
 response.write "                      <td>" & vbcrlf
 response.write "                          <input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to List"" class=""button"" onclick=""location.href='datamgr_list.asp" & lcl_url_parameters & "'"" />" & vbcrlf
 response.write "                      </td>" & vbcrlf
 response.write "                      <td align=""right"">" & vbcrlf
 response.write "                          <input type=""button"" name=""cancelImportButton"" id=""cancelImportButton"" value=""Cancel Import"" class=""button"" />" & vbcrlf
 response.write "                          <input type=""button"" name=""startImportButton"" id=""startImportButton"" value=""Start Import"" class=""button"" />" & vbcrlf
 response.write "                      </td>" & vbcrlf
 response.write "                  </tr>" & vbcrlf
 response.write "                </table>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td>&nbsp;</td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td>" & vbcrlf

'BEGIN: Import: Step 1 --------------------------------------------------------
 response.write "    <fieldset name=""step1"" id=""step1"" class=""fieldset"">" & vbcrlf
 response.write "      <legend>Step 1</legend>" & vbcrlf

 response.write "      <div id=""setupColumns"">" & vbcrlf
 response.write "        <span class=""instructions"">Ensure that the following columns exist on egov_dm_import_data.</span><br />" & vbcrlf
 response.write "        <input type=""checkbox"" class=""setupColumns"" name=""dbcolumn_dm_importdata_id"" id=""dbcolumn_dm_importdata_id"" value=""1"" onclick=""calculateTotalDBColumns();"" /> dm_importdata_id (int NOT NULL)<br />" & vbcrlf
 response.write "        <input type=""checkbox"" class=""setupColumns"" name=""dbcolumn_dm_importid"" id=""dbcolumn_dm_importid"" value=""1"" onclick=""calculateTotalDBColumns();"" /> dm_importid (int NOT NULL)<br />" & vbcrlf
 response.write "        <input type=""checkbox"" class=""setupColumns"" name=""dbcolumn_orgid"" id=""dbcolumn_orgid"" value=""1"" onclick=""calculateTotalDBColumns();"" /> orgid (int NULL)<br />" & vbcrlf
 response.write "        <input type=""checkbox"" class=""setupColumns"" name=""dbcolumn_dm_typeid"" id=""dbcolumn_dm_typeid"" value=""1"" onclick=""calculateTotalDBColumns();"" /> dm_typeid (int NULL)<br />" & vbcrlf
 response.write "        <input type=""checkbox"" class=""setupColumns"" name=""dbcolumn_dmid"" id=""dbcolumn_dmid"" value=""1"" onclick=""calculateTotalDBColumns();"" /> dmid (int NULL)<br />" & vbcrlf
 response.write "        <input type=""hidden"" name=""total_dbcolumns"" id=""total_dbcolumns"" value=""0"" size=""3"" maxlength=""10"" />" & vbcrlf
 response.write "      </div>" & vbcrlf

 response.write "      <div class=""instructions"">" & vbcrlf
 response.write "        Select the [Organization] and [DataMgr Type] we are importing data into.<br />" & vbcrlf
 response.write "      </div>" & vbcrlf
 response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""importoptions_label"">" & vbcrlf
 response.write "               Organization:" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td class=""importoptions_dropdown"">" & vbcrlf
 response.write "                <select name=""orgid"" id=""orgid"">" & vbcrlf
                                   displayOrgOptions lcl_orgid
 response.write "                </select>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""importoptions_label"">" & vbcrlf
 response.write "                DM Type:" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td class=""importoptions_dropdown"">" & vbcrlf
 response.write "                <select name=""dm_typeid"" id=""dm_typeid"">" & vbcrlf
                                   displayDMTypeOptions session("orgid"), lcl_dm_typeid
 response.write "                </select>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
 response.write "        <tr>" & vbcrlf
 response.write "            <td colspan=""2"">" & vbcrlf
 response.write "                <input type=""button"" name=""assignOrgDMTButton"" id=""assignOrgDMTButton"" value=""Assign Organization and DM Type"" class=""button"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "    </table>" & vbcrlf
 response.write "    </fieldset>" & vbcrlf
'END: Import: Step 1 ----------------------------------------------------------

'BEGIN: Import: Step 2 --------------------------------------------------------
 response.write "    <p>" & vbcrlf
 response.write "    <fieldset name=""step2"" id=""step2"" class=""fieldset"">" & vbcrlf
 response.write "      <legend>Step 2</legend>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "      <span class=""instructions"">Are categories included on the spreadsheet?</span>" & vbcrlf
 response.write "      <select name=""hasCategories"" id=""hasCategories"">" & vbcrlf
 response.write "        <option value=""Y"">Yes</option>" & vbcrlf
 response.write "        <option value=""N"">No</option>" & vbcrlf
 response.write "      </select><br />" & vbcrlf
 response.write "      <div>" & vbcrlf
 response.write "        <span class=""instructions"">Ensure that the following column exists on egov_dm_import_data.</span><br />" & vbcrlf
 response.write "        <input type=""checkbox"" name=""dbcolumn_categoryid"" id=""dbcolumn_categoryid"" value=""1"" onclick=""checkboxEnableFields('dbcolumn_categoryid');"" /> categoryid (int NULL)" & vbcrlf
 response.write "      </div>" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "        <input type=""button"" name=""setupCategoriesButton"" id=""setupCategoriesButton"" value=""Continue"" class=""button"" />" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>&nbsp;</p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "      <span class=""instructions"">Are sub-categories included on the spreadsheet?</span><br />" & vbcrlf
 response.write "      <select name=""hasSubCategories"" id=""hasSubCategories"">" & vbcrlf
 response.write "        <option value=""1to1"">Yes - 1 Sub-Category per Category</option>" & vbcrlf
 response.write "        <option value=""N"">No</option>" & vbcrlf
 response.write "      </select><br />" & vbcrlf
 response.write "      <div>" & vbcrlf
 response.write "        <span class=""instructions"">Ensure that the following column exists on egov_dm_import_data.</span><br />" & vbcrlf
 response.write "        <input type=""checkbox"" name=""dbcolumn_subcategoryid"" id=""dbcolumn_subcategoryid"" value=""1"" onclick=""checkboxEnableFields('dbcolumn_subcategoryid');"" /> subcategoryid (int NULL)" & vbcrlf
 response.write "      </div>" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "        <input type=""button"" name=""setupSubCategoriesButton"" id=""setupSubCategoriesButton"" value=""Continue"" class=""button"" />" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "    </fieldset>" & vbcrlf
 response.write "    </p>" & vbcrlf
'END: Import: Step 2 ----------------------------------------------------------

'BEGIN: Import: Step 3 --------------------------------------------------------
 response.write "    <p>" & vbcrlf
 response.write "    <fieldset name=""step3"" id=""step3"" class=""fieldset"">" & vbcrlf
 response.write "      <legend>Step 3</legend>" & vbcrlf
 response.write "      <div class=""instructions"">" & vbcrlf
 response.write "        For each column on the egov_dm_import_data table select the DataMgr field that it is to be imported into.<br />" & vbcrlf
 response.write "        <p>" & vbcrlf
 response.write "          Do you want to validate the addresses?" & vbcrlf
 response.write "          <select name=""validateAddresses"" id=""validateAddresses"">" & vbcrlf
 response.write "            <option value=""Y"">Yes</option>" & vbcrlf
 response.write "            <option value=""N"">No</option>" & vbcrlf
 response.write "          </select>" & vbcrlf
 response.write "        </p>" & vbcrlf
 response.write "      </div>"
 response.write "      <p>" & vbcrlf
 response.write "        <input type=""button"" name=""addDBColumnRowButton"" id=""addDBColumnRowButton"" value=""Add Column"" class=""button"" />" & vbcrlf
 response.write "        <input type=""text"" name=""addTotalColumns"" id=""addTotalColumns"" value="""" size=""5"" maxlength=""10"" />" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "      <table id=""dbColumnTable"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"">" & vbcrlf
 response.write "        <tr align=""left"">" & vbcrlf
 response.write "            <th class=""columnHeader"">DB Column Name</th>" & vbcrlf
 response.write "            <th>Transfer Data to DM Field...</th>" & vbcrlf
 'response.write "            <th align=""center"">Is<br />Address<br />Field</th>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "      </table>" & vbcrlf
 response.write "      <div id=""totaltransferfields_label"" align=""right""><strong>Total DB Columns to Transfer: </strong>[<span id=""totaltransferfields_value"">0</span>]</div>" & vbcrlf
 response.write "      <input type=""text"" name=""totalfields"" id=""totalfields"" value=""0"" size=""3"" maxlength=""10"" />" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "        <input type=""button"" name=""beginImportButton"" id=""beginImportButton"" value=""Begin Import"" class=""button"" onclick=""beginImport()"" />" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "    </fieldset>" & vbcrlf
 response.write "    </p>" & vbcrlf
'END: Import: Step 3 ----------------------------------------------------------

 response.write "            </td>" & vbcrlf
 response.write "            <td>" & vbcrlf
 response.write "                <fieldset id=""importResults"" class=""fieldset"">" & vbcrlf
 response.write "                  <legend>Import Results</legend>" & vbcrlf
 response.write "                  <p>" & vbcrlf
 response.write "                    <span id=""startImport""></span>" & vbcrlf
 response.write "                    <span id=""assignOrgDMType_results""></span>" & vbcrlf
 response.write "                    <span id=""setupCategories_results""></span>" & vbcrlf
 response.write "                    <span id=""setupSubCategories_results""></span>" & vbcrlf
 response.write "                    <span id=""setupAddresses_results""></span>" & vbcrlf
 response.write "                    <span id=""dbcolumnrows_results""></span>" & vbcrlf
 response.write "                    <span id=""importdata_results""></span>" & vbcrlf
 response.write "                  </p>" & vbcrlf
 response.write "                </fieldset>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "      </table>" & vbcrlf
 response.write "    </p>" & vbcrlf

 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
 response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"--> 
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayDMTypeOptions(iOrgID, iSC_DMTypeID)

  sOrgID                 = 0
  sSC_DMTypeID           = ""
  lcl_selected_dm_typeid = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iSC_DMTypeID <> "" then
     sSC_DMTypeID = clng(iSC_DMTypeID)
  end if

  sSQL = "SELECT dm_typeid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND isActive = 1 "
  sSQL = sSQL & " AND isTemplate = 0 "

'  if sSC_DMTypeID <> "" then
'     sSQL = sSQL & " AND dm_typeid = " & sSC_DMTypeID
'  end if

  sSQL = sSQL & " ORDER BY description "

 	set oDisplayDMTypeOptions = Server.CreateObject("ADODB.Recordset")
	 oDisplayDMTypeOptions.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayDMTypeOptions.eof then
     do while not oDisplayDMTypeOptions.eof

        if sSC_DMTypeID = oDisplayDMTypeOptions("dm_typeid") then
           lcl_selected_dm_typeid = " selected=""selected"""
        else
           lcl_selected_dm_typeid = ""
        end if

        response.write "  <option value=""" & oDisplayDMTypeOptions("dm_typeid") & """" & lcl_selected_dm_typeid & ">" & oDisplayDMTypeOptions("description") & "</option>" & vbcrlf

        oDisplayDMTypeOptions.movenext
     loop
  end if

  oDisplayDMTypeOptions.close
  set oDisplayDMTypeOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayTransferFieldOptions(iOrgID, iDM_TypeID)
  sOrgID        = 0
  sDM_TypeID = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iDM_TypeID <> "" then
     sDM_TypeID = clng(iDM_TypeID)
  end if

  sSQL = sSQL & "SELECT DISTINCT "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQl = sSQL & " dms.sectionname, "
  sSQL = sSQL & " dmtf.dm_fieldid, "
  sSQL = sSQL & " dmtf.section_fieldid, "
  sSQL = sSQL & " dmsf.fieldname "
  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_types dmt "
  sSQL = sSQL &            " ON dmt.dm_typeid = dmtf.dm_typeid "
  sSQL = sSQL &            " AND dmt.isActive = 1 "
  sSQL = sSQL &            " AND dmt.isTemplate = 0 "
  sSQL = sSQL &            " AND dmt.orgid = " & sOrgID
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &            " ON dmts.dm_sectionid = dmtf.dm_sectionid "
  sSQL = sSQL &            " AND dmts.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections dms "
  sSQL = sSQL &            " ON dms.sectionid = dmts.sectionid "
  sSQL = sSQL &            " AND dms.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &            " ON dmsf.section_fieldid = dmtf.section_fieldid "
  sSQL = sSQL &            " AND dmsf.isActive = 1 "
  sSQL = sSQL & " WHERE dmtf.orgid = " & sOrgID
  sSQL = sSQL & " AND dmtf.dm_typeid = " & sDM_TypeID
  sSQL = sSQL & " ORDER BY dms.sectionname, dmsf.fieldname "

 	set oDMTransferFieldsOptions = Server.CreateObject("ADODB.Recordset")
	 oDMTransferFieldsOptions.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMTransferFieldsOptions.eof then
     lcl_bgcolor            = "#ffffff"

     do while not oDMTransferFieldsOptions.eof

        response.write "  <option value=""dmsectionid" & oDMTransferFieldsOptions("dm_sectionid") & "_dmfieldid" & oDMTransferFieldsOptions("dm_fieldid") & """>" & oDMTransferFieldsOptions("sectionname") & ": " & oDMTransferFieldsOptions("fieldname") & "</option>" & vbcrlf

        oDMTransferFieldsOptions.movenext
     loop
  end if

  oDMTransferFieldsOptions.close
  set oDMTransferFieldsOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayOrgOptions(iOrgID)

  dim sSQL, lcl_orgid

  lcl_orgid = 0

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

 	sSQL = "SELECT "
  sSQL = sSQL & " orgid, "
  sSQL = sSQL & " orgcity "
  sSQL = sSQL & " FROM organizations "
	 sSQL = sSQL & " WHERE isdeactivated = 0 "
  sSQL = sSQL & " ORDER BY orgcity "

 	set oGetOrgOptions = Server.CreateObject("ADODB.Recordset")
 	oGetOrgOptions.Open sSQL, Application("DSN"), 3, 1

  if not oGetOrgOptions.eof then
     do while not oGetOrgOptions.eof
        if lcl_orgid = clng(oGetOrgOptions("orgid")) then
           lcl_selected_org = " selected=""selected"""
        else
           lcl_selected_org = ""
        end if

        response.write "  <option value=""" & oGetOrgOptions("orgid") & """" & lcl_selected_org & ">" & oGetOrgOptions("orgcity") & "</option>" & vbcrlf

        oGetOrgOptions.movenext
     loop
  end if

  oGetOrgOptions.close
  set oGetOrgOptions = nothing

end sub
%>