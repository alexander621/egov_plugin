<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!-- #include file="datamgr_build_sections_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: maintain_dmt_section.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DM Type Section
'
' MODIFICATION HISTORY
' 1.0 02/24/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

'Retrieve the dmid to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 lcl_dmid      = 0
 lcl_sectionid = 0
 lcl_dm_typeid = 0

 if request("dmid") <> "" then
    lcl_dmid = request("dmid")
 end if

 if request("sectionid") <> "" then
    lcl_sectionid = request("sectionid")
 end if

'Get Section Info
 getSectionInfo lcl_sectionid, lcl_sectionname, lcl_sectiontype

'Get DMTypeID
 lcl_dm_typeid = getDMTypeID_byDMID(lcl_dmid)

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature             = "datamgr_maint"
 lcl_featurename         = getFeatureName(lcl_feature)
 lcl_display_featurename = lcl_featurename
 lcl_display_featurename = replace(lcl_display_featurename,"Maintain ","")

 if request("f") <> "" then
    if not containsApostrophe(request("f")) then
       lcl_feature     = request("f")
       lcl_featurename = getFeatureName(lcl_feature)
    end if
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
   	response.redirect sLevel & "permissiondenied.asp?f=" & lcl_feature
 end if

'Build return parameters
 lcl_url_parameters = ""

 if lcl_feature <> "datamgr_maint" then
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

'Check for org features
 lcl_orghasfeature_feature            = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain   = orghasfeature(lcl_feature)
 lcl_orghasfeature_issue_location     = orghasfeature("issue location")
 lcl_orghasfeature_large_address_list = orghasfeature("large address list")

'Check for user permissions
 lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
 lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

'Determine if the user has clicked on the "Import Address Fields" button
 if request("importAddressFields") <> "" then
    lcl_importAddressFields = request("importAddressFields")
 else
    lcl_importAddressFields = "N"
 end if

 if lcl_importAddressFields = "Y" then
    if lcl_orghasfeature_large_address_list then
       lcl_importstreet_number  = request("residentstreetnumber")
       lcl_importstreet_address = request("streetaddress")
    else
       lcl_importstreet_number  = ""
       lcl_importstreet_address = request("streetaddress")
    end if

    lcl_importsortstreetname = request("sortstreetname")
 else
    lcl_importstreet_number  = ""
    lcl_importstreet_address = ""
    lcl_importsortstreetname = ""
 end if

'Get the layoutid
 lcl_layoutid = getDMTLayoutID(lcl_dm_typeid)

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"

    if lcl_success = "SU" then
       lcl_onload = lcl_onload & "window.opener.location.reload();"
    end if
 end if

'Determine if this section has an "ADDRESS" field.  Used to enable javascript to set up address field(s).
 lcl_addressfield_exists = checkForAddressFieldInSection(lcl_sectionid)

'If the "large address" feature is turned on then enable/disable the "Import Address Fields" button
 'if lcl_orghasfeature_large_address_list then
 '   lcl_onload = lcl_onload & "checkAddressButtons();"
 'end if

'Show/Hide all "hidden" fields.  (HIDDEN = hide, TEXT = show)
 lcl_hidden = "hidden"

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_displayfeaturename%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />

 	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
  <script type="text/javascript" src="../scripts/removespaces.js"></script>
  <script type="text/javascript" src="../scripts/selectAll.js"></script>
  <script type="text/javascript" src="../scripts/tooltip_new.js"></script>
  <script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <% 'Past this version and the code doesn't like the "checked" attribute for the radion options in the "hours" section %>
  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>
<% '  <script type="text/javascript" src="https://github.com/jquery/jquery-ui.git"></script> %>


<script language="javascript">
$(document).ready(function(){
<%
 'BEGIN: Hours jQuery ----------------------------------------------------------
  if lcl_sectiontype = "HOURS" then
%>
  $('.editHoursInfo').css("display", "none");
		$('.editHoursInfo').each(function(){
    var lcl_id = this.id;
        lcl_id = lcl_id.replace('editHoursInfo','');

    $('#editHoursButton'+lcl_id).click(function() {
      $('#editHoursInfo'+lcl_id).slideDown('slow',function() {
      });
    });

    function enableDisableHoursFields(iSelection) {
      var lcl_hours_selection = 'OTHER';
      var lcl_fieldid         = '';

      //Determine which "time" fields are enabled/disabled
      if(iSelection == '') {
         var lcl_fieldvalue      = $('#fieldvalue'+lcl_id).val();
         var lcl_fieldvalue_colon = lcl_fieldvalue.indexOf(':');
         var lcl_fieldvalue_am    = lcl_fieldvalue.indexOf('AM');
         var lcl_fieldvalue_pm    = lcl_fieldvalue.indexOf('PM');

         if(lcl_fieldvalue_colon > -1 || lcl_fieldvalue_am > -1 || lcl_fieldvalue_pm > -1) {
            lcl_hours_selection = 'TIME';
         } else {
            lcl_hours_selection = 'OTHER';
         }
      } else {
         lcl_hours_selection = iSelection;
      }

      if(lcl_hours_selection == 'TIME') {
         lcl_fieldid    = '#selection_time' + lcl_id;
         lcl_time_hours = '';
         lcl_time_other = 'disabled';
      } else {
         lcl_fieldid    = '#selection_other' + lcl_id;
         lcl_time_hours = 'disabled';
         lcl_time_other = '';
      }

      //$(lcl_fieldid).attr('checked','checked');
      //$('#time_other'    + lcl_id).attr('disabled',lcl_time_other);
      //$('#START_HOURS'   + lcl_id).attr('disabled',lcl_time_hours);
      //$('#START_MINUTES' + lcl_id).attr('disabled',lcl_time_hours);
      //$('#START_AMPM'    + lcl_id).attr('disabled',lcl_time_hours);
      //$('#END_HOURS'     + lcl_id).attr('disabled',lcl_time_hours);
      //$('#END_MINUTES'   + lcl_id).attr('disabled',lcl_time_hours);
      //$('#END_AMPM'      + lcl_id).attr('disabled',lcl_time_hours);
      $(lcl_fieldid).prop('checked','checked');
      $('#time_other'    + lcl_id).prop('disabled',lcl_time_other);
      $('#START_HOURS'   + lcl_id).prop('disabled',lcl_time_hours);
      $('#START_MINUTES' + lcl_id).prop('disabled',lcl_time_hours);
      $('#START_AMPM'    + lcl_id).prop('disabled',lcl_time_hours);
      $('#END_HOURS'     + lcl_id).prop('disabled',lcl_time_hours);
      $('#END_MINUTES'   + lcl_id).prop('disabled',lcl_time_hours);
      $('#END_AMPM'      + lcl_id).prop('disabled',lcl_time_hours);
    }

    //Initialize hours fields
    enableDisableHoursFields('');

    $('#selection_time'+lcl_id).click(function() {
      enableDisableHoursFields('TIME');
    });

    $('#selection_other'+lcl_id).click(function() {
      enableDisableHoursFields('OTHER');
    });

    $('#hoursSaveButton'+lcl_id).click(function() {
      var lcl_fieldvalue      = "";
      var lcl_starttime       = "";
      var lcl_endtime         = "";
      var lcl_check_starttime = "";
      var lcl_check_endtime   = "";

      //Get the "start/end times" or "other" values and build the value to be stored.
      //if($('#time_other'+lcl_id).val() != '') {
      if($('#selection_other' + lcl_id).prop('checked')) {
         lcl_fieldvalue = $('#time_other'+lcl_id).val();
      } else {

         lcl_starttime  = $('#START_HOURS'+lcl_id).val();
         lcl_starttime += ':';
         lcl_starttime += $('#START_MINUTES'+lcl_id).val();
         lcl_starttime += ' ';
         lcl_starttime += $('#START_AMPM'+lcl_id).val();

         lcl_endtime  = $('#END_HOURS'+lcl_id).val();
         lcl_endtime += ':';
         lcl_endtime += $('#END_MINUTES'+lcl_id).val();
         lcl_endtime += ' ';
         lcl_endtime += $('#END_AMPM'+lcl_id).val();

         lcl_check_starttime = lcl_starttime;
         lcl_check_starttime = lcl_check_starttime.replace(':','');
         lcl_check_starttime = lcl_check_starttime.replace(' ','');

         lcl_check_endtime = lcl_starttime;
         lcl_check_endtime = lcl_check_endtime.replace(':','');
         lcl_check_endtime = lcl_check_endtime.replace(' ','');

         if(lcl_check_starttime != "") {
            lcl_fieldvalue = lcl_starttime;
         }

         if(lcl_check_endtime != "") {
            if(lcl_fieldvalue != "") {
               lcl_fieldvalue = lcl_fieldvalue + ' - ' + lcl_endtime;
            } else {
               lcl_fieldvalue = lcl_endtime;
            }
         }
      }

      $('#fieldvalue'+lcl_id).val(lcl_fieldvalue);

      //Save the value
      var lcl_dm_typeid    = $('#dm_typeid').val();
      var lcl_dmid         = $('#dmid').val();
      var lcl_dm_valueid   = $('#dm_valueid'+lcl_id).val();
      var lcl_dm_sectionid = $('#dm_sectionid'+lcl_id).val();
      var lcl_dm_fieldid   = $('#dm_fieldid'+lcl_id).val();

      $.post('update_dm_value.asp', {
          userid:       '<%=session("userid")%>',
          orgid:        '<%=session("orgid")%>',
          dm_typeid:    lcl_dm_typeid,
          dmid:         lcl_dmid,
          dm_sectionid: lcl_dm_sectionid,
          dm_fieldid:   lcl_dm_fieldid,
          dm_valueid:   lcl_dm_valueid,
          fieldvalue:   lcl_fieldvalue,
          isAjax:       'Y'
        }, function(result) {
          displayScreenMsg(result);
          $('#display_fieldvalue'+lcl_id).html(lcl_fieldvalue);
          window.opener.location.reload();
          $('#editHoursInfo'+lcl_id).slideUp('slow',function() {
          });
      });
    });
  });
<%
  end if
 'END: Hours jQuery -----------------------------------------------------------
%>

});

<%
 'BEGIN: Check for the "issue location" feature ------------------------------
  if lcl_orghasfeature_issue_location AND lcl_addressfield_exists then
     lcl_addresstype = ""

     response.write "$(document).ready(function(){" & vbcrlf

     if lcl_orghasfeature_large_address_list then
        response.write "  $('#validaddresslist').hide();" & vbcrlf
        lcl_addresstype = "LARGE"
     end if

     response.write "  enableDisableAddressFields('');" & vbcrlf         

    'Street Number - onChange ---------------------------------------------
     response.write "  $('#residentstreetnumber').change(function() {" & vbcrlf
     response.write "    clearMsg('residentstreetnumber');" & vbcrlf
     response.write "    clearMsg('validateAddress');" & vbcrlf
     response.write "    enableDisableAddressFields('');" & vbcrlf
     response.write "    if($('#residentstreetnumber').val() != '') {" & vbcrlf
     response.write "       $('#ques_issue2').val('');" & vbcrlf
     response.write "    }" & vbcrlf
     response.write "  });" & vbcrlf

    'Stret Address - onChange ---------------------------------------------
     response.write "  $('#streetaddress').change(function() {" & vbcrlf
     response.write "    clearMsg('streetaddress');" & vbcrlf
     response.write "    clearMsg('validateAddress');" & vbcrlf
     response.write "    enableDisableAddressFields('');" & vbcrlf
     response.write "    if($('#streetaddress').val() != '0000') {" & vbcrlf
     response.write "       $('#ques_issue2').val('');" & vbcrlf
     response.write "    }" & vbcrlf
     response.write "  });" & vbcrlf

    'Other Address - onChange ---------------------------------------------
     response.write "  $('#ques_issue2').change(function() {" & vbcrlf
     response.write "    enableDisableAddressFields('');" & vbcrlf
     response.write "    if($('#ques_issue2').val() != '') {" & vbcrlf
     response.write "       $('#residentstreetnumber').val('');" & vbcrlf
     response.write "       $('#streetaddress').val('0000');" & vbcrlf
     response.write "       $('#validstreet').val('N');" & vbcrlf
     response.write "    }" & vbcrlf
     response.write "  });" & vbcrlf

    'Import Address Fields Button - onClick --------------------------------------
     'response.write "  $('#importAddress').click(function() {" & vbcrlf
     'response.write "    $('#importAddressFields').val('Y');" & vbcrlf
     '//response.write "    $('#maintain_dmt_section').attr('action','maintain_dmt_section.asp');" & vbcrlf
     'response.write "    $('#maintain_dmt_section').prop('action','maintain_dmt_section.asp');" & vbcrlf
     'response.write "    $('#maintain_dmt_section').submit();" & vbcrlf
     'response.write "  });" & vbcrlf

     response.write "});" & vbcrlf

    'BEGIN: Check Address -------------------------------------------------
     response.write "function checkAddress(iFunction, iValidate) {" & vbcrlf
     response.write "  clearScreenMsg();" & vbcrlf
     response.write "  var lcl_streetnumber = $('#residentstreetnumber').val();" & vbcrlf
     response.write "  var lcl_streetname   = $('#streetaddress').val();" & vbcrlf
     response.write "  var lcl_otheraddress = $('#ques_issue2').val();" & vbcrlf

     response.write "  if(lcl_otheraddress == '') {" & vbcrlf
     response.write "     lcl_success = validateAddress();" & vbcrlf

    'Validate the street number and name entered to determine if it is a valid address in the system for the org
     response.write "     if(lcl_success) {" & vbcrlf
     response.write "        $.post('checkaddress.asp', {" & vbcrlf
     response.write "           addresstype: '" & lcl_addresstype & "'," & vbcrlf
     response.write "           stnumber:    lcl_streetnumber," & vbcrlf
     response.write "           stname:      lcl_streetname," & vbcrlf
     response.write "           returntype:  'CHECK'" & vbcrlf
     response.write "         }, function(result) {" & vbcrlf
     response.write "           displayValidAddressList(result);" & vbcrlf
     response.write "        });" & vbcrlf
     response.write "     }" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     if(lcl_streetnumber != '' || lcl_streetname != '0000') {" & vbcrlf
     response.write "        lcl_success = validateAddress();" & vbcrlf
     response.write "        if(! lcl_success) {" & vbcrlf
     response.write "           FinalCheck('NOT FOUND',1);" & vbcrlf
     response.write "        }" & vbcrlf
     response.write "     }" & vbcrlf
     response.write "  }" & vbcrlf
     response.write "}" & vbcrlf
    'END: Check Address ---------------------------------------------------

    'BEGIN: Validate Address ----------------------------------------------
     response.write "function validateAddress() {" & vbcrlf
     response.write "  clearMsg('residentstreetnumber');" & vbcrlf
     response.write "  clearMsg('streetaddress');" & vbcrlf
     response.write "  clearMsg('validateAddress');" & vbcrlf

    'Remove any extra spaces
     response.write "  $('#residentstreetnumber').val(jQuery.trim($('#residentstreetnumber').val()));" & vbcrlf

    'Check the number for non-numeric values
     response.write "  if($('#residentstreetnumber').val() != '') {" & vbcrlf
     response.write "     var rege = /^\d+$/;" & vbcrlf
     response.write "     var Ok = rege.exec($('#residentstreetnumber').val());" & vbcrlf

     response.write "     if ( ! Ok ) {" & vbcrlf
     response.write "         $('#residentstreetnumber').focus();" & vbcrlf
     response.write "         inlineMsg(document.getElementById(""residentstreetnumber"").id,'<strong>Invalid Value: </strong> The Street Number must be numeric.',10,'residentstreetnumber');" & vbcrlf
     response.write "         return false;" & vbcrlf
     response.write "     } else {" & vbcrlf

    'Check that they picked a street name
     response.write "        if ($('#streetaddress').val() == '0000') {" & vbcrlf
     response.write "            $('#streetaddress').focus();" & vbcrlf
     response.write "            inlineMsg(document.getElementById(""streetaddress"").id,'<strong>Required Field: </strong> Please select a street name from the list before validating the address.',10,'streetaddress');" & vbcrlf
     response.write "  	         return false;" & vbcrlf
     response.write "        } else {" & vbcrlf
     response.write "            return true;" & vbcrlf
     response.write "        }" & vbcrlf
     response.write "    	}" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     if ($('#streetaddress').val() == '0000') {" & vbcrlf

     if lcl_orghasfeature_large_address_list then
        response.write "         inlineMsg(document.getElementById(""validateAddress"").id,'<strong>Required Field: </strong> At least one address field must be entered before attempting to validate.',10,'validateAddress');" & vbcrlf
     else
        response.write "         inlineMsg(document.getElementById(""streetaddress"").id,'<strong>Required Field: </strong> An address must be entered before attempting to validate.',10,'validateAddress');" & vbcrlf
     end if

     response.write "         return false;" & vbcrlf

     if lcl_orghasfeature_large_address_list then
        response.write "     } else {" & vbcrlf
        response.write "         $('#residentstreetnumber').focus();" & vbcrlf
        response.write "         inlineMsg(document.getElementById(""residentstreetnumber"").id,'<strong>Required Field: </strong> Street Number',10,'residentstreetnumber');" & vbcrlf
        response.write "         return false;" & vbcrlf
     end if

     response.write "     }" & vbcrlf
     response.write "  }" & vbcrlf

     response.write "  return true;" & vbcrlf
     response.write "}" & vbcrlf
    'END: Validate Address ------------------------------------------------

    'BEGIN: Final Check ---------------------------------------------------
     response.write "function FinalCheck( sResults, iFalseCount ) {" & vbcrlf
     response.write "  if (sResults == 'FOUND CHECK') {" & vbcrlf
     response.write "      $('#validstreet').val('Y');" & vbcrlf
     response.write "      $('#validaddresslist').hide('slow');" & vbcrlf
     response.write "      enableDisableAddressFields('');" & vbcrlf
     response.write "  } else if (sResults == 'SUBMIT') {" & vbcrlf
     response.write "      if($('#ques_issue2').val() == '') {" & vbcrlf
     response.write "         var lcl_streetnumber = $('#residentstreetnumber').val();" & vbcrlf
     response.write "         var lcl_streetname   = $('#streetaddress').val();" & vbcrlf
     response.write "      }" & vbcrlf

     response.write "      if(iFalseCount > 0) {" & vbcrlf
     response.write "         return false;" & vbcrlf
     response.write "      } else {" & vbcrlf
     response.write "         document.getElementById(""maintain_dmt_section"").submit();" & vbcrlf
     response.write "         return true;" & vbcrlf
     response.write "      }" & vbcrlf
     response.write "  }else{" & vbcrlf
     response.write "      if ((sResults == 'FOUND SELECT')||(sResults == 'FOUND KEEP')) {" & vbcrlf
     response.write "           if (sResults == 'FOUND SELECT') {" & vbcrlf
     response.write "               $('#validstreet').val('Y');" & vbcrlf
     response.write "           }else{" & vbcrlf
     response.write "               $('#validstreet').val('N');" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "           $('#validaddresslist').hide('slow');" & vbcrlf
     response.write "           enableDisableAddressFields('');" & vbcrlf
     response.write "      }else{" & vbcrlf
     response.write "           if($('#ques_issue2').val() != '') {" & vbcrlf
     response.write "              $('#validaddresslist').hide('slow');" & vbcrlf
     response.write "              enableDisableAddressFields('');" & vbcrlf
     response.write "           } else {" & vbcrlf
     response.write "              $('#validaddresslist').show('slow');" & vbcrlf
     response.write "              enableDisableAddressFields('disabled');" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "      }" & vbcrlf
     response.write "  }" & vbcrlf
     response.write "}" & vbcrlf
    'END: Final Check -----------------------------------------------------

    'BEGIN: Enable/Disable Address Fields ---------------------------------
     response.write "function enableDisableAddressFields(iMode) {" & vbcrlf
     response.write "  var lcl_mode = '';" & vbcrlf

     response.write "  if(iMode != '') {" & vbcrlf
     response.write "     lcl_mode = iMode;" & vbcrlf
     response.write "  }" & vbcrlf

     'If mode = disabled then simply disable all of the fields/buttons
     'If mode = '' (enabled) then check to see which field(s) are populated to determine which field(s) to enable
     'response.write "  if(lcl_mode == 'disabled') {" & vbcrlf
     'response.write "     $('#residentstreetnumber').attr('disabled','disabled');" & vbcrlf
     'response.write "     $('#streetaddress').attr('disabled','disabled');" & vbcrlf
     'response.write "     $('#ques_issue2').attr('disabled','disabled');" & vbcrlf
     'response.write "     $('#validateAddress').attr('disabled','disabled');" & vbcrlf
     'response.write "     $('#importAddress').attr('disabled','disabled');" & vbcrlf
     'response.write "     $('#latitude').attr('disabled','');" & vbcrlf
     'response.write "     $('#longitude').attr('disabled','');" & vbcrlf
     'response.write "  } else {" & vbcrlf
     'response.write "     $('#residentstreetnumber').attr('disabled','');" & vbcrlf
     'response.write "     $('#streetaddress').attr('disabled','');" & vbcrlf
     'response.write "     $('#ques_issue2').attr('disabled','');" & vbcrlf
     'response.write "     $('#latitude').attr('disabled','');" & vbcrlf
     'response.write "     $('#longitude').attr('disabled','');" & vbcrlf

     'response.write "     if($('#ques_issue2').val() != '') {" & vbcrlf
     'response.write "        $('#validateAddress').attr('disabled','disabled');" & vbcrlf
     'response.write "        $('#importAddress').attr('disabled','disabled');" & vbcrlf
     'response.write "     } else {" & vbcrlf
     'response.write "        if($('#residentstreetnumber').val() != '') {" & vbcrlf
     'response.write "           $('#validateAddress').attr('disabled','');" & vbcrlf
     'response.write "           $('#importAddress').attr('disabled','');" & vbcrlf
     'response.write "           $('#latitude').attr('disabled','disabled');" & vbcrlf
     'response.write "           $('#longitude').attr('disabled','disabled');" & vbcrlf
     'response.write "        } else {" & vbcrlf
     'response.write "           if($('#streetaddress').val() == '0000') {" & vbcrlf
     'response.write "             $('#validateAddress').attr('disabled','disabled');" & vbcrlf
     'response.write "             $('#importAddress').attr('disabled','disabled');" & vbcrlf
     'response.write "           } else {" & vbcrlf
     'response.write "             $('#latitude').attr('disabled','disabled');" & vbcrlf
     'response.write "             $('#longitude').attr('disabled','disabled');" & vbcrlf
     'response.write "           }" & vbcrlf
     'response.write "        }" & vbcrlf
     'response.write "     }" & vbcrlf
     'response.write "  }" & vbcrlf

     response.write "  if(lcl_mode == 'disabled') {" & vbcrlf
     response.write "     $('#residentstreetnumber').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#streetaddress').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#ques_issue2').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#validateAddress').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#importAddress').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#latitude').prop('disabled','');" & vbcrlf
     response.write "     $('#longitude').prop('disabled','');" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     $('#residentstreetnumber').prop('disabled','');" & vbcrlf
     response.write "     $('#streetaddress').prop('disabled','');" & vbcrlf
     response.write "     $('#ques_issue2').prop('disabled','');" & vbcrlf
     response.write "     $('#latitude').prop('disabled','');" & vbcrlf
     response.write "     $('#longitude').prop('disabled','');" & vbcrlf

     response.write "     if($('#ques_issue2').val() != '') {" & vbcrlf
     response.write "        $('#validateAddress').prop('disabled','disabled');" & vbcrlf
     response.write "        $('#importAddress').prop('disabled','disabled');" & vbcrlf
     response.write "     } else {" & vbcrlf
     response.write "        if($('#residentstreetnumber').val() != '') {" & vbcrlf
     response.write "           $('#validateAddress').prop('disabled','');" & vbcrlf
     response.write "           $('#importAddress').prop('disabled','');" & vbcrlf
     response.write "           $('#latitude').prop('disabled','disabled');" & vbcrlf
     response.write "           $('#longitude').prop('disabled','disabled');" & vbcrlf
     response.write "        } else {" & vbcrlf
     response.write "           if($('#streetaddress').val() == '0000') {" & vbcrlf
     response.write "             $('#validateAddress').prop('disabled','disabled');" & vbcrlf
     response.write "             $('#importAddress').prop('disabled','disabled');" & vbcrlf
     response.write "           } else {" & vbcrlf
     response.write "             $('#latitude').prop('disabled','disabled');" & vbcrlf
     response.write "             $('#longitude').prop('disabled','disabled');" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "        }" & vbcrlf
     response.write "     }" & vbcrlf
     response.write "  }" & vbcrlf
     response.write "}" & vbcrlf
    'END: Enable/Disable Address Fields -----------------------------------

     if lcl_orghasfeature_large_address_list then
       'BEGIN: Display Valid Address List ------------------------------------
        response.write "function displayValidAddressList(iResult) {" & vbcrlf
        response.write "  var lcl_streetnumber = $('#residentstreetnumber').val();" & vbcrlf
        response.write "  var lcl_streetname   = $('#streetaddress').val();" & vbcrlf

       'Determine if the address is "valid" based on the records in egov_residentaddresses for the org
        response.write "  if(iResult == 'FOUND CHECK' || iResult == 'CANCEL') {" & vbcrlf
        response.write "     if(iResult == 'FOUND CHECK') {" & vbcrlf
        response.write "        displayScreenMsg('Address is Valid');" & vbcrlf
        response.write "        $('#validstreet').val('Y');" & vbcrlf
        response.write "     }" & vbcrlf
        response.write "     $('#validaddresslist').hide('slow');" & vbcrlf
        response.write "     enableDisableAddressFields('');" & vbcrlf
        response.write "  } else { " & vbcrlf
        response.write "     displayScreenMsg('Invalid Address');" & vbcrlf
        response.write "     $('#validstreet').val('N');" & vbcrlf
        response.write "     $('#oldstnumber').val(lcl_streetnumber);" & vbcrlf
        response.write "     $('#stname').val(lcl_streetname);" & vbcrlf

       'Display the valid address section and build the list of valid addresses
        response.write "     enableDisableAddressFields('disabled');" & vbcrlf

        response.write "     $('#validaddresslist').show('slow', function() {" & vbcrlf
        response.write "        $.post('checkaddress.asp', {" & vbcrlf
        response.write "           addresstype: '" & lcl_addresstype & "'," & vbcrlf
        response.write "           stnumber:   lcl_streetnumber," & vbcrlf
        response.write "           stname:     lcl_streetname," & vbcrlf
        response.write "           returntype: 'DISPLAY_OPTIONS'" & vbcrlf
        response.write "           }, function(result) {" & vbcrlf
        response.write "              $('#addresspicklist').html(result);" & vbcrlf
        response.write "        });" & vbcrlf
        response.write "     });" & vbcrlf
        response.write "  }" & vbcrlf
        response.write "}" & vbcrlf
       'END: Display Valid Address List --------------------------------------

       'BEGIN: Do Select -----------------------------------------------------
        response.write "function doSelect() {" & vbcrlf
        'response.write "  if($('#stnumber').attr('selectedIndex') < 0) {" & vbcrlf
        response.write "  if($('#stnumber').prop('selectedIndex') < 0) {" & vbcrlf
        response.write "     inlineMsg(document.getElementById(""stnumber"").id,'<strong>Required Field Missing: </strong> Please select a valid address first.',10,'stnumber');" & vbcrlf
        response.write "     return false;" & vbcrlf
        response.write "  }" & vbcrlf

        response.write "  clearScreenMsg();" & vbcrlf
        response.write "  clearMsg('stnumber');" & vbcrlf
        response.write "  $('#residentstreetnumber').val($('#stnumber').val());" & vbcrlf
        response.write "  $('#ques_issue2').val('');" & vbcrlf
        response.write "  FinalCheck('FOUND SELECT',0);" & vbcrlf
        response.write "}" & vbcrlf
       'END: Do Select -------------------------------------------------------

       'BEGIN: Cancel Pick ---------------------------------------------------
        response.write "function cancelPick() {" & vbcrlf
        response.write "  clearScreenMsg();" & vbcrlf
        response.write "  clearMsg('stnumber');" & vbcrlf
        response.write "  displayValidAddressList('CANCEL');" & vbcrlf
        response.write "}" & vbcrlf
       'END: Cancel Pick -----------------------------------------------------

       'BEGIN: Do Keep -------------------------------------------------------
        response.write "function doKeep() {" & vbcrlf
        response.write "  var lcl_streetnumber = $('#oldstnumber').val();" & vbcrlf
        response.write "  var lcl_streetname   = $('#stname').val();" & vbcrlf
        response.write "  var lcl_streetaddress = '';" & vbcrlf

        response.write "  if(lcl_streetnumber != '') {" & vbcrlf
        response.write "     lcl_streetaddress = lcl_streetnumber;" & vbcrlf
        response.write "  }" & vbcrlf

        response.write "  if(lcl_streetname != '') {" & vbcrlf
        response.write "     if(lcl_streetaddress != '') {" & vbcrlf
        response.write "        lcl_streetaddress += ' ';" & vbcrlf
        response.write "        lcl_streetaddress += lcl_streetname;" & vbcrlf
        response.write "     } else {" & vbcrlf
        response.write "        lcl_streetaddress = lcl_streetname;" & vbcrlf
        response.write "     }" & vbcrlf
        response.write "  }" & vbcrlf

        response.write "  $('#ques_issue2').val(lcl_streetaddress);" & vbcrlf
        response.write "  $('#residentstreetnumber').val('');" & vbcrlf
        'response.write "  $('#streetaddress').attr('selectedIndex',0);" & vbcrlf
        response.write "  $('#streetaddress').val('');" & vbcrlf
        response.write "  $('#streetaddress').prop('selectedIndex',0);" & vbcrlf
        response.write "  FinalCheck('FOUND KEEP',0);" & vbcrlf
        response.write "}" & vbcrlf
       'END: Do Keep ---------------------------------------------------------
     end if
  end if
 'END: Check for the "issue location" feature --------------------------------

 'BEGIN: Validate Fields ------------------------------------------------------
  response.write "function validateFields() {" & vbcrlf
  response.write "  var lcl_false_count = 0;" & vbcrlf

 'Need to set the address to the proper field based on feature(s) turned on for the org.
 'Check for "issue location" and "large address" features
  if lcl_orghasfeature_issue_location AND lcl_addressfield_exists then
     if lcl_orghasfeature_large_address_list then
        lcl_addresstype = "LARGE"
     else
        lcl_addresstype = ""
     end if

    'Validate the street number and name entered to determine if it is a valid address in the system for the org
     response.write "  if($('#ques_issue2').val() == '') {" & vbcrlf

'     if lcl_orghasfeature_large_address_list then

        response.write "     lcl_success = validateAddress();" & vbcrlf
        response.write "     if(lcl_success) {" & vbcrlf
        response.write "        var lcl_streetnumber = $('#residentstreetnumber').val();" & vbcrlf
        response.write "        var lcl_streetname   = $('#streetaddress').val();" & vbcrlf

        response.write "        $.post('checkaddress.asp', {" & vbcrlf
        response.write "           addresstype: '" & lcl_addresstype & "',"
        response.write "           stnumber:    lcl_streetnumber," & vbcrlf
        response.write "           stname:      lcl_streetname," & vbcrlf
        response.write "           returntype:  'CHECK'" & vbcrlf
        response.write "         }, function(result) {" & vbcrlf
        response.write "           if(result == 'NOT FOUND') {" & vbcrlf
        response.write "              displayValidAddressList(result);" & vbcrlf
        response.write "              return false;" & vbcrlf
        response.write "           } else {" & vbcrlf
        response.write "              FinalCheck('SUBMIT',0);" & vbcrlf
        response.write "           }" & vbcrlf
        response.write "        });" & vbcrlf
        response.write "     } else {" & vbcrlf
        response.write "       return false;" & vbcrlf
        response.write "     }" & vbcrlf
'     else
'        response.write "     FinalCheck('SUBMIT',0);" & vbcrlf
'     end if

     response.write "  } else {" & vbcrlf
     response.write "     FinalCheck('SUBMIT',0);" & vbcrlf
     response.write "  }" & vbcrlf

  else
     response.write "  if(lcl_false_count > 0) {" & vbcrlf
     response.write "     return false;" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     document.getElementById(""maintain_dmt_section"").submit();" & vbcrlf
     response.write "     return true;" & vbcrlf
     response.write "  } " & vbcrlf
  end if

  response.write "}" & vbcrlf
 'END: Validate Fields --------------------------------------------------------
%>

function editURL(iRowCount, iLineNum, iAction) {
  var sRowCount
  var sLineNum
  var sAction
  var i                = 0;
  var lcl_total_urls   = new Number(0);
  var lcl_url_disabled = false;
  var lcl_fieldvalue   = '';
  var lcl_fieldtype    = '';
  var lcl_displayvalue = '';
  var lcl_error        = 0;
  var lcl_button_label = '';
  var lcl_field_label  = '';
  var lcl_help_url_displayvalue_mouseover = ' onMouseOver="tooltip.show(\'If a DISPLAY VALUE is not entered then the URL will be used as the clickable link.\');"'
  var lcl_help_url_displayvalue_mouseout  = ' onMouseOut="tooltip.hide();"'

  if((iRowCount == '') || (iRowCount == undefined)) {
      sRowCount = '0';
  } else {
      sRowCount = iRowCount;
  }

  if((iLineNum == '') || (iLineNum == undefined)) {
      sLineNum = '0';
  } else {
      sLineNum = iLineNum;
  }

  if((iAction == '') || (iAction == undefined)) {
      sAction = 'CANCEL';
  } else {
      sAction = iAction.toUpperCase();
  }

  if($('#total_urls_' + sRowCount).val() != '') {
     lcl_total_urls = Number($('#total_urls_' + sRowCount).val());
  }

  if(sAction == "SAVE") {
     //Loop through and modify "fieldvalue" based on all rows NOT checked for removal.
     if(lcl_total_urls > 0) {
        for (i=1; i <= lcl_total_urls; i++) {
           if($('#removeURL' + sRowCount + '_' + i).is(':checked') == false) {
              if($('#website_url' + sRowCount + '_' + i).val() == '') {
                 $('#website_url' + sRowCount + '_' + i).focus();
                 inlineMsg(document.getElementById('website_url' + sRowCount + '_' + i).id,'<strong>Required Field</strong>',10,'website_url' + sRowCount + '_' + i);
                 lcl_error = lcl_error + 1;
              } else {
                 if(lcl_fieldvalue != '') {
                    lcl_fieldvalue   += ',';
                    lcl_displayvalue += '<br />';
                 }

                 //build the database value
                 lcl_fieldvalue += '[';
                 lcl_fieldvalue += $('#website_url'  + sRowCount + '_' + i).val();
                 lcl_fieldvalue += ']<';
                 lcl_fieldvalue += $('#website_text' + sRowCount + '_' + i).val();
                 lcl_fieldvalue += '>';

                 //build the display value
                 lcl_displayvalue += '<a href="';
                 lcl_displayvalue += $('#website_url'  + sRowCount + '_' + i).val();
                 lcl_displayvalue += '" target="_blank">';

                 if($('#website_text' + sRowCount + '_' + i).val() != '') {
                    lcl_displayvalue += $('#website_text' + sRowCount + '_' + i).val();
                 } else {
                    lcl_displayvalue += $('#website_url'  + sRowCount + '_' + i).val();
                 }

                 lcl_displayvalue += '</a>';
              }
           }
        }

        if(lcl_error == 0) {
           $('#maintain_url' + sRowCount).hide('slow',function(){
             $('#dm_fieldvalue' + sRowCount).val(lcl_fieldvalue);
             $('#dm_fieldvalue' + sRowCount + '_display').html(lcl_displayvalue);

             $('#editURLButton' + sRowCount).show('slow',function(){
               $('#dm_fieldvalue' + sRowCount + '_display').show('slow');
             });
           });
        }
     }
  } else if(sAction == "ADD") {
    lcl_fieldtype = $('#fieldtype' + sRowCount).val();

    if(lcl_fieldtype.indexOf('WEBSITE') > -1) {
       lcl_button_label = 'Website';
       lcl_field_label  = 'URL';
    } else {
       lcl_button_label = 'Email';
       lcl_field_label  = 'Email';
    }

    var mytbl = document.getElementById('website_table' + sRowCount);
    mytbl     = mytbl.insertRow(lcl_total_urls);

    //Build the cell for the new row.
    var a           = mytbl.insertCell(0);
    var b           = mytbl.insertCell(1);
    var c           = mytbl.insertCell(2);
    var lcl_html_a = '';
    var lcl_html_b = '';
    var lcl_html_c = '';

    //Increase the total rows by one.
    lcl_total_urls = lcl_total_urls + 1;

    lcl_html_a += lcl_field_label + ': ';
    lcl_html_a += '<input type="text" name="website_url'            + sRowCount + '_' + lcl_total_urls + '" id="website_url'          + sRowCount + '_' + lcl_total_urls + '" value="" size="40" maxlength="50" onchange="clearMsg(\'website_url' + sRowCount + '_' + lcl_total_urls + '\');" />';
    lcl_html_a += '<input type="hidden" name="original_website_url' + sRowCount + '_' + lcl_total_urls + '" id="original_website_url' + sRowCount + '_' + lcl_total_urls + '" value="" size="40" maxlength="50" />';

    lcl_html_b += 'Display Value: ';
    lcl_html_b += '<input type="text" name="website_text'            + sRowCount + '_' + lcl_total_urls + '" id="website_text'          + sRowCount + '_' + lcl_total_urls + '" value="" size="30" maxlength="50" />';
    lcl_html_b += '<input type="hidden" name="original_website_text' + sRowCount + '_' + lcl_total_urls + '" id="original_website_text' + sRowCount + '_' + lcl_total_urls + '" value="" size="30" maxlength="50" />&nbsp;';
    lcl_html_b += '<img src="../images/help_graybg.jpg" name="helpFeature_url' + sRowCount + '_' + lcl_total_urls + '" id="helpFeature_url' + sRowCount + '_' + lcl_total_urls + '" class="helpOption"' + lcl_help_url_displayvalue_mouseover + lcl_help_url_displayvalue_mouseout + ' />'

    lcl_html_c += '<input type="checkbox" name="removeURL' + sRowCount + '_' + lcl_total_urls + '" id="removeURL' + sRowCount + '_' + lcl_total_urls + '" value="Y" onclick="editURL(\'' + sRowCount + '\',\'' + sRowCount + '\',\'DELETE\');" /> Remove';

    a.style.whiteSpace = 'nowrap';
    b.style.whiteSpace = 'nowrap';
    c.style.whiteSpace = 'nowrap';

    a.innerHTML = lcl_html_a;
    b.innerHTML = lcl_html_b;
    c.innerHTML = lcl_html_c;

    $('#total_urls_' + sRowCount).val(lcl_total_urls);

  } else if(sAction == "CANCEL") {
     //uncheck and re-enable fields if they have been checked for deletion.
     if(sLineNum > 0) {
        for (i=0; i<=sLineNum; i++) {
           $('#removeURL'    + sRowCount + '_' + i).prop({checked: ''});
           $('#website_url'  + sRowCount + '_' + i).prop({disabled: false});
           $('#website_text' + sRowCount + '_' + i).prop({disabled: false});

           $('#website_url'  + sRowCount + '_' + i).val($('#original_website_url'  + sRowCount + '_' + i).val());
           $('#website_text' + sRowCount + '_' + i).val($('#original_website_text' + sRowCount + '_' + i).val());
        }
     }

     $('#maintain_url' + sRowCount).hide('slow',function(){
       $('#editURLButton' + sRowCount).show('slow',function(){
         $('#dm_fieldvalue' + sRowCount + '_display').show('slow');
       });
     });
  } else if(sAction == "DELETE") {
     if($('#removeURL' + sRowCount + '_' + sLineNum).is(':checked') == true) {
        lcl_url_disabled = true;
     }

     $('#website_url'  + sRowCount + '_' + sLineNum).prop({disabled: lcl_url_disabled});
     $('#website_text' + sRowCount + '_' + sLineNum).prop({disabled: lcl_url_disabled});

  } else {  //EDIT
     lcl_fieldtype  = $('#fieldtype'     + sRowCount).val();
     lcl_fieldvalue = $('#dm_fieldvalue' + sRowCount).val();

     $.post('build_maintainurl_section.asp', {
        rowCount:   sRowCount,
        fieldtype:  lcl_fieldtype,
        fieldvalue: lcl_fieldvalue
     }, function(result) {
        $('#maintain_url' + sRowCount).html(result);

        $('#dm_fieldvalue' + sRowCount + '_display').hide('slow',function(){
          $('#editURLButton' + sRowCount).hide('slow',function(){
            $('#maintain_url' + sRowCount).slideDown('slow');
          });
        });
     });
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

  eval('window.open("' + lcl_url + '", "_blank", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
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

<style>
textarea {
  width:  500px;
  height: 300px;
}
<% if lcl_sectiontype = "HOURS" then %>
.hoursTable {
  width:              350px;
  font-size:          11px;
  color:              #000000;
  margin-bottom:      5px;
  background-color:   #ffffff;
  border:             1pt solid #c0c0c0;
 	-moz-border-radius: 2% 2% 2% 2%;
  border-radius:      2% 2% 2% 2%;
}

.hours_title_left {
   width:        80px;
   font-weight:  bold;
   padding-left: 5px;
}

.hours_title_right {
   text-align: right;
}

.hours_edit {
   background-color: #efefef;
   border-bottom-left-radius:      2% 2%;
   border-bottom-right-radius:     2% 2%;
  	-moz-border-radius-bottomleft:  2% 2%;
  	-moz-border-radius-bottomright: 2% 2%;
}
<% end if %>

.address_fieldset {
   border:                1pt solid #808080;
   border-radius:         5px;
   -moz-border-radius:    5px;
   -webkit-border-radius: 5px;
}

#validaddresslist {
   border:                1pt solid #c0c0c0;
   border-radius:         6px;
  	-moz-border-radius:    6px;
   -webkit-border-radius: 6px;
   background-color:   #efefef;
   margin-top:         4px;
}

#validaddresslist legend {
   border:           1pt solid #c0c0c0;
   border-radius:    4px;
  	-moz-border-radius:    4px;
   -webkit-border-radius: 4px;
   background-color: #ffffff;
   color:            #ff0000;
   padding-left:     4px;
   padding-right:    4px;
}

div#addresspicklist {
  border-radius: 6px;
   -moz-border-radius:    6px;
   -webkit-border-radius: 5px;
}

.maintain_url {
   border:           1pt solid #000000;
   background-color: #c0c0c0;
   padding:          4px;
   color:            #000000;
   font-size:        10pt;
   display:          none;
}

.url_displaytext {
   font-size: 10pt;
   color:     #000000;
}

#screenMsg {
   text-align:  right;
   color:       #ff0000;
   font-size:   10pt;
   font-weight: bold;
}
</style>
</head>
<body onload="<%=lcl_onload%>">
<%
  response.write "<form name=""maintain_dmt_section"" id=""maintain_dmt_section"" method=""post"" action=""maintain_dmt_section_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""dmid"" id=""dmid"" value=""" & lcl_dmid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" value=""" & lcl_dm_typeid & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""sectionid"" id=""sectionid"" value=""" & lcl_sectionid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""screen_mode"" value=""EDIT"" size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""importAddressFields"" id=""importAddressFields"" size=""1"" maxlength=""1"" value=""" & lcl_importAddressFields & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" size=""5"" maxlength=""50"" value=""" & lcl_feature & """ />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>Maintain " & lcl_sectionname & "</strong></font>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf
                            displayButtons "TOP", lcl_sectiontype

                            lcl_dm_sectionid         = ""
                            lcl_sectionIsActive      = ""
                            lcl_isAccountInfoSection = False
                            lcl_sectionlocation      = ""
                            lcl_sectionorder         = ""
                            lcl_totalDraggableItems  = ""
                            lcl_section_mode         = "EDIT"

                           'Determine which section to edit
                            buildSection lcl_dmid, _
                                         lcl_dm_typeid, _
                                         lcl_sectionid, _
                                         lcl_orghasfeature_issue_location, _
                                         lcl_orghasfeature_large_address_list, _
                                         lcl_importAddressFields, _
                                         lcl_importstreet_number, _
                                         lcl_importstreet_address

                            displayButtons "BOTTOM", lcl_sectiontype

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
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

<!--#include file="../admin_footer.asp"-->

</body>
</html>
<%
'------------------------------------------------------------------------------
sub buildSection(iDMID, iDMTypeID, iSectionID, iOrgHasFeature_issueLocation, _
                 iOrgHasFeature_largeAddressList, iImportAddressFields, iStreetNumber, iStreetAddress)

 'Get all of the section info
  getSectionInfo iSectionID, lcl_sectionname, lcl_sectiontype

 'Do NOT display the "Default" section name

  if lcl_sectionname <> "" then
     if ucase(lcl_sectionname) = "DEFAULT" then
        lcl_sectionname = "&nbsp;"
     end if
  end if

  response.write "<div class=""section"" id=""section" & iSectionID & """>" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"" class=""section-content"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf

 'Get all of the DM Type Fields for the DM Type Section the user is working with.
 'Perform an OUTER JOIN to the DM Values table so that we can get all of the "value ids" so we
 '  know where to store the data for each field.  If a "value" record does NOT exist then we know
 '  we must create one in order to store the data for the field.
  sSQL = "SELECT dmtf.dm_fieldid, "
  sSQL = sSQL & " dmtf.dm_typeid, "
  sSQL = sSQL & " dmtf.dm_sectionid, "
  sSQL = sSQL & " dmtf.section_fieldid, "
  sSQL = sSQL & " dmtf.displayFieldName, "
  sSQL = sSQL & " dmtf.orgid, "
  sSQL = sSQL & " dmsf.fieldtype, "
  sSQL = sSQL & " dmsf.fieldname, "
  sSQL = sSQL & " dmsf.isMultiLine, "
  sSQL = sSQL & " dmsf.hasAddLinkButton, "
  sSQL = sSQL & " dmv.dmid, "
  sSQL = sSQL & " dmv.dm_valueid, "

  if iDMID <> "" then
     sSQL = sSQL & " dmv.fieldvalue "
  else
     sSQL = ssQL & " '' as fieldvalue "
  end if

  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &                 " ON dmtf.section_fieldid = dmsf.section_fieldid "
  sSQL = sSQL &                 " AND dmsf.isActive = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &                 " ON dmtf.dm_sectionid = dmts.dm_sectionid "
  sSQL = sSQL &                 " AND dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL &                 " AND dmts.sectionid = " & iSectionID

  if iDMID <> "" then
     sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_values dmv "
     sSQL = sSQL &                      " ON dmtf.dm_fieldid = dmv.dm_fieldid "
     sSQL = sSQL &                      " AND dmv.dm_typeid = " & iDMTypeID
     sSQL = sSQL &                      " AND dmv.dmid = " & iDMID
  end if

  sSQL = sSQL & " WHERE dmtf.dm_sectionid IN (select distinct dmts.dm_sectionid "
  sSQL = sSQL &                             " from egov_dm_types_sections dmts "
  sSQL = sSQL &                             " where dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL &                             " and dmts.sectionid = " & iSectionID & ") "
  sSQL = sSQL & " ORDER BY dmsf.displayOrder, dmtf.resultsOrder "

 'Determine how to display the section for editing
  if lcl_sectiontype = "HOURS" then
     editSection_hours sSQL
  else
     editSection sSQL, _
                 iOrgHasFeature_issueLocation, _
                 iOrgHasFeature_largeAddressList, _
                 iImportAddressFields, _
                 iStreetNumber, _
                 iStreetAddress, _
                 iDMID
  end if

  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub editSection(iSQL, iOrgHasFeature_issueLocation, iOrgHasFeature_largeAddressList, _
                iImportAddressFields, p_street_number, p_street_address, iDMID)

  set oEditDMValues = Server.CreateObject("ADODB.Recordset")
  oEditDMValues.Open iSQL, Application("DSN"), 3, 1

  if not oEditDMValues.eof then
     iRowCount           = 0
     lcl_field_maxlength = "4000"

    'Check to see if the user wants to import the address fields because the street number/name has been changed.
     if iImportAddressFields = "Y" then
        if iOrgHasFeature_largeAddressList then
           GetAddressInfoLarge oEditDMValues("orgid"), _
                               p_street_number, _
                               p_street_address, _
                               sNumber, _
                               sPrefix, _
                               sAddress, _
                               sSuffix, _
                               sDirection, _
                               sCity, _
                               sState, _
                               sZip, _
                               sValidStreet, _
                               sLatitude, _
                               sLongitude
        else
    		    	GetAddressInfo p_street_address, _
                          sNumber, _
                          sPrefix, _
                          sAddress, _
                          sSuffix, _
                          sDirection, _
                          sCity, _
                          sState, _
                          sZip, _
                          sValidStreet, _
                          sLatitude, _
                          sLongitude
        end if

       'These fields are NOT to be overridden during the import
        sSortStreetName = request("sortstreetname")
     else
       'Get the address info from the DM Data itself
        getAddressInfo_byDMID iDMID, _
                              sNumber, _
                              sPrefix, _
                              sAddress, _
                              sSuffix, _
                              sDirection, _
                              sCity, _
                              sState, _
                              sZip, _
                              sValidStreet, _
                              sLatitude, _
                              sLongitude

       'If the org does NOT have the "issue location" feature turned on then all addresses entered are considered "Invalid"
        if not iOrgHasFeature_IssueLocation then
           sValidStreet = "N"
        end if
     end if

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">" & vbcrlf

     do while not oEditDMValues.eof
        iRowCount            = iRowCount + 1
        lcl_orgid            = oEditDMValues("orgid")
        lcl_fieldname        = oEditDMValues("fieldname")
        lcl_fieldvalue       = oEditDMValues("fieldvalue")
        lcl_fieldtype        = oEditDMValues("fieldtype")
        lcl_displayFieldName = oEditDMValues("displayFieldName")
        lcl_displayMap       = getDMTypeDisplayMap(oEditDMValues("dm_typeid"))

       'Determine if the fieldname is displayed
        'if lcl_displayFieldName then
        if lcl_fieldname = "" then
           lcl_fieldname = "&nbsp;"
        end if

        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <td nowrap=""nowrap"">" & lcl_fieldname & "</td>" & vbcrlf
        response.write "      <td>" & vbcrlf

       'Check for "specialty" fields
        if lcl_fieldtype = "ADDRESS" then
           response.write "<fieldset id=""address_fieldset"" class=""address_fieldset"">" & vbcrlf

          'Determine how to pull the address info.
          '- Check to see if the org has the "issue location" feature on.
          '- If "yes" then check to see if the org has the "large address list" feature on.
           if iOrgHasFeature_IssueLocation then
              if sValidStreet = "Y" then
                 if iOrgHasFeature_LargeAddressList then
                    lcl_street_name = buildStreetAddress("", sPrefix, sAddress, sSuffix, sDirection)

                    DisplayLargeAddressList lcl_orgid, sNumber, sPrefix, sAddress, sSuffix, sDirection
                 else
                    DisplayAddress lcl_orgid, sNumber, sAddress
                 end if

                 lcl_display_other_address = ""
                 lcl_displayAddress        = sNumber & " " & sAddress
              else
                 if iOrgHasFeature_LargeAddressList then
                    DisplayLargeAddressList lcl_orgid, "", "", "", "", ""
                 else
                    DisplayAddress lcl_orgid, "", ""
                 end if

                 lcl_display_other_address = sNumber

                 if lcl_display_other_address <> "" then
                    lcl_display_other_address = lcl_display_other_address & " " & sAddress
                 else
                    lcl_display_other_address = sAddress
                 end if

                 lcl_displayAddress = lcl_display_other_address
              end if

              'if iDMID <> 0 then
                 'response.write "<input type=""button"" id=""importAddress"" class=""button"" value=""Import Address Fields"" onclick=""getAddressFields()"" />" & vbcrlf
              '   response.write "<input type=""button"" id=""importAddress"" class=""button"" value=""Import Address Fields"" />" & vbcrlf
              'end if

              response.write "<br /> - Or Other Not Listed - <br /> " & vbcrlf
              'lcl_address_onchange = " onchange=""save_address();checkAddressButtons()"""
              lcl_address_onchange = ""

           else
              lcl_display_other_address = sAddress
              lcl_displayAddress        = sAddress
              lcl_address_onchange      = ""
           end if

           response.write "          <input type=""text"" name=""ques_issue2"" id=""ques_issue2"" class=""correctionstextbox"" size=""50"" maxlength=""75"" value=""" & lcl_display_other_address & """" & lcl_address_onchange & " />" & vbcrlf
           response.write "          <input type=""hidden"" name=""validstreet"" id=""validstreet"" value=""" & sValidStreet & """ />" & vbcrlf
           response.write "    <br /><input type=""" & lcl_hidden & """ name=""dm_fieldvalue" & iRowCount & """ id=""dm_fieldvalue" & iRowCount & """ value=""" & lcl_displayAddress & """ size=""50"" maxlength=""" & lcl_field_maxlength & """ />" & vbcrlf

          'Only build the "invalid address" section if the org has the "issue location" and "large address list" features
           if iOrgHasFeature_IssueLocation AND iOrgHasFeature_LargeAddressList then
              response.write "    <fieldset id=""validaddresslist"">" & vbcrlf
              response.write "      <legend>Invalid Address</legend>" & vbcrlf
              response.write "      <p>The address you entered does not match any in the system. " & vbcrlf
              response.write "      You can select a valid address from the list, or if you are certain the address you entered is correct " & vbcrlf
              response.write "      click the ""Use the address I entered"" button, to continue.</p>" & vbcrlf
              'response.write "      <form name=""frmAddress"" action=""addresspicker.asp"" method=""post"">" & vbcrlf
              response.write "      			<strong>The address you entered</strong><br />" & vbcrlf
              response.write "      			<input type=""text"" name=""oldstnumber"" id=""oldstnumber"" value="""" disabled=""disabled"" size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
              response.write "      			<input type=""text"" name=""stname"" id=""stname"" value="""" disabled=""disabled"" size=""50"" maxlength=""50"" />" & vbcrlf
              response.write "      			<div id=""addresspicklist""></div>" & vbcrlf
              response.write "      			<input type=""button"" name=""validpick"" id=""validpick"" value=""Use the valid address selected"" class=""button"" onclick=""doSelect();"" />" & vbcrlf
              response.write "      			<input type=""button"" name=""invalidpick"" id=""invalidpick"" value=""Use the address I entered"" class=""button"" onclick=""doKeep();"" />" & vbcrlf
              response.write "      			<input type=""button"" name=""cancelpick"" id=""cancelpick"" value=""Cancel"" class=""button"" onclick=""cancelPick();"" />" & vbcrlf
              'response.write "      		</form>" & vbcrlf
              response.write "    </fieldset>" & vbcrlf
           end if

           response.write "</fieldset>" & vbcrlf

          'Latitude/Longitude Instructions ------------------------------------
           if lcl_displayMap then
              response.write "For Mapping, enter the latitude and longitude.<br />" & vbcrlf
              response.write """Valid"" addresses may auto-populate these values when available.<br />" & vbcrlf
              'response.write "If not, you can search for latitude and longitude values <a href=""javascript:openWin('http://www.batchgeocode.com/lookup/', 1000, 600);"">here.</a>" & vbcrlf
              response.write "If not, you can search for latitude and longitude values <a href=""javascript:openWin('datamgr_geocode_finder.asp?popup=Y&fname=maintain_dmt_section&lat=latitude&long=longitude', 800, 700);"">here.</a>" & vbcrlf
           end if

        elseif lcl_fieldtype = "LATITUDE" OR lcl_fieldtype = "LONGITUDE" then

          'If this is an import then the LATITUDE and LONGITUDE fields must be overridden
           'lcl_fieldvalue = oMPTFields("fieldvalue")

           if lcl_fieldtype = "LATITUDE" then
              lcl_fieldname = "latitude"

              if iImportAddressFields = "Y" then
                 lcl_fieldvalue = sLatitude
              'else
              '   lcl_fieldvalue = request("latitude")
              end if

           elseif lcl_fieldtype = "LONGITUDE" then
              lcl_fieldname = "longitude"

              if iImportAddressFields = "Y" then
                 lcl_fieldvalue = sLongitude
              'else
              '   lcl_fieldvalue = request("longitude")
              end if

           end if

           response.write "          <input type=""text"" name=""" & lcl_fieldname & """ id=""" & lcl_fieldname & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""" & lcl_field_maxlength & """ onchange=""clearMsg('" & lcl_fieldname & "');"" />" & vbcrlf
           response.write "          <input type=""" & lcl_hidden & """ name=""dm_fieldvalue" & iRowCount & """ id=""dm_fieldvalue" & iRowCount & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""" & lcl_field_maxlength & """ onchange=""clearMsg('dm_fieldvalue" & iRowCount & "');"" />" & vbcrlf

        else
           if instr(lcl_fieldtype,"WEBSITE") > 0 OR instr(lcl_fieldtype,"EMAIL") > 0 then
              lcl_display_value = buildURLDisplayValue(lcl_fieldtype, lcl_fieldvalue)

              response.write "<input type=""hidden"" name=""dm_fieldvalue" & iRowCount & """ id=""dm_fieldvalue" & iRowCount & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""" & lcl_field_maxlength & """ onchange=""clearMsg('dm_fieldvalue" & iRowCount & "');""" & lcl_style_div & " />" & vbcrlf
              response.write "<table border=""0"" class=""url_displaytext"">" & vbcrlf
              response.write "  <tr valign=""top"">" & vbcrlf
              response.write "      <td>" & vbcrlf
              response.write "          <input type=""button"" name=""editURLButton" & iRowCount & """ id=""editURLButton" & iRowCount & """ value=""Edit"" class=""button"" onclick=""editURL('" & iRowCount & "','','EDIT');"" />" & vbcrlf
              response.write "      </td>" & vbcrlf
              response.write "      <td id=""dm_fieldvalue" & iRowCount & "_display"">" & lcl_display_value & "</td>" & vbcrlf
              response.write "  </tr>" & vbcrlf
              response.write "</table>" & vbcrlf
              response.write "<div id=""maintain_url" & iRowCount & """ class=""maintain_url""></div>" & vbcrlf
           else
             'Determine if this field is considered a "multi-line"
              if oEditDMValues("isMultiLine") then
                 response.write "<textarea name=""dm_fieldvalue" & iRowCount & """ id=""dm_fieldvalue" & iRowCount & """ rows=""5"" cols=""49"" maxlength=""" & lcl_field_maxlength & """>" & lcl_fieldvalue & "</textarea>" & vbcrlf
              else
                 response.write "<input type=""text"" name=""dm_fieldvalue" & iRowCount & """ id=""dm_fieldvalue" & iRowCount & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""" & lcl_field_maxlength & """ onchange=""clearMsg('dm_fieldvalue" & iRowCount & "');""" & lcl_style & " />" & vbcrlf
              end if
           end if
        end if

        response.write "          <input type=""hidden"" name=""dm_valueid"   & iRowCount & """ id=""dm_valueid"   & iRowCount & """ value=""" & oEditDMValues("dm_valueid")   & """ size=""5"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_fieldid"   & iRowCount & """ id=""dm_fieldid"   & iRowCount & """ value=""" & oEditDMValues("dm_fieldid")   & """ size=""5"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_sectionid" & iRowCount & """ id=""dm_sectionid" & iRowCount & """ value=""" & oEditDMValues("dm_sectionid") & """ size=""5"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""hidden"" name=""fieldtype"    & iRowCount & """ id=""fieldtype"    & iRowCount & """ value=""" & lcl_fieldtype & """ size=""20"" maxlength=""100"" />" & vbcrlf

       'Determine if we are to display the "Add a Link"
        if oEditDMValues("hasAddLinkButton") then
           if instr(lcl_fieldtype,"WEBSITE") < 1 AND instr(lcl_fieldtype,"EMAIL") < 1 then
              response.write "          <input type=""button"" value=""Add a Link"" class=""button"" onclick=""doPicker('maintain_dmt_section.dm_fieldvalue" & iRowCount & "','Y','Y','Y','Y');"" />" & vbcrlf
           end if
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oEditDMValues.movenext
     loop

     response.write "</table>" & vbcrlf
  end if

  oEditDMValues.close
  set oEditDMValues = nothing

  response.write "<input type=""hidden"" name=""totalfields"" id=""totalfields"" value=""" & iRowCount & """ size=""5"" maxlength=""10"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub editSection_hours(iSQL)

  set oEditDMValues = Server.CreateObject("ADODB.Recordset")
  oEditDMValues.Open iSQL, Application("DSN"), 3, 1

  if not oEditDMValues.eof then
     iRowCount           = 0
     lcl_field_maxlength = "4000"

     do while not oEditDMValues.eof
        iRowCount      = iRowCount + 1
        lcl_fieldvalue = oEditDMValues("fieldvalue")
        lcl_fieldname  = "&nbsp;"

        if oEditDMValues("fieldname") <> "" then
           lcl_fieldname = "<strong>" & oEditDMValues("fieldname") & "</strong>"
        end if

       'If there is a fieldname AND it is to be HIDDEN then still show the label, but in ITALICS so the user knows it's hidden.
        if not oEditDMValues("displayFieldName") then
           lcl_fieldname = "<em>" & lcl_fieldname & "</em>"
        end if

       'Break the value into hours, minutes, and AM/PM values
        lcl_time_start    = ""
        lcl_time_end      = ""
        lcl_hours_start   = ""
        lcl_hours_end     = ""
        lcl_minutes_start = ""
        lcl_minutes_end   = ""
        lcl_ampm_start    = ""
        lcl_ampm_end      = ""
        lcl_other_value   = ""

        if  lcl_fieldvalue <> "" _
        AND instr(lcl_fieldvalue,":") > 0 _
        AND instr(lcl_fieldvalue," ") > 0 _
        AND instr(lcl_fieldvalue," - ") > 0 then
           lcl_hours = split(lcl_fieldvalue," - ")
           lcl_time_start = lcl_hours(0)
           lcl_time_end   = lcl_hours(1)

          'Break down the start time
           lcl_position_colon1 = instr(lcl_time_start,":")
           lcl_hours_start     = replace(LEFT(lcl_time_start, 2),":","")
           lcl_minutes_start   = MID(lcl_time_start, lcl_position_colon1+1, 2)
           lcl_ampm_start      = RIGHT(lcl_time_start, 2)

          'Break down the end time
           lcl_position_colon2 = instr(lcl_time_end,":")
           lcl_hours_end       = replace(LEFT(lcl_time_end, 2),":","")
           lcl_minutes_end     = MID(lcl_time_end, lcl_position_colon2+1, 2)
           lcl_ampm_end        = RIGHT(lcl_time_end, 2)
        else
          'If the value does not fit into the "time" dropdown lists then we know it's a custom entry
           lcl_other_value = lcl_fieldvalue
        end if

        response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""hoursTable"">" & vbcrlf
        response.write "  <tr valign=""top"">" & vbcrlf
        response.write "      <td class=""hours_title_left"">" & lcl_fieldname & "</td>" & vbcrlf
        response.write "      <td class=""hours_title_middle"" id=""display_fieldvalue" & iRowCount & """>" & lcl_fieldvalue & "</td>" & vbcrlf
        response.write "      <td class=""hours_title_right""><input type=""button"" name=""editHoursButton" & iRowCount & """ id=""editHoursButton" & iRowCount & """ class=""button"" value=""Edit"" /></td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "  <tr>" & vbrlf
        response.write "      <td colspan=""3"" class=""hours_edit"">" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_valueid"   & iRowCount & """ id=""dm_valueid"   & iRowCount & """ value=""" & oEditDMValues("dm_valueid")   & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_fieldid"   & iRowCount & """ id=""dm_fieldid"   & iRowCount & """ value=""" & oEditDMValues("dm_fieldid")   & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_sectionid" & iRowCount & """ id=""dm_sectionid" & iRowCount & """ value=""" & oEditDMValues("dm_sectionid") & """ />" & vbcrlf
        response.write "          <input type=""hidden"" name=""fieldvalue"   & iRowCount & """ id=""fieldvalue"   & iRowCount & """ value=""" & lcl_fieldvalue & """ size=""20"" maxlength=""4000"" />" & vbcrlf

        response.write "          <div name=""editHoursInfo" & iRowCount & """ id=""editHoursInfo" & iRowCount & """ class=""editHoursInfo"">" & vbcrlf
        response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
        response.write "            <tr>" & vbcrlf
        response.write "                <td rowspan=""2"">" & vbcrlf
        response.write "                    <input type=""radio"" name=""selection_time" & iRowCount & """ id=""selection_time" & iRowCount & """ value=""TIME"" />" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "                <td>Start:</td>" & vbcrlf
        response.write "                <td nowrap=""nowrap"" width=""100%"">" & vbcrlf
                                            buildHoursFields "START", "HOURS",   iRowCount, lcl_hours_start
                                            buildHoursFields "START", "MINUTES", iRowCount, lcl_minutes_start
                                            buildHoursFields "START", "AMPM",    iRowCount, lcl_ampm_start
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr>" & vbcrlf
        response.write "                <td>End:</td>" & vbcrlf
        response.write "                <td nowrap=""nowrap"" width=""100%"">" & vbcrlf
                                            buildHoursFields "END", "HOURS",   iRowCount, lcl_hours_end
                                            buildHoursFields "END", "MINUTES", iRowCount, lcl_minutes_end
                                            buildHoursFields "END", "AMPM",    iRowCount, lcl_ampm_end
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr>" & vbcrlf
        response.write "                <td>" & vbcrlf
        response.write "                    <input type=""radio"" name=""selection_time" & iRowCount & """ id=""selection_other" & iRowCount & """ value=""OTHER"" />" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "                <td colspan=""2"">" & vbcrlf
        response.write "                    <p>" & vbcrlf
        response.write "                    Other:<br />" & vbcrlf
        response.write "                    <input type=""text"" name=""time_other" & iRowCount & """ id=""time_other" & iRowCount & """ value=""" & lcl_other_value & """ size=""50"" maxlength=""100"" /><br />" & vbcrlf
        response.write "                    <em>(i.e. Closed, By Appointment Only, etc)</em>" & vbcrlf
        response.write "                    </p>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "            <tr>" & vbcrlf
        response.write "                <td colspan=""3"" align=""center"">" & vbcrlf
        response.write "                    <input type=""button"" name=""hoursSaveButton" & iRowCount & """ id=""hoursSaveButton" & iRowCount & """ class=""button"" value=""Save Changes"" />" & vbcrlf
        'response.write "                    <input type=""button"" name=""cancelButton" & iRowCount & """ id=""cancelButton" & iRowCount & """ class=""button"" value=""Cancel"" />" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
        response.write "          </table>" & vbcrlf
        response.write "          </div>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
        response.write "</table>" & vbcrlf

        oEditDMValues.movenext
     loop
  end if

  oEditDMValues.close
  set oEditDMValues = nothing

  response.write "<input type=""hidden"" name=""totalfields"" id=""totalfields"" value=""" & iRowCount & """ size=""5"" maxlength=""10"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
function getDMTypeID_byDMID(iDMID)

  lcl_return = 0

  if iDMID <> "" then
     sSQL = "SELECT dm_typeid "
     sSQL = sSQL & " FROM egov_dm_data "
     sSQL = sSQL & " WHERE dmid = " & iDMID

     set oGetDMTIDbyDMID = Server.CreateObject("ADODB.Recordset")
     oGetDMTIDbyDMID.Open sSQL, Application("DSN"), 3, 1

     if not oGetDMTIDbyDMID.eof then
        lcl_return = oGetDMTIDbyDMID("dm_typeid")
     end if

     oGetDMTIDbyDMID.close
     set oGetDMTIDbyDMID = nothing
  end if

  getDMTypeID_byDMID = lcl_return

end function

'------------------------------------------------------------------------------
sub displaybuttons(iTopBottom, iSectionType)

  if iTopBottom <> "" then
     lcl_topbottom = UCASE(iTopBottom)
  else
     lcl_topbottom = "TOP"
  end if

  if lcl_topbottom = "BOTTOM" then
     lcl_style_div = "margin-top:5px;"
  else
     lcl_style_div = "margin-bottom:5px;"
  end if

  if iSectionType <> "" then
     lcl_sectiontype = UCASE(iSectionType)
  else
     lcl_sectiontype = ""
  end if

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  response.write "  <input type=""button"" name=""cancelButton"" id=""cancelButton"" class=""button"" value=""Close Window"" onclick=""parent.close();"" />" & vbcrlf

  if lcl_sectiontype <> "HOURS" then
     response.write "  <input type=""button"" name=""saveChangesButton"" id=""saveChangesButton"" class=""button"" value=""Save Changes"" onclick=""validateFields();"" />" & vbcrlf
  end if

  response.write "</div>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub buildHoursFields(iStartEnd, iOptionType, iRowCount, iFieldValue)

'  response.write "["& iFieldValue & "]" & vbcrlf
  response.write "<select name=""" & iStartEnd & "_" & iOptionType & iRowCount & """ id=""" & iStartEnd & "_" & iOptionType & iRowCount & """>" & vbcrlf
  response.write "  <option value=""""></option>" & vbcrlf

  if iOptionType = "AMPM" then

     sFieldValue = cstr(iFieldValue)

     lcl_selected_am = ""
     lcl_selected_pm = ""

     if sFieldValue = "PM" then
        lcl_selected_pm = " selected=""selected"""
     elseif sFieldValue = "AM" then
        lcl_selected_am = " selected=""selected"""
     end if

     response.write "  <option value=""AM""" & lcl_selected_am & ">AM</option>" & vbcrlf
     response.write "  <option value=""PM""" & lcl_selected_pm & ">PM</option>" & vbcrlf
  else
     if iOptionType = "MINUTES" then
        for i = 0 to 55 step 5
           sFieldValue = trim(cstr(iFieldValue))

           if i < 10 then
              lcl_option_value = "0" & i
           else
              lcl_option_value = i
           end if

           if sFieldValue = trim(lcl_option_value) then
              lcl_selected_minutes = " selected=""selected"""
           else
              lcl_selected_minutes = ""
           end if

           response.write "  <option value=""" & lcl_option_value & """" & lcl_selected_minutes & ">" & lcl_option_value & "</option>" & vbcrlf
        next
     else
        sFieldValue = trim(iFieldValue)

        for i = 1 to 12
           if sFieldValue = trim(i) then
              lcl_selected_hours = " selected=""selected"""
           else
              lcl_selected_hours = ""
           end if

           response.write "  <option value=""" & i & """" & lcl_selected_hours & ">" & i & "</option>" & vbcrlf
        next
     end if
  end if

  response.write "</select>" & vbcrlf

  if iOptionType = "HOURS" then
     response.write "&nbsp;:&nbsp;" & vbcrlf
  elseif iOptionType = "MINUTES" then
     response.write "&nbsp;"
  end if

end sub

'------------------------------------------------------------------------------
function getTotalTextAreas(iDMTypeID, iDMID, iSectionID)

  lcl_return = 0

  sSQL = "SELECT count(dmtf.dm_fieldid) as total_elements "
  sSQL = sSQL & " FROM egov_dm_types_fields dmtf "
  sSQL = sSQL &      " INNER JOIN egov_dm_sections_fields dmsf "
  sSQL = sSQL &                 " ON dmtf.section_fieldid = dmsf.section_fieldid "
  sSQL = sSQL &                 " AND dmsf.isActive = 1 "
  sSQL = sSQL &                 " AND dmsf.isMultiLine = 1 "
  sSQL = sSQL &      " INNER JOIN egov_dm_types_sections dmts "
  sSQL = sSQL &                 " ON dmtf.dm_sectionid = dmts.dm_sectionid "
  sSQL = sSQL &                 " AND dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL &                 " AND dmts.sectionid = " & iSectionID

  if iDMID <> "" then
     sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_values dmv "
     sSQL = sSQL &                      " ON dmtf.dm_fieldid = dmv.dm_fieldid "
     sSQL = sSQL &                      " AND dmv.dm_typeid = " & iDMTypeID
     sSQL = sSQL &                      " AND dmv.dmid = " & iDMID
  end if

  sSQL = sSQL & " WHERE dmtf.dm_sectionid IN (select distinct dmts.dm_sectionid "
  sSQL = sSQL &                             " from egov_dm_types_sections dmts "
  sSQL = sSQL &                             " where dmts.dm_typeid = " & iDMTypeID
  sSQL = sSQL &                             " and dmts.sectionid = " & iSectionID & ") "

  set oGetTotalTextAreas = Server.CreateObject("ADODB.Recordset")
  oGetTotalTextAreas.Open sSQL, Application("DSN"), 3, 1

  if not oGetTotalTextAreas.eof then
     lcl_return = oGetTotalTextAreas("total_elements")
  end if

  oGetTotalTextAreas.close
  set oGetTotalTextAreas = nothing

  getTotalTextAreas = lcl_return

end function

'------------------------------------------------------------------------------
function getDMTypeDisplayMap(iDMTypeID)
  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT displayMap "
  sSQL = sSQL & " FROM egov_dm_types "
  sSQL = sSQL & " WHERE dm_typeid = " & iDMTypeID

  set oGetDMTypeDisplayMap = Server.CreateObject("ADODB.Recordset")
  oGetDMTypeDisplayMap.Open sSQL, Application("DSN"), 3, 1

  if not oGetDMTypeDisplayMap.eof then
     lcl_return = oGetDMTypeDisplayMap("displayMap")
  end if

  oGetDMTypeDisplayMap.close
  set oGetDMTypeDisplayMap = nothing

  getDMTypeDisplayMap = lcl_return

end function
%>