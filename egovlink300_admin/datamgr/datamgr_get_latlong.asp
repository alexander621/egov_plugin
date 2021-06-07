<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_get_latlong.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This screen pulls all non-valid addresses and attempts to pull a latitude and longitude value for each address from the GoogleMaps API.
'
' MODIFICATION HISTORY
' 1.0  01/16/2012 David Boyer - Initial Version
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

'Set up Google geocoder
 lcl_onload = lcl_onload & "initialize();"

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

'Get Org City and State
 lcl_org_city  = getOrgCityState(session("orgid"), "orgcity")
 lcl_org_state = getOrgCityState(session("orgid"), "orgstate")

'Get search parameters
 lcl_sc_dm_importid = 0

 if request.ServerVariables("REQUEST_METHOD") = "POST" then
    if request("sc_dm_importid") <> "" then
       if isnumeric(request("sc_dm_importid")) then
          lcl_sc_dm_importid = request("sc_dm_importid")
          lcl_sc_dm_importid = clng(lcl_sc_dm_importid)
       end if
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

  .fieldset legend {
     color: #800000;
  }

  .noWrap {
     white-space: nowrap;
  }

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
</style>

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

  <script type="text/javascript" src="https://maps.google.com/maps/api/js?sensor=false"></script>

<script language="javascript">
<!--

var geocoder;
var map;

function initialize() {
  geocoder = new google.maps.Geocoder();

  var myLatlng = new google.maps.LatLng(39.16563, -84.541111);
  var myOptions = {
        mapTypeId: google.maps.MapTypeId.ROADMAP,  //maptypes: ROADMAP, SATELLITE, HYBRID, TERRAIN
        zoom:      13,
        center:    myLatlng
      }
 
      //map = new google.maps.Map(document.getElementById("map_canvas"), myOptions);
}

$(document).ready(function(){
  var lcl_importLineNumber;

  //Retrieve latitude/longitude
  //$('#cancelButton').prop('disabled',true);
  //$('#getLatLongValuesButton').prop('disabled',true);
  //$('#nonValidAddressesTable').css('display','none');

  //BEGIN: Set up "help" definitions ------------------------------------------
  $('#helpDMImportID_text').hide();
  $('#helpCityState_text').hide();
  //$('#helpIncludeOnlyImportedNVA_text').hide();

  $('#helpDMImportID').click(function() {
     $('#helpDMImportID_text').toggle('slow');
  });

  $('#helpCity').click(function() {
     $('#helpCityState_text').toggle('slow');
  });

  $('#helpState').click(function() {
     $('#helpCityState_text').toggle('slow');
  });
  //$('#helpIncludeOnlyImportedNVA').click(function() {
  //   $('#helpIncludeOnlyImportedNVA_text').toggle('slow');
  //});
  //END: Set up "help" definitions --------------------------------------------

  //BEGIN: Cancel Button ------------------------------------------------------
  //$('#cancelButton').click(function() {

  //   if($('#cancelButton').val() == 'Clear Results') {
  //      $('#cancelButton').val('Cancel');
  //   }

  //   $('#sc_dm_importid').prop('disabled',false);
     //$('#sc_includeOnlyImportedNVA').prop('disabled',false);
  //   $('#sc_city').prop('disabled',false);
  //   $('#sc_state').prop('disabled',false);
  //   $('#sc_latitude').prop('disabled',false);
  //   $('#sc_longitude').prop('disabled',false);
  //   $('#getLatLongValuesButton').prop('disabled',true);
  //   $('.nonValidAddressRow').remove();
  //   $('.nonValidAddressTotalRow').remove();
     //$('#nonValidAddressesTable').css('display','none');
  //   $('#getNonValidAddressesButton').prop('disabled',false);
  //   $('#cancelButton').prop('disabled',true);
  //});
  //END: Cancel Button --------------------------------------------------------

  //BEGIN: Get Non-Valid Addresses Button -------------------------------------
//  $('#getNonValidAddressesButton').click(function() {
//     var lcl_orgid;
//     var lcl_dm_importid;
//     var lcl_sc_dm_importid         = '';
//     var lcl_includeOnlyImportedNVA = '';
//     var lcl_row_html               = '';
//     var lcl_bgcolor                = '#ffffff';
//     var lcl_totaladdresses         = 0;
//     var i                          = 0;

//     $('#cancelButton').prop('disabled',false);
//     $('#getNonValidAddressesButton').prop('disabled',true);
//     $('#sc_dm_importid').prop('disabled',true);
//     $('#sc_includeOnlyImportedNVA').prop('disabled',true);
//     $('#sc_city').prop('disabled',true);
//     $('#sc_state').prop('disabled',true);
//     $('#sc_latitude').prop('disabled',true);
//     $('#sc_longitude').prop('disabled',true);

//     lcl_orgid          = $('#orgid').val();
//     lcl_dm_importid    = $('#dm_importid').val();
//     lcl_sc_dm_importid = $('#sc_dm_importid').val();

//     if(document.getElementById('sc_includeOnlyImportedNVA').checked) {
//        lcl_includeOnlyImportedNVA = $('#sc_includeOnlyImportedNVA').val();
//     }

//     $.post('datamgr_import_from_spreadsheet_action.asp', {
//        userid:                 '<%=session("userid")%>',
//        orgid:                  lcl_orgid,
//        includeOnlyImportedNVA: lcl_includeOnlyImportedNVA,
//        sc_dm_importid:         lcl_sc_dm_importid,
//        action:                 'GET_NONVALID_ADDRESSES',
//        isAjax:                 'Y'
//     }, function(result) {
//        lcl_row_html = result;

//        $('#nonValidAddressesTable').show(function() {
//           $('#nonValidAddressesTable').append(lcl_row_html);
//           $('#getLatLongValuesButton').prop('disabled',false);
//        });
//     });

//     $('#getLatLongValuesButton').prop('disabled',false);

//  });
  //END: Get Non-Valid Addresses Button ---------------------------------------

  //BEGIN: Get Lat/Long Button ------------------------------------------------
  $('#getLatLongValuesButton').click(function(){
    var lcl_total_nvaddresses = Number($('#nvaddresses_total').val());
    var lcl_false_count       = Number(0);

    if(lcl_total_nvaddresses > 0) {
       var lcl_starting_linenum = 0;
       var lcl_sc_latitude      = "";
       var lcl_sc_longitude     = "";

       lcl_sc_latitude  = $('#sc_latitude').val();
       lcl_sc_longitude = $('#sc_longitude').val();

       if(lcl_sc_longitude == '') {
          document.getElementById("sc_longitude").focus();
          inlineMsg(document.getElementById("sc_longitude").id,'<strong>Required Field Missing: </strong> Latitude (where the latitude value is to be stored.)',10,'sc_longitude');
          lcl_false_count = lcl_false_count + 1;
       }

       if(lcl_sc_latitude == '') {
          document.getElementById("sc_latitude").focus();
          inlineMsg(document.getElementById("sc_latitude").id,'<strong>Required Field Missing: </strong> Latitude (where the latitude value is to be stored.)',10,'sc_latitude');
          lcl_false_count = lcl_false_count + 1;
       }

    } else {
       document.getElementById("getLatLongValuesButton").focus();
       inlineMsg(document.getElementById("getLatLongValuesButton").id,'<strong>Invalid: </strong> No non-valid addresses are available.',10,'getLatLongValuesButton');
       lcl_false_count = lcl_false_count + 1;
    }

    if(lcl_false_count > 0) {
       return false
    } else {
       lcl_starting_linenum = $('#nvaddresses_linenum').val();

       getLatLng(lcl_starting_linenum);
    }

  });
  //END: Get Lat/Long Button --------------------------------------------------

});

function getLatLng() {
  var lcl_orgid;
  var lcl_dmid;
  var lcl_dm_importid;
  var lcl_sc_latitude        = '';
  var lcl_sc_longitude       = '';
  var lcl_address            = '';
  var lcl_city               = '';
  var lcl_state              = '';
  var lcl_default_city       = '';
  var lcl_default_state      = '';
  var lcl_nvaddress_address  = '';
  var lcl_nvaddress_city     = '';
  var lcl_nvaddress_state    = '';
  var lcl_latitude           = '';
  var lcl_longitude          = '';
  var lcl_total_lines        = 10;
  var lcl_total_nv_addresses = 0;
  var lcl_linenum;

  lcl_orgid              = '<%=lcl_orgid%>';
  lcl_linenum            = Number($('#nvaddresses_linenum').val());
  lcl_total_nv_addresses = Number($('#nvaddresses_total').val());
  lcl_default_city       = $('#sc_city').val();
  lcl_default_state      = $('#sc_state').val();
  lcl_sc_latitude        = $('#sc_latitude').val();
  lcl_sc_longitude       = $('#sc_longitude').val();
  lcl_total_lines        = Number(lcl_linenum + 10);

  if((lcl_linenum+1) > lcl_total_nv_addresses) {
     displayScreenMsg('Saving latitude/longitude values...');

     for(i = 1; i <= lcl_total_nv_addresses; i++) {
        lcl_dmid        = $('#nvaddresses_dmid'        + i).val();
        lcl_dm_importid = $('#nvaddresses_dm_importid' + i).val();
        lcl_latitude    = $('#nvaddresses_latitude'    + i).val();
        lcl_longitude   = $('#nvaddresses_longitude'   + i).val();

        $.post('datamgr_import_from_spreadsheet_action.asp', {
           orgid:                 lcl_orgid,
           dmid:                  lcl_dmid,
           linecount:             i,
           dm_importid:           lcl_dm_importid,
           sc_latitude:           lcl_sc_latitude,
           sc_longitude:          lcl_sc_longitude,
           nvaddresses_latitude:  lcl_latitude,
           nvaddresses_longitude: lcl_longitude,
           action:                'UPDATE_LAT_LONG',
           isAjax:                'Y'
        }, function(result) {
           var lcl_linecount = Number(result);

           $('#displayStatus' + result).html('<span class="redText">Saved</span>');

           if(lcl_linecount == lcl_total_nv_addresses) {
              //$('#cancelButton').val('Clear Results');
              //$('#cancelButton').prop('disabled',false);
              $('#returnButton').prop('disabled',false);
           }
        });
     }

  } else {    
     if(lcl_total_lines > lcl_total_nv_addresses) {
        lcl_total_lines = lcl_total_nv_addresses;
     }

     //So we aren't starting from zero (0) and/or repeating a line already completed, add "1".
     lcl_linenum = Number(lcl_linenum + 1);

     displayScreenMsg('Processing lines ' + lcl_linenum + ' - ' + lcl_total_lines + '...');

     if(lcl_total_lines > 0) {
        for(i = lcl_linenum; i <= lcl_total_lines; i++) {
           lcl_address         = '';
           lcl_nvaddress_city  = '';
           lcl_nvaddress_state = '';

           if($('#nvaddresses_address' + i)) {
              lcl_address = $('#nvaddresses_address' + i).val();
           }

           if($('#nvaddresses_city' + i)) {
              lcl_nvaddress_city = $('#nvaddresses_city' + i).val();
           }

           if($('#nvaddresses_state' + i)) {
              lcl_nvaddress_state = $('#nvaddresses_state' + i).val();
           }

           if(lcl_nvaddress_city != '') {
              lcl_city = lcl_nvaddress_city;
           } else {
              lcl_city = lcl_default_city;
           }

           if(lcl_nvaddress_state != '') {
              lcl_state = lcl_nvaddress_state;
           } else {
              lcl_state = lcl_default_state;
           }

           $('#nvaddresses_city'  + i).val(lcl_city);
           $('#nvaddresses_state' + i).val(lcl_state);

           if(lcl_city != '') {
              if(lcl_address != '') {
                 lcl_address = lcl_address + ', ' + lcl_city;
              } else {
                 lcl_address = lcl_city;
              }
           }

           if(lcl_state != '') {
              if(lcl_address != '') {
                 lcl_address = lcl_address + ', ' + lcl_state;
              } else {
                 lcl_address = lcl_state;
              }
           }

           setupLatLong(lcl_address, lcl_total_nv_addresses, lcl_total_lines, i);
        }
     }
  }
}

function setupLatLong(p_address, p_total_addresses, p_total_lines, p_linenum) {
  geocoder.geocode( { 'address': p_address}, function(results, status) {
    var lcl_latlng;
    var lcl_latitude;
    var lcl_longitude;
    var lcl_latlng_index;

    if (status == google.maps.GeocoderStatus.OK) {
        lcl_latlng = results[0].geometry.location;
        lcl_latlng = lcl_latlng.toString();
        lcl_latlng = lcl_latlng.replace('(', '');
        lcl_latlng = lcl_latlng.replace(')', '');

        lcl_latlng_index = lcl_latlng.indexOf(',');
        lcl_latitude     = lcl_latlng.substr(0,lcl_latlng_index);
        lcl_longitude    = lcl_latlng.substr(lcl_latlng_index+2);
    }

    if(document.getElementById('nvaddresses_latitude'  + p_linenum)) {
       document.getElementById('nvaddresses_latitude'  + p_linenum).value = lcl_latitude;
    }

    if(document.getElementById('nvaddresses_longitude'  + p_linenum)) {
       document.getElementById('nvaddresses_longitude' + p_linenum).value = lcl_longitude;
    }

    $('#nvaddresses_latitude'  + p_linenum).prop('disabled',true);
    $('#nvaddresses_longitude' + p_linenum).prop('disabled',true);
    $('#nvaddresses_city'      + p_linenum).prop('disabled',true);
    $('#nvaddresses_state'     + p_linenum).prop('disabled',true);

    document.getElementById('nvaddresses_linenum').value = p_linenum;

    if(p_linenum == p_total_lines) {
       if(p_total_lines <= p_total_addresses) {
          window.setTimeout("getLatLng()", (10 * 1000));
       }
    }
  });
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

//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
 response.write "<form name=""retrieveLatLong"" id=""retrieveLatLong"" method=""post"" action=""datamgr_get_latlong.asp"">" & vbcrlf
 response.write "  <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
 response.write "  <input type=""hidden"" name=""dm_importid"" id=""dm_importid"" value="""" size=""5"" maxlength=""10"" />" & vbcrlf
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
 'response.write "                <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
 'response.write "                  <tr>" & vbcrlf
 'response.write "                      <td>" & vbcrlf
 response.write "                          <input type=""button"" name=""returnButton"" id=""returnButton"" value=""Return to List"" class=""button"" onclick=""location.href='datamgr_list.asp" & lcl_url_parameters & "'"" />" & vbcrlf
 'response.write "                      </td>" & vbcrlf
 'response.write "                      <td align=""right"">" & vbcrlf
 'response.write "                          <input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" />" & vbcrlf
 'response.write "                          <input type=""button"" name=""getNonValidAddressesButton"" id=""getNonValidAddressesButton"" value=""Get Non-Valid Addresses"" class=""button"" />" & vbcrlf
 'response.write "                          <input type=""button"" name=""getLatLongValuesButton"" id=""getLatLongValuesButton"" value=""Get Lat/Long Values"" class=""button"" />" & vbcrlf
 'response.write "                      </td>" & vbcrlf
 'response.write "                  </tr>" & vbcrlf
 'response.write "                </table>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td>&nbsp;</td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td>" & vbcrlf

'BEGIN: Import: Non-Valid Addresses -------------------------------------------
 response.write "    <p>" & vbcrlf
 response.write "    <fieldset name=""nonValidAddresses"" id=""nonValidAddresses"" class=""fieldset"">" & vbcrlf
 response.write "      <legend>Get Latitude and Longitude</legend>" & vbcrlf
 response.write "      <div class=""instructions"">" & vbcrlf
 response.write "        Check for all ""non-valid"" street addresses for imported data ONLY.  Return any/all that exist in a list with input boxes for latitude and longitude values.<br />" & vbcrlf
 response.write "      </div>"
 response.write "      <p>" & vbcrlf
 response.write "      <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""noWrap"">DM ImportID:</td>" & vbcrlf
 response.write "            <td class=""noWrap"">" & vbcrlf
 response.write "                <input type=""text"" name=""sc_dm_importid"" id=""sc_dm_importid"" value="""" size=""5"" maxlength=""100"" />" & vbcrlf
 response.write "                <img src=""../images/help.jpg"" name=""helpDMImportID"" id=""helpDMImportID"" align=""top"" class=""helpOption"" alt=""Click for more info"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "            <td width=""100%"" align=""center"" class=""noWrap"">" & vbcrlf
 response.write "                &nbsp;" & vbcrlf
 'response.write "                <input type=""checkbox"" name=""sc_includeOnlyImportedNVA"" id=""sc_includeOnlyImportedNVA"" value=""Y"" checked=""checked"" /> Include ONLY imported, Non-Valid Addresses" & vbcrlf
 'response.write "                <img src=""../images/help.jpg"" name=""helpIncludeOnlyImportedNVA"" id=""helpIncludeOnlyImportedNVA"" class=""helpOption"" alt=""Click for more info"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr>" & vbcrlf
 response.write "            <td colspan=""3"">" & vbcrlf
 response.write "                <div name=""helpDMImportID_text"" id=""helpDMImportID_text"" class=""helpOptionText"">" & vbcrlf
 response.write "                  <p><strong>E-GOV TIP:</strong><br />This option is used to limit the non-valid addresses returned in the list to a specific import.</p>" & vbcrlf
 response.write "                </div>" & vbcrlf
 'response.write "                <div name=""helpIncludeOnlyImportedNVA_text"" id=""helpIncludeOnlyImportedNVA_text"" class=""helpOptionText"">" & vbcrlf
 'response.write "                  <p><strong>E-GOV TIP:</strong><br />This option is used to limit the non-valid addresses returned in the list to only those that have been imported or all non-valid addresses.</p>" & vbcrlf
 'response.write "                </div>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""noWrap"">City: </td>" & vbcrlf
 response.write "            <td colspan=""2"">" & vbcrlf
 response.write "                <input type=""text"" name=""sc_city"" id=""sc_city"" value=""" & lcl_org_city & """ size=""20"" maxlength=""100"" />" & vbcrlf
 response.write "                <img src=""../images/help.jpg"" name=""helpCity"" id=""helpCity"" align=""top"" class=""helpOption"" alt=""Click for more info"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""noWrap"">State: </td>" & vbcrlf
 response.write "            <td colspan=""2"">" & vbcrlf
 response.write "                <input type=""text"" name=""sc_state"" id=""sc_state"" value=""" & lcl_org_state & """ size=""20"" maxlength=""100"" />" & vbcrlf
 response.write "                <img src=""../images/help.jpg"" name=""helpState"" id=""helpState"" align=""top"" class=""helpOption"" alt=""Click for more info"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr>" & vbcrlf
 response.write "            <td colspan=""3"">" & vbcrlf
 response.write "                <div name=""helpCityState_text"" id=""helpCityState_text"" class=""helpOptionText"">" & vbcrlf
 response.write "                  <p><strong>E-GOV TIP:</strong><br />These options are the city/state assigned to the Organization via the ""Org Properties"" screen. " & vbcrlf
 response.write "                     When pulling the results list, the city and state are pulled from the egov_dm_data record.  If the city and/or state or NULL then you can either " & vbcrlf
 response.write "                     enter a specific city/state for a specific record or have these organization city/state values be automatically used when retrieving the latitude/longitude values.</p>" & vbcrlf
 response.write "                </div>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""noWrap"">Latitude: </td>" & vbcrlf
 response.write "            <td colspan=""2"">" & vbcrlf
 response.write "                <select name=""sc_latitude"" id=""sc_latitude"" class=""transferFieldData"" onchange=""clearMsg('sc_latitude');"">" & vbcrlf
 response.write "                  <option value="""">&nbsp;</option>" & vbcrlf
                                   displayTransferFieldOptions lcl_orgid
 response.write "                </select>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr valign=""top"">" & vbcrlf
 response.write "            <td class=""noWrap"">Longitude: </td>" & vbcrlf
 response.write "            <td colspan=""2"">" & vbcrlf
 response.write "                <select name=""sc_longitude"" id=""sc_longitude"" class=""transferFieldData"" onchange=""clearMsg('sc_longitude');"">" & vbcrlf
 response.write "                  <option value="""">&nbsp;</option>" & vbcrlf
                                   displayTransferFieldOptions lcl_orgid
 response.write "                </select>" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "        <tr>" & vbcrlf
 response.write "            <td colspan=""3"">" & vbcrlf
 'response.write "                <input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" />" & vbcrlf
 response.write "                <input type=""submit"" name=""getNonValidAddressesButton"" id=""getNonValidAddressesButton"" value=""Get Non-Valid Addresses"" class=""button"" />" & vbcrlf
 response.write "                <input type=""button"" name=""getLatLongValuesButton"" id=""getLatLongValuesButton"" value=""Get Lat/Long Values"" class=""button"" />" & vbcrlf
 response.write "            </td>" & vbcrlf
 response.write "        </tr>" & vbcrlf
 response.write "      </table>" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 response.write "      <table id=""nonValidAddressesTable"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"">" & vbcrlf
 response.write "        <tr align=""left"" id=""columnHeaderRow"">" & vbcrlf
 response.write "            <th class=""columnHeader"">Address</th>" & vbcrlf
 response.write "            <th>City</th>" & vbcrlf
 response.write "            <th>State</th>" & vbcrlf
 response.write "            <th>Latitude</th>" & vbcrlf
 response.write "            <th>Longitude</th>" & vbcrlf
 response.write "            <th align=""center"">DM ImportID</th>" & vbcrlf
 response.write "            <th>&nbsp;</th>" & vbcrlf
 response.write "        </tr>" & vbcrlf

 if request.ServerVariables("REQUEST_METHOD") = "POST" then
    displayNonValidAddressRows lcl_sc_dm_importid
 end if

 response.write "      </table>" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "      <p>" & vbcrlf
 'response.write "        <input type=""button"" name=""beginImportButton"" id=""beginImportButton"" value=""Begin Import"" class=""button"" onclick=""beginImport()"" />" & vbcrlf
 response.write "      </p>" & vbcrlf
 response.write "    </fieldset>" & vbcrlf
 response.write "    </p>" & vbcrlf
'END: Import: Non-Valid Addresses ---------------------------------------------

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
sub displayTransferFieldOptions(iOrgID)
  sOrgID = 0

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  sSQL = sSQL & "SELECT DISTINCT "
  sSQL = sSQL & " dmtf.dm_typeid, "
  sSQL = sSQL & " dmt.description, "
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
  sSQL = sSQL & " ORDER BY dmt.description, dms.sectionname, dmsf.fieldname "

 	set oDMTransferFieldsOptions = Server.CreateObject("ADODB.Recordset")
	 oDMTransferFieldsOptions.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMTransferFieldsOptions.eof then
     lcl_bgcolor            = "#ffffff"

     do while not oDMTransferFieldsOptions.eof

        response.write "  <option value=""dmtypeid" & oDMTransferFieldsOptions("dm_typeid") & "_dmsectionid" & oDMTransferFieldsOptions("dm_sectionid") & "_dmfieldid" & oDMTransferFieldsOptions("dm_fieldid") & """>[" & oDMTransferFieldsOptions("description") & "] " & oDMTransferFieldsOptions("sectionname") & ": " & oDMTransferFieldsOptions("fieldname") & "</option>" & vbcrlf

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

'------------------------------------------------------------------------------
function getOrgCityState(iOrgID, iOrgColumn)
  dim lcl_return, lcl_orgid, lcl_org_column

  lcl_return     = ""
  lcl_orgid      = 0
  lcl_org_column = ""

  if iOrgID <> "" then
     lcl_orgid = clng(iOrgID)
  end if

  if iOrgColumn <> "" then
     if not containsApostrophe(iOrgColumn) then
        lcl_org_column = iOrgColumn
     end if
  end if

'  if lcl_org_column <> "" then
     sSQL = "SELECT " & lcl_org_column & " AS orgColumn "
     sSQL = sSQL & " FROM organizations "
     sSQL = sSQL & " WHERE orgid = " & lcl_orgid

    	set oGetOrgCityState = Server.CreateObject("ADODB.Recordset")
    	oGetOrgCityState.Open sSQL, Application("DSN"), 3, 1
'dtb_debug(sSQL)
     if not oGetOrgCityState.eof then
        lcl_return = oGetOrgCityState("orgColumn")
     end if

     oGetOrgCityState.close
     set oGetOrgCityState = nothing

'  end if

  getOrgCityState = lcl_return

end function

'------------------------------------------------------------------------------
sub displayNonValidAddressRows(iSC_DMImportID)

  dim lcl_linecount, lcl_bgcolor, sSC_DMImportID

  lcl_linecount  = 0
  lcl_bgcolor    = ""
  sSC_DMImportID = 0

  if iSC_DMImportID <> "" then
     if isnumeric(iSC_DMImportID) then
        sSC_DMImportID = clng(iSC_DMImportID)
     end if
  end if

  sSQLnv = "SELECT "
  sSQLnv = sSQLnv & " dmid, "
  sSQLnv = sSQLnv & " dm_typeid, "
  sSQLnv = sSQLnv & " streetaddress, "
  sSQLnv = sSQLnv & " sortstreetname, "
  sSQLnv = sSQLnv & " city, "
  sSQLnv = sSQLnv & " state, "
  sSQLnv = sSQLnv & " latitude, "
  sSQLnv = sSQLnv & " longitude, "
  sSQLnv = sSQLnv & " dm_importid "
  sSQLnv = sSQLnv & " FROM egov_dm_data "
  sSQLnv = sSQLnv & " WHERE validstreet = 'N' "
  sSQLnv = sSQLnv & " AND streetaddress <> '' "
  sSQLnv = sSQLnv & " AND streetaddress IS NOT NULL "
  sSQLnv = sSQLnv & " AND (latitude = 0 OR latitude is null OR longitude = 0 OR longitude is null) "

  if sSC_DMImportID > 0 then
     sSQLnv = sSQLnv & " AND dm_importid = " & sSC_DMImportID
  else
     sSQLnv = sSQLnv & " AND dm_importid <> '' "
     sSQLnv = sSQLnv & " AND dm_importid IS NOT NULL "
  end if

  sSQLnv = sSQLnv & " ORDER BY sortstreetname, streetaddress "

  set oGetNonValidAddresses = Server.CreateObject("ADODB.Recordset")
  oGetNonValidAddresses.Open sSQLnv, Application("DSN"), 3, 1

  if not oGetNonValidAddresses.eof then
     do while not oGetNonValidAddresses.eof
        lcl_linecount  = lcl_linecount + 1
        lcl_bgcolor    = changeBGColor(lcl_bgcolor,"#ffffff","#eeeeee")
        lcl_snumber    = ""
        lcl_sprefix    = ""
        lcl_ssuffix    = ""
        lcl_sdirection = ""

        lcl_saddress   = oGetNonValidAddresses("streetaddress")
        lcl_saddress   = replace(lcl_saddress,chr(10),"")
        lcl_saddress   = replace(lcl_saddress,chr(13),"")

        lcl_dm_address = buildStreetAddress(lcl_snumber, lcl_sprefix, lcl_saddress, lcl_ssuffix, lcl_sdirection)

        response.write "  <tr class=""nonValidAddressRow"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "       <td style=""white-space: nowrap"">" & vbcrlf
        response.write "           <input type=""hidden"" name=""nvaddresses_dmid"    & lcl_linecount & """ id=""nvaddresses_dmid"    & lcl_linecount & """ value=""" & oGetNonValidAddresses("dmid") & """ size=""10"" maxlength=""100"" />" & vbcrlf
        response.write "           <input type=""hidden"" name=""nvaddresses_address" & lcl_linecount & """ id=""nvaddresses_address" & lcl_linecount & """ value=""" & lcl_dm_address                & """ size=""10"" maxlength=""100"" />" & vbcrlf
        response.write lcl_linecount & ". " & lcl_dm_address
        response.write "       </td>" & vbcrlf
        response.write "       <td>" & vbcrlf
        response.write "           <input type=""text"" name=""nvaddresses_city" & lcl_linecount & """ id=""nvaddresses_city" & lcl_linecount & """ value=""" & oGetNonValidAddresses("city") & """ size=""20"" maxlength=""50"" />" & vbcrlf
        response.write "       </td>" & vbcrlf
        response.write "       <td>" & vbcrlf
        response.write "           <input type=""text"" name=""nvaddresses_state" & lcl_linecount & """ id=""nvaddresses_state" & lcl_linecount & """ value=""" & oGetNonValidAddresses("state") & """ size=""3"" maxlength=""20"" />" & vbcrlf
        response.write "       </td>" & vbcrlf
        response.write "       <td>" & vbcrlf
        response.write "           <input type=""text"" name=""nvaddresses_latitude" & lcl_linecount & """ id=""nvaddresses_latitude" & lcl_linecount & """ value=""" & oGetNonValidAddresses("latitude") & """ size=""20"" maxlength=""100"" />" & vbcrlf
        response.write "       </td>" & vbcrlf
        response.write "       <td>" & vbcrlf
        response.write "           <input type=""text"" name=""nvaddresses_longitude" & lcl_linecount & """ id=""nvaddresses_longitude" & lcl_linecount & """ value=""" & oGetNonValidAddresses("longitude") & """ size=""20"" maxlength=""100"" />" & vbcrlf
        response.write "       </td>" & vbcrlf
        response.write "       <td align=""center"">" & vbcrlf
        response.write             oGetNonValidAddresses("dm_importid") & vbcrlf
        response.write "           <input type=""hidden"" name=""nvaddresses_dm_importid" & lcl_linecount & """ id=""nvaddresses_dm_importid" & lcl_linecount & """ value=""" & oGetNonValidAddresses("dm_importid") & """ size=""10"" maxlength=""100"" />" & vbcrlf
        response.write "       </td>" & vbcrlf
        response.write "       <td>" & vbcrlf
        response.write "           <span id=""displayStatus" & lcl_linecount & """></span>" & vbcrlf
        response.write "       </td>" & vbcrlf
        response.write "   </tr>" & vbcrlf

        oGetNonValidAddresses.movenext
     loop
  end if

  response.write "  <tr class=""nonValidAddressTotalRow""bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
  response.write "      <td align=""right"" colspan=""7"">" & vbcrlf
  response.write "          <input type=""hidden"" name=""nvaddresses_total"" id=""nvaddresses_total"" value=""" & lcl_linecount & """ size=""5"" maxlength=""100"" />" & vbcrlf
  response.write "          <input type=""hidden"" name=""nvaddresses_linenum"" id=""nvaddresses_linenum"" value=""0"" size=""5"" maxlength=""100"" />" & vbcrlf
  response.write "          <strong>Total Non-Valid Addresses: </strong>[" & lcl_linecount & "]" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  set oGetNonValidAddresses = nothing

end sub
%>
