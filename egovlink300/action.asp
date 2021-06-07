<!DOCTYPE HTML>
<%
intZoom = 4
dblLat = 40.042916434756286
dblLng = -95.9609852298176

%>
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_rye.asp" //-->
<%
'response.redirect("feature_offline.asp")
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: action_.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Action Line Search Results.
'
' MODIFICATION HISTORY
' ?.?	05/08/07	 Steve Loar - Changes to problem location to handle larger cities
' 6.2	07/16/07 	Steve Loar - Changes to handle email/request type problems.
' 6.3 11/16/07  David Boyer - Now validates contact email address to ensure it is entered in proper format.
' 6.4 01/22/08  David Boyer - Added "isFeatureOffline" check to screen.
' 6.5 05/08/08  David Boyer - Now disable the SUBMIT FORM button when clicked to submit request.
' 6.6 05/29/09  David Boyer - Added check to see if "Additional Information" textarea is displayed or not.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'To help prevent hacks.
 if NOT isnumeric(request("actionid")) then
    response.redirect "action.asp"
 end if

'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 dim sError, oActionOrg, iSectionID, sDocumentTitle, blnDisplayMobileOptions_takePic, blnShowMapInput

'Capture current path
 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
 session("RedirectLang") = "Return to Action Line"

 set oActionOrg = New classOrganization

'Show/Hide hidden fields.  To Hide = "HIDDEN", To Show = "TEXT"
 lcl_hidden = "hidden"

'Check for org features
 lcl_orghasfeature_issue_location     = OrgHasFeature(iorgid,"issue location")
 lcl_orghasfeature_large_address_list = OrgHasFeature(iorgid,"large address list")
 lcl_orghasfeature_actionline_formcreator_mobileoptions = OrgHasFeature(iorgid, "actionline_formcreator_mobileoptions")
 lcl_orghasfeature_hide_login         = OrgHasFeature(iorgid,"actionline_hide_login")

 'response.write 
 if not OrgHasFeature(iorgid,"action line") then response.redirect "default.asp"

'Check for org "edit displays"
 lcl_orghasdisplay_action_page_title     = OrgHasDisplay(iorgid,"action page title")
 lcl_orghasdisplay_action_tracking_title = OrgHasDisplay(iorgid,"action tracking title")
 lcl_orghasdisplay_action_list_title     = OrgHasDisplay(iorgid,"action list title")

'Check for cookies
 lcl_cookie_userid = request.cookies("userid")

'Build the Title
 lcl_title = sOrgName

 if iorgid <> 7 then
    lcl_title = "E-Gov Services " & lcl_title
 end if

'Build parameters
 lcl_sc_request_type = request("sc_request_type")
 if instr(lcl_sc_request_type,"<") > 0 or instr(lcl_sc_request_type,"<") > 0 then lcl_sc_request_type = ""

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'------------------------------------------------------------------------------
'Determine if this screen is being accessed by a mobile device
' *** Script provided by detectmobilebrowsers.com
'------------------------------------------------------------------------------
 dim u, b, v, lcl_isMobileDevice

 set u = Request.ServerVariables("HTTP_USER_AGENT")
 set b = new RegExp
 set v = new RegExp

 lcl_isMobileDevice = false

 b.Pattern    = "(android|bb\d+|meego).+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|mobile.+firefox|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|series(4|6)0|symbian|treo|up\.(browser|link)|vodafone|wap|windows (ce|phone)|xda|xiino"
 v.Pattern    = "1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|yas\-|your|zeto|zte\-"
 b.IgnoreCase = true
 v.IgnoreCase = true
 b.Global     = true
 v.Global     = true

 if b.test(u) or v.test(Left(u,4)) then
    'response.redirect("http://detectmobilebrowser.com/mobile")
    lcl_isMobileDevice = true
 end if
%>
<html>
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

		<title><%=lcl_title%></title>

<script src='https://www.google.com/recaptcha/api.js'></script>
 	<link type="text/css" rel="stylesheet" href="css/styles.css" />
 	<link type="text/css" rel="stylesheet" href="global.css" />
 	<link type="text/css" rel="stylesheet" href="css/style_<%=iorgid%>.css" />

<style type="text/css">
  .alertmsg {
     color:       #ff0000;
     font-weight: bold;
     font-style:  italic;
  }

  .sectionTitle {
     font-weight:     bold;
     text-decoration: underline;
  }

  .fieldset_warning {
     border:           2px solid #000000;
     border-radius:    5px;
     background-color: #ff0000;
     margin-bottom:    10px;
     margin-top:       10px;
     padding:          10px;
     /* width:            574px; */
  }

  .fieldset_warning,
  .fieldset_warning p {
     color:       #ffff00;
     /* font-size:   10px; */
     font-size: 0.875em;
     font-weight: bold;
  }

  .address_fieldset {
     border:        1pt solid #808080;
     border-radius: 6px;
  }

  #validaddresslist {
     border:             1pt solid #c0c0c0;
     border-radius: 6px;
     background-color:   #efefef;
     margin-top:         4px;
  }

  #validaddresslist legend {
     border:           1pt solid #c0c0c0;
     border-radius:    6px;
     background-color: #ffffff;
     color:            #ff0000;
     padding-left:     4px;
     padding-right:    4px;
  }

  div#addresspicklist {
     border-radius: 6px;
  }

  #screenMsg {
     text-align:    right;
     color:         #ff0000;
     font-weight:   bold;
     margin-bottom: 5px;
     width:         574px;
  }

  #contactInfoMsg {
     margin:      10px 0px;
     text-align:  center;
     color:       #ff0000;
     font-weight: bold;
     font-style:  italic;
  }

  .fieldset
  {
     margin: 10px 0px;
     background-color: #e0e0e0;
     border-radius: 6px;
  }

  .fieldset legend
  {
     background-color: #ffffff;
     padding: 4px 8px;
     border: 1pt solid #808080;
     border-radius: 6px;
     font-size: 1.125em;
     color: #800000;
  }

#rightSideDiv
{
   float: right;
}

  #formsDiv li
  {
     list-style: none;
     display: inline-table;
     width: 48%;
  }

  #formsDiv ul
  {
     margin: 0px;
     padding: 0px;
  }

  .nestedFieldset
  {
     margin: 10px 0px;
     background-color: #eeeeee;
     border-radius: 6px;
  }

  .nestedFieldset legend
  {
     font-size: 1em;
     color: #0000ff;
  }

/*--------------------------------------------------------------------------------
BEGIN: Set up for screens with max of 800px
----------------------------------------------------------------------------------*/
@media screen and (max-width: 768px) 
{
   .indent20
   {
      padding: 5px;
   }

   #centercontent
   {
      width: 100%;
      margin-left: 0px;
   }


   img
   {
      /*width: 100%; */
      /* min-height: 1px; */
   } 

.accountmenu
{
   width: 38px;
   height: 25px;
}

   .fieldset
   {
     /*margin-right: 10px;*/
     padding:0;
   }

    #formsDiv li
   {
      float: left;
      display: block;
      width: 100%;
   }
}


.respIndent
{
	 margin-top:20px; 
	 margin-left:20px;
}

 
.frmBottom
{
 max-width:450px;
 text-align:right;
}

@media screen and (max-width: 480px) 
{
	.tblResponsive 
	{
		width:95%;
	}
	.tblResponsive td:first-child
	{
		text-align:left;
	}
	.tblResponsive td
	{
		display:block;
	}

	.tblResponsive textarea.formtextarea
	{
		width:100%;
		max-width:450px;
	}
	.inputResponsive
	{
		width:100%;
	}

	.address_fieldset
	{
		display:inline-block;
	}
	#screenMsg
	{
		width:100%;
	}
	.respIndent
	{
	 	margin-top:0; 
	 	margin-left:0;
	}
	.frmBottom
	{
 		text-align:center;
	}
	.respHide
	{
		display:none;
	}
	.respHeader
	{
		max-height:145px;
		height:auto;
	}

	.fieldset legend,
	#formsDiv li
	{
		padding:0;
	}
	input[type="text"].phonenum
	{
		width:auto !important;
	}

	#map_wrapper
	{
		left: 1% !important;
	}
	<% if iorgid = "37" then %>
	<% end if %>
}

</style>

 	<script type="text/javascript" src="scripts/modules.js"></script>
 	<script type="text/javascript" src="scripts/easyform.js"></script>
  <script type="text/javascript" src="scripts/ajaxLib.js"></script>
  <script type="text/javascript" src="scripts/removespaces.js"></script>
  <script type="text/javascript" src="scripts/setfocus.js"></script>
  <script type="text/javascript" src="scripts/textareamaxlength.js"></script>
  <script type="text/javascript" src="scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="scripts/jquery-1.9.1.min.js"></script>
  <script type='text/javascript' src='scripts/jquery.ui.core.min.js?ver=1.10.3'></script>
  <script type='text/javascript' src='scripts/jquery.ui.position.min.js?ver=1.10.3'></script>
<script type='text/javascript' src='scripts/jquery.ui.widget.min.js?ver=1.10.3'></script>
<script type='text/javascript' src='scripts/jquery.ui.menu.min.js?ver=1.10.3'></script>
<script type='text/javascript' src='scripts/jquery.ui.autocomplete.min.js?ver=1.10.3'></script>

  <script type="text/javascript" src="scripts/zebra_datepicker.js"></script>
  <link rel="stylesheet" href="css/zebra_datepicker.css" type="text/css">
  <link rel='stylesheet' id='smoothness-css'  href='https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css?ver=3.6.1' type='text/css' media='all' />

<script type="text/javascript">

function initMaps()
{
	if (initMap && typeof initMap === 'function') { initMap();}
	if (initalizeOtherReports && typeof initalizeOtherReports === 'function') { initalizeOtherReports();}
	
}

$(document).ready(function(){
  $('#searchButton_trackingLookup').click(function() {
     $('#frmActionLookup').submit();
  });

  $('#searchButton_requestTypes').click(function() {
     searchRequestTypes();
  });
<%
 'BEGIN: Check for the "issue location" feature ------------------------------
  lcl_addressfield_exists = true

  if lcl_orghasfeature_issue_location AND lcl_addressfield_exists then
     lcl_addresstype = ""

//     response.write "$(document).ready(function(){" & vbcrlf

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

//     response.write "});" & vbcrlf
  end if
%>
});
<%
  if lcl_orghasfeature_issue_location AND lcl_addressfield_exists then

    'BEGIN: Check Address -------------------------------------------------
     response.write "function checkAddress(iFunction, iValidate) {" & vbcrlf
     response.write "  var lcl_streetnumber = $('#residentstreetnumber').val();" & vbcrlf
     response.write "  var lcl_streetname   = $('#streetaddress').val();" & vbcrlf
     response.write "  var lcl_otheraddress = $('#ques_issue2').val();" & vbcrlf
     response.write "  var lcl_isFinalCheck = 'N';" & vbcrlf
     response.write "  clearScreenMsg();" & vbcrlf

     response.write "  if(iFunction == 'FinalCheck') {" & vbcrlf
     response.write "     lcl_isFinalCheck = 'Y';" & vbcrlf
     response.write "  }" & vbcrlf

     response.write "  $('#isFinalCheck').val(lcl_isFinalCheck);" & vbcrlf

    'Validate the street number and name entered to determine if it is a valid address in the system for the org
     response.write "  if(lcl_otheraddress == '') {" & vbcrlf
     response.write "     lcl_success = validateAddress();" & vbcrlf

     response.write "     if(lcl_success) {" & vbcrlf
     response.write "        $.post('checkaddress_actionline.asp', {" & vbcrlf
     response.write "           addresstype: '" & lcl_addresstype & "'," & vbcrlf
     response.write "           stnumber:    lcl_streetnumber," & vbcrlf
     response.write "           stname:      lcl_streetname," & vbcrlf
     response.write "           returntype:  'CHECK'" & vbcrlf
     response.write "         }, function(result) {" & vbcrlf
     response.write "           displayValidAddressList(result);" & vbcrlf
     response.write "        });" & vbcrlf
     response.write "     } else {" & vbcrlf
     response.write "        if(lcl_isFinalCheck == 'Y') {"
     response.write "           if(ValidateInput()) {" & vbcrlf
     response.write "              isemailentered();" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "        }" & vbcrlf
     response.write "     }" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     if(lcl_streetnumber != '' || lcl_streetname != '0000') {" & vbcrlf
     response.write "        lcl_success = validateAddress();" & vbcrlf

     response.write "        if(! lcl_success) {" & vbcrlf
     response.write "           FinalCheck('NOT FOUND',1);" & vbcrlf
     response.write "        }" & vbcrlf
     response.write "     } else {" & vbcrlf
     response.write "        if(lcl_isFinalCheck == 'Y') {"
     response.write "           if(ValidateInput()) {" & vbcrlf
     response.write "              isemailentered();" & vbcrlf
     response.write "           }" & vbcrlf
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
     response.write "  var lcl_isFinalCheck = $('#isFinalCheck').val();" & vbcrlf

     response.write "  if (sResults == 'FOUND CHECK') {" & vbcrlf
     response.write "      $('#validstreet').val('Y');" & vbcrlf
     response.write "      $('#validaddresslist').hide('slow');" & vbcrlf
     response.write "      enableDisableAddressFields('');" & vbcrlf

     response.write "      if(lcl_isFinalCheck == 'Y') {"
     response.write "         if(ValidateInput()) {" & vbcrlf
     response.write "            isemailentered();" & vbcrlf
     response.write "         }" & vbcrlf
     response.write "      }" & vbcrlf
     response.write "  } else if (sResults == 'SUBMIT') {" & vbcrlf
     response.write "      if($('#ques_issue2').val() == '') {" & vbcrlf
     response.write "         var lcl_streetnumber = $('#residentstreetnumber').val();" & vbcrlf
     response.write "         var lcl_streetname   = $('#streetaddress').val();" & vbcrlf
     response.write "      }" & vbcrlf

     response.write "      if(iFalseCount > 0) {" & vbcrlf
     response.write "         return false;" & vbcrlf
     response.write "      } else {" & vbcrlf
     response.write "         $('#frmRequestAction').submit();" & vbcrlf
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

     response.write "           if(lcl_isFinalCheck == 'Y') {"
     response.write "              if(ValidateInput()) {" & vbcrlf
     response.write "                 isemailentered();" & vbcrlf
     response.write "              }" & vbcrlf
     response.write "           }" & vbcrlf
     response.write "      }else{" & vbcrlf
     response.write "           if($('#ques_issue2').val() != '') {" & vbcrlf
     response.write "              $('#validaddresslist').hide('slow');" & vbcrlf
     response.write "              enableDisableAddressFields('');" & vbcrlf

     response.write "              if(lcl_isFinalCheck == 'Y') {"
     response.write "                 if(ValidateInput()) {" & vbcrlf
     response.write "                    isemailentered();" & vbcrlf
     response.write "                 }" & vbcrlf
     response.write "              }" & vbcrlf
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

     response.write "  if(lcl_mode == 'disabled') {" & vbcrlf
     response.write "     $('#residentstreetnumber').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#streetaddress').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#ques_issue2').prop('disabled','disabled');" & vbcrlf
     response.write "     $('#validateAddress').prop('disabled','disabled');" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     $('#residentstreetnumber').prop('disabled','');" & vbcrlf
     response.write "     $('#streetaddress').prop('disabled','');" & vbcrlf
     response.write "     $('#ques_issue2').prop('disabled','');" & vbcrlf

     response.write "     if($('#ques_issue2').val() != '') {" & vbcrlf
     response.write "        $('#validateAddress').prop('disabled','disabled');" & vbcrlf
     response.write "     } else {" & vbcrlf
     response.write "        if($('#residentstreetnumber').val() != '') {" & vbcrlf
     response.write "           $('#validateAddress').prop('disabled','');" & vbcrlf
     response.write "        } else {" & vbcrlf
     response.write "           if($('#streetaddress').val() == '0000') {" & vbcrlf
     response.write "             $('#validateAddress').prop('disabled','disabled');" & vbcrlf
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
        response.write "  var lcl_isFinalCheck = $('#isFinalCheck').val();" & vbcrlf

       'Determine if the address is "valid" based on the records in egov_residentaddresses for the org
        response.write "  if(iResult == 'FOUND CHECK' || iResult == 'CANCEL') {" & vbcrlf
        response.write "     if(iResult == 'FOUND CHECK') {" & vbcrlf
        response.write "        displayScreenMsg('Address is Valid');" & vbcrlf
        response.write "        $('#validstreet').val('Y');" & vbcrlf
        response.write "     }" & vbcrlf
        response.write "     $('#validaddresslist').hide('slow');" & vbcrlf
        response.write "     enableDisableAddressFields('');" & vbcrlf

        response.write "     if(iResult != 'CANCEL' && lcl_isFinalCheck == 'Y') {" & vbcrlf
        //response.write "     if(iResult != 'CANCEL') {" & vbcrlf
        response.write "        if(ValidateInput()) {" & vbcrlf
        response.write "           isemailentered();" & vbcrlf
        response.write "        }" & vbcrlf
        response.write "     }" & vbcrlf
        response.write "  } else { " & vbcrlf
        response.write "     displayScreenMsg('Invalid Address');" & vbcrlf
        response.write "     $('#validstreet').val('N');" & vbcrlf
        response.write "     $('#oldstnumber').val(lcl_streetnumber);" & vbcrlf
        response.write "     $('#stname').val(lcl_streetname);" & vbcrlf

        response.write "     enableDisableAddressFields('disabled');" & vbcrlf

        response.write "     $('#validaddresslist').show('slow', function() {" & vbcrlf
        response.write "        $.post('checkaddress_actionline.asp', {" & vbcrlf
        response.write "           addresstype: '" & lcl_addresstype & "'," & vbcrlf
        response.write "           stnumber:    lcl_streetnumber," & vbcrlf
        response.write "           stname:      lcl_streetname," & vbcrlf
        response.write "           returntype:  'DISPLAY_OPTIONS'" & vbcrlf
        response.write "        }, function(result) {" & vbcrlf
        response.write "           $('#addresspicklist').html(result);" & vbcrlf
        response.write "        });" & vbcrlf
        response.write "     });" & vbcrlf
        response.write "  }" & vbcrlf
        response.write "}" & vbcrlf
       'END: Display Valid Address List --------------------------------------

       'BEGIN: Do Select -----------------------------------------------------
        response.write "function doSelect() {" & vbcrlf
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
     response.write "  } else {" & vbcrlf
     response.write "     FinalCheck('SUBMIT',0);" & vbcrlf
     response.write "  }" & vbcrlf

  else
     response.write "  if(lcl_false_count > 0) {" & vbcrlf
     response.write "     return false;" & vbcrlf
     response.write "  } else {" & vbcrlf
     response.write "     document.getElementById(""frmRequestAction"").submit();" & vbcrlf
     response.write "     return true;" & vbcrlf
     response.write "  } " & vbcrlf
  end if

  response.write "}" & vbcrlf
 'END: Validate Fields --------------------------------------------------------
%>

		function ValidateInput() {
			//document.getElementById('btnSubmit').disabled = true;
			var response = grecaptcha.getResponse();

			if(response.length == 0)
			{
    				//reCaptcha not verified
				alert("Sorry, but the CAPTCHA field is required. Please check that box before submitting again.");
				return false;
			}

			if(document.frmRequestAction.frmsubjecttext.value != '') {
   			alert("Please remove any input from the Internal Only field at the bottom of the form.");
				  document.frmRequestAction.frmsubjecttext.focus();
//  				document.getElementById('btnSubmit').disabled=false;
		  		return false;
			}

			//Determine if the Contact Phone is displayed or not.
			lcl_showContactPhone = document.getElementById("showContactPhone").value;

			if(lcl_showContactPhone=="True") {
   			// Set the Phone number
   				var Phone_exists = eval(document.frmRequestAction["cot_txtDaytime_Phone"]);

   				if(Phone_exists) {
			      	document.frmRequestAction.cot_txtDaytime_Phone.value = document.frmRequestAction.skip_user_areacode.value + document.frmRequestAction.skip_user_exchange.value + document.frmRequestAction.skip_user_line.value;
   				}

   				// Set the Fax
   				var Fexists = eval(document.frmRequestAction["cot_txtFax"]);

   				if(Fexists) {
   		   		document.frmRequestAction.cot_txtFax.value = document.frmRequestAction.skip_fax_areacode.value + document.frmRequestAction.skip_fax_exchange.value + document.frmRequestAction.skip_fax_line.value;
   				}
			}



			var tf = true;
			<%if request.querystring("actionid") = "17890" then%>
			//Validate Selected Date
			jQuery.ajax({
         			url:    'validateryedate.asp' 
                  			+ '?date=' 
                  			+ $("#startdate").val(),
         			success: function(result) {
					if (result == "NO")
					{
						alert("You must select a non-weekend or holiday for your start date.");
						tf = false;
					}
                  			},
         			async:   false
    			});      
			//Make sure no duplicates in the last year
			jQuery.ajax({
         			url:    'rye_yearcheck.asp' 
                  			+ '?date=' 
                  			+ $("#startdate").val()
                  			+ '&address=' 
                  			+ $("#issuelocation").val(),
         			success: function(result) {
					if (result != "PASS")
					{
						alert("You cannot submit a Registration for this address that starts before " + result + ".");
						tf = false;
					}
                  			},
         			async:   false
    			});      
			if (tf)
			{
				//Make sure no duplicates in the last year
				jQuery.ajax({
         				url:    'rye_yearcheck.asp' 
                  				+ '?date=' 
                  				+ $("#enddate").html()
                  				+ '&address=' 
                  				+ $("#issuelocation").val(),
         				success: function(result) {
						if (result != "PASS")
						{
							alert("You cannot submit a Registration for this address that starts before " + result + ".");
							tf = false;
						}
                  				},
         				async:   false
    				});      
			}
			if ($("#residentaddressid").val() == "0")
			{
				alert("You must select a valid address.");
				tf = false;
			}
			<% end if %>

                	if (document.getElementById("mapLat") && marker != undefined)
                	{
                        	document.getElementById("mapLat").value = marker.getPosition().lat();
                        	document.getElementById("mapLng").value = marker.getPosition().lng();
                	}
			


			
			if (!tf)
			{
				return false
			}
			else 
			{
				return  validateForm('frmRequestAction');
			}

		}

function isemailentered() {
  var lcl_showContactEmail = document.getElementById("showContactEmail").value;

  //1. Check to see if the contact fields are displayed.
  if (lcl_showContactEmail=="True") {

    		//2. If "send confirmation email" is CHECKED
      if (document.frmRequestAction.chkSendEmail.checked == true) {

    						//3. Check to see if the email address was entered
						    if (document.frmRequestAction.cot_txtEmail.value != '') {
  								    //If YES then submit the form
        						document.frmRequestAction.submit();
				    		}else{
  						    		alert('You have selected to have an email confirmation sent, but have not entered an email address. \nPlease either enter an email address or uncheck the option to have a confirmation sent.');
              document.frmRequestAction.cot_txtEmail.focus();
//              document.getElementById('btnSubmit').disabled=false;
    						}
    		}else{
						    document.frmRequestAction.submit();
    		}
  }else{
		    document.frmRequestAction.submit();
  }
}

function autoTab(input,len, e) {
  var keyCode = (isNN) ? e.which : e.keyCode; 
		var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

		if(input.value.length >= len && !containsElement(filter,keyCode)) {
  			input.value = input.value.slice(0, len);
   		var addNdx = 1;

			  while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") {
    				addNdx++;
    				//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
   		}
   		input.form[(getIndex(input)+addNdx) % input.form.length].focus();
		}

function containsElement(arr, ele) {
  var found = false, index = 0;
		while(!found && index < arr.length)
 			if(arr[index] == ele)
	   			found = true;
			 else
				   index++;
  	   	return found;
}

function getIndex(input) {
		var index = -1, i = 0, found = false;

		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
				return index;
		}
		return true;
	}

function ShowMap() {
  var lcl_url_map = '';

		//alert(document.frmRequestAction.skip_address.value);
		//if (document.frmRequestAction.skip_address.value == 0)
		if($('#streetaddress').val() == 0) {
		 		alert('Please Select an address from the list.');
 				//document.frmRequestAction.skip_address.focus();
	 			$('#streetaddress').focus();
 				return;
		}

  lcl_url_map  = '<%=Application("MAP_URL")%>action_line_map.asp';
  lcl_url_map += '?residentaddressid=' + $('#streetaddress').val();

		//eval('window.open("<% 'Application("MAP_URL")%>action_line_map.asp?residentaddressid=' + document.frmRequestAction.skip_address.value + '", "_map", "width=700,height=500,toolbar=0,statusbar=0,resizable,scrollbars=1,menubar=0,left=0,top=0")');
		eval('window.open("' + lcl_url_map + '", "_map", "width=840,height=660,toolbar=0,statusbar=0,resizable,scrollbars=1,menubar=0,left=0,top=0")');
}

		function openWin2(url, name) 
		{
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		}

		// Validate Tracking added by Steve Loar - 12/30/2005
		function ValidateTracking( form )
		{
			//return true;
			var rege = /^\d+$/;
			var Ok = rege.exec(form.REQUEST_ID.value);

			if (! Ok)
			{
				alert ("Tracking Numbers must be numeric. Please try your search again.");
				form.REQUEST_ID.focus();
				form.REQUEST_ID.select();
				return false;
			}
			return true;
		}

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

		function getinfo()
		{
			if (document.frmRequestAction.chkSameAs.checked) 
			{
				// CHECK USE ABOVE
				document.frmRequestAction.ques_issue3.value = document.frmRequestAction.cot_txtCity.value;
				document.frmRequestAction.ques_issue4.value = document.frmRequestAction.cot_txtState_vSlash_Province.value;
				document.frmRequestAction.ques_issue5.value = document.frmRequestAction.cot_txtZIP_vSlash_Postal_Code.value;
			}


			else 
			{
				// UNCHECKED CLEAR VALUES
				document.frmRequestAction.ques_issue3.value = '';
				document.frmRequestAction.ques_issue4.value = '';
				document.frmRequestAction.ques_issue5.value = '';
			}
		}
 	var winHandle;
 	var w = (screen.width - 640)/2;
 	var h = (screen.height - 450)/2;

function validateCheckbox(p_field) {
  //get the total options
  var lcl_total_options = document.getElementById("total_options_"+p_field).innerHTML;
  var lcl_total_checked = 0;

  for (i = 1; i <= lcl_total_options-1; i++) {
       if(document.getElementById(p_field+"_"+i).checked==true) {
          lcl_total_checked = lcl_total_checked + 1;
       }
  }
  if(lcl_total_checked == 0) {
     document.getElementById(p_field+"_"+lcl_total_options).checked=true;
  }else{
     document.getElementById(p_field+"_"+lcl_total_options).checked=false;
  }
}

function searchRequestTypes() {
  lcl_url_params   = "";
  lcl_request_type = document.getElementById("sc_request_type").value;

  if(lcl_request_type !="") {
     lcl_url_params = "?sc_request_type=" + encodeURIComponent(lcl_request_type);
  }

  location.href="action.asp" + lcl_url_params;
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
<!--#Include file="include_top.asp"-->
<%
 'BEGIN: Build the Welcome message --------------------------------------------
  lcl_org_name        = oActionOrg.GetOrgName()
  lcl_org_state       = oActionOrg.GetState()
  lcl_org_featurename = oActionOrg.GetOrgFeatureName("action line")

  oActionOrg.buildWelcomeMessage iorgid, _
                                 lcl_orghasdisplay_action_page_title, _
                                 lcl_org_name, _
                                 lcl_org_state, _
                                 lcl_org_featurename

  response.write "<br />" & vbcrlf

  RegisteredUserDisplay( "" )
 'END: Build the Welcome message ----------------------------------------------

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf

  if trim(request("actionid")) <> "" then
	  on error resume next
    	iActionId       = CLng(Track_DBSafe(request("actionid")))
	if err.number <> 0 then
      		response.redirect("action.asp")
	end if
	on error goto 0
     sFormIsEnabled  = FormIsEnabled(iActionID)
     sFormIsInternal = isFormInternal(iActionID)

    	if IsNumeric(iActionId) AND sFormIsEnabled AND NOT sFormIsInternal then
       'Do not need to perform a CLNG.  It is overkill.
       'At this point we already know that the value is numeric
       'iActionId = CLng(iActionId)  'Commented out for bug fix if value is a number, but too long for a CLNG (March 08, 2008)
       'We are also checking to ensure that the form IS enabled.
     			subDisplayActionForm iActionId, _
                             iorgid
     else
      		response.redirect("action.asp")
     end if
  else

    'BEGIN: Visitor Tracking --------------------------------------------------
    	iSectionID     = 2
    	sDocumentTitle = "MAIN"
    	sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
    	datDate        = Date()	
    	datDateTime    = Now()
    	sVisitorIP     = request.servervariables("REMOTE_ADDR")
    	'Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
    'END: Visitor Tracking ----------------------------------------------------

     response.write "<div id=""formsDiv"">" & vbcrlf
     response.write "<ul>" & vbcrlf

    'BEGIN: Right-side Column (floated to right) ------------------------------
     response.write "  <li id=""rightSideDiv"">" & vbcrlf

    'FAQ Button
     if iorgid = 15 then 
        response.write "<p><input type=""button"" onclick=""window.location='faq.asp'"" value=""Frequently Asked Questions"" class=""button"" /></p>" & vbcrlf
     end if

    'Register/Login Links
     if sOrgRegistration AND (lcl_cookie_userid = "" OR lcl_cookie_userid = "-1") AND lcl_orghasfeature_hide_login = false then
        response.write "    <fieldset class=""fieldset"">" & vbcrlf
        response.write "      <legend>Personalized Services</legend>" & vbcrlf
        response.write "      &#8226;&nbsp;<a href=""user_login.asp"">Click here to Login</a><br />" & vbcrlf
        response.write "      &#8226;&nbsp;<a href=""register.asp"">Click here to Register</a>" & vbcrlf
        response.write "    </fieldset>" & vbcrlf
     end if

     response.write "<div>" & vbcrlf
     response.write sActionDescription & vbcrlf
     response.write "</div>" & vbcrlf

     response.write "  </li>" & vbcrlf
    'END: Right-side Column (floated to right) --------------------------------

    'BEGIN: Left-side Column --------------------------------------------------
     response.write "  <li>" & vbcrlf

    'BEGIN: Tracking Lookup ---------------------------------------------------
     lcl_displaytitle_actiontracking = "Check the Status of an " & lcl_org_featurename & " Request"

     if lcl_orghasdisplay_action_tracking_title then
        lcl_displaytitle_actiontracking = GetOrgDisplay(iorgid,"action tracking title")
     end if

     response.write "    <fieldset class=""fieldset"">" & vbcrlf
     response.write "      <legend>" & lcl_displaytitle_actiontracking & "</legend>" & vbcrlf
     response.write "      <form name=""frmActionLookup"" id=""frmActionLookup"" action=""action_request_lookup.asp"" method=""post"" onsubmit=""return ValidateTracking(this);"">" & vbcrlf
     response.write "      <table border=""0"" id=""searchTable"">" & vbcrlf
     response.write "        <tr>" & vbcrlf
     response.write "            <td>" & vbcrlf
     response.write "                <span class=""label"">Tracking Number:</span>" & vbcrlf
     response.write "                <input type=""text"" name=""REQUEST_ID"" id=""REQUEST_ID"" />" & vbcrlf
     response.write "                <input type=""button"" name=""searchButton_trackingLookup"" id=""searchButton_trackingLookup"" value=""Search"" />" & vbcrlf
     response.write "            </td>" & vbcrlf
     response.write "        </tr>" & vbcrlf
     response.write "      </table>" & vbcrlf
     response.write "      </form>" & vbcrlf
     response.write "    </fieldset>" & vbcrlf
    'END: Tracking Lookup -----------------------------------------------------

    'BEGIN: Display Forms -----------------------------------------------------
     lcl_displaytitle_actionlist = "Create a New " & lcl_org_featurename & " Request"

     if lcl_orghasdisplay_action_list_title then
        lcl_displaytitle_actionlist = GetOrgDisplay( iOrgId, "action list title" )
     end if

     response.write "    <fieldset class=""fieldset"">" & vbcrlf
     response.write "      <legend>" & lcl_displaytitle_actionlist & "</legend>" & vbcrlf
     response.write "      <fieldset class=""nestedFieldset"">" & vbcrlf
     response.write "        <legend>Search for Request Type(s)</legend>" & vbcrlf
     response.write "        <form name=""frmSearchForms"" id=""frmSearchForms"" action=""action.asp?list=true"" method=""post"">" & vbcrlf
     response.write "        <table border=""0"" id=""requestTypeTable"">" & vbcrlf
     response.write "          <tr>" & vbcrlf
     response.write "              <td>" & vbcrlf
     response.write "                  <span class=""label"">Request Type:</span>" & vbcrlf
     response.write "                  <input type=""text"" name=""sc_request_type"" id=""sc_request_type"" value=""" & lcl_sc_request_type & """ maxlength=""50"" />" & vbcrlf
     response.write "                  <input type=""button"" name=""searchButton_requestTypes"" id=""searchButton_requestTypes"" value=""Search"" />" & vbcrlf
     response.write "              </td>" & vbcrlf
     response.write "          </tr>" & vbcrlf
     response.write "        </table>" & vbcrlf
     response.write "        </form>" & vbcrlf
     response.write "      </fieldset>" & vbcrlf

     response.write "      <div>" & vbcrlf
                             fnListForms iorgid, _
                                         lcl_sc_request_type
     response.write "      </div>" & vbcrlf

     response.write "    </fieldset>" & vbcrlf
    'END: Display Forms -------------------------------------------------------

     response.write "  </li>" & vbcrlf
    'END: Left-side Column ----------------------------------------------------

     response.write "</ul>" & vbcrlf
     response.write "</div>" & vbcrlf

  end if

  set oActionOrg = nothing 

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="include_bottom.asp"-->  
<%
'------------------------------------------------------------------------------
sub fnListForms(p_orgid, _
                p_sc_request_type)
	dim sOrgID

 sOrgID          = 0
	sLastCategory   = "NONE_START"
 iSC_RequestType = ""

 if p_orgid <> "" then
    sOrgID = clng(p_orgid)
 end if

 if p_sc_request_type <> "" then
    iSC_RequestType = ucase(p_sc_request_type)
    iSC_RequestType = dbsafe(iSC_RequestType)
    iSC_RequestType = "'%" & iSC_RequestType & "%'"
 end if

	sSQL = "SELECT * "
 sSQL = sSQL & " FROM dbo.egov_form_list_200 "
 sSQL = sSQL & " WHERE orgid = " & sOrgID
 sSQL = sSQL & " AND form_category_id <> 6 "
 sSQL = sSQL & " AND action_form_internal <> 1 "
 sSQL = sSQL & " AND action_form_displayOnList = 1 "

 if iSC_RequestType <> "" then
    sSQL = sSQL & " AND (UPPER(form_category_name) LIKE (" & iSC_RequestType & ") "
    sSQL = sSQL & "  OR  UPPER(action_form_name) LIKE ("   & iSC_RequestType & ")) "
 end if

 sSQL = sSQL & " ORDER BY form_category_sequence, action_form_name "
 'response.write sSQL & "<br /><br />"

	set oForms = Server.CreateObject("ADODB.Recordset")
	oForms.Open sSQL, Application("DSN") , 3, 1
	
	if not oForms.eof then
	   do while not oForms.eof
   			 sTopic                          = Server.URLEncode(sCurrentCategory & " > " & oForms("action_form_name"))
    			sCurrentCategory                = oForms("form_category_name")
       sDisplayMobileOptions_geoLoc    = oForms("display_mobileoptions_geoloc")
       sDisplayMobileOptions_takePic   = oForms("display_mobileoptions_takepic")
       sActionURL                      = "action.asp?actionid=" & oForms("action_form_id")

  		  	if sLastCategory = "NONE_START" then
       			response.write "<p class=""actionCategory"">" & vbcrlf
			      	response.write "  <strong><a class=""actionjump"" name=""" & oForms("form_category_id")  & """>" & sCurrentCategory & "</a></strong>" & vbcrlf
          response.write "</p>" & vbcrlf
       end if

    			if (sCurrentCategory <> sLastCategory) AND (sLastCategory <> "NONE_START") then
  		     	response.write "<p class=""actionCategory"">" & vbcrlf
		      		response.write "  <strong><a class=""actionjump"" name=""" & oForms("form_category_id")  & """>" & sCurrentCategory & "</a></strong>" & vbcrlf
          response.write "</p>" & vbcrlf
       end if

    'Determine if "Mobile Options" have been turned on.
    '  If "yes" then check to see if the "Mobile Options - Take Pic" option is enabled.
    '  If "yes" then change the URL to the action line picture taking screen.
     'lcl_isMobileDevice = true

     'if lcl_isMobileDevice then
     '   if lcl_orghasfeature_actionline_formcreator_mobileoptions then
     '      sActionURL = "action_takepic.asp?formid=" & oForms("action_form_id")
           'response.redirect "action_takepic.asp?formid=" & iFormID
     '   end if
     'end if

  		  	response.write "<p class=""actionItem"">" & vbcrlf
    			response.write "-&nbsp;<a href=""" & sActionURL & """>" & oForms("action_form_name") &  "</a>" & vbcrlf
       response.write "</p>" & vbcrlf

    			oForms.movenext

    			sLastCategory = sCurrentCategory
    loop
 else
	   response.write "<p style=""padding-top:10px;"" class=""alertmsg"">" & vbcrlf
    response.write "  <center>No action forms enabled.</center>" & vbcrlf
    response.write "</p>" & vbcrlf
	end if

	Set oForms = Nothing 

end sub

'------------------------------------------------------------------------------
sub subDisplayActionForm( ByVal iFormID, ByVal iorgid )
	 dim sSql, oForm

 'Declare variables
 	dim sTitle,sIntroText,sFooterText,sMask,blnEmergencyNote,sEmergencyText,blnIssueDisplay
 	dim sIssueMask,iStreetNumberInputType,iStreetAddressInputType,sIssueName,sIssueDesc,sIssueQues,sHideIssueLocAddInfo

	'Get form information
 	sSQL = "SELECT * "
	 sSQL = sSQL & " FROM egov_action_request_forms "
 	sSQL = sSQL & " WHERE action_form_id = '" & CLng(iFormID) & "'"

 	set oForm = Server.CreateObject("ADODB.Recordset")
	 oForm.Open sSQL, Application("DSN"), 3, 1

 	if not oForm.eof then
   		sTitle                          = oForm("action_form_name")
   		sIntroText                      = oForm("action_form_description")
   		sFooterText                     = oForm("action_form_footer")
	   	sMask                           = oForm("action_form_contact_mask")
	  	 blnEmergencyNote                = oForm("action_form_emergency_note")
  	 	sEmergencyText                  = oForm("action_form_emergency_text")
 	  	blnIssueDisplay                 = oForm("action_form_display_issue")
	  	 sIssueMask                      = oForm("action_form_issue_mask")
   		iStreetNumberInputType          = oForm("issuestreetnumberinputtype")
	   	iStreetAddressInputType         = oForm("issuestreetaddressinputtype")
	 	  sIssueName                      = oForm("issuelocationname")
   		sIssueDesc                      = oForm("issuelocationdesc")
   		sIssueQues                      = oForm("issuequestion")
   		sHideIssueLocAddInfo            = oForm("hideIssueLocAddInfo")
       		blnDisplayMobileOptions_takePic   = oForm("display_mobileoptions_takepic")
       		'blnDisplayMobileOptions_geoLoc   = oForm("display_mobileoptions_geoloc")
		blnShowMapInput 		= oForm("showmapinput")
	 'response.write blnDisplayMobileOptions_takePic  & "####"

   		if trim(sIssueName) = "" OR isnull(sIssueName) then
      		sIssueName = "Issue/Problem Location:"
   		end if

   		if isnull(sIssueDesc) then
        sIssueDesc = "Please select the closest street number/streetname of problem "
        sIssueDesc = sIssueDesc & "location from list or select ""Choose street from dropdown"". "
        sIssueDesc = sIssueDesc & "Provide any additional information on problem location in the box below."
     end if

  		'Determine if the Contact Info fields are displayed
  		 lcl_checkForContactFields = 0
	  	 lcl_displayContactFields  = False
		   lcl_displayContactEmail   = False

   		for i = 1 to 10
     			if isDisplay(sMASK,i) then
       				lcl_checkForContactFields = lcl_checkForContactFields + 1
     			else
       				lcl_checkForContactFields = lcl_checkForContactFields
    	 		end if
   		next

  		'If the value is greater than zero is there is at least one contact field that is displayed/required.
	   	if lcl_checkForContactFields > 0 then
     			lcl_displayContactFields = True
    	end if
 	else
	   	response.redirect("action.asp")
 	end if

 	oForm.Close
 	set oForm = nothing 

	'BEGIN: Visitor Tracking -----------------------------------------------------
 	iSectionID     = 22
 	sDocumentTitle = sTitle
 	sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
 	datDate        = Date()	
 	datDateTime    = Now()
 	sVisitorIP     = request.servervariables("REMOTE_ADDR")
 	Call LogPageVisit(iSectionID,sDocumentTitle,sURL,datDate,datDateTime,sVisitorIP,iorgid)
	'END: Visitor Tracking -------------------------------------------------------

	'if iorgid = "37" then
 		'response.write "<form name=""frmRequestAction"" id=""frmRequestAction"" action=""action_cgi_twf.asp?list=true"" enctype=""multipart/form-data"" method=""post"">" & vbcrlf
	'elseif iorgid = "5" then
 		'response.write "<form name=""frmRequestAction"" id=""frmRequestAction"" action=""action_cgi_twf.asp?list=true"" method=""post"">" & vbcrlf
	'else
 		'response.write "<form name=""frmRequestAction"" id=""frmRequestAction"" action=""action_cgi.asp?list=true""  method=""post"">" & vbcrlf
	 if blnDisplayMobileOptions_takePic and lcl_orghasfeature_actionline_formcreator_mobileoptions then
 		response.write "<form name=""frmRequestAction"" id=""frmRequestAction"" action=""action_cgi.asp?list=true""  enctype=""multipart/form-data"" method=""post"">" & vbcrlf
	else
 		response.write "<form name=""frmRequestAction"" id=""frmRequestAction"" action=""action_cgi.asp?list=true""  method=""post"">" & vbcrlf
	end if
 	response.write "  <input type=""" & lcl_hidden & """ name=""actionid"" value=""" & iFormID & """ />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""actiontitle"" value=""" & sTitle & """ />" & vbcrlf
 	'response.write "  <input type=""" & lcl_hidden & """ name=""validstreet"" value=""N"" />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""control_field"" id=""control_field"" value="""" size=""20"" maxlength=""4001"" />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""showContactInfo"" id=""showContactInfo"" value=""" & lcl_displayContactFields & """ size=""5"" maxlength=""10"" />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""showContactEmail"" id=""showContactEmail"" value=""" & isDisplay(sMASK,4) & """ size=""5"" maxlength=""10"" />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""showContactPhone"" id=""showContactPhone"" value=""" & isDisplay(sMASK,5) & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""" & lcl_hidden & """ name=""isFinalCheck"" id=""isFinalCheck"" value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf
 	response.write "<div class=""respIndent"" >" & vbcrlf

 'BEGIN: Screen Message -------------------------------------------------------
  response.write "<div id=""screenMsg""></div>" & vbcrlf
 'END: Screen Message ---------------------------------------------------------


	'BEGIN: Emergency Note -------------------------------------------------------
 	if blnEmergencyNote then
	   	'response.write "<div class=""warning"">" & sEmergencyText & "</div>" & vbcrlf
     response.write "<fieldset class=""fieldset_warning"">" & vbcrlf
     response.write    sEmergencyText & vbcrlf
     response.write "</fieldset>" & vbcrlf
 	end if
	'END: Emergency Note ---------------------------------------------------------

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>" & sTitle & "</legend>" & vbcrlf

	'BEGIN: Title ----------------------------------------------------------------
' 	response.write "<div class=""box_header4"">" & sTitle & "</div>" & vbcrlf
	'END: Title ------------------------------------------------------------------

' 	response.write "<div class=""group"">" & vbcrlf
'	 response.write "<div class=""orgadminboxf"">" & vbcrlf

	'BEGIN: Register/Login Links -------------------------------------------------
 	if sOrgRegistration AND (lcl_cookie_userid = "" OR lcl_cookie_userid = "-1") and lcl_orghasfeature_hide_login = false then
	   	response.write "<div id=""loginRegisterFieldset"" align=""right"">" & vbcrlf
   		response.write "  <a href=""user_login.asp"">Click here to Login</a> | " & vbcrlf
   		response.write "  <a href=""register.asp"">Click here to Register</a>" & vbcrlf
   		response.write "</div>" & vbcrlf
 	end if
	'END: Register/Login Links --------------------------------------------------

	'BEGIN: Intro Information ---------------------------------------------------
		lcl_intro_text = " - <i>Introduction text is currently blank</i> - "

 	if sIntroText <> "" then
	   	lcl_intro_text = sIntroText
 	end if

	 response.write "<p>" & lcl_intro_text & "</p>" & vbcrlf
	'END: Intro Information ------------------------------------------------------
 	if iorgid = "37" then
		%><link href="//www.egovlink.com/eclink/css/otherreports.css" rel="stylesheet"><%
		OtherReports
	end if

	response.write "<p class=""alertmsg"">* Information is required.</p>" & vbcrlf

	'BEGIN: Contact Information --------------------------------------------------
 	if lcl_displayContactFields then
	   	response.write "<p>" & vbcrlf
   		response.write "   <span class=""sectionTitle"">Contact Information</span><br />" & vbcrlf
   		response.write "   <table class=""tblResponsive"">" & vbcrlf
                          DrawContactTable(sMask)
   		response.write "   </table>" & vbcrlf

     if lcl_cookie_userid <> "" and lcl_cookie_userid <> "-1" then
        response.write "<div id=""contactInfoMsg"">" & vbcrlf
        response.write "  *** Contact information will be updated when the form is submitted. ***" & vbcrlf
        response.write "</div>" & vbcrlf
     end if

   		response.write "</p>" & vbcrlf
  end if
	'END: Contact Information ----------------------------------------------------

	'BEGIN: Problem/Issue Location -----------------------------------------------
 	if lcl_orghasfeature_issue_location then
	   	if blnIssueDisplay then
     			response.write "<p>" & vbcrlf
		  	   response.write "   <span id=""sIssueName"" class=""sectionTitle"">" & sIssueName & "</span>" & vbcrlf
   	  		response.write "   <p>" & sIssueDesc & "</p>" & vbcrlf
                           displayIssueLocation_new iorgid, _
                                                    lcl_orghasfeature_issue_location, _
                                                    lcl_orghasfeature_large_address_list, _
                                                    sIssueMask, _
                                                    sIssueQues, _
                                                    sHideIssueLocAddInfo

                        			'subDisplayIssueLocation sIssueMask, _
                           '                        iStreetNumberInputType, _
                           '                        iStreetAddressInputType, _
                           '                        sIssueQues, _
                           '                        sHideIssueLocAddInfo
     			response.write "</p>" & vbcrlf
   		end if
 	end if
	'END: Problem/Issue Location -------------------------------------------------

	'BEGIN: Form Field Information -----------------------------------------------
 	response.write "<p>" & vbcrlf
	                   subDisplayQuestions iFormID, sMask
 	response.write "</p>" & vbcrlf
	 response.write "<p class=""alertmsg"">* Information is required.</p>" & vbcrlf
	'END: Form Field Information -------------------------------------------------

	if request.querystring("actionid") <> "17890" then
	'BEGIN: Ending Notes ---------------------------------------------------------
 	lcl_footer_text = " - <i>Footer text is currently blank</i>"

	 if sFooterText <> "" then
   		lcl_footer_text = sFooterText
 	end if

	 response.write "<p>" & lcl_footer_text & "</p>" & vbcrlf
	'END: Ending Notes -----------------------------------------------------------
	end if

  response.write "</fieldset>" & vbcrlf

	'BEGIN: Problem Field --------------------------------------------------------
 	response.write "<div id=""problemtextfield1"">" & vbcrlf
	 response.write "  Internal Use Only, Leave Blank: <input type=""text"" name=""frmsubjecttext"" id=""problemtextinput"" value="""" size=""6"" />" & vbcrlf
 	response.write "  <input type=""" & lcl_hidden & """ name=""problemorg"" value=""" & iorgid & """ /><br />" & vbcrlf
	 response.write "  <strong>Please leave this field blank and remove any values that have been populated for it.</strong>" & vbcrlf
 	response.write "</div>" & vbcrlf
	'END: Problem Field ----------------------------------------------------------

	response.write "</form>" & vbcrlf

end sub

'------------------------------------------------------------------------------
function DrawContactTable(sMASK)

  ' BEGIN: GET USER PERSONNEL INFORMATION IF USER IS LOGGED INTO WEBSITE
  If sOrgRegistration Then 
    If lcl_cookie_userid <> "" and lcl_cookie_userid <> "-1" Then
		
    		iUserID = lcl_cookie_userid
	
    		sSQL = "SELECT * FROM egov_users WHERE userid = " & iUserID
    		Set oInfo = Server.CreateObject("ADODB.Recordset")
    		oInfo.Open sSQL, Application("DSN") , 3, 1

    		If NOT oInfo.EOF Then
     			'USER FOUND SET VALUES
      			sFirstName    = oInfo("userfname")
   						sLastName     = oInfo("userlname")
   						sAddress      = oInfo("useraddress")
   						sCity         = oInfo("usercity")
   						sState        = oInfo("userstate")
   						sZip          = oInfo("userzip")
   						sEmail        = oInfo("useremail")
   						sHomePhone    = oInfo("userhomephone")
   						sWorkPhone    = oInfo("userworkphone")
   						sBusinessName = oInfo("userbusinessname")
   						sFax          = oInfo("userfax")
			 		Else
   					'USER NOT FOUND SET VALUES TO EMPTY
   						sFirstName    = ""
   						sLastName     = ""
   						sAddress      = ""
   						sCity         = ""
   						sState        = ""
   						sZip          = ""
   						sEmail        = ""
   						sHomePhone    = ""
   						sWorkPhone    = ""
   						sFax          = ""
   						sBusinessName = ""
   		 End If

    		Set oInfo = Nothing
    End If
  end if

  'First Name -------------------------------------------------------------------
  if isDisplay(sMASK,1) then
    if isRequired(sMASK,1) <> "" then
     		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtFirst_Name-text/req"" value=""First Name"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
    response.write "            <span class=""cot-text-emphasized"">" & isRequired(sMASK,1) & "</span>" & vbcrlf
    response.write "            First Name:" & vbcrlf
    response.write "          </span>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "          <span class=""cot-text-emphasized"" title=""This field is required""> " & vbcrlf
    response.write "        		  <input type=""text"" value=""" & sFirstName & """ name=""cot_txtFirst_Name"" id=""txtFirst_Name"" style=""width:150px;"" maxlength=""50"" />" & vbcrlf
    response.write "        		</span>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Last Name --------------------------------------------------------------------
 if isDisplay(sMASK,2) then
    if isRequired(sMASK,2) <> "" then
       response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtLast_Name-text/req"" value=""Last Name"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write "          <span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
    response.write "            <span class=""cot-text-emphasized"">" & isRequired(sMASK,2) & "</span>" & vbcrlf
    response.write "            Last Name:" & vbcrlf
    response.write "          </span>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "        		<span class=""cot-text-emphasized"" title=""This field is required"">" & vbcrlf
    response.write "          		<input type=""text"" value=""" & sLastName & """ name=""cot_txtLast_Name"" id=""txtLast_Name"" style=""width:150px;"" maxlength=""50"" />" & vbcrlf
    response.write "        		</span>" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Business Name ----------------------------------------------------------------
 if isDisplay(sMASK,3) then
    if isRequired(sMASK,3) <> "" then
       response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtBusiness_Name-text/req"" value=""Business Name"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,3) & "Business Name:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "        		<input type=""text"" value=""" & sBusinessName & """ name=""cot_txtBusiness_Name"" id=""txtBusiness_Name"" style=""width:300px;"" maxlength=""255"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Email -----------------------------------------------------------------------
 if isDisplay(sMASK,4) then
   	if isRequired(sMASK,4) <> "" then
     		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtEmail-text/req"" value=""Email Address"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,4) & "Email:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "          <input type=""text"" value=""" & sEmail & """ name=""cot_txtEmail"" id=""txtEmail"" style=""width:300px;"" maxlength=""512"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 'else
 '   response.write "  <input type=""text"" value=""" & sEmail & """ name=""cot_txtEmail"" id=""txtEmail"" style=""width:300px;"" maxlength=""512"" />" & vbcrlf
 end if

'Daytime Phone ---------------------------------------------------------------
 if isDisplay(sMASK,5) then
   	if isRequired(sMASK,5) <> "" then
     		response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtDaytime_Phone-text/number/req"" value=""Daytime Phone"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,5) & "Daytime Phone:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "        		<input type=""" & lcl_hidden & """ value=""" & sHomePhone & """ name=""cot_txtDaytime_Phone"" />" & vbcrlf
    response.write "       		(<input class=""phonenum"" type=""text"" value=""" & left(sHomePhone,3) & """ name=""skip_user_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
    response.write "        		<input class=""phonenum"" type=""text"" value=""" & mid(sHomePhone,4,3) & """ name=""skip_user_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
    response.write "        		<input class=""phonenum"" type=""text"" value=""" & mid(sHomePhone,7,4) & """ name=""skip_user_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Fax --------------------------------------------------------------------------
 if isDisplay(sMASK,6) then
    if isRequired(sMASK,6) <> "" then
    			response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtFax/req"" value=""Fax"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,6) & "Fax:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "      			(<input class=""phonenum"" type=""text"" value=""" & left(sFax,3) & """ name=""skip_fax_areacode"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />)&nbsp;" & vbcrlf
    response.write "      				<input class=""phonenum"" type=""text"" value=""" & mid(sFax,4,3) & """ name=""skip_fax_exchange"" onKeyUp=""return autoTab(this, 3, event);"" size=""3"" maxlength=""3"" />&ndash;" & vbcrlf
    response.write "      				<input class=""phonenum"" type=""text"" value=""" & right(sFax,4) & """ name=""skip_fax_line"" onKeyUp=""return autoTab(this, 4, event);"" size=""4"" maxlength=""4"" />" & vbcrlf
    response.write "      				<input type=""" & lcl_hidden & """ value=""" & sFax & """ name=""cot_txtFax"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Address ----------------------------------------------------------------------
 if isDisplay(sMASK,7) then
    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,7) & "Address:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "       			<input type=""text"" value=""" & sAddress & """ name=""cot_txtStreet"" id=""txtStreet"" style=""width:300px;"" maxlength=""255"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf

    if isRequired(sMASK,7) <> "" then
    			response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtStreet/req"" value=""Street"" />" & vbcrlf
    end if
 end if

'City -------------------------------------------------------------------------
 if isDisplay(sMASK,8) then
    if isRequired(sMASK,8) <> "" then
       response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtCity/req"" value=""City"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,8) & "City:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "        		<input type=""text"" value=""" & sCity & """ name=""cot_txtCity"" id=""txtCity"" style=""width:300;"" maxlength=""50"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'State ------------------------------------------------------------------------
 if isDisplay(sMASK,9) then
    if isRequired(sMASK,9) <> "" then
       response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtState_vSlash_Province/req"" value=""State"" />" & vbrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,9) & "State:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "        		<input type=""text"" value=""" & sState & """ name=""cot_txtState_vSlash_Province"" id=""txtState_vSlash_Province"" size=""5"" maxlength=""2"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if

'Zip --------------------------------------------------------------------------
 if isDisplay(sMASK,10) then
    if isRequired(sMASK,10) <> "" then
       response.write "<input type=""" & lcl_hidden & """ name=""ef:cot_txtZIP_vSlash_Postal_Code/req"" value=""Zipcode"" />" & vbcrlf
    end if

    response.write "  <tr>" & vbcrlf
    response.write "      <td align=""right"">" & vbcrlf
    response.write            isRequired(sMASK,10) & "ZIP:" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "      <td>" & vbcrlf
    response.write "         	<input type=""text"" value=""" & sZip & """ name=""cot_txtZIP_vSlash_Postal_Code"" id=""txtZIP_vSlash_Postal_Code"" size=""10"" maxlength=""10"" />" & vbcrlf
    response.write "      </td>" & vbcrlf
    response.write "  </tr>" & vbcrlf
 end if
end function

'------------------------------------------------------------------------------
function DBsafe( ByVal strDB )
 	Dim sNewString

 	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
 	sNewString = Replace( strDB, "'", "''" )
 	sNewString = Replace( sNewString, "<", "&lt;" )
 	DBsafe = sNewString

end function

'------------------------------------------------------------------------------
sub subDisplayQuestions( ByVal iFormID, ByVal sMask )
	Dim sSQL, oQuestions

	sSQL = "SELECT * "
	sSQL = sSQL & " FROM egov_action_form_questions "
	sSQL = sSQL & " WHERE formid = " & CLng(iFormID)
	sSQL = sSQL & " AND (isinternalonly <> 1 OR isinternalonly IS NULL) "
	'sSQL = sSQL & " AND orgid = " & iorgid
	sSQL = sSQL & " ORDER BY sequence"

 	set oQuestions = Server.CreateObject("ADODB.Recordset")
	oQuestions.Open sSQL, Application("DSN") , 3, 1

 	if not oQuestions.eof then
	  		response.write "<table class=""tblResponsive"">" & vbcrlf

    	while not oQuestions.eof
        iQuestionCount = iQuestionCount + 1

     		'Determine if required
      		sisRequired = oQuestions("isRequired")

      		if sisRequired = True then
         		sisRequired = " <font color=""#ff0000"">*</font> "
      		else
        			sisRequired = ""
      		end if

    		 'Tracking current form configuration for editing later
        lcl_answerlist = oQuestions("answerlist")

        if lcl_answerlist <> "" then
           lcl_answerlist = replace(lcl_answerlist,"""","&quot;")
        end if

        response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("fieldtype")   & """ name=""fieldtype"" />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ value=""" & lcl_answerlist            & """ name=""answerlist"" />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("isRequired")  & """ name=""isRequired"" />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("sequence")    & """ name=""sequence"" />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("pdfformname") & """ name=""pdfformname"" />" & vbcrlf
        response.write "<input type=""" & lcl_hidden & """ value=""" & oQuestions("pushfieldid") & """ name=""pushfieldid"" />" & vbcrlf

       'Format the Prompt and Value
        lcl_prompt_req_value = oQuestions("prompt")
        lcl_prompt_display   = oQuestions("prompt")
        lcl_prompt_value     = oQuestions("prompt")

        if lcl_prompt_value <> "" then
           lcl_prompt_value = replace(lcl_prompt_value,"""","&quot;")
        end if

        if lcl_prompt_req_value <> "" then
           lcl_prompt_req_value = left(lcl_prompt_req_value,75) & "..."
           lcl_prompt_req_value = replace(lcl_prompt_req_value,"""","&quot;")
        end if

    		select case oQuestions("fieldtype")
       			case "2"
     			    'Build RADIO Question ----------------------------------------------
         		 	if sisRequired <> "" then
     		     		response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-radio/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
					end if

          			response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf

          			response.write "  <tr><td class=""question"">" & sisRequired & lcl_prompt_display & "</td></tr>" & vbcrlf

					if oQuestions("answerlist") <> "" then
             			arrAnswers = split(oQuestions("answerlist"),chr(10))

				response.write "<tr><td>"
				intNumAnswers = ubound(arrAnswers)
				if intNumAnswers > 7 or intNumAnswers = 2 or intNumAnswers = 5 then
					intDivAmt = intNumAnswers/3
				else
					intDivAmt = intNumAnswers/2
				end if
				intCount = 0
				response.write "<table><tr><td valign=""top"">"
             			for alist = 0 to ubound(arrAnswers)
					intCount = intCount + 1
							if arrAnswers(alist) <> "" then
								lcl_answerslist = arrAnswers(alist)
							else
								lcl_answerslist = ""
							end if

							'Format the value
							if lcl_answerslist <> "" then
								lcl_answerslist = replace(lcl_answerslist,"""","&quot;")
							end if

							response.write "  <label class=""actionradio"" style=""white-space:nobr;""><input type=""radio"" name=""fmquestion" & iQuestionCount & """ value=""" & lcl_answerslist & """ class=""formradio"" />" & arrAnswers(alist) & "</label><br />" & vbcrlf
					if intCount > intDivAmt and alist < intNumAnswers then
						intCount = 0
						response.write "</td><td valign=""top"">"
					end if
             			next
				response.write "</td></tr></table>"
				response.write "</td></tr>"
                    end if

                    response.write "<tr style=""display: none"">" & vbcrlf
                    response.write "    <td><input type=""radio"" name=""fmquestion" & iQuestionCount & """ value=""default_novalue"" CHECKED /></td>" & vbcrlf
                    response.write "</tr>" & vbcrlf
          			response.write "<tr><td>&nbsp;</td></tr>" & vbcrlf

       		case "4"
                'Build SELECT Question --------------------------------------------
                if sisRequired <> "" then
             			response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-select/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
                end if

          			response.write "<input type=""hidden"" name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_display & """ />" & vbcrlf

          			response.write "  <tr><td class=""question"">" & sisRequired & lcl_prompt_display & "</td></tr>" & vbcrlf

          			arrAnswers = split(oQuestions("answerlist"),chr(10))
			
          			response.write "  <tr>" & vbcrlf
                    response.write "      <td>" & vbcrlf
                    response.write "          <select class=""formselect"" name=""fmquestion" & iQuestionCount & """ />" & vbcrlf

          			for alist = 0 to ubound(arrAnswers)
             				response.write "            <option value=""" & formatSelectOptionValue(arrAnswers(alist)) & """>" & arrAnswers(alist) & "</option>" & vbcrlf
          			next

                    response.write "          </select>" & vbcrlf
                    response.write "      </td>" & vbcrlf
                    response.write "  </tr>" & vbcrlf
          			response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

       			case "6"
            'Build CHECKBOX Question ------------------------------------------
       						if sisRequired <> "" then
       			  				response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-checkbox/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
       						end if

       						response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf

       						response.write "  <tr><td class=""question"">" & sisRequired & lcl_prompt_display & "</td></tr>" & vbcrlf

						response.write "<tr><td>"
       						arrAnswers = split(oQuestions("answerlist"),chr(10))
				intNumAnswers = ubound(arrAnswers)
				if intNumAnswers > 7 or intNumAnswers = 2 or intNumAnswers = 5 then
					intDivAmt = intNumAnswers/3
				else
					intDivAmt = intNumAnswers/2
				end if
				intCount = 0
				response.write "<table><tr><td valign=""top"">"
			
       			   i = 0
       						for alist = 0 to ubound(arrAnswers)
       			       i = i + 1
			       intCount = intCount + 1

                'Format the value
                 lcl_answerslist = arrAnswers(alist)

                 if lcl_answerslist <> "" then
                    lcl_answerslist = replace(lcl_answerslist,"""","&quot;")
                 end if

       					   		response.write "  <label style=""white-space:nowrap;""><input type=""checkbox"" name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & "_" & i & """  value=""" & lcl_answerslist & """ class=""formcheckbox"" onclick=""validateCheckbox('fmquestion" & iQuestionCount & "')"" />" & trim(arrAnswers(alist)) & "</label><br />" & vbcrlf
					if intCount > intDivAmt and alist < intNumAnswers then
						intCount = 0
						response.write "</td><td valign=""top"">"
					end if
       						next
				response.write "</td></tr></table>"
						response.write "</td></tr>"
       			   i = i + 1

       			   response.write "  <tr style=""display: none"">" & vbcrlf
       			   response.write "      <td>" & vbcrlf
       			   response.write "          <input type=""checkbox"" name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & "_" & i & """ value=""default_novalue"" CHECKED onclick=""validateCheckbox('fmquestion" & iQuestionCount & "')"" />" & vbcrlf
       			   response.write "          <span id=""total_options_fmquestion" & iQuestionCount & """>" & i & "</span>" & vbcrlf
       			   response.write "      </td>" & vbcrlf
      				   response.write "  </tr>" & vbcrlf
       						response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

       			case "8"
            'Build TEXT Question ----------------------------------------------
					if sisRequired <> "" then
						response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-text/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
					end if

					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
					response.write "  <tr><td class=""question"">" & sisRequired & lcl_prompt_display & "</td></tr>" & vbcrlf
					if iFormID = "17890" and oQuestions("prompt") = "Start Date" then
						sFirstDate = FindNextRyeBusinessDay(DateAdd("d",7,Date()))
						sFirstDate = RIGHT("0" & Month(sFirstDate),2) & "/" & RIGHT("0" & Day(sFirstDate),2) & "/" & Year(sFirstDate)
						%><script>
						$(document).ready(function() {

    							// assuming the controls you want to attach the plugin to 
    							// have the "datepicker" class set
							$('#startdate').Zebra_DatePicker({
            							first_day_of_week: 0,
            							format: 'm/d/Y',
								show_clear_date: false,
            							show_select_today: false,
								disabled_dates: ['* * * 0,6'],
  								direction: ['<%=sFirstDate%>', false],
								onSelect: function () {
									var enddate = new Date($('#startdate').val());
									enddate.setDate(enddate.getDate() + 29);
									var yyyy = enddate.getFullYear().toString();
									var mm = (enddate.getMonth()+1).toString(); // getMonth() is zero-based
									var dd  = enddate.getDate().toString();
									$("#enddate").html(mm + "/" + dd + "/" + yyyy);
								}
								});
							
 							});
						</script><%
						response.write "<tr><td><input id=""startdate"" name=""fmquestion" & iQuestionCount & """ value=""" & sFirstDate & """ type=""text"" style=""width:90px;"" readonly maxlength=""" & lcl_text_field_length & """ /></td></tr>"& vbcrlf
						response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf
						response.write "<tr><td class=""question"">End Date</td></tr>"& vbcrlf
						response.write "<tr><td><div id=""enddate"">" & DateAdd("d",29,sFirstDate) & "</div></td></tr>"& vbcrlf
					elseif iFormID = "17890" and oQuestions("prompt") = "Address" then
						%><script>
						$(document).ready(function() {
							$("#issuelocation").autocomplete({
							source: function( request, response ) {
								//alert( request.term );
								$("#residentaddressid").val("0");
								$("#addressinsystemmsg").html("We do not consider this to be a valid address.");
								jQuery("#addressinsystemmsg").toggleClass("addressfound", false);
								jQuery("#addressinsystemmsg").toggleClass("addressnotfound", true);
								jQuery.ajax({
									url: "http://apidev.egovlink.com/api/ActionForm/GetAddressList?callback=?",
									type: "GET",
									dataType: "jsonp",
									contentType: "application/json",
									data: {
										_OrgId : 153,
										_MatchString: request.term,
										_MaxRows: 15
									},
									success: function( data ) {
										response( jQuery.map( data, function( item ) {
											$('#ui-id-1').css('display', 'block');
											return {
												label: item.streetaddress,
												value: item.streetaddress, 
												residentadressid: item.residentadressid
											}
										}));
									}
								});
							},
							minLength: 2,
							select: function( event, ui ) {
								$("#residentaddressid").val( ui.item ? ui.item.residentadressid : "0" );
								$("#addressinsystemmsg").html("We consider this to be a valid address.");
								jQuery("#addressinsystemmsg").toggleClass("addressfound", true);
								jQuery("#addressinsystemmsg").toggleClass("addressnotfound", false);
								//$("#residentaddressval").val($("#issuelocation").val());
								//alert($("#residentaddressval").val());
							},
							close: function( event, ui ) {
								//alert($("#issuelocation").val());
								$("#residentaddressval").val($("#issuelocation").val());
							},
							change: function( event, ui ) {
								if ($("#residentaddressval").val() != $("#issuelocation").val())
								{
									$("#residentaddressid").val("0");
									$("#addressinsystemmsg").html("We do not consider this to be a valid address.");
									jQuery("#addressinsystemmsg").toggleClass("addressfound", false);
									jQuery("#addressinsystemmsg").toggleClass("addressnotfound", true);
								}
							}
						});
						});
						</script><%
						response.write "<tr><td>"
							response.write "<input id=""issuelocation"" name=""fmquestion" & iQuestionCount & """ value="""" type=""text"" style=""width:300px;"" maxlength=""" & lcl_text_field_length & """ />"
							response.write "<div id=""addressinsystemmsg"">Start typing an address in the box above and then select one from the popup list if there is a match.</div>"
							response.write "<input type=""" & lcl_hidden & """ name=""ef:residentaddressid-text/nonzero-req"" value=""Address"">" & vbcrlf
							response.write "<input type=""hidden"" id=""residentaddressid"" value="""" />"
							response.write "<input type=""hidden"" id=""residentaddressval"" value="""" />"
						response.write "</td></tr>"& vbcrlf
					else
						response.write "  <tr><td><input name=""fmquestion" & iQuestionCount & """ value="""" type=""text"" style=""width:300px;"" maxlength=""100"" /></td></tr>" & vbcrlf
					end if
					response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

				 case "10"
		  'Build TEXTAREA Question ------------------------------------------
					if sisRequired <> "" then
						  response.write "<input type=""" & lcl_hidden & """ name=""ef:fmquestion" & iQuestionCount & "-textarea/req"" value=""" & lcl_prompt_req_value & """ />" & vbcrlf
					end if

					response.write "<input type=""" & lcl_hidden & """ name=""fmname" & iQuestionCount & """ value=""" & lcl_prompt_value & """ />" & vbcrlf
					response.write "  <tr><td class=""question"">" & sisRequired & lcl_prompt_display & "</td></tr>" & vbcrlf
					response.write "  <tr><td><textarea name=""fmquestion" & iQuestionCount & """ id=""fmquestion" & iQuestionCount & """ class=""formtextarea"" maxlength=""4000""></textarea></td></tr>" & vbcrlf 'onkeydown=""document.getElementById('control_field').value=this.value;"" onkeyup=""javascript:checkFieldLength(this.value,4000,'N',this)"" onblur=""javascript:checkFieldLength(this.value,4000,'N',this)""></textarea></td></tr>" & vbcrlf
					response.write "  <tr><td>&nbsp;</td></tr>" & vbcrlf

				case else

    		end select

		    oQuestions.movenext
    	wend

		response.write "</table>" & vbcrlf

	end if

	 oQuestions.close
	 set oQuestions = nothing

	'Determine if the form is diplaying the issue/problem location section
 	sSQL = "SELECT action_form_display_issue FROM egov_action_request_forms WHERE action_form_id = " & CLng(iFormID)

	 set oForm = Server.CreateObject("ADODB.Recordset")
 	oForm.Open sSQL, Application("DSN"), 3, 1
	
	 blnIssueDisplay = oForm("action_form_display_issue")

	 'Determine if we should be showing the file upload field
	 'response.write blnDisplayMobileOptions_takePic  & "####"
	 'if iorgid = "37" then response.write blnDisplayMobileOptions_geoLoc  & "####"
	 'response.write lcl_orghasfeature_actionline_formcreator_mobileoptions & "###"
	 if blnDisplayMobileOptions_takePic and lcl_orghasfeature_actionline_formcreator_mobileoptions then%>
		<fieldset>
		<legend><strong>Attach a Picture or File (max size: 30MB):</strong></legend>
		<input type="file" name="file1" />
                    <div id="morefiles"></div>
                    <a href="javascript:AddMoreFiles()">Add Another FIle</a>
		    <script>
                	var filecount = 1;
                	function AddMoreFiles()
                	{
                        	filecount = filecount + 1;
	
                        	var input=document.createElement('input');
                        	var linebreak = document.createElement("br");
                        	input.type="file";
                        	input.name="file" + filecount
                        	document.getElementById('morefiles').appendChild(input);
                        	document.getElementById('morefiles').appendChild(linebreak);
	
                	}
		    </script>
				
		</fieldset>
	 <%end if %>
	 <%if blnShowMapInput and lcl_orghasfeature_actionline_formcreator_mobileoptions then%>
		<fieldset>
		<legend><strong>Click the map where the issue is located.  Drag the map point if it needs to be moved.</strong></legend>
		<style>
			.mappointmap {
				height:250px;
				width:100%;
			}
		</style>
    		<div id="map" class="mappointmap"></div>
                <input type="hidden" id="mapLat" name="mapLat" value="" />
                <input type="hidden" id="mapLng" name="mapLng" value="" />
    		<script src="https://maps.googleapis.com/maps/api/js?key=AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM&callback=initMaps&libraries=&v=weekly" defer ></script>
    <script>
    var map, marker;
    var pointPlaced = false;
      (function(exports) {
        "use strict";

        function initMap() {
          map = new google.maps.Map(document.getElementById("map"), {
            zoom: <%=intZoom%>,
            center: {lat: <%=dblLat%>, lng: <%=dblLng%>}
          });
        // Try HTML5 geolocation.
        if (navigator.geolocation) {
          navigator.geolocation.getCurrentPosition(function(position) {
            var pos = {
              lat: position.coords.latitude,
              lng: position.coords.longitude
            };

            map.setCenter(pos);
          }, function() {
            //handleLocationError(true, infoWindow, map.getCenter());
          });
        } else {
          // Browser doesn't support Geolocation
          //handleLocationError(false, infoWindow, map.getCenter());
        }
          map.addListener("click", function(e) {
            placeMarkerAndPanTo(e.latLng, map);
          });
        }

        function placeMarkerAndPanTo(latLng, map) {
		if (!pointPlaced)
		{
          		marker = new google.maps.Marker({
            			position: latLng,
				draggable:true,
            			map: map
          			});
          		map.panTo(latLng);
			map.setZoom(12);
			pointPlaced = true;
        	}
	}

        exports.initMap = initMap;
        exports.placeMarkerAndPanTo = placeMarkerAndPanTo;
      })((this.window = this.window || {}));


    </script>
				
		</fieldset>
	 <%end if %>
		<br />
		<fieldset style="padding:0;">
		<legend><strong>CAPTCHA</strong></legend>
		<div class="g-recaptcha" style="width:305px;display:block;margin:0 auto;" data-sitekey="6LcVxxwUAAAAAEYHUr3XZt3fghgcbZOXS6PZflD-"></div>
		 </fieldset>
		<br />
	<%

	'SEND EMAIL
	'-------------------------------------------------------------------------
	'- This IF statement sets up the SUBMIT FORM button and the validation for the screen.
	'- The IF statement checks to see if the Contact Email is "turned on" and if so there is customized
	'validation that takes place specifically dealing with the Contact Email field.  Otherwise, the 
	'validation is not needed.
	'-------------------------------------------------------------------------
	if iFormID = "17890" then
		response.write "<input type=""hidden"" name=""chkSendEmail"" value=""YES"" /> " & vbcrlf
	elseif iFormID = "17968" then
		response.write "<input type=""hidden"" name=""chkSendEmail"" value=""NO"" /> " & vbcrlf
	elseif isDisplay(sMASK,4) then
		'-------------------------------------------------------------------------
		'If email is listed in personal information form then show option to send email to user
		response.write "<div class=""frmBottom"">" & vbcrlf
		response.write "<input type=""checkbox"" name=""chkSendEmail"" checked=""checked"" value=""YES"" /> " & vbcrlf
		response.write "Check here to have email confirmation of this submission." & vbcrlf
		response.write "</div>" & vbcrlf
	end if

	if blnIssueDisplay AND lcl_orghasfeature_issue_location then
		if lcl_orghasfeature_large_address_list then
			lcl_onclick = "checkAddress( 'FinalCheck', 'yes' );"
		else
			'lcl_onclick = "if (ValidateInput()) {isemailentered();}"
			lcl_onclick = "if(($('#streetaddress').val()=='0000')&&($('#ques_issue2').val()=='')){"
			lcl_onclick = lcl_onclick & "alert('Required field missing: Address');"
			lcl_onclick = lcl_onclick & "$('#streetaddress').focus();"
			lcl_onclick = lcl_onclick & "}else{"
			lcl_onclick = lcl_onclick &   "if(ValidateInput()) {"
			lcl_onclick = lcl_onclick &       "isemailentered();"
			'lcl_onclick = lcl_onclick &   "}else{"
			'lcl_onclick = lcl_onclick &       "document.getElementById('btnSubmit').disabled=false;"
			lcl_onclick = lcl_onclick &   "}"
			lcl_onclick = lcl_onclick & "}"
		end if
	else
		lcl_onclick = "if(ValidateInput()) {"
		lcl_onclick = lcl_onclick & "isemailentered();"
		'lcl_onclick = lcl_onclick & "}else{"
		'lcl_onclick = lcl_onclick & "document.getElementById('btnSubmit').disabled=false;"
		lcl_onclick = lcl_onclick & "}"
	end if

	'Submit button with email address check
	response.write "<div class=""frmBottom"">" & vbcrlf
	'response.write "<input name=""btnSubmit"" id=""btnSubmit"" class=""actionbtn"" type=""button"" onclick=""" & lcl_onclick & """ value=""SUBMIT FORM"" />" & vbcrlf
	if iFormID = "17890" then
 response.write "<input type=""button"" name=""btnSubmit"" id=""btnSubmit"" onclick=""" & lcl_onclick & """ value=""SUBMIT FORM AND PRINT REGISTRATION"" />" & vbcrlf
	else
 response.write "<input type=""button"" name=""btnSubmit"" id=""btnSubmit"" onclick=""" & lcl_onclick & """ value=""SUBMIT FORM"" />" & vbcrlf
 	end if
	response.write "</div>" & vbcrlf
 '--------------------------------------------------------------------------
'  			else  'Contact Email is NOT turned on
 '--------------------------------------------------------------------------
'        if blnIssueDisplay AND lcl_orghasfeature_issue_location then
'           if lcl_orghasfeature_large_address_list then
'              lcl_onclick = "checkAddress( 'FinalCheck', 'yes' );"
'           else
'              'lcl_onclick = "if (ValidateInput()) {document.frmRequestAction.submit();}"
'              lcl_onclick = "if(document.frmRequestAction.skip_address.value=='0000'&&document.frmRequestAction.ques_issue2.value==''){"
'              lcl_onclick = lcl_onclick & "alert('Required field missing: Address');"
'              lcl_onclick = lcl_onclick & "document.frmRequestAction.skip_address.focus();"
'              lcl_onclick = lcl_onclick & "}else{"
'              lcl_onclick = lcl_onclick &   "if (ValidateInput()) {"
'              lcl_onclick = lcl_onclick &       "isemailentered();"
'              lcl_onclick = lcl_onclick &   "}else{"
'              lcl_onclick = lcl_onclick &       "document.getElementById('btnSubmit').disabled=false;"
'              lcl_onclick = lcl_onclick &   "}"
'              lcl_onclick = lcl_onclick & "}"
'           end if
'        else
'           lcl_onclick = "if (ValidateInput()) {"
'           lcl_onclick = lcl_onclick & "isemailentered();"
'           lcl_onclick = lcl_onclick & "}else{"
'           lcl_onclick = lcl_onclick & "document.getElementById('btnSubmit').disabled=false;"
'           lcl_onclick = lcl_onclick & "}"
'        end if

    			'Submit button with NO email address check
'   	  		response.write "<div style=""width:450px;text-align:right;"">" & vbcrlf
'        response.write "<input name=""btnSubmit"" id=""btnSubmit"" class=""actionbtn"" type=""button"" onclick=""" & lcl_onclick & """  value=""SUBMIT FORM"" />" & vbcrlf
'        response.write "</div>" & vbcrlf
'  			end if

 '-----------------------------------------------------------------------------
'  end if  'End Contact Email check
 '-----------------------------------------------------------------------------

end sub

'------------------------------------------------------------------------------
Function isRequired(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "2" Then
		  sReturnValue = " <font color=""#ff0000"">*</font> "
	Else
		  sReturnValue = ""
	End If

	isRequired = sReturnValue
End Function

'------------------------------------------------------------------------------
Function isDisplay(sMASK,iField)
	sValue = Mid(sMask,iField,1)
	
	If sValue = "1" or sValue = "2" Then
		  sReturnValue = True
	Else
		  sReturnValue = False
	End If

	isDisplay = sReturnValue
End Function

'------------------------------------------------------------------------------
sub displayIssueLocation_new(iOrgID, _
                             iOrgHasFeature_issueLocation, _
                             iOrgHasFeature_largeAddressList, _
                             iIssueMask, _
                             iIssueQues, _
                             sHideIssueLocAddInfo)

  dim sOrgID

  sOrgID                  = 0
  lcl_street_number       = ""
  lcl_street_address      = ""
  lcl_issueMask           = ""
  lcl_issueQuestion       = ""
  lcl_isRequired_address  = ""
  lcl_isRequired_addinfo  = ""
  lcl_hideIssueLocAddInfo = false

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iIssueMask = "" or IsNull(iIssueMask) then
   		lcl_issueMask = "121111"
  end if

  if trim(iIssueQues) = "" OR isnull(trim(iIssueQues)) then
	  		lcl_issueQuestion = "Provide any additional information on problem location in the box below."
  else
     lcl_issueQuestion = iIssueQues
  end if

  if sHideIssueLocAddInfo <> "" then
     lcl_hideIssueLocAddInfo = sHideIssueLocAddInfo
  end if

  if isDisplay(lcl_issueMask,2) then
     lcl_isRequired_address = isRequired(lcl_issueMask,2)
     lcl_isRequired_address = lcl_isRequired_address & " "
  end if

  if isDisplay(lcl_issueMask,6) AND not lcl_hideIssueLocAddInfo then 
     lcl_isRequired_addinfo = isRequired(lcl_issueMask,6)
  end if

  if iOrgHasFeature_largeAddressList then
     GetAddressInfoLarge sOrgID, _
                         lcl_street_number, _
                         lcl_street_address, _
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
    	GetAddressInfo lcl_street_address, _
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

  response.write "<p>" & vbcrlf
  response.write "<table class=""tblResponsive"" border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td nowrap=""nowrap"">" & lcl_isRequired_address & "Address</td>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <fieldset id=""address_fieldset"" class=""address_fieldset"">" & vbcrlf

 'Determine how to pull the address info.
 '- Check to see if the org has the "issue location" feature on.
 '- If "yes" then check to see if the org has the "large address list" feature on.
  if iOrgHasFeature_IssueLocation then
     if sValidStreet = "Y" then
        if iOrgHasFeature_LargeAddressList then
           lcl_street_name = buildStreetAddress("", sPrefix, sAddress, sSuffix, sDirection)

           displayLargeAddressList_new sOrgID, _
                                       sNumber, _
                                       sPrefix, _
                                       sAddress, _
                                       sSuffix, _
                                       sDirection
        else
           displayAddress_new sOrgID, _
                              sNumber, _
                              sAddress
        end if

        lcl_display_other_address = ""
        lcl_displayAddress        = sNumber & " " & sAddress
     else
        if iOrgHasFeature_LargeAddressList then
           displayLargeAddressList_new sOrgID, _
                                       "", _
                                       "", _
                                       "", _
                                       "", _
                                       ""
        else
           displayAddress_new sOrgID, _
                              "", _
                              ""
        end if

        lcl_display_other_address = sNumber

        if lcl_display_other_address <> "" then
           lcl_display_other_address = lcl_display_other_address & " " & sAddress
        else
           lcl_display_other_address = sAddress
        end if

        lcl_displayAddress = lcl_display_other_address
     end if

     response.write "<br /> - Or Other Not Listed - <br /> " & vbcrlf
     lcl_address_onchange = ""

  else
     lcl_display_other_address = sAddress
     lcl_displayAddress        = sAddress
     lcl_address_onchange      = ""
  end if

  response.write "          <input type=""text"" name=""ques_issue2"" id=""ques_issue2"" class=""correctionstextbox inputResponsive"" size=""50"" maxlength=""75"" value=""" & lcl_display_other_address & """" & lcl_address_onchange & " />" & vbcrlf
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
     'response.write "      			<input type=""button"" name=""validpick"" id=""validpick"" value=""Use the valid address selected"" class=""button"" onclick=""doSelect();"" />" & vbcrlf
     'response.write "      			<input type=""button"" name=""invalidpick"" id=""invalidpick"" value=""Use the address I entered"" class=""button"" onclick=""doKeep();"" />" & vbcrlf
     'response.write "      			<input type=""button"" name=""cancelpick"" id=""cancelpick"" value=""Cancel"" class=""button"" onclick=""cancelPick();"" />" & vbcrlf
     response.write "      			<input type=""button"" name=""validpick"" id=""validpick"" value=""Use the valid address selected"" onclick=""doSelect();"" />" & vbcrlf
     response.write "      			<input type=""button"" name=""invalidpick"" id=""invalidpick"" value=""Use the address I entered"" onclick=""doKeep();"" />" & vbcrlf
     response.write "      			<input type=""button"" name=""cancelpick"" id=""cancelpick"" value=""Cancel"" onclick=""cancelPick();"" />" & vbcrlf
     'response.write "      		</form>" & vbcrlf
     response.write "    </fieldset>" & vbcrlf
  end if

  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf

 'Unit
  if isDisplay(lcl_issueMask,2) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td align=""right"">Unit:&nbsp;</td>" & vbcrlf
     response.write "      <td><input type=""text"" name=""streetunit"" id=""streetunit"" size=""8"" maxlength=""10"" /></td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

 'Additional Info
  if isDisplay(lcl_issueMask,6) AND not lcl_hideIssueLocAddInfo then 
     'response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  	 	response.write "  <tr>" & vbcrlf
     response.write "      <td>&nbsp;</td>" & vbcrlf
     response.write "      <td>" & lcl_isRequired_addinfo & lcl_issueQuestion & "<br />" & vbcrlf
   		response.write "          <textarea name=""ques_issue6"" id=""ques_issue6"" class=""formtextarea"" maxlength=""512""></textarea>" & vbcrlf ' onkeydown=""document.getElementById('control_field').value=this.value;"" onkeyup=""javascript:checkFieldLength(this.value,512,'N',this)""></textarea>" & vbcrlf
     response.write "      </td>" & vbcrlf
	    response.write "  </tr>" & vbcrlf

     if lcl_isRequired_addinfo <> "" then
  	 	   response.write "<input type=""" & lcl_hidden & """ name=""ef:ques_issue6-textarea/req"" value=""Issue\problem additional information field"" />" & vbcrlf
   	 end if
  end if

  response.write "</table>" & vbcrlf
  response.write "</p>" & vbcrlf
end sub

'------------------------------------------------------------------------------
sub subDisplayIssueLocation( sIssueMask, iStreetNumberInputType, iStreetAddressInputType, sIssueQues, sHideIssueLocAddInfo )
	
 if sIssueMask = "" or IsNull(sIssueMask) then
  		sIssueMask = "121111"
 end if

 if trim(sIssueQues) = "" OR isnull(trim(sIssueQues)) then
	 		sIssueQues = "Provide any additional information on problem location in the box below."
 end if

 response.write "<table border=""1"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf

 if isDisplay(sIssueMask,2) then
  		response.write "  <tr>" & vbcrlf
	  	response.write "      <td valign=""top"" align=""right"">" & isRequired(sIssueMask,2) & " Address: &nbsp; </td>" & vbcrlf
		  response.write "      <td>" & vbcrlf
                              fnDrawInputType "3", "1"
 	 	response.write "      </td>" & vbcrlf
	 	 response.write "  </tr>" & vbcrlf

    response.write "  <tr>" & vbcrlf
  		response.write "      <td align=""right"">Unit:&nbsp;</td>" & vbcrlf
		  response.write "      <td><input type=""text"" name=""streetunit"" size=""8"" maxlength=""10"" /></td>" & vbcrlf
  		response.write "  </tr>" & vbcrlf
 end if

 if isDisplay(sIssueMask,6) AND not sHideIssueLocAddInfo then
  		sZipRequired = isRequired(sIssueMask,6)

    response.write "  <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf

   'Unit
 	 	response.write "  <tr>" & vbcrlf
    response.write "      <td>&nbsp;</td>" & vbcrlf
    response.write "      <td>" & isRequired(sIssueMask,6) & sIssueQues & "<br />" & vbcrlf
    'response.write "          <textarea name=""ques_issue6"" class=""formtextarea""></textarea></td>" & vbcrlf
  		response.write "          <textarea name=""ques_issue6"" id=""ques_issue6"" class=""formtextarea"" maxlength=""512""></textarea>" & vbcrlf ' onkeydown=""document.getElementById('control_field').value=this.value;"" onkeyup=""javascript:checkFieldLength(this.value,512,'N',this)""></textarea>" & vbcrlf
    response.write "      </td>" & vbcrlf
	   response.write "  </tr>" & vbcrlf

 		'IF REQUIRED ADD JAVASCRIPT CHECK
    if sZipRequired <> "" then
		 	   response.write "<input type=""" & lcl_hidden & """ name=""ef:ques_issue6-textarea/req"" value=""Issue\problem additional information field"" />" & vbcrlf
  	 end if
 end if
	
 response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sPrefix, ByRef sAddress, _
                    ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, ByRef sZip, _
                    ByRef sValidStreet, ByRef sLatitude, ByRef sLongitude )

 dim iResidentAddressID

 sValidStreet       = "N"
 iResidentAddressID = 0

 if sResidentAddressID <> "" then
    iResidentAddressID = clng(sResidentAddressID)
 end if

 sNumber           = ""
 sPrefix           = ""
 sAddress          = ""
 sSuffix           = ""
 sDirection        = ""
	sLatitude         = ""
 sLongitude        = ""
 sCity             = ""
 sState            = ""
 sZip              = ""
 sValidStreet      = "N"
 sLatitude         = ""
 sLongitude        = ""

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
	sSQL = sSQL & " WHERE residentaddressid = " & iResidentAddressID

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
Sub fnDrawInputType(iInputType,iAddressPart)

'	sReturnValue = DisplayAddress( iorgid, "R", 1 ) & " <br> - Or Other Not Listed - <br> "
'	sReturnValue = sReturnValue & "<input name=""ques_issue2"" type=""text"" />"


Select Case iInputType
 	Case "1"
     		'TEXT BOX ONLY
     			If iAddressPart = 2 Then
      				'response.write "<input name=""ques_issue1"" type=text>"
     			Else
        			response.write "<input name=""ques_issue2"" id=""ques_issue2"" type=""text"" size=""60"" maxlength=""75"" />" & vbcrlf
     			End If

  Case "2"
    			'SELECT BOX ONLY
      		If clng(iAddressPart) = 1 Then
       				DisplayAddress iorgid, "R"
    			'Else
    			'	Call DisplayAddressNumber( iorgid, "R" , 0 )
     			End If

  Case "3"
    			'TEXT OR SELECT BOX
     			If clng(iAddressPart) = 2 Then
       			'Call DisplayAddressNumber( iorgid, "R", 1 ) 
      				'response.write " <br> - Or Other Not Listed - <br> "

   		 	   response.write "<input type=""hidden"" name=""issuelocation-text/req"" value="""" />" & vbcrlf

     			Else
       				if lcl_orghasfeature_large_address_list then  'This is in include_top_functions.asp
         					DisplayLargeAddressList iorgid, "R" 

       		 	   response.write "<input type=""hidden"" name=""issuelocation-text/req"" value=""largeaddresslist"" />" & vbcrlf
       				Else
         					DisplayAddress iorgid, "R"

       		 	   response.write "<input type=""hidden"" name=""issuelocation-text/req"" value=""smalladdresslist"" />" & vbcrlf

         					If CityHasGeopointAddresses( iorgid, "R" ) Then 
           						response.write "&nbsp; <input type=""button"" class=""button"" name=""btnMap"" value=""View on Map"" onclick=""ShowMap();"" />" & vbcrlf
         					End If 
       				End If 
       				response.write "<br /> - Or Other Not Listed - <br />" & vbcrlf
    		 	End If
			
    			'If iAddressPart = 1 Then
    				response.write "<input name=""ques_issue2"" id=""ques_issue2"" type=""text"" size=""60"" maxlength=""75"" />" & vbcrlf
    			'Else
			'	response.write  "<input name=""ques_issue1"" type=text>"
			'End If

		Case Else
		   	 'DEFAULT TO TEXT BOX ONLY
     			If clng(iAddressPart) = 2 Then
      				'response.write "<input name=""ques_issue1"" type=text>"
     			Else
       				response.write "<input name=""ques_issue2"" id=""ques_issue2"" type=""text"" />" & vbcrlf
     			End If
End Select

End Sub 

'------------------------------------------------------------------------------
'Function CityHasGeopointAddresses( iOrgId, sResidenttype )
'	Dim sSql, oGeoPoints

'	CityHasGeopointAddresses = False 
'	sSQL = "SELECT count(latitude) as geopoints "
' sSQL = sSQL & " FROM egov_residentaddresses "
' sSQL = sSQL & " WHERE orgid = " & iOrgId
 'sSQL = sSQL & " AND residenttype='" & sResidenttype & "' "
' sSQL = sSQL & " AND excludefromactionline = 0 "
' sSQL = sSQL & " AND latitude is not NULL "

'	Set oGeoPoints = Server.CreateObject("ADODB.Recordset")
'	oGeoPoints.Open sSQL, Application("DSN"), 0, 1

'	If Not oGeoPoints.EOF Then 
'		If CLng(oGeoPoints("geopoints")) > CLng(0) Then
'			CityHasGeopointAddresses = True 
'		End If 
'	End If 
'	oGeoPoints.close
'	Set oGeoPoints = Nothing 

'End Function 

'------------------------------------------------------------------------------
function FormIsEnabled( iActionId )
  dim sSql, oForm

  sSQL = "SELECT isnull(action_form_enabled,0) AS action_form_enabled "
  sSQL = sSQL & " FROM egov_action_request_forms "
  sSQL = sSQL & " WHERE action_form_id = '" & iActionId  & "'"

  set oForm = Server.CreateObject("ADODB.Recordset")
  oForm.Open sSQL, Application("DSN"), 0, 1

  if not oForm.eof then
 	  	FormIsEnabled = oForm("action_form_enabled")
  else
   		FormIsEnabled = False
  end if

  oForm.close
  set oForm = nothing 

end function

'------------------------------------------------------------------------------
function isFormInternal( iActionId )
  dim sSql, oForm

  sSQL = "SELECT isnull(action_form_internal,0) AS action_form_internal "
  sSQL = sSQL & " FROM egov_action_request_forms "
  sSQL = sSQL & " WHERE action_form_id = '" & iActionId  & "'"

  set oForm = Server.CreateObject("ADODB.Recordset")
  oForm.Open sSQL, Application("DSN"), 0, 1

  if not oForm.eof then
 	  	isFormInternal = oForm("action_form_internal")
  else
   		isFormInternal = False
  end if

  oForm.close
  set oForm = nothing 

end function

'------------------------------------------------------------------------------
function formatSelectOptionValue(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,chr(10),"")
     lcl_return = replace(lcl_return,chr(13),"")
     lcl_return = replace(lcl_return,"<br>","")
     lcl_return = replace(lcl_return,"<br />","")
     lcl_return = replace(lcl_return,"<BR>","")
     lcl_return = replace(lcl_return,"<BR />","")
     lcl_return = replace(lcl_return,"""","&quot;")
  end if

  formatSelectOptionValue = lcl_return

end function

sub OtherReports()
	sSQL = "SELECT r.[OrgID] ,category_id, action_autoid ,[mobileoption_latitude] ,[mobileoption_longitude] ,f.action_form_name " _
  		& " FROM egov_actionline_requests r " _
  		& " INNER JOIN egov_action_request_forms f ON f.action_form_id = r.category_id " _
  		& " WHERE r.orgid = " & iorgid & " and r.category_id = '" & track_dbsafe(request("actionid")) & "' and mobileoption_latitude IS NOT NULL and mobileoption_latitude <> 0 " _
  		& " AND status IN ('SUBMITTED','IN PROGRESS','WAITING') "
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then %>
	<p><button id="otherrptsbtn" type="button" class="button" onclick="showMapModal();">Click here to see the other reports again</button></p>
	<div id="cover">
    		<div id="map_wrapper">
        		<div id="closemodal" class="closemodal" onclick="hideModal();"><img src="https://www.egovlink.com/permitcity/admin/images/close-icon.png" width="15" height="15" /></div>
			<div id="mapstuff" style="height:100%;width:100%;">
        			<h3>Take a look at other open reports to see if this has been reported already:</h3>
        			<div id="map_canvas" class="mapping"></div>
			</div>
			<div id="successMessage" style="display:none;"><img src="https://www.egovlink.com/eclink/images/success.gif" /><br />We've recorded your report in our system.<br /><a href="javascript:showMapModal();">Click Here</a> to make another report.</div>
		</div>
	</div>
	<script>
		<% ListCoordinates oRs %>
	</script>
	<script type="text/javascript" src="//www.egovlink.com/eclink/scripts/otherreports.js"></script>
	<!---script type="text/javascript" src="//maps.googleapis.com/maps/api/js?sensor=false&key=AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"></script-->
	<%
	end if

	oRs.Close
	Set oRs = Nothing
end sub 
sub ListCoordinates(oRs)
    	response.write "markers = [" & vbcrlf
	strJS = ""
	Do While Not oRs.EOF
       		strJS = strJS & "['', " & oRs("mobileoption_latitude") & "," & oRs("mobileoption_longitude") & "],"
		oRs.MoveNext
	loop
	response.write left(strJS,len(strJS)-1) & vbcrlf
   	response.write "];" & vbcrlf

	oRs.MoveFirst

	strJS = ""
    	response.write "infoWindowContent = [" & vbcrlf
	Do While Not oRs.EOF
       		strJS = strJS & "['<div class=""info_content"">' + " 
       		strJS = strJS & "'<h3>" & oRs("action_form_name") & "</h3>' + " 
       		strJS = strJS & "'<p>Click <a href=""javascript:upvote(\'" & oRs("action_autoid") & "\');"">HERE</a> to also report this issue.</p>' +        '</div>'], " 
		oRs.MoveNext
	loop
	response.write left(strJS,len(strJS)-1) & vbcrlf
	response.write "];" & vbcrlf
end sub

%>
