<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a Blog entry
'
' MODIFICATION HISTORY
' 1.0 03/15/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"mappoints_maint") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Retrieve the mappointid to be maintain.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("mappointid") <> "" then
    lcl_mappointid = request("mappointid")

    if isnumeric(lcl_mappointid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "mappoints_list.asp"
    end if
 else
    lcl_screen_mode     = "ADD"
    lcl_sendToLabel     = "Create"
    lcl_mappointid = 0
 end if

'Retrieve the search options
 lcl_sc_mappoint_typeid = ""

 if request("sc_mappoint_typeid") <> "" then
    lcl_sc_mappoint_typeid = request("sc_mappoint_typeid")
 end if

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_mappoint_typeid", lcl_sc_mappoint_typeid)

'Check for org features
 lcl_orghasfeature_issue_location     = orghasfeature("issue location")
 lcl_orghasfeature_large_address_list = orghasfeature("large address list")

'Check for user permissions
' lcl_userhaspermission_rssfeeds_mayorsblog = userhaspermission(session("userid"),"rssfeeds_mayorsblog")

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

'Set up local variables
 lcl_mappoint_typeid    = 0
 lcl_orgid              = session("orgid")
 lcl_createdbyid        = 0
 lcl_createdbydate      = ""
 lcl_createdbyname      = ""
 lcl_lastmodifiedbyid   = 0
 lcl_lastmodifiedbydate = ""
 lcl_lastmodifiedbyname = ""
 lcl_isInactive         = 0
 lcl_checked_isInactive = " checked=""checked"""
 lcl_number             = ""
 lcl_prefix             = ""
 lcl_address            = ""
 lcl_suffix             = ""
 lcl_direction          = ""
 lcl_city               = ""
 lcl_state              = ""
 lcl_zip                = ""
 lcl_validstreet        = "N"
 lcl_sortstreetname     = ""
 'sLatitude              = ""
 'sLongitude             = ""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the blog
    sSQL = "SELECT mp.mappointid, mp.mappoint_typeid, mp.orgid, mp.createdbyid, mp.createdbydate, mp.isInactive, "
    sSQL = sSQL & " mp.lastmodifiedbyid, mp.lastmodifiedbydate, "
    sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
    sSQL = sSQL & " mp.streetnumber, mp.streetprefix, mp.streetaddress, mp.streetsuffix, mp.streetdirection, mp.sortstreetname, "
    sSQL = sSQL & " mp.city, mp.state, mp.zip, mp.validstreet "
    'sSQL = sSQL & " mp.latitude, mp.longitude "
    sSQL = sSQL & " FROM egov_mappoints mp "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON mp.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON mp.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE mp.mappointid = " & lcl_mappointid

    set oMapPoints = Server.CreateObject("ADODB.Recordset")
    oMapPoints.Open sSQL, Application("DSN"), 3, 1

    if not oMapPoints.eof then
       lcl_mappointid         = oMapPoints("mappointid")
       lcl_mappoint_typeid    = oMapPoints("mappoint_typeid")
       lcl_orgid              = oMapPoints("orgid")
       lcl_createdbyid        = oMapPoints("createdbyid")
       lcl_createdbydate      = oMapPoints("createdbydate")
       lcl_createdbyname      = oMapPoints("createdbyname")
       lcl_lastmodifiedbyid   = oMapPoints("lastmodifiedbyid")
       lcl_lastmodifiedbydate = oMapPoints("lastmodifiedbydate")
       lcl_lastmodifiedbyname = oMapPoints("lastmodifiedbyname")
       lcl_isInactive         = oMapPoints("isInactive")
       lcl_number             = oMapPoints("streetnumber")
       lcl_prefix             = oMapPoints("streetprefix")
       lcl_address            = oMapPoints("streetaddress")
       lcl_suffix             = oMapPoints("streetsuffix")
       lcl_direction          = oMapPoints("streetdirection")
       lcl_city               = oMapPoints("city")
       lcl_state              = oMapPoints("state")
       lcl_zip                = oMapPoints("zip")
       'sLatitude              = oMapPoints("latitude")
       'sLongitude             = oMapPoints("longitude")
       lcl_sortstreetname     = oMapPoints("sortstreetname")
       lcl_validstreet        = oMapPoints("validstreet")

      'Determine if the checkbox(es) are checked or not
       if oMapPoints("isInactive") then
          lcl_checked_isInactive = ""
       end if
    else

       lcl_add_params = setupUrlParameters(lcl_url_parameters, "success", "NE")

       response.redirect("mappoints_list.asp" & lcl_add_params)
    end if

    oMapPoints.close
    set oMapPoints = nothing
 end if

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
 end if

'If the "large address" feature is turned on then enable/disable the "Import Address Fields" button
 if lcl_orghasfeature_large_address_list then
    lcl_onload = lcl_onload & "checkImportAddressBtn();"
 end if

'Show/Hide all "hidden" fields.  (HIDDEN = hide, TEXT = show)
 lcl_hidden = "text"

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {Map-Points - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
var control_field = "";

function confirmDelete() {
  //var r = confirm('Are you sure you want to delete the "' + document.getElementById("title").value + '" blog entry?  \r NOTE: Any/All comments will be deleted as well.');
//  var r = confirm('Are you sure you want to delete: "' + document.getElementById("description").value + '"');
  var r = confirm('Are you sure you want to delete this Map-Point?');
  if (r==true) {

    <%
      lcl_delete_params = lcl_url_parameters
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "user_action", "DELETE")
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "mappointid", lcl_mappointid)
    %>
      location.href="mappoints_action.asp<%=lcl_delete_params%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count = 0;

//  if(document.getElementById("description").value=="") {
//     inlineMsg(document.getElementById("description").id,'<strong>Required Field Missing: </strong> Description',10,'description');
//     lcl_focus       = document.getElementById("description");
//     lcl_false_count = lcl_false_count + 1;
//  }else{
//     clearMsg("description");
//  }

  //---------------------------------------------------------------------------
  //Check the Map-Point Category Fields
  //---------------------------------------------------------------------------
//  lcl_total_fields = document.getElementById("totalFields").value;
//  lcl_i_start         = 1;

//  for (i=lcl_total_fields; lcl_i_start<=i; -- i) {
//       if(document.getElementById("fieldname"+i).value == "") {
          //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
//          if(document.getElementById("deleteField"+i).checked!=true && (document.getElementById("fieldname"+i).value != "" || document.getElementById("displayInResults"+i).value != "" || document.getElementById("resultsOrder"+i).value != "")) {
//         			 inlineMsg(document.getElementById("fieldname"+i).id,'<strong>Required Field Missing: </strong>Field Name',8,'fieldname'+i);
//             lcl_false_count = lcl_false_count + 1;
//          }else{
//             clearMsg('fieldname'+i);
//             document.getElementById("fieldname"+i).value = '';
//          }

//          if(lcl_false_count == 1) {
//             lcl_focus = document.getElementById("fieldname"+i);
//          }
//       }else{
//          clearMsg('fieldname'+i);
//     		}
//  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("mappoints_maint").submit();
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

function checkImportAddressBtn() {
  <%
    if lcl_mappointid <> 0 then
       response.write "document.getElementById(""importAddress"").disabled=true;" & vbcrlf

      'If a non-valid address is entered then disable the "Import Address" button.
      'Determine which fields are valid depending on the which address list the org has
       lcl_address_js_code = ""

       if lcl_orghasfeature_large_address_list then
          lcl_address_js_code = "document.getElementById(""residentstreetnumber"").value!="""" || "
          lcl_address_js_code = lcl_address_js_code & "document.getElementById(""streetaddress"").value!=""0000"""
       else
          lcl_address_js_code = "document.getElementById(""streetaddress"").value!=""0000"""
       end if

       response.write "if(" & lcl_address_js_code & ") {" & vbcrlf
       response.write "   document.getElementById(""importAddress"").disabled=false;" & vbcrlf
       response.write "}" & vbcrlf
    end if
  %>
}

function getAddressFields() {
  document.getElementById("importAddressFields").value = "Y";
  document.getElementById("mappoints_maint").action = "mappoints_maint.asp";
  document.getElementById("mappoints_maint").submit();
}

function validateAddress() {
  // Remove any extra spaces
  document.mappoints_maint.residentstreetnumber.value = removeSpaces(document.mappoints_maint.residentstreetnumber.value);

  // check the number for non-numeric values
  var rege = /^\d+$/;
  var Ok = rege.exec(document.mappoints_maint.residentstreetnumber.value);

		if ( ! Ok ) {
    		alert("The Resident Street Number cannot be blank and must be numeric.");
	 	   setfocus(document.mappoints_maint.residentstreetnumber);
   		 return false;
  }

  // check that they picked a street name
  if ( document.mappoints_maint.streetaddress.value == '0000') {
 	   	alert("Please select a street name from the list first.");
    		setfocus(document.mappoints_maint.streetaddress);
   	 	return false;
 	}

  return true;

}

function save_address() {
  // Check to see if the address is on file or if it is a custom address
  if (document.mappoints_maint.residentstreetnumber.value=="" && document.mappoints_maint.streetaddress.options[document.mappoints_maint.streetaddress.selectedIndex].value=="0000") {
 	  		// Submit form "as is"
  	 		document.mappoints_maint.validstreet.value = 'N';
  }else{
 	  		// If address is seleted from list then clear out the custom address field
  	 		document.mappoints_maint.ques_issue2.value = '';
  	 		document.mappoints_maint.validstreet.value = 'Y';
  }
}

  //Set up global variables
 	var winHandle;
 	var w = (screen.width - 640)/2;
 	var h = (screen.height - 450)/2;

<% if lcl_orghasfeature_large_address_list then %>
function checkAddress( sReturnFunction, sSave ) {
  // build url
  var lcl_url_params;
  lcl_url_params += 'stnumber=' + document.mappoints_maint.residentstreetnumber.value;
  lcl_url_params += '&stname=' + document.mappoints_maint.streetaddress.value;

  if(document.mappoints_maint.ques_issue2.value=="") {
     lcl_success = validateAddress();
     if(lcl_success) {
      		// This is here because window.open in the Ajax callback routine will not work
      		//winHandle = eval('window.open("../addresspicker.asp?saving=' + sSave + '&stnumber=' + document.mappoints_maint.residentstreetnumber.value + '&stname=' + document.mappoints_maint.streetaddress.value + '&sCheckType=' + sReturnFunction + '&formname=mappoints_maint", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
      		//self.focus();

      		// Fire off Ajax routine
      		doAjax('../action_line/checkaddress.asp', lcl_url_params, sReturnFunction, 'get', '0');
     }
  }else{
     if(document.mappoints_maint.residentstreetnumber.value!="" || document.mappoints_maint.streetaddress.value!="0000") {
        document.mappoints_maint.ques_issue2.value="";
      		//winHandle = eval('window.open("../addresspicker.asp?saving=' + sSave + '&stnumber=' + document.mappoints_maint.residentstreetnumber.value + '&stname=' + document.mappoints_maint.streetaddress.value + '&sCheckType=' + sReturnFunction + '&formname=mappoints_maint", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
      		doAjax('../action_line/checkaddress.asp', lcl_url_params, sReturnFunction, 'get', '0');
     }else{
      //if(sReturnFunction!="FinalCheck") {
      //   alert("A non-listed address has been entered.");
      //   document.mappoints_maint.validstreet.value = 'N';
      //}else{
      //   document.mappoints_maint.validstreet.value = 'Y';
           FinalCheck('NOT FOUND');
      //}
     }
  }
}

function CheckResults( sResults ) {
  // Process the Ajax CallBack when the validate address button is clicked
  if (sResults == 'FOUND CHECK') {
    		//if(winHandle != null && ! winHandle.closed) { 
   	  //			winHandle.close();
   			//}
	 	  	document.mappoints_maint.ques_issue2.value = '';
      document.mappoints_maint.validstreet.value = 'Y';
  		 	alert("This is a valid address in the system.");
  }else{
      document.mappoints_maint.validstreet.value = 'N';
 	  		//winHandle.focus();
      PopAStreetPicker('CheckResults','no');
  }
}

function FinalCheck( sResults ) {
  if (sResults == 'FOUND CHECK') {
    		//if(winHandle != null && ! winHandle.closed) { 
   	  //			winHandle.close();
   			//}
      document.mappoints_maint.validstreet.value = 'Y';
      document.mappoints_maint.submit();
  }else{
      if ((sResults == 'FOUND SELECT')||(sResults == 'FOUND KEEP')) {
     		    //if(winHandle != null && ! winHandle.closed) { 
        	  //			winHandle.close();
   			     //}

           if (sResults == 'FOUND SELECT') {
               document.mappoints_maint.validstreet.value = 'Y';
           }else{
               document.mappoints_maint.validstreet.value = 'N';
           }

           document.mappoints_maint.submit();
      }else{
           //document.mappoints_maint.validstreet.value = 'N';
         		//if(winHandle != null && ! winHandle.closed) { 
           //   winHandle.focus();
           //}else{
           //   document.mappoints_maint.submit();
           //}
           if(document.mappoints_maint.ques_issue2.value!="") {
              document.mappoints_maint.submit();
           } else {
              PopAStreetPicker('FinalCheck','no');
           }
      }
  }
}

function PopAStreetPicker( sReturnFunction, sSave )	{
		// pop up the address picker
		//winHandle = eval('window.open("addresspicker.asp?saving=' + sSave + '&stnumber=' + document.register.residentstreetnumber.value + '&stname=' + document.register.skip_address.value + '&sCheckType=' + sReturnFunction + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');

  // build url parameters
  var lcl_url = "";
  lcl_url += "../action_line/addresspicker.asp";
  lcl_url += "?saving="     + sSave;
  lcl_url += "&stnumber="   + document.mappoints_maint.residentstreetnumber.value;
  lcl_url += "&stname="     + document.mappoints_maint.streetaddress.value;
  lcl_url += "&sCheckType=" + sReturnFunction;
  lcl_url += "&formname=mappoints_maint";

  winHandle = eval('window.open("' + lcl_url + '", "_picker", "width=640,height=450,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}
<% end if %>

function openWin(p_url, p_width, p_height) {
  w = 600;
  h = 400;

  if((p_width!="")&&(p_width!=undefined)) {
      w = p_width;
  }

  if((p_height!="")&&(p_height!=undefined)) {
      h = p_height;
  }

  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_url = p_url;

  eval('window.open("' + lcl_url + '", "_picker", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
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

<div id="centercontent">
<table border="0" cellspacing="0" cellpadding="10" width="800" class="start">
  <form name="mappoints_maint" id="mappoints_maint" method="post" action="mappoints_action.asp">
    <input type="<%=lcl_hidden%>" name="mappointid" value="<%=lcl_mappointid%>" size="5" maxlength="5" />
    <input type="<%=lcl_hidden%>" name="screen_mode" value="<%=lcl_screen_mode%>" size="4" maxlength="4" />
    <input type="<%=lcl_hidden%>" name="user_action" id="user_action" value="" size="4" maxlength="20" />
    <input type="<%=lcl_hidden%>" name="orgid" value="<%=session("orgid")%>" size="4" maxlength="10" />
    <input type="<%=lcl_hidden%>" name="validstreet" id="validstreet" value="<%=lcl_validstreet%>" />
    <input type="<%=lcl_hidden%>" name="sortstreetname" id="sortstreetname" value="<%=lcl_sortstreetname%>" />
    <input type="<%=lcl_hidden%>" name="city" id="city" value="<%=lcl_city%>" />
    <input type="<%=lcl_hidden%>" name="state" id="state" value="<%=lcl_state%>" />
    <input type="<%=lcl_hidden%>" name="zip" id="zip" value="<%=lcl_zip%>" />
    <input type="<%=lcl_hidden%>" name="importAddressFields" id="importAddressFields" value="<%=lcl_importAddressFields%>" size="1" maxlength="1" />
    <input type="<%=lcl_hidden%>" name="sc_mappoint_typeid" id="sc_mappoint_typeid" value="<%=lcl_sc_mappoint_typeid%>" />
  <tr>
      <td>
          <font size="+1"><strong>Map-Points: <%=lcl_screen_mode%></strong></font><br />
          <input type="button" name="backButton" id="backButton" value="Back to List" class="button" onclick="location.href='mappoints_list.asp<%=lcl_url_parameters%>';" />
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
  <tr valign="top">
      <td colspan="2">
          <p>
          <% displayButtons "TOP", lcl_screen_mode %>
          <table border="0" cellspacing="0" cellpadding="3" class="tableadmin">
            <tr>
                <th align="left">Map-Point</th>
                <th align="right">
                    <input type="checkbox" name="isInactive" id="isInactive" value="Y"<%=lcl_checked_isInactive%> /> Active
                </th>
            </tr>
            <tr>
                <td nowrap="nowrap">Map-Point Category:</td>
                <td>
                    <%
                     'If the user is editing then display the Map-Point Type "description".  Do not allow the value to be changed.
                     'If the user is adding then display the dropdown list of Map-Point Types.
                      if lcl_screen_mode = "EDIT" then
                         lcl_displayMPT_description = getMapPointTypeDescription(lcl_mappoint_typeid)

                         response.write "<input type=""" & lcl_hidden & """ name=""mappoint_typeid"" id=""mappoint_typeid"" value=""" & lcl_mappoint_typeid & """ />" & vbcrlf
                         response.write "<span style=""color:#800000;"">" & lcl_displayMPT_description & "</span>" & vbcrlf
                      else
                         response.write "<select name=""mappoint_typeid"" id=""mappoint_typeid"">" & vbcrlf
                                           displayMapPointTypes session("orgid"), lcl_mappoint_typeid
                         response.write "</select>" & vbcrlf
                      end if
                    %>
                </td>
            </tr>
            <tr><td colspan="2">&nbsp;</td></tr>
            <%
              displayMapPointFields session("orgid"), lcl_mappointid, lcl_mappoint_typeid, lcl_orghasfeature_issue_location, lcl_orghasfeature_large_address_list, _
                                    lcl_importAddressFields, lcl_importstreet_number, lcl_importstreet_address, lcl_importsortstreetname, _
                                    lcl_number, lcl_prefix, lcl_address, lcl_suffix, lcl_direction, lcl_city, lcl_state, lcl_zip, lcl_sortstreetname, lcl_validstreet
  
              if lcl_screen_mode = "EDIT" then
                 response.write "<tr>" & vbcrlf
                 response.write "    <td nowrap=""nowrap"">Created By:</td>" & vbcrlf
                 response.write "    <td style=""color:#800000"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
                 response.write "</tr>" & vbcrlf
                 response.write "<tr>" & vbcrlf
                 response.write "    <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
                 response.write "    <td style=""color:#800000"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
                 response.write "</tr>" & vbcrlf
              else
                 response.write "<tr><td colspan=""2""></td></tr>" & vbcrlf
                 response.write "<tr><td colspan=""2""></td></tr>" & vbcrlf
              end if
            %>
          </table>
            <%
             'Retrieve any/all fields related to this Map-Point Type
              'displayMPTypesFields session("orgid"), lcl_mappoint_typeid

             'Display the bottom row of buttons
              displayButtons "BOTTOM", lcl_screen_mode
            %>
          </p>
      </td>
  </tr>
</table>
</div>

<%
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
sub displayMapPointFields(iOrgID, iMapPointID, iMapPointTypeID, iOrgHasFeature_IssueLocation, iOrgHasFeature_LargeAddressList, iImportAddressFields, _
                         iImportStreetNumber, iImportStreetName, iImportSortStreetName, iNumber, iPrefix, iAddress, iSuffix, iDirection, iCity, _
                         iState, iZip, iSortStreetName, iValidStreet)

  lcl_display_other_address = ""
  lcl_displayAddress        = ""
  iRowCount                 = 0

 'Retrieve all of the Map-Point Type Fields
  sSQL = "SELECT mpf.mp_fieldid, mpf.mappoint_typeid, mpf.fieldname, UPPER(mpf.fieldtype) as fieldtype, mpf.displayInResults, mpf.resultsOrder, "
  sSQL = sSQL & " mpv.fieldvalue "
  sSQL = sSQL & " FROM egov_mappoints_types_fields mpf "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_mappoints_values mpv ON mpf.mp_fieldid = mpv.mp_fieldid "

  if iMapPointID <> "" then
     sSQL = sSQL &   " AND mpv.mappointid = " & iMapPointID
  end if

  sSQL = sSQL & " WHERE mpf.mappoint_typeid = " & iMapPointTypeID
  sSQL = sSQL & " ORDER BY mpf.resultsOrder "

  set oMPTFields = Server.CreateObject("ADODB.Recordset")
  oMPTFields.Open sSQL, Application("DSN"), 3, 1

  if not oMPTFields.eof then

    'BEGIN: Import address fields ---------------------------------------------
     if iImportAddressFields = "Y" then
        GetAddressInfoNew iOrgHasFeature_LargeAddressList, iOrgID, iImportStreetNumber, iImportStreetName, sNumber, sPrefix, sAddress, sSuffix, sDirection, _
                          sLatitude, sLongitude, sCity, sState, sZip, sCounty, sParcelID, sListedOwner, sLegalDescription, sResidentType, _
                          sRegisteredUserID, sValidStreet

       'These fields are NOT to be overridden during the import
        sSortStreetName   = iImportSortStreetName
     end if
    'END: Import address fields -----------------------------------------------

     do while not oMPTFields.eof

        iRowCount = iRowCount + 1

        response.write "  <tr>" & vbcrlf
        response.write "      <td nowrap=""nowrap"">" & oMPTFields("fieldname") & ":</td>" & vbcrlf
        response.write "      <td nowrap=""nowrap"" valign=""top"">" & vbcrlf

       'Check for "specialty" fields
        if oMPTFields("fieldtype") = "ADDRESS" then
           response.write "<fieldset>" & vbcrlf

          'Check to see if the user wants to import the address fields because the street number/name has been changed.
           if iImportAddressFields <> "Y" then
              sNumber           = iNumber
              sPrefix           = iPrefix
              sAddress          = iAddress
              sSuffix           = iSuffix
              sDirection        = iDirection
              sCity             = iCity
              sState            = iState
              sZip              = iZip
              sSortStreetName   = iSortStreetName

             'If the org does NOT have the "issue location" feature turned on then all addresses entered are considered "Invalid"
              if iOrgHasFeature_IssueLocation then
                 sValidStreet = iValidStreet
              end if

              'sUnit             = oIssueLocation("streetunit")
              'sCounty           = oIssueLocation("county")
              'sParcelID         = oIssueLocation("parcelidnumber")
              'sListedOwner      = oIssueLocation("listedowner")
              'sLegalDescription = oIssueLocation("legaldescription")
              'sComments         = oIssueLocation("comments")
              'sResidentType     = oIssueLocation("residenttype")
              'sRegisterUserID   = oIssueLocation("registereduserid")
           end if

          'Determine how to pull the address info.
          '- Check to see if the org has the "issue location" feature on.
          '- If "yes" then check to see if the org has the "large address list" feature on.
           if iOrgHasFeature_IssueLocation then
              if iValidStreet = "Y" then
                 if iOrgHasFeature_LargeAddressList then
                    lcl_street_name = buildStreetAddress("", sPrefix, sAddress, sSuffix, sDirection)

                    DisplayLargeAddressList iOrgID, sNumber, sAddress
                 else
                    DisplayAddress iOrgID, sNumber, sAddress
                 end if

                 lcl_display_other_address = ""
                 lcl_displayAddress        = sNumber & " " & sAddress
              else
                 if iOrgHasFeature_LargeAddressList then
                    DisplayLargeAddressList iOrgID, "", ""
                 else
                    DisplayAddress iOrgID, "", ""
                 end if

                 lcl_display_other_address = sNumber

                 if lcl_display_other_address <> "" then
                    lcl_display_other_address = lcl_display_other_address & " " & sAddress
                 else
                    lcl_display_other_address = sAddress
                 end if

                 lcl_displayAddress = lcl_display_other_address
              end if

              if iMapPointID <> 0 then
                 response.write "<input type=""button"" id=""importAddress"" class=""button"" value=""Import Address Fields"" onclick=""getAddressFields()"" />" & vbcrlf
              end if

              response.write "<br /> - Or Other Not Listed - <br /> " & vbcrlf
           else
              lcl_display_other_address = sAddress
              lcl_displayAddress        = sAddress
           end if

           response.write "          <input type=""text"" name=""ques_issue2"" id=""ques_issue2"" class=""correctionstextbox"" size=""50"" maxlength=""75"" value=""" & lcl_display_other_address & """ onchange=""save_address();checkImportAddressBtn()"" />" & vbcrlf
           response.write "    <br /><input type=""" & lcl_hidden & """ name=""mp_fieldvalue" & iRowCount & """ id=""mp_fieldvalue" & iRowCount & """ value=""" & lcl_displayAddress & """ size=""50"" maxlength=""500"" />" & vbcrlf
           response.write "</fieldset>" & vbcrlf
           response.write "For Mapping, enter the latitude and longitude.<br />" & vbcrlf
           response.write """Valid"" addresses may auto-populate these values when available.<br />" & vbcrlf
           response.write "If not, you can search for latitude and longitude values <a href=""javascript:openWin('http://www.batchgeocode.com/lookup/', 1000, 600);"">here.</a>" & vbcrlf
           'response.write "<a href=""http://www.batchgeocode.com/lookup/"" target=""_blank"">here.</a>" & vbcrlf
        elseif oMPTFields("fieldtype") = "STATUS" then

          'Get the statusid for this Map-Point
           getMapPointStatusInfo iMapPointID, lcl_statusid, lcl_statusname

           response.write "<select name=""statusid"" id=""statusid"" onchange=""clearMsg('status');"">" & vbcrlf
                             displayMPTypeStatuses iOrgID, lcl_statusid, False
           response.write "</select>" & vbcrlf
           response.write "    <br /><input type=""" & lcl_hidden & """ name=""mp_fieldvalue" & iRowCount & """ id=""mp_fieldvalue" & iRowCount & """ value=""" & lcl_statusname & """ size=""50"" maxlength=""500"" />" & vbcrlf

        elseif oMPTFields("fieldtype") = "LATITUDE" OR oMPTFields("fieldtype") = "LONGITUDE" then

          'If this is an import then the LATITUDE and LONGITUDE fields must be overridden
           lcl_fieldvalue = oMPTFields("fieldvalue")

           if oMPTFields("fieldtype") = "LATITUDE" then
              lcl_fieldname = "latitude"

              if iImportAddressFields = "Y" then
                 lcl_fieldvalue = sLatitude
              'else
              '   lcl_fieldvalue = request("latitude")
              end if

           elseif oMPTFields("fieldtype") = "LONGITUDE" then
              lcl_fieldname = "longitude"

              if iImportAddressFields = "Y" then
                 lcl_fieldvalue = sLongitude
              'else
              '   lcl_fieldvalue = request("longitude")
              end if

           end if

           response.write "          <input type=""text"" name=""" & lcl_fieldname & """ id=""" & lcl_fieldname & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""500"" onchange=""clearMsg('" & lcl_fieldname & "');"" />" & vbcrlf
           response.write "          <input type=""" & lcl_hidden & """ name=""mp_fieldvalue" & iRowCount & """ id=""mp_fieldvalue" & iRowCount & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""500"" onchange=""clearMsg('mp_fieldvalue" & iRowCount & "');"" />" & vbcrlf

        else
           lcl_fieldvalue = oMPTFields("fieldvalue")

           response.write "          <input type=""text"" name=""mp_fieldvalue" & iRowCount & """ id=""mp_fieldvalue" & iRowCount & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""500"" onchange=""clearMsg('mp_fieldvalue" & iRowCount & "');"" />" & vbcrlf
        end if

        response.write "          <input type=""" & lcl_hidden & """ name=""mp_fieldid" & iRowCount & """ id=""mp_fieldid" & iRowCount & """ value=""" & oMPTFields("mp_fieldid") & """ size=""5"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""" & lcl_hidden & """ name=""fieldtype" & iRowCount & """ id=""fieldtype" & iRowCount & """ value=""" & oMPTFields("fieldtype") & """ size=""20"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""" & lcl_hidden & """ name=""fieldname" & iRowCount & """ id=""fieldname" & iRowCount & """ value=""" & oMPTFields("fieldname") & """ size=""20"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""" & lcl_hidden & """ name=""displayInResults" & iRowCount & """ id=""displayInResults" & iRowCount & """ value=""" & oMPTFields("displayInResults") & """ size=""5"" maxlength=""10"" />" & vbcrlf
        response.write "          <input type=""" & lcl_hidden & """ name=""resultsOrder" & iRowCount & """ id=""resultsOrder" & iRowCount & """ value=""" & oMPTFields("resultsOrder") & """ size=""5"" maxlength=""100"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oMPTFields.movenext
     loop
  end if

  oMPTFields.close
  set oMPTFields = nothing

  response.write "<input type=""" & lcl_hidden & """ name=""totalFields"" id=""totalFields"" value=""" & iRowCount & """ size=""5"" maxlength=""100"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode)

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

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='mappoints_types_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addAnotherButton"" id=""addAnotherButton"" value=""Add Another"" class=""button"" onclick=""return validateFields('ADDANOTHER');"" />" & vbcrlf
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""return validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub

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