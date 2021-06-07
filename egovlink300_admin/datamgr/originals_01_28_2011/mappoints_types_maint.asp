<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints_types_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a Blog entry
'
' MODIFICATION HISTORY
' 1.0 03/05/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'sSQL = "SELECT orgid, latitude, longitude, mappoints_defaultzoomlevel "
'sSQL = sSQL & " FROM organizations "
'sSQL = sSQL & " WHERE mappoints_defaultzoomlevel is not null or mappoints_defaultzoomlevel <> '' "

'set rs1 = Server.CreateObject("ADODB.Recordset")
'rs1.Open sSQL, Application("DSN"), 3, 1

'if not rs1.eof then
'   do while not rs1.eof

'      sSQL = "UPDATE egov_mappoints_types SET "
'      sSQL = sSQL & " latitude = "  & rs1("latitude")  & ", "
'      sSQL = sSQL & " longitude = " & rs1("longitude") & ", "
'      sSQL = sSQL & " mappoints_defaultzoomlevel = '" & rs1("mappoints_defaultzoomlevel") & "' "
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
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel               = "../"  'Override of value from common.asp
 lcl_onload           = ""
 lcl_isRootAdmin      = False
 lcl_isTemplate       = False
 lcl_isTemplate_url   = ""
 lcl_isTemplate_title = ""

'Determine if there is a specific feature associated to the Map-Point Types
'  "Is Limited": means that the admin user is maintaining a specific feature instead of the root admin viewing ALL MapPoint Types
'      (i.e. "Is Limited" will be true when an admin clicks on "Maintain Available Properties (fields)", but NOT 
'            when a root admin clicks on "Maintain MapPoint Types" and then selects a specific MapPoint Type to edit.
 lcl_mappoint_typeid = ""
 lcl_feature         = "mappoints_types_maint"
 lcl_isLimited       = False
 lcl_pagetitle       = "Map-Point Types"
 lcl_sectiontitle    = lcl_pagetitle & ": Fields"

 if request("f") <> "" AND request("f") <> "mappoints_types_maint" then
    lcl_feature      = request("f")
    lcl_isLimited    = True
    lcl_pagetitle    = getFeatureName(lcl_feature)
    lcl_sectiontitle = lcl_pagetitle

   'Retrieve the MapPoint_TypeID
    if request("mappoint_typeid") <> "" then
       lcl_mappoint_typeid = request("mappoint_typeid")
    else
       lcl_mappoint_typeid = getMapPointTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

       if lcl_mappoint_typeid = 0 then
         	response.redirect sLevel & "permissiondenied.asp"
       end if
    end if
 else
    if request("mappoint_typeid") <> "" then
       lcl_mappoint_typeid = request("mappoint_typeid")
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

'Retrieve the mappoint_typeid to be maintain.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 'if request("mappoint_typeid") <> "" then
 if lcl_mappoint_typeid <> "" then
    'lcl_mappoint_typeid = request("mappoint_typeid")

    if isnumeric(lcl_mappoint_typeid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    else
       response.redirect "mappoints_types_list.asp"
    end if
 else
    lcl_screen_mode     = "ADD"
    lcl_sendToLabel     = "Create"
    lcl_mappoint_typeid = 0
 end if

'Set up local variables
 lcl_description                = ""
 lcl_isActive                   = 1
 lcl_createdbyid                = 0
 lcl_createdbydate              = ""
 lcl_createdbyname              = ""
 lcl_lastmodifiedbyid           = 0
 lcl_lastmodifiedbydate         = ""
 lcl_lastmodifiedbyname         = ""
 lcl_mappointcolor              = "green"
 lcl_feature_public             = ""
 lcl_feature_maintain           = ""
 lcl_feature_maintain_fields    = ""
 lcl_displayMap                 = 1
 lcl_latitude                   = ""
 lcl_longitude                  = ""
 sLat                           = ""
 sLng                           = ""
 lcl_mapPoints_defaultZoomLevel = ""
 lcl_checked_isActive           = " checked=""checked"""
 lcl_checked_displayMap         = " checked=""checked"""
 lcl_checked_useAdvancedSearch  = ""

'Get the Latitude and Longitude for the org
 GetCityPoint session("orgid"), sLat, sLng

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the blog
    sSQL = "SELECT t.mappoint_typeid, "
    sSQL = sSQL & " t.description, "
    sSQL = sSQL & " t.isActive, "
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
    sSQL = sSQL & " t.displayMap, "
    sSQL = sSQL & " t.useAdvancedSearch, "
    sSQL = sSQL & " t.latitude, "
    sSQL = sSQL & " t.longitude, "
    sSQL = sSQL & " t.mappoints_defaultzoomlevel "
    sSQL = sSQL & " FROM egov_mappoints_types t "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON t.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON t.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE t.mappoint_typeid = " & lcl_mappoint_typeid

    set oMPTypes = Server.CreateObject("ADODB.Recordset")
    oMPTypes.Open sSQL, Application("DSN"), 3, 1

    if not oMPTypes.eof then
       lcl_description                = oMPTypes("description")
       lcl_isActive                   = oMPTypes("isActive")
       lcl_createdbyid                = oMPTypes("createdbyid")
       lcl_createdbydate              = oMPTypes("createdbydate")
       lcl_createdbyname              = oMPTypes("createdbyname")
       lcl_lastmodifiedbyid           = oMPTypes("lastmodifiedbyid")
       lcl_lastmodifiedbydate         = oMPTypes("lastmodifiedbydate")
       lcl_lastmodifiedbyname         = oMPTypes("lastmodifiedbyname")
       lcl_mappointcolor              = oMPTypes("mappointcolor")
       lcl_feature_public             = oMPTypes("feature_public")
       lcl_feature_maintain           = oMPTypes("feature_maintain")
       lcl_feature_maintain_fields    = oMPTypes("feature_maintain_fields")
       lcl_displayMap                 = oMPTypes("displayMap")
       lcl_useAdvancedSearch          = oMPTypes("useAdvancedSearch")
       lcl_latitude                   = oMPTypes("latitude")
       lcl_longitude                  = oMPTypes("longitude")
       lcl_mapPoints_defaultZoomLevel = oMPTypes("mappoints_defaultzoomlevel")

      'If the Map-Point Type is NOT "active" then do NOT "check" the checkbox
       if not lcl_isActive then
          lcl_checked_isActive = ""
       end if

       if not lcl_displayMap then
          lcl_checked_displayMap = ""
       end if

       if lcl_useAdvancedSearch then
          lcl_checked_useAdvancedSearch = " checked=""checked"""
       end if

    else
       response.redirect("mappoints_types_list.asp?success=NE")
    end if

    oMPTypes.close
    set oMPTypes = nothing

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
    if lcl_mapPoints_defaultZoomLevel = "" OR isnull(lcl_mapPoints_defaultZoomLevel) then
 		   	lcl_mapPoints_defaultZoomLevel = "13"
    end If

 end if

'Check for org features
' lcl_orghasfeature_rssfeeds_mayorsblog = orghasfeature("rssfeeds_mayorsblog")

'Check for user permissions
' lcl_userhaspermission_rssfeeds_mayorsblog = userhaspermission(session("userid"),"rssfeeds_mayorsblog")

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

'Check for associated Map-Points to determine if this MapPointType can/cannot be deleted.
'*** NOTE: if this IS a template then allow then bypass the check.
 if lcl_isTemplate then
    lcl_canDelete = true
 else
    lcl_canDelete = False
 end if

 if lcl_isRootAdmin AND not lcl_isTemplate then
    'lcl_canDelete = True
    lcl_canDelete = checkForMapPointsByMapPointTypeID(lcl_mappoint_typeid)

    lcl_onload = lcl_onload & "displayFeature('feature_public','displayfeature_public');"
    lcl_onload = lcl_onload & "displayFeature('feature_maintain','displayfeature_maintain');"
    lcl_onload = lcl_onload & "displayFeature('feature_maintain_fields','displayfeature_maintain_fields');"
 end if

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"
' lcl_onload  = lcl_onload & "enableDisableMapSetupFields();"

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

  <script type="text/javascript" src="../scripts/jquery-1.4.4.min.js"></script>


<style type="text/css">
  .hidden            { display: none; }
  .requiredFieldsMsg { color: #800000; }
</style>

<script language="javascript">

$(document).ready(function(){
  $('#requiredFieldsMsg').addClass('requiredFieldsMsg');

  $('#displayMap').click(function() {
    if($('#displayMap').attr('checked')) {
       $('#requiredFieldsMsg').toggleClass('requiredFieldsMsg');
       $('#requiredFieldsMsg').show('slow');
       $('#mappointcolor').attr('disabled',false);
       $('#latitude').attr('disabled',false);
       $('#longitude').attr('disabled',false);
       $('#getLatLongButton').attr('disabled',false);
       $('#mappoints_defaultzoomlevel').attr('disabled',false);
    } else {
       $('#requiredFieldsMsg').toggleClass('requiredFieldsMsg');
       $('#requiredFieldsMsg').hide('slow');
       $('#mappointcolor').attr('disabled',true);
       $('#latitude').attr('disabled',true);
       $('#longitude').attr('disabled',true);
       $('#getLatLongButton').attr('disabled',true);
       $('#mappoints_defaultzoomlevel').attr('disabled',true);
    }
  });
});


var control_field = "";

function enableDisableMapSetupFields() {
  lcl_isFieldChecked = false;

  if(document.getElementById("displayMap")) {
     lcl_isFieldChecked = document.getElementById("displayMap").checked;
  }

  document.getElementById("requiredFieldsMsg").style.display     = "none";
  document.getElementById("mappointcolor").disabled              = true;
  document.getElementById("latitude").disabled                   = true;
  document.getElementById("longitude").disabled                  = true;
  document.getElementById("mappoints_defaultzoomlevel").disabled = true;

  if(lcl_isFieldChecked) {
     document.getElementById("requiredFieldsMsg").style.display     = "inline";
     document.getElementById("mappointcolor").disabled              = false;
     document.getElementById("latitude").disabled                   = false;
     document.getElementById("longitude").disabled                  = false;
     document.getElementById("mappoints_defaultzoomlevel").disabled = false;

     if(document.getElementById("mappoints_defaultzoomlevel").value == "") {
        document.getElementById("mappoints_defaultzoomlevel").value = "13";
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
      location.href="mappoints_types_action.asp?user_action=DELETE&mappoint_typeid=<%=lcl_mappoint_typeid & lcl_isTemplate_url%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count    = 0;
  var lcl_isFieldChecked = false;

  //---------------------------------------------------------------------------
  //Check the Map-Point Types Fields
  //---------------------------------------------------------------------------
<% if lcl_screen_mode <> "ADD" then %>
  lcl_total_fields = document.getElementById("totalFields").value;
  lcl_i_start         = 1;

  for (i=lcl_total_fields; lcl_i_start<=i; -- i) {
       if(document.getElementById("fieldname"+i).value == "") {
          //Check to see if the other fields are NULL.  If so then allow the row to pass through validation and "check" the "Remove Flag"
          if(document.getElementById("deleteField"+i).checked!=true && (document.getElementById("fieldname"+i).value != "" || document.getElementById("displayInResults"+i).value != "" || document.getElementById("resultsOrder"+i).value != "")) {
         			 inlineMsg(document.getElementById("fieldname"+i).id,'<strong>Required Field Missing: </strong>Field Name',8,'fieldname'+i);
             lcl_false_count = lcl_false_count + 1;
          }else{
             clearMsg('fieldname'+i);
             document.getElementById("fieldname"+i).value = '';
          }

          if(lcl_false_count == 1) {
             lcl_focus = document.getElementById("fieldname"+i);
          }
       }else{
          clearMsg('fieldname'+i);
     		}
  }

  if(document.getElementById("displayMap")) {
     lcl_isFieldChecked = document.getElementById("displayMap").checked;
  }

  if(lcl_isFieldChecked) {
     if(document.getElementById("mappoints_defaultzoomlevel").value=="") {
        //inlineMsg(document.getElementById("mappoints_defaultzoomlevel").id,'<strong>Required Field Missing: </strong> Zoom Level',10,'mappoints_defaultzoomlevel');
        //lcl_focus       = document.getElementById("mappoints_defaultzoomlevel");
        //lcl_false_count = lcl_false_count + 1;
        document.getElementById("mappoints_defaultzoomlevel").value = "13";
     }else{
    				var rege = /^\d+$/;
				    var Ok = rege.exec(document.getElementById("mappoints_defaultzoomlevel").value);

    				if ( ! Ok ) {
           inlineMsg(document.getElementById("mappoints_defaultzoomlevel").id,'<strong>Invalid Value: </strong> Zoom Level must be numeric.',10,'mappoints_defaultzoomlevel');
           lcl_focus       = document.getElementById("mappoints_defaultzoomlevel");
           lcl_false_count = lcl_false_count + 1;
    			} else {
           clearMsg("mappoints_defaultzoomlevel");
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
     document.getElementById("mappoints_types_maint").submit();
     return true;
  }
}

function addFieldRow() {
  var mytbl     = document.getElementById('addFieldTBL');
  var totalrows = Number(document.getElementById("totalFields").value);

  //Increase the total rows by one.  This is index for the new row.
  totalrows = totalrows+1;

  //Set up the new row.
  var row = mytbl.insertRow(totalrows);
      row.id = "addFieldRow" + totalrows;

  //Set the background color.  Odd rows: "#eeeeee", Even rows: "#ffffff"
  var lcl_rowbg   = "";
  var lcl_evenodd = totalrows/2;
      lcl_evenodd = lcl_evenodd.toString();

  if(lcl_evenodd.indexOf('.') > 0) {
     lcl_rowbg = "#eeeeee";
  }else{
     lcl_rowbg = "#ffffff";
  }

  mytbl.style.background = lcl_rowbg;

  //Build the cells for the new row.
  var a = row.insertCell(0);  //Row Count
  var b = row.insertCell(1);  //Field Name
  var c = row.insertCell(2);  //Display In Results
  var d = row.insertCell(3);  //Display On Info Page
  var e = row.insertCell(4);  //Results Order
  var f = row.insertCell(5);  //In Public Search
  var g = row.insertCell(6);  //Include "Add a Link"
  var h = row.insertCell(7);  //Display as Multi-Line
  var i = row.insertCell(8);  //Delete Row and Additional Info

  //Row Count
  a.innerHTML = totalrows + '. ';

  //Field Name
  var cell_fieldname = document.createElement('input');
      cell_fieldname.type      = 'text';
      cell_fieldname.name      = 'fieldname' + totalrows;
      cell_fieldname.id        = 'fieldname' + totalrows;
      cell_fieldname.size      = '50';
      cell_fieldname.maxLength = '100';
      cell_fieldname.onchange  = function() { clearMsg('fieldname' + totalrows); };
  b.appendChild(cell_fieldname);
<%
  if lcl_isRootAdmin and not lcl_isLimited then
     response.write "  var cell_fieldtype = document.createElement('input');" & vbcrlf
     response.write "      cell_fieldtype.type      = 'text';" & vbcrlf
     response.write "      cell_fieldtype.name      = 'fieldtype' + totalrows;" & vbcrlf
     response.write "      cell_fieldtype.id        = 'fieldtype' + totalrows;" & vbcrlf
     response.write "      cell_fieldtype.size      = '15';" & vbcrlf
     response.write "      cell_fieldtype.maxLength = '100';" & vbcrlf

     response.write "  var cell_fieldtype_label1 = document.createElement('span');" & vbcrlf
     response.write "      cell_fieldtype_label1.innerHTML = '<br /><strong>Field Type: </strong>(code use ONLY)&nbsp;';" & vbcrlf
'     response.write "  var cell_fieldtype_label2 = document.createElement(""span"");" & vbcrlf
'     response.write "      cell_fieldtype_label2.innerHTML = "" (code use ONLY)&nbsp;"";" & vbcrlf

     response.write "  b.align='right';" & vbcrlf
     response.write "  b.appendChild(cell_fieldtype_label1);" & vbcrlf
     response.write "  b.appendChild(cell_fieldtype);" & vbcrlf
'     response.write "  b.appendChild(cell_fieldtype_label2);" & vbcrlf
  end if
%>
  //Display In Results
  var cell_displayresults = document.createElement('input');
      cell_displayresults.type      = 'checkbox';
      cell_displayresults.name      = 'displayInResults' + totalrows;
      cell_displayresults.id        = 'displayInResults' + totalrows;
      cell_displayresults.value     = '1';
      cell_displayresults.checked   = 'checked';
  c.align = 'center';
  c.appendChild(cell_displayresults);

  //Display On Info Page
  var cell_displayonpage = document.createElement('input');
      cell_displayonpage.type      = 'checkbox';
      cell_displayonpage.name      = 'displayInInfoPage' + totalrows;
      cell_displayonpage.id        = 'displayInInfoPage' + totalrows;
      cell_displayonpage.value     = '1';
      cell_displayonpage.checked   = 'checked';
  d.align = 'center';
  d.appendChild(cell_displayonpage);

  //Results Order
  var cell_resultsorder = document.createElement('input');
      cell_resultsorder.type      = 'text';
      cell_resultsorder.name      = 'resultsOrder' + totalrows;
      cell_resultsorder.id        = 'resultsOrder' + totalrows;
      cell_resultsorder.size      = '3';
      cell_resultsorder.maxLength = '5';
      cell_resultsorder.value     = totalrows;
  e.align = 'center';
  e.appendChild(cell_resultsorder);

  //In Public Search
  var cell_publicsearch = document.createElement('input');
      cell_publicsearch.type      = 'checkbox';
      cell_publicsearch.name      = 'inPublicSearch' + totalrows;
      cell_publicsearch.id        = 'inPublicSearch' + totalrows;
      cell_publicsearch.value     = '1';
      cell_publicsearch.checked   = '';
  f.align = 'center';
  f.appendChild(cell_publicsearch);

  //Include "Add a Link"
  var cell_addalink = document.createElement('input');
      cell_addalink.type      = 'checkbox';
      cell_addalink.name      = 'hasAddLinkButton' + totalrows;
      cell_addalink.id        = 'hasAddLinkButton' + totalrows;
      cell_addalink.value     = '1';
      cell_addalink.checked   = '';
  g.align = 'center';
  g.appendChild(cell_addalink);

  //Display as Multi-Line
  var cell_multiline = document.createElement('input');
      cell_multiline.type      = 'checkbox';
      cell_multiline.name      = 'isMultiLine' + totalrows;
      cell_multiline.id        = 'isMultiLine' + totalrows;
      cell_multiline.value     = '1';
      cell_multiline.checked   = '';
  h.align = 'center';
  h.appendChild(cell_multiline);

  //Delete Row and Additional Info
  var cell_deleterow = document.createElement('input');
      cell_deleterow.type      = 'checkbox';
      cell_deleterow.name      = 'deleteField' + totalrows;
      cell_deleterow.id        = 'deleteField' + totalrows;
      cell_deleterow.value     = 'Y';
      cell_deleterow.checked   = '';

  var cell_deleterow2 = document.createElement('input');
      cell_deleterow2.type      = 'hidden';
      cell_deleterow2.name      = 'mp_fieldid' + totalrows;
      cell_deleterow2.id        = 'mp_fieldid' + totalrows;
      cell_deleterow2.size      = '5';
      cell_deleterow2.maxLength = '';

  i.align = 'center';
  i.appendChild(cell_deleterow);
  i.appendChild(cell_deleterow2);

  //update the total row count.
  document.getElementById('totalFields').value = totalrows;
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
     lcl_iframe_url  = 'getMapPointTypeTemplateFields.asp';
     lcl_iframe_url += '?mappoint_typeid=' + encodeURIComponent(lcl_mptid);
     lcl_iframe_url += '&isRootAdmin='     + encodeURIComponent(lcl_isRootAdmin);
     lcl_iframe_url += '&isLimited='       + encodeURIComponent(lcl_isLimited);
     lcl_iframe_url += '&isTemplate='      + encodeURIComponent(lcl_isTemplate);
     lcl_iframe_url += '&isDisplayOnly='   + encodeURIComponent(lcl_isDisplayOnly);

     lcl_iframe_width  = "760";
     lcl_iframe_height = "300";
  }

  document.getElementById("previewMPTemplateFields").width  = lcl_iframe_width;
  document.getElementById("previewMPTemplateFields").height = lcl_iframe_height;
  document.getElementById("previewMPTemplateFields").src    = lcl_iframe_url;
}

function displayTemplateFields(p_code) {
  document.getElementById("previewMPTemplateFields").innerHTML = p_code;
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
  response.write "<form name=""mappoints_types_maint"" id=""mappoints_types_maint"" method=""post"" action=""mappoints_types_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""mappoint_typeid"" value=""" & lcl_mappoint_typeid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""screen_mode"" value=""" & lcl_screen_mode & """ size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""user_action"" id=""user_action"" value="" size=""4"" maxlength=""20"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""t"" id=""t"" value=""" & request("t") & """ size=""5"" maxlength=""5"" />" & vbcrlf

  if not lcl_isRootAdmin then
     response.write "  <input type=""hidden"" name=""feature_public"" id=""feature_public"" value=""" & lcl_feature_public & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain"" id=""feature_maintain"" value=""" & lcl_feature_maintain & """ />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain_fields"" id=""feature_maintain_fields"" value=""" & lcl_feature_maintain_fields & """ />" & vbcrlf
  end if

  if lcl_isTemplate then
     response.write "  <input type=""hidden"" name=""isTemplate"" id=""isTemplate"" value=""Y"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""useAdvancedSearch"" id=""useAdvancedSearch"" value=""Y"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""displayMap"" id=""displayMap"" value=""Y"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""mappointcolor"" id=""mappointcolor"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""latitude"" id=""latitude"" value="""" size=""15"" maxlength=""10"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""longitude"" id=""longitude"" value="""" size=""15"" maxlength=""10"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""mappoints_defaultzoomlevel"" id=""mappoints_defaultzoomlevel"" value="""" size=""10"" maxlength=""10"" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_public"" id=""feature_public"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain"" id=""feature_maintain"" value="""" />" & vbcrlf
     response.write "  <input type=""hidden"" name=""feature_maintain_fields"" id=""feature_maintain_fields"" value="""" />" & vbcrlf
  else
     response.write "  <input type=""hidden"" name=""isTemplate"" id=""isTemplate"" value=""N"" />" & vbcrlf
  end if

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_pagetitle & ": " & lcl_screen_mode & lcl_isTemplate_title & "</strong></font><br />" & vbcrlf

  if not lcl_isLimited then
     response.write "<input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='mappoints_types_list.asp" & replace(lcl_isTemplate_url,"&","?") & "';"" />" & vbcrlf
  end if

  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf
                            displayButtons "TOP", lcl_screen_mode, lcl_canDelete, lcl_isLimited, lcl_isTemplate
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
     response.write "                <td></td>" & vbcrlf
     response.write "                <td><input type=""checkbox"" name=""useAdvancedSearch"" id=""useAdvancedSearch"" value=""Y""" & lcl_checked_useAdvancedSearch & "/>&nbsp;Use Advanced Search (public-side enabled)</td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""2"">" & vbcrlf
     response.write "                    <p>" & vbcrlf
     response.write "                    <fieldset>" & vbcrlf
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
        response.write "                                  <strong>ADDRESS, LATITUDE, and LONGITUDE</strong>.  Without these values <br />" & vbcrlf
        response.write "                                  the map will NOT function properly." & vbcrlf
     end if

     response.write "                                </div>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                        <tr>" & vbcrlf
     response.write "                            <td nowrap=""nowrap"">MapPoint Color:</td>" & vbcrlf
     response.write "                            <td colspan=""3"">" & vbcrlf
     response.write "                                <select name=""mappointcolor"" id=""mappointcolor"">" & vbcrlf
                                                       displayMapPointColors lcl_mappointcolor
     response.write "                                </select>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
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
     response.write "			                      			    <input type=""text"" name=""mappoints_defaultzoomlevel"" id=""mappoints_defaultzoomlevel"" value=""" & lcl_mapPoints_defaultZoomLevel & """ size=""10"" maxlength=""10"" onchange=""clearMsg('mappoints_defaultzoomlevel');"" />" & vbcrlf
     response.write "			                      			    <span style=""color:#800000"">Zoom Levels: 0 to 21+ || Max Zoom OUT: 0 || Default Zoom: 13</span>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                      </table>" & vbcrlf
     response.write "                    </fieldset>" & vbcrlf
     response.write "                    </p>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if

  if lcl_screen_mode = "EDIT" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td nowrap=""nowrap"" style=""height:15px"">Created By:</td>" & vbcrlf
     response.write "      <td style=""color:#800000"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
     response.write "      <td style=""color:#800000"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  'else
  '   response.write "  <tr><td colspan=""2""></td></tr>" & vbcrlf
  '   response.write "  <tr><td colspan=""2""></td></tr>" & vbcrlf
  end if

  if lcl_isRootAdmin AND not lcl_isTemplate then
     response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
     response.write "            <tr valign=""top"">" & vbcrlf
     response.write "                <td nowrap=""nowrap"">Feature<br />(public-side):</td>" & vbcrlf
     response.write "                <td>" & vbcrlf
     response.write "                    <select name=""feature_public"" id=""feature_public"" onchange=""displayFeature('feature_public','displayfeature_public');"">" & vbcrlf
     response.write "                      <option value=""""></option>" & vbcrlf
                                           showFeatureOptions lcl_feature_public
     response.write "                    </select>" & vbcrlf
     response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_public"" id=""helpFeature_public"" alt=""Used to connect the Map-Points public displayed feature (home page option - i.e. Available Properties feature) to this Map-Point Type."" /><br />" & vbcrlf
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
     response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_maintain"" id=""helpFeature_maintain"" alt=""Used to determine which feature/permission to use when accessing the maintenance screen."" /><br />" & vbcrlf
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
     response.write "                    <img src=""../images/help.jpg"" name=""helpFeature_maintain_fields"" id=""helpFeature_maintain_fields"" alt=""Used to determine which feature/permission to use when accessing the Map-Point Types maintenance screen for admins.  It will limit the screen to this specific Map-Point Type from the navigation menu."" /><br />" & vbcrlf
     response.write "                    <span id=""displayfeature_maintain_fields"" style=""color:#800000;""></span>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf

     if lcl_screen_mode = "ADD" then
        response.write "            <tr valign=""top"">" & vbcrlf
        response.write "                <td colspan=""2"">" & vbcrlf
        response.write "                    <fieldset>" & vbcrlf
        response.write "                      <legend>Field Setup&nbsp;</legend>" & vbcrlf
        response.write "                      <table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
        response.write "                        <tr>" & vbcrlf
        response.write "                            <td>" & vbcrlf
        response.write "                                <div style=""color:#800000; margin-bottom:5px;"">" & vbcrlf
        response.write "                                  *** NOTE: Any changes made to these template fields will NOT be recognized " & vbcrlf
        response.write "                                  when creating a new MapPoint Type.<br />" & vbcrlf
        response.write "                                  All changes to template fields need to be made to the template itself " & vbcrlf
        response.write "                                  before creating a MapPoint Type." & vbcrlf
        response.write "                                </div>" & vbcrlf
        response.write "                                Template: <select name=""MPTemplateID"" id=""MPTemplateID"">" & vbcrlf
                                                          displayTemplateOptions
        response.write "                                </p>" & vbcrlf
        response.write "                                <input type=""button"" name=""previewButton"" id=""previewButton"" value=""Preview Template"" class=""button"" onclick=""getTemplateFields('MPTemplateID')"" />" & vbcrlf
        response.write "                            </td>" & vbcrlf
        response.write "                        </tr>" & vbcrlf
        response.write "                        <tr>" & vbcrlf
        response.write "                            <td>" & vbcrlf
        response.write "                                <iframe name=""previewMPTemplateFields"" id=""previewMPTemplateFields"" frameborder=""0"" src="""" width=""0"" height=""0""></iframe>" & vbcrlf
        response.write "                            </td>" & vbcrlf
        response.write "                        </tr>" & vbcrlf
        response.write "                      </table>" & vbcrlf
        response.write "                    </fieldset>" & vbcrlf
        response.write "                </td>" & vbcrlf
        response.write "            </tr>" & vbcrlf
     end if

  end if

  response.write "          </table>" & vbcrlf
  response.write "          </p>" & vbcrlf
  response.write "          <p>" & vbcrlf

 'Retrieve any/all fields related to this Map-Point Type
 'ONLY show these field if "screen mode" = "EDIT"
  if lcl_screen_mode = "EDIT" then
     displayMPTypesFields session("orgid"), lcl_mappoint_typeid, lcl_isRootAdmin, lcl_isLimited, False
  end if

 'Display the bottom row of buttons
  displayButtons "BOTTOM", lcl_screen_mode, lcl_canDelete, lcl_isLimited, lcl_isTemplate

  response.write "          </p>" & vbcrlf
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
     response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='mappoints_types_list.asp" & lcl_return_parameters & "'"" />" & vbcrlf
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

  response.write "<div>" & vbcrlf

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