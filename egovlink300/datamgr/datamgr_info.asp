<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!-- #include file="datamgr_build_sections_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_info.asp
' AUTHOR:   David Boyer
' CREATED:  04/29/2011
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the custom "Data Manager" data for a specific DM Data record.
'
' MODIFICATION HISTORY
' 1.0  04/29/11	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 Dim oDataMgr

 set oDataMgr = New classOrganization

 lcl_dmid       = ""
 lcl_feature    = ""
 lcl_return_url = ""

'Determine which DM Data record is to be display
 if request("dm") <> "" then
    if not containsApostrophe(request("dm")) then
       lcl_dmid = trim(request("dm"))
    end if
 end if

'Validate the DMID
 if lcl_dmid <> "" then
    if isnumeric(lcl_dmid) then
       on error resume next
       lcl_dmid = CLng(lcl_dmid)
       if err.number <> 0 then response.redirect "datamgr.asp"
       on error goto 0
    else
       response.redirect "datamgr.asp"
    end if
 else
    response.redirect "datamgr.asp"
 end if

'Validate the feature
 if request("f") <> "" then
    if not containsApostrophe(request("f")) then
       lcl_feature = request("f")

       if lcl_return_url <> "" then
          lcl_return_url = lcl_return_url & "&f=" & lcl_feature
       else
          lcl_return_url = "?f=" & lcl_feature
       end if
    end if
 end if

'Validate the DM TypeID
 if request("d") <> "" then
    if not containsApostrophe(request("d")) then
       lcl_dmt = request("d")

       if lcl_return_url <> "" then
          lcl_return_url = lcl_return_url & "&d=" & lcl_dmt
       else
          lcl_return_url = "?d=" & lcl_dmt
       end if
    end if
 end if

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

'Set up the page variables
 lcl_dm_typeid        = 0
 lcl_description      = ""
 lcl_dm_latitude      = ""
 lcl_dm_longitude     = ""
 lcl_dm_displaymap    = 1
 lcl_layoutid         = 0
 lcl_defaultzoomlevel = "13"
 lcl_enableOwnerMaint = false
 lcl_assignedto       = 0
 
'Retrieve the DM Data
 sSQL = "SELECT "
 sSQL = sSQL & " dmd.dm_typeid, "
 sSQL = sSQL & " dmt.description, "
 sSQL = sSQL & " dmd.latitude, "
 sSQL = sSQL & " dmd.longitude, "
 sSQL = sSQL & " dmt.displayMap, "
 sSQL = sSQL & " dmt.layoutid, "
 sSQL = sSQL & " dmt.defaultzoomlevel, "
 sSQL = sSQL & " dmt.enableOwnerMaint, "
 sSQL = sSQL & " dmt.assignedto "
 sSQL = sSQL & " FROM egov_dm_data dmd, "
 sSQL = sSQL &      " egov_dm_types dmt "
 sSQL = sSQL & " WHERE dmd.dm_typeid = dmt.dm_typeid "
 sSQL = sSQL & " AND dmd.orgid = " & iorgid
 sSQL = sSQL & " AND dmd.dmid = " & lcl_dmid
 'sSQL = sSQL & " AND dmd.isActive = 1 "

	set oDMDInfo = Server.CreateObject("ADODB.Recordset")
	oDMDInfo.Open sSQL, Application("DSN"), 3, 1

 if not oDMDInfo.eof then
    lcl_dm_typeid        = oDMDInfo("dm_typeid")
    lcl_description      = oDMDInfo("description")
    lcl_dm_latitude      = oDMDInfo("latitude")
    lcl_dm_longitude     = oDMDInfo("longitude")
    lcl_dm_displaymap    = oDMDInfo("displayMap")
    lcl_layoutid         = oDMDInfo("layoutid")
    lcl_defaultzoomlevel = oDMDInfo("defaultzoomlevel")
    lcl_enableOwnerMaint = oDMDInfo("enableOwnerMaint")
    lcl_assignedto       = oDMDInfo("assignedto")
 else
    lcl_dmid_exists = checkDMIDExists(iOrgID, lcl_dmid)

    if not lcl_dmid_exists then
       response.redirect "datamgr.asp" & lcl_return_url
    end if
 end if

 oDMDInfo.close
 set oDMDInfo = nothing

'Check for cookies
 lcl_cookie_userid = ""

 if request.cookies("userid") <> "" then
    lcl_cookie_userid = request.cookies("userid")
 end if

'Determine if there is an owner and if user is signed in if he/she is the owner
 lcl_ownerExists = checkDMOwnerExists(lcl_dmid)

 getDMOwnerEditorInfo lcl_dmid, _
                      lcl_cookie_userid, _
                      lcl_ownerid, _
                      lcl_ownertype, _
                      lcl_isOwner, _
                      lcl_isApproved, _
                      lcl_isWaitingApproval

'Set up page redirect
 session("RedirectPage") = request.servervariables("script_name") & "?" & request.querystring()
%>
<html>
<head>

 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
<!--  <link rel="stylesheet" type="text/css" href="layout_styles.css" /> -->

 	<script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/easyform.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/setfocus.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>

  <script type="text/javascript">var addthis_config = {"data_track_clickback":true};</script>
  <script type="text/javascript" src="https://s7.addthis.com/js/250/addthis_widget.js#pubid=egovlink"></script>
<%
'BEGIN: Google Maps javascript ------------------------------------------------
'Build the Google Maps javascript ONLY if we are displaying the map
 if lcl_dm_displaymap _
 AND (lcl_dm_latitude <> "" AND not isnull(lcl_dm_latitude)) _
 AND (lcl_dm_longitude <> "" AND not isnull(lcl_dm_longitude)) then
      lcl_onload   = "initialize(" & lcl_dm_latitude & "," & lcl_dm_longitude & ");"

    response.write "  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, user-scalable=no"" />" & vbcrlf
	sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"
    response.write "  <script type=""text/javascript"" src=""https://maps.google.com/maps/api/js?sensor=false&key=" & sGoogleMapAPIKey & """></script>" & vbcrlf
%>
<script type="text/javascript">
//  function openGoogleMap() {
//    var lcl_google_address = "";
//    var lcl_googlemap_url  = "";

//    lcl_google_address = document.getElementById("googleAddress").value;

//    if(lcl_google_address != "") {
//       lcl_googlemap_url += "http://maps.google.com/maps?hl=en&expIds=17259,17291,27615,27846,28155&sugexp=ldymls&xhr=t";
//       lcl_googlemap_url += "&q=" + lcl_google_address;
//       lcl_googlemap_url += "&cp=17&um=1&ie=UTF-8&sa=N&tab=wl";
//    }

//    openWin(lcl_googlemap_url,"","");
//  }

  //Set up global variables
  var myLatLng;
  var myOptions;

  function initialize(iLat, iLng) {
    //var myLatLng  = new google.maps.LatLng(iLat, iLng);
    myLatLng = new google.maps.LatLng(iLat, iLng);
    var sv   = new google.maps.StreetViewService();

    //var myOptions = {
    myOptions = {
       mapTypeId: google.maps.MapTypeId.ROADMAP,  //maptypes: ROADMAP, SATELLITE, HYBRID, TERRAIN
       zoom:      <%=lcl_defaultzoomlevel%>,
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

       //getPanoramaByLocation will return the nearest pano when the
       //given radius is 50 meters or less.
       //"processSVData is the callback from for the "getPanoramaByLocation".
       //   The parameters are passed to it from Google on the callback.
       sv.getPanoramaByLocation(myLatLng, 50, processSVData);
    }
  }

function processSVData(data, status) {
  if (status == google.maps.StreetViewStatus.OK) {

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
  } else {
      //Hide the canvas
      document.getElementById("map_canvas_streetview").style.display               = 'none';
      document.getElementById("map_canvas_streetview_navigation").style.display    = 'none';
      document.getElementById("map_canvas_streetview_note").style.display          = 'none';
      document.getElementById("map_canvas_streetview_getDirections").style.display = 'none';
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

function openGoogleURL() {
  w = 1000;
  h = 700;
  l = (screen.availWidth/2)-(w/2);
  t = (screen.availHeight/2)-(h/2);

  var lcl_googlemap_url  = '';
  var lcl_google_address = '';
  var lcl_google_city    = '';
  var lcl_google_state   = '';

  if(document.getElementById('googleADDRESS')) {
     lcl_google_address = document.getElementById('googleADDRESS').value;
  }

  //First we try and pull the values from the fields in section.
  if(document.getElementById('googleCITY')) {
     lcl_google_city = document.getElementById('googleCITY').value;
  }

  if(document.getElementById('googleSTATE')) {
     lcl_google_state = document.getElementById('googleSTATE').value;
  }

  //If the value is still empty, default it to the org's properties.
  if(lcl_google_city == '') {
     lcl_google_city = '<%=sDefaultCity%>';
  }

  if(lcl_google_state == '') {
     lcl_google_state = '<%=sDefaultState%>';
  }

  if(lcl_google_city != '') {
     if(lcl_google_address != '') {
        lcl_google_address += '+' + lcl_google_city;
     } else {
        lcl_google_address = lcl_google_city;
     }
  }

  if(lcl_google_state != '') {
     if(lcl_google_address != '') {
        lcl_google_address += '+' + lcl_google_state;
     } else {
        lcl_google_address = lcl_google_state;
     }
  }

  if(lcl_google_address != '') {
     //lcl_googlemap_url += "http://maps.google.com/maps?hl=en&expIds=17259,17291,27615,27846,28155&sugexp=ldymls&xhr=t";
     //lcl_googlemap_url += "&q=" + lcl_google_address;
     //lcl_googlemap_url += "&cp=17&um=1&ie=UTF-8&sa=N&tab=wl";
     lcl_googlemap_url += 'http://maps.google.com/maps';
     lcl_googlemap_url += '?q=' + lcl_google_address;
     lcl_googlemap_url += '&hl=en&z=16';
  } else {
     lcl_googlemap_url = 'http://maps.google.com/';
  }

  eval('window.open("' + lcl_googlemap_url + '", "_googlemap_dot", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}
</script>
<%
 end if
'END: Google Maps javascript --------------------------------------------------
%>
<script type="text/javascript">
  $(document).ready(function(){
    $('#returnButton').click(function(){
      location.href='datamgr.asp<%=lcl_return_url%>';
    });

    //$('#map_canvas_dot').click(function(){
    //  openGoogleURL();
    //});

    //$('#map_canvas_streetview').click(function(){
    //  openGoogleURL();
    //});

<% if lcl_enableOwnerMaint then %>
    $('#myDataMgrButton').click(function() {
      var lcl_url = '';

      lcl_url += 'mydatamgr.asp';
      lcl_url += '?f=<%=lcl_feature%>';

      location.href = lcl_url;
    });

    $('#ownerRequestButton').click(function() {
      $.post('send_dm_email.asp', {
        orgid:  '<%=iOrgID%>',
        userid: '<%=lcl_cookie_userid%>',
        dmid:   '<%=lcl_dmid%>',
        action: 'REQUEST_OWNER',
        isAjax: 'Y'
     }, function(result) {
        if(result == "sent") {
           $('#ownerRequestButton').hide('slow');
           displayScreenMsg('Your request has been successfully submitted.');
        }
     });

    });
<% end if %>
  });

function openEdit() {
  //w = 1000;
  //h = 700;
  //l = (screen.availWidth/2)-(w/2);
  //t = (screen.availHeight/2)-(h/2);

  editURL  = 'datamgr_maint.asp';
  editURL += '?dmid=<%=lcl_dmid%>';
  editURL += "&f=<%=lcl_feature%>";

  //eval('window.open("' + editURL + '", "_mpt_layout", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');  
  location.href = editURL;
}

function openLogin() {
  //w = 1000;
  //h = 700;
  //l = (screen.availWidth/2)-(w/2);
  //t = (screen.availHeight/2)-(h/2);

  editURL  = '../user_login.asp';

  //eval('window.open("' + editURL + '", "_mpt_layout", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');  
  location.href = editURL;
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

<style type="text/css">
  html       { height: 100% }
  body       { height: 100%; margin: 0px; padding: 0px; background-color: #efefef; }
  #screenMsg { color: #ff0000; text-align: right; }
</style>

</head>
<!--#include file="../include_top.asp"-->
<%
  RegisteredUserDisplay("../")

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""datamgr_centercontent"">" & vbcrlf

 'BEGIN: Buton Row ------------------------------------------------------------
  response.write "<table id=""buttonRow"">" & vbcrlf
  response.write "  <tr><td colspan=""2"" align=""right""><span id=""screenMsg"">&nbsp;</span></td></tr>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <input type=""button"" name=""returnButton"" id=""returnButton"" class=""button"" value=""Return to " & lcl_description & """ />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right"">&nbsp;" & vbcrlf

 '1. Check to see if the user is logged in.
 '2. If "yes" then check to see if there is an "owner" for this dm data record.
 '3. If "yes" then check to see if the user IS the owner.  If the user IS the owner then redirect to "edit" page.
 '4. If there is NO owner then show the "Request to become an Owner" button.
 '5. If there IS an owner, but the user is NOT the owner then:
 '   a. Check to see if the user is an editor.
 '   b. If NOT an editor then check to see if user has already requested to become an EDITOR and has been DENIED.
 '      - If "DENIED" then hide all of the buttons.
 '      - If "APPROVED" then redirect user to "edit" page.
  if lcl_enableOwnerMaint then
     if lcl_cookie_userid <> "" then
        if lcl_ownerExists then
           if lcl_ownerid > 0 then
              if lcl_isApproved OR lcl_isWaitingApproval then
                 response.redirect "datamgr_maint.asp?dmid=" & lcl_dmid & "&f=" & lcl_feature
              end if
           'else
           '   response.write "<input type=""button"" name=""editorRequestButton"" id=""editorRequestButton"" class=""button"" value=""Request to become an Editor"" />" & vbcrlf
           end if
        else
           if NOT lcl_isOwner AND NOT lcl_isWaitingApproval AND not lcl_isApproved then
              response.write "<input type=""button"" name=""ownerRequestButton"" id=""ownerRequestButton"" class=""button"" value=""Request to become an Owner"" />" & vbcrlf
           end if
        end if

        response.write "<input type=""button"" name=""myDataMgrButton"" id=""myDataMgrButton"" class=""button"" value=""My " & lcl_description & """ />" & vbcrlf
     else
        response.write "<input type=""button"" name=""loginButton"" id=""loginButton"" class=""button"" value=""Sign up/Log in"" onclick=""openLogin()"" />" & vbcrlf
     end if
  end if

  response.write "          &nbsp;&nbsp;&nbsp;" & vbcrlf
                            displayAddThisButtonNew iOrgID
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
 'END: Buton Row --------------------------------------------------------------

 'BEGIN: Build the Layout -----------------------------------------------------
    'Build the Layout
    'Retrieve any/all fields related to this DM Data
     lcl_displayFieldsetLegend    = False
     lcl_displayFieldsetBorder    = False
     lcl_displayAvailableSections = False
     lcl_section_mode             = "PUBLIC_VIEW"

     buildDMLayout lcl_layoutid, lcl_dm_typeid, lcl_dmid, lcl_displayFieldsetLegend, _
                   lcl_displayFieldsetBorder, lcl_displayAvailableSections, lcl_section_mode

 'END: Build the Layout -------------------------------------------------------

 'BEGIN: Google Map -----------------------------------------------------------
  'if lcl_dm_displaymap then
  '   response.write "<div id=""map_canvas"">&nbsp;</div>" & vbcrlf
  'end if
 'END: Google Map -------------------------------------------------------------

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!-- #include file="../include_bottom.asp" -->
