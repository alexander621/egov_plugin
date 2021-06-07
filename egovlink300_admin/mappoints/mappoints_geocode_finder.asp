<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints_geocode_finder.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the entries in the Blog
'
' MODIFICATION HISTORY
' 1.0 04/19/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Determine if the parent feature is "offline"
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Setup the page variables
 lcl_isPopup          = False
 lcl_formname         = ""
 lcl_return_latitude  = ""
 lcl_return_longitude = ""

'Determine if this is a popup or not.
 if request("popup") = "Y" then
    lcl_isPopup = True
 end if

'Retrieve the formname
 if request("fname") <> "" then
    lcl_formname = request("fname")
 end if

'Retrieve the latitude and longitude fields to pass the values back to.
 if request("lat") <> "" then
    lcl_return_latitude = request("lat")
 end if

 if request("long") <> "" then
    lcl_return_longitude = request("long")
 end if

 lcl_pagetitle = "Please enter the complete address (i.e. 4303 Hamilton Ave, Cincinnati, OH)"
 lcl_success   = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

 GetCityPoint session("orgid"), sLat, sLng

'Check for org features
' lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
' lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

'Check for user permissions
' lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
' lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

 lcl_onload = lcl_onload & "initialize();"
 lcl_onload = lcl_onload & "document.getElementById('address').focus();"
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

 	<script src="https://maps.google.com/maps?file=api&amp;v=2&amp;key=<%=GetGoogleMapApiKey()%>" type="text/javascript"></script>

<script language="javascript">
<!--
		var map;
  var myPano;
  var geocoder      = null;
  var control_field = "";

		function initialize() {
		  
		  if(GBrowserIsCompatible()) {
   			 map      = new GMap2(document.getElementById("map"));
       geocoder = new GClientGeocoder(); 

   			 map.addControl(new GMapTypeControl());
   			 map.addControl(new GLargeMapControl3D());
   			 map.setCenter(new GLatLng(<%=sLat%>, <%=sLng%>), 13);

       // Enable the additional map types within the map type collection
       map.enableRotation();

   			 var point;
		  }
		}

  function getGeoCoordinates() {
     var lcl_geo_coordinates = '';

     document.getElementById("geocoordinates").innerHTML = "";

     lcl_address = document.getElementById("address").value;

     if (lcl_address != null && lcl_address !="") {
         if (geocoder) {
             geocoder.getLatLng(
                lcl_address,
                function(point) {
                   if (!point) {
                       alert(lcl_address + " not found");

                       lcl_geo_coordinates += 'Try finding your coordinates here: <a href="http://www.batchgeocode.com/lookup/" target="_blank">here.</a>';

                       document.getElementById("geocoordinates").innerHTML = lcl_geo_coordinates;

                   } else {
                       var marker = new GMarker(point);
                       var myPoint;
                       var myPointIndex;
                       var lcl_latitude;
                       var lcl_longitude;
                       var lcl_display_info    = '';

                       myPoint = point.toString();
                       myPoint = myPoint.replace("(","");
                       myPoint = myPoint.replace(")","");

                       myPointIndex  = myPoint.indexOf(",");
                       lcl_latitude  = myPoint.substr(0,myPointIndex);
                       lcl_longitude = myPoint.substr(myPointIndex+2);

                       lcl_display_info += '<table border="0" cellspacing="0" cellpadding="2">';
                       lcl_display_info +=   '<tr>';
                       lcl_display_info +=       '<td colspan="2"><strong>' + lcl_address + '</strong></td>';
                       lcl_display_info +=   '</tr>';
                       lcl_display_info +=   '<tr>';
                       lcl_display_info +=       '<td>Latitude:</td>';
                       lcl_display_info +=       '<td>' + lcl_latitude + '</td>';
                       lcl_display_info +=   '</tr>';
                       lcl_display_info +=   '<tr>';
                       lcl_display_info +=       '<td>Longitude:</td>';
                       lcl_display_info +=       '<td>' + lcl_longitude + '</td>';
                       lcl_display_info +=   '</tr>';
                       lcl_display_info += '</table>';

                       map.addOverlay(marker);
                       marker.openInfoWindowHtml(lcl_display_info);

                       lcl_geo_coordinates += '<span style="color:#800000;">';
                       //lcl_geo_coordinates += 'SEARCHED: ';
                       lcl_geo_coordinates += '<strong>Latitude: </strong>' + lcl_latitude;
                       lcl_geo_coordinates += '&nbsp;-&nbsp;';
                       lcl_geo_coordinates += '<strong>Longitude: </strong>' + lcl_longitude;
                       lcl_geo_coordinates += '</span>';
                       lcl_geo_coordinates += '&nbsp;<input type="button" name="useSearchedButton" id="useSearchedButton" value="USE" class="button" onclick="returnGeocodes(\'SEARCHED\');" />';


                       document.getElementById("geocoordinates").innerHTML = lcl_geo_coordinates;
                       //alert(myPoint + " [" + lcl_latitude + "] - [" + lcl_longitude + "]");

                       document.getElementById("searched_latitude").value  = lcl_latitude;
                       document.getElementById("searched_longitude").value = lcl_longitude;
                   }
                }
             );
         }
     }
  }

  function returnGeocodes(p_type) {
    //Determine which values to return
    if (p_type == "CLICKED") {
        lcl_lat  = "clicked_latitude";
        lcl_long = "clicked_longitude";
    } else {
        lcl_lat  = "searched_latitude";
        lcl_long = "searched_longitude";
    }

    lcl_lat  = document.getElementById(lcl_lat).value;
    lcl_long = document.getElementById(lcl_long).value;

    //Verify that the return fields exist and a value has been found then return it.
    if (window.opener.document.<%=lcl_formname%>.<%=lcl_return_latitude%>) {
        window.opener.document.<%=lcl_formname%>.<%=lcl_return_latitude%>.value = lcl_lat;
    }

    if (window.opener.document.<%=lcl_formname%>.<%=lcl_return_longitude%>) {
        window.opener.document.<%=lcl_formname%>.<%=lcl_return_longitude%>.value = lcl_long;
    }
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

<%
   ShowHeader sLevel

  if not lcl_isPopup then
%>
<!--#Include file="../menu/menu.asp"--> 
<% end if %>

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;">
            <p>
            <table border="0" cellspacing="0" cellpadding="0" width="1000px">
              <tr>
                  <td><font size="+1"><strong><%=lcl_pagetitle%></strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;">&nbsp;</span></td>
              </tr>
            </table>
            </p>
          </div>
      </td>
  </tr>
  <tr>
      <td>
          <form action="#" onsubmit="getGeoCoordinates(); return false">
            <input type="hidden" name="searched_latitude" id="searched_latitude" value="" size="10" maxlength="50" />
            <input type="hidden" name="searched_longitude" id="searched_longitude" value="" size="10" maxlength="50" />
            <input type="hidden" name="clicked_latitude" id="clicked_latitude" value="" size="10" maxlength="50" />
            <input type="hidden" name="clicked_longitude" id="clicked_longitude" value="" size="10" maxlength="50" />
            <p>
              <input type="text" size="60" name="address" id="address" value="" />
              <input type="submit" value="Find!" />
            </p>
            <div id="geocoordinates" style="padding-bottom:5px;"></div>
            <div id="map" style="width:600px; height:400px;"></div>
          </form>

<%
  if lcl_isPopup then
     response.write "<div style=""width:600px; height:400px; text-align:center;"">" & vbcrlf
     response.write "<input type=""button"" name=""closeButton"" id=""closeButton"" value=""Close Window"" class=""button"" onclick=""parent.close();"" />" & vbcrlf
     response.write "</div>" & vbcrlf
  end if
%>
      </td>
  </tr>
</table>

  </div>
</div>
	
<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
