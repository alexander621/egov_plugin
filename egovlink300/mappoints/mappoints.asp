<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints.asp
' AUTHOR:   David Boyer
' CREATED:  03/05/2010
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays the Mayor's Blog
'
' MODIFICATION HISTORY
' 1.0  03/05/10	 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"
'Check to see if the feature is offline
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 Dim oMapPoints

 set oMapPoints = New classOrganization

 lcl_permission_check = "mappoints"
 lcl_feature          = ""
 lcl_mpt              = ""

'Determine what type of "map points" are to be displayed
'If there is only a single Map-Point Type then show those map-points.
'If there are mulitple then grab the first available
'Check to see if the "map point" exists
'NOTE: request("m") = mappoint_typeid
 if request("m") <> "" then
    if not containsApostrophe(request("m")) and isnumeric(request("m")) then
       lcl_mpt = request("m")
    end if
 else

   'Determine if we are accessing a specific Map-Point via a feature (link).
   'For example, features may be created for specific Map-Point Types such as "Available Properties".
   'Those features may be turned on for the public so that there will be a link in the menu bar, welcome page, and footer.
   'Those links will access this page.  The problem we run into is that we don't know which "Available Properties" map-point to
   'retrieve since it's a different ID for each org.  Passing in the feature name and having it associated to the Map-Point Type
   'will let us find the correct Map-Point Type.
   'NOTE: request("f") = feature name of the mappoint type to be displayed.
    if request("f") <> "" then
       if not containsApostrophe(request("f")) then
          lcl_feature          = Track_DBSafe(request("f"))
          lcl_permission_check = lcl_feature

          lcl_mpt = getMapPointTypeByFeature(iorgid, lcl_feature)
       end if
    end if
 end if

'If both a MapPoint Type ID or Feature have NOT been passed in then try and find a MapPoint Type ID to default to.
'If I cannot be found after this check then MapPoints has not been set up correctly.  
'The feature needs to be turned on more than likely.
 if lcl_mpt = "" then
   'Check to see if org has only one Map-Point Type.
   'If "yes" then show the Map-Points for that Map-Point Type
   'If "no" then grab the first one in the list (ordered by description)
    lcl_mappointtypes = getMapPointTypes(iorgid)

    if lcl_mappointtypes <> "" then
       sSQL = "SELECT distinct mpt.mappoint_typeid "
       sSQL = sSQL & " FROM egov_mappoints_types mpt, egov_mappoints mp "
       sSQL = sSQL & " WHERE mpt.mappoint_typeid = mp.mappoint_typeid "
       sSQL = sSQL & " AND mp.orgid = " & iorgid
       sSQL = sSQL & " AND mpt.mappoint_typeid IN (" & lcl_mappointtypes & ") "
       sSQL = sSQL & " AND mpt.isActive = 1 "
       sSQL = sSQL & " AND mp.isActive = 1 "
       sSQL = sSQL & " AND mp.latitude is not null "
       sSQL = sSQL & " AND mp.latitude <> 0.00 "
       sSQL = sSQL & " AND mp.longitude is not null "
       sSQL = sSQL & " AND mp.longitude <> 0.00 "

       set oGetDefaultMPTypeID = Server.CreateObject("ADODB.Recordset")
       oGetDefaultMPTypeID.Open sSQL, Application("DSN"), 3, 1

       if not oGetDefaultMPTypeID.eof then
          lcl_mpt = oGetDefaultMPTypeID("mappoint_typeid")
       end if

       oGetDefaultMPTypeID.close
       set oGetDefaultMPTypeID = nothing
    end if

    if lcl_mpt = "" then
       lcl_mpt = 0
    end if
 end if

'First Check 
 getMapPointsTypeInfo lcl_mpt, iorgid, lcl_feature, lcl_total_mptypes, lcl_mappoint_typeid, lcl_description, _
                      lcl_mappointcolor, lcl_displayMap, lcl_useAdvancedSearch

'Get the city's map "center point"
 GetCityPoint iorgid, sLat, sLng, sZoom

'Retrieve the search parameters
 'lcl_blogMonth = trim(request("blogMonth"))
 'lcl_blogYear  = trim(request("blogYear"))

'If BOTH the blogMonth AND blogYear are blank then grab the Month/Year of the latest, active, blog entry.
 'if  (lcl_blogMonth = "" OR isnull(lcl_blogMonth)) _
 'AND (lcl_blogYear = ""  OR isnull(lcl_blogYear)) then
 '     getCurrentBlogArchive iorgid, lcl_blogMonth, lcl_blogYear
 'end if

'Set to the current month/year if either are blank.
 'if lcl_blogMonth = "" OR isnull(lcl_blogMonth) then
 '   lcl_blogMonth = month(now)
 'end if

 'if lcl_blogYear = "" OR isnull(lcl_blogYear) then
 '   lcl_blogYear = year(now)
 'end if

'Build the query to be used within the mapping functions
 lcl_query = "SELECT mp.mappointid, "
 lcl_query = lcl_query & " mp.mappoint_typeid, "
 lcl_query = lcl_query & " mptf.mp_fieldid, "
 lcl_query = lcl_query & " mpv.fieldtype, "
 lcl_query = lcl_query & " mptf.fieldname, "
 lcl_query = lcl_query & " mpv.fieldvalue, "
 lcl_query = lcl_query & " mptf.displayInResults, "
 lcl_query = lcl_query & " mptf.resultsOrder, "
 lcl_query = lcl_query & " mp.streetnumber, "
 lcl_query = lcl_query & " mp.streetprefix, "
 lcl_query = lcl_query & " mp.streetaddress, "
 lcl_query = lcl_query & " mp.streetsuffix, "
 lcl_query = lcl_query & " mp.streetdirection, "
 lcl_query = lcl_query & " mp.latitude, "
 lcl_query = lcl_query & " mp.longitude, "
 lcl_query = lcl_query & " isnull(isnull(mp.mappointcolor,mpt.mappointcolor),'green') as mappointcolor, "
 lcl_query = lcl_query & " mpt.displayMap, "
 lcl_query = lcl_query & " mpt.useAdvancedSearch "
 lcl_query = lcl_query & " FROM egov_mappoints mp, "
 lcl_query = lcl_query &      " egov_mappoints_types mpt, "
 lcl_query = lcl_query &      " egov_mappoints_types_fields mptf, "
 lcl_query = lcl_query &      " egov_mappoints_values mpv "
 lcl_query = lcl_query & " WHERE mp.mappoint_typeid = mpt.mappoint_typeid "
 lcl_query = lcl_query & " AND mpt.mappoint_typeid = mptf.mappoint_typeid "
 lcl_query = lcl_query & " AND mpv.mp_fieldid = mptf.mp_fieldid "
 lcl_query = lcl_query & " AND mpv.mappointid = mp.mappointid "
 lcl_query = lcl_query & " AND mpv.mappoint_typeid = mp.mappoint_typeid "
 lcl_query = lcl_query & " AND mp.orgid = " & iorgid
 lcl_query = lcl_query & " AND mp.mappoint_typeid = " & lcl_mappoint_typeid
 lcl_query = lcl_query & " AND mpt.isActive = 1 "
 lcl_query = lcl_query & " AND mp.isActive = 1 "
 lcl_query = lcl_query & " AND mptf.displayInResults = 1 "
 lcl_query = lcl_query & " AND mp.latitude is not null "
 lcl_query = lcl_query & " AND mp.latitude <> 0.00 "
 lcl_query = lcl_query & " AND mp.longitude is not null "
 lcl_query = lcl_query & " AND mp.longitude <> 0.00 "

'Get the search criteria fields
 sSQL = "SELECT mp_fieldid, "
 sSQL = sSQL & " mappoint_typeid, "
 sSQL = sSQL & " fieldname, "
 sSQL = sSQL & " fieldtype "
 sSQL = sSQL & " FROM egov_mappoints_types_fields "
 'sSQL = sSQL & " WHERE mappoint_typeid = " & lcl_mpt
 sSQL = sSQL & " WHERE mappoint_typeid = " & lcl_mappoint_typeid
 sSQL = sSQL & " AND inPublicSearch = 1 "
'dtb_debug(sSQL)
 set oGetSearchCriteria = Server.CreateObject("ADODB.Recordset")
 oGetSearchCriteria.Open sSQL, Application("DSN"), 3, 1

 if not oGetSearchCriteria.eof then
    lcl_line_count = 0

    if NOT lcl_useAdvancedSearch then
       lcl_sc_searchfield = request("sc_searchfield_0")
    end if

    do while not oGetSearchCriteria.eof
       lcl_line_count = lcl_line_count + 1

      'Determine which search criteria layout to display.
      'if "lcl_useAdvancedSearch" = TRUE then show ALL of the fields selected to display in the search as searchable fields.
      'if "lcl_useAdvancedSearch" = FALSE then show only a single textbox and use that value to search all of the fields selected as searchable fields.
       if lcl_useAdvancedSearch then
          lcl_sc_searchfield = request("sc_searchfield_" & oGetSearchCriteria("mp_fieldid"))
       end if

       'if request("sc_searchfield_" & oGetSearchCriteria("mp_fieldid")) <> "" then
       '   lcl_fieldvalue = UCASE(request("sc_searchfield_" & oGetSearchCriteria("mp_fieldid")))
       if lcl_sc_searchfield <> "" then
          lcl_fieldvalue = UCASE(lcl_sc_searchfield)
          lcl_fieldvalue = dbsafe(lcl_fieldvalue)

          if NOT lcl_useAdvancedSearch then
             if lcl_line_count = 1 then
                lcl_query = lcl_query & " AND ("
             else
                lcl_query = lcl_query & " OR "
             end if
          else
             lcl_query = lcl_query & " AND "
          end if

          lcl_query = lcl_query &      " mp.mappointid in ("
          lcl_query = lcl_query &      " select distinct mpv" & oGetSearchCriteria("mp_fieldid") & ".mappointid "
          lcl_query = lcl_query &      " from egov_mappoints_values mpv" & oGetSearchCriteria("mp_fieldid")
          lcl_query = lcl_query &      " where UPPER(mpv" & oGetSearchCriteria("mp_fieldid")& ".fieldvalue) LIKE ('%" & lcl_fieldvalue & "%') "
          lcl_query = lcl_query &      " AND mpv" & oGetSearchCriteria("mp_fieldid") & ".mp_fieldid = " & oGetSearchCriteria("mp_fieldid")
          lcl_query = lcl_query &      ") "

          'lcl_query = lcl_query & " AND ("
          'lcl_query = lcl_query &      " mptf.mp_fieldid = " & oGetSearchCriteria("mp_fieldid")
          'lcl_query = lcl_query &      " AND UPPER(mpv.fieldvalue) LIKE ('%" & lcl_fieldvalue & "%') "
          'lcl_query = lcl_query &      ") "
       else
          lcl_fieldvalue = ""
       end if

       oGetSearchCriteria.movenext
    loop

    if NOT lcl_useAdvancedSearch then
       if lcl_sc_searchfield <> "" then
          lcl_query = lcl_query & ")"
       end if
    end if

 end if
'response.write lcl_query & "<br />"
 oGetSearchCriteria.close
 set oGetSearchCriteria = nothing

 lcl_query = lcl_query & " ORDER BY mp.mappointid, mptf.resultsOrder "

'Check for org "edit displays"
 lcl_orghasdisplay_mappoints_intro = OrgHasDisplay(iorgid,"mappoints_intro")

'Get the local date/time
 lcl_local_datetime = ConvertDateTimetoTimeZone(iOrgID)

'Set up the BODY "onload" and "onunload"
 if lcl_displayMap then
    'lcl_onload   = "initialize();listOrderInit();"
    lcl_onload   = "initialize();"
    lcl_onunload = "GUnload();"
 end if
%>
<html>
<head>
 	<title>E-Gov Services - <%=sOrgName%></title>

 	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
 	<link rel="stylesheet" type="text/css" href="../global.css" />
 	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

  <link rel="stylesheet" type="text/css" href="mapstyle.css" />

 	<script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/easyform.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/setfocus.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>

<% if lcl_displayMap then %>
 	<!--script src="http://maps.google.com/maps?file=api&amp;v=3&amp;key=<%=GetGoogleMapApiKey(iorgid)%>" type="text/javascript"></script-->
	<script type="text/javascript" src="https://maps.google.com/maps/api/js?sensor=false&key=<%= sGoogleMapAPIKey %>"></script>
<% end if %>
 	<script type="text/javascript" src="../scripts/column_sorting.js"></script>

	<script type="text/javascript">
  var sorter = new TINY.table.sorter("sorter");

function listOrderInit() {
  sorter.head      = "head";
  sorter.asc       = "asc";
  sorter.desc      = "desc";
  sorter.even      = "evenrow";
  sorter.odd       = "oddrow";
  sorter.evensel   = "evenselected";
  sorter.oddsel    = "oddselected";
  //sorter.paginate  = true;
  //sorter.currentid = "currentpage";
  //sorter.limitid   = "pagelimit";
  sorter.init("mappoints",0);
}
</script>

<script language="javascript">
<% if lcl_displayMap then %>
  var infowindow = [];
  var bubbleInfo = [];
  function openInfoWindow(iRowCount,lat,lng) {
    infowindow.push(new google.maps.InfoWindow({
       size:     new google.maps.Size(50,50),
      position: new google.maps.LatLng(lat,lng),
       content:  bubbleInfo[iRowCount]
    }));

    google.maps.event.addListener(gmarkers[iRowCount], 'click', function() {
       infowindow[iRowCount].open(map,gmarkers[iRowCount]);
    });
  }
  var map;
		var gmarkers     = new Array(); 
		var badaddresses = new Array();

		var cm_map;
		var cm_mapMarkers = [];
		var cm_mapHTMLS   = [];

		// Create a base icon for all of our markers that specifies the
		// shadow, icon dimensions, etc.
		/*
		var cm_baseIcon = new GIcon();
		cm_baseIcon.shadow           = "http://www.google.com/mapfiles/shadow50.png";
		cm_baseIcon.iconSize         = new GSize(20, 34);
		cm_baseIcon.shadowSize       = new GSize(37, 34);
		cm_baseIcon.iconAnchor       = new GPoint(9, 34);
		cm_baseIcon.infoWindowAnchor = new GPoint(9, 2);
		cm_baseIcon.infoShadowAnchor = new GPoint(18, 25);
		*/

		var param_wsId              = "od6";
		var param_ssKey             = '<%=GetGoogleMapApiKey(iorgid)%>';
		var param_useSidebar        = true;
		var param_titleColumn       = "title";
		var param_descriptionColumn = "description";
		var param_latColumn         = "latitude";
		var param_lngColumn         = "longitude";
		var param_rankColumn        = "rank";
		var param_iconType          = "";
		var param_iconOverType      = "";
//		var param_iconType          = "<%=lcl_mappointcolor%>";
//		var param_iconOverType      = "<%=lcl_mappointcolor%>";


		//var points = new Array();  // for panning to a point
		var i  = 0;
		var ba = 0;
		var map;
  var myPano;
  var geocoder = null;
		var side_bar_html = '';

		function initialize() {
		  
		  //if(GBrowserIsCompatible()) {
		  if(1==1) {
       //myPano = new GStreetviewPanorama(document.getElementById("pano"));
   			 //map = new GMap2(document.getElementById("map"));
    var myLatlng = new google.maps.LatLng(<%=sLat%>, <%=sLng%>);
    var myOptions = {
       mapTypeId: google.maps.MapTypeId.ROADMAP,  //maptypes: ROADMAP, SATELLITE, HYBRID, TERRAIN
       zoom:      <%=sZoom%>,
       center:    myLatlng
    }
			 map = new google.maps.Map(document.getElementById("map"), myOptions);


   			 //map.addControl(new GLargeMapControl());
   			 //map.addControl(new GMapTypeControl());
   			 //map.addControl(new GLargeMapControl3D());
   			 //map.setCenter(new GLatLng(<%=sLat%>, <%=sLng%>), <%=sZoom%>);

       // Enable the additional map types within the map type collection
       //map.enableRotation();

       // Enable the street view overlay
       showHideStreetView("HIDE");

//       svOverlay = new GStreetviewOverlay();
//       map.addOverlay(svOverlay);

//       GEvent.addListener(map,"click", function(overlay,latlng) { 
//         if(latlng) {
//            showHideStreetView("SHOW");
//            myPano.setLocationAndPOV(latlng);
//         }
//       });

       //GEvent.addListener(map,"dblclick", function(overlay,latlng) { 
       //  if(latlng) {
       //     showHideStreetView("SHOW");
       //     myPano.setLocationAndPOV(latlng);
       //  }
       //});

       //GEvent.addListener(myPano, "error", handleNoFlash); 
    
       // Create all of the map-points
   			 var point;
     <%
       if lcl_total_mptypes > 0 then
          showPoints lcl_query, iorgid, lcl_mappoint_typeid

          response.write "side_bar_html = '<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""text-align:left"">' + side_bar_html + '</table>';" & vbcrlf
          response.write "document.getElementById(""side_bar"").innerHTML = side_bar_html;" & vbcrlf
       end if
     %>
		  }
		}

		//Creates a marker at the given point with the given number label
//function createMarker(point, rank, pointcolor, mappointlabel, sMsg) {
function createMarker(lat,lng, rank, pointcolor, mappointlabel, sMsg) {
		  //var markerOpts     = {};
		  //var nIcon          = new GIcon(cm_baseIcon);
    		  var lcl_pointcolor = "green";

                  if(pointcolor != "") {
                     lcl_pointcolor = pointcolor;
                  }

		  //markerOpts.icon  = nIcon;
		  //markerOpts.title = mappointlabel;		 

		  //var marker = new GMarker(point, markerOpts);

    		var pinColor = "FE7569";
    		if (lcl_pointcolor == "blue")
    		{
	    		pinColor = "839afa";
    		}
    		if (lcl_pointcolor == "green")
    		{
	    		pinColor = "92e415";
    		}
		if(lcl_pointcolor == "pink")
		{
			pinColor = "fb85f7"
		}
		  
    		var pinImage = new google.maps.MarkerImage("http://chart.apis.google.com/chart?chst=d_map_pin_letter&chld=" + (rank) + "|" + pinColor,
        		new google.maps.Size(21, 34),
        		new google.maps.Point(0,0),
        		new google.maps.Point(10, 34));
		
	 var marker = new google.maps.Marker({
               		position: new google.maps.LatLng(lat,lng),
               		map: map,
               		icon: pinImage,
       			animation: google.maps.Animation.DROP
            	});

		  //var marker = null;

		  //INCREMENT MARKER ARRAY TO REFERENCE OFF MAP LINKS
		  gmarkers[i]=marker;
		  //alert(i + " " + sMsg);
		  bubbleInfo[i] = sMsg;
		openInfoWindow(i,lat,lng);

    side_bar_html += '  <tr>';
    side_bar_html += '      <td><img src="mappoint_colors/bg_' + pointcolor + '.jpg" width="15" height="10" style="border:1pt solid #000000" valign="middle" />&nbsp;</td>';
    side_bar_html += '      <td>' + rank + '. <a href="javascript:myclick(' + i + ')">' + mappointlabel + '</a></td>';
    side_bar_html += '  </tr>';
		  i++;



		}

		//This function picks up the click and opens the corresponding info window
		function myclick(i) {
    google.maps.event.trigger(gmarkers[i], 'click', function() {
       infowindow[i].open(map,gmarkers[i]);
    });
		}

		function checkplot(iIndex) {

		  // IF FALSE POINT WAS NOT PLOTTED
		  //if (isbadaddress(iIndex) != true) {
		  //	alert('Error: This address could not be plotted!');
		  //}
		  //else
		  //{
		  // POINT WAS PLOTTED. MOVE TO MAP TO SHOW MARKER
		  gmarkers[iIndex - 1].show();
		  GEvent.trigger(gmarkers[iIndex - 1], 'click');
		  location.href='#';
		  //}
		}

  function enableStreetView() {
    var lcl_field = document.getElementById("enableStreetView");

    if(lcl_field.checked == true) {
       myPano    = new GStreetviewPanorama(document.getElementById("pano"));
       svOverlay = new GStreetviewOverlay();
       map.addOverlay(svOverlay);

       GEvent.addListener(map,"click", function(overlay,latlng) { 
         if(latlng) {
            showHideStreetView("SHOW");
            myPano.setLocationAndPOV(latlng);
            document.getElementById("side_bar").style.height="800px";
         }
       });

       //GEvent.addListener(map,"dblclick", function(overlay,latlng) { 
       //  if(latlng) {
       //     showHideStreetView("SHOW");
       //     myPano.setLocationAndPOV(latlng);
       //  }
       //});

       GEvent.addListener(myPano, "error", handleNoFlash); 
    } else {
       map.removeOverlay(svOverlay);
       showHideStreetView("HIDE");
       document.getElementById("side_bar").style.height="400px";
    }
  }

  function showHideStreetView(p_mode) {
    lcl_showHide = "none";

    if(p_mode == "SHOW" && document.getElementById("enableStreetView").checked == true) {
       lcl_showHide = "block";
    }

    document.getElementById("pano").style.display                 = lcl_showHide;
    document.getElementById("hideStreetViewButton").style.display = lcl_showHide;
  }

  function handleNoFlash(errorCode) { 
//     if (errorCode == 603) {
     if (errorCode == "FLASH_UNAVAILABLE") {
         alert("Error: Flash doesn't appear to be supported by your browser"); 
         return;
     }
  } 
<% end if %>
  //The two functions handle the row highlight and un-highlight for result lists when the mouse cursor moves over and off a record
  function mouseOverRow( oRow ) {
    oRow.style.backgroundColor = '#93bee1';
    oRow.style.cursor          = 'pointer';
  }

  function mouseOutRow( oRow ) {	
    oRow.style.backgroundColor = '';
    oRow.style.cursor          = '';
  }

function changeMap(p_MPTypeID) {
  location.href="mappoints.asp?m=" + p_MPTypeID;
}

function openMPInfo(p_ID) {

  lcl_feature = '<%=lcl_feature%>';
  lcl_mpt_id  = '<%=lcl_mappoint_typeid%>';

  var lcl_mappoint_url;
  lcl_mappoint_url  = "mappointsinfo.asp";
  lcl_mappoint_url += "?m=" + p_ID;

  if(lcl_mpt_id != "") {
     lcl_mappoint_url += "&mpt=" + lcl_mpt_id;
  }

  if(lcl_feature != "") {
     lcl_mappoint_url += "&f=" + lcl_feature;
  }

  location.href = lcl_mappoint_url;
}

</script>
</head>
<!--#include file="../include_top.asp"-->
<p>
<table border="0" cellspacing="0" cellpadding="0" width="800">
  <tr>
      <td>
          <font class="pagetitle"><%=lcl_description%> Map</font>
      </td>
      <td align="right">
          &nbsp;
      </td>
  </tr>
</table>
</p>
<% RegisteredUserDisplay("../") %>

<div id="content">
  <div id="centercontent">
<%
 'Determine if there is an "org display"
  if lcl_orghasdisplay_mappoints_intro then
					response.write GetOrgDisplay( iOrgId, "mappoints_intro" ) & vbcrlf
  end if

 'Show a dropdown list containing all of the active Map-Point Types to select from if more than one exist
  if lcl_total_mptypes > 1 then
     response.write "<div style=""margin-bottom:5px;"">" & vbcrlf
     response.write "  Show Map:" & vbcrlf
     response.write "  <select name=""mappoint_typeid"" id=""mappoint_typeid"" onchange=""changeMap(this.value);"">" & vbcrlf
                         displayMapPointTypes iorgid, lcl_mappoint_typeid
     response.write "  </select>" & vbcrlf
     response.write "<div>" & vbcrlf
  end if

  if lcl_displayMap then
%>
<p>
<table cellpadding="2" cellspacing="0" border="0" id="maptable">
  <tr valign="top">
      <td>
          <!--div align="right">Enable Street View: <input type="checkbox" name="enableStreetView" id="enableStreetView" value="Y" onclick="enableStreetView();" /></div-->
          <div name="map" id="map" align="center" style="width:600px; height:400px; border:1pt solid #000000;"></div>
          <div align="right"><a href="http://maps.google.com/support/bin/static.py?page=guide.cs&guide=21670&topic=21671&answer=144350" color="#0000ff" target="_blank">How to navigate in Google Maps</a></div>
          <div name="pano" id="pano" style="width:600px; height:400px"></div>
          <div name="hideStreetViewButton" id="hideStreetViewButton" align="right"><a href="javascript:showHideStreetView('HIDE');" color="#0000ff">Hide Street View</a></div>
      </td>
      <td align="center" valign="top" style="width: 200px;" nowrap="nowrap">
      <%
        if lcl_description <> "" then
           response.write "<strong>" & lcl_description & "</strong><br /><hr />" & vbcrlf
        end if
      %>
      <div id="side_bar" style="overflow:auto; width:210px; height:400px;"></div>
      </td>
  </tr>
</table>
</p>
<%
  end if

 'BEGIN: Search Criteria ------------------------------------------------------
  lcl_fields_per_line       = 3
  lcl_fields_per_line_count = 0
  lcl_sc_linecount          = 0

  sSQL = "SELECT mp_fieldid, "
  sSQL = sSQL & " mappoint_typeid, "
  sSQL = sSQL & " fieldname, "
  sSQL = sSQL & " fieldtype "
  sSQL = sSQL & " FROM egov_mappoints_types_fields "
  sSQL = sSQL & " WHERE mappoint_typeid = " & lcl_mpt
  sSQL = sSQL & " AND inPublicSearch = 1 "
  sSQL = sSQL & " ORDER BY resultsOrder "

  set oBuildSearchCriteria = Server.CreateObject("ADODB.Recordset")
  oBuildSearchCriteria.Open sSQL, Application("DSN"), 3, 1

  if not oBuildSearchCriteria.eof then
     if lcl_useAdvancedSearch then
        do while not oBuildSearchCriteria.eof
           lcl_sc_linecount          = lcl_sc_linecount + 1
           lcl_fields_per_line_count = lcl_fields_per_line_count + 1

           if lcl_sc_linecount = 1 then
              response.write "<p>" & vbcrlf
              response.write "   <fieldset>" & vbcrlf
              response.write "     <legend>Search Criteria&nbsp;</legend>" & vbcrlf
              response.write "     <p>" & vbcrlf
              response.write "     <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
              response.write "       <form name=""searchForm"" id=""searchForm"" method=""post"" action=""mappoints.asp"">" & vbcrlf
              response.write "         <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
              response.write "       <tr>" & vbcrlf
           end if

           if lcl_fields_per_line_count > lcl_fields_per_line then
              lcl_fields_per_line_count = 1
              response.write "       </tr>" & vbcrlf
              response.write "       <tr>" & vbcrlf
           end if

           response.write "           <td>" & oBuildSearchCriteria("fieldname") & ":</td>" & vbcrlf
           response.write "           <td><input type=""text"" name=""sc_searchfield_" & oBuildSearchCriteria("mp_fieldid") & """ id=""sc_searchfield_" & oBuildSearchCriteria("mp_fieldid") & """ value=""" & request("sc_searchfield_" & oBuildSearchCriteria("mp_fieldid")) & """ /></td>" & vbcrlf

           oBuildSearchCriteria.movenext
        loop

        if lcl_sc_linecount > 0 then
           response.write "       </tr>" & vbcrlf
           response.write "     </table>" & vbcrlf
           response.write "     <p>" & vbcrlf
           response.write "       <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
           response.write "     </p>" & vbcrlf
           response.write "   </fieldset>" & vbcrlf
           response.write "</p>" & vbcrlf
     end if

     else

        response.write "<p>" & vbcrlf
        response.write "   <fieldset>" & vbcrlf
        response.write "     <legend>Search Criteria&nbsp;</legend>" & vbcrlf
        response.write "     <p>" & vbcrlf
        response.write "     <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
        response.write "       <form name=""searchForm"" id=""searchForm"" method=""post"" action=""mappoints.asp"">" & vbcrlf
        response.write "         <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
        response.write "       <tr>" & vbcrlf
        response.write "           <td>Search:</td>" & vbcrlf
        response.write "           <td><input type=""text"" name=""sc_searchfield_0"" id=""sc_searchfield_0"" value=""" & request("sc_searchfield_0") & """ size=""30"" /></td>" & vbcrlf
        response.write "       </tr>" & vbcrlf
        response.write "     </table>" & vbcrlf
        response.write "     <p>" & vbcrlf
        response.write "       <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
        response.write "     </p>" & vbcrlf
        response.write "   </fieldset>" & vbcrlf
        response.write "</p>" & vbcrlf

     end if

     oBuildSearchCriteria.close
     set oBuildSearchCriteria = nothing

  end if
 'END: Search Criteria --------------------------------------------------------
%>
<p>
   <strong><%=lcl_description%></strong><br />
   <% showList lcl_query, iorgid, lcl_mappoint_typeid %>
</p>
<p>&nbsp;</p>

  </div>
</div>
<!-- #include file="../include_bottom.asp" -->
<%
'------------------------------------------------------------------------------
sub showPoints(p_SQL, p_orgid, iMapPointTypeID)

	 iPointCount             = 1
  iRowCount               = 0
  lcl_showpoints          = ""
  lcl_displayheaders      = ""
  lcl_previous_mappointid = 0
  lcl_latitude            = 0.00
  lcl_longitude           = 0.00

 	set oPoints = Server.CreateObject("ADODB.Recordset")
 	oPoints.Open p_SQL, Application("DSN"), 3, 1

  if not oPoints.eof then
    	do while not oPoints.eof

        iRowCount = iRowCount + 1

        if lcl_previous_mappointid <> oPoints("mappointid") then
           if iRowCount > 1 then
              lcl_showpoints = lcl_showpoints & "  <tr><td colspan=""2"" align=""center""><a href=""javascript:openMPInfo(" & lcl_previous_mappointid & ");"" color=""#0000ff"">[more details...]</a></td></tr>"
              lcl_showpoints = lcl_showpoints & "</table>"
              lcl_showpoints = lcl_showpoints & "</div>"

           			'response.write "point = new GLatLng(" & lcl_latitude & "," & lcl_longitude & ");" & vbcrlf
           			'response.write "map.addOverlay(createMarker(point, " & iPointCount & ",'" & lcl_mpcolor & "','" & lcl_label & "','" & lcl_showpoints & "'));" & vbcrlf

           			response.write "createMarker('" & lcl_latitude & "','" & lcl_longitude & "', " & iPointCount & ",'" & lcl_mpcolor & "','" & lcl_label & "','" & lcl_showpoints & "');" & vbcrlf

              iPointCount    = iPointCount + 1
              iRowCount      = 1
              lcl_latitude   = 0.00
              lcl_longitude  = 0.00
           end if

           lcl_showpoints = ""
           lcl_showpoints = lcl_showpoints & "<div class=""info"">"
           lcl_showpoints = lcl_showpoints & "<table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""text-align:left;"">"
           lcl_showpoints = lcl_showpoints &   "<tr>"
           lcl_showpoints = lcl_showpoints &       "<td colspan=""2"" align=""center""><strong># " & iPointCount & "</strong></td>"
           lcl_showpoints = lcl_showpoints &   "</tr>"
        end if

        if oPoints("fieldvalue") <> "" then
           lcl_fieldvalue = ""

           if oPoints("fieldvalue") <> "" then
              'lcl_fieldvalue = replace(replace(replace(oPoints("fieldvalue"),chr(10),""),chr(13),"<br />"),"'","\'")
              lcl_fieldvalue = formatFieldValue(oPoints("fieldvalue"))
           end if

           lcl_showpoints = lcl_showpoints &   "<tr>"
           lcl_showpoints = lcl_showpoints &       "<td><strong>" & oPoints("fieldname") & ": </strong></td>"
           lcl_showpoints = lcl_showpoints &       "<td>"         & lcl_fieldvalue       & "</td>"
           lcl_showpoints = lcl_showpoints &   "</tr>"
        end if

        lcl_previous_mappointid = oPoints("mappointid")
        lcl_label               = buildStreetAddress(oPoints("streetnumber"), oPoints("streetprefix"), oPoints("streetaddress"), oPoints("streetsuffix"), oPoints("streetdirection"))
        lcl_label               = replace(lcl_label,"'","\'")
        lcl_latitude            = oPoints("latitude")
        lcl_longitude           = oPoints("longitude")
        lcl_mpcolor             = oPoints("mappointcolor")

        oPoints.movenext
     loop

     if iRowCount > 1 then
        lcl_showpoints = lcl_showpoints & "  <tr><td colspan=""2"" align=""center""><a href=""javascript:openMPInfo(" & lcl_previous_mappointid & ");"" color=""#0000ff"">[more details...]</a></td></tr>"
        lcl_showpoints = lcl_showpoints & "</table>"
        lcl_showpoints = lcl_showpoints & "</div>"

        'response.write "point = new GLatLng(" & lcl_latitude & "," & lcl_longitude & ");" & vbcrlf
        'response.write "map.addOverlay(createMarker(point, " & iPointCount & ",'" & lcl_mpcolor & "','" & lcl_label & "','" & lcl_showpoints & "'));" & vbcrlf

     	response.write "createMarker('" & lcl_latitude & "','" & lcl_longitude & "', " & iPointCount & ",'" & lcl_mpcolor & "','" & lcl_label & "','" & lcl_showpoints & "');" & vbcrlf
     end if

  end if

  oPoints.close
  set oPoints = nothing

end sub

'------------------------------------------------------------------------------
sub showList(p_SQL, p_orgid, iMapPointTypeID)

  lcl_columns_displayed = "N"
  lcl_results_displayed = "N"

 'BEGIN: Map-Point Values - Columns -------------------------------------------
  iColumnCount        = 0
  lcl_display_headers = ""
  'lcl_display_headers = lcl_display_headers & "<table id=""mappoints"" cellpadding=""2"" cellspacing=""0"" border=""0"" class=""mappoints_sortable"" style=""width:950px"">" & vbcrlf

  response.write "<table id=""mappoints"" cellpadding=""2"" cellspacing=""0"" border=""0"" class=""mappoints_sortable"" style=""width:950px"">" & vbcrlf
  response.write "  <thead>" & vbcrlf

 'Get the fieldnames to be used as column headers
  sSQL = "SELECT distinct mptf.fieldname, mptf.resultsOrder "
  sSQL = sSQL & " FROM egov_mappoints_types_fields mptf, egov_mappoints mp, egov_mappoints_types mpt "
  sSQL = sSQL & " WHERE mptf.mappoint_typeid = mp.mappoint_typeid "
  sSQL = sSQL & " AND mptf.mappoint_typeid = mpt.mappoint_typeid "
  sSQL = sSQL & " AND mptf.orgid = " & p_orgid
  sSQL = sSQL & " AND mptf.mappoint_typeid = " & iMapPointTypeID
  sSQL = sSQL & " AND mpt.isActive = 1 "
  sSQL = sSQL & " AND mp.isActive = 1 "
  sSQL = sSQL & " AND mptf.displayInResults = 1 "
  sSQL = sSQL & " ORDER BY resultsOrder "

  set oMPColumns = Server.CreateObject("ADODB.Recordset")
  oMPColumns.Open sSQL, Application("DSN"), 3, 1

  if not oMPColumns.eof then
     'lcl_display_headers = lcl_display_headers & "  <thead>" & vbcrlf
     lcl_display_headers = lcl_display_headers & "  <tr valign=""bottom"">" & vbcrlf
     lcl_display_headers = lcl_display_headers & "      <th nowrap=""nowrap"" style=""width:75px""><span>Map #</span></th>" & vbcrlf
     lcl_display_headers = lcl_display_headers & "      <th class=""nosort"">&nbsp;</th>" & vbcrlf

     do while not oMPColumns.eof
        iColumnCount = iColumnCount + 1
        lcl_display_headers = lcl_display_headers & "      <th><span>" & oMPColumns("fieldname") & "</span></th>" & vbcrlf

        oMPColumns.movenext
     loop

     lcl_display_headers = lcl_display_headers & "  </tr>" & vbcrlf
     'lcl_display_headers = lcl_display_headers & "  </thead>" & vbcrlf

     response.write lcl_display_headers

  else
     response.write "  <tr><th class=""nosort"">&nbsp;</th></tr>" & vbcrlf
  end if

  oMPColumns.close
  set oMPColumns = nothing
 'END: Map-Point Values - Columns ---------------------------------------------

  response.write "  </thead>" & vbcrlf
  response.write "  <tbody>" & vbcrlf

 'BEGIN: Map-Point Values - Rows ----------------------------------------------
  iRowCount               = 1
  lcl_previous_mappointid = 0
  lcl_bgcolor             = "#ffffff"
  lcl_scripts             = ""

'dtb_debug(p_SQL)
  set oMapPointsList = Server.CreateObject("ADODB.Recordset")
  oMapPointsList.Open p_SQL, Application("DSN"), 3, 1

  if not oMapPointsList.eof then
     lcl_results_displayed = "Y"
     lcl_scripts           = lcl_scripts & "listOrderInit();"

     do while not oMapPointsList.eof

        'if iRowCount = 1 then
        '   response.write lcl_display_headers
        '   response.write "  <tbody>" & vbcrlf
        'end if

        if lcl_previous_mappointid <> oMapPointsList("mappointid") then
           'lcl_onclick = " onclick=""checkplot(" & iRowCount & ");"""
           lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
           lcl_onclick = " onclick=""openMPInfo(" & oMapPointsList("mappointid") & ");"""

           if iRowCount > 1 then
              response.write "  </tr>" & vbcrlf
           end if

           response.write "  <tr bgcolor=""" & lcl_bgcolor & """ onmouseover=""mouseOverRow(this);"" onmouseout=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
           response.write "      <td" & lcl_onclick & " align=""center"">" & iRowCount & ".</td>" & vbcrlf
           response.write "      <td align=""center"" onclick=""myclick(" & iRowCount-1 & ");"">" & vbcrlf
           'response.write "          <img src=""http://gmaps-samples.googlecode.com/svn/trunk/markers/" & oMapPointsList("mappointcolor") & "/marker" & iRowCount & ".png"" width=""14"" height=""24"" />" & vbcrlf
           response.write "          <img src=""mappoint_colors/bg_" & oMapPointsList("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
        end if

       'Setup/Format the fieldvalue for display
        'lcl_fieldvalue = "&nbsp;"
        lcl_fieldvalue = oMapPointsList("fieldvalue")

        'if oMapPointsList("fieldvalue") <> "" then
        '   lcl_fieldvalue = replace(replace(oMapPointsList("fieldvalue"),chr(10),""),chr(13),"<br />")
        if lcl_fieldvalue <> "" then
           lcl_fieldvalue = replace(replace(lcl_fieldvalue,chr(10),""),chr(13),"<br />")
        end if

        response.write "      <td" & lcl_onclick & ">" & lcl_fieldvalue & "</td>" & vbcrlf

       'Determine if the row count increases
        if lcl_previous_mappointid <> oMapPointsList("mappointid") then
           iRowCount = iRowCount + 1
        end if

        lcl_previous_mappointid = oMapPointsList("mappointid")

        oMapPointsList.movenext

     loop

     'response.write "  </tr>" & vbcrlf
     'response.write "  </tbody>" & vbcrlf
     'response.write "</table>" & vbcrlf

    'BEGIN: Results List Navigation -------------------------------------------
     'response.write "<div id=""controls"">" & vbcrlf
     'response.write "  <div id=""perpage"">" & vbcrlf
     'response.write "		  <select onchange=""sorter.size(this.value)"">" & vbcrlf
     'response.write "			   <option value=""3"">3</option>" & vbcrlf
     'response.write "			   <option value=""5"">5</option>" & vbcrlf
     'response.write "				  <option value=""10"" selected=""selected"">10</option>" & vbcrlf
     'response.write "				  <option value=""20"">20</option>" & vbcrlf
     'response.write "				  <option value=""50"">50</option>" & vbcrlf
     'response.write "				  <option value=""100"">100</option>" & vbcrlf
     'response.write "			 </select>" & vbcrlf
     'response.write "			  <span>Entries Per Page</span>" & vbcrlf
     'response.write "		</div>" & vbcrlf
     'response.write "		<div id=""navigation"">" & vbcrlf
     'response.write "		 	<img src=""images/first.gif"" width=""16"" height=""16"" alt=""First Page"" onclick=""sorter.move(-1,true)"" />" & vbcrlf
     'response.write "		 	<img src=""images/previous.gif"" width=""16"" height=""16"" alt=""First Page"" onclick=""sorter.move(-1)"" />" & vbcrlf
     'response.write "		 	<img src=""images/next.gif"" width=""16"" height=""16"" alt=""First Page"" onclick=""sorter.move(1)"" />" & vbcrlf
     'response.write "		 	<img src=""images/last.gif"" width=""16"" height=""16"" alt=""Last Page"" onclick=""sorter.move(1,true)"" />" & vbcrlf
     'response.write "		</div>" & vbcrlf
     'response.write "		<div id=""text"">Displaying Page <span id=""currentpage""></span> of <span id=""pagelimit""></span></div>" & vbcrlf
     'response.write "	</div>" & vbcrlf
    'END: Results List Navigation ---------------------------------------------

  else
     'response.write "  <tr bgcolor=""" & lcl_bgcolor & """ onmouseover=""mouseOverRow(this);"" onmouseout=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
     response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
     response.write "      <td colspan=""" & iColumnCount + 2 & """ class=""nosort"">" & vbcrlf
     response.write "          <p>No " & lcl_description & " could be found that match your search criteria.</p>" & vbcrlf
     response.write "      </td>" & vbcrlf
  end if

 	oMapPointsList.close
 	set oMapPointsList = nothing
 'END: Map-Point Values - Rows ------------------------------------------------

  response.write "  </tr>" & vbcrlf
  response.write "  </tbody>" & vbcrlf
  response.write "</table>" & vbcrlf

  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if

end sub
%>
