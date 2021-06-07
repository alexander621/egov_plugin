<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectorroute.asp
' AUTHOR: Steve Loar
' CREATED: 08/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays the inspection route details for printing.
'
' MODIFICATION HISTORY
' 1.0   08/12/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iRowCount, sLat, sLng, sGoogleMapAPIKey

iRowCount = 0

sSql = session("sSql")  ' This is filled in by the calling page
sSql = sSql & " ORDER BY routeorder, I.scheduleddate"

'GET CITY'S MAP CENTER POINT
 GetCityPoint sLat, sLng 

sGoogleMapAPIKey = GetGoogleMapApiKey()
sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"

%>

<html>
<head>
	<title>E-Gov Permit Inspection Map</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<!--script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=<%= sGoogleMapAPIKey %>" type="text/javascript"></script-->
	<script type="text/javascript" src="https://maps.google.com/maps/api/js?sensor=false&key=<%= sGoogleMapAPIKey %>"></script>

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		//<![CDATA[


  		var infowindow = [];
  		var bubbleInfo = [];
		var gmarkers     = new Array(); 
  function openInfoWindow(iRowCount,point) {
    infowindow.push(new google.maps.InfoWindow({
       size:     new google.maps.Size(50,50),
      position: point,
       content:  bubbleInfo[iRowCount]
    }));

    google.maps.event.addListener(gmarkers[iRowCount], 'click', function() {
       infowindow[iRowCount].open(map,gmarkers[iRowCount]);
    });
  }
		var badaddresses = new Array();

		var cm_map;
		var cm_mapMarkers = [];
		var cm_mapHTMLS = [];

		// Create a base icon for all of our markers that specifies the
		// shadow, icon dimensions, etc.
		/*
		var cm_baseIcon = new GIcon();
		cm_baseIcon.shadow = "http://www.google.com/mapfiles/shadow50.png";
		cm_baseIcon.iconSize = new GSize(20, 34);
		cm_baseIcon.shadowSize = new GSize(37, 34);
		cm_baseIcon.iconAnchor = new GPoint(9, 34);
		cm_baseIcon.infoWindowAnchor = new GPoint(9, 2);
		cm_baseIcon.infoShadowAnchor = new GPoint(18, 25);
		var param_wsId = "od6";
		var param_ssKey = '<% = sGoogleMapAPIKey %>';
		var param_useSidebar = true;
		var param_titleColumn = "title";
		var param_descriptionColumn = "description";
		var param_latColumn = "latitude";
		var param_lngColumn = "longitude";
		var param_rankColumn = "rank";
		var param_iconType = "green";
		var param_iconOverType = "green";
		*/



		//var points = new Array();  // for panning to a point
		var i  = 0;
		var ba = 0;
		var map;
		var side_bar_html = '';

		function load() {
		  
		  /*
		  if(GBrowserIsCompatible()) {
			 map = new GMap2(document.getElementById("map"));
			 map.addControl(new GLargeMapControl());
			 map.addControl(new GMapTypeControl());
			 map.setCenter(new GLatLng(<%=sLat%>, <%=sLng%>), 13);

			 var point;

			 <% ShowPoints sSql %>

			 document.getElementById("side_bar").innerHTML = side_bar_html;
		  }
		*/
		}
  		function initialize() {
    			var myLatlng = new google.maps.LatLng(<%=sLat%>, <%=sLng%>);
    			var myOptions = {
       					mapTypeId: google.maps.MapTypeId.ROADMAP,  //maptypes: ROADMAP, SATELLITE, HYBRID, TERRAIN
       					zoom:      13,
       					center:    myLatlng
    			}

    			map = new google.maps.Map(document.getElementById("map"), myOptions);

			<% ShowPoints sSql %>
			 document.getElementById("side_bar").innerHTML = side_bar_html;
  		}

		//Creates a marker at the given point with the given number label
		function createMarker(point, rank, permitno, sMsg) {

		  //var markerOpts = {};
		  //var nIcon = new GIcon(cm_baseIcon);
    		  var lcl_pointcolor = "green";

    		var pinImage = new google.maps.MarkerImage("https://chart.apis.google.com/chart?chst=d_map_pin_letter&chld=" + (rank) + "|92e415",
        		new google.maps.Size(21, 34),
        		new google.maps.Point(0,0),
        		new google.maps.Point(10, 34));

	 var marker = new google.maps.Marker({
               		position: point,
               		map: map,
               		icon: pinImage,
       			animation: google.maps.Animation.DROP
            	});
		/*
		  if(rank > 0 && rank < 100) {
			nIcon.imageOut = "http://gmaps-samples.googlecode.com/svn/trunk/" +
				"markers/" + param_iconType + "/marker" + rank + ".png";
			nIcon.imageOver = "http://gmaps-samples.googlecode.com/svn/trunk/" +
				"markers/" + param_iconOverType + "/marker" + rank + ".png";
			nIcon.image = nIcon.imageOut; 
		  } else { 
			nIcon.imageOut = "http://gmaps-samples.googlecode.com/svn/trunk/" +
				"markers/" + param_iconType + "/blank.png";
			nIcon.imageOver = "http://gmaps-samples.googlecode.com/svn/trunk/" +
				"markers/" + param_iconOverType + "/blank.png";
			nIcon.image = nIcon.imageOut;
		  }

		  markerOpts.icon = nIcon;
		  markerOpts.title = permitno;		 
		  var marker = new GMarker(point, markerOpts);
		  //INCREMENT MARKER ARRAY TO REFERENCE OFF MAP LINKS
		  */
		  bubbleInfo[i] = sMsg;
		  gmarkers[i]=marker;
		openInfoWindow(i,point);
		  // add a line to the side_bar html
          side_bar_html += '<a href="javascript:myclick(' + i + ')">#' + rank + ' ' + permitno + '</a><br />';


//infowindow[i].open(map,gmarkers[i]);

		  //google.maps.event.addListener(gmarkers[i], 'click', function() {
			////marker.openInfoWindowHtml(sMsg);
       			//infowindow[i].open(map,gmarkers[i]);
		  //});
		  /*
			
		  GEvent.addListener(marker, "click", function() {
			marker.openInfoWindowHtml(sMsg);
		  });
		  GEvent.addListener(marker, "mouseover", function() {
			marker.setImage(marker.getIcon().imageOver);
		  });
		  GEvent.addListener(marker, "mouseout", function() {
			marker.setImage(marker.getIcon().imageOut);
		  });
		  GEvent.addListener(marker, "infowindowopen", function() {
			marker.setImage(marker.getIcon().imageOver);
		  });
		  GEvent.addListener(marker, "infowindowclose", function() {
			marker.setImage(marker.getIcon().imageOut);
		  });
		  */
		  i++;
		  return marker;

		}

		// This function picks up the click and opens the corresponding info window
		function myclick(i) {
			//gmarkers[i].show();
			//GEvent.trigger(gmarkers[i], "click");
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
		  gmarkers[iIndex].show();
		  GEvent.trigger(gmarkers[iIndex], 'click');
		  location.href='#';
		  //}
		}

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

	//]]>

	//-->
	</script>

</head>

<body onload="initialize()">
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		
	<!--BEGIN: PAGE TITLE-->
	<!--END: PAGE TITLE-->

	<table cellpadding="2" cellspacing="0" border="1" id="maptable">
		<tr><td valign="top">
			<div align="center" id="map" style="width: 870px; height: 640px"></div>
		</td><td valign="top" style="width: 100px;" nowrap="nowrap">
			Inspections<br /><hr />
			<div id="side_bar"></div>
		</td></tr>
	</table>


	<!--BEGIN: PAGE TITLE-->
	<p>
		<font size="+1"><strong>Permit Inspection Route</strong></font><br /><br />
	</p>
	<!--END: PAGE TITLE-->

	<%  ShowList sSql %>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void GetCityPoint ByRef sLat, ByRef sLng 
'------------------------------------------------------------------------------------------------------------
Sub GetCityPoint( ByRef sLat, ByRef sLng )
   'Get the point to center the map
    Dim sSql, oRs

    sSql = "SELECT latitude, longitude FROM organizations WHERE orgid = " & session("orgid")

    Set oRs = Server.CreateObject("ADODB.Recordset")
    oRs.Open sSQL, Application("DSN"), 3, 1

    If Not oRs.EOF Then 
       sLat = oRs("latitude")
       sLng = oRs("longitude")
    End If 

    oRs.Close
    Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowPoints
'------------------------------------------------------------------------------------------------------------
Sub ShowPoints( ByVal sSql )
	Dim oRs, iPointCount, sMsg, iCurrentPermitId, iLongitude

	iPointCount = 0
	iCurrentPermitId = CLng(0)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iPointCount = iPointCount + 1
		If CDbl(oRs("latitude")) <> CDbl(0.00) Then 
			iLongitude = oRs("longitude") 
			sMsg = "<div class=""info"">"
			sMsg = sMsg & "<table border=""0"" cellpadding=""1"" cellspacing=""1"">"
			sMsg = sMsg & "<tr><td align=""left""><b>Route Order #:</b></td><td align=""left"">" & iPointCount & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left""><b>Permit #:</b></td><td align=""left"">" & GetPermitNumber( oRs("permitid") ) & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left""><b>Inspection:</b></td><td align=""left"">" & oRs("permitinspectiontype") & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left"" valign=""top""><b>Location:</b></td><td align=""left"">"
			sMsg = sMsg &  oRs("residentstreetnumber")
			If oRs("residentstreetprefix") <> "" Then
				sMsg = sMsg &  " " & oRs("residentstreetprefix")
			End If 
			sMsg = sMsg &  " " & oRs("residentstreetname")
			If oRs("streetsuffix") <> "" Then
				sMsg = sMsg &  " " & oRs("streetsuffix")
			End If 
			If oRs("streetdirection") <> "" Then
				sMsg = sMsg &  " " & oRs("streetdirection")
			End If 
			sMsg = sMsg &  " " & oRs("residentunit")
			sMsg = sMsg &  "<br />" & oRs("residentcity")
			sMsg = sMsg & "</td></tr></table>"
			sMsg = sMsg & "</div>"

			If CDbl(oRs("latitude")) <> CDbl(0.00) Then 
				response.write  vbcrlf & "point = new google.maps.LatLng(" & oRs("latitude") & "," & iLongitude & ");"
				'response.write vbcrlf & "map.addOverlay(createMarker(point, " & iPointCount & ",'" & GetPermitNumber( oRs("permitid") ) & "','" & sMsg & "'));" 
				response.write vbcrlf & "createMarker(point, " & iPointCount & ",'" & GetPermitNumber( oRs("permitid") ) & "','" & sMsg & "');" 
			End If 
		End If 

		oRs.MoveNext
    Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowList sSql 
'------------------------------------------------------------------------------------------------------------
Sub ShowList( ByVal sSql )
	Dim oRs, iRowCount, iPlotCount

	iRowCount = clng(0)
	iPlotCount = clng(-1)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<table id=""inspectorroute"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write "<tr><th>Route<br />Order</th><th>Permit #</th><th>Permit Type</th><th>Address</th><th>Scheduled</th><th>Inspection</th><th>Inspection<br />Status</th><th>Reinspection</th><th>Final</th><th>Contact</th><th>Notes</th><th>Inspector</th></tr>"

		Do While Not oRs.EOF
			'If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				response.write vbcrlf & "<tr onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				iRowCount = iRowCount + 1
				' Route Order
				response.write "<td  align=""center"""
				If oRs("locationtype") = "address" Then 
					iPlotCount = iPlotCount + 1
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				response.write iRowCount & "</td>"
				
				' Permit No
				response.write "<td nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount & ");"" "
				End If 
				response.write ">"
				response.write GetPermitNumber( oRs("permitid") )
				response.write "</td>"

				' Permit Type
				response.write "<td align=""center"" nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				response.write oRs("permittype") & "</td>"
				
				' Address/Location
				response.write "<td nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount & ");"" "
				End If 
				response.write ">"
				Select Case oRs("locationtype")

					Case "address"
						response.write oRs("residentstreetnumber")
						If oRs("residentstreetprefix") <> "" Then
							response.write " " & oRs("residentstreetprefix")
						End If 
						response.write " " & oRs("residentstreetname")
						If oRs("streetsuffix") <> "" Then
							response.write " " & oRs("streetsuffix")
						End If 
						If oRs("streetdirection") <> "" Then
							response.write " " & oRs("streetdirection")
						End If 
						response.write " " & oRs("residentunit")
						response.write "<br />" & oRs("residentcity")

					Case "location"
						response.write Replace(oRs("permitlocation"),Chr(10),"<br />")

					Case Else 
						response.write "&nbsp;"

				End Select  

				response.write "</td>"
				
				' Scheduled date and time 
				response.write "<td align=""center"" nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				response.write FormatDateTime( oRs("scheduleddate"),2 )
				If oRs("scheduledtime") <> "" Then
					response.write "<br />" & oRs("scheduledtime") & " " & oRs("scheduledampm")
				End If 
				response.write "</td>"
				
				' Inspection Type
				response.write "<td align=""center"" nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				response.write oRs("permitinspectiontype") & "</td>"
				
				' Status
				response.write "<td align=""center"" nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				response.write oRs("inspectionstatus") & "</td>"
				
				' Required
				response.write "<td align=""center"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				If oRs("isreinspection") Then
					response.write "Yes" 
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"

				' Final Insp
				response.write "<td align=""center"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				If oRs("isfinal") Then
					response.write "Yes" 
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"

				' Contact
				response.write "<td align=""center"" nowrap=""nowrap"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				If oRs("contact") = "" And oRs("contactphone") = "" Then
					response.write "&nbsp;"
				Else 
					response.write oRs("contact") & "<br />" & oRs("contactphone")
				End If 
				response.write "</td>"

				' Notes
				response.write "<td"
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				If oRs("schedulingnotes") <> "" Then 
					response.write oRs("schedulingnotes") 
				Else
					response.write "&nbsp;"
				End If 
				response.write "</td>"

				' Inspector
				response.write "<td align=""center"""
				If oRs("locationtype") = "address" Then 
					response.write " onclick=""checkplot("& iPlotCount &");"" "
				End If 
				response.write ">"
				response.write oRs("FirstName") & " " & oRs("LastName")
				response.write "</td>"

				response.write "</tr>"
			'End If 

			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</table>"
	Else
		response.write vbcrlf & "<p>No Permits could be found that match your search criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
