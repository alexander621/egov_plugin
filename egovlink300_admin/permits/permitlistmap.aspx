<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectorroute.asp
' AUTHOR: Steve Loar
' CREATED: 02/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays the permits on Google Map.
'
' MODIFICATION HISTORY
' 1.0   02/10/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, iRowCount, sLat, sLng

iRowCount = 0

sSql = session("sSql")  ' This is filled in by the calling page

'GET CITY'S MAP CENTER POINT
GetCityPoint sLat, sLng 


%>

<html>
<head>
	<title>E-Gov Permit Locations Map</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=<%= GetGoogleMapApiKey %>" type="text/javascript"></script>

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		//<![CDATA[

		var gmarkers     = new Array(); 
		var badaddresses = new Array();

		var cm_map;
		var cm_mapMarkers = [];
		var cm_mapHTMLS = [];

		// Create a base icon for all of our markers that specifies the
		// shadow, icon dimensions, etc.
		var cm_baseIcon = new GIcon();
		cm_baseIcon.shadow = "http://www.google.com/mapfiles/shadow50.png";
		cm_baseIcon.iconSize = new GSize(20, 34);
		cm_baseIcon.shadowSize = new GSize(37, 34);
		cm_baseIcon.iconAnchor = new GPoint(9, 34);
		cm_baseIcon.infoWindowAnchor = new GPoint(9, 2);
		cm_baseIcon.infoShadowAnchor = new GPoint(18, 25);
		var param_wsId = "od6";
		var param_ssKey = '<% = GetGoogleMapApiKey %>';
		var param_useSidebar = true;
		var param_titleColumn = "title";
		var param_descriptionColumn = "description";
		var param_latColumn = "latitude";
		var param_lngColumn = "longitude";
		var param_rankColumn = "rank";
		var param_iconType = "green";
		var param_iconOverType = "green";

		//var points = new Array();  // for panning to a point
		var i  = 0;
		var ba = 0;
		var map;
		var side_bar_html = '';

		function load() 
		{
			if( GBrowserIsCompatible() ) 
			{
				map = new GMap2(document.getElementById("map"));
				map.addControl(new GLargeMapControl());
				map.addControl(new GMapTypeControl());
				map.setCenter(new GLatLng(<%=sLat%>, <%=sLng%>), 13);

				var point;

				<% ShowPoints sSql %>

				document.getElementById("side_bar").innerHTML = side_bar_html;
			}
		}

		//Creates a marker at the given point with the given number label
		function createMarker(point, rank, permitno, sMsg) 
		{

		  var markerOpts = {};
		  var nIcon = new GIcon(cm_baseIcon);

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
		  gmarkers[i]=marker;
		  // add a line to the side_bar html
          side_bar_html += '<a href="javascript:myclick(' + i + ')">#' + rank + ' ' + permitno + '</a><br />';
		  i++;
			
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
		  return marker;

		}

		// This function picks up the click and opens the corresponding info window
		function myclick(i) {
			gmarkers[i].show();
			GEvent.trigger(gmarkers[i], "click");
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

		function doClose()
		{
			window.close();
			window.opener.focus();
		}

	//]]>

	//-->
	</script>

</head>

<body onload="load()" onunload="GUnload()">
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		
		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Permit Locations Map</strong></font><br /><br />
		</p>
		<!--END: PAGE TITLE-->

		<table cellpadding="2" cellspacing="0" border="1" id="maptable">
			<tr><td>
				<div align="center" id="map" style="width: 870px; height: 640px"></div>
			</td><td valign="top" style="width: 100px;" nowrap="nowrap">
				Permits<br /><hr />
				<div id="side_bar"></div>
			</td></tr>
		</table>

		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Mapped Permits</strong></font><br /><br />
		</p>
		<!--END: PAGE TITLE-->

		<%  ShowList( sSql ) %>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Sub GetCityPoint( ByRef sLat, ByRef sLng )
'------------------------------------------------------------------------------------------------------------
Sub GetCityPoint( ByRef sLat, ByRef sLng )
   'Get the point to center the map
    Dim sSql, oPoint

    sSql = "SELECT latitude, longitude FROM organizations WHERE orgid = " & session("orgid")

    Set oPoint = Server.CreateObject("ADODB.Recordset")
    oPoint.Open sSQL, Application("DSN"), 3, 1

    If Not oPoint.EOF Then 
       sLat = oPoint("latitude")
       sLng = oPoint("longitude")
    End If 

    oPoint.close
    Set oPoint = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowPoints_old( sSql )
'------------------------------------------------------------------------------------------------------------
Sub ShowPoints_old( ByVal sSql )
	Dim oPoint, iPointCount, sMsg, iCurrentPermitId, iLongitude

	iPointCount = 0
	iCurrentPermitId = CLng(0)


	Set oPoint = Server.CreateObject("ADODB.Recordset")
	oPoint.Open sSQL, Application("DSN"), 3, 1

	Do While Not oPoint.EOF
		If CDbl(oPoint("latitude")) <> CDbl(0.00) Then 
			iPointCount = iPointCount + 1
			iLongitude = oPoint("longitude") 
			sMsg = "<div class=""info"">"
			sMsg = sMsg & "<b># " & iPointCount & "</b><br />"
			sMsg = sMsg & "<b>Permit: </b>" & GetPermitNumber( oPoint("permitid") ) & "<br />"
			sMsg = sMsg & "<b>Type: </b>" & oPoint("permittype") & " &ndash; " & oPoint("permittypedesc") & "<br />"
			sMsg = sMsg & "<b>Location: </b>"
			sMsg = sMsg &  oPoint("residentstreetnumber")
			If oPoint("residentstreetprefix") <> "" Then
				sMsg = sMsg &  " " & oPoint("residentstreetprefix")
			End If 
			sMsg = sMsg &  " " & oPoint("residentstreetname")
			If oPoint("streetsuffix") <> "" Then
				sMsg = sMsg &  " " & oPoint("streetsuffix")
			End If 
			If oPoint("streetdirection") <> "" Then
				sMsg = sMsg &  " " & oPoint("streetdirection")
			End If 
			sMsg = sMsg &  " " & oPoint("residentunit")
			sMsg = sMsg &  "<br />" & oPoint("residentcity")
			sMsg = sMsg & "</div>"
			response.write  vbcrlf & "point = new GLatLng(" & oPoint("latitude") & "," & iLongitude & ");"
			response.write vbcrlf & "map.addOverlay(createMarker(point, " & iPointCount & ",'" & GetPermitNumber( oPoint("permitid") ) & "','" & sMsg & "'));" 
		End If 
		oPoint.MoveNext
    Loop 

	oPoint.close
	Set oPoint = Nothing 
End Sub 


 '------------------------------------------------------------------------------------------------------------
' Sub ShowPoints( sSql )
'------------------------------------------------------------------------------------------------------------
Sub ShowPoints( ByVal sSql )
	Dim oPoint, iPointCount, sMsg, iCurrentPermitId, iLongitude

	iPointCount = 0
	iCurrentPermitId = CLng(0)


	Set oPoint = Server.CreateObject("ADODB.Recordset")
	oPoint.Open sSQL, Application("DSN"), 3, 1

	Do While Not oPoint.EOF
		If CDbl(oPoint("latitude")) <> CDbl(0.00) Then 
			iPointCount = iPointCount + 1
			iLongitude = oPoint("longitude") 
			sMsg = "<div class=""info"">"
			sMsg = sMsg & "<table border=""0"" cellpadding=""1"" cellspacing=""1"">"
			sMsg = sMsg & "<tr><td>&nbsp;</td><td align=""left""><b># " & iPointCount & "</b></td></tr>"
			sMsg = sMsg & "<tr><td align=""left""><b>Permit: </b></td><td align=""left"">" & GetPermitNumber( oPoint("permitid") ) & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left""><b>Type: </b></td><td align=""left"">" & oPoint("permittype") & " &ndash; " & oPoint("permittypedesc") & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left"" valign=""top""><b>Location: </b></td><td align=""left"">"
			sMsg = sMsg &  oPoint("residentstreetnumber")
			If oPoint("residentstreetprefix") <> "" Then
				sMsg = sMsg &  " " & oPoint("residentstreetprefix")
			End If 
			sMsg = sMsg &  " " & oPoint("residentstreetname")
			If oPoint("streetsuffix") <> "" Then
				sMsg = sMsg &  " " & oPoint("streetsuffix")
			End If 
			If oPoint("streetdirection") <> "" Then
				sMsg = sMsg &  " " & oPoint("streetdirection")
			End If 
			sMsg = sMsg &  " " & oPoint("residentunit")
			sMsg = sMsg &  "<br />" & oPoint("residentcity")
			sMsg = sMsg & "</td></tr></table>"
			sMsg = sMsg & "</div>"
			response.write  vbcrlf & "point = new GLatLng(" & oPoint("latitude") & "," & iLongitude & ");"
			response.write vbcrlf & "map.addOverlay(createMarker(point, " & iPointCount & ",'" & GetPermitNumber( oPoint("permitid") ) & "','" & sMsg & "'));" 
		End If 
		oPoint.MoveNext
    Loop 

	oPoint.close
	Set oPoint = Nothing 
End Sub 


'------------------------------------------------------------------------------------------------------------
' Sub ShowList( sSql )
'------------------------------------------------------------------------------------------------------------
Sub ShowList( sSql )
	Dim oRs 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<table id=""inspectorroute"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write "<tr><th>Map #</th><th>Permit #</th><th>Permit Type</th><th>Address</th><th>Applicant</th><th>Status</th></tr>"

		Do While Not oRs.EOF
			If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				response.write vbcrlf & "<tr onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
				iRowCount = iRowCount + 1
				' Map #
				response.write "<td align=""center"" onclick=""checkplot("& iRowCount &");"">" & iRowCount & "</td>"
				
				' Permit No
				response.write "<td align=""center"" nowrap=""nowrap"" onclick=""checkplot("& iRowCount &");"">"
				sPermitNo = GetPermitNumber( oRs("permitid") )
				If sPermitNo = "" Then 
					response.write "&nbsp;"
				Else
					response.write sPermitNo
				End If 
				response.write "</td>"

				' Permit Type
				response.write "<td nowrap=""nowrap"" onclick=""checkplot("& iRowCount &");""> " & oRs("permittype") & " &ndash; " & oRs("permittypedesc") & "</td>"
				
				' Address
				response.write "<td nowrap=""nowrap"" onclick=""checkplot("& iRowCount &");"">"
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
				response.write "</td>"
				
				' Applicant
				response.write "<td onClick=""checkplot("& iRowCount &");"">"
				response.write GetPermitApplicantName( oRs("permitid") )
				response.write "</td>"

				' Permit Status
				If oRs("isonhold") Or oRs("isvoided") Or oRs("isexpired") Then 
					response.write "<td align=""center"" onClick=""checkplot("& iRowCount &");"">"
					If oRs("isonhold") Then 
						response.write "On Hold"
					Else
						If oRs("isvoided") Then 
							response.write "Voided"
						Else
							response.write "Expired"
						End If 
					End If 
					response.write "</td>"
				Else
					response.write "<td align=""center"" onClick=""checkplot("& iRowCount &");"">" & oRs("permitstatus") & "</td>"
				End If 

				response.write "</tr>"
			End If 

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
