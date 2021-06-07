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

sSql = session("PermitListSql")  ' This is filled in by the calling page

'GET CITY'S MAP CENTER POINT
GetCityPoint sLat, sLng 

sGoogleMapAPIKey = "AIzaSyCvkUmkSSC8QVN4h21QSUNaiKi_7b4e1eM"

%>

<html>
<head>
	<title>E-Gov Permit Locations Map</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script type="text/javascript" src="https://maps.google.com/maps/api/js?sensor=false&key=<%= sGoogleMapAPIKey %>"></script>

	<script language="javascript" src="../scripts/modules.js"></script>

	<script language="Javascript">
	<!--

		//<![CDATA[

		var gmarkers     = new Array(); 
		var badaddresses = new Array();

		var cm_map;
		var cm_mapMarkers = [];
		var cm_mapHTMLS = [];
		var mappoints  = [];
		var bubbleInfo = [];
		var permitno = [];
		var markers    = [];
		var infowindow = [];
		var map;




		//var points = new Array();  // for panning to a point
		var i  = 0;
		var ba = 0;
		var side_bar_html = '';

		function load() {
		  
    			var lcl_latitude  = '<%=sLat%>';
    			var lcl_longitude = '<%=sLng%>';

    			var latlng = new google.maps.LatLng(lcl_latitude, lcl_longitude);
    			var myOptions = {
        			zoom: 13,
        			center: latlng,
        			mapTypeId: google.maps.MapTypeId.ROADMAP
    			};
			
    			map = new google.maps.Map(document.getElementById("map"), myOptions);

			 <% ShowPoints sSql %>
    			for (var i=0; i < mappoints.length; i++)
    			{
       				createMarker(i);
    			}

			 document.getElementById("side_bar").innerHTML = side_bar_html;
		}

		//Creates a marker at the given point with the given number label
		function createMarker(rank) {
          		side_bar_html += '<a href="javascript:infowindow[' + rank + '].open(map,markers[' + rank + '])">#' + (rank+1) + ' ' + permitno[rank] + '</a><br />';

    			var image = new google.maps.MarkerImage("https://chart.apis.google.com/chart?chst=d_map_pin_letter&chld=" + (rank+1) + "|92e415",
        			new google.maps.Size(21, 34),
        			new google.maps.Point(0,0),
        			new google.maps.Point(10, 34));
	
    			markers.push(new google.maps.Marker({
       				position:  mappoints[rank],
       				map:       map,
       				draggable: false,
       				animation: google.maps.Animation.DROP,
       				icon:      image
    			}));
    			infowindow.push(new google.maps.InfoWindow({
       				size:     new google.maps.Size(50,50),
       				position:  mappoints[rank],
       				content:  bubbleInfo[rank]
    			}));

    			google.maps.event.addListener(markers[rank], 'click', function() {
       				infowindow[rank].open(map,markers[rank]);
    			});

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
		  gmarkers[iIndex].show();
		  GEvent.trigger(gmarkers[iIndex], 'click');
		  location.href='#';
		  //}
		}

		function doClose()
		{
			//window.close();
			//window.opener.focus();
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

	//]]>


	//-->
	</script>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<style>
			#content * { font-family:'Open Sans', sans-serif !important; font-size 14px !important;}
			.ui-button, .ui-button:hover { background-color: #2C5F93; color:white; margin-bottom:5px; }
			input, select, input.button { font-size: 14px; }
		</style>

</head>

<body onload="load();">
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		

		<table cellpadding="2" cellspacing="0" border="1" id="maptable">
			<tr><td valign="top">
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

		<%  ShowList sSql  %>

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
' Sub GetCityPoint( ByRef sLat, ByRef sLng )
'------------------------------------------------------------------------------------------------------------
Sub GetCityPoint( ByRef sLat, ByRef sLng )
   'Get the point to center the map
    Dim sSql, oRs

    sSql = "SELECT latitude, longitude FROM organizations WHERE orgid = " & session("orgid")

    Set oRs = Server.CreateObject("ADODB.Recordset")
    oRs.Open sSql, Application("DSN"), 3, 1

    If Not oRs.EOF Then 
       sLat = oRs("latitude")
       sLng = oRs("longitude")
    End If 

    oRs.close
    Set oRs = Nothing 

End Sub 


 '------------------------------------------------------------------------------------------------------------
' void ShowPoints sSql 
'------------------------------------------------------------------------------------------------------------
Sub ShowPoints( ByVal sSql )
	Dim oRs, iPointCount, sMsg, iCurrentPermitId, iLongitude

	iPointCount = clng(0)
	iCurrentPermitId = CLng(0)

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		' count them all so we keep they numbers the same as in the list
		iPointCount = iPointCount + 1

		If CDbl(oRs("latitude")) <> CDbl(0.00) Then 
			iPlotCount = iPlotCount + 1
			iLongitude = oRs("longitude") 
			sMsg = "<div class=""info"">"
			sMsg = sMsg & "<table border=""0"" cellpadding=""1"" cellspacing=""1"">"
			sMsg = sMsg & "<tr><td>&nbsp;</td><td align=""left""><b># " & iPointCount & "</b></td></tr>"
			sMsg = sMsg & "<tr><td align=""left""><b>Permit: </b></td><td align=""left"">" & GetPermitNumber( oRs("permitid") ) & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left""><b>Type: </b></td><td align=""left"">" & oRs("permittype") & " &ndash; " & oRs("permittypedesc") & "</td></tr>"
			sMsg = sMsg & "<tr><td align=""left"" valign=""top""><b>Location: </b></td><td align=""left"">"
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
			'response.write  vbcrlf & "point = new GLatLng(" & oRs("latitude") & "," & iLongitude & ");"
			'response.write vbcrlf & "map.addOverlay(createMarker(point, " & iPointCount & ",'" & GetPermitNumber( oRs("permitid") ) & "','" & sMsg & "'));" 

			response.write vbcrlf & "mappoints.push(new google.maps.LatLng(" & oRs("latitude") & "," & iLongitude & "));"
			response.write vbcrlf & "bubbleInfo.push('" & sMsg & "');"
			response.write vbcrlf & "permitno.push('" & GetPermitNumber(oRs("permitid")) & "');"

			'response.write vbcrlf & "createMarker(" & iPointCount & ",'" & GetPermitNumber(oRs("permitid")) & "');"
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
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<table id=""inspectorroute"" cellpadding=""2"" cellspacing=""0"" border=""0"">"
		response.write "<tr><th>Map #</th><th>Permit #</th><th>Permit Type</th><th>Address/Location</th><th>Applicant</th><th>Status</th></tr>"

		Do While Not oRs.EOF
			response.write vbcrlf & "<tr onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
			iRowCount = iRowCount + 1
			' Map #
			response.write "<td align=""center"""
			If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				iPlotCount = iPlotCount + 1
				response.write " onclick=""checkplot("& iPlotCount &");"""
			End If 
			response.write ">" & iRowCount & "</td>"
			
			' Permit No
			response.write "<td align=""center"" nowrap=""nowrap"""
			If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				response.write " onclick=""checkplot("& iPlotCount &");"""
			End If 
			response.write ">"
			sPermitNo = GetPermitNumber( oRs("permitid") )
			If sPermitNo = "" Then 
				response.write "&nbsp;"
			Else
				response.write sPermitNo
			End If 
			response.write "</td>"

			' Permit Type
			response.write "<td nowrap=""nowrap"""
			If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				response.write " onclick=""checkplot("& iPlotCount &");"""
			End If 
			response.write "> " & oRs("permittype") & " &ndash; " & oRs("permittypedesc") & "</td>"
			
			' Address
			response.write "<td nowrap=""nowrap"""
			If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				response.write " onclick=""checkplot("& iPlotCount &");"""
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
					response.write oRs("permitlocation")

				Case Else 
					response.write "&nbsp;"

			End Select  

			response.write "</td>"
			
			' Applicant
			response.write "<td"
			If CDbl(oRs("latitude")) <> CDbl(0.00) Then
				response.write " onClick=""checkplot("& iPlotCount &");"""
			End If 
			response.write ">"
			response.write GetPermitApplicantName( oRs("permitid") )
			response.write "</td>"

			' Permit Status
			If oRs("isonhold") Or oRs("isvoided") Or oRs("isexpired") Then 
				response.write "<td align=""center"""
				If CDbl(oRs("latitude")) <> CDbl(0.00) Then
					response.write " onClick=""checkplot("& iPlotCount &");"""
				End If 
				response.write ">"
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
				response.write "<td align=""center"""
				If CDbl(oRs("latitude")) <> CDbl(0.00) Then
					response.write " onClick=""checkplot("& iPlotCount &");"""
				End If 
				response.write ">" & oRs("permitstatus") & "</td>"
			End If 

			response.write "</tr>"

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
